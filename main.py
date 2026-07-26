"""
API de collecte — Étude clinique ANALOG.

Deux liens, deux endpoints :
  POST /api/submit/medecin   -> une soumission médecin (T0 "avant" ou T1 "après", champ timepoint)
  POST /api/submit/patient   -> une soumission patient (14 items + pré-questionnaire)
  GET  /api/export?key=ADMIN_KEY -> génère et renvoie un .xlsx à jour (temps réel : lit direct la DB)
  GET  /api/health           -> ping

Données anonymes (pas de nom, pas d'identifiant patient) -> pas d'exigence HDS
sur ce backend. Le médecin renseigne lui-même un identifiant (M01...) pour
permettre l'appariement avant/après ; ce n'est pas une donnée patient.

Sécurité (pilote, ~10 médecins / ~500 patients) :
  - STUDY_KEY : clé partagée envoyée par le frontend à chaque soumission.
    Filtre anti-bot basique (visible dans le JS du formulaire), pas une
    authentification forte.
  - ADMIN_KEY : clé pour télécharger l'export -> à garder confidentielle.
"""
import io
import json
import os
import datetime
from typing import Optional

from fastapi import FastAPI, HTTPException, Query, Depends
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from starlette.responses import StreamingResponse

from database import SessionLocal, Submission, init_db
import scoring

STUDY_KEY = os.environ.get("STUDY_KEY", "changeme-study-key")
ADMIN_KEY = os.environ.get("ADMIN_KEY", "changeme-admin-key")
ALLOWED_ORIGINS = os.environ.get("ALLOWED_ORIGINS", "*").split(",")

VALID_TYPES = {"medecin", "patient"}

app = FastAPI(title="ANALOG Study API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


@app.on_event("startup")
def on_startup():
    init_db()


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


class SubmissionPayload(BaseModel):
    study_key: str
    timepoint: Optional[str] = None          # T0 / T1 (médecin uniquement)
    medecin_id: Optional[str] = None         # auto-déclaré (médecin uniquement)
    centre: Optional[str] = None
    age: Optional[str] = None
    anciennete: Optional[str] = None         # médecin
    logiciel_actuel: Optional[str] = None    # médecin
    type_intervention: Optional[str] = None  # patient
    mode_prise_charge: Optional[str] = None  # patient
    support: Optional[str] = None            # patient
    answers: dict


@app.get("/api/health")
def health():
    return {"status": "ok", "time": datetime.datetime.utcnow().isoformat()}


@app.post("/api/submit/{questionnaire_type}")
def submit(questionnaire_type: str, payload: SubmissionPayload, db=Depends(get_db)):
    if questionnaire_type not in VALID_TYPES:
        raise HTTPException(404, f"Type de questionnaire inconnu: {questionnaire_type}")
    if payload.study_key != STUDY_KEY:
        raise HTTPException(403, "Clé d'étude invalide")

    row = Submission(
        questionnaire_type=questionnaire_type,
        medecin_id=payload.medecin_id,
        timepoint=payload.timepoint,
        centre=payload.centre,
        age=payload.age,
        anciennete=payload.anciennete,
        logiciel_actuel=payload.logiciel_actuel,
        type_intervention=payload.type_intervention,
        mode_prise_charge=payload.mode_prise_charge,
        support=payload.support,
        answers_json=json.dumps(payload.answers, ensure_ascii=False),
    )
    db.add(row)
    db.commit()
    db.refresh(row)
    return {"status": "recorded", "id": row.id}


def _autosize(ws):
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max(length + 2, 10), 60)


@app.get("/api/export")
def export(key: str = Query(...), db=Depends(get_db)):
    """
    Export temps réel : lit l'état actuel de la base à chaque appel, donc
    toujours à jour au moment du téléchargement. Pas besoin de régénérer
    un fichier en tâche de fond.
    """
    if key != ADMIN_KEY:
        raise HTTPException(403, "Clé admin invalide")

    rows = db.query(Submission).order_by(Submission.submitted_at).all()
    medecin_rows = [r for r in rows if r.questionnaire_type == "medecin"]
    patient_rows = [r for r in rows if r.questionnaire_type == "patient"]

    wb = Workbook()
    wb.remove(wb.active)

    # --- Feuille brute médecin (T0 et T1 mélangés, colonne timepoint distingue) ---
    ws = wb.create_sheet("Donnees_brutes_medecin")
    all_keys = []
    for r in medecin_rows:
        for k in json.loads(r.answers_json).keys():
            if k not in all_keys:
                all_keys.append(k)
    headers = ["id", "medecin_id", "timepoint", "centre", "age", "anciennete", "logiciel_actuel", "submitted_at"] + all_keys
    ws.append(headers)
    for r in medecin_rows:
        answers = json.loads(r.answers_json)
        ws.append([
            r.id, r.medecin_id, r.timepoint, r.centre, r.age, r.anciennete, r.logiciel_actuel,
            r.submitted_at.isoformat() if r.submitted_at else None,
        ] + [answers.get(k) for k in all_keys])
    _autosize(ws)

    # --- Feuille brute patient ---
    ws = wb.create_sheet("Donnees_brutes_patient")
    all_keys_p = []
    for r in patient_rows:
        for k in json.loads(r.answers_json).keys():
            if k not in all_keys_p:
                all_keys_p.append(k)
    headers_p = ["id", "centre", "age", "type_intervention", "mode_prise_charge", "support", "submitted_at"] + all_keys_p
    ws.append(headers_p)
    for r in patient_rows:
        answers = json.loads(r.answers_json)
        ws.append([
            r.id, r.centre, r.age, r.type_intervention, r.mode_prise_charge, r.support,
            r.submitted_at.isoformat() if r.submitted_at else None,
        ] + [answers.get(k) for k in all_keys_p])
    _autosize(ws)

    # --- Feuille "Scores calculés" ---
    ws = wb.create_sheet("Scores_calcules")
    ws.append([
        "medecin_id", "centre",
        "SUS_avant", "SUS_apres", "delta_SUS",
        "NASA_avant", "NASA_apres", "delta_NASA",
        "Bloc3_avant_/100", "Bloc3_apres_/100", "delta_Bloc3",
        "Incident_declare", "Incident_risque_patient",
    ])

    by_med_tp = {}
    for r in medecin_rows:
        if not r.medecin_id:
            continue
        by_med_tp.setdefault(r.medecin_id, {})[r.timepoint] = (json.loads(r.answers_json), r.centre)

    for mid, tp_map in sorted(by_med_tp.items()):
        a, centre_a = tp_map.get("T0", ({}, None))
        b, centre_b = tp_map.get("T1", ({}, None))
        sus_a, sus_b = scoring.sus_score(a, "av_"), scoring.sus_score(b, "ap_")
        nasa_a, nasa_b = scoring.nasa_tlx_score(a, "av_"), scoring.nasa_tlx_score(b, "ap_")
        bloc3_a, bloc3_b = scoring.bloc3_score(a, "av_"), scoring.bloc3_score(b, "ap_")
        ws.append([
            mid, centre_a or centre_b,
            sus_a, sus_b, (round(sus_b - sus_a, 1) if sus_a is not None and sus_b is not None else None),
            nasa_a, nasa_b, (nasa_b - nasa_a if nasa_a is not None and nasa_b is not None else None),
            bloc3_a, bloc3_b, (bloc3_b - bloc3_a if bloc3_a is not None and bloc3_b is not None else None),
            b.get("ap_inc1"), b.get("ap_inc2"),
        ])
    _autosize(ws)

    # --- Feuille "Scores patient" ---
    ws = wb.create_sheet("Scores_patient")
    ws.append(["id", "centre", "age", "type_intervention", "mode_prise_charge", "support", "score_/100",
               "acces_espace_patient", "preconsultation", "technique", "impression_globale"])
    for r in patient_rows:
        answers = json.loads(r.answers_json)
        score, dims = scoring.patient_score(answers)
        ws.append([
            r.id, r.centre, r.age, r.type_intervention, r.mode_prise_charge, r.support, score,
            dims["acces_espace_patient"], dims["preconsultation"], dims["technique"], dims["impression_globale"],
        ])
    _autosize(ws)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"ANALOG_export_{datetime.date.today().isoformat()}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )
