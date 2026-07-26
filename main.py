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
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from starlette.responses import StreamingResponse

from database import SessionLocal, Submission, init_db
import scoring
import labels

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


# ============================== EXPORT EXCEL ==============================

HEADER_FILL = PatternFill(start_color="26215C", end_color="26215C", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
HEADER_ALIGN = Alignment(wrap_text=True, vertical="center")
TITLE_FONT = Font(bold=True, size=14, color="26215C")
SUBTITLE_FONT = Font(italic=True, size=10, color="71717A")
KPI_LABEL_FONT = Font(size=10, color="71717A")
KPI_VALUE_FONT = Font(bold=True, size=18, color="26215C")


def _num(v):
    """Convertit en int/float si possible, pour que les cellules Excel soient de vrais nombres (SOMME, MOYENNE...)."""
    if v is None or v == "":
        return None
    try:
        if isinstance(v, str) and "." in v:
            return float(v)
        return int(v)
    except (TypeError, ValueError):
        return v


def _style_header_row(ws, row_idx, n_cols):
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN


def _add_table(ws, n_rows, n_cols, name):
    if n_rows < 1:
        return
    ref = f"A1:{get_column_letter(n_cols)}{n_rows + 1}"
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showRowStripes=True, showFirstColumn=False
    )
    ws.add_table(table)


def _autosize(ws, max_width=55, header_height=42):
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max(length + 2, 12), max_width)
    ws.row_dimensions[1].height = header_height
    ws.freeze_panes = "A2"


def _kpi_card(ws, row, col, label_text, value):
    ws.cell(row=row, column=col, value=label_text).font = KPI_LABEL_FONT
    ws.cell(row=row + 1, column=col, value=value).font = KPI_VALUE_FONT


@app.get("/api/export")
def export(key: str = Query(...), db=Depends(get_db)):
    """
    Export temps réel : lit l'état actuel de la base à chaque appel, donc
    toujours à jour au moment du téléchargement.
    """
    if key != ADMIN_KEY:
        raise HTTPException(403, "Clé admin invalide")

    rows = db.query(Submission).order_by(Submission.submitted_at).all()
    medecin_rows = [r for r in rows if r.questionnaire_type == "medecin"]
    patient_rows = [r for r in rows if r.questionnaire_type == "patient"]
    medecin_avant = [r for r in medecin_rows if r.timepoint == "T0"]
    medecin_apres = [r for r in medecin_rows if r.timepoint == "T1"]

    wb = Workbook()
    wb.remove(wb.active)

    # ---------------------------------------------------------------
    # FEUILLE 1 — RÉSUMÉ
    # ---------------------------------------------------------------
    ws = wb.create_sheet("Résumé")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Étude clinique ANALOG — Tableau de bord"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Généré le {datetime.datetime.now().strftime('%d/%m/%Y à %H:%M')}"
    ws["A2"].font = SUBTITLE_FONT

    _kpi_card(ws, 4, 1, "Réponses médecin — Avant (T0)", len(medecin_avant))
    _kpi_card(ws, 4, 3, "Réponses médecin — Après (T1)", len(medecin_apres))
    _kpi_card(ws, 4, 5, "Réponses patient", len(patient_rows))

    # Scores moyens médecins (sur les paires avant/après complètes uniquement)
    pair_scores = []
    avant_by_med = {r.medecin_id: json.loads(r.answers_json) for r in medecin_avant if r.medecin_id}
    apres_by_med = {r.medecin_id: json.loads(r.answers_json) for r in medecin_apres if r.medecin_id}
    for mid in set(avant_by_med) & set(apres_by_med):
        a, b = avant_by_med[mid], apres_by_med[mid]
        pair_scores.append((
            scoring.sus_score(a, "av_"), scoring.sus_score(b, "ap_"),
            scoring.nasa_tlx_score(a, "av_"), scoring.nasa_tlx_score(b, "ap_"),
            scoring.bloc3_score(a, "av_"), scoring.bloc3_score(b, "ap_"),
        ))

    def _avg(vals):
        vals = [v for v in vals if v is not None]
        return round(sum(vals) / len(vals), 1) if vals else "—"

    if pair_scores:
        sus_a = _avg([p[0] for p in pair_scores]); sus_b = _avg([p[1] for p in pair_scores])
        nasa_a = _avg([p[2] for p in pair_scores]); nasa_b = _avg([p[3] for p in pair_scores])
        b3_a = _avg([p[4] for p in pair_scores]); b3_b = _avg([p[5] for p in pair_scores])
    else:
        sus_a = sus_b = nasa_a = nasa_b = b3_a = b3_b = "—"

    _kpi_card(ws, 7, 1, "SUS moyen — Avant / Après (/100)", f"{sus_a} → {sus_b}")
    _kpi_card(ws, 7, 3, "NASA-TLX moyen — Avant / Après (/50, ↓ = mieux)", f"{nasa_a} → {nasa_b}")
    _kpi_card(ws, 7, 5, "Bloc 3 moyen — Avant / Après (/100)", f"{b3_a} → {b3_b}")

    patient_scores = []
    for r in patient_rows:
        s, _ = scoring.patient_score(json.loads(r.answers_json))
        if s is not None:
            patient_scores.append(s)
    _kpi_card(ws, 10, 1, "Satisfaction patient moyenne (/100)", _avg(patient_scores))
    _kpi_card(ws, 10, 3, "Nb médecins ayant répondu (identifiants distincts)", len(set(avant_by_med) | set(apres_by_med)))
    _kpi_card(ws, 10, 5, "Nb paires Avant/Après complètes", len(pair_scores))

    ws["A13"] = "Cet onglet se recalcule automatiquement à chaque nouvel export — retélécharger le fichier pour une vue à jour."
    ws["A13"].font = SUBTITLE_FONT
    for col_letter, width in zip("ABCDEF", [4, 32, 4, 32, 4, 32]):
        ws.column_dimensions[col_letter].width = width

    # ---------------------------------------------------------------
    # FEUILLE 2 — SCORES MÉDECIN (une ligne par médecin, avant/après/delta)
    # ---------------------------------------------------------------
    ws = wb.create_sheet("Scores médecin")
    headers = [
        "ID médecin", "Centre",
        "SUS — Avant (/100)", "SUS — Après (/100)", "Δ SUS",
        "NASA-TLX — Avant (/50)", "NASA-TLX — Après (/50)", "Δ NASA-TLX (négatif = charge réduite)",
        "Bloc 3 — Avant (/100)", "Bloc 3 — Après (/100)", "Δ Bloc 3",
        "Incident déclaré", "Risque patient signalé",
    ]
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))

    by_med_tp = {}
    for r in medecin_rows:
        if not r.medecin_id:
            continue
        by_med_tp.setdefault(r.medecin_id, {})[r.timepoint] = (json.loads(r.answers_json), r.centre)

    n_data_rows = 0
    for mid, tp_map in sorted(by_med_tp.items()):
        a, centre_a = tp_map.get("T0", ({}, None))
        b, centre_b = tp_map.get("T1", ({}, None))
        sus_a_v, sus_b_v = scoring.sus_score(a, "av_"), scoring.sus_score(b, "ap_")
        nasa_a_v, nasa_b_v = scoring.nasa_tlx_score(a, "av_"), scoring.nasa_tlx_score(b, "ap_")
        bloc3_a_v, bloc3_b_v = scoring.bloc3_score(a, "av_"), scoring.bloc3_score(b, "ap_")
        ws.append([
            mid, centre_a or centre_b,
            sus_a_v, sus_b_v, (round(sus_b_v - sus_a_v, 1) if sus_a_v is not None and sus_b_v is not None else None),
            nasa_a_v, nasa_b_v, (nasa_b_v - nasa_a_v if nasa_a_v is not None and nasa_b_v is not None else None),
            bloc3_a_v, bloc3_b_v, (bloc3_b_v - bloc3_a_v if bloc3_a_v is not None and bloc3_b_v is not None else None),
            {"oui": "Oui", "non": "Non"}.get(b.get("ap_inc1"), b.get("ap_inc1")),
            {"oui": "Oui", "non": "Non"}.get(b.get("ap_inc2"), b.get("ap_inc2")),
        ])
        n_data_rows += 1
    _add_table(ws, n_data_rows, len(headers), "ScoresMedecin")
    _autosize(ws)

    # ---------------------------------------------------------------
    # FEUILLE 3 — SCORES PATIENT
    # ---------------------------------------------------------------
    ws = wb.create_sheet("Scores patient")
    headers = [
        "Centre", "Âge", "Type d'intervention", "Mode de prise en charge", "Support utilisé",
        "Score global (/100)", "Accès & espace patient (/100)", "Préconsultation (/100)",
        "Fonctionnement technique (/100)", "Impression globale (/100)", "Date de soumission",
    ]
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))
    for r in patient_rows:
        answers = json.loads(r.answers_json)
        score, dims = scoring.patient_score(answers)
        ws.append([
            r.centre, _num(r.age), r.type_intervention, r.mode_prise_charge, r.support, score,
            dims["acces_espace_patient"], dims["preconsultation"], dims["technique"], dims["impression_globale"],
            r.submitted_at.strftime("%d/%m/%Y %H:%M") if r.submitted_at else None,
        ])
    _add_table(ws, len(patient_rows), len(headers), "ScoresPatient")
    _autosize(ws)

    # ---------------------------------------------------------------
    # FEUILLE 4 — RÉPONSES BRUTES MÉDECIN (en-têtes lisibles)
    # ---------------------------------------------------------------
    ws = wb.create_sheet("Réponses brutes médecin")
    all_keys = []
    for r in medecin_rows:
        for k in json.loads(r.answers_json).keys():
            if k not in all_keys:
                all_keys.append(k)
    meta_headers = ["ID médecin", "Moment", "Centre", "Âge", "Ancienneté (années)", "Logiciel actuel", "Date de soumission"]
    headers = meta_headers + [labels.label(k) for k in all_keys]
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))
    for r in medecin_rows:
        answers = json.loads(r.answers_json)
        ws.append([
            r.medecin_id, {"T0": "Avant", "T1": "Après"}.get(r.timepoint, r.timepoint), r.centre,
            _num(r.age), _num(r.anciennete), r.logiciel_actuel,
            r.submitted_at.strftime("%d/%m/%Y %H:%M") if r.submitted_at else None,
        ] + [_num(answers.get(k)) if k not in ("inc1", "ap_inc1", "ap_inc2") else answers.get(k) for k in all_keys])
    _add_table(ws, len(medecin_rows), len(headers), "BrutesMedecin")
    _autosize(ws, max_width=70)

    # ---------------------------------------------------------------
    # FEUILLE 5 — RÉPONSES BRUTES PATIENT (en-têtes lisibles)
    # ---------------------------------------------------------------
    ws = wb.create_sheet("Réponses brutes patient")
    all_keys_p = []
    for r in patient_rows:
        for k in json.loads(r.answers_json).keys():
            if k not in all_keys_p:
                all_keys_p.append(k)
    meta_headers_p = ["Centre", "Âge", "Type d'intervention", "Mode de prise en charge", "Support", "Date de soumission"]
    headers = meta_headers_p + [labels.label(k) for k in all_keys_p]
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))
    for r in patient_rows:
        answers = json.loads(r.answers_json)
        ws.append([
            r.centre, _num(r.age), r.type_intervention, r.mode_prise_charge, r.support,
            r.submitted_at.strftime("%d/%m/%Y %H:%M") if r.submitted_at else None,
        ] + [_num(answers.get(k)) if k != "support" else answers.get(k) for k in all_keys_p])
    _add_table(ws, len(patient_rows), len(headers), "BrutesPatient")
    _autosize(ws, max_width=70)

    wb.move_sheet("Résumé", offset=-len(wb.sheetnames))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"ANALOG_export_{datetime.date.today().isoformat()}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )
