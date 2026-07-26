"""
Calcul des scores — réplique exacte de la logique JS embarquée dans les 2
questionnaires HTML (medecin.html, patient.html), pour que l'export Excel
donne les mêmes chiffres que ce que le médecin/patient voit en direct.

Le questionnaire médecin est une page unique avec bascule Avant/Après :
les champs sont préfixés av_ (avant, T0) ou ap_ (après, T1) pour ne jamais
se mélanger dans une même soumission. `prefix` sélectionne lequel lire.
"""


def _get(answers, key):
    v = answers.get(key)
    if v in (None, "", "null"):
        return None
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def sus_score(answers, prefix=""):
    """SUS standard (10 items) -> score /100."""
    total, n = 0, 0
    for i in range(1, 11):
        v = _get(answers, f"{prefix}s{i}")
        if v is None:
            continue
        n += 1
        total += (v - 1) if i % 2 == 1 else (5 - v)
    if n == 0:
        return None
    # reproduit le JS a l'identique -> pas de normalisation si <10 items repondus
    return round(total * 2.5, 1)


def nasa_tlx_score(answers, prefix=""):
    """NASA-TLX simplifie (5 items) -> score /50 (brut, sans normalisation)."""
    total, n = 0, 0
    for i in range(1, 6):
        v = _get(answers, f"{prefix}n{i}")
        if v is None:
            continue
        n += 1
        total += (10 - v) if i in (1, 2, 5) else v
    if n == 0:
        return None
    return total


def bloc3_score(answers, prefix="", n_items=18):
    """Bloc 3 (18 items) -> score /100."""
    total, n = 0, 0
    for i in range(1, n_items + 1):
        v = _get(answers, f"{prefix}c{i}")
        if v is None:
            continue
        n += 1
        total += v
    return round((total / (n * 5)) * 100) if n else None


def patient_score(answers, n_items=14):
    """Questionnaire patient (p1..p14) -> score /100 + sous-scores par dimension."""
    total, n = 0, 0
    for i in range(1, n_items + 1):
        v = _get(answers, f"p{i}")
        if v is None:
            continue
        n += 1
        total += v
    global_score = round((total / (n * 5)) * 100) if n else None

    dims = {}
    for label, (start, end) in {
        "acces_espace_patient": (1, 4),
        "preconsultation": (5, 9),
        "technique": (10, 11),
        "impression_globale": (12, 14),
    }.items():
        t, n2 = 0, 0
        for i in range(start, end + 1):
            v = _get(answers, f"p{i}")
            if v is None:
                continue
            n2 += 1
            t += v
        dims[label] = round((t / (n2 * 5)) * 100) if n2 else None
    return global_score, dims
