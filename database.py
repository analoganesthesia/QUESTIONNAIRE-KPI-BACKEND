"""
Modèle de données — Étude clinique ANALOG.

Une table unique 'submissions' : une ligne = une soumission de questionnaire
(médecin T0, médecin T1, ou patient). Les réponses individuelles (s1, c4, p12...)
sont stockées en JSON dans answers_json — pas besoin de migrer le schéma si un
item de questionnaire change.

SQLite par défaut (fichier local, zéro configuration, largement suffisant pour
100 consultations x 3 questionnaires). Pour passer sur Postgres (recommandé
si le volume de l'étude s'étend), changer uniquement DATABASE_URL.
"""
import os
import datetime
from sqlalchemy import create_engine, Column, Integer, String, DateTime, Text
from sqlalchemy.orm import declarative_base, sessionmaker

DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./study.db")

connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine_kwargs = {"connect_args": connect_args}
if not DATABASE_URL.startswith("sqlite"):
    # pool_pre_ping : teste la connexion avant chaque usage et la régénère si elle est
    # périmée (cas Neon qui met la base en veille par inactivité) -> évite l'échec
    # "au premier essai" suivi d'un succès au second.
    # pool_recycle : force le renouvellement des connexions de toute façon après 280s,
    # avant que Neon/Postgres ne les ferme lui-même côté serveur.
    engine_kwargs["pool_pre_ping"] = True
    engine_kwargs["pool_recycle"] = 280
engine = create_engine(DATABASE_URL, **engine_kwargs)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


class Submission(Base):
    __tablename__ = "submissions"

    id = Column(Integer, primary_key=True, index=True)
    questionnaire_type = Column(String, index=True, nullable=False)  # medecin | patient
    medecin_id = Column(String, index=True, nullable=True)           # auto-déclaré par le médecin (M01...), sert à apparier T0/T1
    timepoint = Column(String, nullable=True)                        # T0 | T1 (médecin uniquement)
    centre = Column(String, nullable=True)
    age = Column(String, nullable=True)
    anciennete = Column(String, nullable=True)                       # médecin
    logiciel_actuel = Column(String, nullable=True)                  # médecin
    type_intervention = Column(String, nullable=True)                # patient
    mode_prise_charge = Column(String, nullable=True)                # patient: ambulatoire / hospitalisation
    support = Column(String, nullable=True)                          # patient: ordinateur/telephone/tablette
    answers_json = Column(Text, nullable=False)                      # dump JSON de toutes les réponses (items s/n/c ou p)
    submitted_at = Column(DateTime, default=datetime.datetime.utcnow)


def init_db():
    Base.metadata.create_all(bind=engine)
