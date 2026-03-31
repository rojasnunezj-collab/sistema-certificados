# ====================================================================
# --- BLOQUE 0: Imports e Inicialización de Entorno ---
# ====================================================================
import os
from dotenv import load_dotenv
import streamlit as st

load_dotenv()

# ====================================================================
# --- BLOQUE 1: IDs Globales (Sheets y Drive Folder) ---
# ====================================================================
# Google IDs
ID_SHEET_REPOSITORIO = "14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU"
ID_SHEET_CONTROL = "14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU"

CARPETAS_DESTINO = {
    "EPMI S.A.C.": {
        "Comercialización": "1NZc-nfGHw5bnkCAv0TdQYW_bPM_UkKC-", # El que me pasaste
        "Disposición Final 1": "12PMJ1d-CSWo64m7aNQRQj2yGHFdp9B9S",
        "Disposición Final 2": "12PMJ1d-CSWo64m7aNQRQj2yGHFdp9B9S"
    },
    "INECOVE S.A.C.": {
        "Comercialización": "1NZc-nfGHw5bnkCAv0TdQYW_bPM_UkKC-",
        "Disposición Final 1": "12PMJ1d-CSWo64m7aNQRQj2yGHFdp9B9S",
        "Disposición Final 2": "12PMJ1d-CSWo64m7aNQRQj2yGHFdp9B9S"
    }
}

# ====================================================================
# --- BLOQUE 2: Diccionario de Plantillas Documentales ---
# ====================================================================
PLANTILLAS = {
    "EPMI S.A.C.": {
        "Comercialización": "1d09vmlBlW_4yjrrz5M1XM8WpCvzTI4f11pERDbxFvNE",
        "Disposición Final 1": "1QqqVJ2vCiAjiKKGt_zEpaImUB-q3aRurSiXjMEU--eg",
        "Disposición Final 2": "1fpdZef3Fe3tl00yAuM0Cehx2_o3AusrErcOJisBtdBM"
    },
    "INECOVE S.A.C.": {
        "Comercialización": os.getenv("TEMPLATE_INECOVE_ID", "1MPzCwxR538osP3_br4VrTDybplqpTBtB08Jo"),
        "Disposición Final 1": os.getenv("TEMPLATE_INECOVE_PELIGROSO_ID", "1W-HyVSivqug13gBRBclBuICAOSBUHm1WN5cnqtMQcZY"),
        "Disposición Final 2": os.getenv("TEMPLATE_INECOVE_PELIGROSO_ID", "1W-HyVSivqug13gBRBclBuICAOSBUHm1WN5cnqtMQcZY")
    }
}

# ====================================================================
# --- BLOQUE 3: Configuración y Fallbacks de Google Cloud (Vertex) ---
# ====================================================================
# Google Cloud Project Info (for Vertex AI)
GCP_PROJECT = os.getenv("GCP_PROJECT_ID")

# Prioridad: secretos.json > secrets (Streamlit) > env
if os.path.exists("secretos.json"):
    import json
    try:
        with open("secretos.json", "r") as f:
            secretos_data = json.load(f)
            if "project_id" in secretos_data:
                GCP_PROJECT = secretos_data["project_id"]
    except: pass

if not GCP_PROJECT and "gcp_service_account" in st.secrets:
    GCP_PROJECT = st.secrets["gcp_service_account"].get("project_id")

GCP_LOCATION = os.getenv("GCP_LOCATION", "us-central1")
