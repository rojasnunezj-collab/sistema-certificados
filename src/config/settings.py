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
ID_SHEET_SIGERSOL = os.getenv("ID_SHEET_SIGERSOL", "10fES_h_4c8AdwaAAa46nuhTwbhNoifhCMzagCPI_4HI")
ID_SHEET_PETRAMAS = os.getenv("ID_SHEET_PETRAMAS", ID_SHEET_CONTROL)

# Hojas especializadas en Google Sheets para Petramás
SHEET_CERTIFICADOS_PETRAMAS = "Certificados_Petramas"
SHEET_SUSTENTOS_CAMPO = "Sustentos_Campo"

# Estructura obligatoria de carpetas Google Drive para Petramás
CARPETAS_PETRAMAS = {
    "root": os.getenv("DRIVE_PETRAMAS_ROOT", "1hBZhJ5DX-nri8l79-KrVqxoCXuxAItNu"),
    "certificados_root": "1fkHtb0BcGgHlE6PXbrYIEOdE267RPlxW",
    "sustentos_root": "16pBskzUWjrusbEsr_jvyR2mqKVZN1awY",
    "expedientes_root": "1OBBSMNLslYowxBSc8TLvdH_Zhmy9vRVF",
    "certificados": {
        "EPMI S.A.C.": "1qn4ASHsBzj8vJKc2JVoXH-qJKCh9a6ku",
        "INECOVE S.A.C.": "14Jz6yiTjqp-jTG2W_OKbsnR_zE6v7fcc",
        "TRANSPORT & SERVICE ALHUAY S.A.C.": "1cdGN3HtIhyDuHc4yodZew0UdDYv2Qn89"
    },
    "sustentos": {
        "EPMI S.A.C.": "1CMe0hTVCN5kfNxbJ0575Br6AnRFKJ0R_",
        "INECOVE S.A.C.": "1OZGDSpMcmbTp1ySWJumMJdXToaUNOysB",
        "TRANSPORT & SERVICE ALHUAY S.A.C.": "134k-n-yXMm-VFdxViZIgkDEV7M0VIz90"
    },
    "expedientes": {
        "EPMI S.A.C.": "1f4QTNEeN6lC31NYAe9cLaE2g91-gU3H1",
        "INECOVE S.A.C.": "1xdPgXvI7cwgl9tKaR7gzKPBorlblQj_H",
        "TRANSPORT & SERVICE ALHUAY S.A.C.": "1mJwsQWRSAWBVHt_CDUPPz6Sh8oBBUVaY"
    }
}

# Configuración de las 3 Empresas Operadoras (EO-RS)
EMPRESAS_EORS = {
    "EPMI S.A.C.": {
        "ruc": "20601955938",
        "hoja_historial": "Historial",
        "prefijo_cert": "CERT-SER",
        "alias": ["EPMI", "EPMI S.A.C.", "EMPRESA EPMI S.A.C."]
    },
    "INECOVE S.A.C.": {
        "ruc": "20607873330",
        "hoja_historial": "Historial_INECOVE",
        "prefijo_cert": "CERT-SER",
        "alias": ["INECOVE", "INECOVE S.A.C.", "INVERSIONES ECOLOGICAS DEL PERU"]
    },
    "TRANSPORT & SERVICE ALHUAY S.A.C.": {
        "ruc": "20600862596",
        "hoja_historial": "Historial_ALHUAY",
        "prefijo_cert": "CERT-SER",
        "alias": ["ALHUAY", "TRANSPORT & SERVICE ALHUAY", "TRANSPORT & SERVICE ALHUAY S.A.C."]
    }
}

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
    },
    "TRANSPORT & SERVICE ALHUAY S.A.C.": {
        "Comercialización": "1i44hDtMx7kHm_2g1UyXeNo9cdMw_XSks",
        "Disposición Final 1": "1F4NVyV7_EISPYu2eHMWeFcd7ChWEKe7o"
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
    },
    "TRANSPORT & SERVICE ALHUAY S.A.C.": {
        "Disposición Final 1": "1pNvIbFugtCqDxCk0nMgDh3hEHyk75lKgDz_ILU9_ayM"
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

try:
    if not GCP_PROJECT and "gcp_service_account" in st.secrets:
        GCP_PROJECT = st.secrets["gcp_service_account"].get("project_id")
except Exception:
    pass

GCP_LOCATION = os.getenv("GCP_LOCATION", "us-central1")
