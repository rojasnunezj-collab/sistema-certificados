import streamlit as st
import pandas as pd
import google.generativeai as genai
import json
import io
import base64
import re
import time
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from docxtpl import DocxTemplate

# ==========================================
# 1. CONFIGURACIÓN
# ==========================================
st.set_page_config(page_title="Sistema Certificados", layout="wide")

# --- GESTIÓN DE API KEY ---
API_KEY = None
try:
    if "GEMINI_API_KEY" in st.secrets:
        API_KEY = st.secrets["GEMINI_API_KEY"]
except: pass

if not API_KEY:
    try:
        if "gcp_service_account" in st.secrets and "GEMINI_API_KEY" in st.secrets["gcp_service_account"]:
            API_KEY = st.secrets["gcp_service_account"]["GEMINI_API_KEY"]
    except: pass

if not API_KEY:
    API_KEY = "FALTA_CONFIGURAR"

ID_SHEET_REPOSITORIO = "14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU"
ID_SHEET_CONTROL = "14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU" 

PLANTILLAS = {
    "EPMI S.A.C.": {
        "Comercialización/Disposición Final": "1d09vmlBlW_4yjrrz5M1XM8WpCvzTI4f11pERDbxFvNE",
        "Peligroso y No Peligroso": "1QqqVJ2vCiAjiKKGt_zEpaImUB-q3aRurSiXjMEU--eg"
    },
    "INECOVE S.A.C.": {
        "Comercialización/Disposición Final": "1MPzCwxR538osP3_br4VrTDybplqpTBtB08Jo",
        "Peligroso y No Peligroso": "1W-HyVSivqug13gBRBclBuICAOSBUHm1WN5cnqtMQcZY"
    }
}

# ==========================================
# 2. FUNCIONES
# ==========================================
def obtener_servicios():
    scopes = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
    creds = None
    try:
        if "gcp_service_account" in st.secrets:
            creds = service_account.Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    except: pass

    if not creds:
        try:
            creds = service_account.Credentials.from_service_account_file('secretos.json', scopes=scopes)
        except: return None, None

    try:
        return build('drive', 'v3', credentials=creds), build('sheets', 'v4', credentials=creds)
    except Exception as e:
        st.error(f"Error Google: {e}")
        return None, None

def registrar_en_control(datos_fila):
    _, sheets = obtener_servicios()
    if not sheets: return False
    try:
        sheets.spreadsheets().values().append(
            spreadsheetId=ID_SHEET_CONTROL, range="'historial'!A:J",
            valueInputOption="USER_ENTERED", body={"values": [datos_fila]}
        ).execute()
        return True
    except Exception as e:
        st.error(f"Error Excel: {e}")
        return False

# --- FORMATOS ---
def obtener_fin_de_mes(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, "%d/%m/%Y")
        next_month = dt.replace(day=28) + timedelta(days=4)
        res = next_month - timedelta(days=next_month.day)
        return res.strftime("%d/%m/%Y")
    except: return fecha_str

def limpiar_descripcion(texto):
    if not texto: return ""
    return re.sub(r'VEN\s*-\s*AMB\s*-\s*', '', str(texto).strip(), flags=re.IGNORECASE).strip()

def formato_nompropio(texto):
    return str(texto).strip().title() if texto else ""

def normalizar_fecha(fecha_str):
    if not fecha_str: return datetime.now().strftime("%d/%m/%Y")
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
        try: return datetime.strptime(fecha_str.strip(), fmt).strftime("%d/%m/%Y")
        except: continue
    return fecha_str 

def formatear_guia(serie_str):
    if not serie_str or '-' not in str(serie_str): return serie_str
    try:
        p = str(serie_str).split('-')
        if len(p) == 2: return f"{p[0].strip()}-{str(int(p[1].strip()))}"
    except: pass
    return serie_str

@st.cache_data(show_spinner=False, ttl=10)
def leer_sheet_seguro(pestaña):
    _, s = obtener_servicios()
    if not s: return pd.DataFrame()
    try:
        r = s.spreadsheets().values().get(spreadsheetId=ID_SHEET_REPOSITORIO, range=f"'{pestaña}'!A1:Z1000").execute()
        v = r.get('values', [])
        if not v: return pd.DataFrame()
        return pd.DataFrame(v[1:], columns=v[0])
    except: return pd.DataFrame()

def procesar_guia_ia(pdf_bytes):
    try:
        if "FALTA" in API_KEY:
            st.error("⚠️ Falta API Key")
            return None
        genai.configure(api_key=API_KEY.strip())
    except: return None

    # === RASTREADOR DE MODELOS INTELIGENTE ===
    model = None
    lista_modelos_visibles = []
    
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                lista_modelos_visibles.append(m.name)
        
        candidato = next((m for m in lista_modelos_visibles if 'flash' in m and '1.5' in m), None)
        if not candidato:
            candidato = next((m for m in lista_modelos_visibles if 'flash' in m), None)
        if not candidato:
            candidato = next((m for m in lista_modelos_visibles if 'pro' in m and '1.5' in m), None)

        if candidato:
            model = genai.GenerativeModel(candidato)
        else:
            st.warning(f"⚠️ No encontré modelos 'Flash'. Disponibles: {lista_modelos_visibles}")
            return None

    except Exception as e:
        st.error(f"❌ Error buscando modelos: {e}")
        return None

    # === INSTRUCCIONES DETALLADAS ===
    prompt = """
    Actúa como un experto en extracción de datos OCR. Analiza este documento (Guía de Remisión) y extrae la siguiente información en formato JSON estricto.

    ESTRUCTURA JSON REQUERIDA:
    {
        "fecha": "dd/mm/yyyy", 
        "serie": "T001-000000", 
        "vehiculo": "PLACA (ej: B2F-837)", 
        "punto_partida": "Dirección completa de partida", 
        "punto_llegada": "Dirección completa de llegada", 
        "destinatario": "Razón Social del Destinatario", 
        "items": [
            {
                "desc": "Descripción detallada del bien", 
                "cant": "Número exacto (ej: 749.00)", 
                "um": "Unidad de medida (UND, KG, NIU, etc)", 
                "peso": "Peso total numérico (ej: 500.00)"
            }
        ]
    }

    REGLAS CRÍTICAS DE EXTRACCIÓN:
    1. **TABLA DE ITEMS:**
       - Busca columnas como "Descripción", "Cantidad", "U.M.", "Peso Total" o "Peso".
       - **¡IMPORTANTE!** Diferencia entre CANTIDAD (bultos) y PESO (kg). Si solo hay una cifra numérica junto a la unidad, asígnala al campo más lógico. Si ves 'KGM' o 'KG', es peso. Si ves 'UND' o 'NIU', es cantidad.
       - Si el peso está vacío o es 0, busca si está en otra columna cercana.

    2. **PUNTO DE PARTIDA Y OBSERVACIONES:**
       - Revisa el campo "OBSERVACIONES" al final del documento. 
       - Si menciona un lugar específico como "FUNDO...", "PLANTA...", "POZO...", extrae ESE LUGAR y agrégalo al final de 'punto_partida' separado por " - ".
       - Ignora textos genéricos sobre residuos o devolución de envases en las observaciones.

    3. **FECHA Y SERIE:**
       - La fecha suele estar arriba a la derecha o en 'Fecha de Emisión'.
       - La serie tiene formato XXXX-XXXXXXX (ej: EG07-0004331).

    Responde SOLO con el JSON.
    """
