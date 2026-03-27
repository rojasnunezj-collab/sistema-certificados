# ====================================================================
# --- BLOQUE 0: Imports ---
# ====================================================================
import streamlit as st
import io
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from src.config.settings import ID_SHEET_CONTROL, DRIVE_FOLDER_ID, ID_SHEET_REPOSITORIO

# ====================================================================
# --- BLOQUE 1: Lectura Segura de Google Sheets (Caché Opcional) ---
# ====================================================================
@st.cache_data(show_spinner=False, ttl=10)
def leer_sheet_seguro(pestaña):
    """Lectura segura de Google Sheets con cacheo parcial"""
    _, s = obtener_servicios()
    if not s: return pd.DataFrame()
    try:
        r = s.spreadsheets().values().get(spreadsheetId=ID_SHEET_REPOSITORIO, range=f"'{pestaña}'!A1:Z1000").execute()
        v = r.get('values', [])
        if not v: return pd.DataFrame()
        return pd.DataFrame(v[1:], columns=v[0])
    except Exception as e:
        st.error(f"Error leyendo pestaña {pestaña}: {e}")
        return pd.DataFrame()

# ====================================================================
# --- BLOQUE 2: Flujo de Autenticación y Obtención de Credenciales ---
# ====================================================================
def obtener_servicios():
    import os
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from google.oauth2 import service_account

    scopes = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
    creds = None
    
    # ================================================================
    # SOLUCIÓN CUENTAS PERSONALES: Flujo OAuth 2.0 (token.json)
    # Autentica la app como tu usuario de Gmail, usando tus 15GB libres
    # ================================================================
    if os.path.exists('token.json'):
        try:
            creds = Credentials.from_authorized_user_file('token.json', scopes)
        except: pass
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        elif os.path.exists('credentials.json'):
            # Abre el navegador local para otorgar permiso (solo la 1ra vez)
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', scopes)
            creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

    # Fallback 1: Identidad de Service Account (Limitado a 0 bytes sin Workspace)
    if not creds:
        cred_file = next((p for p in ["secretoslocal.json", "secretos_local.json", "secretos.json"] if os.path.exists(p)), None)
        if cred_file:
            try:
                creds = service_account.Credentials.from_service_account_file(cred_file, scopes=scopes)
            except Exception as e:
                st.warning(f"No se pudo cargar {cred_file}: {e}")

    # Fallback 2: Streamlit Secrets
    if not creds:
        try:
            if "gcp_service_account" in st.secrets:
                info = dict(st.secrets["gcp_service_account"])
                # FIX: Auto-correcion de typos en secrets (ej: https=// -> https://)
                for k, v in info.items():
                    if isinstance(v, str):
                        info[k] = v.replace("https=//", "https://")
                creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
        except Exception:
            pass

    try:
        return build('drive', 'v3', credentials=creds), build('sheets', 'v4', credentials=creds)
    except Exception as e:
        st.error(f"⚠️ Error conectando con Google API: {e}")
        return None, None

# ====================================================================
# --- BLOQUE 3: Funciones de Escritura y Subida (Sheets y Drive) ---
# ====================================================================
def registrar_en_control(datos_fila):
    _, sheets = obtener_servicios()
    if not sheets: return False
    try:
        sheets.spreadsheets().values().append(
            spreadsheetId=ID_SHEET_CONTROL, range="'historial'!A:J",
            valueInputOption="USER_ENTERED", 
            insertDataOption="INSERT_ROWS",
            body={"values": [datos_fila]}
        ).execute()
        return True
    except: return False

def subir_a_drive(contenido_bytes, nombre_archivo, tipo_flujo):
    """Sube el archivo a Drive enrutándolo a la carpeta correcta según el tipo"""
    
    # ENRUTADOR DINÁMICO DE CARPETAS
    if "Comercialización" in tipo_flujo:
        folder_id = "1NZc-nfGHw5bnkCAv0TdQYW_bPM_UkKC-" # Carpeta Comercialización
    else:
        folder_id = "12PMJ1d-CSWo64m7aNQRQj2yGHFdp9B9S" # Carpeta Servicios (Disp. Final)

    drive, _ = obtener_servicios()
    if not drive: return None
    
    try:
        file_metadata = {
            'name': f"{nombre_archivo}.docx", 
            'mimeType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'parents': [folder_id]
        }
        
        media = MediaIoBaseUpload(
            io.BytesIO(contenido_bytes), 
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
            resumable=True
        )
        # FIX: supportsAllDrives=True delega la cuota a la Unidad Compartida (Shared Drive)
        file = drive.files().create(
            body=file_metadata, 
            media_body=media, 
            fields='id, webViewLink',
            supportsAllDrives=True
        ).execute()
        return file.get('webViewLink')
    except Exception as e:
        st.error(f"Error subiendo a Drive: {e}")
        return None