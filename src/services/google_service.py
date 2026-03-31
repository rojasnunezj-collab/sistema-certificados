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
@st.cache_data(show_spinner=False, ttl=600)
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
    from google.oauth2 import service_account

    scopes = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
    creds = None
    
    # 1. Intentar cargar desde Streamlit Secrets (NUBE) - ¡Prioridad #1!
    if "google" in st.secrets:
        try:
            info = dict(st.secrets["google"])
            # Auto-corrección de formato por si acaso
            for k, v in info.items():
                if isinstance(v, str):
                    info[k] = v.replace("https=//", "https://")
            creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
        except Exception as e:
            st.error(f"Error cargando credenciales de secrets en Google Service: {e}")

    # 2. Si no hay secrets, intentar cargar desde archivo local (PC)
    if not creds:
        cred_file = next((p for p in ["secretoslocal.json", "secretos_local.json", "secretos.json"] if os.path.exists(p)), None)
        if cred_file:
            try:
                creds = service_account.Credentials.from_service_account_file(cred_file, scopes=scopes)
            except Exception as e:
                st.warning(f"No se pudo cargar archivo local {cred_file}: {e}")

    # 3. Construir los servicios
    if creds:
        try:
            return build('drive', 'v3', credentials=creds), build('sheets', 'v4', credentials=creds)
        except Exception as e:
            st.error(f"⚠️ Error conectando con Google API: {e}")
            return None, None
    else:
        st.error("No se encontraron credenciales válidas para conectar a Google.")
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

import io
import unicodedata
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

import unicodedata

def subir_a_drive(contenido_bytes, nombre_archivo, tipo_flujo, carpeta_id=None):
    """Sube el archivo a Drive enrutándolo a la carpeta correcta según el tipo"""
    
    # --- ENRUTADOR DINÁMICO ---
    # Si NO le mandamos una carpeta específica desde app.py, usa tu lógica original como respaldo de emergencia
    if not carpeta_id:
        def normalizar(texto):
            texto = str(texto).lower() # Todo a minúsculas
            return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
            
        tipo_seguro = normalizar(tipo_flujo) 
        if "comercializacion" in tipo_seguro:
            carpeta_id = "1NZc-nfGHw5bnkCAv0TdQYW_bPM_UkKC-" # Comercialización
        else:
            carpeta_id = "12PMJ1d-CSWo64m7aNQRQj2yGHFdp9B9S" # Servicios (Disp. Final)

    drive, _ = obtener_servicios()
    if not drive: return None
    
    try:
        file_metadata = {
            'name': f"{nombre_archivo}.docx", 
            'mimeType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'parents': [carpeta_id]  # <--- AQUÍ GUARDARÁ EN LA CARPETA EXACTA
        }
        
        from googleapiclient.http import MediaIoBaseUpload
        import io
        
        media = MediaIoBaseUpload(
            io.BytesIO(contenido_bytes), 
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
            resumable=True
        )
        
        file = drive.files().create(
            body=file_metadata, 
            media_body=media, 
            fields='id, webViewLink',
            supportsAllDrives=True
        ).execute()
        return file.get('webViewLink')
        
    except Exception as e:
        print(f"Error subiendo a Drive: {e}") 
        return None

def obtener_plantilla_drive(empresa_nombre, tipo_certificado, drive_service):
    """Busca la plantilla en blanco en la nueva carpeta de Drive."""
    CARPETA_PLANTILLAS_ID = '1_kY1h6PwlhDPl8BjG7u0fbGMT1AWyn3n' # Carpeta de plantillas en blanco
    
    palabra_flujo = "Comercializacion" if "Comercialización" in tipo_certificado else "Final"
    
    query = (f"'{CARPETA_PLANTILLAS_ID}' in parents "
             f"and name contains '{empresa_nombre}' "
             f"and name contains '{palabra_flujo}' "
             f"and trashed=false")
             
    # CAMBIO 1: Le pedimos a Google que nos devuelva el 'mimeType' (el tipo de archivo)
    resultados = drive_service.files().list(q=query, spaces='drive', fields='files(id, name, mimeType)').execute()
    archivos = resultados.get('files', [])
    
    if not archivos:
        raise Exception(f"No se encontró plantilla en Drive para '{empresa_nombre}' y '{palabra_flujo}'. Revisa la carpeta de plantillas.")
        
    archivo = archivos[0]
    archivo_id = archivo['id']
    mime_type = archivo.get('mimeType', '')
    
    # CAMBIO 2: La bifurcación anti-Error 403
    if 'application/vnd.google-apps.document' in mime_type:
        # Si Google lo convirtió a Google Docs, lo EXPORTAMOS como Word a la fuerza
        request = drive_service.files().export_media(fileId=archivo_id, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    else:
        # Si sigue siendo un Word puro (.docx), lo descargamos normal
        request = drive_service.files().get_media(fileId=archivo_id)
        
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        
    fh.seek(0)
    return fh

def subir_modelo_a_drive(nombre_archivo, contenido_bytes, drive_service):
    """Sube el certificado modelo terminado a su carpeta exclusiva."""
    CARPETA_DESTINO_MODELOS = '1LUErbILxjVHnzuHkdWaeAMI4HnLg1c7E' # Carpeta de modelos terminados
    
    file_metadata = {
        'name': nombre_archivo,
        'parents': [CARPETA_DESTINO_MODELOS]
    }
    media = MediaIoBaseUpload(io.BytesIO(contenido_bytes), 
                              mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
                              resumable=True)
    archivo = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    return archivo.get('webViewLink')

@st.cache_data(ttl=600)
def obtener_mapa_plantillas_drive(es_modelo=False):
    """Escanea Drive y mapea según el modo seleccionado (Normal o Modelo)."""
    from src.services.google_service import obtener_servicios
    drive_service, _ = obtener_servicios()
    
    if not drive_service: 
        return {"EPMI S.A.C.": ["Comercialización", "Disposición Final"]}
    
    # --- EL INTERRUPTOR DE CARPETAS ---
    if es_modelo:
        ID_CARPETA = '1_kY1h6PwlhDPl8BjG7u0fbGMT1AWyn3n' # Carpeta de MODELOS
        modo_texto = "MODO MODELO"
    else:
        ID_CARPETA = '1EwbYAbyv2uMsSn0yXZd0vTPuCoPMbzKs' # Carpeta de NORMALES
        modo_texto = "MODO NORMAL"
    
    query = f"'{ID_CARPETA}' in parents and trashed = false"
    
    try:
        resultados = drive_service.files().list(q=query, fields='files(name, mimeType)').execute()
        archivos = resultados.get('files', [])
        mapa = {}
        
        print(f"\n--- REVISANDO DRIVE ({modo_texto}): {len(archivos)} ARCHIVOS ---")

        for arch in archivos:
            nombre_raw = arch.get('name', '').upper()
            if arch.get('mimeType') == 'application/vnd.google-apps.folder':
                continue

            print(f"-> Analizando: {nombre_raw}")

            # --- LÓGICA DE EMPRESA SEGÚN EL MODO ---
            if es_modelo:
                # En modo modelo, no importa el nombre, la categoría es "MODELO"
                empresa = "MODELO"
            else:
                # En modo normal, buscamos la empresa real
                if "INECOVE" in nombre_raw: empresa = "INECOVE"
                elif "BETA" in nombre_raw: empresa = "BETA"
                else: empresa = "EPMI S.A.C."
            
            if empresa not in mapa:
                mapa[empresa] = []
            
            # DETECCIÓN DE SERVICIO (Igual que antes)
            if "COMERCIALIZACION" in nombre_raw or "COMERCIALIZACIÓN" in nombre_raw:
                mapa[empresa].append("Comercialización")
            if "FINAL" in nombre_raw:
                mapa[empresa].append("Disposición Final")

        for emp in mapa:
            mapa[emp] = sorted(list(set(mapa[emp])))

        print(f"--- MAPA GENERADO: {mapa} ---\n")
        return mapa if mapa else {"EPMI S.A.C.": ["Comercialización"]}

    except Exception as e:
        print(f"🚨 Error al leer Drive: {e}")
        return {"EPMI S.A.C.": ["Comercialización"]}

@st.cache_data(ttl=600) 
def obtener_clientes_desde_sheets():
    """Lee la pestaña 'CLIENTES' del Excel y devuelve un diccionario {Nombre: RUC}."""
    SHEET_ID = '14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU'
    RANGO = 'CLIENTES!A2:B' 
    
    from src.services.google_service import obtener_servicios 
    # LA CORRECCIÓN: Llamamos a los dos robots en el orden correcto
    drive_service, sheet_service = obtener_servicios() 
    
    if not sheet_service: 
        st.error("❌ No hay conexión a los servicios de Google.")
        return {}
    
    try:
        resultado = sheet_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=RANGO).execute()
        filas = resultado.get('values', [])
        
        clientes_dict = {}
        for fila in filas:
            if len(fila) >= 2: 
                nombre = str(fila[0]).strip().upper()
                ruc = str(fila[1]).strip()
                if nombre:
                    clientes_dict[nombre] = ruc
                    
        if not clientes_dict:
            st.warning("⚠️ El bot entró al Excel pero no encontró texto en las columnas A y B.")

        return clientes_dict
    except Exception as e:
        # Si falla, ahora sí lo imprimirá en tu terminal negra para que lo veamos
        st.error(f"🚨 Error real leyendo el Excel: {e}") 
        return {}

@st.cache_data(ttl=600)
def obtener_datos_empresas_desde_sheets():
    """Lee la pestaña 'EMPRESAS' del Excel y devuelve un diccionario con RUC y Registro."""
    SHEET_ID = '14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU'
    RANGO = 'EMPRESAS!A2:C' # A=Empresa, B=RUC, C=Registro
    
    from src.services.google_service import obtener_servicios 
    _, sheet_service = obtener_servicios()
    
    if not sheet_service: return {}
    
    try:
        resultado = sheet_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=RANGO).execute()
        filas = resultado.get('values', [])
        
        datos_dict = {}
        for fila in filas:
            if len(fila) >= 2:
                nombre = str(fila[0]).strip().upper()
                ruc = str(fila[1]).strip()
                # Si hay columna C, la tomamos; si no, ponemos pendiente
                reg = str(fila[2]).strip() if len(fila) >= 3 else "Pendiente"
                
                if nombre:
                    datos_dict[nombre] = {"ruc": ruc, "reg": reg}
                    
        return datos_dict
    except Exception as e:
        print(f"Error leyendo base de datos de empresas: {e}")
        return {}