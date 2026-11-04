# ====================================================================
# --- BLOQUE 0: Imports ---
# ====================================================================
import streamlit as st
import io
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from src.config.settings import ID_SHEET_CONTROL, ID_SHEET_REPOSITORIO

# ====================================================================
# --- BLOQUE 1: Lectura Segura de Google Sheets (Caché Opcional) ---
# ====================================================================
@st.cache_data(show_spinner=False, ttl=60)
def leer_sheet_seguro(pestaña):
    """Lectura segura de Google Sheets con cacheo parcial y normalización robusta de columnas"""
    _, s = obtener_servicios()
    if not s: return pd.DataFrame()
    try:
        r = s.spreadsheets().values().get(spreadsheetId=ID_SHEET_REPOSITORIO, range=f"'{pestaña}'!A1:Z1000").execute()
        v = r.get('values', [])
        if not v or len(v) == 0: return pd.DataFrame()
        raw_headers = v[0]
        num_cols = len(raw_headers)
        if num_cols == 0: return pd.DataFrame()
        
        headers = [str(c).strip() if str(c).strip() else f"Col_{i+1}" for i, c in enumerate(raw_headers)]
        
        # Normalizar cada fila al número exacto de columnas de la cabecera
        filas = []
        for row in v[1:]:
            if len(row) < num_cols:
                filas.append(list(row) + [''] * (num_cols - len(row)))
            else:
                filas.append(list(row[:num_cols]))
                
        return pd.DataFrame(filas, columns=headers)
    except Exception as e:
        print(f"[Aviso] Error leyendo pestaña {pestaña}: {e}")
        return pd.DataFrame()

# ====================================================================
# --- BLOQUE 2: Flujo de Autenticación y Obtención de Credenciales ---
# ====================================================================
def obtener_servicios():
    import os
    from google.oauth2 import service_account

    scopes = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
    creds = None
    
    # 0. Intentar cargar User Credentials (token.json) desde Streamlit Secrets (¡Prioridad Máxima anti-cuota!)
    try:
        if "token" in st.secrets:
            tinfo = dict(st.secrets["token"])
            from google.oauth2.credentials import Credentials
            creds = Credentials(
                token=tinfo.get('token'),
                refresh_token=tinfo.get('refresh_token'),
                token_uri=tinfo.get('token_uri', "https://oauth2.googleapis.com/token"),
                client_id=tinfo.get('client_id'),
                client_secret=tinfo.get('client_secret'),
                scopes=scopes
            )
    except Exception:
        pass

    # 1. Intentar cargar cuenta de servicio desde Streamlit Secrets (NUBE) - ¡Prioridad #2!
    try:
        if not creds:
            for s_key in ["google", "gcp_service_account", "google_service_account", "service_account"]:
                if s_key in st.secrets:
                    info = dict(st.secrets[s_key])
                    # Auto-corrección de formato por si acaso
                    for k, v in info.items():
                        if isinstance(v, str):
                            info[k] = v.replace("https=//", "https://")
                    creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
                    if creds:
                        break
    except Exception:
        pass

    # 2. Si no hay secrets, intentar cargar archivos locales (PC) - Prioridad local
    if not creds:
        base_app_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        
        # Primero buscar token.json local (prioridad local anti-cuota)
        token_candidates = [
            "token.json", 
            os.path.join(base_app_dir, "token.json"),
            os.path.join(os.path.dirname(os.path.dirname(base_app_dir)), "BotGestion", "Bot_lia_guias_notas", "token.json"),
            os.path.join(os.path.dirname(os.path.dirname(base_app_dir)), "SistemaPostmortem", "token.json"),
            os.path.join(os.path.dirname(base_app_dir), "BotGestion", "Bot_lia_guias_notas", "token.json"),
            os.path.join(os.path.dirname(base_app_dir), "SistemaPostmortem", "token.json")
        ]
        token_file = next((p for p in token_candidates if os.path.exists(p)), None)
        if token_file:
            try:
                from google.oauth2.credentials import Credentials
                creds = Credentials.from_authorized_user_file(token_file, scopes)
            except Exception as e:
                st.warning(f"No se pudo cargar token.json ({token_file}): {e}")
        
        # Si no hay token.json, probar con las service accounts locales
        if not creds:
            posibles_cred = [
                "secretoslocal.json", "secretos_local.json", "secretos.json", "credenciales_gcp.json",
                os.path.join(base_app_dir, "secretoslocal.json"),
                os.path.join(base_app_dir, "secretos.json"),
                os.path.join(base_app_dir, "credenciales_gcp.json")
            ]
            cred_file = next((p for p in posibles_cred if os.path.exists(p)), None)
            if cred_file:
                try:
                    creds = service_account.Credentials.from_service_account_file(cred_file, scopes=scopes)
                except Exception as e:
                    st.warning(f"No se pudo cargar cuenta de servicio local {cred_file}: {e}")

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
def registrar_en_control(datos_fila, hoja_destino="Historial"):
    _, sheets = obtener_servicios()
    if not sheets: return False
    try:
        # Asegurar dinámicamente que la pestaña de historial exista
        try:
            meta = sheets.spreadsheets().get(spreadsheetId=ID_SHEET_CONTROL).execute()
            existing_titles = [s['properties']['title'] for s in meta.get('sheets', [])]
            if hoja_destino not in existing_titles:
                sheets.spreadsheets().batchUpdate(
                    spreadsheetId=ID_SHEET_CONTROL,
                    body={'requests': [{'addSheet': {'properties': {'title': hoja_destino}}}]}
                ).execute()
                headers = [['Fecha de emision', 'Empresa', 'Fundo', 'Correlativo', 'Certificado', 'Guia', 'Link Guia', 'Link Documento', 'Link pdf', 'Observaciones', 'Mes']]
                sheets.spreadsheets().values().update(
                    spreadsheetId=ID_SHEET_CONTROL,
                    range=f"'{hoja_destino}'!A1:K1",
                    valueInputOption='USER_ENTERED',
                    body={'values': headers}
                ).execute()
        except Exception as e_sheet:
            print(f"Aviso comprobando pestaña {hoja_destino}: {e_sheet}")

        res = sheets.spreadsheets().values().append(
            spreadsheetId=ID_SHEET_CONTROL, range=f"'{hoja_destino}'!A:J",
            valueInputOption="USER_ENTERED", 
            insertDataOption="INSERT_ROWS",
            body={"values": [datos_fila]}
        ).execute()
        import re
        updates = res.get('updates', {})
        updated_range = updates.get('updatedRange', '')
        m = re.search(r'(\d+)', updated_range.split('!')[-1])
        fila_idx = int(m.group(1)) if m else None
        try:
            leer_sheet_seguro.clear()
        except Exception:
            pass
        return fila_idx if fila_idx else True
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
        import streamlit as st
        st.error(f"Google Drive rechazó la subida. Motivo técnico: {str(e)}")
        print(f"Error subiendo a Drive: {e}") 
        return None

def obtener_plantilla_drive(empresa_nombre, tipo_certificado, drive_service):
    """
    Busca la plantilla en blanco en Drive.
    Si se especifica una empresa real, prioriza su plantilla oficial configurada por ID.
    Si no existe ID o es 'MODELO', busca en la carpeta de plantillas modelo en Drive.
    """
    import io
    from googleapiclient.http import MediaIoBaseDownload
    
    tipo_flujo_limpio = str(tipo_certificado).strip()
    
    # 1. Si es una empresa real, intentar usar su plantilla oficial por ID
    if empresa_nombre and str(empresa_nombre).upper() != "MODELO":
        try:
            cfg_emisor = buscar_config_empresa(empresa_nombre)
            id_p = cfg_emisor.get('plantillas', {}).get(tipo_flujo_limpio)
            if not id_p:
                from src.config.settings import PLANTILLAS
                id_p = PLANTILLAS.get(empresa_nombre, {}).get(tipo_flujo_limpio)
            if not id_p and cfg_emisor.get('plantillas'):
                id_p = next(iter(cfg_emisor['plantillas'].values()))
                
            if id_p:
                f_meta = drive_service.files().get(fileId=id_p, fields='mimeType', supportsAllDrives=True).execute()
                m_type = f_meta.get('mimeType', '')
                if 'application/vnd.google-apps.document' in m_type:
                    req = drive_service.files().export_media(fileId=id_p, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                else:
                    req = drive_service.files().get_media(fileId=id_p, supportsAllDrives=True)
                    
                fh = io.BytesIO()
                dl = MediaIoBaseDownload(fh, req)
                done = False
                while not done: _, done = dl.next_chunk()
                try:
                    from src.utils.document_utils import sanitizar_tags_docx
                    fh = sanitizar_tags_docx(fh)
                except Exception:
                    pass
                fh.seek(0)
                return fh
        except Exception as e_cfg:
            print(f"[Aviso] Fallback a carpeta modelo para {empresa_nombre}: {e_cfg}")

    # 2. Respaldo: Buscar en la carpeta de plantillas de modelos
    CARPETA_PLANTILLAS_ID = '1_kY1h6PwlhDPl8BjG7u0fbGMT1AWyn3n' # Carpeta de plantillas en blanco
    palabra_flujo = "Comercializacion" if "Comercializaci" in tipo_certificado else "Final"
    
    # Búsqueda tentativa por empresa + flujo
    query = (f"'{CARPETA_PLANTILLAS_ID}' in parents "
             f"and name contains '{empresa_nombre}' "
             f"and name contains '{palabra_flujo}' "
             f"and trashed=false")
    resultados = drive_service.files().list(q=query, spaces='drive', fields='files(id, name, mimeType)').execute()
    archivos = resultados.get('files', [])
    
    # Si no hay match específico, buscar la plantilla genérica del flujo
    if not archivos:
        query_fallback = (f"'{CARPETA_PLANTILLAS_ID}' in parents "
                          f"and (name contains '{palabra_flujo}' or name contains 'MODELO') "
                          f"and trashed=false")
        resultados = drive_service.files().list(q=query_fallback, spaces='drive', fields='files(id, name, mimeType)').execute()
        archivos = resultados.get('files', [])
        
    if not archivos:
        raise Exception(f"No se encontró plantilla en Drive para '{empresa_nombre}' y '{palabra_flujo}'. Revisa la carpeta de plantillas.")
        
    archivo = archivos[0]
    archivo_id = archivo['id']
    mime_type = archivo.get('mimeType', '')
    
    if 'application/vnd.google-apps.document' in mime_type:
        request = drive_service.files().export_media(fileId=archivo_id, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    else:
        request = drive_service.files().get_media(fileId=archivo_id)
        
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        
    try:
        from src.utils.document_utils import sanitizar_tags_docx
        fh = sanitizar_tags_docx(fh)
    except Exception:
        pass
    fh.seek(0)
    return fh

def subir_modelo_a_drive(nombre_archivo, contenido_bytes, drive_service):
    """Sube el certificado modelo terminado a su carpeta exclusiva."""
    try:
        CARPETA_DESTINO_MODELOS = '1LUErbILxjVHnzuHkdWaeAMI4HnLg1c7E' # Carpeta de modelos terminados
        
        file_metadata = {
            'name': nombre_archivo,
            'parents': [CARPETA_DESTINO_MODELOS]
        }
        media = MediaIoBaseUpload(io.BytesIO(contenido_bytes), 
                                  mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
                                  resumable=True)
        archivo = drive_service.files().create(
            body=file_metadata, 
            media_body=media, 
            fields='id, webViewLink',
            supportsAllDrives=True
        ).execute()
        return archivo.get('webViewLink')
    except Exception as e:
        import streamlit as st
        st.error(f"Google Drive rechazó la subida del Modelo. Motivo técnico: {str(e)}")
        print(f"Error subiendo modelo a Drive: {e}")
        return None

@st.cache_data(ttl=60)
def obtener_mapa_plantillas_drive(es_modelo=False):
    """
    Escanea la configuración de empresas y mapea los servicios disponibles.
    Permite siempre seleccionar cualquiera de las empresas generadoras reales,
    incluso si se activa el modo de Certificado Modelo.
    """
    mapa_dinamico = {}
    try:
        cfgs = obtener_configuracion_empresas()
        if cfgs:
            for emp, cfg in cfgs.items():
                mapa_dinamico[emp] = cfg.get("servicios", ["Comercialización", "Disposición Final 1", "Disposición Final 2"])
    except Exception as e_cfg:
        print(f"[Aviso] Error leyendo configuracion empresas para mapa: {e_cfg}")

    if not mapa_dinamico:
        try:
            from src.config.settings import PLANTILLAS
            for emp, servs in PLANTILLAS.items():
                mapa_dinamico[emp] = list(servs.keys())
        except Exception:
            pass

    if es_modelo:
        # Añadir opción genérica MODELO si se desea prueba libre
        if "MODELO" not in mapa_dinamico:
            mapa_dinamico["MODELO"] = ["Comercialización", "Disposición Final"]

    return mapa_dinamico if mapa_dinamico else {"EPMI S.A.C.": ["Comercialización"]}

@st.cache_data(ttl=600) 
def obtener_clientes_desde_sheets():
    """Lee la pestaña 'CLIENTES' del Excel y devuelve un diccionario {Nombre: RUC}."""
    SHEET_ID = '14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU'
    RANGO = 'CLIENTES!A2:B' 
    
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

# --- FALLBACK SEGURO DE CARPETAS DE GUÍAS ---
CARPETAS_GUIAS_DEFAULT = {
    "EPMI S.A.C.": {"hechas": "", "recibidas": ""},
    "INECOVE S.A.C.": {"hechas": "", "recibidas": ""},
    "TRANSPORT & SERVICE ALHUAY S.A.C.": {
        "hechas": "1iYLDP6DE4e5bPCtpJ4Bgb16XKVvT4yCk",
        "recibidas": "1WSyZhKo7XqcHu-Zq4emkEhGTQ63EYEBp"
    },
    "P&P INVERSIONES COMERCIALES Y AMBIENTALES S.A.C.": {
        "hechas": "1R_4asjtOQG3Bxgbtx1SJjWYnNQBweovM",
        "recibidas": "1vPvVcNHb52rq7xi-nMK0_dkG10YaOTB9"
    }
}

def obtener_dict_carpetas_guias():
    try:
        import src.config.settings as _stg
        if not hasattr(_stg, 'CARPETAS_GUIAS'):
            import importlib
            _stg = importlib.reload(_stg)
        return getattr(_stg, 'CARPETAS_GUIAS', CARPETAS_GUIAS_DEFAULT)
    except Exception:
        return CARPETAS_GUIAS_DEFAULT

@st.cache_data(ttl=300)
def obtener_configuracion_empresas():
    """
    Lee dinámicamente la pestaña 'EMPRESAS' del Excel configurador.
    Extrae para cada empresa: RUC, Registro MINAM, Hoja de Historial independiente,
    Carpetas de Destino (CARPETA_...) y Plantillas (PLANTILLA_...).
    Permite agregar nuevas empresas y carpetas directamente desde el Sheet sin tocar código.
    """
    from src.config.settings import PLANTILLAS, CARPETAS_DESTINO

    config_empresas = {}
    try:
        df = leer_sheet_seguro("EMPRESAS")
        if df.empty:
            for emp_nombre in PLANTILLAS.keys():
                config_empresas[emp_nombre] = {
                    "nombre": emp_nombre,
                    "ruc": "",
                    "reg": "",
                    "hoja_historial": "Historial",
                    "carpetas": CARPETAS_DESTINO.get(emp_nombre, {}),
                    "plantillas": PLANTILLAS.get(emp_nombre, {}),
                    "servicios": list(PLANTILLAS.get(emp_nombre, {}).keys())
                }
            return config_empresas

        for _, row in df.iterrows():
            emp_nombre = str(row.get('EMPRESA', '')).strip()
            if not emp_nombre or emp_nombre.lower() == 'nan':
                continue
            ruc = str(row.get('RUC', '')).strip() if str(row.get('RUC', '')).strip().lower() != 'nan' else ""
            reg = str(row.get('REGISTRO', '')).strip() if str(row.get('REGISTRO', '')).strip().lower() != 'nan' else ""
            hoja_hist = str(row.get('HOJA_HISTORIAL', '')).strip()
            if not hoja_hist or hoja_hist.lower() == 'nan':
                emp_clean_token = emp_nombre.replace(" S.A.C.", "").replace(" SAC", "").replace(".", "").strip().split()[0]
                hoja_hist = f"Historial_{emp_clean_token}" if "EPMI" not in emp_nombre.upper() else "Historial"

            carpetas = {}
            carpetas_guias = {'hechas': '', 'recibidas': ''}
            plantillas = {}
            id_sheet_sigersol = ""

            for col in df.columns:
                col_clean = str(col).strip().upper()
                val = str(row.get(col, '')).strip()
                if not val or val.lower() == 'nan':
                    continue
                clean_id = extraer_id_drive(val)
                if col_clean.startswith('CARPETA_'):
                    op_key = col_clean.replace('CARPETA_', '')
                    if 'HECHAS' in op_key:
                        carpetas_guias['hechas'] = clean_id
                    elif 'RECIBIDAS' in op_key:
                        carpetas_guias['recibidas'] = clean_id
                    elif 'COMERCIALIZACI' in op_key:
                        carpetas['Comercialización'] = clean_id
                    elif 'FINAL_1' in op_key:
                        carpetas['Disposición Final 1'] = clean_id
                    elif 'FINAL_2' in op_key:
                        carpetas['Disposición Final 2'] = clean_id
                    elif 'FINAL' in op_key:
                        carpetas['Disposición Final 1'] = clean_id
                        carpetas['Disposición Final 2'] = clean_id
                    else:
                        carpetas[op_key.capitalize()] = clean_id
                elif col_clean.startswith('PLANTILLA_'):
                    op_key = col_clean.replace('PLANTILLA_', '')
                    if 'COMERCIALIZACI' in op_key:
                        plantillas['Comercialización'] = clean_id
                    elif 'FINAL_1' in op_key:
                        plantillas['Disposición Final 1'] = clean_id
                    elif 'FINAL_2' in op_key:
                        plantillas['Disposición Final 2'] = clean_id
                    elif 'FINAL' in op_key:
                        plantillas['Disposición Final 1'] = clean_id
                        plantillas['Disposición Final 2'] = clean_id
                    else:
                        plantillas[op_key.capitalize()] = clean_id
                elif 'SIGERSOL' in col_clean:
                    id_sheet_sigersol = clean_id

            # Fallbacks desde settings.py si la empresa no tuvo plantilla o carpeta en el Sheet
            if not carpetas and emp_nombre in CARPETAS_DESTINO:
                carpetas = dict(CARPETAS_DESTINO[emp_nombre])
            if not plantillas and emp_nombre in PLANTILLAS:
                plantillas = dict(PLANTILLAS[emp_nombre])
            if not id_sheet_sigersol:
                from src.config.settings import SHEETS_SIGERSOL, ID_SHEET_SIGERSOL
                id_sheet_sigersol = SHEETS_SIGERSOL.get(emp_nombre, ID_SHEET_SIGERSOL)

            dict_cg = obtener_dict_carpetas_guias()
            cfg_cg = dict_cg.get(emp_nombre, {})
            if not carpetas_guias.get('hechas'):
                carpetas_guias['hechas'] = cfg_cg.get('hechas', '')
            if not carpetas_guias.get('recibidas'):
                carpetas_guias['recibidas'] = cfg_cg.get('recibidas', '')

            # Los servicios ofertados dependen estrictamente de las plantillas existentes.
            # Si solo tiene una plantilla, solo saldrá esa plantilla. Si tiene más, saldrán todas.
            servicios = [s for s, pid in plantillas.items() if pid and str(pid).strip() and str(pid).strip().lower() != 'nan']
            if not servicios:
                servicios = [s for s, cid in carpetas.items() if cid and str(cid).strip() and str(cid).strip().lower() != 'nan']
            if not servicios:
                servicios = ['Disposición Final 1']

            config_empresas[emp_nombre] = {
                'nombre': emp_nombre,
                'ruc': ruc,
                'reg': reg,
                'hoja_historial': hoja_hist,
                'carpetas': carpetas,
                'carpetas_guias': carpetas_guias,
                'plantillas': plantillas,
                'servicios': servicios,
                'id_sheet_sigersol': id_sheet_sigersol
            }

        return config_empresas
    except Exception as e:
        print(f"Error parseando configuración de empresas: {e}")
        return {}

def buscar_config_empresa(empresa_nombre, config_empresas=None):
    """Busca y retorna la configuración de una empresa emisora con soporte para coincidencias flexibles."""
    if not config_empresas:
        config_empresas = obtener_configuracion_empresas()
    if not empresa_nombre:
        return next(iter(config_empresas.values())) if config_empresas else {}
    
    emp_target = str(empresa_nombre).strip().upper()
    # 1. Búsqueda exacta
    for emp_k, cfg in config_empresas.items():
        if emp_k.strip().upper() == emp_target:
            return cfg
            
    # 2. Búsqueda normalizada (sin S.A.C., SAC, etc.)
    emp_corta = emp_target.replace(" S.A.C.", "").replace(" SAC", "").replace(".", "").strip()
    for emp_k, cfg in config_empresas.items():
        k_corta = emp_k.strip().upper().replace(" S.A.C.", "").replace(" SAC", "").replace(".", "").strip()
        if emp_corta in k_corta or k_corta in emp_corta:
            return cfg
            
    # 3. Fallback al primero o a un dict por defecto
    return next(iter(config_empresas.values())) if config_empresas else {}

def obtener_sheet_sigersol_destino(emisor_nombre=None, cliente_nombre=None) -> str:
    """
    Resuelve el ID de Google Sheet de destino para SIGERSOL.
    Prioridad:
    1. Si el cliente tiene un sheet configurado en la pestaña 'CLIENTES' (columna SIGERSOL)
    2. Si la empresa emisora tiene un sheet configurado en la pestaña 'EMPRESAS' (columna SIGERSOL)
    3. Mapeo fallback en settings.py (SHEETS_SIGERSOL)
    4. Fallback global ID_SHEET_SIGERSOL
    """
    from src.config.settings import ID_SHEET_SIGERSOL, SHEETS_SIGERSOL

    # 1. Verificar si en la pestaña CLIENTES existe una columna SIGERSOL
    if cliente_nombre:
        try:
            df_clientes = leer_sheet_seguro("CLIENTES")
            if not df_clientes.empty and 'EMPRESA' in df_clientes.columns:
                col_sig = next((c for c in df_clientes.columns if 'SIGERSOL' in str(c).upper()), None)
                if col_sig:
                    cli_target = str(cliente_nombre).strip().upper()
                    for _, r in df_clientes.iterrows():
                        r_emp = str(r.get('EMPRESA', '')).strip().upper()
                        if r_emp == cli_target or (r_emp and (r_emp in cli_target or cli_target in r_emp)):
                            v_link = str(r.get(col_sig, '')).strip()
                            if v_link and v_link.lower() != 'nan':
                                cid = extraer_id_drive(v_link)
                                if cid:
                                    return cid
        except Exception as e:
            print(f"[Aviso] No se pudo leer SIGERSOL desde pestaña CLIENTES: {e}")

    # 2. Verificar configuración de empresa emisora
    if emisor_nombre:
        cfg = buscar_config_empresa(emisor_nombre)
        if cfg and cfg.get('id_sheet_sigersol'):
            return cfg['id_sheet_sigersol']
        
        # Búsqueda en SHEETS_SIGERSOL con normalización
        em_upper = str(emisor_nombre).strip().upper()
        for k_emp, sid in SHEETS_SIGERSOL.items():
            k_upper = k_emp.upper()
            if k_upper in em_upper or em_upper in k_upper or any(token in em_upper for token in k_upper.split() if len(token) > 3):
                return sid

    return ID_SHEET_SIGERSOL

def obtener_carpeta_guias(empresa_nombre, tipo_guia="hechas") -> str:
    """
    Obtiene el ID de la carpeta de Drive para guías ('hechas' o 'recibidas') según la empresa emisora.
    Prioridad:
    1. Configuración dinámica en la pestaña 'EMPRESAS' de Google Sheets (columnas CARPETA_GUIAS_HECHAS / CARPETA_GUIAS_RECIBIDAS)
    2. Diccionario de respaldo CARPETAS_GUIAS en src/config/settings.py
    """
    tipo_key = "recibidas" if "recib" in str(tipo_guia).lower() else "hechas"
    if empresa_nombre:
        cfg = buscar_config_empresa(empresa_nombre)
        if cfg and cfg.get('carpetas_guias'):
            cid = cfg['carpetas_guias'].get(tipo_key)
            if cid:
                return cid

    # Fallback directo seguro
    CARPETAS_GUIAS = obtener_dict_carpetas_guias()
    em_target = str(empresa_nombre or "").strip().upper()
    for k_emp, dic_c in CARPETAS_GUIAS.items():
        k_upper = k_emp.upper()
        if k_upper in em_target or em_target in k_upper or any(token in em_target for token in k_upper.split() if len(token) > 3):
            cid = dic_c.get(tipo_key)
            if cid:
                return cid
    return ""

def guardar_guia_en_drive(contenido_bytes, nombre_archivo, empresa_nombre=None, tipo_guia="recibidas", carpeta_id=None):
    """
    Sube y guarda una guía (PDF, JPG, PNG, etc.) en Google Drive en la carpeta correspondiente de la empresa ('hechas' o 'recibidas').
    Retorna un diccionario con { 'ok': bool, 'id': file_id, 'link': webViewLink, 'nombre': nombre_archivo } o {'ok': False, 'error': msg}.
    """
    import io
    from googleapiclient.http import MediaIoBaseUpload

    cid = carpeta_id or obtener_carpeta_guias(empresa_nombre, tipo_guia)
    if not cid:
        err_msg = f"No se encontró carpeta de guías configurada para empresa '{empresa_nombre}' (tipo: '{tipo_guia}')."
        print(f"[Aviso] {err_msg}")
        return {'ok': False, 'error': err_msg}

    drive, _ = obtener_servicios()
    if not drive:
        return {'ok': False, 'error': "No se pudo inicializar el servicio de Google Drive."}

    nom_lower = str(nombre_archivo).lower()
    if nom_lower.endswith('.pdf'):
        mimetype = 'application/pdf'
    elif nom_lower.endswith('.png'):
        mimetype = 'image/png'
    elif nom_lower.endswith('.jpg') or nom_lower.endswith('.jpeg'):
        mimetype = 'image/jpeg'
    else:
        mimetype = 'application/pdf' if '.' not in nombre_archivo else 'application/octet-stream'

    nom_final = nombre_archivo if '.' in nombre_archivo else f"{nombre_archivo}.pdf"

    try:
        file_metadata = {
            'name': nom_final,
            'mimeType': mimetype,
            'parents': [cid]
        }
        media = MediaIoBaseUpload(
            io.BytesIO(contenido_bytes),
            mimetype=mimetype,
            resumable=True
        )
        file = drive.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, name, webViewLink',
            supportsAllDrives=True
        ).execute()
        return {
            'ok': True,
            'id': file.get('id'),
            'link': file.get('webViewLink'),
            'nombre': file.get('name')
        }
    except Exception as e:
        err_str = str(e)
        print(f"[Error] Subida de guía a Drive: {err_str}")
        return {'ok': False, 'error': err_str}

@st.cache_data(ttl=600)
def obtener_datos_empresas_desde_sheets():
    """Lee la pestaña 'EMPRESAS' del Excel y devuelve un diccionario con RUC y Registro."""
    cfgs = obtener_configuracion_empresas()
    if cfgs:
        datos_dict = {}
        for emp, c in cfgs.items():
            datos_dict[emp.upper()] = {"ruc": c["ruc"], "reg": c["reg"]}
            emp_corta = emp.replace(" S.A.C.", "").replace(" SAC", "").replace(".", "").strip().upper()
            datos_dict[emp_corta] = {"ruc": c["ruc"], "reg": c["reg"]}
        return datos_dict
    return {}

# ====================================================================
# --- BLOQUE 4: Funciones de Repositorio Masivo ---
# ====================================================================
ID_SHEET_GUIAS = "14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU"

@st.cache_data(show_spinner=False, ttl=60)
def obtener_catalogo_guias(_servicio_sheets):
    """Extrae listado de empresas, meses y fundos anidados para los selectboxes."""
    if not _servicio_sheets: return {}
    try:
        r = _servicio_sheets.spreadsheets().values().get(spreadsheetId=ID_SHEET_GUIAS, range="'Guias_recibidas'!A2:J").execute()
        v = r.get('values', [])
        
        catalogo = {}
        meses_es = {"01":"Enero", "02":"Febrero", "03":"Marzo", "04":"Abril", "05":"Mayo", "06":"Junio", "07":"Julio", "08":"Agosto", "09":"Septiembre", "10":"Octubre", "11":"Noviembre", "12":"Diciembre"}
        
        for fila in v:
            # Solo consideramos si tiene al menos Fecha(0), Empresa(3), Fundo(4)
            if len(fila) >= 6:
                fecha_str = str(fila[0]).strip()
                empresa = str(fila[3]).strip()
                fundo = str(fila[4]).strip()
                bitacora = str(fila[7]).strip() if len(fila) > 7 else ""
                
                # REGLA ESTRICTA: Filtrar lo que ya fue generado
                if "✅ Nuevo" in bitacora:
                    continue
                
                mes_formateado = "Sin Fecha"
                # Parseo manual de fecha DD/MM/YYYY o M/D/YYYY
                if "/" in fecha_str:
                    partes = fecha_str.split("/")
                    if len(partes) >= 2:
                        mes_num = partes[1].zfill(2)
                        anio = partes[2] if len(partes) > 2 else ""
                        nombre_mes = meses_es.get(mes_num, mes_num)
                        mes_formateado = f"{nombre_mes} {anio}".strip()
                elif fecha_str:
                    mes_formateado = fecha_str

                if empresa:
                    if empresa not in catalogo:
                        catalogo[empresa] = {}
                    if mes_formateado not in catalogo[empresa]:
                        catalogo[empresa][mes_formateado] = set()
                    if fundo:
                        catalogo[empresa][mes_formateado].add(fundo)
                        
        # Formatear sets a listas ordenadas de fundos
        for emp in catalogo:
            for mes in catalogo[emp]:
                catalogo[emp][mes] = sorted(list(catalogo[emp][mes]))
        return catalogo
    except Exception as e:
        print(f"Error catalogo guias: {e}")
        return {}

def buscar_guias_repositorio(servicio_sheets, empresa, fundo, mes, tipo=None):
    """Filtra y devuelve archivos a descargar, validando que no estén procesados y extrayendo el tipo de Columna J."""
    if not servicio_sheets: return []
    try:
        r = servicio_sheets.spreadsheets().values().get(spreadsheetId=ID_SHEET_REPOSITORIO, range="'Guias_recibidas'!A2:J").execute()
        v = r.get('values', [])
        
        resultados = []
        meses_es = {"01":"Enero", "02":"Febrero", "03":"Marzo", "04":"Abril", "05":"Mayo", "06":"Junio", "07":"Julio", "08":"Agosto", "09":"Septiembre", "10":"Octubre", "11":"Noviembre", "12":"Diciembre"}
        import unicodedata
        
        for i, fila in enumerate(v):
            fila_real = i + 2 # Fila en excel (A2 en adelante)
            if len(fila) >= 6:
                fecha_str = str(fila[0]).strip()
                f_empresa = str(fila[3]).strip()
                f_fundo = str(fila[4]).strip()
                f_archivo = str(fila[5]).strip()
                bitacora = str(fila[7]).strip() if len(fila) > 7 else ""
                
                mes_formateado = "Sin Fecha"
                if "/" in fecha_str:
                    partes = fecha_str.split("/")
                    if len(partes) >= 2:
                        mes_num = partes[1].zfill(2)
                        anio = partes[2] if len(partes) > 2 else ""
                        nombre_mes = meses_es.get(mes_num, mes_num)
                        mes_formateado = f"{nombre_mes} {anio}".strip()
                elif fecha_str:
                    mes_formateado = fecha_str
                
                tipo_raw = str(fila[9]).strip() if len(fila) > 9 else ""
                tipo_norm = ''.join(c for c in unicodedata.normalize('NFD', tipo_raw) if unicodedata.category(c) != 'Mn').lower().replace(" ", "")
                
                # Regla automática desde Columna J:
                # - Comercialización -> plantilla 'Comercialización'
                # - Disposición Final / Servicios -> plantilla 'Disposición Final 1'
                if "comercializacion" in tipo_norm:
                    tipo_detectado = "Comercialización"
                    plantilla_flujo = "Comercialización"
                elif "disposicion" in tipo_norm or "servicio" in tipo_norm or "final" in tipo_norm:
                    tipo_detectado = "Disposición Final"
                    plantilla_flujo = "Disposición Final 1"
                else:
                    tipo_detectado = tipo_raw if tipo_raw else "Comercialización"
                    plantilla_flujo = "Comercialización" if "comercial" in tipo_detectado.lower() else "Disposición Final 1"

                # Coincidir con los filtros y que no contenga marca de 'Nuevo' procesado
                coincide_tipo = True if tipo is None else (tipo_detectado == tipo or tipo_raw == tipo)
                if f_empresa == empresa and f_fundo == fundo and mes_formateado == mes and coincide_tipo and f_archivo:
                    if "✅ Nuevo" not in bitacora:
                        numero_guia = str(fila[1]).strip() if len(fila) > 1 else "S/N"
                        resultados.append({
                            "nombre": f_archivo, 
                            "fila": fila_real, 
                            "numero_guia": numero_guia,
                            "tipo_detectado": tipo_detectado,
                            "tipo_flujo": plantilla_flujo,
                            "tipo_raw": tipo_raw
                        })
        return resultados
    except Exception as e:
        print(f"Error buscar guias: {e}")
        return []

def extraer_id_drive(valor):
    """
    Extrae de forma robusta el ID de un archivo o carpeta de Google Drive / Google Docs / Sheets.
    Admite URLs de tipo docs.google.com/document/d/<id>, drive.google.com/file/d/<id>,
    drive.google.com/open?id=<id>, o strings de ID directos.
    """
    import re
    if not valor:
        return None
    val_str = str(valor).strip()
    # Si es directamente un ID de Drive (25 a 65 caracteres alfanuméricos, guiones o guiones bajos)
    if re.match(r'^[a-zA-Z0-9_-]{25,65}$', val_str):
        return val_str
    
    patrones = [
        r'drive\.google\.com/(?:drive/)?(?:u/\d+/)?folders/([a-zA-Z0-9_-]+)',
        r'folders/([a-zA-Z0-9_-]+)',
        r'drive\.google\.com/file/d/([a-zA-Z0-9_-]+)',
        r'docs\.google\.com/(?:document|spreadsheets|presentation)/d/([a-zA-Z0-9_-]+)',
        r'drive\.google\.com/open\?id=([a-zA-Z0-9_-]+)',
        r'drive\.google\.com/uc\?.*?id=([a-zA-Z0-9_-]+)',
        r'[?&]id=([a-zA-Z0-9_-]+)',
        r'/d/([a-zA-Z0-9_-]+)'
    ]
    for pat in patrones:
        m = re.search(pat, val_str)
        if m:
            return m.group(1)
            
    return val_str

def descargar_guias_drive(servicio_drive, nombres_archivos):
    """Busca y descarga los archivos del Drive, los devuelve en memoria."""
    import io
    import re
    from googleapiclient.http import MediaIoBaseDownload
    
    if not servicio_drive or not nombres_archivos: return []
    
    archivos_memoria = []
    for nombre in nombres_archivos:
        archivo_id = extraer_id_drive(nombre)
        es_posible_id = (archivo_id != nombre) or (re.match(r'^[a-zA-Z0-9_-]{25,65}$', str(nombre).strip()) is not None)
        try:
            if es_posible_id:
                meta = servicio_drive.files().get(fileId=archivo_id, fields='name, mimeType', supportsAllDrives=True).execute()
                archivo_name = meta.get('name', 'archivo_descargado')
                mime_type = meta.get('mimeType', '')
                
                # Si es un documento nativo de Google Docs, Sheets, etc, exportarlo
                if mime_type == 'application/vnd.google-apps.document':
                    req = servicio_drive.files().export_media(fileId=archivo_id, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                elif mime_type == 'application/vnd.google-apps.spreadsheet':
                    req = servicio_drive.files().export_media(fileId=archivo_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                else:
                    req = servicio_drive.files().get_media(fileId=archivo_id)
                    
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, req)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                fh.seek(0)
                fh.name = archivo_name
                archivos_memoria.append(fh)
            else:
                # Usamos contains para atrapar variaciones o diferentes extensiones (JPG, PNG, PDF, etc)
                query = f"name contains '{nombre}' and trashed=false"
                res = servicio_drive.files().list(
                    q=query, 
                    spaces='drive', 
                    corpora='allDrives',
                    includeItemsFromAllDrives=True, 
                    supportsAllDrives=True, 
                    fields='files(id, name)'
                ).execute()
                items = res.get('files', [])
                if items:
                    archivo_id = items[0]['id']
                    req = servicio_drive.files().get_media(fileId=archivo_id)
                    fh = io.BytesIO()
                    downloader = MediaIoBaseDownload(fh, req)
                    done = False
                    while not done:
                        status, done = downloader.next_chunk()
                    fh.seek(0)
                    
                    # Mock streamlit file properties
                    fh.name = items[0]['name']
                    archivos_memoria.append(fh)
        except Exception as e:
            print(f"Error descargando {nombre}: {e}")
            
    return archivos_memoria

def actualizar_bitacora_guias(servicio_sheets, filas):
    """Actualiza la columna H de las filas iteradas en un Batch Update."""
    from datetime import datetime, timedelta
    if not servicio_sheets or not filas: return False
    try:
        hora_lima = datetime.utcnow() - timedelta(hours=5)
        marca = f"✅ Nuevo: {hora_lima.strftime('%d/%m/%Y %H:%M:%S')}"
        
        data = []
        for fila in filas:
            data.append({
                "range": f"'Guias_recibidas'!H{fila}",
                "values": [[marca]]
            })
            
        body = {
            "valueInputOption": "USER_ENTERED",
            "data": data
        }
        servicio_sheets.spreadsheets().values().batchUpdate(spreadsheetId=ID_SHEET_REPOSITORIO, body=body).execute()
        return True
    except Exception as e:
        print(f"Error actualizando bitácora: {e}")
        return False

def buscar_actualizar_guia(servicio_sheets, num_guia):
    from datetime import datetime, timedelta
    import re
    if not servicio_sheets or not str(num_guia).strip() or str(num_guia).strip() == "S/N": return False
    try:
        r = servicio_sheets.spreadsheets().values().get(spreadsheetId=ID_SHEET_REPOSITORIO, range="'Guias_recibidas'!A2:H").execute()
        v = r.get('values', [])
        
        hora_lima = datetime.utcnow() - timedelta(hours=5)
        marca = f"✅ Nuevo: {hora_lima.strftime('%d/%m/%Y %H:%M:%S')}"
        
        def norm_g(s):
            t = re.sub(r'(?i)[nº°\s_-]', '', str(s)).upper()
            return "".join([p.lstrip('0') if p.isdigit() else p for p in re.findall(r'[A-Z]+|[0-9]+', t)])

        target_norm = norm_g(num_guia)

        for i, fila in enumerate(v):
            if len(fila) > 1:
                val_col = str(fila[1]).strip().upper()
                if val_col == str(num_guia).strip().upper() or (target_norm and norm_g(val_col) == target_norm):
                    fila_excel = i + 2
                    body = {"values": [[marca]]}
                    servicio_sheets.spreadsheets().values().update(
                        spreadsheetId=ID_SHEET_REPOSITORIO,
                        range=f"'Guias_recibidas'!H{fila_excel}",
                        valueInputOption="USER_ENTERED",
                        body=body
                    ).execute()
        return True
    except Exception as e:
        print(f"Error actualizando marca de control {num_guia}: {e}")
        return False

@st.cache_data(ttl=600, show_spinner=False)
def obtener_usuarios_roles():
    """Descarga los roles de usuario desde la hoja 'Usuario_Roles'."""
    import time
    _, servicio_sheets = obtener_servicios()
    
    if not servicio_sheets:
        return {}
        
    SHEET_ID = "1OglcHUuRkD6LErUMevr-tOmAsLZpQEk2OKn8NyJN8mo" 
    
    intentos = 3
    for intento in range(intentos):
        try:
            # Apunta a la Base de Datos Exclusiva de Seguridad (RBAC)
            r = servicio_sheets.spreadsheets().values().get(
                spreadsheetId=SHEET_ID, 
                range="'Usuario_Roles'!A2:D"
            ).execute()
            
            v = r.get('values', [])
            usuarios = {}
            for fila in v:
                if len(fila) >= 4:
                    email = str(fila[0]).strip().lower()
                    nombre = str(fila[1]).strip()
                    rol = str(fila[2]).strip()
                    estado = str(fila[3]).strip()
                    
                    if email:
                        usuarios[email] = {
                            "Nombre": nombre,
                            "Rol": rol,
                            "Estado": estado
                        }
            return usuarios
        except Exception as e:
            if intento < intentos - 1:
                time.sleep(2)
                continue
            import streamlit as st
            st.error(f"Error cargando dict de roles: {e}")
            return None

def registrar_auditoria_sistema(correo, modo, guias_leidas, guias_certificadas):
    """Registra en la pestaña 'Control' del Sheet RBAC."""
    SHEET_ID = "1OglcHUuRkD6LErUMevr-tOmAsLZpQEk2OKn8NyJN8mo" 
    from datetime import datetime
    _, sheets = obtener_servicios()
    if not sheets: return False
    
    fecha = datetime.now().strftime("%d/%m/%Y")
    hora = datetime.now().strftime("%H:%M:%S")
    
    datos = [fecha, hora, str(correo), str(modo), str(guias_leidas), str(guias_certificadas)]
    
    try:
        sheets.spreadsheets().values().append(
            spreadsheetId=SHEET_ID, range="'Control'!A:F",
            valueInputOption="USER_ENTERED", 
            insertDataOption="INSERT_ROWS",
            body={"values": [datos]}
        ).execute()
        return True
    except Exception as e:
        print(f"Error registrando auditoría: {e}")
        return False

# ====================================================================
# --- BLOQUE 5: Unificador de Documentos (Certificado + Guías) ---
# ====================================================================

def es_guia_valida_para_unir(observacion):
    """
    Determina si una fila en Registro_Guias es válida para unir.
    SI DICE ERRADA -> NO (devuelve False).
    Excepción: si dice 'GUIA CORRECTA' (es la versión rectificada) -> devuelve True.
    """
    obs = str(observacion).strip().upper()
    if not obs or obs in ['NONE', 'NAN', '-', '']:
        return True
    if "GUIA CORRECTA" in obs:
        return True
    for patron in ["ERRADA", "GUIA ERRADA", "ERROR", "SE HIZO MAL", "ME EQUIVOQUE", "ANULADA"]:
        if patron in obs:
            return False
    return True

def obtener_carpeta_padre_archivo_drive(servicio_drive, file_id_o_nombre):
    """
    Obtiene el ID de la carpeta contenedora (parent) de un archivo en Google Drive.
    """
    import re
    if not servicio_drive or not file_id_o_nombre:
        return None
    try:
        f_id = extraer_id_drive(file_id_o_nombre)
        es_posible_id = (f_id != file_id_o_nombre) or (re.match(r'^[a-zA-Z0-9_-]{25,65}$', str(file_id_o_nombre).strip()) is not None)
        if not es_posible_id:
            # Buscar por nombre
            res = servicio_drive.files().list(
                q=f"name = '{file_id_o_nombre}' and trashed = false",
                fields="files(id, parents)",
                pageSize=1,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            files = res.get('files', [])
            if files and files[0].get('parents'):
                return files[0]['parents'][0]
            return None

        meta = servicio_drive.files().get(
            fileId=f_id,
            fields='id, parents',
            supportsAllDrives=True
        ).execute()
        parents = meta.get('parents', [])
        if parents:
            return parents[0]
    except Exception as e:
        print(f"Aviso al obtener carpeta contenedora en Drive ({file_id_o_nombre}): {e}")
    return None

def obtener_link_archivo_drive(servicio_drive, nombre_archivo, carpeta_id=None):
    """Obtiene el enlace webViewLink de un archivo en Drive por ID o nombre, con filtro opcional por carpeta."""
    import re
    if not servicio_drive or not nombre_archivo:
        return None
    try:
        archivo_id = extraer_id_drive(nombre_archivo)
        es_posible_id = (archivo_id != nombre_archivo) or (re.match(r'^[a-zA-Z0-9_-]{25,65}$', str(nombre_archivo).strip()) is not None)
        if es_posible_id:
            meta = servicio_drive.files().get(fileId=archivo_id, fields='webViewLink', supportsAllDrives=True).execute()
            return meta.get('webViewLink')
        else:
            if carpeta_id:
                query = f"'{carpeta_id}' in parents and name contains '{nombre_archivo}' and trashed=false"
            else:
                query = f"name contains '{nombre_archivo}' and trashed=false"
            res = servicio_drive.files().list(
                q=query, 
                spaces='drive', 
                corpora='allDrives',
                includeItemsFromAllDrives=True, 
                supportsAllDrives=True, 
                fields='files(id, webViewLink)'
            ).execute()
            items = res.get('files', [])
            if items:
                return items[0].get('webViewLink')
            elif carpeta_id:
                # Reintento sin filtro de carpeta por si fue movido
                res2 = servicio_drive.files().list(
                    q=f"name contains '{nombre_archivo}' and trashed=false",
                    spaces='drive',
                    corpora='allDrives',
                    includeItemsFromAllDrives=True,
                    supportsAllDrives=True,
                    fields='files(id, webViewLink)'
                ).execute()
                items2 = res2.get('files', [])
                if items2:
                    return items2[0].get('webViewLink')
    except Exception as e:
        print(f"Error obteniendo link Drive para {nombre_archivo}: {e}")
    return None

def descomponer_guia_para_match(texto):
    """
    Descompone un texto o código de guía extrayendo variantes canónicas:
    serie-correlativo, seriecorrelativo, correlativo limpio, etc.
    Soporta formatos tipo TR16-Nº 0006306, TR16-6306, EG03-00000494, 1002-2859, etc.
    """
    import re
    if not texto:
        return set()
    s = str(texto).strip().upper()
    if s in ['NONE', 'NAN', 'S/N', 'SIN GUIA', '-']:
        return set()
    
    # Limpiar prefijos de tipo de documento
    s = re.sub(r'(?i)\b(GUIA|GU[IÍ]A|GUA|REMISI[OÓ]N|TRANSPORTE|TRANSPORTISTA|REMITENTE)\b', ' ', s)
    # Limpiar prefijos de número tipo N°, Nº, Nro, No., etc.
    s = re.sub(r'(?i)\bN[º°oO\.\#\:]*\s*', ' ', s)
    s = re.sub(r'[º°\#]', ' ', s)
    
    variantes = set()
    
    # 1. Separar por guión, barra, espacio o subguión
    partes = [p.strip() for p in re.split(r'[-\s/_]+', s) if p.strip()]
    if len(partes) >= 2:
        serie_raw = partes[0]
        corr_raw = partes[-1]
        
        # Normalizar serie (ej: TR016 -> TR16 y TR016; EG03 -> EG3 y EG03)
        m_serie = re.match(r'^([A-Z]*)(\d*)$', serie_raw)
        if m_serie:
            letras, num_s = m_serie.groups()
            num_s_clean = num_s.lstrip('0')
            serie_norm1 = f"{letras}{num_s_clean}"
            serie_norm2 = f"{letras}{num_s}"
        else:
            serie_norm1 = serie_raw
            serie_norm2 = serie_raw
            
        corr_clean = corr_raw.lstrip('0') or "0"
            
        variantes.add(f"{serie_norm1}-{corr_clean}")
        variantes.add(f"{serie_norm1}{corr_clean}")
        variantes.add(f"{serie_norm2}-{corr_clean}")
        variantes.add(f"{serie_norm2}{corr_clean}")
        variantes.add(f"{serie_norm1}-{corr_raw}")
        variantes.add(f"{serie_norm1}{corr_raw}")
        if len(corr_clean) >= 4:
            variantes.add(corr_clean)
    else:
        m = re.match(r'^([A-Z]+)(\d+)$', s)
        if m:
            letras, numeros = m.groups()
            variantes.add(f"{letras}{numeros.lstrip('0')}")
            variantes.add(s)
        else:
            t = re.sub(r'[^A-Z0-9]', '', s)
            if t:
                variantes.add(t)
                
    return variantes

def coinciden_guias_expediente(guia_buscada, guia_candidata):
    """Compara si dos guías coinciden de forma estricta y segura, evitando falsos positivos."""
    if not guia_buscada or not guia_candidata:
        return False
    g_c_str = str(guia_candidata).strip().upper()
    if not g_c_str or g_c_str in ['NONE', 'NAN', 'S/N', 'SIN GUIA', '-']:
        return False
        
    v_busc = descomponer_guia_para_match(guia_buscada)
    v_cand = descomponer_guia_para_match(guia_candidata)
    if not v_busc or not v_cand:
        return False
        
    inter = v_busc.intersection(v_cand)
    # Evitar matches espurios de solo números muy cortos (menores a 4 dígitos)
    inter_validas = {x for x in inter if not x.isdigit() or len(x) >= 4}
    return len(inter_validas) > 0

def buscar_guias_asociadas_para_unir(sheets_service, drive_service, guias_lista):
    """
    Busca en Registro_Guias las guías de Transporte y Remisión vinculadas al certificado.
    Filtra automáticamente filas con 'ERRADA'.
    Retorna lista con detalles y enlaces de Drive.
    """
    import re
    if not sheets_service or not guias_lista:
        return []
    
    # Aplanar y separar guías que vengan separadas por comas o saltos de línea
    guias_a_buscar = []
    for item in guias_lista:
        if not item: continue
        for sub in re.split(r'[,;\n]+', str(item)):
            sub_c = sub.strip()
            if sub_c and sub_c.upper() not in ['NONE', 'NAN', 'S/N', 'SIN GUIA', '-']:
                guias_a_buscar.append(sub_c)
                
    if not guias_a_buscar:
        return []
    
    try:
        r = sheets_service.spreadsheets().values().get(
            spreadsheetId=ID_SHEET_REPOSITORIO, 
            range="'Registro_Guias'!A1:O"
        ).execute()
        filas = r.get('values', [])
        if len(filas) < 2: return []
        
        resultados = []
        
        for idx, row in enumerate(filas[1:]):
            row_idx = idx + 2
            col1_guia = str(row[1]).strip() if len(row) > 1 else ""
            col2_ligada = str(row[2]).strip() if len(row) > 2 else ""
            tipo_guia = str(row[3]).strip().upper() if len(row) > 3 else ""
            col8_hecha = str(row[8]).strip() if len(row) > 8 else ""
            col9_recibida = str(row[9]).strip() if len(row) > 9 else ""
            observacion = str(row[11]).strip() if len(row) > 11 else ""
            
            # Mapeo según Tipo Guía:
            # Si es TRANSPORTISTA:
            #   N° Guía (col1) es Transporte (EG03-...), Archivo col8 (Guia hecha) es Transporte
            #   Guía ligada (col2) es Remisión (TR...), Archivo col9 (Guia recibida) es Remisión
            # Si es REMITENTE:
            #   N° Guía (col1) es Remisión (EG07-...), Archivo col8 (Guia hecha) es Remisión
            #   Guía ligada (col2) es Transporte (TR...), Archivo col9 (Guia recibida) es Transporte
            if "TRANSP" in tipo_guia or col1_guia.startswith("EG03"):
                num_transporte = col1_guia
                archivo_transporte = col8_hecha
                num_remision = col2_ligada
                archivo_remision = col9_recibida
            else:
                num_remision = col1_guia
                archivo_remision = col8_hecha
                num_transporte = col2_ligada
                archivo_transporte = col9_recibida
                
            match = False
            for g_busc in guias_a_buscar:
                if (coinciden_guias_expediente(g_busc, num_remision) or 
                    coinciden_guias_expediente(g_busc, num_transporte) or 
                    coinciden_guias_expediente(g_busc, col1_guia) or 
                    coinciden_guias_expediente(g_busc, col2_ligada)):
                    match = True
                    break
                    
            if match:
                es_valida = es_guia_valida_para_unir(observacion)
                link_trans = obtener_link_archivo_drive(drive_service, archivo_transporte) if archivo_transporte else None
                link_rem = obtener_link_archivo_drive(drive_service, archivo_remision) if archivo_remision else None
                
                resultados.append({
                    "fila": row_idx,
                    "num_transporte": num_transporte,
                    "num_remision": num_remision,
                    "archivo_transporte": archivo_transporte,
                    "archivo_remision": archivo_remision,
                    "observacion": observacion,
                    "valida": es_valida,
                    "link_transporte": link_trans,
                    "link_remision": link_rem
                })
                
        return resultados
    except Exception as e:
        print(f"Error buscando guías asociadas: {e}")
        return []

def descargar_archivo_drive_por_id_o_nombre(servicio_drive, file_id_o_nombre):
    """Descarga un archivo específico de Drive (exporta a Word si es Google Doc, o usa headRevisionId para binarios y evitar cualquier caché)."""
    import io
    from googleapiclient.http import MediaIoBaseDownload
    if not servicio_drive or not file_id_o_nombre:
        return None
    try:
        archivo_id = extraer_id_drive(file_id_o_nombre)
        
        try:
            meta = servicio_drive.files().get(
                fileId=archivo_id, 
                fields='id, name, mimeType, headRevisionId, modifiedTime', 
                supportsAllDrives=True
            ).execute()
            mime_type = meta.get('mimeType', '')
            archivo_name = meta.get('name', 'archivo_descargado')
            head_rev = meta.get('headRevisionId')
            
            if mime_type == 'application/vnd.google-apps.document':
                req = servicio_drive.files().export_media(fileId=archivo_id, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            elif head_rev:
                # Descarga directa por revisionId para garantizar la última versión sin caché intermedio
                try:
                    req = servicio_drive.revisions().get_media(fileId=archivo_id, revisionId=head_rev)
                except Exception:
                    req = servicio_drive.files().get_media(fileId=archivo_id)
            else:
                req = servicio_drive.files().get_media(fileId=archivo_id)
                
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)
            fh.name = archivo_name
            return fh
        except Exception as e_get:
            archivos = descargar_guias_drive(servicio_drive, [str(file_id_o_nombre)])
            if archivos:
                return archivos[0]
            print(f"Error descargando por ID {archivo_id}: {e_get}")
    except Exception as e:
        print(f"Error descargando {file_id_o_nombre}: {e}")
    return None

def obtener_info_revision_documento_drive(servicio_drive, file_id_o_nombre):
    """
    Inspecciona en tiempo real el archivo en Google Drive y devuelve:
    - nombre: nombre del archivo
    - modified_pe: fecha y hora formateada en hora Perú (UTC-5)
    - tiempo_rel: texto amigable ('hace 15 segundos', 'hace 2 minutos', etc.)
    - diff_seg: diferencia en segundos desde la última modificación
    - filas_resumen: lista de filas de tabla extraídas del .docx para previsualización inmediata
    - es_reciente: True si fue modificado en los últimos 15 minutos
    """
    from datetime import datetime, timezone, timedelta
    from docx import Document
    if not servicio_drive or not file_id_o_nombre:
        return None
    try:
        archivo_id = extraer_id_drive(file_id_o_nombre)
        meta = servicio_drive.files().get(
            fileId=archivo_id, 
            fields='id, name, mimeType, modifiedTime, headRevisionId', 
            supportsAllDrives=True
        ).execute()
        
        mod_str = meta.get('modifiedTime')
        if not mod_str:
            return None
            
        dt_utc = datetime.fromisoformat(mod_str.replace('Z', '+00:00'))
        dt_pe = dt_utc.astimezone(timezone(timedelta(hours=-5)))
        ahora_pe = datetime.now(timezone(timedelta(hours=-5)))
        diff_seg = int((ahora_pe - dt_pe).total_seconds())

        if diff_seg < 60:
            tiempo_rel = f"hace {max(1, diff_seg)} segundo(s)"
        elif diff_seg < 3600:
            tiempo_rel = f"hace {diff_seg // 60} minuto(s)"
        else:
            tiempo_rel = f"hace {diff_seg // 3600} hora(s)"

        # Descargar y extraer vista rápida de texto/tablas
        fh = descargar_archivo_drive_por_id_o_nombre(servicio_drive, archivo_id)
        filas_resumen = []
        if fh:
            try:
                doc = Document(fh)
                for t in doc.tables:
                    for row in t.rows[:6]:
                        vals = [c.text.strip() for c in row.cells if c.text.strip()]
                        if vals:
                            filas_resumen.append(" | ".join(vals))
            except Exception:
                pass

        return {
            "id": archivo_id,
            "name": meta.get('name', ''),
            "modified_pe": dt_pe.strftime("%d/%m/%Y %H:%M:%S"),
            "tiempo_rel": tiempo_rel,
            "diff_seg": diff_seg,
            "es_reciente": diff_seg < 900,
            "filas_resumen": filas_resumen
        }
    except Exception as e:
        print(f"Error obteniendo info de revisión en Drive: {e}")
        return None

def convertir_imagen_a_pdf(imagen_bytes, nombre_archivo=""):
    """Convierte bytes de imagen (JPG, PNG, JPEG, WEBP) a páginas PDF en memoria."""
    import io
    from PIL import Image
    nombre_lower = str(nombre_archivo).lower()
    
    es_imagen = (
        nombre_lower.endswith(('.jpg', '.jpeg', '.png', '.webp', '.bmp')) or
        imagen_bytes[:4] in [b'\xff\xd8\xff\xe0', b'\xff\xd8\xff\xe1', b'\x89PNG']
    )
    if es_imagen:
        try:
            img = Image.open(io.BytesIO(imagen_bytes))
            if img.mode != 'RGB':
                img = img.convert('RGB')
            out_pdf = io.BytesIO()
            img.save(out_pdf, format='PDF')
            out_pdf.seek(0)
            return out_pdf.getvalue()
        except Exception as e:
            print(f"Error convirtiendo imagen a PDF ({nombre_archivo}): {e}")
    return imagen_bytes

def convertir_docx_a_pdf(docx_bytes):
    """Convierte un documento Word .docx a PDF. Prioridad: MS Word local (win32com). Fallback: Drive API."""
    import io
    import os
    import tempfile
    
    # 1. Intentar con MS Word local
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as f_in:
            f_in.write(docx_bytes)
            docx_path = f_in.name
        pdf_path = docx_path.replace('.docx', '.pdf')
        
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17) # 17 = wdFormatPDF
        doc.Close()
        word.Quit()
        
        with open(pdf_path, 'rb') as f_out:
            pdf_bytes = f_out.read()
            
        if os.path.exists(docx_path): os.remove(docx_path)
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return pdf_bytes
    except Exception as e:
        print(f"Aviso: win32com no disponible ({e}). Intentando fallback con Google Drive API...")
        
    # 2. Fallback vía Google Drive API
    drive_service, _ = obtener_servicios()
    if drive_service:
        try:
            from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
            media = MediaIoBaseUpload(io.BytesIO(docx_bytes), mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', resumable=False)
            temp_file = drive_service.files().create(
                body={'name': 'temp_conv_pdf', 'mimeType': 'application/vnd.google-apps.document'},
                media_body=media,
                fields='id'
            ).execute()
            temp_id = temp_file.get('id')
            req = drive_service.files().export_media(fileId=temp_id, mimeType='application/pdf')
            fh = io.BytesIO()
            dl = MediaIoBaseDownload(fh, req)
            done = False
            while not done: _, done = dl.next_chunk()
            drive_service.files().delete(fileId=temp_id).execute()
            return fh.getvalue()
        except Exception as e2:
            print(f"Fallback Drive API PDF conversion falló: {e2}")
            
    return None

def unir_tres_documentos_pdf(cert_docx_bytes, remision_file_obj, transporte_file_obj):
    """
    Une en orden estricto los 3 documentos:
    1. Certificado (Página 1)
    2. Guía de Remisión (Página 2)
    3. Guía de Transporte (Página 3)
    Convierte imágenes (JPG/PNG) a PDF si es necesario.
    """
    import io
    from pypdf import PdfWriter, PdfReader
    
    # 1. Convertir Certificado a PDF
    cert_pdf = convertir_docx_a_pdf(cert_docx_bytes)
    if not cert_pdf:
        raise Exception("No se pudo convertir el Certificado Word a PDF.")
        
    writer = PdfWriter()
    
    # Agregar páginas del Certificado
    reader_cert = PdfReader(io.BytesIO(cert_pdf))
    for page in reader_cert.pages:
        writer.add_page(page)
        
    # 2. Agregar páginas de la Guía de Remisión
    if remision_file_obj:
        rem_bytes = remision_file_obj.getvalue() if hasattr(remision_file_obj, 'getvalue') else remision_file_obj.read()
        rem_nombre = getattr(remision_file_obj, 'name', '')
        rem_pdf = convertir_imagen_a_pdf(rem_bytes, rem_nombre)
        try:
            reader_rem = PdfReader(io.BytesIO(rem_pdf))
            for page in reader_rem.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"Error procesando páginas de Guía de Remisión: {e}")
            
    # 3. Agregar páginas de la Guía de Transporte
    if transporte_file_obj:
        trans_bytes = transporte_file_obj.getvalue() if hasattr(transporte_file_obj, 'getvalue') else transporte_file_obj.read()
        trans_nombre = getattr(transporte_file_obj, 'name', '')
        trans_pdf = convertir_imagen_a_pdf(trans_bytes, trans_nombre)
        try:
            reader_trans = PdfReader(io.BytesIO(trans_pdf))
            for page in reader_trans.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"Error procesando páginas de Guía de Transporte: {e}")
            
    output = io.BytesIO()
    writer.write(output)
    output.seek(0)
    return output.getvalue()

def subir_pdf_a_drive(contenido_bytes, nombre_archivo, tipo_flujo, carpeta_id=None):
    """Sube el archivo PDF unificado a Drive enrutándolo a la carpeta correcta."""
    import unicodedata
    import io
    from googleapiclient.http import MediaIoBaseUpload
    
    if not carpeta_id:
        def normalizar(texto):
            texto = str(texto).lower()
            return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
            
        tipo_seguro = normalizar(tipo_flujo) 
        if "comercializacion" in tipo_seguro:
            carpeta_id = "1NZc-nfGHw5bnkCAv0TdQYW_bPM_UkKC-" # Comercialización
        else:
            carpeta_id = "12PMJ1d-CSWo64m7aNQRQj2yGHFdp9B9S" # Servicios (Disp. Final)

    drive, _ = obtener_servicios()
    if not drive: return None
    
    try:
        nom = nombre_archivo if nombre_archivo.lower().endswith('.pdf') else f"{nombre_archivo}.pdf"
        file_metadata = {
            'name': nom, 
            'mimeType': 'application/pdf',
            'parents': [carpeta_id]
        }
        media = MediaIoBaseUpload(
            io.BytesIO(contenido_bytes), 
            mimetype='application/pdf', 
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
        import streamlit as st
        st.error(f"Google Drive rechazó la subida del PDF. Motivo: {str(e)}")
        print(f"Error subiendo PDF a Drive: {e}") 
        return None

def actualizar_link_pdf_historial(servicio_sheets, fila_historial, link_pdf, correlativo=None, link_doc=None, hoja_destino="Historial"):
    """
    Actualiza la columna I ('Link pdf') en la fila correspondiente de la hoja de historial especificada.
    Si se pasa fila_historial válida, la usa directamente.
    Si no, busca la fila por correlativo o link_doc para máxima precisión y evitar filas duplicadas.
    """
    if not servicio_sheets or not link_pdf:
        return False
    try:
        fila_destino = fila_historial if (fila_historial and isinstance(fila_historial, int) and fila_historial > 1) else None
        
        # Si no se tiene el número exacto de fila, buscar en la hoja
        if not fila_destino:
            r = servicio_sheets.spreadsheets().values().get(
                spreadsheetId=ID_SHEET_CONTROL,
                range=f"'{hoja_destino}'!A:I"
            ).execute()
            filas = r.get('values', [])
            for idx, row in enumerate(filas):
                num_fila = idx + 1
                corr_row = str(row[3]).strip() if len(row) > 3 else ""
                doc_row = str(row[7]).strip() if len(row) > 7 else ""
                if link_doc and str(link_doc).strip() and str(link_doc).strip() in doc_row:
                    fila_destino = num_fila
                    break
                if correlativo and str(correlativo).strip() and str(correlativo).strip() == corr_row:
                    fila_destino = num_fila
                    
        if fila_destino:
            body = {"values": [[link_pdf]]}
            servicio_sheets.spreadsheets().values().update(
                spreadsheetId=ID_SHEET_CONTROL,
                range=f"'{hoja_destino}'!I{fila_destino}",
                valueInputOption="USER_ENTERED",
                body=body
            ).execute()
            return True
        else:
            print(f"No se pudo determinar la fila en {hoja_destino} para correlativo={correlativo}, link={link_doc}")
            return False
    except Exception as e:
        print(f"Error actualizando link PDF en {hoja_destino}: {e}")
        return False

def obtener_catalogo_servicios_por_categoria(df_servicios):
    """
    Parsea el DataFrame de la pestaña 'SERVICIOS' agrupando las opciones por categoría
    (COMERCIALIZACION vs SERVICIOS). Ignora filas de cabecera repetidas y limpia textos.
    """
    import pandas as pd
    secciones = {
        "COMERCIALIZACION": {"titulos": [], "servicios": [], "residuos": []},
        "SERVICIOS": {"titulos": [], "servicios": [], "residuos": []}
    }
    if df_servicios is None or df_servicios.empty:
        return secciones

    KEYWORDS_HEADER = ["COMERCIALIZACION", "COMERCIALIZACIÓN", "SERVICIOS", "DISPOSICION FINAL", "DISPOSICIÓN FINAL"]
    seccion_actual = "COMERCIALIZACION"

    for _, row in df_servicios.iterrows():
        c0 = str(row.iloc[0]).strip() if len(row) > 0 and pd.notna(row.iloc[0]) else ""
        c1 = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
        c2 = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""

        c0_norm = c0.upper().replace("Ó", "O").replace("Í", "I").replace("Á", "A")
        if c0_norm in [k.replace("Ó", "O").replace("Í", "I").replace("Á", "A") for k in KEYWORDS_HEADER]:
            seccion_actual = "COMERCIALIZACION" if "COMERC" in c0_norm else "SERVICIOS"
            continue

        if seccion_actual in secciones:
            if c0 and c0 not in secciones[seccion_actual]["titulos"]:
                secciones[seccion_actual]["titulos"].append(c0)
            if c1 and c1 not in secciones[seccion_actual]["servicios"]:
                secciones[seccion_actual]["servicios"].append(c1)
            if c2 and c2 not in secciones[seccion_actual]["residuos"]:
                secciones[seccion_actual]["residuos"].append(c2)

    return secciones

def buscar_datos_certificado_en_historial(servicio_sheets, correlativo=None, servicio_drive=None, hoja_destino="Historial", carpeta_empresa_id=None, limite=100):
    """
    Busca certificados en la pestaña de historial especificada por número correlativo o texto.
    Si correlativo es None o vacío, retorna los certificados más recientes hasta el límite indicado.
    Si se proporciona servicio_drive, resuelve automáticamente los enlaces web (webViewLink)
    de Drive utilizando la carpeta de la empresa para evitar confusiones de nombres.
    """
    if not servicio_sheets:
        return []
    try:
        r = servicio_sheets.spreadsheets().values().get(
            spreadsheetId=ID_SHEET_CONTROL,
            range=f"'{hoja_destino}'!A:J"
        ).execute()
        filas = r.get('values', [])
        if len(filas) < 2:
            return []

        corr_query = str(correlativo).strip().upper() if correlativo is not None else ""
        corr_num = None
        if corr_query.isdigit():
            corr_num = int(corr_query)

        resultados = []
        for idx, row in enumerate(filas[1:]):
            num_fila = idx + 2
            corr_row = str(row[3]).strip().upper() if len(row) > 3 else ""
            empresa_row = str(row[1]).strip().upper() if len(row) > 1 else ""
            fundo_row = str(row[2]).strip().upper() if len(row) > 2 else ""
            guia_row = str(row[5]).strip().upper() if len(row) > 5 else ""
            
            match = False
            if not corr_query:
                # Sin filtro de búsqueda: devolver todas las filas para que la UI muestre el listado reciente
                match = True
            elif corr_row and corr_row == corr_query:
                match = True
            elif corr_num is not None and corr_row.isdigit() and int(corr_row) == corr_num:
                match = True
            elif corr_query:
                # Búsqueda flexible por correlativo o nombre de cliente/guía
                if corr_query in corr_row or corr_query in empresa_row or corr_query in fundo_row or corr_query in guia_row:
                    match = True

            if match:
                raw_guia = str(row[6]).strip() if len(row) > 6 else ""
                raw_doc = str(row[7]).strip() if len(row) > 7 else ""
                raw_pdf = str(row[8]).strip() if len(row) > 8 else ""

                url_doc = raw_doc
                if raw_doc and not raw_doc.startswith(('http://', 'https://')):
                    if servicio_drive:
                        link_d = obtener_link_archivo_drive(servicio_drive, raw_doc, carpeta_id=carpeta_empresa_id)
                        if link_d:
                            url_doc = link_d

                url_pdf = raw_pdf
                if raw_pdf and not raw_pdf.startswith(('http://', 'https://')):
                    if servicio_drive:
                        link_p = obtener_link_archivo_drive(servicio_drive, raw_pdf, carpeta_id=carpeta_empresa_id)
                        if link_p:
                            url_pdf = link_p

                resultados.append({
                    "fila": num_fila,
                    "fecha": str(row[0]).strip() if len(row) > 0 else "",
                    "empresa": str(row[1]).strip() if len(row) > 1 else "",
                    "fundo": str(row[2]).strip() if len(row) > 2 else "",
                    "correlativo": corr_row,
                    "tipo_cert": str(row[4]).strip() if len(row) > 4 else "",
                    "guias": str(row[5]).strip() if len(row) > 5 else "",
                    "link_guia": raw_guia,
                    "link_doc": url_doc,
                    "raw_doc": raw_doc,
                    "link_pdf": url_pdf,
                    "raw_pdf": raw_pdf,
                    "observacion": str(row[9]).strip() if len(row) > 9 else "",
                    "hoja_historial": hoja_destino
                })

        if not corr_query and len(resultados) > limite:
            return list(reversed(resultados))[:limite]
        return resultados
    except Exception as e:
        print(f"Error buscando datos de certificado en Historial ({hoja_destino}): {e}")
        return []

def guardar_pdf_en_carpeta_docs(servicio_drive, pdf_bytes, nombre_archivo, file_id_docs=None, file_id_pdf_existente=None, tipo_flujo="Comercialización", carpeta_fallback_id=None):
    """
    Guarda o actualiza el archivo PDF en Google Drive exactamente en la carpeta
    donde está alojado el archivo Google Docs/Word asociado.
    Si ya existía el PDF en Drive, intenta actualización in-place para preservar URL y permisos.
    Retorna tupla: (link_pdf, file_id_pdf)
    """
    import io
    from googleapiclient.http import MediaIoBaseUpload

    if not servicio_drive:
        return None, None

    # 1. Resolver carpeta contenedora del Docs
    carpeta_destino = None
    if file_id_docs:
        try:
            carpeta_destino = obtener_carpeta_padre_archivo_drive(servicio_drive, file_id_docs)
        except Exception as e_p:
            print(f"Aviso obteniendo carpeta del Docs: {e_p}")
    if not carpeta_destino:
        carpeta_destino = carpeta_fallback_id

    nom_pdf = nombre_archivo.strip() if nombre_archivo else "Expediente.pdf"
    if not nom_pdf.lower().endswith('.pdf'):
        nom_pdf += '.pdf'

    # 2. Intentar actualización in-place si ya existe el PDF en Drive
    if file_id_pdf_existente:
        try:
            media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype='application/pdf', resumable=True)
            res_up = servicio_drive.files().update(
                fileId=file_id_pdf_existente,
                body={'name': nom_pdf},
                media_body=media,
                fields='id, name, webViewLink',
                supportsAllDrives=True
            ).execute()
            link = res_up.get('webViewLink')
            if link:
                return link, file_id_pdf_existente
        except Exception as e_up:
            print(f"Aviso: Falló actualización in-place de PDF ({e_up}). Se creará en la carpeta del Docs.")

    # 3. Subir como nuevo archivo en la carpeta del Docs
    try:
        body = {
            'name': nom_pdf,
            'mimeType': 'application/pdf'
        }
        if carpeta_destino:
            body['parents'] = [carpeta_destino]

        media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype='application/pdf', resumable=True)
        res_new = servicio_drive.files().create(
            body=body,
            media_body=media,
            fields='id, name, webViewLink',
            supportsAllDrives=True
        ).execute()
        return res_new.get('webViewLink'), res_new.get('id')
    except Exception as e_cr:
        print(f"Error creando PDF en carpeta del Docs ({carpeta_destino}): {e_cr}")
        return None, None

def sobrescribir_o_subir_pdf_drive(servicio_drive, file_id_existente, contenido_bytes, nombre_archivo="Expediente.pdf", tipo_flujo="Comercialización", carpeta_id=None):
    """
    Intenta actualizar in-place el archivo existente en Google Drive para preservar su ID y URL pública.
    Asegura que el nombre del archivo se conserve exactamente como se especificó.
    Si no existe o falla, sube un nuevo archivo en la carpeta designada.
    """
    import io
    from googleapiclient.http import MediaIoBaseUpload

    if servicio_drive and file_id_existente:
        try:
            media = MediaIoBaseUpload(io.BytesIO(contenido_bytes), mimetype='application/pdf', resumable=True)
            body = {'name': nombre_archivo} if nombre_archivo else {}
            res = servicio_drive.files().update(
                fileId=file_id_existente,
                body=body,
                media_body=media,
                fields='id, name, webViewLink',
                supportsAllDrives=True
            ).execute()
            link = res.get('webViewLink')
            if link:
                return link
        except Exception as e:
            print(f"Aviso: Falló actualización in-place en Drive ({e}). Se procederá a subir como nuevo archivo.")

    # Fallback: Subir como nuevo archivo
    return subir_pdf_a_drive(contenido_bytes, nombre_archivo, tipo_flujo, carpeta_id=carpeta_id)

def registrar_edicion_en_historial(servicio_sheets, num_fila, link_pdf, usuario_editor="Usuario", obs_extra="", hoja_destino="Historial"):
    """
    Actualiza la fila correspondiente en la pestaña de historial especificada:
    - Columna I ('Link pdf'): Enlace actualizado del PDF unificado.
    - Columna J ('Obs'): Bitácora con fecha, hora y usuario editor.
    """
    from datetime import datetime, timedelta
    if not servicio_sheets or not num_fila or num_fila < 2:
        return False
    try:
        ahora_pe = (datetime.utcnow() - timedelta(hours=5)).strftime("%d/%m/%Y %H:%M")
        nota_audit = f"Actualizado el {ahora_pe} por {usuario_editor}"
        if obs_extra:
            nota_audit += f" | {obs_extra}"

        body = {
            "values": [[link_pdf, nota_audit]]
        }
        servicio_sheets.spreadsheets().values().update(
            spreadsheetId=ID_SHEET_CONTROL,
            range=f"'{hoja_destino}'!I{num_fila}:J{num_fila}",
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()
        return True
    except Exception as e:
        print(f"Error registrando edición en {hoja_destino}: {e}")
        return False

