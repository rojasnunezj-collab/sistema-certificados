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

    # === INSTRUCCIONES DETALLADAS (PROMPT RESTAURADO Y MEJORADO) ===
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
    
    try:
        res = model.generate_content([prompt, {"mime_type": "application/pdf", "data": base64.b64encode(pdf_bytes).decode('utf-8')}])
        match = re.search(r'\{.*\}', res.text.replace("```json", "").replace("```", ""), re.DOTALL)
        if match: return json.loads(match.group(0))
        return None
    except Exception as e:
        if "429" in str(e):
             st.warning("⏳ Velocidad: Pausa automática activada...")
             time.sleep(5) 
             return None
        st.error(f"❌ Error lectura ({candidato}): {e}")
        return None

# ==========================================
# 3. INTERFAZ
# ==========================================
if 'ocr_data' not in st.session_state: st.session_state['ocr_data'] = None
if 'df_items' not in st.session_state: st.session_state['df_items'] = pd.DataFrame()
if 'datos_log_pendientes' not in st.session_state: st.session_state['datos_log_pendientes'] = {}

with st.sidebar:
    with st.expander("❓ Ayuda Rápida"):
        st.markdown("""
        1. **Sube PDF**.
        2. Clic **Procesar**.
        3. **Revisa** datos.
        4. **Descarga** Word/Excel.
        5. **Registra** link final.
        """)
    st.divider()
    st.header("⚙️ Configuración")
    empresa_firma = st.selectbox("Empresa", list(PLANTILLAS.keys()))
    tipo_plantilla = st.selectbox("Plantilla", ["Comercialización/Disposición Final", "Peligroso y No Peligroso"])
    if st.button("🔄 Reiniciar"): st.cache_data.clear(); st.rerun()

st.title("Generador de Certificados (v4.5)")

if 'repo_data' not in st.session_state:
    st.session_state['repo_data'] = {
        "emisores": leer_sheet_seguro("EMPRESAS"),
        "clientes": leer_sheet_seguro("CLIENTES"),
        "servicios": leer_sheet_seguro("COMERCIALIZACION")
    }
repo = st.session_state['repo_data']

archivos = st.file_uploader("1. Subir Guías (PDF)", type=["pdf"], accept_multiple_files=True)

if archivos:
    if st.button("🔍 Procesar Guías con IA"):
        prog = st.progress(0)
        items, grl = [], None
        errores = 0
        total = len(archivos)
        
        for i, arc in enumerate(archivos):
            # === FRENO DE MANO ===
            if i > 0: time.sleep(4) 
            
            d = procesar_guia_ia(arc.read())
            if d:
                if not grl: grl = d
                s, f, p = formatear_guia(d.get('serie','S/N')), d.get('fecha',''), d.get('vehiculo','')
                for it in d.get('items', []):
                    it.update({'guia_origen': s, 'fecha_origen': f, 'placa_origen': p})
                    items.append(it)
            else: errores += 1
            prog.progress((i+1)/total)
        
        time.sleep(0.5); prog.empty()
        
        if grl and items:
            st.session_state['ocr_data'] = grl
            df = pd.DataFrame(items)
            for c in ['desc','cant','um','peso','fecha_origen','guia_origen','placa_origen']:
                if c not in df.columns: df[c] = ""
            df['peso'] = df['peso'].replace("", "0.00").replace("None", "0.00")
            df['desc'] = df['desc'].apply(limpiar_descripcion)
            df['fecha_origen'] = df['fecha_origen'].apply(normalizar_fecha)
            st.session_state['df_items'] = df
            st.success(f"✅ Procesado: {total-errores} correctos.")
        else: st.error("❌ No se pudieron extraer datos.")

# EDICIÓN
if st.session_state['ocr_data']:
    ocr = st.session_state['ocr_data']
    st.markdown("### 2. Validación")
    
    c1, c2, c3, c4 = st.columns(4)
    v_corr = c1.text_input("Correlativo", "001")
    fecha_base = normalizar_fecha(ocr.get('fecha'))
    cont_f = c2.container()
    
    opt_f = c2.radio("Regla Fecha:", ["Comercialización (+2)", "Disposición Final (Fin de Mes)"], label_visibility="collapsed")
    f_calc = fecha_base
    tipo_operacion_simple = "" 
    try:
        if "Comercialización" in opt_f: 
            f_calc = (datetime.strptime(fecha_base, "%d/%m/%Y")+timedelta(days=2)).strftime("%d/%m/%Y")
            tipo_operacion_simple = "Comercialización"
        else: 
            f_calc = obtener_fin_de_mes(fecha_base)
            tipo_operacion_simple = "Disposición Final"
    except: tipo_operacion_simple = "Indefinido"

    v_fec_emis = cont_f.text_input("F. Emisión", value=f_calc)

    if len(archivos) > 1:
        v_guia, v_placa = c3.text_input("Guía", "VARIAS"), c4.text_input("Placa", "VARIAS")
    else:
        v_guia, v_placa = c3.text_input("Guía", formatear_guia(ocr.get('serie'))), c4.text_input("Placa", ocr.get('vehiculo'))

    v_partida = st.text_input("Partida", formato_nompropio(ocr.get('punto_partida','')))
    v_llegada = st.text_input("Llegada", formato_nompropio(ocr.get('punto_llegada','')))
    v_dest = st.text_input("Destinatario", ocr.get('destinatario',''))

    v_items = st.data_editor(st.session_state['df_items'], num_rows="dynamic", use_container_width=True,
        column_config={"guia_origen": st.column_config.TextColumn("Guía", disabled=True)})
    
    c_a, c_b = st.columns(2)
    with c_a:
        v_emi = st.selectbox("Emisor", repo['emisores']['EMPRESA'].unique() if not repo['emisores'].empty else [])
        v_ruc_e, v_reg_e = "", ""
        if not repo['emisores'].empty and v_emi:
            try:
                row_e = repo['emisores'][repo['emisores']['EMPRESA']==v_emi].iloc[0]
                v_ruc_e, v_reg_e = str(row_e['RUC']), str(row_e['REGISTRO'])
            except: pass
        st.caption(f"RUC: {v_ruc_e} | REG: {v_reg_e}")
        v_tit = st.selectbox("Título Certificado", repo['servicios'].iloc[:,0].unique() if not repo['servicios'].empty else [])

    with c_b:
        v_cli = st.selectbox("Cliente", repo['clientes']['EMPRESA'].unique() if not repo['clientes'].empty else [])
        v_ruc_c = ""
        if not repo['clientes'].empty and v_cli:
            try:
                row_c = repo['clientes'][repo['clientes']['EMPRESA']==v_cli].iloc[0]
                v_ruc_c = str(row_c['RUC'])
            except: pass
        st.caption(f"RUC: {v_ruc_c}")
        v_serv = st.selectbox("Servicio", repo['servicios'].iloc[:,1].unique() if not repo['servicios'].empty else [])
        v_res = st.selectbox("Residuo", repo['servicios'].iloc[:,2].unique() if not repo['servicios'].empty else [])

    dest_final = v_dest if "EPMI" not in str(v_dest).upper() else "EPMI S.A.C."

    st.divider()
    tab1, tab2 = st.tabs(["1️⃣ Generar", "2️⃣ Registrar"])

    with tab1:
        c_gen, c_exc = st.columns(2)
        with c_gen:
            if st.button("📄 Generar Word", type="primary"):
                drive, _ = obtener_servicios()
                if drive:
                    try:
                        id_p = PLANTILLAS[empresa_firma][tipo_plantilla]
                        req = drive.files().export_media(fileId=id_p, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                        fh = io.BytesIO()
                        dl = MediaIoBaseDownload(fh, req)
                        done = False
                        while not done: _, done = dl.next_chunk()
                        
                        doc = DocxTemplate(io.BytesIO(fh.getvalue()))
                        ctx = {
                            "CORRELATIVO": v_corr, "TITULO": v_tit, "REGISTRO": v_reg_e,
                            "EMPRESA": v_emi, "RUC_EMPRESA": v_ruc_e, "RUC": v_ruc_e, 
                            "CLIENTE": v_cli, "RUC_CLIENTE": v_ruc_c, "RAZON_SOCIAL_CLIENTE": v_cli,
                            "SERVICIO_O_COMPRA": v_serv, "TIPO_DE_RESIDUO": v_res,
                            "PUNTO_PARTIDA": v_partida, "DIRECCION_EMPRESA": v_llegada, 
                            "EMPRESA_2": dest_final, "FECHA_EMISION": v_fec_emis, "items": [] 
                        }
                        doc.render(ctx)
                        buf = io.BytesIO()
                        doc.save(buf)
                        st.session_state['word_buffer'] = buf.getvalue()
                        
                        peso_t = sum([float(str(x).replace(',','')) for x in v_items['peso'] if str(x).replace(',','').replace('.','').isdigit()])
                        
                        name_safe = f"{empresa_firma} - {tipo_operacion_simple} - {v_corr}".replace("/", "-")
                        st.session_state['nombre_archivo_final'] = name_safe
                        st.session_state['datos_log_pendientes'] = {
                            "fec_emis": v_fec_emis, "emi": v_emi, "tit": tipo_operacion_simple, 
                            "cli": v_cli, "ruc_c": v_ruc_c, "guia": v_guia, "res": v_res,
                            "cert_name": name_safe, "peso": peso_t              
                        }
                        st.success("✅ Generado")
                    except Exception as e: st.error(f"Error: {e}")

            if 'word_buffer' in st.session_state:
                fn = st.session_state.get('nombre_archivo_final', "Borrador")
                st.download_button("📩 Bajar Word", st.session_state['word_buffer'], f"{fn}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        with c_exc:
            df_x = st.session_state['df_items'].copy()
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as w: df_x.to_excel(w, index=False)
            fn_x = st.session_state.get('nombre_archivo_final', "Tabla")
            st.download_button("📊 Bajar Excel", out.getvalue(), f"Tabla {fn_x}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        u_d = st.text_input("Link DOC:")
        u_p = st.text_input("Link PDF:")
        if st.button("🏁 Registrar"):
            if not st.session_state.get('datos_log_pendientes') or not u_d or not u_p:
                st.warning("⚠️ Faltan datos")
            else:
                d = st.session_state['datos_log_pendientes']
                f = [d['fec_emis'], d['emi'], d['tit'], d['cli'], d['ruc_c'], d['guia'], "FINALIZADO", d['cert_name'], u_d, u_p]
                if registrar_en_control(f): st.success("✅ Registrado"); st.balloons()

st.caption("--- FIN DEL SISTEMA V4.5 (LECTURA PRECISA) ---")
