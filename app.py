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
# 1. CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="Sistema Certificados", layout="wide")

# ==========================================
# 2. GESTIÓN DE CREDENCIALES (API KEY)
# ==========================================
API_KEY = None
try:
    if "GEMINI_API_KEY" in st.secrets:
        API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    pass

if not API_KEY:
    try:
        if "gcp_service_account" in st.secrets and "GEMINI_API_KEY" in st.secrets["gcp_service_account"]:
            API_KEY = st.secrets["gcp_service_account"]["GEMINI_API_KEY"]
    except:
        pass

if not API_KEY:
    API_KEY = "FALTA_CONFIGURAR"

# IDs de las Hojas de Cálculo (Base de Datos)
ID_SHEET_REPOSITORIO = "14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU"
ID_SHEET_CONTROL = "14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU" 

# IDs de los archivos de Plantillas de Word en Google Drive
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
# 3. FUNCIONES DE CONEXIÓN A GOOGLE (DRIVE/SHEETS)
# ==========================================
def obtener_servicios():
    """Conecta con Google Drive y Sheets usando las credenciales."""
    scopes = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
    creds = None
    
    # Intentar cargar desde Streamlit Secrets
    try:
        if "gcp_service_account" in st.secrets:
            creds = service_account.Credentials.from_service_account_info(
                st.secrets["gcp_service_account"], scopes=scopes
            )
    except:
        pass

    # Intentar cargar desde archivo local (si estamos en PC)
    if not creds:
        try:
            creds = service_account.Credentials.from_service_account_file('secretos.json', scopes=scopes)
        except Exception as e:
            return None, None

    try:
        service_drive = build('drive', 'v3', credentials=creds)
        service_sheets = build('sheets', 'v4', credentials=creds)
        return service_drive, service_sheets
    except Exception as e:
        st.error(f"Error conectando servicios Google: {e}")
        return None, None

def registrar_en_control(datos_fila):
    """Guarda una fila nueva en la pestaña 'historial'."""
    _, service_sheets = obtener_servicios()
    if not service_sheets:
        return False
    try:
        service_sheets.spreadsheets().values().append(
            spreadsheetId=ID_SHEET_CONTROL,
            range="'historial'!A:J",
            valueInputOption="USER_ENTERED",
            body={"values": [datos_fila]}
        ).execute()
        return True
    except Exception as e:
        st.error(f"Error guardando en Excel (Pestaña historial): {e}")
        return False

# ==========================================
# 4. FUNCIONES DE LIMPIEZA Y FORMATO
# ==========================================
def obtener_fin_de_mes(fecha_str):
    """Calcula el último día del mes."""
    try:
        dt = datetime.strptime(fecha_str, "%d/%m/%Y")
        next_month = dt.replace(day=28) + timedelta(days=4)
        res = next_month - timedelta(days=next_month.day)
        return res.strftime("%d/%m/%Y")
    except:
        return fecha_str

def limpiar_descripcion(texto):
    """Elimina prefijos basura de la descripción."""
    if not texto: return ""
    texto_str = str(texto).strip()
    # Eliminar 'VEN - AMB - ' y variantes
    texto_limpio = re.sub(r'VEN\s*-\s*AMB\s*-\s*', '', texto_str, flags=re.IGNORECASE)
    return texto_limpio.strip()

def formato_nompropio(texto):
    """Convierte texto a Título (Mayúscula inicial)."""
    if not texto: return ""
    return str(texto).strip().title()

def normalizar_fecha(fecha_str):
    """Intenta convertir cualquier formato de fecha a dd/mm/yyyy."""
    if not fecha_str: return datetime.now().strftime("%d/%m/%Y")
    formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"]
    for fmt in formatos:
        try:
            return datetime.strptime(fecha_str.strip(), fmt).strftime("%d/%m/%Y")
        except:
            continue
    return fecha_str 

def formatear_guia(serie_str):
    """Formatea la serie de la guía (quita ceros extra)."""
    if not serie_str or '-' not in str(serie_str): return serie_str
    try:
        partes = str(serie_str).split('-')
        if len(partes) == 2:
            return f"{partes[0].strip()}-{str(int(partes[1].strip()))}"
    except:
        pass
    return serie_str

@st.cache_data(show_spinner=False, ttl=10)
def leer_sheet_seguro(pestaña):
    """Lee datos de Google Sheets de forma segura."""
    _, service_sheets = obtener_servicios()
    if not service_sheets: return pd.DataFrame()
    try:
        result = service_sheets.spreadsheets().values().get(
            spreadsheetId=ID_SHEET_REPOSITORIO, range=f"'{pestaña}'!A1:Z1000"
        ).execute()
        values = result.get('values', [])
        if not values: return pd.DataFrame()
        return pd.DataFrame(values[1:], columns=values[0])
    except Exception as e:
        st.warning(f"No se pudo leer la pestaña '{pestaña}': {e}")
        return pd.DataFrame()

# ==========================================
# 5. MOTOR DE INTELIGENCIA ARTIFICIAL (GEMINI)
# ==========================================
def procesar_guia_ia(pdf_bytes):
    """Envía el PDF a Gemini para extraer datos."""
    
    # 1. Configuración de API
    try:
        if "FALTA" in API_KEY or API_KEY is None:
            st.error("⚠️ ERROR: No se detectó la API KEY. Revisa los secretos.")
            return None
        genai.configure(api_key=API_KEY.strip())
    except Exception as e:
        st.error(f"❌ Error Configuración API: {e}")
        return None

    # 2. Selección Inteligente del Modelo (Para evitar error 404)
    model = None
    lista_modelos_visibles = []
    
    try:
        # Obtenemos la lista real de modelos disponibles
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                lista_modelos_visibles.append(m.name)
        
        # Buscamos el mejor candidato (FLASH 1.5)
        candidato = next((m for m in lista_modelos_visibles if 'flash' in m and '1.5' in m), None)
        
        # Si no está, buscamos cualquier Flash
        if not candidato:
            candidato = next((m for m in lista_modelos_visibles if 'flash' in m), None)
            
        # Si no, usamos el Pro 1.5
        if not candidato:
            candidato = next((m for m in lista_modelos_visibles if 'pro' in m and '1.5' in m), None)

        if candidato:
            model = genai.GenerativeModel(candidato)
        else:
            st.warning(f"⚠️ No encontré modelos compatibles. Disponibles: {lista_modelos_visibles}")
            return None

    except Exception as e:
        st.error(f"❌ Error buscando modelos: {e}")
        return None

    # 3. Prompt (Instrucciones) MEJORADO PARA ORDENAR DATOS
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

    REGLAS DE ORDENAMIENTO (CRÍTICO):
    1. **Items y Columnas:** - Identifica correctamente las columnas de la tabla. A veces 'Cantidad' y 'Peso' están juntas.
       - Si ves una columna con valores pequeños (ej: 10, 50) y unidades como 'UND', 'BULTOS', 'NIU' -> Eso es "cant".
       - Si ves una columna con valores más grandes o decimales y unidades como 'KGM', 'KG' -> Eso es "peso".
       - Si el peso parece ser 0, busca en columnas adyacentes.

    2. **Observaciones:**
       - Revisa el campo "OBSERVACIONES". Si menciona un lugar específico (ej: "FUNDO...", "PLANTA...", "POZO..."), extráelo.
       - Agrégalo al final de 'punto_partida' separado por un guion " - ".
       - Ignora textos sobre devolución de envases.

    3. **Datos Generales:**
       - Fecha: Formato dd/mm/yyyy.
       - Serie: Formato XXXX-XXXXXXX.

    Responde SOLO con el JSON.
    """
    
    # 4. Ejecución con Reintentos (Para evitar error de velocidad)
    max_intentos = 3
    for intento in range(max_intentos):
        try:
            res = model.generate_content([prompt, {"mime_type": "application/pdf", "data": base64.b64encode(pdf_bytes).decode('utf-8')}])
            
            # Limpieza de respuesta para obtener solo el JSON
            texto_limpio = res.text.replace("```json", "").replace("```", "")
            match = re.search(r'\{.*\}', texto_limpio, re.DOTALL)
            
            if match:
                return json.loads(match.group(0))
            else:
                return None
                
        except Exception as e:
            # Si es error de cuota (429), esperamos y reintentamos
            if "429" in str(e):
                if intento < max_intentos - 1:
                    time.sleep(10) # Espera 10 segundos
                    continue 
                else:
                    st.warning("⏳ Se superó el límite de velocidad. Espera un momento.")
                    return None
            else:
                # Otros errores
                st.error(f"❌ Error lectura: {e}")
                return None
    return None

# ==========================================
# 6. INTERFAZ GRÁFICA (FRONTEND)
# ==========================================
# Inicializar variables de estado
if 'ocr_data' not in st.session_state: st.session_state['ocr_data'] = None
if 'df_items' not in st.session_state: st.session_state['df_items'] = pd.DataFrame()
if 'datos_log_pendientes' not in st.session_state: st.session_state['datos_log_pendientes'] = {}

# --- BARRA LATERAL ---
with st.sidebar:
    with st.expander("❓ ¿Cómo usar el sistema?"):
        st.markdown("""
        **1. Sube tus Guías:** Arrastra los PDFs al recuadro principal.
        **2. Procesa:** Dale clic al botón **'🔍 Procesar Guías con IA'**.
        **3. Revisa y Edita:** Verifica los datos en la pantalla.
        **4. Descarga:** Ve a la pestaña **'1️⃣ Generar Materiales'**.
        **5. Registra:** Pega el link final en **'2️⃣ Registrar Final'**.
        """)
    st.divider()
    
    st.header("⚙️ Configuración")
    empresa_firma = st.selectbox("Empresa", list(PLANTILLAS.keys()))
    tipo_plantilla = st.selectbox("Plantilla", ["Comercialización/Disposición Final", "Peligroso y No Peligroso"])
    
    if st.button("🔄 Recargar Página"):
        st.cache_data.clear()
        st.rerun()

# --- PANTALLA PRINCIPAL ---
st.title("Generador de Certificados con IA")

# Cargar datos de repositorios
if 'repo_data' not in st.session_state:
    st.session_state['repo_data'] = {
        "emisores": leer_sheet_seguro("EMPRESAS"),
        "clientes": leer_sheet_seguro("CLIENTES"),
        "servicios": leer_sheet_seguro("COMERCIALIZACION")
    }
repo = st.session_state['repo_data']

# --- SECCIÓN 1: CARGA DE ARCHIVOS ---
archivos = st.file_uploader("1. Subir Guías (PDF)", type=["pdf"], accept_multiple_files=True)

if archivos:
    if st.button("🔍 Procesar Guías con IA"):
        prog = st.progress(0)
        items, grl = [], None
        errores = 0
        total_archivos = len(archivos)
        
        for i, arc in enumerate(archivos):
            # Pequeña pausa para no saturar
            if i > 0: time.sleep(2)
            
            d = procesar_guia_ia(arc.read())
            if d:
                if not grl: grl = d
                # Formatos previos
                s = formatear_guia(d.get('serie','S/N'))
                f = d.get('fecha','')
                p = d.get('vehiculo','')
                
                for it in d.get('items', []):
                    it.update({'guia_origen': s, 'fecha_origen': f, 'placa_origen': p})
                    items.append(it)
            else:
                errores += 1
            
            prog.progress((i+1)/total_archivos)
        
        time.sleep(0.5)
        prog.empty()
        
        if grl and items:
            st.session_state['ocr_data'] = grl
            df = pd.DataFrame(items)
            
            # Asegurar columnas
            for c in ['desc','cant','um','peso','fecha_origen','guia_origen','placa_origen']:
                if c not in df.columns: df[c] = ""
            
            # Limpiezas finales
            df['peso'] = df['peso'].replace("", "0.00").replace("None", "0.00")
            df['desc'] = df['desc'].apply(limpiar_descripcion)
            df['fecha_origen'] = df['fecha_origen'].apply(normalizar_fecha)
            
            st.session_state['df_items'] = df
            st.success(f"✅ Éxito: {total_archivos-errores} guías procesadas.")
        else:
            st.error("❌ No se pudieron extraer datos.")

# --- SECCIÓN 2: VALIDACIÓN Y EDICIÓN ---
if st.session_state['ocr_data']:
    ocr = st.session_state['ocr_data']
    st.markdown("### 2. Validación de Datos")
    
    # Fila 1 de inputs
    c1, c2, c3, c4 = st.columns(4)
    v_corr = c1.text_input("Correlativo", "001")
    
    # Lógica de Fechas
    fecha_base = normalizar_fecha(ocr.get('fecha'))
    cont_f = c2.container()
    opt_f = c2.radio("Regla Fecha:", ["Comercialización (+2)", "Disposición Final (Fin de Mes)"], label_visibility="collapsed")
    
    try:
        if "Comercialización" in opt_f: 
            dt_base = datetime.strptime(fecha_base, "%d/%m/%Y")
            f_calc = (dt_base + timedelta(days=2)).strftime("%d/%m/%Y")
            tipo_operacion_simple = "Comercialización"
        else: 
            f_calc = obtener_fin_de_mes(fecha_base)
            tipo_operacion_simple = "Disposición Final"
    except:
        f_calc = fecha_base
        tipo_operacion_simple = "Indefinido"

    v_fec_emis = cont_f.text_input("F. Emisión", value=f_calc)

    if len(archivos) > 1:
        v_guia, v_placa = c3.text_input("Guía", "VARIAS / VER DETALLE"), c4.text_input("Placa", "VARIAS / VER DETALLE")
    else:
        v_guia = c3.text_input("Guía", formatear_guia(ocr.get('serie')))
        v_placa = c4.text_input("Placa", ocr.get('vehiculo'))

    # Fila 2 de inputs
    v_partida = st.text_input("Partida", formato_nompropio(ocr.get('punto_partida','')))
    v_llegada = st.text_input("Llegada", formato_nompropio(ocr.get('punto_llegada','')))
    v_dest = st.text_input("Destinatario", ocr.get('destinatario',''))

    # Tabla Editable
    v_items = st.data_editor(
        st.session_state['df_items'],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "guia_origen": st.column_config.TextColumn("Guía", disabled=True),
            "peso": st.column_config.TextColumn("Peso (Kg)"),
            "cant": st.column_config.TextColumn("Cantidad"),
            "desc": st.column_config.TextColumn("Descripción")
        }
    )
    
    # Selectores finales
    c_a, c_b = st.columns(2)
    with c_a:
        lista_emisores = repo['emisores']['EMPRESA'].unique() if not repo['emisores'].empty else []
        v_emi = st.selectbox("Emisor", lista_emisores)
        v_ruc_e, v_reg_e = "", ""
        if not repo['emisores'].empty and v_emi:
            try:
                row_e = repo['emisores'][repo['emisores']['EMPRESA']==v_emi].iloc[0]
                v_ruc_e, v_reg_e = str(row_e['RUC']), str(row_e['REGISTRO'])
            except: pass
        st.caption(f"RUC: {v_ruc_e} | REG: {v_reg_e}")
        
        lista_titulos = repo['servicios'].iloc[:,0].unique() if not repo['servicios'].empty else []
        v_tit = st.selectbox("Título del Certificado", lista_titulos)

    with c_b:
        lista_clientes = repo['clientes']['EMPRESA'].unique() if not repo['clientes'].empty else []
        v_cli = st.selectbox("Cliente", lista_clientes)
        v_ruc_c = ""
        if not repo['clientes'].empty and v_cli:
            try:
                row_c = repo['clientes'][repo['clientes']['EMPRESA']==v_cli].iloc[0]
                v_ruc_c = str(row_c['RUC'])
            except: pass
        st.caption(f"RUC: {v_ruc_c}")
        
        lista_servicios = repo['servicios'].iloc[:,1].unique() if not repo['servicios'].empty else []
        v_serv = st.selectbox("Servicio", lista_servicios)
        
        lista_residuos = repo['servicios'].iloc[:,2].unique() if not repo['servicios'].empty else []
        v_res = st.selectbox("Residuo", lista_residuos)

    dest_final = v_dest if "EPMI" not in str(v_dest).upper() else "EPMI S.A.C."

    st.divider()
    
    # --- PESTAÑAS DE ACCIÓN ---
    tab1, tab2 = st.tabs(["1️⃣ Generar Materiales", "2️⃣ Registrar Final"])

    with tab1:
        st.info("ℹ️ Genera y descarga los archivos.")
        col_btn_1, col_btn_2 = st.columns(2)
        
        # Generar WORD
        with col_btn_1:
            if st.button("📄 Generar Word (Borrador)", type="primary"):
                service_drive, _ = obtener_servicios()
                if service_drive:
                    try:
                        id_p = PLANTILLAS[empresa_firma][tipo_plantilla]
                        
                        # Descargar plantilla
                        request = service_drive.files().export_media(
                            fileId=id_p,
                            mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                        )
                        fh = io.BytesIO()
                        downloader = MediaIoBaseDownload(fh, request)
                        done = False
                        while not done:
                            _, done = downloader.next_chunk()
                        
                        # Renderizar plantilla
                        doc = DocxTemplate(io.BytesIO(fh.getvalue()))
                        
                        context = {
                            "CORRELATIVO": v_corr,
                            "TITULO": v_tit,
                            "REGISTRO": v_reg_e,
                            "EMPRESA": v_emi,
                            "RUC_EMPRESA": v_ruc_e,
                            "RUC": v_ruc_e, 
                            "CLIENTE": v_cli,
                            "RUC_CLIENTE": v_ruc_c,
                            "RAZON_SOCIAL_CLIENTE": v_cli,
                            "SERVICIO_O_COMPRA": v_serv,
                            "TIPO_DE_RESIDUO": v_res,
                            "PUNTO_PARTIDA": v_partida,
                            "DIRECCION_EMPRESA": v_llegada, 
                            "EMPRESA_2": dest_final,
                            "FECHA_EMISION": v_fec_emis,
                            "items": [] # La tabla se llena fuera, esto es por compatibilidad
                        }
                        doc.render(context)
                        
                        # Guardar en buffer
                        buf = io.BytesIO()
                        doc.save(buf)
                        st.session_state['word_buffer'] = buf.getvalue()
                        
                        # Calcular peso total
                        peso_t = 0.0
                        for x in v_items['peso']:
                            try:
                                peso_t += float(str(x).replace(',',''))
                            except: pass
                        
                        # Preparar nombre y datos log
                        nombre_certificado_completo = f"{empresa_firma} - {tipo_operacion_simple} - {v_corr}"
                        nombre_archivo_safe = nombre_certificado_completo.replace("/", "-").replace("\\", "-")
                        
                        st.session_state['nombre_archivo_final'] = nombre_archivo_safe
                        st.session_state['datos_log_pendientes'] = {
                            "fec_emis": v_fec_emis, "emi": v_emi, "tit": tipo_operacion_simple, 
                            "cli": v_cli, "ruc_c": v_ruc_c, "guia": v_guia, "res": v_res,
                            "cert_name": nombre_certificado_completo, "peso": peso_t              
                        }
                        st.success("✅ Word generado correctamente.")
                    except Exception as e:
                        st.error(f"Error generando Word: {e}")

            if 'word_buffer' in st.session_state:
                fname = st.session_state.get('nombre_archivo_final', f"Borrador_{v_corr}")
                st.download_button(
                    label="📩 Descargar Word",
                    data=st.session_state['word_buffer'],
                    file_name=f"{fname}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        # Generar EXCEL
        with col_btn_2:
            df_ex = pd.DataFrame()
            n_filas = len(v_items)
            
            # Construir Excel de Items
            df_ex['Fecha'] = v_items.get('fecha_origen', [fecha_base]*n_filas)
            df_ex['Vehículo'] = v_items.get('placa_origen', [v_placa]*n_filas)
            df_ex['Guía'] = v_items.get('guia_origen', [v_guia]*n_filas)
            df_ex['Descripción'] = v_items['desc']
            df_ex['Cantidad'] = v_items['cant']
            df_ex['U.M.'] = v_items['um']
            df_ex['Peso (Kg)'] = v_items['peso']
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_ex.to_excel(writer, index=False)
            
            fname_excel = st.session_state.get('nombre_archivo_final', f"Tabla_{v_corr}")
            st.download_button(
                label="📊 Descargar Excel",
                data=output.getvalue(),
                file_name=f"Tabla {fname_excel}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # --- PESTAÑA REGISTRO ---
    with tab2:
        url_doc = st.text_input("🔗 Link del Documento Final (DOC/PDF):")
        url_pdf = st.text_input("🔗 Link de las Guías Escaneadas (PDF):")
        
        if st.button("🏁 Guardar Registro en Historial"):
            if not st.session_state.get('datos_log_pendientes') or not url_doc or not url_pdf:
                st.warning("⚠️ Faltan datos (Genera el Word primero) o faltan los links.")
            else:
                d_log = st.session_state['datos_log_pendientes']
                fila_excel = [
                    d_log['fec_emis'],
                    d_log['emi'],
                    d_log['tit'],
                    d_log['cli'], 
                    d_log['ruc_c'],
                    d_log['guia'],
                    "FINALIZADO", 
                    d_log['cert_name'],
                    url_doc,
                    url_pdf
                ]
                
                if registrar_en_control(fila_excel):
                    st.balloons()
                    st.success("✅ ¡Registro exitoso! La información se guardó en la pestaña 'historial'.")
                else:
                    st.error("Hubo un error al guardar en el Excel de control.")

st.divider()
st.caption("Sistema de Certificados - Versión Estable Completa")
