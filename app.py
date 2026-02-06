import streamlit as st
import pandas as pd
import google.generativeai as genai
import json
import io
import base64
import re
import time
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
import docx.oxml.simpletypes
from docx.oxml.simpletypes import ST_TwipsMeasure, Twips
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

# ==========================================
# 0. PATCHES & SETUP
# ==========================================

# Monkey patch para corregir error de parsing de floats en Twips
original_convert_from_xml = ST_TwipsMeasure.convert_from_xml

@classmethod
def patch_convert_from_xml(cls, str_value):
    try:
        return Twips(int(str_value))
    except ValueError:
        try:
            return Twips(int(float(str_value)))
        except:
            return original_convert_from_xml(str_value)

ST_TwipsMeasure.convert_from_xml = patch_convert_from_xml

load_dotenv()

st.set_page_config(page_title="Sistema Certificados", layout="wide")

# ==========================================
# 1. CREDENCIALES Y CONSTANTES
# ==========================================
API_KEY = os.getenv("GEMINI_API_KEY")
if not API_KEY and "GEMINI_API_KEY" in st.secrets:
    API_KEY = st.secrets["GEMINI_API_KEY"]

if not API_KEY:
    try:
        if "gcp_service_account" in st.secrets and "GEMINI_API_KEY" in st.secrets["gcp_service_account"]:
            API_KEY = st.secrets["gcp_service_account"]["GEMINI_API_KEY"]
    except: pass

if not API_KEY:
    st.error("🚨 ERROR: Falta API KEY. Configura el archivo .env o st.secrets.")
    st.stop()

# IDs Google
ID_SHEET_REPOSITORIO = "14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU"
ID_SHEET_CONTROL = "14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU" 
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")
if not DRIVE_FOLDER_ID and "DRIVE_FOLDER_ID" in st.secrets:
    DRIVE_FOLDER_ID = st.secrets["DRIVE_FOLDER_ID"]

PLANTILLAS = {
    "EPMI S.A.C.": {
        "Comercialización/Disposición Final": os.getenv("TEMPLATE_EPMI_ID", "1d09vmlBlW_4yjrrz5M1XM8WpCvzTI4f11pERDbxFvNE"),
        "Peligroso y No Peligroso": os.getenv("TEMPLATE_EPMI_PELIGROSO_ID", "1QqqVJ2vCiAjiKKGt_zEpaImUB-q3aRurSiXjMEU--eg")
    },
    "INECOVE S.A.C.": {
        "Comercialización/Disposición Final": os.getenv("TEMPLATE_INECOVE_ID", "1MPzCwxR538osP3_br4VrTDybplqpTBtB08Jo"),
        "Peligroso y No Peligroso": os.getenv("TEMPLATE_INECOVE_PELIGROSO_ID", "1W-HyVSivqug13gBRBclBuICAOSBUHm1WN5cnqtMQcZY")
    }
}

# ==========================================
# 2. CONEXIÓN GOOGLE DRIVE/SHEETS
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
    except: return None, None

def registrar_en_control(datos_fila):
    _, sheets = obtener_servicios()
    if not sheets: return False
    try:
        sheets.spreadsheets().values().append(
            spreadsheetId=ID_SHEET_CONTROL, range="'historial'!A:J",
            valueInputOption="USER_ENTERED", body={"values": [datos_fila]}
        ).execute()
        return True
    except: return False

def subir_a_drive(contenido_bytes, nombre_archivo):
    drive, _ = obtener_servicios()
    if not drive: return None
    try:
        file_metadata = {'name': f"{nombre_archivo}.docx", 'mimeType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}
        if DRIVE_FOLDER_ID:
            file_metadata['parents'] = [DRIVE_FOLDER_ID]
        
        media = MediaIoBaseUpload(io.BytesIO(contenido_bytes), mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', resumable=True)
        file = drive.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        return file.get('webViewLink')
    except Exception as e:
        st.error(f"Error subiendo a Drive: {e}")
        return None

# ==========================================
# 3. FORMATOS Y UTILIDADES
# ==========================================
def limpiar_monto(valor):
    """
    Convierte string a float manejando comas y puntos extra.
    Regla: 
    1. Comas se borran.
    2. Si hay >1 punto, se mantienen solo el ultimo (el decimal).
    """
    if not valor: return 0.0
    s = str(valor).strip()
    
    # 1. Eliminar comas siempre
    s = s.replace(',', '')
    
    # 2. Manejo de multiples puntos (ej: 3.580.00 -> 3580.00)
    if s.count('.') > 1:
        parts = s.split('.')
        # Unir todo menos el último con nada, y pegar el último con punto
        s = "".join(parts[:-1]) + '.' + parts[-1]
    
    # 3. Limpieza final caracteres no numéricos
    s = re.sub(r'[^\d.]', '', s)
    try:
        return float(s)
    except:
        return 0.0

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

# ==========================================
# 4. LOGICA DOCX (INYECCIÓN DE TABLA)
# ==========================================
def set_borders(table):
    """Fallback para bordes manuales"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4') 
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        borders.append(border)
    tblPr.append(borders)

def set_cell_background(cell, color_hex):
    """Establece color de fondo de celda"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_table_margins(table, top=0, bottom=0, left=10, right=10):
    """Reduce márgenes internos de celdas"""
    tblPr = table._tbl.tblPr
    tblCellMar = parse_xml(f'''
    <w:tblCellMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:top w:w="{top}" w:type="dxa"/>
        <w:left w:w="{left}" w:type="dxa"/>
        <w:bottom w:w="{bottom}" w:type="dxa"/>
        <w:right w:w="{right}" w:type="dxa"/>
    </w:tblCellMar>
    ''')
    tblPr.append(tblCellMar)

def inyectar_tabla_en_docx(doc_io, data_items, servicio_global):
    """
    Recibe un BytesIO con el DOCX ya renderizado por docxtpl.
    Busca el marcador [[TABLA_NOTAS]] y lo reemplaza por una tabla real usando python-docx.
    """
    doc = Document(doc_io)
    
    # Buscar el párrafo con el marcador
    target_paragraph = None
    for p in doc.paragraphs:
        if '[[TABLA_NOTAS]]' in p.text:
            target_paragraph = p
            break
            
    if target_paragraph:
        # Limpiar el texto del marcador
        target_paragraph.text = target_paragraph.text.replace('[[TABLA_NOTAS]]', '')
        
        # Crear tabla y aplicar estilo/bordes
        table = doc.add_table(rows=1, cols=7)
        try:
            table.style = 'Table Grid'
        except:
            set_borders(table)
            
        # Ajustar ancho de tabla y columnas
        table.autofit = False
        table.allow_autofit = False
        
        # Padding de celdas: 0.05 pulgadas para equilibrio vertical
        # 0.05 inches * 1440 = 72 twips
        set_table_margins(table, top=72, bottom=72, left=30, right=30)

        # Anchos PROPORCIONALES (Total 7.5")
        # Fecha, Placa, Guia, Cant, Medida, Peso = 10% (0.75")
        # Descripcion = 40% (3.0")
        widths = [Inches(0.75), Inches(0.75), Inches(0.75), Inches(3.0), Inches(0.75), Inches(0.75), Inches(0.75)]
        
        for i, col in enumerate(table.columns):
            col.width = widths[i]
        
        # Encabezados
        encabezados = ['Fecha', 'Placa', 'N° Guía', 'Descripción', 'Cantidad', 'Medida', 'Peso']
        hdr_cells = table.rows[0].cells
        for i, nombre in enumerate(encabezados):
            cell = hdr_cells[i]
            cell.text = nombre
            cell.width = widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            # Fondo Verde Estilo Sheets (#70ad47)
            set_cell_background(cell, "70ad47")
            
            # Centrar todos los párrafos
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # LIMPIEZA VERTICAL STRICTA
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1
                
                if p.runs:
                    run = p.runs[0]
                else:
                    run = p.add_run(nombre)
                    
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0) # Negro
                run.font.name = 'Calibri'
                run.font.size = Pt(9)
        
        # Datos
        for item in data_items:
            row_cells = table.add_row().cells
            
            # Preparar valores
            vals = [
                str(item.get('fecha_origen', '')),
                str(item.get('placa_origen', '')),
                str(item.get('guia_origen', '')),
                str(item.get('desc', '')),
                str(item.get('cant', '')),
                str(item.get('um', '')),
                f"{str(item.get('peso', '0'))} kg" if not 'kg' in str(item.get('peso', '0')).lower() else str(item.get('peso', '0'))
            ]
            
            for idx, valor in enumerate(vals):
                cell = row_cells[idx]
                cell.text = valor
                cell.width = widths[idx]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                
                # Centrar todos los párrafos
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    # LIMPIEZA VERTICAL STRICTA
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1
                    
                    if p.runs:
                        run = p.runs[0]
                        run.font.name = 'Calibri'
                        run.font.size = Pt(9)
                    else: 
                        run = p.add_run(valor) # Fallback
                        run.font.name = 'Calibri'
                        run.font.size = Pt(9)

        # Mover tabla después del párrafo
        tbl, p = table._tbl, target_paragraph._p
        p.addnext(tbl)

    new_buffer = io.BytesIO()
    doc.save(new_buffer)
    return new_buffer.getvalue()

# ==========================================
# 5. INTELIGENCIA ARTIFICIAL (GEMINI)
# ==========================================
def procesar_guia_ia(pdf_bytes):
    try:
        genai.configure(api_key=API_KEY.strip())
    except: return None

    modelo = "gemini-flash-latest"

    prompt = """
    Actúa como experto OCR y extrae los datos de esta Guía de Remisión a JSON:
    {
        "fecha": "dd/mm/yyyy", 
        "serie": "T001-000000", 
        "vehiculo": "PLACA", 
        "punto_partida": "Dirección completa", 
        "punto_llegada": "Dirección completa", 
        "destinatario": "Razón Social", 
        "items": [
            {
                "desc": "Descripción del bien", 
                "cant": "Número (ej: 100)", 
                "um": "Unidad (UND, NIU)", 
                "peso": "Peso (ej: 500.00)"
            }
        ]
    }
    """

    try:
        time.sleep(2) 
        model = genai.GenerativeModel(modelo)
        res = model.generate_content([prompt, {"mime_type": "application/pdf", "data": base64.b64encode(pdf_bytes).decode('utf-8')}])
        
        texto_limpio = res.text.replace("```json", "").replace("```", "")
        match = re.search(r'\{.*\}', texto_limpio, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        else:
            return None

    except Exception as e:
        st.error(f"Error IA: {e}")
        return None

# ==========================================
# 6. INTERFAZ STREAMLIT
# ==========================================
if 'ocr_data' not in st.session_state: st.session_state['ocr_data'] = None
if 'df_items' not in st.session_state: st.session_state['df_items'] = pd.DataFrame()
if 'datos_log_pendientes' not in st.session_state: st.session_state['datos_log_pendientes'] = {}

with st.sidebar:
    st.header("Configuración")
    empresa_firma = st.selectbox("Empresa", list(PLANTILLAS.keys()))
    tipo_plantilla = st.selectbox("Plantilla", ["Comercialización/Disposición Final", "Peligroso y No Peligroso"])
    if st.button("Recargar"): st.cache_data.clear(); st.rerun()

st.title("Generador de Certificados (V8.1 - Tabla Final)")

if 'repo_data' not in st.session_state:
    st.session_state['repo_data'] = {
        "emisores": leer_sheet_seguro("EMPRESAS"),
        "clientes": leer_sheet_seguro("CLIENTES"),
        "servicios": leer_sheet_seguro("COMERCIALIZACION")
    }
repo = st.session_state['repo_data']

archivos = st.file_uploader("Subir Guías (PDF)", type=["pdf"], accept_multiple_files=True)

if archivos:
    if st.button("🔍 Procesar"):
        prog = st.progress(0)
        items, grl = [], None
        errores = 0
        total = len(archivos)
        
        for i, arc in enumerate(archivos):
            if i > 0: time.sleep(3) # Pausa
            
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
        else: st.error("❌ Falló el proceso. Revisa los mensajes de error.")

if st.session_state['ocr_data']:
    ocr = st.session_state['ocr_data']
    st.markdown("### Validación")
    c1, c2, c3, c4 = st.columns(4)
    v_corr = c1.text_input("Correlativo", "001")
    fecha_base = normalizar_fecha(ocr.get('fecha'))
    cont_f = c2.container()
    
    opt_f = c2.radio("Regla Fecha:", ["Comercialización (+2)", "Disposición Final (Fin de Mes)"], label_visibility="collapsed")
    try:
        if "Comercialización" in opt_f: 
            f_calc = (datetime.strptime(fecha_base, "%d/%m/%Y")+timedelta(days=2)).strftime("%d/%m/%Y")
            tipo_operacion_simple = "Comercialización"
        else: 
            f_calc = obtener_fin_de_mes(fecha_base)
            tipo_operacion_simple = "Disposición Final"
    except: 
        f_calc = fecha_base
        tipo_operacion_simple = "Indefinido"

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
        v_tit = st.selectbox("Título", repo['servicios'].iloc[:,0].unique() if not repo['servicios'].empty else [])

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
    tab1, tab2 = st.tabs(["Generar", "Registrar"])

    with tab1:
        if st.button("📄 Generar Word", type="primary"):
            drive, _ = obtener_servicios()
            if drive:
                try:
                    # 1. Descargar plantilla
                    id_p = PLANTILLAS[empresa_firma][tipo_plantilla]
                    req = drive.files().export_media(fileId=id_p, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                    fh = io.BytesIO()
                    dl = MediaIoBaseDownload(fh, req)
                    done = False
                    while not done: _, done = dl.next_chunk()
                    
                    # 2. Renderizar Docxtpl
                    doc = DocxTemplate(io.BytesIO(fh.getvalue()))
                    ctx = {
                        "CORRELATIVO": v_corr, "TITULO": v_tit, "REGISTRO": v_reg_e,
                        "EMPRESA": v_emi, "RUC_EMPRESA": v_ruc_e, "RUC": v_ruc_e, 
                        "CLIENTE": v_cli, "RUC_CLIENTE": v_ruc_c, "RAZON_SOCIAL_CLIENTE": v_cli,
                        "SERVICIO_O_COMPRA": v_serv, "TIPO_DE_RESIDUO": v_res,
                        "PUNTO_PARTIDA": v_partida, "DIRECCION_EMPRESA": v_llegada, 
                        "EMPRESA_2": dest_final, "FECHA_EMISION": v_fec_emis
                    }
                    doc.render(ctx)
                    buf_tpl = io.BytesIO()
                    doc.save(buf_tpl)

                    # 3. Inyectar Tabla
                    items_para_tabla = st.session_state['df_items'].to_dict('records')
                    final_bytes = inyectar_tabla_en_docx(io.BytesIO(buf_tpl.getvalue()), items_para_tabla, v_serv)
                    
                    st.session_state['word_buffer'] = final_bytes
                    
                    # Calcular Peso Total Seguro
                    peso_t = sum([limpiar_monto(x) for x in v_items['peso']])
                    name_safe = f"{empresa_firma} - {tipo_operacion_simple} - {v_corr}".replace("/", "-")
                    st.session_state['nombre_archivo_final'] = name_safe
                    
                    # 4. Subir a Drive automáticamente
                    link_drive = subir_a_drive(final_bytes, name_safe)
                    st.session_state['link_drive_generado'] = link_drive

                    st.session_state['datos_log_pendientes'] = {
                        "fec_emis": v_fec_emis, "emi": v_emi, "tit": tipo_operacion_simple, 
                        "cli": v_cli, "ruc_c": v_ruc_c, "guia": v_guia, "res": v_res,
                        "cert_name": name_safe, "peso": peso_t              
                    }
                    
                    if link_drive:
                        st.success(f"✅ Generado y Subido a Drive: {name_safe}")
                        st.markdown(f"[📂 Abrir en Drive]({link_drive})")
                    else:
                        st.success("✅ Generado (Subida a Drive falló, revisa credenciales)")

                except Exception as e: st.error(f"Error: {e}")
            else:
                st.error("No se pudo conectar con Google Drive.")

        if 'word_buffer' in st.session_state:
            fn = st.session_state.get('nombre_archivo_final', "Borrador")
            st.download_button("📩 Bajar Copia Local", st.session_state['word_buffer'], f"{fn}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    with tab2:
        u_d = st.text_input("Link DOC:", value=st.session_state.get('link_drive_generado', ''))
        u_p = st.text_input("Link PDF:")
        if st.button("🏁 Registrar"):
            if not st.session_state.get('datos_log_pendientes') or not u_d or not u_p:
                st.warning("⚠️ Faltan datos (Link Doc/PDF o generar primero)")
            else:
                d = st.session_state['datos_log_pendientes']
                f = [d['fec_emis'], d['emi'], d['tit'], d['cli'], d['ruc_c'], d['guia'], "FINALIZADO", d['cert_name'], u_d, u_p]
                if registrar_en_control(f): st.success("✅ Registrado en Sheets"); st.balloons()
