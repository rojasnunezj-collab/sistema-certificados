# ====================================================================
# --- BLOQUE 1: Imports y Configuración Inicial ---
# ====================================================================
import streamlit as st
import pandas as pd
import io
import os
import time
from datetime import datetime, timedelta

# Importamos ÚNICAMENTE el motor de Vertex (eliminamos la función vieja)
from src.services.vertex_service import procesar_guia_ia_vertex

from src.services.google_service import obtener_servicios, registrar_en_control, leer_sheet_seguro, subir_a_drive
from src.config.settings import PLANTILLAS
from src.utils.document_utils import inyectar_tabla_en_docx
from src.utils.format_utils import (
    limpiar_monto, formato_inteligente, normalizar_fecha, 
    limpiar_descripcion, formatear_guia, obtener_fin_de_mes,
    formato_nompropio
)
from docxtpl import DocxTemplate
from googleapiclient.http import MediaIoBaseDownload

# --- Configuración Inicial ---
st.set_page_config(page_title="Certificador AI V2", layout="wide")

if 'datos_extraidos' not in st.session_state:
    st.session_state.datos_extraidos = None

if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0
# ====================================================================
# --- BLOQUE 2: UI - Setup Inicial y Sidebar ---
# ====================================================================
# --- UI: Encabezado ---
st.title("📄 Generador de Certificados")
st.info("Plataforma optimizada con Vertex AI Gemini")

# --- Barra Lateral (Configuración) ---
with st.sidebar:
    st.header("Configuración de Flujo")
    empresa_firma = st.selectbox("Empresa Firmanente", list(PLANTILLAS.keys()))
    
    # AQUÍ ESTÁ EL CAMBIO: Desglosamos en las 3 opciones exactas
    tipo_flujo = st.selectbox(
        "Tipo de Certificado",
        ["Comercialización", "Disposición Final 1", "Disposición Final 2"]
    )
    
    modo_manual = st.toggle("Llenado Manual (Sin PDF)", value=False)
    
    if st.button("Limpiar Sesión", use_container_width=True):
        # 1. Borramos todas las variables de la memoria (¡sin piedad!)
        for key in ['ocr_data', 'df_items', 'datos_extraidos', 'extraccion']:
            if key in st.session_state:
                del st.session_state[key]
        
        # 2. Rotamos la llave para obligar a Streamlit a destruir el componente del PDF
        st.session_state.uploader_key += 1
        
        # 3. Recargamos la página al instante
        st.rerun()

# ====================================================================
# --- BLOQUE 3: UI - Ingesta y Procesamiento de Archivos ---
# ====================================================================
if not modo_manual:
    archivos = st.file_uploader("Sube tus guías", type=["pdf"], accept_multiple_files=True, key=f"uploader_{st.session_state.uploader_key}")

    if archivos:
        if st.button("🔍 Procesar"):
            prog = st.progress(0)
            items, grl = [], None
            errores = 0
            total = len(archivos)
            
            for i, arc in enumerate(archivos):
                d = procesar_guia_ia_vertex(arc.read())
                if d:
                    if not grl: grl = d 
                    s, f, p = d.get('serie','S/N'), d.get('fecha',''), d.get('vehiculo','')
                    for it in d.get('items', []):
                        it.update({'guia_origen': s, 'fecha_origen': f, 'placa_origen': p})
                        items.append(it)
                else: errores += 1
                prog.progress((i+1)/total)
            
            time.sleep(0.5); prog.empty()
            
            if items:
                # AQUÍ GUARDAMOS EN LA CAJA CORRECTA
                st.session_state['ocr_data'] = grl if grl else {}
                df = pd.DataFrame(items)
                for c in ['desc','cant','um','peso','fecha_origen','guia_origen','placa_origen']:
                    if c not in df.columns: df[c] = ""
                
                # --- LIMPIEZA BASE OBLIGATORIA ---
                df['peso'] = df['peso'].apply(lambda x: formato_inteligente(limpiar_monto(x)))
                df['cant'] = df['cant'].apply(lambda x: formato_inteligente(limpiar_monto(x)))
                df['desc'] = df['desc'].astype(str).str.upper()
                df['um'] = df['um'].apply(lambda x: 'KG' if 'KILO' in str(x).upper() else 'GLN' if 'GALO' in str(x).upper() else 'UNID' if 'UNIDA' in str(x).upper() else str(x).upper())
                df['desc'] = df['desc'].apply(limpiar_descripcion)
                
                def forzar_limpieza_guia(guia_str):
                    s = str(guia_str).strip()
                    if '-' in s:
                        partes = s.split('-', 1)
                        prefijo = partes[0].strip()
                        numeros = ''.join(filter(str.isdigit, partes[1]))
                        if numeros:
                            return f"{prefijo}-{int(numeros)}"
                    return s
                
                df['guia_origen'] = df['guia_origen'].apply(forzar_limpieza_guia)
                df['fecha_origen'] = df['fecha_origen'].apply(normalizar_fecha)

                def fecha_a_entero(fecha_str):
                    try:
                        p = str(fecha_str).strip().split('/')
                        if len(p) == 3: return int(f"{p[2]}{p[1]}{p[0]}")
                    except: pass
                    return 99999999 
                    
                df['_llave_orden'] = df['fecha_origen'].apply(fecha_a_entero)
                df = df.sort_values(by='_llave_orden', ascending=True)
                df = df.drop(columns=['_llave_orden']).reset_index(drop=True)

                st.session_state['df_items'] = df
                st.success(f"✅ Procesado, Limpiado y Ordenado: {len(items)} items de {total} archivos.")
            else: st.error("❌ Falló: No se encontraron items.")

else:
    # --- FASE 3: BYPASS MODO MANUAL ---
    st.info("✍️ Modo Llenado Manual Activado: Crea un certificado desde cero sin subir archivos.")
    if st.button("✨ Generar Plantilla en Blanco"):
        # Inyectamos datos vacíos en la memoria para despertar al Bloque 4
        st.session_state['ocr_data'] = {}
        df_vacio = pd.DataFrame([{
            'desc': '', 'cant': '0', 'um': 'UNID', 'peso': '0.00',
            'fecha_origen': datetime.now().strftime("%d/%m/%Y"),
            'guia_origen': '', 'placa_origen': ''
        }])
        st.session_state['df_items'] = df_vacio
        st.rerun()

# ====================================================================
# --- BLOQUE 4: UI - Validación y Edición ---
# ====================================================================
# Ahora verificamos las variables CORRECTAS
if 'ocr_data' in st.session_state and 'df_items' in st.session_state:
    grl = st.session_state['ocr_data']
    df_items = st.session_state['df_items']
    
    st.markdown("### Validación")
    
    st.markdown('''
            <style>
            div[data-baseweb="input"] > input[aria-label="Correlativo"] {
                background-color: #FFFF00 !important;
                color: black !important;
                font-weight: bold !important;
            }
            </style>
        ''', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    # --- MEJORA: CÁLCULO INTELIGENTE DEL CORRELATIVO DESDE SHEETS ---
    siguiente_corr = "001" # Valor por defecto si la hoja está vacía
    try:
        from src.services.google_service import leer_sheet_seguro
        import pandas as pd
        
        df_historial = leer_sheet_seguro("Historial")
        
        if not df_historial.empty and "Correlativo" in df_historial.columns:
            # 1. Definir qué estamos buscando según el menú lateral
            palabra_clave = "Comercialización" if "Comercialización" in tipo_flujo else "Final"
            
            # 2. Filtrar el historial donde CUALQUIER columna contenga la palabra clave
            # (Así no dependemos de saber el nombre exacto de la columna del tipo de certificado)
            mask = df_historial.astype(str).apply(lambda x: x.str.contains(palabra_clave, case=False, na=False)).any(axis=1)
            df_filtrado = df_historial[mask]
            
            if not df_filtrado.empty:
                # 3. Extraer números, ignorar textos rotos, sacar el máximo y sumar 1
                max_corr = pd.to_numeric(df_filtrado["Correlativo"], errors='coerce').max()
                if pd.notna(max_corr):
                    siguiente_corr = f"{int(max_corr) + 1:03d}"
    except Exception as e:
        st.warning(f"⚠️ Aviso: No se pudo auto-calcular el correlativo ({e}). Se usará 001.")
        
    # Inyectamos el número calculado directamente en el input amarillo
    v_corr = c1.text_input("Correlativo", value=siguiente_corr)
    
    fecha_base = grl.get('fecha', datetime.now().strftime("%d/%m/%Y"))
    
    # --- Magia Robusta: Selección automática forzada (CORREGIDO) ---
    with c2:
        if "Comercialización" in tipo_flujo:
            st.info("📅 COMERCIALIZACIÓN (FIN DE MES)")
            f_calc = obtener_fin_de_mes(fecha_base)
            tipo_op = "Comercialización"
            v_fec_emis = st.text_input("F. Emisión", value=f_calc)
        else:
            st.info("📅 DISPOSICIÓN FINAL +2")
            try:
                f_calc = (datetime.strptime(fecha_base, "%d/%m/%Y") + timedelta(days=2)).strftime("%d/%m/%Y")
            except ValueError:
                f_calc = (datetime.now() + timedelta(days=2)).strftime("%d/%m/%Y")
            tipo_op = "Disposición Final"
            v_fec_emis = st.text_input("F. Emisión", value=f_calc)
    
    # Extraemos la guía y placa LIMPIAS desde la tabla procesada
    guia_limpia = df_items['guia_origen'].iloc[0] if not df_items.empty else grl.get('serie', '')
    placa_limpia = df_items['placa_origen'].iloc[0] if not df_items.empty else grl.get('vehiculo', '')
    
    # Si estamos en manual, siguen desbloqueadas para escribir
    v_guia = c3.text_input("Guía", guia_limpia, disabled=not modo_manual)
    v_placa = c4.text_input("Placa", placa_limpia, disabled=not modo_manual)

    v_partida = st.text_input("Partida", formato_nompropio(grl.get('punto_partida', '')))
    v_llegada = st.text_input("Llegada", formato_nompropio(grl.get('punto_llegada', '')))
    v_dest = st.text_input("Destinatario", grl.get('destinatario', ''))

    v_items_df = st.data_editor(df_items, num_rows="dynamic", use_container_width=True)

    from src.services.google_service import leer_sheet_seguro
    if 'repo' not in st.session_state:
        st.session_state.repo = {
            "emisores": leer_sheet_seguro("EMPRESAS"),
            "clientes": leer_sheet_seguro("CLIENTES"),
            "servicios": leer_sheet_seguro("COMERCIALIZACION")
        }
    repo = st.session_state.repo

    c_a, c_b = st.columns(2)
    with c_a:
        st.markdown("**Emisor**")
        v_emi = empresa_firma 
        st.info(f"🏢 {v_emi}")
        
        v_ruc_e, v_reg_e = "", ""
        if not repo['emisores'].empty:
            try:
                row_e = repo['emisores'][repo['emisores']['EMPRESA'] == v_emi].iloc[0]
                v_ruc_e, v_reg_e = str(row_e['RUC']), str(row_e['REGISTRO'])
            except: pass
        st.caption(f"RUC: {v_ruc_e} | REG: {v_reg_e}")
        
        v_tit = st.selectbox("Título", repo['servicios'].iloc[:,0].unique() if not repo['servicios'].empty else [])

    with c_b:
        # --- Extracción Defensiva del Cliente (CORREGIDO) ---
        cliente_crudo = grl.get('cliente') or grl.get('razon_social') or grl.get('empresa') or ""
        ruc_crudo = grl.get('ruc_cliente') or grl.get('ruc') or ""
        
        # EL CAMBIO: Forzamos el texto a MAYÚSCULAS absolutas usando .upper()
        cliente_ocr = str(cliente_crudo).upper() if cliente_crudo else ""

        v_cli = st.text_input("Cliente (Obligatorio)", value=cliente_ocr)
        if not v_cli:
            st.warning("⚠️ No se detectó Cliente en la guía. Debes ingresarlo manualmente.")

        v_ruc_c = st.text_input("RUC Cliente", value=ruc_crudo)
        
        v_serv = st.selectbox("Servicio", repo['servicios'].iloc[:,1].unique() if not repo['servicios'].empty else [])
        v_res = st.selectbox("Residuo", repo['servicios'].iloc[:,2].unique() if not repo['servicios'].empty else [])

    dest_final = v_dest if "EPMI" not in str(v_dest).upper() else "EPMI S.A.C."

    st.divider()
# ====================================================================
# --- BLOQUE 5: UI - Generación de Word, Descarga y Registro en Sheets ---
# ====================================================================
    if st.button("GENERAR CERTIFICADO", type="primary"):
        drive, _ = obtener_servicios()
        if drive:
            try:
                id_p = PLANTILLAS[empresa_firma][tipo_flujo]
                req = drive.files().export_media(fileId=id_p, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                fh = io.BytesIO()
                from googleapiclient.http import MediaIoBaseDownload
                dl = MediaIoBaseDownload(fh, req)
                done = False
                while not done: _, done = dl.next_chunk()
                
                doc = DocxTemplate(io.BytesIO(fh.getvalue()))
                
                # Contexto para inyectar en el Word
                ctx = {
                    "CORRELATIVO": v_corr, "TITULO": v_tit, "REGISTRO": v_reg_e,
                    "EMPRESA": v_emi, "RUC_EMPRESA": v_ruc_e, "RUC": v_ruc_e, 
                    "CLIENTE": v_cli, "RUC_CLIENTE": v_ruc_c, "RAZON_SOCIAL_CLIENTE": v_cli,
                    "SERVICIO_O_COMPRA": v_serv, "TIPO_DE_RESIDUO": v_res,
                    "PUNTO_PARTIDA": v_partida, 
                    "DIRECCION_EMPRESA": v_llegada, 
                    "DIRECCION_LLEGADA": v_llegada, 
                    "LLEGADA": v_llegada,
                    "EMPRESA_2": dest_final, "FECHA_EMISION": v_fec_emis,
                    "DESTINATARIO_FINAL": v_dest
                }
                doc.render(ctx)
                buf_tpl = io.BytesIO()
                doc.save(buf_tpl)

                items_para_tabla = v_items_df.to_dict('records')
                final_bytes = inyectar_tabla_en_docx(io.BytesIO(buf_tpl.getvalue()), items_para_tabla)
                
                # --- LÓGICA DE NOMENCLATURA ESTRICTA ---
                # 1. Extraer nombre de Fundo/Planta
                partes_partida = str(v_partida).split(' - ')
                nombre_crudo = partes_partida[-1].strip() if len(partes_partida) > 1 else "Sede Principal"
                
                # 2. Aniquilar palabras reservadas (Planta, Fundo, Sede) usando Regex (Case Insensitive)
                import re
                nombre_limpio = re.sub(r'(?i)^(Planta|Fundo|Sede|Sucursal|Predio)\s+', '', nombre_crudo).strip()
                
                # 3. Determinar etiqueta
                etiqueta_tipo = "Comercializacion" if "Comercialización" in tipo_flujo else "Servicio"
                
                # 4. Ensamblar con formato NOMPROPIO (Title Case) absoluto
                from src.utils.format_utils import formato_nompropio
                cli_format = formato_nompropio(v_cli)
                planta_format = formato_nompropio(nombre_limpio)
                etiq_format = formato_nompropio(etiqueta_tipo)
                
                nombre_archivo_final = f"{cli_format} - {planta_format} - {etiq_format} - {v_corr}"
                
                # Guardar en sesión
                st.session_state.word_buffer = final_bytes
                st.session_state.nombre_generado = nombre_archivo_final
                st.session_state.generado = True
                
                st.success(f"✅ Certificado Generado: {nombre_archivo_final}")
                st.balloons()
            except Exception as e: st.error(f"Error: {e}")

    if st.session_state.get('generado'):
        nombre_safe = st.session_state.get('nombre_generado', 'Certificado')
        st.download_button("📩 Descargar Certificado", st.session_state.word_buffer, f"{nombre_safe}.docx")
        
    if st.button("💾 Registrar y Subir a Drive"):
        # Validación: Asegurar que el certificado ya se generó
        if not st.session_state.get('generado') or not st.session_state.get('word_buffer'):
            st.warning("⚠️ Primero debes darle al botón 'GENERAR CERTIFICADO'.")
        else:
            with st.spinner("Subiendo a Google Drive y registrando en Sheets... 🚀"):
                peso_t = sum([float(str(x).replace(',','.')) for x in v_items_df['peso']]) if 'peso' in v_items_df.columns else 0
                nombre_safe = st.session_state.get('nombre_generado', f"Certificado {v_corr}")
                buffer = st.session_state.word_buffer
                
                # 1. Ejecutar el enrutador hacia Drive
                link_drive = subir_a_drive(buffer, nombre_safe, tipo_flujo)
                
                # 2. Armar la fila de datos para Sheets
                link_final = link_drive if link_drive else "Error de Permisos en Drive"
                
                # --- NUEVA LÓGICA DE EXTRACCIÓN Y ORDEN PARA SHEETS ---
                partes_partida = str(v_partida).split(' - ')
                nombre_crudo = partes_partida[-1].strip() if len(partes_partida) > 1 else "Sede Principal"
                import re; nombre_limpio = re.sub(r'(?i)^(Planta|Fundo|Sede|Sucursal|Predio)\s+', '', nombre_crudo).strip()
                from src.utils.format_utils import formato_nompropio
                
                val_empresa = formato_nompropio(v_cli)
                val_fundo = formato_nompropio(nombre_limpio)
                val_cert = "Comercialización" if "Comercialización" in tipo_flujo else "Servicios"
                val_guia = str(v_guia).strip()
                
                # Columnas Excel: 0(Fecha), 1(Empresa), 2(Fundo), 3(Correlativo), 4(Certificado), 5(Guia), 6(L Guia), 7(L Doc), 8(L pdf), 9(Obs)
                datos_log = [v_fec_emis, val_empresa, val_fundo, v_corr, val_cert, val_guia, "", link_final, "", ""]

                
                # 3. Disparar el guardado en Google Sheets
                if registrar_en_control(datos_log):
                    if link_drive:
                        st.success("✅ ¡Operación Exitosa! Documento en Drive y base de datos actualizada.")
                        st.markdown(f"[🔗 Clic aquí para ver el documento en Drive]({link_drive})")
                    else:
                        st.warning("⚠️ El registro se guardó en el Excel, pero Drive rechazó el archivo (Falta compartir la carpeta con el correo del Service Account).")
                else:
                    st.error("❌ Falló la conexión con Sheets.")