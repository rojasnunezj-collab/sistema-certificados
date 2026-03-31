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

# --- MEJORA: Añadimos leer_sheet_seguro a la lista de importaciones ---
from src.services.google_service import (
    obtener_servicios, subir_a_drive, obtener_plantilla_drive, 
    subir_modelo_a_drive, obtener_mapa_plantillas_drive, 
    obtener_datos_empresas_desde_sheets, registrar_en_control,
    leer_sheet_seguro  # <--- Agregado aquí
)

from src.config.settings import PLANTILLAS, CARPETAS_DESTINO # <-- Añade esto
from src.utils.document_utils import inyectar_tabla_en_docx
from src.utils.format_utils import (
    limpiar_monto, formato_inteligente, normalizar_fecha, 
    limpiar_descripcion, formatear_guia, obtener_fin_de_mes,
    formato_nompropio
)
from docxtpl import DocxTemplate
from googleapiclient.http import MediaIoBaseDownload

# --- CARGA DE BASES DE DATOS (REPOS) ---
# Esto garantiza que 'repo' exista siempre en toda la App
if 'repo' not in st.session_state:
    with st.spinner("Conectando con bases de datos de Google..."):
        st.session_state.repo = {
            "emisores": leer_sheet_seguro("EMPRESAS"),
            "clientes": leer_sheet_seguro("CLIENTES"),
            "servicios": leer_sheet_seguro("SERVICIOS")
        }

# Definimos la variable 'repo' global para el resto del código
repo = st.session_state.repo

# --- Configuración Inicial
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

# --- Barra Lateral (Líneas 40 a 100 aprox) ---
with st.sidebar:
    st.header("Configuración de Flujo")
    
    # 1. Controles principales (4 espacios de indentación)
    es_modelo = st.checkbox("📝 Generar como Certificado Modelo", value=False)
    modo_manual = st.toggle("🔴 Llenado Manual (Sin PDF)", value=False)
    
    if es_modelo:
        st.info("💡 MODO MODELO ACTIVO: Se usarán las plantillas de prueba.")
    
    st.divider()
    
    # 2. Carga del Menú Dinámico
    from src.services.google_service import obtener_mapa_plantillas_drive
    mapa_plantillas = obtener_mapa_plantillas_drive(es_modelo=es_modelo)
    
    # 3. Selectores de Empresa y Servicio
    lista_empresas = list(mapa_plantillas.keys())
    
    if not lista_empresas:
        lista_empresas = ["Esperando conexión con Drive..."]
        
    empresa_firma = st.selectbox("Empresa Firmante", options=lista_empresas)
    
    servicios_base = mapa_plantillas.get(empresa_firma, [])
    opciones_finales = []
    
    # ==========================================
    # 🚧 LECTURA DINÁMICA DESDE DRIVE 🚧
    # ==========================================
    if es_modelo:
        # 1. SI ES MODELO: Pasa directamente lo que Drive encuentre (Escalable 100%)
        opciones_finales = servicios_base
        if not opciones_finales:
            opciones_finales = ["No se detectaron plantillas modelo en Drive"]
    else:
        # 2. SI ES NORMAL: Mantiene tu lógica original
        for serv in servicios_base:
            if serv == "Disposición Final":
                opciones_finales.extend(["Disposición Final 1", "Disposición Final 2"])
            else:
                opciones_finales.append(serv)
                
        if not opciones_finales:
            opciones_finales = ["Comercialización", "Disposición Final 1", "Disposición Final 2"]
            
    tipo_flujo = st.selectbox("Tipo de Certificado", options=opciones_finales)

    st.divider() # <--- Aquí ya no habrá error
    
    if st.sidebar.button("Limpiar Sesión", use_container_width=True):
        # 1. Borramos SOLO las variables del certificado actual (dejando vivo a 'repo')
        llaves_basura = ['ocr_data', 'df_items', 'word_buffer', 'nombre_generado', 'generado']
        for k in llaves_basura:
            if k in st.session_state:
                del st.session_state[k]
        
        # 2. Truco de magia: Le cambiamos el DNI al file_uploader para que nazca vacío
        if 'uploader_key' not in st.session_state:
            st.session_state.uploader_key = 1
        else:
            st.session_state.uploader_key += 1
            
        # 3. Recargamos la pantalla suavemente
        st.rerun()


# ====================================================================
# --- BLOQUE 3: UI - Ingesta y Procesamiento de Archivos ---
# ====================================================================
if not modo_manual:
    # Verifica que tu línea sea así (usa uploader_key):
    archivos = st.file_uploader("Sube tus guías", type=["pdf"], accept_multiple_files=True, key=f"uploader_{st.session_state.get('uploader_key', 0)}")

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
    if es_modelo:
        st.success("🛠️ MODO MODELO ACTIVADO: Estás creando un certificado de prueba. Se usará la plantilla de Google Drive y no afectará tus correlativos.")
        texto_boton = "✨ Generar Plantilla para Modelo"
    else:
        st.info("✍️ Modo Llenado Manual Activado: Crea un certificado oficial desde cero sin subir archivos.")
        texto_boton = "✨ Generar Plantilla en Blanco"

    if st.button(texto_boton):
        # Inyectamos datos vacíos en la memoria para despertar al Bloque 4
        st.session_state['ocr_data'] = {}
        df_vacio = pd.DataFrame([{
            'desc': '', 'cant': '0', 'um': 'UNID', 'peso': '0.00',
            'fecha_origen': '',  # <--- ¡AQUÍ ESTÁ LA MAGIA! Ahora nace vacío.
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
            # BIFURCACIÓN PARA BÚSQUEDA INDEPENDIENTE
            if es_modelo:
                if "Comercialización" in tipo_flujo:
                    palabra_clave = "M-COM"
                else:
                    palabra_clave = "M-FIN"
            elif "Comercialización" in tipo_flujo:
                palabra_clave = "Comercialización"
            else:
                palabra_clave = "Final"
            
            # 2. Filtrar el historial donde CUALQUIER columna contenga la palabra clave
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

    # 2. COLUMNAS: Partimos la pantalla en 2 mitades
    c_a, c_b = st.columns(2, gap="large")

    # ==========================================
    # ⬅️ LADO IZQUIERDO: EMISOR 
    # ==========================================
    with c_a:
        st.subheader("Emisor")
        
        if es_modelo:
            # --- SI ES MODELO: Cajas en blanco, editables y sin buscar en Excel ---
            st.caption("💡 Datos de la empresa que emite (Modo Modelo):")
            emisor_nombre = st.text_input("Nombre Emisor", value="", key="em_nom_mod")
            emisor_ruc = st.text_input("RUC Emisor", value="", key="em_ruc_mod")
            emisor_reg = st.text_input("Registro Emisor", value="", key="em_reg_mod")
        else:
            # --- SI ES NORMAL: Lógica de búsqueda dinámica en Excel ---
            info_emisor = None
            try:
                from src.services.google_service import leer_sheet_seguro
                import pandas as pd
                
                df_empresas = leer_sheet_seguro("EMPRESAS")
                
                if not df_empresas.empty:
                    nombres_excel = df_empresas.iloc[:, 0].astype(str).str.strip().str.upper()
                    empresa_target = empresa_firma.strip().upper()
                    
                    fila = df_empresas[nombres_excel == empresa_target]
                    
                    if fila.empty:
                        empresa_corta = empresa_target.replace(" S.A.C.", "").replace(" SAC", "").replace(".", "").strip()
                        fila = df_empresas[nombres_excel.str.contains(empresa_corta, na=False)]
                    
                    if not fila.empty:
                        info_emisor = {
                            'ruc': str(fila.iloc[0, 1]).strip(),
                            'reg': str(fila.iloc[0, 2]).strip()
                        }
            except Exception as e:
                st.error(f"Error al conectar con la base de datos de empresas: {e}")        

            st.caption("💡 Datos del Emisor")

            if info_emisor:
                emisor_nombre = st.text_input("Nombre Emisor", value=empresa_firma, disabled=True, key="em_nom_lock")
                emisor_ruc = st.text_input("RUC Emisor", value=info_emisor['ruc'], disabled=True, key="em_ruc_lock")
                emisor_reg = st.text_input("Registro Emisor", value=info_emisor['reg'], disabled=True, key="em_reg_lock")
                st.success("✅ Datos verificados desde la base de datos.")
            else:
                st.error("❌ ERROR DE SINCRONIZACIÓN")
                st.warning(f"La empresa '{empresa_firma}' no existe en la pestaña 'empresas' del Excel.")
                emisor_nombre = empresa_firma
                emisor_ruc = ""
                emisor_reg = ""

        # TÍTULO: Va exactamente debajo del registro del Emisor
        v_tit = st.selectbox(
            "Título", 
            repo['servicios'].iloc[:,0].unique() if not repo['servicios'].empty else ["CERTIFICADO DE MANEJO"]
        )
            
        # --- CONEXIÓN DE VARIABLES PARA EL WORD ---
        v_emi = emisor_nombre
        v_emp_e = emisor_nombre
        v_ruc_e = emisor_ruc
        v_reg_e = emisor_reg
        

    # ==========================================
    # ➡️ LADO DERECHO: CLIENTE Y SERVICIOS 
    # ==========================================
    with c_b:

        st.subheader("Cliente")
        st.caption("💡 Datos de cliente")
        
        cliente_crudo = grl.get('cliente') or grl.get('razon_social') or grl.get('empresa') or ""
        ruc_crudo = grl.get('ruc_cliente') or grl.get('ruc') or ""
        cliente_ocr = str(cliente_crudo).upper() if cliente_crudo else ""

        if es_modelo:
            v_cli = st.text_input("Cliente (Modelo)", value="", key="cl_nom_mod")
            v_ruc_c = st.text_input("RUC Cliente (Modelo)", value="", key="cl_ruc_mod")
        elif modo_manual:
            from src.services.google_service import obtener_clientes_desde_sheets
            diccionario_clientes = obtener_clientes_desde_sheets()
            opciones_clientes = [""] + list(diccionario_clientes.keys())
            v_cli = st.selectbox("Cliente (Desde Base de Datos)", options=opciones_clientes)
            
            ruc_encontrado = diccionario_clientes.get(v_cli, "") if v_cli else ""
            v_ruc_c = st.text_input("RUC Cliente", value=ruc_encontrado)
        else:
            v_cli = st.text_input("Cliente (Extraído)", value=cliente_ocr)
            v_ruc_c = st.text_input("RUC Cliente", value=ruc_crudo)
        
        v_serv = st.selectbox(
            "Servicio", 
            repo['servicios'].iloc[:,1].unique() if not repo['servicios'].empty else ["Sin Datos"]
        )
        v_res = st.selectbox(
            "Residuo", 
            repo['servicios'].iloc[:,2].unique() if not repo['servicios'].empty else ["Sin Datos"]
        )
        

# ====================================================================
# --- BLOQUE 5: UI - Generación de Word, Descarga y Registro en Sheets ---
# ====================================================================
st.divider()

# --- 1. PROCESO DE GENERACIÓN (BOTÓN PRIMARIO - Oculto secuencialmente) ---
# -- NUEVA LÓGICA STRICTA: Comprobar que todos los datos están llenos --
# Revisamos las variables críticas que definimos en el Bloque 4:
# v_cli (Cliente), v_ruc_c (RUC Cliente) no deben estar vacíos.
# v_items_df (Tabla) no debe estar vacía (tener al menos una línea manual en image 3).

# Simplificamos validación a campos clave vacíos en imagen 3
# Nota: "Llenado Manual (Sin PDF)" en imagen 3 implica llenar datos a mano y tabla.
# --- Lógica blindada para evitar NameError cuando la app recién abre ---
v_cli_seguro = locals().get('v_cli', '')
v_ruc_seguro = locals().get('v_ruc_c', '')
v_df_seguro = locals().get('v_items_df', None)

if str(v_cli_seguro).strip() != "" and str(v_ruc_seguro).strip() != "" and v_df_seguro is not None and not v_df_seguro.empty:

    # EL BOTÓN SOLO APARECE AQUÍ, SI formulario_completo es VERDADERO
    if st.button("GENERAR CERTIFICADO", type="primary"):
        drive, _ = obtener_servicios()
        if drive:
            try:
                if es_modelo:
                    # RUTA MODELOS: Buscar en la nueva carpeta de Drive
                    from src.services.google_service import obtener_plantilla_drive
                    fh = obtener_plantilla_drive(empresa_firma, tipo_flujo, drive)
                    doc = DocxTemplate(fh)
                else:
                    # RUTA OFICIAL: Usar la plantilla configurada en PLANTILLAS
                    id_p = PLANTILLAS[empresa_firma][tipo_flujo]
                    req = drive.files().export_media(fileId=id_p, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                    fh = io.BytesIO()
                    from googleapiclient.http import MediaIoBaseDownload
                    dl = MediaIoBaseDownload(fh, req)
                    done = False
                    while not done: _, done = dl.next_chunk()
                    doc = DocxTemplate(io.BytesIO(fh.getvalue()))
                    
                # Contexto BLINDADO para inyectar en el Word
                ctx = {
                    # --- VARIABLES COMUNES ---
                    "CORRELATIVO": v_corr, 
                    "TITULO": v_tit, 
                    "REGISTRO": emisor_reg,
                    "CLIENTE": v_cli, 
                    "RUC_CLIENTE": v_ruc_c, 
                    "RAZON_SOCIAL_CLIENTE": v_cli,
                    "SERVICIO_O_COMPRA": v_serv, 
                    "TIPO_DE_RESIDUO": v_res,
                    "PUNTO_PARTIDA": v_partida, 
                    "DIRECCION_EMPRESA": v_llegada, 
                    "DIRECCION_LLEGADA": v_llegada, 
                    "LLEGADA": v_llegada,
                    "EMPRESA_2": emisor_nombre,
                    "FECHA_EMISION": v_fec_emis,
                    "DESTINATARIO_FINAL": emisor_nombre,
                    
                    # --- VARIABLES PARA PLANTILLAS NORMALES ---
                    "EMPRESA": emisor_nombre, 
                    "RUC_EMPRESA": emisor_ruc, 
                    "RUC": emisor_ruc,
                    
                    # --- VARIABLES PARA PLANTILLAS MODELO ---
                    "EMISOR": emisor_nombre,        
                    "RUC_EMISOR": emisor_ruc        
                }
                doc.render(ctx)
                buf_tpl = io.BytesIO()
                doc.save(buf_tpl)

                items_para_tabla = v_items_df.to_dict('records')
                final_bytes = inyectar_tabla_en_docx(io.BytesIO(buf_tpl.getvalue()), items_para_tabla)
                
                # --- LÓGICA DE NOMENCLATURA ESTRICTA ---
                if es_modelo:
                    # SI ES MODELO: Usamos el formato corto y en mayúsculas
                    cliente_limpio = str(v_cli).strip().upper()
                    tipo_corto = "M-COM" if "Comercialización" in tipo_flujo else "M-FIN"
                    nombre_archivo_final = f"{cliente_limpio} - {tipo_corto} - {v_corr}"
                else:
                    # SI ES NORMAL: Usamos la ruta larga de siempre
                    partes_partida = str(v_partida).split(' - ')
                    nombre_crudo = partes_partida[-1].strip() if len(partes_partida) > 1 else "Sede Principal"
                    
                    import re
                    nombre_limpio = re.sub(r'(?i)^(Planta|Fundo|Sede|Sucursal|Predio)\s+', '', nombre_crudo).strip()
                    
                    etiqueta_tipo = "Comercializacion" if "Comercialización" in tipo_flujo else "Servicio"
                    
                    from src.utils.format_utils import formato_nompropio
                    cli_format = formato_nompropio(v_cli)
                    planta_format = formato_nompropio(nombre_limpio)
                    etiq_format = formato_nompropio(etiqueta_tipo)
                    
                    nombre_archivo_final = f"{cli_format} - {planta_format} - {etiq_format} - {v_corr}"
                
                # GUARDAR EN SESIÓN PARA PERSISTENCIA
                st.session_state.word_buffer = final_bytes
                st.session_state.nombre_generado = nombre_archivo_final
                st.session_state.generado = True
                
                st.success(f"✅ Certificado Generado: {nombre_archivo_final}")
                st.balloons()
                st.rerun()

            except Exception as e: 
                st.error(f"Error: {e}")

# --- 2. MOSTRAR DESCARGA Y REGISTRO (SOLO SI YA SE GENERÓ) ---
if st.session_state.get('generado'):
    
    nombre_safe = st.session_state.get('nombre_generado', 'Certificado')
    
    st.download_button(
        label="📩 Descargar Certificado", 
        data=st.session_state.word_buffer, 
        file_name=f"{nombre_safe}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    if st.button("💾 Registrar y Subir a Drive"):
        with st.spinner("Subiendo a Google Drive y registrando en Sheets... 🚀"):
            # Calculamos peso total para el log
            peso_t = sum([float(str(x).replace(',','.')) for x in v_items_df['peso'] if str(x).strip() not in ['None', '', 'nan']]) if 'peso' in v_items_df.columns else 0
            
            buffer = st.session_state.word_buffer
            nombre_safe = st.session_state.nombre_generado
            
            servicio_drive, _ = obtener_servicios()
            
          # 1. Ejecutar el enrutador hacia Drive
if es_modelo:
    from src.services.google_service import subir_modelo_a_drive
    link_drive = subir_modelo_a_drive(f"{nombre_safe}.docx", buffer, servicio_drive)
else:
    # --- AHORA SÍ: Usamos tus variables reales ---
    carpeta_exacta = CARPETAS_DESTINO[empresa_firma][tipo_flujo] 
    
    # Le enviamos esa carpeta exacta a la función
    link_drive = subir_a_drive(buffer, nombre_safe, tipo_flujo, carpeta_id=carpeta_exacta)
            
            # 2. Armar la fila de datos para Sheets
    link_final = link_drive if link_drive else "Error de Permisos en Drive"
            
            # --- LÓGICA DE EXTRACCIÓN PARA SHEETS (REPLICADA) ---
    partes_partida = str(v_partida).split(' - ')
    nombre_crudo = partes_partida[-1].strip() if len(partes_partida) > 1 else "Sede Principal"
    import re; nombre_limpio = re.sub(r'(?i)^(Planta|Fundo|Sede|Sucursal|Predio)\s+', '', nombre_crudo).strip()
            
    from datetime import datetime
    val_empresa = str(v_cli).strip().upper()
    val_fundo = str(nombre_limpio).strip().upper()
            
    if es_modelo:
        val_cert = "M-COM" if "Comercialización" in tipo_flujo else "M-FIN"
    elif "Comercialización" in tipo_flujo:
        val_cert = "COMERCIALIZACIÓN"
    elif "Final" in tipo_flujo or "Disposición" in tipo_flujo:
        val_cert = "FINAL"
    else:
        val_cert = "SERVICIOS"
            
            # --- 1. Lógica para capturar MÚLTIPLES guías ---
    if not v_items_df.empty and 'guia_origen' in v_items_df.columns:
        guias_lista = [str(g).strip().upper() for g in v_items_df['guia_origen'].unique() if str(g).strip() not in ['None', '', 'nan']]
        val_guia_completa = ", ".join(guias_lista)
    else:
        val_guia_completa = str(v_guia).strip().upper()

            # --- 2. Fecha y armado de datos para Sheets ---
    fecha_registro = datetime.now().strftime("%d/%m/%Y")
    datos_log = [fecha_registro, val_empresa, val_fundo, v_corr, val_cert, val_guia_completa, "", link_final, "", ""]
            
    if registrar_en_control(datos_log):
        if link_drive:
            st.success("✅ ¡Operación Exitosa! Documento en Drive y base de datos actualizada.")
            st.markdown(f"[🔗 Clic aquí para ver el documento en Drive]({link_drive})")
            st.cache_data.clear() 
        else:
            st.warning("⚠️ El registro se guardó en el Excel, pero Drive rechazó el archivo.")
    else:
        st.error("❌ Falló la conexión con Sheets.")