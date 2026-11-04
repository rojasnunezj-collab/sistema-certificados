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

# --- MEJORA: Autorecarga y protección contra desfase de módulos en caliente en Streamlit Cloud ---
import sys
import importlib

# Forzar recarga de submódulos locales si ya están en memoria para garantizar código fresco tras git push
for _m in ['src.config.settings', 'src.utils.document_utils', 'src.utils.format_utils', 'src.services.google_service']:
    if _m in sys.modules:
        try:
            importlib.reload(sys.modules[_m])
        except Exception:
            pass

try:
    import src.services.google_service as _gs_mod
    obtener_servicios = getattr(_gs_mod, 'obtener_servicios', None)
    subir_a_drive = getattr(_gs_mod, 'subir_a_drive', None)
    obtener_plantilla_drive = getattr(_gs_mod, 'obtener_plantilla_drive', None)
    subir_modelo_a_drive = getattr(_gs_mod, 'subir_modelo_a_drive', None)
    obtener_mapa_plantillas_drive = getattr(_gs_mod, 'obtener_mapa_plantillas_drive', None)
    obtener_datos_empresas_desde_sheets = getattr(_gs_mod, 'obtener_datos_empresas_desde_sheets', None)
    registrar_en_control = getattr(_gs_mod, 'registrar_en_control', None)
    leer_sheet_seguro = getattr(_gs_mod, 'leer_sheet_seguro', None)
    obtener_catalogo_guias = getattr(_gs_mod, 'obtener_catalogo_guias', None)
    buscar_guias_repositorio = getattr(_gs_mod, 'buscar_guias_repositorio', None)
    descargar_guias_drive = getattr(_gs_mod, 'descargar_guias_drive', None)
    actualizar_bitacora_guias = getattr(_gs_mod, 'actualizar_bitacora_guias', None)
    buscar_actualizar_guia = getattr(_gs_mod, 'buscar_actualizar_guia', None)
    buscar_guias_asociadas_para_unir = getattr(_gs_mod, 'buscar_guias_asociadas_para_unir', None)
    descargar_archivo_drive_por_id_o_nombre = getattr(_gs_mod, 'descargar_archivo_drive_por_id_o_nombre', None)
    unir_tres_documentos_pdf = getattr(_gs_mod, 'unir_tres_documentos_pdf', None)
    subir_pdf_a_drive = getattr(_gs_mod, 'subir_pdf_a_drive', None)
    actualizar_link_pdf_historial = getattr(_gs_mod, 'actualizar_link_pdf_historial', None)
    extraer_id_drive = getattr(_gs_mod, 'extraer_id_drive', None)
    obtener_catalogo_servicios_por_categoria = getattr(_gs_mod, 'obtener_catalogo_servicios_por_categoria', None)
    convertir_docx_a_pdf = getattr(_gs_mod, 'convertir_docx_a_pdf', None)
    buscar_datos_certificado_en_historial = getattr(_gs_mod, 'buscar_datos_certificado_en_historial', None)
    sobrescribir_o_subir_pdf_drive = getattr(_gs_mod, 'sobrescribir_o_subir_pdf_drive', None)
    registrar_edicion_en_historial = getattr(_gs_mod, 'registrar_edicion_en_historial', None)
    obtener_info_revision_documento_drive = getattr(_gs_mod, 'obtener_info_revision_documento_drive', None)
    obtener_configuracion_empresas = getattr(_gs_mod, 'obtener_configuracion_empresas', None)
    buscar_config_empresa = getattr(_gs_mod, 'buscar_config_empresa', None)
    obtener_sheet_sigersol_destino = getattr(_gs_mod, 'obtener_sheet_sigersol_destino', lambda em=None, cl=None: "10fES_h_4c8AdwaAAa46nuhTwbhNoifhCMzagCPI_4HI")
    obtener_carpeta_guias = getattr(_gs_mod, 'obtener_carpeta_guias', lambda emp, tg='hechas': "")
    guardar_guia_en_drive = getattr(_gs_mod, 'guardar_guia_en_drive', lambda *a, **k: {'ok': False, 'error': 'Función de respaldo no disponible'})
    guardar_pdf_en_carpeta_docs = getattr(_gs_mod, 'guardar_pdf_en_carpeta_docs', None)
    obtener_carpeta_padre_archivo_drive = getattr(_gs_mod, 'obtener_carpeta_padre_archivo_drive', None)
except Exception as err_import:
    import traceback
    st.error(f"❌ Error al inicializar servicios de Google: {err_import}")
    st.code(traceback.format_exc())
    st.stop()

try:
    import src.config.settings as _settings_mod
    PLANTILLAS = getattr(_settings_mod, 'PLANTILLAS', {})
    CARPETAS_DESTINO = getattr(_settings_mod, 'CARPETAS_DESTINO', {})
except Exception:
    PLANTILLAS, CARPETAS_DESTINO = {}, {}

try:
    import src.utils.document_utils as _doc_utils_mod
    inyectar_tabla_en_docx = _doc_utils_mod.inyectar_tabla_en_docx
    sustituir_certificado_en_pdf = _doc_utils_mod.sustituir_certificado_en_pdf
    sanitizar_tags_docx = getattr(_doc_utils_mod, 'sanitizar_tags_docx', None)
except Exception:
    from src.utils.document_utils import inyectar_tabla_en_docx, sustituir_certificado_en_pdf
    sanitizar_tags_docx = None

if sanitizar_tags_docx is None:
    def sanitizar_tags_docx(docx_input):
        import zipfile
        import re
        if isinstance(docx_input, (bytes, bytearray)):
            in_buf = io.BytesIO(docx_input)
        elif hasattr(docx_input, 'getvalue'):
            in_buf = io.BytesIO(docx_input.getvalue())
        elif hasattr(docx_input, 'read'):
            try: docx_input.seek(0)
            except Exception: pass
            in_buf = io.BytesIO(docx_input.read())
        else:
            in_buf = io.BytesIO(docx_input)
        tag_map = {
            'SERVICIO O COMPRA': 'SERVICIO_O_COMPRA',
            'RAZON SOCIAL CLIENTE': 'RAZON_SOCIAL_CLIENTE',
            'RUC CLIENTE': 'RUC_CLIENTE',
            'PUNTO PARTIDA': 'PUNTO_PARTIDA',
            'TIPO DE RESIDUO': 'TIPO_DE_RESIDUO',
            'F. EMISION': 'FECHA_EMISION',
            'F.EMISION': 'FECHA_EMISION',
            'FECHA EMISION': 'FECHA_EMISION',
            'DIRECCION EMPRESA': 'DIRECCION_EMPRESA',
            'DIRECCION LLEGADA': 'DIRECCION_LLEGADA',
            'EMPRESA 2': 'EMPRESA_2',
            'REGISTRO 2': 'REGISTRO_2',
            'RAZON SOCIAL': 'RAZON_SOCIAL',
            'RUC GENERADOR': 'RUC_GENERADOR'
        }
        out_buf = io.BytesIO()
        try:
            in_buf.seek(0)
            with zipfile.ZipFile(in_buf, 'r') as zin, zipfile.ZipFile(out_buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)
                    if item.filename.startswith('word/') and item.filename.endswith('.xml'):
                        xml_text = content.decode('utf-8', errors='ignore')
                        def replace_tag(match):
                            full_match = match.group(0)
                            tag_body = match.group(1).strip()
                            if tag_body in tag_map:
                                return f"{{{{{tag_map[tag_body]}}}}}"
                            if any(s in tag_body for s in [' ', '\t']) and not any(kw in tag_body.lower() for kw in [' if ', ' for ', ' in ', '|', '==', '!=', 'elif', 'else']):
                                cleaned = re.sub(r'[\.\s]+', '_', tag_body).strip('_')
                                return f"{{{{{cleaned}}}}}"
                            return full_match
                        xml_clean = re.sub(r'\{\{([^}]+)\}\}', replace_tag, xml_text)
                        content = xml_clean.encode('utf-8')
                    zout.writestr(item, content)
            out_buf.seek(0)
            return out_buf
        except Exception:
            in_buf.seek(0)
            return in_buf
from src.utils.format_utils import (
    limpiar_monto, formato_inteligente, normalizar_fecha, 
    limpiar_descripcion, formatear_guia, limpiar_guia_o_factura, obtener_fin_de_mes,
    formato_nompropio, es_formato_guia, es_formato_placa,
    verificar_consistencia_guia_placa, validar_inconsistencias_df
)
from docxtpl import DocxTemplate
import urllib.parse
import requests
import secrets

DEFAULT_COLUMNAS_CONFIG = [
    {"id": "fecha_origen", "nombre": "Fecha", "width_ratio": 1.0},
    {"id": "placa_origen", "nombre": "Placa", "width_ratio": 1.0},
    {"id": "guia_origen", "nombre": "N° Guía", "width_ratio": 1.1},
    {"id": "desc", "nombre": "Descripción", "width_ratio": 3.2},
    {"id": "cant", "nombre": "Cantidad", "width_ratio": 0.9},
    {"id": "um", "nombre": "Medida", "width_ratio": 0.9},
    {"id": "peso", "nombre": "Peso (Kg)", "width_ratio": 1.0}
]

# --- Configuración Inicial Streamlit (Debe ser el primer comando de Streamlit ejecutado)
st.set_page_config(page_title="Certificador AI V2", layout="wide")

def mostrar_login_google():
    """Genera la URL OAuth puramente REST con validación de estado anti-CSRF."""
    client_id = str(st.secrets["gcp_oauth"]["client_id"]).strip()
    redirect_uri = str(st.secrets["gcp_oauth"]["redirect_uri"]).strip()
    
    # Generar token anti-CSRF persistido en sesión
    if 'oauth_state' not in st.session_state:
        st.session_state['oauth_state'] = secrets.token_urlsafe(32)
    
    # Construimos el auth URL limpio con parámetro state
    params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": "openid https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/userinfo.profile",
        "prompt": "consent",
        "access_type": "offline",
        "state": st.session_state['oauth_state']
    }
    
    auth_url = "https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode(params)
    
    c_btn1, c_btn2, c_btn3 = st.columns([1,2,1])
    with c_btn2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.link_button(
            "Identificarse de forma segura con Google", 
            auth_url, 
            type="primary", 
            use_container_width=True
        )
            
    st.stop()

def verificar_retorno_oauth():
    """Atrapa el callback ?code= e invoca un POST explícito, validando token de estado anti-CSRF."""
    if 'code' in st.query_params:
        try:
            # 1. Validación de seguridad Anti-CSRF
            returned_state = st.query_params.get('state')
            if isinstance(returned_state, list): returned_state = returned_state[0]
            expected_state = st.session_state.get('oauth_state')
            
            if expected_state and returned_state != expected_state:
                st.error("⛔ Violación de seguridad: Parámetro 'state' de OAuth no coincide (posible ataque CSRF).")
                st.query_params.clear()
                st.stop()

            code = st.query_params['code']
            if isinstance(code, list): code = code[0]
                
            # Intercambio Crudo REST hacia Google Cloud (Blindado)
            token_url = "https://oauth2.googleapis.com/token"
            data = {
                "code": code,
                "client_id": str(st.secrets["gcp_oauth"]["client_id"]).strip(),
                "client_secret": str(st.secrets["gcp_oauth"]["client_secret"]).strip(),
                "redirect_uri": str(st.secrets["gcp_oauth"]["redirect_uri"]).strip(),
                "grant_type": "authorization_code"
            }
            res = requests.post(token_url, data=data, timeout=15).json()
            
            if "error" in res:
                st.error(f"Rechazo en Fase Token: {res.get('error_description', res['error'])}")
                st.stop()
                
            access_token = res["access_token"]
            
            # Recuperar perfil de Usuario
            usr_resp = requests.get(
                "https://www.googleapis.com/oauth2/v1/userinfo",
                headers={"Authorization": f"Bearer {access_token}"},
                timeout=10
            ).json()
            
            st.session_state['usuario_email'] = usr_resp.get('email', '').strip().lower()
            
            # Limpiar estado de seguridad y parámetros de query para evitar reutilizaciones
            st.session_state.pop('oauth_state', None)
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"Falla crítica procesando autenticación REST: {e}")
            st.stop()

if 'metricas_exitosos' not in st.session_state:
    st.session_state['metricas_exitosos'] = 0
if 'metricas_errores' not in st.session_state:
    st.session_state['metricas_errores'] = 0

# ====================================================================
# --- LOGICA DE GATEKEEPER (RBAC) ---
# ====================================================================
# 1. Evaluar si la URL trae un token de acceso pendiendo de intercambio
verificar_retorno_oauth()

# 2. Bloqueo 1: Denegar si no existe un token de sesión
if 'usuario_email' not in st.session_state:
    st.warning("🔒 Acceso Restringido: Requiere autenticación de empleado.")
    mostrar_login_google()

# 3. Empalmar email en memoria contra la base estocástica de Sheets
correo_actual = st.session_state['usuario_email']
from src.services.google_service import obtener_usuarios_roles
bd_usuarios = obtener_usuarios_roles()

if bd_usuarios is None:
    st.warning("⚠️ Hay intermitencias con el servidor de Google (Error 503). Por favor, intenta recargar la página en unos segundos.")
    st.stop()

# 4. Bloqueo 2: Denegar si no hace match o si su cuenta fue deshabilitada.
if correo_actual not in bd_usuarios or bd_usuarios[correo_actual].get('Estado', '').strip().upper() != 'ACTIVO':
    st.error(f"⛔ Acceso denegado: El usuario '{correo_actual}' no cuenta con un rol Asignado o está inactivo.")
    st.stop()

# 5. Adquisición Exitosa: Inyectar el Rol en el hilo continuo
st.session_state['usuario_rol'] = bd_usuarios[correo_actual].get('Rol', 'User')
st.session_state['usuario_nombre'] = bd_usuarios[correo_actual].get('Nombre', correo_actual)

with st.sidebar:
    st.info(f"👤 Conectado como: **{st.session_state['usuario_nombre']}**")
    st.divider()

# ====================================================================
# --- CARGA DE BASES DE DATOS (REPOS) PROTEGIDA TRAS AUTENTICACIÓN ---
# ====================================================================
# Se ejecuta ÚNICAMENTE después de validar que el usuario está autenticado y activo
if 'repo' not in st.session_state:
    with st.spinner("Conectando con bases de datos de Google..."):
        st.session_state.repo = {
            "emisores": leer_sheet_seguro("EMPRESAS"),
            "clientes": leer_sheet_seguro("CLIENTES"),
            "servicios": leer_sheet_seguro("SERVICIOS"),
            "direcciones": leer_sheet_seguro("Direcciones")
        }
elif "direcciones" not in st.session_state.repo:
    st.session_state.repo["direcciones"] = leer_sheet_seguro("Direcciones")

# Definimos la variable 'repo' global para el resto del código
repo = st.session_state.repo

if 'datos_extraidos' not in st.session_state:
    st.session_state.datos_extraidos = None

if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

modulo_actual = st.radio("Módulo", ["📄 Generador de Certificados", "🔍 Buscar y Actualizar Certificados", "📑 Expedientes Petramás", "🏢 Sigersol"], horizontal=True, label_visibility="collapsed")

if modulo_actual == "📄 Generador de Certificados":
    # ====================================================================
    # --- BLOQUE 2: UI - Setup Inicial y Sidebar ---
    # ====================================================================
    # --- UI: Encabezado ---
    st.title("📄 Generador de Certificados")
    st.info("Plataforma optimizada con Vertex AI Gemini")
    
    # --- Barra Lateral (Líneas 40 a 100 aprox) ---
if modulo_actual == "📄 Generador de Certificados":
    with st.sidebar:
        st.header("Configuración de Flujo")
        
        # --- Lógica de Exclusión Mutua para Modalidades ---
        if 'toggle_modelo' not in st.session_state:
            st.session_state['toggle_modelo'] = False
        if 'toggle_repo' not in st.session_state:
            st.session_state['toggle_repo'] = False
        if 'toggle_manual' not in st.session_state:
            st.session_state['toggle_manual'] = False

        def _limpiar_temporales_flujo():
            for k in ['guias_repo', 'repo_tipo_flujo', 'repo_tipo_detectado', 'procesar_ya',
                      'subido_drive_link', 'subido_drive_id', 'subido_drive_nombre',
                      'subido_fila_historial', 'subido_carpeta_exacta', 'subido_tipo_flujo',
                      'subido_tipo_cod', 'subido_v_corr', 'subido_destino_final', 'subido_guias_lista',
                      'pdf_unido_buffer', 'pdf_unido_link', 'pdf_unido_nombre']:
                if k in st.session_state:
                    del st.session_state[k]

        def _on_change_modelo():
            if st.session_state.get('toggle_modelo'):
                st.session_state['toggle_repo'] = False
                st.session_state['toggle_manual'] = False
            _limpiar_temporales_flujo()

        def _on_change_repo():
            if st.session_state.get('toggle_repo'):
                st.session_state['toggle_modelo'] = False
                st.session_state['toggle_manual'] = False
            _limpiar_temporales_flujo()

        def _on_change_manual():
            if st.session_state.get('toggle_manual'):
                st.session_state['toggle_modelo'] = False
                st.session_state['toggle_repo'] = False
            _limpiar_temporales_flujo()

        # 1. Controles principales mutuamente excluyentes
        es_modelo = st.toggle("📝 Generar como Certificado Modelo", key="toggle_modelo", on_change=_on_change_modelo)
        repositorio_masivo = st.toggle("🗄️ Repositorio Masivo", key="toggle_repo", on_change=_on_change_repo)
        modo_manual = st.toggle("🔴 Llenado Manual (Sin PDF)", key="toggle_manual", on_change=_on_change_manual)
        
        if es_modelo:
            st.info("💡 MODO MODELO ACTIVO: Se usarán las plantillas de prueba.")
        elif repositorio_masivo:
            st.info("🗄️ MODO REPOSITORIO ACTIVO: Guías desde Google Sheets.")
        elif modo_manual:
            st.info("🔴 MODO MANUAL ACTIVO: Llenado de datos sin PDF.")
        else:
            st.caption("📄 Modo Estándar: Carga y procesamiento OCR de PDFs.")
        
        st.divider()
        
        # 2. Carga del Menú Dinámico
        mapa_plantillas = obtener_mapa_plantillas_drive(es_modelo=es_modelo)
        
        # 3. Selectores de Empresa y Servicio
        lista_empresas = list(mapa_plantillas.keys())
        
        if not lista_empresas:
            lista_empresas = ["Esperando conexión con Drive..."]
            
        empresa_firma = st.selectbox("Empresa Firmante", options=lista_empresas)
        
        servicios_base = mapa_plantillas.get(empresa_firma, [])
        opciones_finales = []
        
        # ==========================================
        # 🚧 LECTURA DINÁMICA DESDE DRIVE / SHEETS 🚧
        # ==========================================
        if es_modelo:
            # 1. SI ES MODELO: Pasa directamente lo que Drive encuentre (Escalable 100%)
            opciones_finales = servicios_base
            if not opciones_finales:
                opciones_finales = ["No se detectaron plantillas modelo en Drive"]
        else:
            # 2. SI ES NORMAL: Lista exactamente los servicios que tienen plantilla configurada
            for serv in servicios_base:
                if serv and serv not in opciones_finales:
                    opciones_finales.append(serv)
                    
            if not opciones_finales:
                opciones_finales = ["Disposición Final 1"]
                
        if repositorio_masivo and st.session_state.get('repo_tipo_flujo'):
            tipo_flujo_auto = st.session_state['repo_tipo_flujo']
            if tipo_flujo_auto in opciones_finales:
                idx_repo = opciones_finales.index(tipo_flujo_auto)
                sel_tipo = tipo_flujo_auto
            else:
                idx_repo = 0
                sel_tipo = opciones_finales[0] if opciones_finales else tipo_flujo_auto
            tipo_flujo = st.selectbox(
                "Tipo de Certificado", 
                options=opciones_finales, 
                index=idx_repo, 
                disabled=True, 
                help="Definido automáticamente desde la Columna J de las guías en el repositorio"
            )
            tipo_flujo = sel_tipo
        else:
            tipo_flujo = st.selectbox("Tipo de Certificado", options=opciones_finales)

        st.divider()
        if st.button("Limpiar Sesión Activa", key="btn_limpiar_cert", use_container_width=True):
            llaves_protegidas = ['repo', 'usuario_rol', 'usuario_email', 'metricas_exitosos', 'metricas_errores']
            for k in list(st.session_state.keys()):
                if k not in llaves_protegidas:
                    del st.session_state[k]
            st.session_state.uploader_key = st.session_state.get('uploader_key', 0) + 1
            st.rerun()

        if st.session_state.get('usuario_rol') == 'Admin':
            st.divider()
            with st.expander("🛠️ Admin Tools"):
                st.warning("Controles Elevados")
                st.markdown("### 📊 Rendimiento de Sesión")
                col_c1, col_c2 = st.columns(2)
                col_c1.metric(label="Certificados", value=st.session_state.get('metricas_exitosos', 0), delta="Esta sesión")
                col_c2.metric(label="Errores", value=st.session_state.get('metricas_errores', 0), delta="Alertas", delta_color="inverse")
                st.divider()
                if st.button("Forzar Purga de Caché GCP", key="btn_purge_cert", use_container_width=True):
                    st.cache_data.clear()
                    st.success("Toda la Memoria RAM del entorno purgó Sheets y Drive.")


if modulo_actual == "📄 Generador de Certificados":

    # ====================================================================
    # --- BLOQUE 3: UI - Ingesta y Procesamiento de Archivos ---
    # ====================================================================
    if not modo_manual:
        if repositorio_masivo:
            st.subheader("🗄️ Búsqueda en Repositorio Masivo")
            drv, sht = obtener_servicios()
            cat = obtener_catalogo_guias(sht) if sht else {}
            
            if not cat:
                st.info("✅ ¡Todo está al día! No hay certificados pendientes por generar en el repositorio.")
            else:
                c1, c2, c3 = st.columns(3)
                # Caja 1: Empresa
                r_empresa = c1.selectbox("1. Empresa", options=list(cat.keys()), index=None, placeholder="Seleccione...", key="repo_empresa")
                
                # Caja 2: Mes
                opciones_mes = list(cat.get(r_empresa, {}).keys()) if r_empresa else []
                r_mes = c2.selectbox("2. Mes", options=opciones_mes, index=None, placeholder="Seleccione...", disabled=not r_empresa, key="repo_mes")
                
                # Caja 3: Fundo
                opciones_fundo = cat.get(r_empresa, {}).get(r_mes, []) if (r_empresa and r_mes) else []
                r_fundo = c3.selectbox("3. Fundo/Planta", options=opciones_fundo, index=None, placeholder="Seleccione...", disabled=not r_mes, key="repo_fundo")
                
                if st.button("🔍 Buscar Guías en Repositorio", disabled=not (r_empresa and r_fundo and r_mes)):
                    res = buscar_guias_repositorio(sht, r_empresa, r_fundo, r_mes)
                    if res:
                        st.success(f"✅ Se encontraron {len(res)} guías nuevas para procesar.")
                        st.session_state['guias_repo'] = res
                        
                        # Auto-detección desde Columna J de las guías encontradas:
                        tipo_detectado = res[0].get('tipo_detectado', 'Comercialización')
                        tipo_flujo_auto = res[0].get('tipo_flujo', 'Comercialización')
                        st.session_state['repo_tipo_detectado'] = tipo_detectado
                        st.session_state['repo_tipo_flujo'] = tipo_flujo_auto
                        st.session_state['repo_tipo_seleccionado'] = tipo_detectado
                        st.rerun()
                    else:
                        st.warning("No se encontraron guías pendientes para estos filtros.")
                        if 'guias_repo' in st.session_state: del st.session_state['guias_repo']
                        if 'repo_tipo_detectado' in st.session_state: del st.session_state['repo_tipo_detectado']
                        if 'repo_tipo_flujo' in st.session_state: del st.session_state['repo_tipo_flujo']
                        if 'repo_tipo_seleccionado' in st.session_state: del st.session_state['repo_tipo_seleccionado']
                        
                if st.session_state.get('repo_tipo_flujo'):
                    st.info(f"📋 **Tipo detectado (Columna J):** {st.session_state.get('repo_tipo_detectado')} ➔ **Plantilla:** {st.session_state.get('repo_tipo_flujo')}")
                        
                if st.session_state.get('guias_repo'):
                    opciones_nombres = [r['nombre'] for r in st.session_state['guias_repo']]
                    
                    def fmt_guia(nombre):
                        for r in st.session_state.get('guias_repo', []):
                            if r['nombre'] == nombre:
                                return r.get('numero_guia', nombre)
                        return nombre

                    archivos_seleccionados_usuario = st.multiselect(
                        "Filtro: Selecciona Guías a Procesar", 
                        options=opciones_nombres, 
                        default=opciones_nombres,
                        format_func=fmt_guia
                    )
                    
                    if not archivos_seleccionados_usuario:
                        st.warning("Debe seleccionar al menos una guía para procesar.")
                    else:
                        if st.button("🧠 Procesar con IA (OCR)"):
                            with st.spinner(f"Descargando {len(archivos_seleccionados_usuario)} PDFs desde Drive..."):
                                archivos_repo = descargar_guias_drive(drv, archivos_seleccionados_usuario)
                                st.session_state['archivos_mock'] = archivos_repo
                                
                                st.session_state['guias_repo'] = [r for r in st.session_state['guias_repo'] if r['nombre'] in archivos_seleccionados_usuario]
                                st.session_state['procesar_ya'] = True

        if not repositorio_masivo:
            # Verifica que tu línea sea así (usa uploader_key):
            archivos = st.file_uploader("Sube tus guías", type=["pdf"], accept_multiple_files=True, key=f"uploader_{st.session_state.get('uploader_key', 0)}")
            if archivos:
                st.session_state['archivos_subidos_dict'] = {arc.name: arc.getvalue() for arc in archivos}
        else:
            archivos = st.session_state.get('archivos_mock', None)

        # Mostrar error solo si el repositorio intentó descargar y devolvió vacío
        if repositorio_masivo and isinstance(archivos, list) and len(archivos) == 0:
            st.error("❌ Drive rechazó la búsqueda. El archivo no existe o tiene una extensión oculta/diferente.")

        if archivos:
            ejecutar_ocr = False
            if not repositorio_masivo and st.button("🔍 Procesar"):
                ejecutar_ocr = True
            elif repositorio_masivo and st.session_state.get('procesar_ya', False):
                ejecutar_ocr = True
                st.session_state['procesar_ya'] = False # Reset

            if ejecutar_ocr:
                prog = st.progress(0)
                items, grl = [], None
                errores = 0
                total = len(archivos)
                st.session_state['total_pdfs_leidos'] = total
                
                for i, arc in enumerate(archivos):
                    arc_bytes = arc.getvalue() if hasattr(arc, 'getvalue') else arc.read()
                    if hasattr(arc, 'seek'):
                        arc.seek(0)
                    d = procesar_guia_ia_vertex(arc_bytes)
                    if d:
                        if not grl: grl = d 
                        s, f, p = d.get('serie','S/N'), d.get('fecha',''), d.get('vehiculo','')
                        # Auto-corrección proactiva si la IA invirtió serie y vehículo
                        if es_formato_placa(s) and es_formato_guia(p):
                            s, p = p, s
                            d['serie'], d['vehiculo'] = s, p
                        for it in d.get('items', []):
                            it.update({
                                'guia_origen': s, 
                                'fecha_origen': f, 
                                'placa_origen': p,
                                'ocr_partida': d.get('punto_partida', ''),
                                'ocr_llegada': d.get('punto_llegada', '')
                            })
                            items.append(it)
                    else: errores += 1
                    prog.progress((i+1)/total)
                
                time.sleep(0.5); prog.empty()
                
                if items:
                    # AQUÍ GUARDAMOS EN LA CAJA CORRECTA
                    st.session_state['ocr_data'] = grl if grl else {}
                    df = pd.DataFrame(items)
                    for c in ['desc','cant','um','peso','fecha_origen','guia_origen','placa_origen', 'ocr_partida', 'ocr_llegada']:
                        if c not in df.columns: df[c] = ""
                    
                    # --- LIMPIEZA BASE OBLIGATORIA ---
                    df['peso'] = df['peso'].apply(lambda x: formato_inteligente(limpiar_monto(x)))
                    df['cant'] = df['cant'].apply(lambda x: formato_inteligente(limpiar_monto(x)))
                    df['desc'] = df['desc'].astype(str).str.upper()
                    df['um'] = df['um'].apply(lambda x: 'KG' if 'KILO' in str(x).upper() else 'GLN' if 'GALO' in str(x).upper() else 'UNID' if 'UNIDA' in str(x).upper() else str(x).upper())
                    df['desc'] = df['desc'].apply(limpiar_descripcion)
                    
                    df['guia_origen'] = df['guia_origen'].apply(limpiar_guia_o_factura)
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
                    st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                    
                    # Sincronizar datos extraídos hacia los campos de cabecera en sesión
                    st.session_state['form_guia'] = limpiar_guia_o_factura(df['guia_origen'].iloc[0] if not df.empty else grl.get('serie', ''))
                    st.session_state['form_placa'] = df['placa_origen'].iloc[0] if not df.empty else grl.get('vehiculo', '')
                    st.session_state['form_partida'] = formato_nompropio(grl.get('punto_partida', ''))
                    st.session_state['form_dest'] = grl.get('destinatario', '')
                    cli_ocr_tmp = grl.get('cliente') or grl.get('razon_social') or grl.get('empresa') or ""
                    st.session_state['form_cli_ocr'] = str(cli_ocr_tmp).upper().strip()
                    st.session_state['form_ruc_ocr'] = grl.get('ruc_cliente') or grl.get('ruc') or ""
                    st.session_state.pop('form_corr', None)
                    st.session_state.pop('form_fec_emis', None)
                    st.session_state.pop('input_partida_man', None)
                    st.session_state.pop('input_cli_ocr', None)
                    st.session_state.pop('input_ruc_ocr', None)
                    st.session_state.pop('_prev_cli_ocr', None)
                    st.session_state.pop('v_items_working', None)
                    # Resetear estado de generación previa para habilitar el botón
                    st.session_state['generado'] = False
                    st.session_state['generando_cert'] = False
                    st.session_state.pop('res_migracion_sigersol', None)
                    st.session_state.pop('word_buffer', None)
                    st.session_state.pop('nombre_generado', None)
                    st.session_state.pop('subido_drive_link', None)
                    st.session_state.pop('subido_drive_id', None)
                    st.session_state.pop('pdf_unido_link', None)
                    st.session_state.pop('pdf_unido_buffer', None)
                    st.session_state.pop('individual_exitosos', None)
                    st.session_state.pop('individual_fallidos', None)
                    st.session_state.pop('modalidad_individual_finalizada', None)
                    st.session_state['msg_generado'] = False
                    st.session_state['msg_descargado'] = False
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
            st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
            st.session_state['generado'] = False
            st.session_state['generando_cert'] = False
            st.session_state.pop('res_migracion_sigersol', None)
            st.session_state.pop('word_buffer', None)
            st.session_state.pop('nombre_generado', None)
            st.session_state.pop('subido_drive_link', None)
            st.session_state.pop('subido_drive_id', None)
            st.session_state.pop('pdf_unido_link', None)
            st.session_state.pop('pdf_unido_buffer', None)
            st.session_state.pop('individual_exitosos', None)
            st.session_state.pop('individual_fallidos', None)
            st.session_state.pop('modalidad_individual_finalizada', None)
            st.session_state['msg_generado'] = False
            st.session_state['msg_descargado'] = False
            st.session_state['form_guia'] = ''
            st.session_state['form_placa'] = ''
            st.session_state['form_partida'] = ''
            st.session_state['form_partida_man'] = ''
            st.session_state['form_dest'] = ''
            st.session_state['form_cli_ocr'] = ''
            st.session_state['form_ruc_ocr'] = ''
            st.session_state['form_ruc_man'] = ''
            st.session_state.pop('form_corr', None)
            st.session_state.pop('form_fec_emis', None)
            st.session_state.pop('input_partida_man', None)
            st.session_state.pop('sel_manual_empresa', None)
            st.session_state.pop('sel_manual_fundo', None)
            st.session_state.pop('sel_cli_manual', None)
            st.session_state.pop('input_ruc_man', None)
            st.session_state.pop('_prev_empresa_sel', None)
            st.session_state.pop('_prev_fundo_sel', None)
            st.session_state.pop('_prev_emp_arriba_sync', None)
            st.session_state.pop('_prev_cli_man', None)
            st.session_state.pop('v_items_working', None)
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
        # --- MEJORA: CÁLCULO INTELIGENTE DEL CORRELATIVO DESDE SHEETS POR EMPRESA ---
        siguiente_corr = "001" # Valor por defecto si la hoja está vacía
        try:
            from src.services.google_service import leer_sheet_seguro, buscar_config_empresa
            import pandas as pd
            
            cfg_emisor = buscar_config_empresa(empresa_firma)
            hoja_hist_emp = cfg_emisor.get('hoja_historial', 'Historial')
            
            # Si el usuario cambió de empresa en el selector, resetear form_corr para recalcular
            if st.session_state.get('last_empresa_calculo_corr') != empresa_firma:
                st.session_state['last_empresa_calculo_corr'] = empresa_firma
                st.session_state['form_corr'] = ""

            df_historial = leer_sheet_seguro(hoja_hist_emp)
            
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
        if 'form_corr' not in st.session_state or not st.session_state.get('form_corr'):
            st.session_state['form_corr'] = siguiente_corr
        v_corr = c1.text_input("Correlativo", value=st.session_state['form_corr'], key="form_corr")
        
        fecha_base = grl.get('fecha', datetime.now().strftime("%d/%m/%Y"))
        
        # --- Magia Robusta: Selección automática forzada (CORREGIDO) ---
        with c2:
            if "Comercialización" in tipo_flujo:
                cliente_crudo_tmp = grl.get('cliente') or grl.get('razon_social') or grl.get('empresa') or ""
                cliente_upper = str(cliente_crudo_tmp).upper().strip()
                
                if ("PROSEMBRA" in cliente_upper or "VILLACURI" in cliente_upper.replace(" ", "")) and str(tipo_flujo).upper().strip() == "COMERCIALIZACIÓN":
                    st.info("📅 REGLA ESPECIAL (HOY)")
                    f_calc = (datetime.utcnow() - timedelta(hours=5)).strftime("%d/%m/%Y")
                else:
                    st.info("📅 COMERCIALIZACIÓN (FIN DE MES)")
                    f_calc = obtener_fin_de_mes(fecha_base)
                tipo_op = "Comercialización"
            else:
                st.info("📅 DISPOSICIÓN FINAL +2")
                try:
                    f_calc = (datetime.strptime(fecha_base, "%d/%m/%Y") + timedelta(days=2)).strftime("%d/%m/%Y")
                except ValueError:
                    f_calc = (datetime.now() + timedelta(days=2)).strftime("%d/%m/%Y")
                tipo_op = "Disposición Final"

            if 'form_fec_emis' not in st.session_state or st.session_state.get('_tipo_flujo_calc') != tipo_flujo:
                st.session_state['form_fec_emis'] = f_calc
                st.session_state['_tipo_flujo_calc'] = tipo_flujo
            v_fec_emis = st.text_input("F. Emisión", value=st.session_state['form_fec_emis'], key="form_fec_emis")
        
        # Extraemos la guía y placa LIMPIAS desde la tabla procesada
        guia_limpia = df_items['guia_origen'].iloc[0] if (not df_items.empty and 'guia_origen' in df_items.columns) else grl.get('serie', '')
        guia_limpia = limpiar_guia_o_factura(guia_limpia)
        placa_limpia = df_items['placa_origen'].iloc[0] if (not df_items.empty and 'placa_origen' in df_items.columns) else grl.get('vehiculo', '')
        
        if 'form_guia' not in st.session_state:
            st.session_state['form_guia'] = guia_limpia
        if 'form_placa' not in st.session_state:
            st.session_state['form_placa'] = placa_limpia

        val_guia_mostrada = st.session_state['form_guia'] if modo_manual else guia_limpia
        val_placa_mostrada = st.session_state['form_placa'] if modo_manual else placa_limpia

        # Si estamos en manual, siguen desbloqueadas para escribir
        v_guia_in = c3.text_input("Guía", value=val_guia_mostrada, disabled=not modo_manual, key="input_guia_txt")
        v_guia = limpiar_guia_o_factura(v_guia_in)
        st.session_state['form_guia'] = v_guia

        v_placa = c4.text_input("Placa", value=val_placa_mostrada, disabled=not modo_manual, key="input_placa_txt")
        st.session_state['form_placa'] = v_placa

        # --- Diagnóstico Temprano de Guía vs Placa ---
        estado_diag, msg_diag = verificar_consistencia_guia_placa(v_guia, v_placa)
        if estado_diag != 'OK':
            c_diag1, c_diag2 = st.columns([3, 1])
            with c_diag1:
                st.warning(f"⚠️ **Inconsistencia detectada:** {msg_diag}")
            with c_diag2:
                if st.button("🔄 Intercambiar Placa ⇄ Guía", key="btn_swap_cabecera", use_container_width=True):
                    df_sw = st.session_state.get('v_items_working', st.session_state.get('df_items', pd.DataFrame())).copy()
                    if not df_sw.empty and 'guia_origen' in df_sw.columns and 'placa_origen' in df_sw.columns:
                        tmp_g = df_sw['guia_origen'].copy()
                        df_sw['guia_origen'] = df_sw['placa_origen']
                        df_sw['placa_origen'] = tmp_g
                        st.session_state['df_items'] = df_sw
                        st.session_state['v_items_working'] = df_sw
                        st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                    if 'ocr_data' in st.session_state and isinstance(st.session_state['ocr_data'], dict):
                        ocr_sw = st.session_state['ocr_data'].copy()
                        s_sw = ocr_sw.get('serie', '')
                        ocr_sw['serie'] = ocr_sw.get('vehiculo', '')
                        ocr_sw['vehiculo'] = s_sw
                        st.session_state['ocr_data'] = ocr_sw
                    # Sincronizar inputs de cabecera
                    tmp_h_g = st.session_state.get('form_guia', '')
                    st.session_state['form_guia'] = st.session_state.get('form_placa', '')
                    st.session_state['form_placa'] = tmp_h_g
                    st.toast("✅ Valores de Guía y Placa intercambiados correctamente.")
                    st.rerun()

        if modo_manual:
            # --- AUTOCOMPLETADO DE DIRECCIÓN EXCLUSIVO PARA LLENADO MANUAL ---
            from src.services.google_service import leer_sheet_seguro
            df_clientes_bd = repo.get("clientes")
            if df_clientes_bd is None or df_clientes_bd.empty:
                df_clientes_bd = leer_sheet_seguro("CLIENTES")
                repo["clientes"] = df_clientes_bd

            df_dir_bd = repo.get("direcciones")
            if df_dir_bd is None or df_dir_bd.empty:
                df_dir_bd = leer_sheet_seguro("Direcciones")
                repo["direcciones"] = df_dir_bd

            # Lista alfabética unificada de empresas desde CLIENTES y Direcciones
            col_emp_cli = next((c for c in df_clientes_bd.columns if 'EMPRESA' in str(c).upper() or 'CLIENTE' in str(c).upper()), 'EMPRESA') if not df_clientes_bd.empty else 'EMPRESA'
            col_emp_dir = next((c for c in df_dir_bd.columns if 'EMPRESA' in str(c).upper() or 'CLIENTE' in str(c).upper()), 'EMPRESA') if not df_dir_bd.empty else 'EMPRESA'
            
            lista_emp_cli = df_clientes_bd[col_emp_cli].dropna().astype(str).str.strip().str.upper().unique().tolist() if not df_clientes_bd.empty and col_emp_cli in df_clientes_bd.columns else []
            lista_emp_dir = df_dir_bd[col_emp_dir].dropna().astype(str).str.strip().str.upper().unique().tolist() if not df_dir_bd.empty and col_emp_dir in df_dir_bd.columns else []
            empresas_unicas = sorted(list(set(lista_emp_cli + lista_emp_dir)))

            c_emp_man, c_fnd_man = st.columns(2)
            with c_emp_man:
                empresa_sel_man = st.selectbox(
                    "🏢 Empresa (Pestaña Clientes):",
                    options=["-- Seleccionar Empresa --"] + empresas_unicas,
                    key="sel_manual_empresa"
                )

            opciones_fundos = []
            mapa_fundo_dir = {}
            if empresa_sel_man and empresa_sel_man != "-- Seleccionar Empresa --":
                if not df_dir_bd.empty:
                    c_emp_dir_b = next((c for c in df_dir_bd.columns if 'EMPRESA' in str(c).upper() or 'CLIENTE' in str(c).upper()), 'EMPRESA')
                    c_fnd_dir_b = next((c for c in df_dir_bd.columns if 'FUNDO' in str(c).upper() or 'PLANTA' in str(c).upper()), 'FUNDO/PLANTA')
                    c_dir_dir_b = next((c for c in df_dir_bd.columns if 'DIRECCION' in str(c).upper() or 'DIRECCIÓN' in str(c).upper()), 'DIRECCION')
                    
                    mask_emp = df_dir_bd[c_emp_dir_b].astype(str).str.strip().str.upper() == empresa_sel_man.strip().upper()
                    df_emp_dir = df_dir_bd[mask_emp]
                    for _, r_dir in df_emp_dir.iterrows():
                        fnd_nom = str(r_dir.get(c_fnd_dir_b, '')).strip()
                        dir_txt = str(r_dir.get(c_dir_dir_b, '')).strip()
                        if fnd_nom and fnd_nom.upper() not in ['NAN', 'NONE']:
                            opciones_fundos.append(fnd_nom)
                            mapa_fundo_dir[fnd_nom] = dir_txt
                        elif dir_txt and dir_txt.upper() not in ['NAN', 'NONE']:
                            opciones_fundos.append("DIRECCIÓN PRINCIPAL")
                            mapa_fundo_dir["DIRECCIÓN PRINCIPAL"] = dir_txt

                # Si no está en Direcciones, buscar en Clientes
                if not mapa_fundo_dir and not df_clientes_bd.empty:
                    c_emp_cli_b = next((c for c in df_clientes_bd.columns if 'EMPRESA' in str(c).upper() or 'CLIENTE' in str(c).upper()), 'EMPRESA')
                    df_emp_c = df_clientes_bd[df_clientes_bd[c_emp_cli_b].astype(str).str.strip().str.upper() == empresa_sel_man.strip().upper()]
                    if not df_emp_c.empty:
                        d_cert = str(df_emp_c.iloc[0].get('DIRECCION CERTIFICADO', '')).strip()
                        d_fisc = str(df_emp_c.iloc[0].get('DOMICILIO FISCAL', '')).strip()
                        d_uno = str(df_emp_c.iloc[0].get('DIRECCION 1', '')).strip()
                        dir_hallada = d_cert if d_cert and d_cert.upper() not in ['NAN', 'NONE'] else (d_fisc if d_fisc and d_fisc.upper() not in ['NAN', 'NONE'] else d_uno)
                        if dir_hallada and dir_hallada.upper() not in ['NAN', 'NONE']:
                            opciones_fundos.append("DIRECCIÓN REGISTRADA")
                            mapa_fundo_dir["DIRECCIÓN REGISTRADA"] = dir_hallada

            with c_fnd_man:
                if opciones_fundos:
                    fundo_sel_man = st.selectbox(
                        "🌱 Fundo / Planta:",
                        options=opciones_fundos,
                        key="sel_manual_fundo"
                    )
                else:
                    fundo_sel_man = None
                    st.selectbox("🌱 Fundo / Planta:", options=["(Sin fundos registrados)"], disabled=True, key="sel_manual_fundo_disabled")

            dir_autocompletada = ""
            if fundo_sel_man and fundo_sel_man in mapa_fundo_dir:
                dir_autocompletada = mapa_fundo_dir[fundo_sel_man]

            cambio_ubicacion = (
                empresa_sel_man != st.session_state.get('_prev_empresa_sel') or
                fundo_sel_man != st.session_state.get('_prev_fundo_sel')
            )
            if 'input_partida_man' not in st.session_state or cambio_ubicacion:
                val_partida_inicial = dir_autocompletada if dir_autocompletada else grl.get('punto_partida', '')
                dir_formateada = formato_nompropio(val_partida_inicial)
                st.session_state['form_partida_man'] = dir_formateada
                st.session_state['input_partida_man'] = dir_formateada
                st.session_state['_prev_empresa_sel'] = empresa_sel_man
                st.session_state['_prev_fundo_sel'] = fundo_sel_man

            v_partida = st.text_input(
                "Partida (Dirección)", 
                key="input_partida_man"
            )
            st.session_state['form_partida_man'] = v_partida
        else:
            if 'form_partida' not in st.session_state or not st.session_state.get('form_partida'):
                st.session_state['form_partida'] = formato_nompropio(grl.get('punto_partida', ''))
            v_partida = st.text_input("Partida", value=st.session_state['form_partida'], key="input_partida_ocr")
            st.session_state['form_partida'] = v_partida
        v_llegada = ""  # Dirección de llegada predeterminada en la plantilla
        
        if 'form_dest' not in st.session_state:
            st.session_state['form_dest'] = grl.get('destinatario', '')
        v_dest = st.text_input("Destinatario", value=st.session_state['form_dest'], key="input_dest_txt")
        st.session_state['form_dest'] = v_dest

        # Inicializar configuración de columnas si no existe
        if 'tabla_columnas_config' not in st.session_state or not isinstance(st.session_state.get('tabla_columnas_config'), list):
            st.session_state['tabla_columnas_config'] = [c.copy() for c in DEFAULT_COLUMNAS_CONFIG]

        # Panel para Mover, Editar Encabezados y Eliminar Columnas
        with st.expander("🛠️ Personalizar Columnas y Encabezados (Mover / Editar / Eliminar)", expanded=False):
            st.markdown("💡 *Edita los nombres de los encabezados, reordena las columnas con las flechas ⬅️ / ➡️ o elimina las que no necesites. La tabla interactiva y el Certificado Word se generarán respetando el orden y los nombres configurados.*")
            
            cols_activas = st.session_state['tabla_columnas_config']
            col_a_borrar = None
            col_a_mover_izq = None
            col_a_mover_der = None
            cambio_encabezado = False
            
            # Encabezado guía de las acciones
            h_c1, h_c2, h_c3, h_c4, h_c5 = st.columns([1.6, 3.2, 0.6, 0.6, 0.8])
            with h_c1:
                st.caption("**ID Columna**")
            with h_c2:
                st.caption("**Nombre en Encabezado**")
            with h_c3:
                st.caption("**Izq.**")
            with h_c4:
                st.caption("**Der.**")
            with h_c5:
                st.caption("**Quitar**")
            
            for idx_c, col_cfg in enumerate(cols_activas):
                col_c1, col_c2, col_c3, col_c4, col_c5 = st.columns([1.6, 3.2, 0.6, 0.6, 0.8])
                with col_c1:
                    st.markdown(f"**Col. {idx_c + 1}:** `{col_cfg['id']}`")
                with col_c2:
                    nuevo_hdr = st.text_input(
                        f"Encabezado para {col_cfg['id']}", 
                        value=col_cfg.get('nombre', col_cfg.get('id', '')), 
                        key=f"hdr_txt_{col_cfg['id']}",
                        label_visibility="collapsed"
                    )
                    if nuevo_hdr != col_cfg.get('nombre'):
                        if st.session_state.get('v_items_working') is not None:
                            st.session_state['df_items'] = st.session_state['v_items_working'].copy()
                        col_cfg['nombre'] = nuevo_hdr
                        cambio_encabezado = True
                with col_c3:
                    if st.button("⬅️", key=f"btn_mv_left_{col_cfg['id']}", disabled=(idx_c == 0), help=f"Mover '{col_cfg.get('nombre')}' hacia la izquierda"):
                        col_a_mover_izq = idx_c
                with col_c4:
                    if st.button("➡️", key=f"btn_mv_right_{col_cfg['id']}", disabled=(idx_c == len(cols_activas) - 1), help=f"Mover '{col_cfg.get('nombre')}' hacia la derecha"):
                        col_a_mover_der = idx_c
                with col_c5:
                    if st.button("🗑️", key=f"btn_del_col_{col_cfg['id']}", help=f"Eliminar la columna '{col_cfg.get('nombre')}' de la tabla"):
                        col_a_borrar = col_cfg['id']

            if cambio_encabezado:
                st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                st.toast("✅ Encabezado actualizado.")
                st.rerun()

            if col_a_mover_izq is not None and col_a_mover_izq > 0:
                if st.session_state.get('v_items_working') is not None:
                    st.session_state['df_items'] = st.session_state['v_items_working'].copy()
                idx_prev = col_a_mover_izq - 1
                cols_activas[col_a_mover_izq], cols_activas[idx_prev] = cols_activas[idx_prev], cols_activas[col_a_mover_izq]
                st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                st.toast("✅ Columna movida hacia la izquierda.")
                st.rerun()

            if col_a_mover_der is not None and col_a_mover_der < len(cols_activas) - 1:
                if st.session_state.get('v_items_working') is not None:
                    st.session_state['df_items'] = st.session_state['v_items_working'].copy()
                idx_next = col_a_mover_der + 1
                cols_activas[col_a_mover_der], cols_activas[idx_next] = cols_activas[idx_next], cols_activas[col_a_mover_der]
                st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                st.toast("✅ Columna movida hacia la derecha.")
                st.rerun()

            if col_a_borrar:
                if len(cols_activas) <= 1:
                    st.warning("⚠️ Debe conservar al menos una columna en la tabla.")
                else:
                    if st.session_state.get('v_items_working') is not None:
                        st.session_state['df_items'] = st.session_state['v_items_working'].copy()
                    st.session_state['tabla_columnas_config'] = [c for c in cols_activas if c['id'] != col_a_borrar]
                    st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                    st.toast("✅ Columna eliminada de la tabla.")
                    st.rerun()

            st.divider()
            c_rest1, c_rest2 = st.columns([2, 1])
            with c_rest1:
                ids_activas = [c['id'] for c in st.session_state['tabla_columnas_config']]
                cols_disponibles_restaurar = [c for c in DEFAULT_COLUMNAS_CONFIG if c['id'] not in ids_activas]
                if cols_disponibles_restaurar:
                    c_sel_r, c_btn_r = st.columns([3, 1])
                    with c_sel_r:
                        col_sel_res = st.selectbox(
                            "Restaurar columna eliminada:",
                            options=[c['id'] for c in cols_disponibles_restaurar],
                            format_func=lambda cid: next(c['nombre'] for c in cols_disponibles_restaurar if c['id'] == cid),
                            key="sel_col_restaurar"
                        )
                    with c_btn_r:
                        st.write("")
                        if st.button("➕ Restaurar", key="btn_ejecutar_restaurar"):
                            if st.session_state.get('v_items_working') is not None:
                                st.session_state['df_items'] = st.session_state['v_items_working'].copy()
                            item_res = next(c.copy() for c in DEFAULT_COLUMNAS_CONFIG if c['id'] == col_sel_res)
                            st.session_state.pop(f"hdr_txt_{col_sel_res}", None)
                            st.session_state['tabla_columnas_config'].append(item_res)
                            st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                            st.toast("✅ Columna restaurada.")
                            st.rerun()
            with c_rest2:
                if st.button("🔄 Restablecer Predeterminados", key="btn_reset_default_cols", help="Restaura las 7 columnas originales con sus nombres por defecto"):
                    if st.session_state.get('v_items_working') is not None:
                        st.session_state['df_items'] = st.session_state['v_items_working'].copy()
                    for c_def in DEFAULT_COLUMNAS_CONFIG:
                        st.session_state.pop(f"hdr_txt_{c_def['id']}", None)
                    st.session_state['tabla_columnas_config'] = [c.copy() for c in DEFAULT_COLUMNAS_CONFIG]
                    st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                    st.toast("✅ Columnas restablecidas a valores predeterminados.")
                    st.rerun()

        cols_activas = st.session_state['tabla_columnas_config']
        ids_activas = [c['id'] for c in cols_activas]
        
        # Asegurarse de que el DataFrame contenga las columnas activas
        for cid in ids_activas:
            if cid not in df_items.columns:
                df_items[cid] = ""
                
        # Construir configuración de columnas con los encabezados personalizados usando Column universal
        col_config_editor = {}
        for c in cols_activas:
            nombre_label = str(c.get('nombre', '')).strip() or c['id']
            col_config_editor[c['id']] = st.column_config.Column(label=nombre_label)
            
        df_para_editor = df_items[ids_activas].copy()

        if 'editor_version' not in st.session_state:
            st.session_state['editor_version'] = 0

        editor_key = f"v_items_editor_v{st.session_state['editor_version']}"

        v_items_df = st.data_editor(
            df_para_editor,
            column_order=ids_activas,
            num_rows="dynamic",
            use_container_width=True,
            column_config=col_config_editor,
            key=editor_key
        )

        v_items_df = v_items_df.copy()
        # Limpieza automática de guías y facturas en la tabla de trabajo
        if 'guia_origen' in v_items_df.columns:
            v_items_df['guia_origen'] = v_items_df['guia_origen'].apply(limpiar_guia_o_factura)

        # Preservar siempre los datos editados en memoria de trabajo sin alterar df_items
        # para que Streamlit mantenga el editor estable sin recargarlo ni borrar celdas
        st.session_state['v_items_working'] = v_items_df

        # --- PANEL DE GESTIÓN Y ELIMINACIÓN DE FILAS ---
        if not v_items_df.empty:
            c_del1, c_del2, c_del3 = st.columns([3, 1.2, 1.2])
            opciones_filas = []
            for idx_f, r in v_items_df.iterrows():
                desc_val = str(r.get('desc', '')).strip()
                desc_corta = (desc_val[:32] + '...') if len(desc_val) > 32 else (desc_val or "Sin descripción")
                cant_val = str(r.get('cant', '')).strip()
                um_val = str(r.get('um', '')).strip()
                peso_val = str(r.get('peso', '')).strip()
                guia_val = str(r.get('guia_origen', '')).strip()
                det_extra = f" | Guía: {guia_val}" if guia_val else ""
                opciones_filas.append(f"Fila {idx_f + 1}: {desc_corta} ({cant_val} {um_val} - {peso_val} KG){det_extra}")

            with c_del1:
                fila_a_eliminar_idx = st.selectbox(
                    "🗑️ Seleccionar fila a eliminar:",
                    options=list(range(len(v_items_df))),
                    format_func=lambda i: opciones_filas[i] if i < len(opciones_filas) else f"Fila {i + 1}",
                    key="sel_eliminar_fila"
                )

            with c_del2:
                st.write("")
                st.write("")
                if st.button("🗑️ Eliminar Fila", type="secondary", use_container_width=True, key="btn_borrar_fila_sel"):
                    if len(v_items_df) <= 1:
                        st.warning("⚠️ La tabla debe conservar al menos una fila para generar el certificado.")
                    else:
                        df_mod = v_items_df.drop(index=fila_a_eliminar_idx).reset_index(drop=True)
                        st.session_state['df_items'] = df_mod
                        st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                        st.toast(f"✅ Fila {fila_a_eliminar_idx + 1} eliminada correctamente.")
                        st.rerun()

            with c_del3:
                st.write("")
                st.write("")
                if st.button("🧹 Limpiar Vacías", use_container_width=True, key="btn_limpiar_filas_vacias", help="Elimina filas sin descripción o completamente en blanco"):
                    if 'desc' in v_items_df.columns:
                        mask_validas = v_items_df['desc'].astype(str).str.strip().ne('') & v_items_df['desc'].astype(str).str.upper().ne('NAN')
                        cant_vacias = len(v_items_df) - mask_validas.sum()
                        if cant_vacias > 0:
                            df_sin_vacias = v_items_df[mask_validas].reset_index(drop=True)
                            if df_sin_vacias.empty:
                                st.warning("⚠️ Todas las filas estaban vacías. Se conservó 1 fila para editar.")
                            else:
                                st.session_state['df_items'] = df_sin_vacias
                                st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                                st.toast(f"🧹 Se eliminaron {cant_vacias} fila(s) vacía(s).")
                                st.rerun()
                        else:
                            st.info("No hay filas vacías para eliminar.")

        from src.services.google_service import leer_sheet_seguro
        if 'repo' not in st.session_state:
            st.session_state.repo = {
                "emisores": leer_sheet_seguro("EMPRESAS"),
                "clientes": leer_sheet_seguro("CLIENTES"),
                "servicios": leer_sheet_seguro("SERVICIOS")
            }
        repo = st.session_state.repo

        # --- CATÁLOGO DE SERVICIOS POR PLANTILLA (COMERCIALIZACIÓN vs DISPOSICIÓN FINAL) ---
        sec_actual = "COMERCIALIZACION" if "comercializa" in str(tipo_flujo).strip().lower() else "SERVICIOS"
        otra_sec = "SERVICIOS" if sec_actual == "COMERCIALIZACION" else "COMERCIALIZACION"
        catalogo_sec = obtener_catalogo_servicios_por_categoria(repo.get('servicios'))

        # Opciones prioritarias de la plantilla actual + resto de opciones disponibles
        opciones_tit = catalogo_sec[sec_actual]["titulos"] + [t for t in catalogo_sec[otra_sec]["titulos"] if t not in catalogo_sec[sec_actual]["titulos"]]
        if not opciones_tit: opciones_tit = ["CERTIFICADO DE MANEJO"]

        opciones_serv = catalogo_sec[sec_actual]["servicios"] + [s for s in catalogo_sec[otra_sec]["servicios"] if s not in catalogo_sec[sec_actual]["servicios"]]
        if not opciones_serv: opciones_serv = ["Sin Datos"]

        opciones_res = catalogo_sec[sec_actual]["residuos"] + [r for r in catalogo_sec[otra_sec]["residuos"] if r not in catalogo_sec[sec_actual]["residuos"]]
        if not opciones_res: opciones_res = ["Sin Datos"]

        # 2. COLUMNAS: Partimos la pantalla en 2 mitades
        c_a, c_b = st.columns(2, gap="large")

        # ==========================================
        # ⬅️ LADO IZQUIERDO: EMISOR 
        # ==========================================
        with c_a:
            st.subheader("Emisor")
            
            if es_modelo:
                from src.services.google_service import buscar_config_empresa
                cfg_mod = buscar_config_empresa(empresa_firma) if empresa_firma != "MODELO" else {}
                nom_def = cfg_mod.get('nombre') or (empresa_firma if empresa_firma != "MODELO" else "")
                ruc_def = cfg_mod.get('ruc', "")
                reg_def = cfg_mod.get('reg', "")

                st.caption("💡 Datos de la empresa que emite (Modo Modelo):")
                emisor_nombre = st.text_input("Nombre Emisor", value=nom_def, key="em_nom_mod")
                emisor_ruc = st.text_input("RUC Emisor", value=ruc_def, key="em_ruc_mod")
                emisor_reg = st.text_input("Registro Emisor", value=reg_def, key="em_reg_mod")
            else:
                # --- SI ES NORMAL: Lógica de búsqueda dinámica en configuración de empresas ---
                info_emisor = None
                try:
                    from src.services.google_service import buscar_config_empresa
                    cfg_emisor = buscar_config_empresa(empresa_firma)
                    if cfg_emisor and cfg_emisor.get('ruc'):
                        info_emisor = {
                            'ruc': cfg_emisor['ruc'],
                            'reg': cfg_emisor['reg']
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
                    st.warning(f"La empresa '{empresa_firma}' no existe en la pestaña 'EMPRESAS' del Excel.")
                    emisor_nombre = empresa_firma
                    emisor_ruc = ""
                    emisor_reg = ""

            # TÍTULO: Va exactamente debajo del registro del Emisor
            v_tit = st.selectbox(
                "Título", 
                options=opciones_tit,
                key=f"sb_titulo_{tipo_flujo}"
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
                
                lista_nombres_cli = list(diccionario_clientes.keys())
                lista_extra_emp = [e for e in locals().get('empresas_unicas', []) if e != "-- Seleccionar Empresa --"]
                todos_clientes = sorted(list(set(lista_nombres_cli + lista_extra_emp)))
                opciones_clientes = [""] + todos_clientes
                
                emp_sel_arriba = locals().get('empresa_sel_man', '')
                
                # Sincronización automática con la empresa elegida arriba en modo manual
                cambio_emp_arriba = (emp_sel_arriba != st.session_state.get('_prev_emp_arriba_sync'))
                if cambio_emp_arriba:
                    if emp_sel_arriba and emp_sel_arriba != "-- Seleccionar Empresa --":
                        match_op = next((op for op in opciones_clientes if op.strip().upper() == emp_sel_arriba.strip().upper()), None)
                        if match_op:
                            st.session_state['sel_cli_manual'] = match_op
                    st.session_state['_prev_emp_arriba_sync'] = emp_sel_arriba

                # Calcular índice por defecto
                idx_default_cli = 0
                cli_actual = st.session_state.get('sel_cli_manual', '')
                if cli_actual and cli_actual in opciones_clientes:
                    idx_default_cli = opciones_clientes.index(cli_actual)
                elif emp_sel_arriba and emp_sel_arriba != "-- Seleccionar Empresa --":
                    for i_op, op in enumerate(opciones_clientes):
                        if op.strip().upper() == emp_sel_arriba.strip().upper():
                            idx_default_cli = i_op
                            break

                v_cli = st.selectbox(
                    "Cliente (Desde Base de Datos)", 
                    options=opciones_clientes, 
                    index=idx_default_cli, 
                    key="sel_cli_manual"
                )
                
                # Buscar RUC correspondiente al cliente seleccionado
                ruc_encontrado = diccionario_clientes.get(v_cli, "") if v_cli else ""
                df_cli_actual = locals().get('df_clientes_bd', pd.DataFrame())
                if not ruc_encontrado and v_cli and not df_cli_actual.empty:
                    col_r_emp = next((c for c in df_cli_actual.columns if 'EMPRESA' in str(c).upper() or 'CLIENTE' in str(c).upper()), 'EMPRESA')
                    col_r_ruc = next((c for c in df_cli_actual.columns if 'RUC' in str(c).upper()), 'RUC')
                    if col_r_emp in df_cli_actual.columns and col_r_ruc in df_cli_actual.columns:
                        match_row = df_cli_actual[df_cli_actual[col_r_emp].astype(str).str.strip().str.upper() == str(v_cli).strip().upper()]
                        if not match_row.empty:
                            ruc_encontrado = str(match_row.iloc[0].get(col_r_ruc, '')).strip()

                cambio_cli = (v_cli != st.session_state.get('_prev_cli_man'))
                if cambio_cli or 'input_ruc_man' not in st.session_state:
                    st.session_state['input_ruc_man'] = ruc_encontrado
                    st.session_state['form_ruc_man'] = ruc_encontrado
                    st.session_state['_prev_cli_man'] = v_cli
                
                v_ruc_c = st.text_input("RUC Cliente", key="input_ruc_man")
                st.session_state['form_ruc_man'] = v_ruc_c
            else:
                from src.services.google_service import obtener_clientes_desde_sheets
                diccionario_clientes = obtener_clientes_desde_sheets()
                
                cli_defecto = cliente_ocr or st.session_state.get('repo_empresa', '')
                if 'input_cli_ocr' not in st.session_state:
                    st.session_state['input_cli_ocr'] = st.session_state.get('form_cli_ocr') or cli_defecto
                v_cli = st.text_input("Cliente (Extraído)", key="input_cli_ocr")
                st.session_state['form_cli_ocr'] = v_cli
                
                # Autocompletado robusto: Si hace match con BD, usa su RUC. Si no, usa el extraído crudo de IA
                ruc_calculado = diccionario_clientes.get(str(v_cli).strip().upper(), ruc_crudo)
                cambio_cli_ocr = (v_cli != st.session_state.get('_prev_cli_ocr'))
                if cambio_cli_ocr or 'input_ruc_ocr' not in st.session_state:
                    st.session_state['input_ruc_ocr'] = ruc_calculado
                    st.session_state['form_ruc_ocr'] = ruc_calculado
                    st.session_state['_prev_cli_ocr'] = v_cli
                v_ruc_c = st.text_input("RUC Cliente", key="input_ruc_ocr")
                st.session_state['form_ruc_ocr'] = v_ruc_c
            
            v_serv = st.selectbox(
                "Servicio", 
                options=opciones_serv,
                key=f"sb_servicio_{tipo_flujo}"
            )
            v_res = st.selectbox(
                "Residuo", 
                options=opciones_res,
                key=f"sb_residuo_{tipo_flujo}"
            )
            

    # ====================================================================
    # --- BLOQUE 5: UI - Generación de Word, Descarga y Registro en Sheets ---
    # ====================================================================
    st.divider()

    def inyectar_estado_sheets_robusto(numero_de_guia):
        # Actualización robusta de estado en Sheets
        try:
            from zoneinfo import ZoneInfo
            from src.config.settings import ID_SHEET_REPOSITORIO
            from src.services.google_service import obtener_servicios
            
            _, sht_drv = obtener_servicios()
            target_ws_title = "Guias_recibidas"
            
            r = sht_drv.spreadsheets().values().get(spreadsheetId=ID_SHEET_REPOSITORIO, range=f"'{target_ws_title}'!B:B").execute()
            col_guias = r.get('values', [])
            
            fila_encontrada = None
            import re
            
            def norm_g(s):
                t = re.sub(r'(?i)[nº°\s_]', '', str(s)).lower()
                return "".join([p.lstrip('0') if p.isdigit() else p for p in re.findall(r'[a-z]+|[0-9]+', t)])

            guia_clean = norm_g(numero_de_guia)
            
            for i, val in enumerate(col_guias):
                if val and len(val) > 0:
                    celda_clean = norm_g(val[0])
                    if celda_clean == guia_clean or (len(guia_clean) >= 4 and guia_clean in celda_clean):
                        fila_encontrada = i + 1
                        break
                    
            if fila_encontrada:
                body = {"values": [[f"✅ Nuevo: {datetime.now(ZoneInfo('America/Lima')).strftime('%d/%m/%Y %H:%M:%S')}"]]}
                sht_drv.spreadsheets().values().update(
                    spreadsheetId=ID_SHEET_REPOSITORIO, range=f"'{target_ws_title}'!H{fila_encontrada}",
                    valueInputOption="USER_ENTERED", body=body
                ).execute()
            else:
                st.warning(f"⚠️ Sheets: No se encontró la guía '{numero_de_guia}' en la columna B de '{target_ws_title}'.")
                
        except Exception as e:
            st.error(f"❌ Error API Sheets con guía {numero_de_guia}: {str(e)}")

    if 'msg_generado' not in st.session_state: st.session_state.msg_generado = False
    if 'msg_descargado' not in st.session_state: st.session_state.msg_descargado = False

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

    if v_df_seguro is not None and not v_df_seguro.empty:

        # --- FIX PARA LLENADO MANUAL ---
        # Propagar la guía y placa ingresadas manualmente hacia la tabla para su registro
        if modo_manual:
            if 'guia_origen' in v_items_df.columns:
                v_items_df['guia_origen'] = v_items_df['guia_origen'].apply(lambda x: v_guia if str(x).strip() in ['', 'None', 'nan'] else limpiar_guia_o_factura(x))
            if 'placa_origen' in v_items_df.columns:
                v_items_df['placa_origen'] = v_items_df['placa_origen'].apply(lambda x: v_placa if str(x).strip() in ['', 'None', 'nan'] else x)
        else:
            if 'guia_origen' in v_items_df.columns:
                v_items_df['guia_origen'] = v_items_df['guia_origen'].apply(limpiar_guia_o_factura)
        st.session_state['v_items_working'] = v_items_df
        guias_unicas_prev = [g for g in v_items_df['guia_origen'].unique() if str(g).strip() not in ['None', '', 'nan']] if 'guia_origen' in v_items_df.columns else []
        if len(guias_unicas_prev) > 1:
            modalidad_gen = st.radio("Modalidad de Generación", [
                "Agrupada (Fusionar las guías seleccionadas en 1 solo certificado)", 
                "Individual (Crear un certificado separado por cada guía seleccionada)"
            ])
        else:
            modalidad_gen = "Agrupada (Fusionar las guías seleccionadas en 1 solo certificado)"
        
        # --- VERIFICACIÓN DE CONSISTENCIA ENTRE PLACA Y GUÍA ANTES DE GENERAR ---
        inconsistencias_tabla = validar_inconsistencias_df(v_items_df)
        estado_cab, msg_cab = verificar_consistencia_guia_placa(v_guia, v_placa)
        hay_error_placa_guia = (estado_cab != 'OK') or bool(inconsistencias_tabla)

        confirmar_forzar = False
        if hay_error_placa_guia:
            with st.container(border=True):
                st.error("⛔ **Atención: Se detectó un posible cruce entre Número de Guía y Placa**")
                if estado_cab != 'OK':
                    st.markdown(f"- **En los campos principales:** {msg_cab}")
                for inc in inconsistencias_tabla:
                    st.markdown(f"- **Fila {inc['fila']} de la tabla:** {inc['detalle']}")

                c_fix1, c_fix2 = st.columns([1, 1])
                with c_fix1:
                    if st.button("🔄 Intercambiar Placa y Guía en toda la tabla", key="btn_fix_swap_gen", use_container_width=True):
                        df_tmp = v_items_df.copy()
                        if 'guia_origen' in df_tmp.columns and 'placa_origen' in df_tmp.columns:
                            tmp_col = df_tmp['guia_origen'].copy()
                            df_tmp['guia_origen'] = df_tmp['placa_origen']
                            df_tmp['placa_origen'] = tmp_col
                            st.session_state['df_items'] = df_tmp
                            st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
                        if 'ocr_data' in st.session_state and isinstance(st.session_state['ocr_data'], dict):
                            st.session_state['ocr_data']['serie'], st.session_state['ocr_data']['vehiculo'] = (
                                st.session_state['ocr_data'].get('vehiculo', ''),
                                st.session_state['ocr_data'].get('serie', '')
                            )
                        st.toast("✅ Valores de Guía y Placa intercambiados correctamente.")
                        st.rerun()
                with c_fix2:
                    confirmar_forzar = st.checkbox(
                        "⚠️ Confirmar que los datos son correctos a pesar de la advertencia",
                        key="chk_forzar_placa_guia"
                    )

        # EL BOTÓN SOLO APARECE AQUÍ, SI formulario_completo es VERDADERO
        ya_generado = bool(st.session_state.get('generado', False))
        esta_generando = bool(st.session_state.get('generando_cert', False))
        bloquear_btn_generar = ya_generado or esta_generando
        es_individual = "Individual" in modalidad_gen

        if ya_generado:
            btn_label = "✅ CERTIFICADOS YA GENERADOS" if es_individual else "✅ CERTIFICADO YA GENERADO"
        elif esta_generando:
            btn_label = "⏳ GENERANDO CERTIFICADOS..." if es_individual else "⏳ GENERANDO CERTIFICADO..."
        else:
            btn_label = "GENERAR CERTIFICADOS" if es_individual else "GENERAR CERTIFICADO"

        if ya_generado:
            col_b1, col_b2 = st.columns([3, 1])
            with col_b1:
                click_generar = st.button(
                    btn_label, 
                    type="primary", 
                    disabled=True, 
                    key="btn_generar_principal", 
                    use_container_width=True
                )
            with col_b2:
                if st.button("🔄 Nuevo Certificado", help="Desbloquea para generar otro certificado o reiniciar", use_container_width=True, key="btn_desbloquear_gen"):
                    st.session_state['generado'] = False
                    st.session_state['generando_cert'] = False
                    st.session_state.pop('res_migracion_sigersol', None)
                    st.session_state.pop('word_buffer', None)
                    st.session_state.pop('nombre_generado', None)
                    st.session_state.pop('subido_drive_link', None)
                    st.session_state.pop('subido_drive_id', None)
                    st.session_state.pop('pdf_unido_link', None)
                    st.session_state.pop('pdf_unido_buffer', None)
                    st.session_state.pop('individual_exitosos', None)
                    st.session_state.pop('individual_fallidos', None)
                    st.session_state.pop('modalidad_individual_finalizada', None)
                    st.session_state['msg_generado'] = False
                    st.session_state['msg_descargado'] = False
                    st.rerun()
            st.info("🔒 **El certificado ya fue generado y sus datos fueron migrados a Sigersol.** El botón de generar se encuentra bloqueado para evitar duplicaciones en Sigersol y Drive. Para generar otro certificado o modificar datos, presiona **'🔄 Nuevo Certificado'** o **'Limpiar Sesión'**.")
        else:
            click_generar = st.button(
                btn_label, 
                type="primary", 
                disabled=bloquear_btn_generar, 
                key="btn_generar_principal", 
                use_container_width=True
            )

        if click_generar:
            if hay_error_placa_guia and not confirmar_forzar:
                st.error("⛔ Operación detenida: Corrige el número de Placa / Guía o marca la casilla de confirmación para proceder.")
                st.stop()
            st.session_state['generando_cert'] = True

            # --- ORDENAMIENTO CRONOLÓGICO PARA EL DOCUMENTO ---
            if 'fecha_origen' in v_items_df.columns:
                def fecha_a_entero_gen(fecha_str):
                    try:
                        p = str(fecha_str).strip().split('/')
                        if len(p) == 3: return int(f"{p[2]}{p[1]}{p[0]}")
                    except: pass
                    return 99999999 
                v_items_df['_llave_orden'] = v_items_df['fecha_origen'].apply(fecha_a_entero_gen)
                v_items_df = v_items_df.sort_values(by='_llave_orden', ascending=True).drop(columns=['_llave_orden']).reset_index(drop=True)

            st.session_state['df_items'] = v_items_df.copy()
            st.session_state['v_items_working'] = v_items_df.copy()
            st.session_state['editor_version'] = st.session_state.get('editor_version', 0) + 1
            if locals().get('repositorio_masivo', False):
                if st.session_state.get('repo_tipo_flujo'):
                    tipo_flujo = st.session_state['repo_tipo_flujo']
                elif st.session_state.get('repo_tipo_seleccionado'):
                    if "comercial" in str(st.session_state['repo_tipo_seleccionado']).lower():
                        tipo_flujo = "Comercialización"
                    else:
                        tipo_flujo = "Disposición Final 1"

            drive, _ = obtener_servicios()
            if drive:
                try:
                    if "Individual" in modalidad_gen:
                        guias_unicas = [g for g in v_items_df['guia_origen'].unique() if str(g).strip() not in ['None', '', 'nan']] if 'guia_origen' in v_items_df.columns else []
                        if not guias_unicas: guias_unicas = ["S/N"]
                        
                        corr_actual_int = int(str(v_corr).strip() or 1)
                        progreso = st.progress(0)
                        from datetime import datetime, timedelta
                        
                        exitosos = []
                        fallidos = []
                        
                        for idx, archivo in enumerate(guias_unicas):
                            st.toast(f"Procesando guía {archivo}...")
                            try:
                                df_filtrado = v_items_df[v_items_df['guia_origen'] == archivo] if ('guia_origen' in v_items_df.columns and archivo != "S/N") else v_items_df
                                corr_str = f"{corr_actual_int:03d}"
                                
                                if es_modelo:
                                    from src.services.google_service import obtener_plantilla_drive
                                    fh = obtener_plantilla_drive(empresa_firma, tipo_flujo, drive)
                                    doc = DocxTemplate(sanitizar_tags_docx(fh))
                                else:
                                    try:
                                        tipo_flujo_limpio = str(tipo_flujo).strip()
                                        cfg_emisor = buscar_config_empresa(empresa_firma)
                                        id_p = cfg_emisor.get('plantillas', {}).get(tipo_flujo_limpio)
                                        if not id_p and empresa_firma in PLANTILLAS:
                                            id_p = PLANTILLAS[empresa_firma].get(tipo_flujo_limpio)
                                        if not id_p and len(cfg_emisor.get('plantillas', {})) == 1:
                                            id_p = next(iter(cfg_emisor['plantillas'].values()))
                                    except KeyError:
                                        id_p = None
                                    if not id_p:
                                        st.error(f"❌ Error Crítico: No se encontró la plantilla '{tipo_flujo_limpio}' para la empresa '{empresa_firma}'.")
                                        st.stop()
                                    
                                    try:
                                        f_meta = drive.files().get(fileId=id_p, fields='mimeType', supportsAllDrives=True).execute()
                                        m_type = f_meta.get('mimeType', '')
                                        
                                        if m_type == 'application/vnd.google-apps.document':
                                            req = drive.files().export_media(fileId=id_p, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                                        else:
                                            req = drive.files().get_media(fileId=id_p, supportsAllDrives=True)
                                            
                                        fh = io.BytesIO()
                                        from googleapiclient.http import MediaIoBaseDownload
                                        dl = MediaIoBaseDownload(fh, req)
                                        done = False
                                        while not done: _, done = dl.next_chunk()
                                        doc = DocxTemplate(sanitizar_tags_docx(fh.getvalue()))
                                    except Exception as e_tpl_ind:
                                        from src.services.google_service import descargar_plantilla_en_blanco
                                        try:
                                            fh_fb = descargar_plantilla_en_blanco(empresa_firma, tipo_flujo)
                                            doc = DocxTemplate(sanitizar_tags_docx(fh_fb))
                                        except Exception:
                                            st.error(f"❌ Error al obtener la plantilla para '{empresa_firma}' (ID: {id_p}): {e_tpl_ind}")
                                            st.stop()
                                    
                                ctx = {
                                    "CORRELATIVO": corr_str, "TITULO": v_tit, "REGISTRO": emisor_reg,
                                    "CLIENTE": v_cli, "RUC_CLIENTE": v_ruc_c, "RAZON_SOCIAL_CLIENTE": v_cli,
                                    "SERVICIO_O_COMPRA": v_serv, "TIPO_DE_RESIDUO": v_res,
                                    "PUNTO_PARTIDA": str(df_filtrado['ocr_partida'].iloc[0]) if 'ocr_partida' in df_filtrado.columns and str(df_filtrado['ocr_partida'].iloc[0]).strip() else v_partida,
                                    "DIRECCION_EMPRESA": str(df_filtrado['ocr_llegada'].iloc[0]) if 'ocr_llegada' in df_filtrado.columns and str(df_filtrado['ocr_llegada'].iloc[0]).strip() else v_llegada, 
                                    "DIRECCION_LLEGADA": str(df_filtrado['ocr_llegada'].iloc[0]) if 'ocr_llegada' in df_filtrado.columns and str(df_filtrado['ocr_llegada'].iloc[0]).strip() else v_llegada,
                                    "LLEGADA": str(df_filtrado['ocr_llegada'].iloc[0]) if 'ocr_llegada' in df_filtrado.columns and str(df_filtrado['ocr_llegada'].iloc[0]).strip() else v_llegada,
                                    "EMPRESA_2": emisor_nombre,
                                    "FECHA_EMISION": (datetime.utcnow() - timedelta(hours=5)).strftime("%d/%m/%Y") if ("PROSEMBRA" in str(v_cli).upper().strip() or "VILLACURI" in str(v_cli).upper().replace(" ", "")) and ("COMERCIALIZACI" in str(tipo_flujo).upper().strip()) else v_fec_emis,
                                    "DESTINATARIO_FINAL": emisor_nombre, "EMPRESA": emisor_nombre, 
                                    "RUC_EMPRESA": emisor_ruc, "RUC": emisor_ruc,
                                    "EMISOR": emisor_nombre, "RUC_EMISOR": emisor_ruc        
                                }
                                doc.render(ctx)
                                buf_tpl = io.BytesIO()
                                doc.save(buf_tpl)
                                final_bytes = inyectar_tabla_en_docx(
                                    io.BytesIO(buf_tpl.getvalue()), 
                                    df_filtrado.to_dict('records'),
                                    columnas_config=st.session_state.get('tabla_columnas_config')
                                )
                                
                                tipo_cod = "COM" if "comercializaci" in str(tipo_flujo).strip().lower() else "SER"
                                
                                val_partida_ctx = str(df_filtrado['ocr_partida'].iloc[0]) if 'ocr_partida' in df_filtrado.columns and str(df_filtrado['ocr_partida'].iloc[0]).strip() else str(v_partida)
                                destino_raw = val_partida_ctx.split(' - ')[-1].strip()
                                import re
                                destino_raw = re.sub(r'(?i)^(Av\.|Avenida|Calle|Jr\.|Jirón|Pasaje|Carretera|Panamericana)\s+', '', destino_raw).strip()
                                destino_limpio = re.sub(r'(?i)^(Planta|Fundo|Sede|Sucursal|Predio)\s+', '', destino_raw).strip().upper()
                                destino_final = str(v_cli).strip().upper() if not destino_limpio or destino_limpio == "NAN" else destino_limpio
                                
                                if ' - ' not in str(v_partida) and len(destino_final.split()) > 1:
                                    destino_final = destino_final.split()[0]
                                    
                                nombre_archivo_final = f"CERT-{tipo_cod}-{corr_str}-{destino_final}"
                                
                                if es_modelo:
                                    from src.services.google_service import subir_modelo_a_drive
                                    link_drive = subir_modelo_a_drive(f"{nombre_archivo_final}.docx", final_bytes, drive)
                                    val_cert = "M-COM" if "Comercialización" in tipo_flujo else "M-FIN"
                                else:
                                    carpeta_exacta = cfg_emisor.get('carpetas', {}).get(tipo_flujo) or CARPETAS_DESTINO.get(empresa_firma, {}).get(tipo_flujo)
                                    if not carpeta_exacta and cfg_emisor.get('carpetas'):
                                        carpeta_exacta = next(iter(cfg_emisor['carpetas'].values()))
                                    link_drive = subir_a_drive(final_bytes, nombre_archivo_final, tipo_flujo, carpeta_id=carpeta_exacta)
                                    val_cert = "COMERCIALIZACIÓN" if "Comercialización" in tipo_flujo else "FINAL"
                                    
                                link_final = link_drive if link_drive else "Error de Permisos en Drive"
                                fecha_registro = (datetime.utcnow() - timedelta(hours=5)).strftime("%d/%m/%Y")
                                datos_log = [fecha_registro, str(v_cli).strip().upper(), destino_final, corr_str, val_cert, str(archivo).upper(), "", link_final, "", ""]
                                hoja_hist_destino = cfg_emisor.get('hoja_historial', 'Historial')
                                registrar_en_control(datos_log, hoja_destino=hoja_hist_destino)
                                
                                # --- MIGRACIÓN DIRECTA A SIGERSOL ---
                                try:
                                    from src.services.sigersol_service import migrar_datos_certificado_a_sigersol
                                    cab_sig = {
                                        "fecha": v_fec_emis,
                                        "cliente": v_cli,
                                        "ruc_cliente": v_ruc_c,
                                        "partida": str(df_filtrado['ocr_partida'].iloc[0]) if 'ocr_partida' in df_filtrado.columns and str(df_filtrado['ocr_partida'].iloc[0]).strip() else v_partida,
                                        "guia": str(archivo).upper(),
                                        "placa": v_placa,
                                        "tipo_residuo": v_res,
                                        "correlativo": corr_str
                                    }
                                    emisor_sig = {
                                        "nombre": emisor_nombre,
                                        "ruc": emisor_ruc,
                                        "registro": emisor_reg,
                                        "id_sheet_sigersol": cfg_emisor.get('id_sheet_sigersol') if 'cfg_emisor' in locals() and cfg_emisor else None
                                    }
                                    res_sig = migrar_datos_certificado_a_sigersol(df_filtrado, cab_sig, tipo_flujo, emisor_sig)
                                    if res_sig.get('ok'):
                                        st.toast(f"✅ Sigersol: {res_sig.get('filas_insertadas')} filas migradas ({res_sig.get('pestana')})")
                                except Exception as e_sig:
                                    print(f"Aviso migración Sigersol individual: {e_sig}")
                                
                                corr_actual_int += 1
                                exitosos.append(nombre_archivo_final)
                                st.session_state['metricas_exitosos'] += 1
                                
                                if link_drive:
                                    st.markdown(f"📄 **Certificado Generado:** [Ver Documento]({link_drive})")
                                    inyectar_estado_sheets_robusto(str(archivo).upper())

                                # Respaldo automático de la guía recibida en la carpeta de Drive de la empresa
                                try:
                                    arc_dict = st.session_state.get('archivos_subidos_dict', {})
                                    nom_arc_match = next((k for k in arc_dict.keys() if str(archivo).upper() in k.upper() or k.upper() in str(archivo).upper()), None)
                                    guia_bytes = arc_dict.get(nom_arc_match) if nom_arc_match else (arc_dict.get(archivo) if archivo in arc_dict else None)
                                    if guia_bytes and not es_modelo:
                                        res_g = guardar_guia_en_drive(
                                            guia_bytes, 
                                            nom_arc_match or (archivo if str(archivo).lower().endswith('.pdf') else f"{archivo}.pdf"), 
                                            empresa_nombre=empresa_firma, 
                                            tipo_guia="recibidas"
                                        )
                                        if res_g and res_g.get('ok'):
                                            st.toast(f"📥 Guía respaldada en Drive ({empresa_firma})")
                                        elif res_g and res_g.get('error'):
                                            print(f"Aviso respaldo guía Drive: {res_g.get('error')}")
                                except Exception as e_guia_up:
                                    print(f"Aviso respaldo automático guía: {e_guia_up}")
                                
                            except Exception as e:
                                fallidos.append((archivo, str(e)))
                                st.session_state['metricas_errores'] += 1
                                continue
                            finally:
                                progreso.progress((idx + 1) / len(guias_unicas))
                                
                        if exitosos:
                            st.success(f"✅ Se generaron {len(exitosos)} certificados exitosamente.")
                            st.balloons()
                            
                        if fallidos:
                            st.error(f"⚠️ Hubo {len(fallidos)} guías que fallaron durante la generación.")
                            with st.expander("Ver detalles de errores"):
                                for f_arch, f_err in fallidos:
                                    st.write(f"- **Guía {f_arch}**: {f_err}")
                                    
                        if locals().get('repositorio_masivo', False) and 'guias_repo' in st.session_state:
                            del st.session_state['guias_repo']
                            
                        st.session_state['generado'] = True
                        st.session_state['generando_cert'] = False
                        st.session_state['individual_exitosos'] = exitosos
                        st.session_state['individual_fallidos'] = fallidos
                        st.session_state['modalidad_individual_finalizada'] = True
                        st.session_state['msg_generado'] = True
                        st.balloons()
                        st.rerun()
                    if es_modelo:
                        from src.services.google_service import obtener_plantilla_drive
                        fh = obtener_plantilla_drive(empresa_firma, tipo_flujo, drive)
                        doc = DocxTemplate(sanitizar_tags_docx(fh))
                    else:
                        try:
                            tipo_flujo_limpio = str(tipo_flujo).strip()
                            cfg_emisor = buscar_config_empresa(empresa_firma)
                            id_p = cfg_emisor.get('plantillas', {}).get(tipo_flujo_limpio)
                            if not id_p and empresa_firma in PLANTILLAS:
                                id_p = PLANTILLAS[empresa_firma].get(tipo_flujo_limpio)
                            if not id_p and len(cfg_emisor.get('plantillas', {})) == 1:
                                id_p = next(iter(cfg_emisor['plantillas'].values()))
                        except KeyError:
                            id_p = None
                        if not id_p:
                            st.error(f"❌ Error Crítico: No se encontró la plantilla '{tipo_flujo_limpio}' para la empresa '{empresa_firma}'.")
                            st.stop()
                        
                        try:
                            f_meta = drive.files().get(fileId=id_p, fields='mimeType', supportsAllDrives=True).execute()
                            m_type = f_meta.get('mimeType', '')
                            
                            if m_type == 'application/vnd.google-apps.document':
                                req = drive.files().export_media(fileId=id_p, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                            else:
                                req = drive.files().get_media(fileId=id_p, supportsAllDrives=True)
                                
                            fh = io.BytesIO()
                            from googleapiclient.http import MediaIoBaseDownload
                            dl = MediaIoBaseDownload(fh, req)
                            done = False
                            while not done: _, done = dl.next_chunk()
                            doc = DocxTemplate(sanitizar_tags_docx(fh.getvalue()))
                        except Exception as e_tpl_cons:
                            from src.services.google_service import descargar_plantilla_en_blanco
                            try:
                                fh_fb = descargar_plantilla_en_blanco(empresa_firma, tipo_flujo)
                                doc = DocxTemplate(sanitizar_tags_docx(fh_fb))
                            except Exception:
                                st.error(f"❌ Error al obtener la plantilla para '{empresa_firma}' (ID: {id_p}): {e_tpl_cons}")
                                st.stop()
                        
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
                        "FECHA_EMISION": (datetime.utcnow() - timedelta(hours=5)).strftime("%d/%m/%Y") if ("PROSEMBRA" in str(v_cli).upper().strip() or "VILLACURI" in str(v_cli).upper().replace(" ", "")) and ("COMERCIALIZACI" in str(tipo_flujo).upper().strip()) else v_fec_emis,
                        "DESTINATARIO_FINAL": emisor_nombre,
                        
                        # --- VARIABLES PARA PLANTILLAS NORMALES ---
                        "EMPRESA": emisor_nombre, 
                        "RUC_EMPRESA": emisor_ruc, 
                        "RUC": emisor_ruc,
                        "REGISTRO_EMISOR": emisor_reg,
                        "REGISTRO_2": emisor_reg,
                        "F_EMISION": v_fec_emis,
                        
                        # --- VARIABLES PARA PLANTILLAS MODELO ---
                        "EMISOR": emisor_nombre,        
                        "RUC_EMISOR": emisor_ruc        
                    }
                    doc.render(ctx)
                    buf_tpl = io.BytesIO()
                    doc.save(buf_tpl)
                    
                    items_para_tabla = v_items_df.to_dict('records')
                    final_bytes = inyectar_tabla_en_docx(
                        io.BytesIO(buf_tpl.getvalue()), 
                        items_para_tabla,
                        columnas_config=st.session_state.get('tabla_columnas_config')
                    )


                    
                    # --- LÓGICA DE NOMENCLATURA ESTRICTA ---
                    tipo_cod = "COM" if "comercializaci" in str(tipo_flujo).strip().lower() else "SER"
                    
                    destino_raw = str(v_partida).split(' - ')[-1].strip()
                    import re
                    destino_raw = re.sub(r'(?i)^(Av\.|Avenida|Calle|Jr\.|Jirón|Pasaje|Carretera|Panamericana)\s+', '', destino_raw).strip()
                    destino_limpio = re.sub(r'(?i)^(Planta|Fundo|Sede|Sucursal|Predio)\s+', '', destino_raw).strip().upper()
                    
                    if not destino_limpio or destino_limpio == "NAN":
                        destino_final = str(v_cli).strip().upper()
                    else:
                        destino_final = destino_limpio
                        
                    if ' - ' not in str(v_partida) and len(destino_final.split()) > 1:
                        destino_final = destino_final.split()[0]
                    
                    nombre_archivo_final = f"CERT-{tipo_cod}-{v_corr}-{destino_final}"
                    
                    # GUARDAR EN SESIÓN PARA PERSISTENCIA
                    st.session_state.word_buffer = final_bytes
                    st.session_state.nombre_generado = nombre_archivo_final
                    st.session_state.tipo_cod = tipo_cod
                    st.session_state.v_corr = v_corr
                    st.session_state.destino_final = destino_final
                    st.session_state.generado = True
                    st.session_state['generando_cert'] = False
                    
                    st.session_state['msg_generado'] = True
                    st.session_state['msg_descargado'] = False
                    
                    # --- MIGRACIÓN AUTOMÁTICA A SIGERSOL ---
                    try:
                        from src.services.sigersol_service import migrar_datos_certificado_a_sigersol
                        cab_sig = {
                            "fecha": v_fec_emis,
                            "cliente": v_cli,
                            "ruc_cliente": v_ruc_c,
                            "partida": v_partida,
                            "guia": v_guia,
                            "placa": v_placa,
                            "tipo_residuo": v_res,
                            "correlativo": v_corr
                        }
                        emisor_sig = {
                            "nombre": emisor_nombre,
                            "ruc": emisor_ruc,
                            "registro": emisor_reg,
                            "id_sheet_sigersol": cfg_emisor.get('id_sheet_sigersol') if 'cfg_emisor' in locals() and cfg_emisor else None
                        }
                        res_sig = migrar_datos_certificado_a_sigersol(
                            df_items=v_items_df,
                            cabecera=cab_sig,
                            tipo_flujo=tipo_flujo,
                            emisor_info=emisor_sig
                        )
                        st.session_state['res_migracion_sigersol'] = res_sig
                    except Exception as e_sig:
                        st.session_state['res_migracion_sigersol'] = {"ok": False, "error": str(e_sig)}
                    
                    st.balloons()
                    st.rerun()

                except Exception as e:
                    st.session_state['generando_cert'] = False
                    st.session_state['generado'] = False
                    st.error(f"Error: {e}")
                    link_final = "Error"

    # --- 2. MOSTRAR DESCARGA Y REGISTRO (SOLO SI YA SE GENERÓ) ---
    if st.session_state.get('generado'):
        if st.session_state.get('modalidad_individual_finalizada'):
            exitosos = st.session_state.get('individual_exitosos', [])
            fallidos = st.session_state.get('individual_fallidos', [])
            if exitosos:
                st.success(f"✅ Se generaron {len(exitosos)} certificados individuales exitosamente y fueron migrados a Sigersol.")
            if fallidos:
                st.error(f"⚠️ Hubo {len(fallidos)} guías que fallaron durante la generación.")
                with st.expander("Ver detalles de errores"):
                    for f_arch, f_err in fallidos:
                        st.write(f"- **Guía {f_arch}**: {f_err}")
            st.info("💡 Cada certificado individual fue subido a Google Drive y registrado en el Historial y Sigersol.")
        else:
            if st.session_state.get('msg_generado'):
                st.success("Certificado generado exitosamente en memoria.")
                
            res_sig = st.session_state.get('res_migracion_sigersol')
            if res_sig:
                if res_sig.get('ya_existia'):
                    st.info(f"🌐 **SIGERSOL:** {res_sig.get('aviso')}")
                elif res_sig.get('ok'):
                    st.success(f"🌐 **SIGERSOL:** Se migraron {res_sig.get('filas_insertadas', 0)} registros a la pestaña **`{res_sig.get('pestana', '')}`** en el documento *3 trimestre*.")
                else:
                    st.warning(f"⚠️ **SIGERSOL (Aviso):** No se pudo migrar automáticamente a Sigersol ({res_sig.get('error', 'Desconocido')}).")
                
            nombre_safe = st.session_state.get('nombre_generado', 'Certificado')
            
            def confirmar_descarga():
                st.session_state['msg_descargado'] = True
                st.session_state['msg_generado'] = False

            if 'word_buffer' in st.session_state and st.session_state.word_buffer:
                st.download_button(
                    label="📩 Descargar Certificado", 
                    data=st.session_state.word_buffer, 
                    file_name=f"{nombre_safe}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    on_click=confirmar_descarga
                )
            
            if st.session_state.get('msg_descargado'):
                st.success("Certificado descargado exitosamente.")

            ya_subido = bool(st.session_state.get('subido_drive_link'))
            if st.button("✅ Ya Registrado y Subido a Drive" if ya_subido else "💾 Registrar y Subir a Drive", disabled=ya_subido, key="btn_subir_drive_historial"):
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
                        cfg_emisor = buscar_config_empresa(empresa_firma)
                        carpeta_exacta = cfg_emisor.get('carpetas', {}).get(tipo_flujo) or CARPETAS_DESTINO.get(empresa_firma, {}).get(tipo_flujo)
                        if not carpeta_exacta and cfg_emisor.get('carpetas'):
                            carpeta_exacta = next(iter(cfg_emisor['carpetas'].values()))
                        link_drive = subir_a_drive(buffer, nombre_safe, tipo_flujo, carpeta_id=carpeta_exacta)
                                
                    # 2. Armar la fila de datos para Sheets
                    link_final = link_drive if link_drive else "Error de Permisos en Drive"
                                
                    # --- LÓGICA DE EXTRACCIÓN PARA SHEETS (REPLICADA) ---
                    nombre_crudo = str(v_partida).split(' - ')[-1].strip()
                    import re
                    nombre_crudo = re.sub(r'(?i)^(Av\.|Avenida|Calle|Jr\.|Jirón|Pasaje|Carretera|Panamericana)\s+', '', nombre_crudo).strip()
                    nombre_limpio = re.sub(r'(?i)^(Planta|Fundo|Sede|Sucursal|Predio)\s+', '', nombre_crudo).strip()
                    
                    if ' - ' not in str(v_partida) and len(nombre_limpio.split()) > 1:
                        nombre_limpio = nombre_limpio.split()[0]
                                
                    from datetime import datetime
                    val_empresa = str(v_cli).strip().upper()
                    val_fundo = str(nombre_limpio).strip().upper()
                                
                    if es_modelo:
                        val_cert = "M-COM" if "Comercialización" in tipo_flujo else "M-FIN"
                    else:
                        val_cert = "COMERCIALIZACIÓN" if "Comercialización" in tipo_flujo else "FINAL"
                                
                    # --- 1. Lógica para capturar MÚLTIPLES guías ---
                    if not v_items_df.empty and 'guia_origen' in v_items_df.columns:
                        guias_lista = [str(g).strip().upper() for g in v_items_df['guia_origen'].unique() if str(g).strip() not in ['None', '', 'nan']]
                        val_guia_completa = ", ".join(guias_lista)
                        num_certificadas = len(guias_lista)
                    else:
                        val_guia_completa = str(v_guia).strip().upper()
                        num_certificadas = 1 if val_guia_completa else 0

                    # --- NUEVO: Inyección del Registro de Auditoría Control ---
                    modo_audio = "Modelo" if es_modelo else ("Manual" if modo_manual else ("OCR Masivo" if repositorio_masivo else "OCR PDF"))
                    from src.services.google_service import registrar_auditoria_sistema
                    registrar_auditoria_sistema(
                        st.session_state.get('usuario_email', 'Desconocido'), 
                        modo_audio, 
                        st.session_state.get('total_pdfs_leidos', 0), 
                        num_certificadas
                    )

                    # --- 2. Fecha y armado de datos para Sheets ---
                    from datetime import datetime, timedelta
                    fecha_registro = (datetime.utcnow() - timedelta(hours=5)).strftime("%d/%m/%Y")
                    v_corr_real = st.session_state.get('v_corr') or str(v_corr)
                    datos_log = [fecha_registro, val_empresa, val_fundo, v_corr_real, val_cert, val_guia_completa, "", link_final, "", ""]
                                
                    hoja_hist_destino = cfg_emisor.get('hoja_historial', 'Historial') if not es_modelo else 'Historial'
                    reg_res = registrar_en_control(datos_log, hoja_destino=hoja_hist_destino)
                    if reg_res:
                        if link_drive:
                            st.session_state['msg_generado'] = False
                            st.session_state['msg_descargado'] = False
                            st.session_state['metricas_exitosos'] += 1
                            
                            st.success("✅ ¡Operación Exitosa! Documento en Drive y base de datos actualizada.")
                            
                            # Guardar metadatos para el flujo de unión de expediente
                            st.session_state['subido_drive_link'] = link_drive
                            st.session_state['subido_drive_id'] = extraer_id_drive(link_drive)
                            st.session_state['subido_drive_nombre'] = nombre_safe
                            st.session_state['subido_fila_historial'] = reg_res if isinstance(reg_res, int) else None
                            st.session_state['subido_hoja_historial'] = hoja_hist_destino
                            carpeta_dest_safe = carpeta_exacta or CARPETAS_DESTINO.get(empresa_firma, {}).get(tipo_flujo, '12PMJ1d-CSWo64m7aNQRQj2yGHFdp9B9S') if not es_modelo else '1LUErbILxjVHnzuHkdWaeAMI4HnLg1c7E'
                            st.session_state['subido_carpeta_exacta'] = carpeta_dest_safe
                            st.session_state['subido_tipo_flujo'] = tipo_flujo
                            tipo_cod_safe = st.session_state.get('tipo_cod') or ("COM" if "comercializaci" in str(tipo_flujo).strip().lower() else "SER")
                            destino_safe = st.session_state.get('destino_final') or (val_fundo if val_fundo and val_fundo != "NAN" else str(v_cli).strip().upper())
                            st.session_state['subido_tipo_cod'] = tipo_cod_safe
                            st.session_state['subido_v_corr'] = v_corr_real
                            st.session_state['subido_destino_final'] = destino_safe
                            st.session_state['subido_guias_lista'] = guias_lista if 'guias_lista' in locals() else ([str(v_guia).strip().upper()] if str(v_guia).strip() else [])
                            
                            # --- NUEVO: Actualizar bitácora del repositorio masivo si aplica ---
                            if num_certificadas > 0:
                                for g_str in guias_lista:
                                    inyectar_estado_sheets_robusto(g_str)
                            elif val_guia_completa:
                                inyectar_estado_sheets_robusto(val_guia_completa)

                            # Respaldo automático de guías recibidas en Drive si fueron cargadas localmente
                            try:
                                subidos_dict = st.session_state.get('archivos_subidos_dict', {})
                                if subidos_dict and not es_modelo and not locals().get('repositorio_masivo', False):
                                    for g_nom, g_bytes in subidos_dict.items():
                                        res_g = guardar_guia_en_drive(
                                            g_bytes,
                                            g_nom,
                                            empresa_nombre=empresa_firma,
                                            tipo_guia="recibidas"
                                        )
                                        if res_g and res_g.get('ok'):
                                            st.toast(f"📥 Guía recibida respaldada en Drive: {g_nom}")
                                        elif res_g and res_g.get('error'):
                                            print(f"Aviso respaldo guía Drive: {res_g.get('error')}")
                            except Exception as e_guia_b:
                                print(f"Aviso respaldo guías recibidas: {e_guia_b}")
                            
                            if locals().get('repositorio_masivo', False) and 'guias_repo' in st.session_state:
                                del st.session_state['guias_repo'] # Limpiar sesión
                                
                            st.markdown(f"📄 **Certificado Generado:** [Ver Documento en Drive]({link_drive})")
                            
                            st.cache_data.clear() 
                        else:
                            st.warning("⚠️ El registro se guardó en el Excel, pero Drive rechazó el archivo.")
                    else:
                        st.error("❌ Falló la conexión con Sheets.")

        # ====================================================================
        # --- SECCIÓN: JUNTAR EXPEDIENTE EN UN SOLO PDF ---
        # ====================================================================
        if st.session_state.get('subido_drive_link'):
            link_drive_actual = st.session_state['subido_drive_link']
            st.divider()
            st.markdown("### 📑 Unión de Expediente (Certificado + Guías)")
            st.info("💡 Puedes abrir el documento con el enlace superior para revisarlo o editarlo directamente en Google Drive. Recuerda verificar que Google Docs indique **'Guardado en Drive'** antes de proceder a juntar los documentos.")
            
            opcion_juntar = st.radio(
                "¿Deseas juntar todo en un solo PDF (Certificado + Guía Remisión + Guía Transporte)?",
                ["No", "Sí"],
                horizontal=True,
                key="radio_juntar_expediente"
            )
            
            if opcion_juntar == "Sí":
                drv, sht = obtener_servicios()
                guias_a_buscar = st.session_state.get('subido_guias_lista', [])
                
                asociadas = buscar_guias_asociadas_para_unir(sht, drv, guias_a_buscar)
                
                guias_validas = [g for g in asociadas if g['valida']]
                guias_erradas = [g for g in asociadas if not g['valida']]
                
                if guias_erradas:
                    with st.expander("⚠️ Guías detectadas como ERRADAS (Omitidas del PDF)", expanded=True):
                        for ge in guias_erradas:
                            st.warning(f"🚫 **Omitida Fila {ge['fila']}**: Transporte `{ge['num_transporte']}` / Remisión `{ge['num_remision']}` — Observación: *'{ge['observacion']}'*")
                
                if not guias_validas:
                    st.warning("⚠️ No se encontraron guías válidas asociadas en 'Registro_Guias' para las guías de este certificado.")
                else:
                    st.markdown("#### Documentos a consolidar:")
                    g_val = guias_validas[0]
                    
                    c_doc1, c_doc2, c_doc3 = st.columns(3)
                    with c_doc1:
                        st.markdown(f"📄 **1. Certificado**\n\n*(Versión editada en Drive)*\n\n[🔗 Ver en Drive]({link_drive_actual})")
                    with c_doc2:
                        link_rem = g_val.get('link_remision') or "#"
                        st.markdown(f"🚛 **2. Guía Remisión**\n\n`{g_val['num_remision']}`\n\n[🔗 Ver en Drive]({link_rem})")
                    with c_doc3:
                        link_trans = g_val.get('link_transporte') or "#"
                        st.markdown(f"🚚 **3. Guía Transporte**\n\n`{g_val['num_transporte']}`\n\n[🔗 Ver en Drive]({link_trans})")
                        
                    # Nombre dinámico generado automáticamente según la operación y editable por el usuario
                    tipo_c = st.session_state.get('subido_tipo_cod') or st.session_state.get('tipo_cod', 'COM')
                    corr_c = str(st.session_state.get('subido_v_corr') or st.session_state.get('v_corr', '001')).strip()
                    dest_c = str(st.session_state.get('subido_destino_final') or st.session_state.get('destino_final', '')).strip()
                    if not dest_c or dest_c.upper() in ['NAN', 'NONE', '']:
                        dest_c = "GENERAL"
                    
                    nombre_pdf_sugerido = f"CERT-{tipo_c}-{corr_c}-REM-TRAN-{dest_c}"
                    
                    nombre_pdf_final = st.text_input(
                        "📌 Nombre del archivo PDF unificado:", 
                        value=nombre_pdf_sugerido, 
                        key="input_nombre_pdf_unido"
                    )
                    
                    ya_unido = bool(st.session_state.get('pdf_unido_link'))
                    c_btn_u1, c_btn_u2 = st.columns([3, 1])
                    with c_btn_u1:
                        btn_ejecutar_union = st.button("✅ Expediente PDF Ya Unificado" if ya_unido else "🚀 Juntar Documentos y Guardar PDF Final", type="primary", disabled=ya_unido, key="btn_ejecutar_union")
                    with c_btn_u2:
                        if ya_unido and st.button("🔄 Nueva Unión", key="btn_reset_union"):
                            for k in ['pdf_unido_link', 'pdf_unido_buffer', 'pdf_unido_nombre', 'res_sigersol_union']:
                                st.session_state.pop(k, None)
                            st.rerun()

                    if btn_ejecutar_union:
                        with st.spinner("Descargando certificado editado de Drive y guías para unificarlas..."):
                            try:
                                # 1. Descargar el Word editado de Drive
                                id_o_link_cert = st.session_state.get('subido_drive_id') or link_drive_actual
                                doc_editado_io = descargar_archivo_drive_por_id_o_nombre(drv, id_o_link_cert)
                                if not doc_editado_io:
                                    st.warning("⚠️ No se pudo descargar la versión de Drive. Se usará la copia local generada inicialmente.")
                                    doc_editado_bytes = st.session_state.word_buffer
                                else:
                                    doc_editado_bytes = doc_editado_io.getvalue()
                                    st.toast("✅ Versión actualizada del certificado descargada de Google Drive.")
                                    
                                # 2. Descargar Guía de Remisión
                                arch_rem_io = None
                                if g_val.get('archivo_remision'):
                                    arch_rem_io = descargar_archivo_drive_por_id_o_nombre(drv, g_val['archivo_remision'])
                                    
                                # 3. Descargar Guía de Transporte
                                arch_trans_io = None
                                if g_val.get('archivo_transporte'):
                                    arch_trans_io = descargar_archivo_drive_por_id_o_nombre(drv, g_val['archivo_transporte'])
                                    
                                # 4. Unir los 3 documentos en orden estricto: Certificado -> Remisión -> Transporte
                                pdf_unido_bytes = unir_tres_documentos_pdf(doc_editado_bytes, arch_rem_io, arch_trans_io)
                                
                                # 5. Subir a Drive
                                carpeta_dest = st.session_state.get('subido_carpeta_exacta')
                                t_flujo = st.session_state.get('subido_tipo_flujo', 'Comercialización')
                                link_pdf_drive = subir_pdf_a_drive(pdf_unido_bytes, nombre_pdf_final, t_flujo, carpeta_id=carpeta_dest)
                                
                                # 6. Actualizar pestaña 'Historial' Columna I ('Link pdf')
                                fila_hist = st.session_state.get('subido_fila_historial')
                                hoja_hist_act = st.session_state.get('subido_hoja_historial', 'Historial')
                                if link_pdf_drive:
                                    actualizar_link_pdf_historial(
                                        sht, 
                                        fila_hist, 
                                        link_pdf_drive, 
                                        correlativo=corr_c, 
                                        link_doc=link_drive_actual,
                                        hoja_destino=hoja_hist_act
                                    )

                                # 7. Sincronizar automáticamente cambios con SIGERSOL (por si el Word fue editado en Drive)
                                try:
                                    from src.services.sigersol_service import sincronizar_certificado_editado_a_sigersol
                                    cert_union_hist = {
                                        'correlativo': corr_c,
                                        'empresa': st.session_state.get('subido_cliente') or str(st.session_state.get('v_cli', '')),
                                        'ruc_cliente': st.session_state.get('subido_ruc_c') or str(st.session_state.get('v_ruc_c', '')),
                                        'fecha': st.session_state.get('subido_fec_emis') or str(st.session_state.get('v_fec_emis', '')),
                                        'fundo': dest_c,
                                        'tipo_cert': t_flujo,
                                        'hoja_historial': hoja_hist_act,
                                        'guias': ", ".join(st.session_state.get('subido_guias_lista', [])) or str(st.session_state.get('v_guia', ''))
                                    }
                                    res_sig_u = sincronizar_certificado_editado_a_sigersol(
                                        docx_bytes=doc_editado_bytes,
                                        pdf_bytes=pdf_unido_bytes,
                                        cert_historial=cert_union_hist
                                    )
                                    st.session_state['res_sigersol_union'] = res_sig_u
                                    if res_sig_u.get('ok'):
                                        acc_u = res_sig_u.get('accion', 'actualizaron')
                                        st.toast(f"✅ Sigersol: Se {acc_u} {res_sig_u.get('filas_afectadas', 0)} fila(s) en '{res_sig_u.get('pestana', '')}'")
                                except Exception as e_sig_u:
                                    print(f"Aviso sincronización Sigersol en unión: {e_sig_u}")
                                    st.session_state['res_sigersol_union'] = {"ok": False, "error": str(e_sig_u)}
                                
                                st.session_state['pdf_unido_buffer'] = pdf_unido_bytes
                                st.session_state['pdf_unido_link'] = link_pdf_drive
                                st.session_state['pdf_unido_nombre'] = nombre_pdf_final
                                st.success("✅ ¡Expediente PDF unificado y registrado exitosamente en Historial!")
                                st.balloons()
                            except Exception as err_union:
                                st.error(f"❌ Error al unir documentos: {err_union}")
                                
                    if st.session_state.get('pdf_unido_link'):
                        st.markdown(f"📄 **PDF Final Unificado:** [Ver en Google Drive]({st.session_state['pdf_unido_link']})")
                        st.download_button(
                            label=f"📩 Descargar {st.session_state.get('pdf_unido_nombre', 'Expediente')}.pdf",
                            data=st.session_state.get('pdf_unido_buffer', b''),
                            file_name=f"{st.session_state.get('pdf_unido_nombre', 'Expediente')}.pdf",
                            mime="application/pdf",
                            key="btn_descarga_pdf_unido"
                        )
                        res_sig_u = st.session_state.get('res_sigersol_union')
                        if res_sig_u:
                            if res_sig_u.get('ok'):
                                acc_u = res_sig_u.get('accion', 'actualizaron')
                                n_u = res_sig_u.get('filas_afectadas', 0)
                                p_u = res_sig_u.get('pestana', '')
                                if acc_u == 'sin_cambios':
                                    st.info(f"🌐 **SIGERSOL:** Los registros correspondientes a este certificado ya estaban sincronizados en **`{p_u}`**.")
                                else:
                                    st.success(f"🌐 **SIGERSOL:** Se {acc_u} exitosamente **{n_u} registro(s)** en la pestaña **`{p_u}`** de Sigersol.")
                            else:
                                st.warning(f"⚠️ **SIGERSOL (Aviso):** No se pudo sincronizar automáticamente con Sigersol: {res_sig_u.get('error', 'Error desconocido')}")
            else:
                st.info("Flujo finalizado. El certificado Word quedó registrado en Google Drive y Sheets.")

elif modulo_actual in ["🔍 Buscar y Actualizar Certificados", "🔄 Actualizar Expediente"]:
    st.title("🔍 Buscar y Actualizar Certificados / Expediente")
    st.info("💡 Localiza un certificado por Empresa Firmante y Correlativo, edítalo en Google Docs, y regenera automáticamente el PDF consolidado en la misma carpeta de Drive, actualizando el Sheet de Control y Sigersol.")

    tab_historial, tab_manual = st.tabs(["📑 Desde Historial (Automático con Drive)", "⚡ Reemplazo Directo (Carga de Archivos)"])

    drv, sht = obtener_servicios()

    with tab_historial:
        st.markdown("### 1. Búsqueda por Empresa Firmante y Número de Certificado")
        cfg_empresas_hist = obtener_configuracion_empresas()
        
        lista_empresas_firmantes = list(cfg_empresas_hist.keys())
        if not lista_empresas_firmantes:
            lista_empresas_firmantes = ["EPMI S.A.C.", "INECOVE S.A.C.", "RENOVA RECICLAJE S.A.C.", "TRANSPORT & SERVICE ALHUAY S.A.C.", "P&P INVERSIONES COMERCIALES Y AMBIENTALES S.A.C."]

        c_emp, c_busq1, c_busq2 = st.columns([2.5, 2, 1.5])
        with c_emp:
            # Paso 1: Selección obligatoria de la empresa firmante primero
            emp_busq_sel = st.selectbox("1. Empresa Firmante:", lista_empresas_firmantes, key="act_empresa_firmante_sel")
            
        cfg_emp_act = cfg_empresas_hist.get(emp_busq_sel, {})
        hoja_destino_act = cfg_emp_act.get('hoja_historial') or ("Historial_RENOVA_RECICLAJE" if "RENOVA" in emp_busq_sel.upper() else "Historial")
        carpeta_empresa_act = cfg_emp_act.get('carpetas', {}).get('Comercialización') or cfg_emp_act.get('carpetas', {}).get('Disposición Final 1') or CARPETAS_DESTINO.get(emp_busq_sel, {}).get('Comercialización')

        # Detectar si cambió la empresa seleccionada para limpiar resultados anteriores
        if st.session_state.get('act_last_empresa') != emp_busq_sel:
            st.session_state['act_last_empresa'] = emp_busq_sel
            for k_del in ['act_certificados_encontrados', 'act_pdf_final_bytes', 'act_pdf_final_link', 'act_pdf_final_nombre', 'act_nuevo_cert_preview', 'act_doc_cloud_info', 'act_res_sigersol']:
                st.session_state.pop(k_del, None)

        with c_busq1:
            # Paso 2: Búsqueda del número de certificado en la pestaña correspondiente
            corr_busqueda = st.text_input("2. N° Certificado / Correlativo:", placeholder="Ej: 001, 1, 045...", key="input_corr_actualizar")

        with c_busq2:
            st.write("")
            st.write("")
            c_btn_b1, c_btn_b2 = st.columns(2)
            with c_btn_b1:
                btn_buscar_corr = st.button("🔍 Buscar", type="primary", use_container_width=True, key="btn_buscar_historial")
            with c_btn_b2:
                btn_recientes = st.button("📋 Recientes", use_container_width=True, key="btn_ver_recientes_hist")

        st.caption(f"📊 Pestaña vinculada en Sheet de Control: **`{hoja_destino_act}`**")

        if (btn_buscar_corr and corr_busqueda) or btn_recientes or (btn_buscar_corr and not corr_busqueda):
            query_buscar = corr_busqueda.strip() if btn_buscar_corr and corr_busqueda else None
            with st.spinner(f"Consultando pestaña '{hoja_destino_act}' y resolviendo enlaces de Drive..."):
                hallados = buscar_datos_certificado_en_historial(
                    sht, 
                    correlativo=query_buscar, 
                    servicio_drive=drv, 
                    hoja_destino=hoja_destino_act,
                    carpeta_empresa_id=carpeta_empresa_act,
                    limite=30
                )
                st.session_state['act_certificados_encontrados'] = hallados
                if not hallados:
                    st.warning(f"⚠️ No se encontró ningún certificado en '{hoja_destino_act}' con el correlativo '{corr_busqueda}'.")
                else:
                    st.success(f"✅ Se encontraron {len(hallados)} registro(s) en '{hoja_destino_act}'.")

        hallados = st.session_state.get('act_certificados_encontrados', [])
        if hallados:
            if len(hallados) > 1:
                opciones_sel = [f"Fila {h['fila']} | Corr: {h['correlativo']} | {h['fecha']} | {h['empresa']} | {h['fundo']} | Guías: {h['guias']}" for h in hallados]
                sel_idx = st.selectbox("Seleccione el certificado específico a editar:", range(len(hallados)), format_func=lambda i: opciones_sel[i], key="sel_cert_multiple")
                cert_sel = hallados[sel_idx]
            else:
                cert_sel = hallados[0]

            st.divider()
            st.markdown("### 2. Información del Certificado Emitido")
            col_info1, col_info2, col_info3 = st.columns(3)
            with col_info1:
                st.markdown(f"🏢 **Empresa:** {cert_sel['empresa']}")
                st.markdown(f"🌱 **Fundo / Destino:** {cert_sel['fundo']}")
                st.markdown(f"📅 **Fecha Emisión:** {cert_sel['fecha']}")
            with col_info2:
                st.markdown(f"🏷️ **Tipo Certificado:** {cert_sel['tipo_cert']}")
                st.markdown(f"📌 **Correlativo:** `{cert_sel['correlativo']}`")
                st.markdown(f"🚛 **Guías Asociadas:** `{cert_sel['guias']}`")
            with col_info3:
                link_doc_actual = cert_sel.get('link_doc', '')
                link_pdf_actual = cert_sel.get('link_pdf', '')
                raw_doc = cert_sel.get('raw_doc', '')
                raw_pdf = cert_sel.get('raw_pdf', '')
                obs_actual = cert_sel.get('observacion', '')

                # 1. Enlace directo a Google Docs
                if link_doc_actual and str(link_doc_actual).startswith(('http://', 'https://')):
                    st.link_button("📄 Abrir y Editar en Google Docs ↗", link_doc_actual, use_container_width=True)
                elif raw_doc:
                    st.markdown(f"📄 **Word:** `{raw_doc}`")
                else:
                    st.caption("📄 Word: Sin registro en Historial")

                # 2. Enlace a PDF
                if link_pdf_actual and str(link_pdf_actual).startswith(('http://', 'https://')):
                    st.link_button("📑 Ver PDF en Google Drive ↗", link_pdf_actual, use_container_width=True)
                elif raw_pdf:
                    st.markdown(f"📑 **PDF:** `{raw_pdf}`")
                else:
                    st.caption("📑 PDF: Aún no generado (Se creará con esta acción)")

                if obs_actual:
                    st.caption(f"📝 *Obs previa:* {obs_actual}")

            st.divider()
            st.markdown("### 3. Edición del Certificado y Verificación en la Nube")
            st.info("""
            📝 **Flujo de Modificación:**
            1. Haz clic en **[📄 Abrir y Editar en Google Docs ↗]** para modificar datos (pesos, placas, correlativo, fechas, etc.).
            2. En Google Docs, confirma que aparezca el icono de nube con el check ✓ **'Guardado en Drive'**.
            3. Haz clic en **'🔄 Recargar de la Nube y Verificar Cambios'** para confirmar que Google Drive ya tiene tu nueva versión antes de regenerar el expediente.
            """)

            doc_target_actual = cert_sel.get('link_doc') or cert_sel.get('raw_doc')
            c_v1, c_v2 = st.columns([2, 1])
            with c_v1:
                btn_refrescar_cloud = st.button("🔄 Recargar de la Nube y Verificar Cambios (Limpiar Caché)", type="secondary", use_container_width=True, key="btn_check_doc_cloud")
            with c_v2:
                btn_limpiar_sesion = st.button("🧹 Limpiar Sesión y Reiniciar", use_container_width=True, key="btn_reset_modulo_actualizar")

            if btn_limpiar_sesion:
                for k in ['act_certificados_encontrados', 'act_pdf_final_bytes', 'act_pdf_final_link', 'act_pdf_final_nombre', 'act_nuevo_cert_preview', 'act_doc_cloud_info', 'act_res_sigersol']:
                    st.session_state.pop(k, None)
                st.cache_data.clear()
                st.rerun()

            if btn_refrescar_cloud:
                with st.spinner("Limpiando memoria caché y consultando el archivo Word directamente en Google Drive..."):
                    st.cache_data.clear()
                    st.session_state.pop('act_pdf_final_bytes', None)
                    st.session_state.pop('act_pdf_final_link', None)
                    st.session_state.pop('act_nuevo_cert_preview', None)
                    info_doc = obtener_info_revision_documento_drive(drv, doc_target_actual)
                    st.session_state['act_doc_cloud_info'] = info_doc
                    if info_doc:
                        st.toast("✅ Documento consultado con éxito en Google Drive.")
                    else:
                        st.warning("No se pudo obtener información del documento en Drive.")

            info_doc = st.session_state.get('act_doc_cloud_info')
            if info_doc:
                with st.container(border=True):
                    st.markdown(f"☁️ **Estado del Documento en Google Drive:** `{info_doc['name']}`")
                    col_st1, col_st2 = st.columns(2)
                    with col_st1:
                        st.markdown(f"🕒 **Última Modificación:** `{info_doc['modified_pe']}`")
                    with col_st2:
                        st.markdown(f"⏱️ **Tiempo transcurrido:** `{info_doc['tiempo_rel']}`")
                    
                    if info_doc.get('es_reciente'):
                        st.success(f"✅ **¡Documento sincronizado en Google Drive!** (Modificado {info_doc['tiempo_rel']}). Ya puedes regenerar el expediente con total seguridad.")
                    else:
                        st.warning(f"""⚠️ **Aviso de Sincronización:** En Google Drive, este archivo figura modificado el **{info_doc['modified_pe']} ({info_doc['tiempo_rel']})**.

Si acabas de editar en Google Docs y no ves tus cambios reflejados, asegúrate de haber cerrado la pestaña de Docs o esperar unos segundos a que Drive vuelque la caché, luego vuelve a presionar **[🔄 Recargar de la Nube]**.""")

                    if info_doc.get('filas_resumen'):
                        with st.expander("📋 Ver datos detectados en la tabla del Word (Primeras filas)", expanded=True):
                            for f_txt in info_doc['filas_resumen']:
                                st.code(f_txt, language=None)

            subida_local_doc = st.file_uploader(
                "📂 (Opcional) Cargar Word (.docx) o PDF corregido desde tu PC:",
                type=["docx", "pdf"],
                key="uploader_cert_corregido_hist"
            )

            st.divider()
            st.markdown("### 4. Regenerar Expediente Completo (Certificado + Guías / Petramás)")
            st.info("⚡ **Acción de regeneración integral:** Se convertirá el certificado modificado a PDF y se regenerará de forma obligatoria el expediente completo (con sus guías de remisión y transporte o constancia Petramás). El documento resultante se guardará en la misma carpeta donde reside el Docs, se registrará en el Sheet de Control y se migrará a SIGERSOL.")

            # Resolver nombre original exacto del archivo PDF
            raw_pdf_val = str(cert_sel.get('raw_pdf', '')).strip()
            nombre_pdf_orig = ""

            if raw_pdf_val and raw_pdf_val.lower().endswith('.pdf') and not raw_pdf_val.startswith(('http://', 'https://')):
                nombre_pdf_orig = raw_pdf_val
            else:
                id_para_nombre = extraer_id_drive(cert_sel.get('link_pdf')) or extraer_id_drive(raw_pdf_val)
                if drv and id_para_nombre:
                    try:
                        meta_drv = drv.files().get(fileId=id_para_nombre, fields='name', supportsAllDrives=True).execute()
                        name_drv = meta_drv.get('name', '')
                        if name_drv and name_drv.lower().endswith('.pdf'):
                            nombre_pdf_orig = name_drv
                    except Exception:
                        pass

            if not nombre_pdf_orig:
                tipo_c = str(cert_sel.get('tipo_cert', 'Comercialización')).strip()
                tipo_cod = "COM" if "comercializa" in tipo_c.lower() else "SER"
                corr_c = str(cert_sel.get('correlativo', '')).strip()
                dest_c = str(cert_sel.get('fundo', 'GENERAL')).strip()
                if not dest_c or dest_c.upper() in ['NAN', 'NONE', '']:
                    dest_c = "GENERAL"
                nombre_pdf_orig = f"CERT-{tipo_cod}-{corr_c}-REM-TRAN-{dest_c}.pdf"

            c_act_col1, c_act_col2 = st.columns([2, 1])
            with c_act_col1:
                nombre_pdf_usuario = st.text_input(
                    "📌 Nombre del archivo PDF consolidado:", 
                    value=nombre_pdf_orig, 
                    key="input_nombre_pdf_actualizado"
                )
            with c_act_col2:
                obs_adicional = st.text_input(
                    "Motivo o detalle de la modificación:", 
                    placeholder="Ej: Corrección de peso / placa", 
                    key="input_motivo_edicion"
                )

            ya_regenerado = bool(st.session_state.get('act_pdf_final_link'))
            c_reg1, c_reg2 = st.columns([3, 1])
            with c_reg1:
                btn_ejecutar_act = st.button(
                    "✅ Expediente Ya Regenerado y Actualizado" if ya_regenerado else "🚀 Regenerar Expediente Completo y Sincronizar", 
                    type="primary", 
                    disabled=ya_regenerado, 
                    use_container_width=True, 
                    key="btn_ejecutar_actualizacion_hist"
                )
            with c_reg2:
                if ya_regenerado and st.button("🔄 Actualizar Otro", key="btn_reset_act_exp", use_container_width=True):
                    for k in ['act_pdf_final_link', 'act_pdf_final_bytes', 'act_pdf_final_nombre', 'act_nuevo_cert_preview', 'act_res_sigersol']:
                        st.session_state.pop(k, None)
                    st.cache_data.clear()
                    st.rerun()

            if btn_ejecutar_act:
                doc_target = cert_sel.get('link_doc') or cert_sel.get('raw_doc')
                pdf_target = cert_sel.get('link_pdf') or cert_sel.get('raw_pdf')
                if not doc_target and not subida_local_doc:
                    st.error("❌ No se encontró el archivo Word en Google Drive ni se adjuntó un archivo local.")
                else:
                    with st.spinner("⏳ Procesando cambios: Descargando de Google Drive, convirtiendo a PDF y regenerando expediente completo..."):
                        try:
                            # 0. Limpiar memoria caché
                            st.cache_data.clear()

                            # 1. Obtener bytes del nuevo certificado en PDF y Word
                            docx_bytes_act = None
                            if subida_local_doc:
                                nombre_subido = subida_local_doc.name.lower()
                                if nombre_subido.endswith('.pdf'):
                                    nuevo_cert_pdf_bytes = subida_local_doc.getvalue()
                                    if doc_target:
                                        try:
                                            d_io = descargar_archivo_drive_por_id_o_nombre(drv, doc_target)
                                            if d_io:
                                                docx_bytes_act = d_io.getvalue()
                                        except Exception:
                                            pass
                                else:
                                    docx_bytes_act = subida_local_doc.getvalue()
                                    nuevo_cert_pdf_bytes = convertir_docx_a_pdf(docx_bytes_act)
                                st.toast("✅ Certificado cargado desde archivo local.")
                            else:
                                doc_drive_io = descargar_archivo_drive_por_id_o_nombre(drv, doc_target)
                                if not doc_drive_io:
                                    raise Exception(f"No se pudo descargar el documento Word '{doc_target}' desde Google Drive.")
                                docx_bytes_act = doc_drive_io.getvalue()
                                nuevo_cert_pdf_bytes = convertir_docx_a_pdf(docx_bytes_act)
                                st.toast("✅ Versión actualizada descargada de Google Docs y convertida a PDF.")

                            if not nuevo_cert_pdf_bytes:
                                raise Exception("Falló la conversión del certificado editado a PDF.")

                            # 2. Regenerar el expediente completo
                            pdf_expediente_final_bytes = None

                            # Caso A: Si ya existía un PDF consolidado previo en Historial o Drive:
                            # Sustituir quirúrgicamente la primera página con el nuevo certificado,
                            # conservando intactas todas las páginas posteriores (guías de remisión, transporte o sustentos Petramás).
                            if pdf_target:
                                try:
                                    pdf_unido_io = descargar_archivo_drive_por_id_o_nombre(drv, pdf_target)
                                    if pdf_unido_io:
                                        pdf_expediente_final_bytes = sustituir_certificado_en_pdf(
                                            pdf_unido_io.getvalue(),
                                            nuevo_cert_pdf_bytes,
                                            num_paginas_reemplazar=1
                                        )
                                        st.toast("✅ Expediente consolidado existente regenerado con la nueva carátula en Página 1.")
                                except Exception as e_sust:
                                    print(f"Aviso sustitución carátula en PDF existente: {e_sust}")

                            # Caso B: Si no había PDF consolidado previo, ensamblar el expediente con las guías
                            if not pdf_expediente_final_bytes:
                                guias_asoc = [g.strip().upper() for g in str(cert_sel.get('guias', '')).split(',') if g.strip()]
                                asociadas = buscar_guias_asociadas_para_unir(sht, drv, guias_asoc) if guias_asoc else []
                                guias_validas = [g for g in asociadas if g.get('valida')]

                                if guias_validas:
                                    g_val = guias_validas[0]
                                    arch_rem_io = descargar_archivo_drive_por_id_o_nombre(drv, g_val.get('archivo_remision')) if g_val.get('archivo_remision') else None
                                    arch_trans_io = descargar_archivo_drive_por_id_o_nombre(drv, g_val.get('archivo_transporte')) if g_val.get('archivo_transporte') else None
                                    pdf_expediente_final_bytes = unir_tres_documentos_pdf(docx_bytes_act or nuevo_cert_pdf_bytes, arch_rem_io, arch_trans_io)
                                    st.toast("✅ Expediente completo ensamblado (Certificado + Guía Remisión + Guía Transporte).")
                                else:
                                    # Si no se detectaron guías anexas en Drive, se toma el certificado PDF con los cambios
                                    pdf_expediente_final_bytes = nuevo_cert_pdf_bytes
                                    st.toast("✅ Certificado PDF generado con éxito.")

                            # 3. Guardar el PDF en la carpeta exacta donde se guarda el Docs
                            carpeta_docs = obtener_carpeta_padre_archivo_drive(drv, doc_target) or carpeta_empresa_act
                            file_id_pdf_existente = extraer_id_drive(cert_sel.get('link_pdf')) if str(cert_sel.get('link_pdf', '')).startswith(('http://', 'https://')) else extraer_id_drive(cert_sel.get('raw_pdf'))

                            nombre_final_pdf = nombre_pdf_usuario.strip() if nombre_pdf_usuario else nombre_pdf_orig
                            if not nombre_final_pdf.lower().endswith('.pdf'):
                                nombre_final_pdf += '.pdf'

                            nuevo_link_drive, nuevo_id_pdf = guardar_pdf_en_carpeta_docs(
                                servicio_drive=drv,
                                pdf_bytes=pdf_expediente_final_bytes,
                                nombre_archivo=nombre_final_pdf,
                                file_id_docs=doc_target,
                                file_id_pdf_existente=file_id_pdf_existente,
                                tipo_flujo=cert_sel.get('tipo_cert', 'Comercialización'),
                                carpeta_fallback_id=carpeta_docs
                            )

                            # 4. Actualizar Sheet de Control (Pestaña correspondiente, Columna I y Columna J)
                            usuario_editor = st.session_state.get('usuario_email', 'Usuario')
                            link_para_historial = nuevo_link_drive or cert_sel.get('link_pdf') or cert_sel.get('raw_pdf')
                            obs_registro = f"Expediente regenerado por {usuario_editor} con cambios en Docs y Sigersol"
                            if obs_adicional:
                                obs_registro += f" | {obs_adicional}"

                            registrar_edicion_en_historial(
                                sht,
                                cert_sel['fila'],
                                link_para_historial,
                                usuario_editor=usuario_editor,
                                obs_extra=obs_registro,
                                hoja_destino=hoja_destino_act
                            )

                            # 5. Sincronizar automáticamente cambios con SIGERSOL correspondiente
                            try:
                                from src.services.sigersol_service import sincronizar_certificado_editado_a_sigersol
                                res_sig_act = sincronizar_certificado_editado_a_sigersol(
                                    docx_bytes=docx_bytes_act,
                                    pdf_bytes=pdf_expediente_final_bytes,
                                    cert_historial=cert_sel
                                )
                                st.session_state['act_res_sigersol'] = res_sig_act
                                if res_sig_act.get('ok'):
                                    accion_txt = res_sig_act.get('accion', 'actualizaron')
                                    st.toast(f"✅ Sigersol: Se {accion_txt} {res_sig_act.get('filas_afectadas', 0)} fila(s) en '{res_sig_act.get('pestana', '')}'")
                            except Exception as e_sig_act:
                                print(f"Error sincronizando Sigersol al actualizar expediente: {e_sig_act}")
                                st.session_state['act_res_sigersol'] = {"ok": False, "error": str(e_sig_act)}

                            st.session_state['act_pdf_final_bytes'] = pdf_expediente_final_bytes
                            st.session_state['act_pdf_final_link'] = link_para_historial
                            st.session_state['act_pdf_final_nombre'] = nombre_final_pdf
                            st.session_state['act_nuevo_cert_preview'] = nuevo_cert_pdf_bytes

                            st.cache_data.clear()
                            st.success("✅ ¡Expediente regenerado exitosamente, guardado en la carpeta de Google Drive y registrado en el Sheet de Control y Sigersol!")
                            st.balloons()
                        except Exception as e_proc:
                            st.error(f"❌ Error durante la regeneración del expediente: {e_proc}")

            if st.session_state.get('act_pdf_final_link'):
                st.markdown(f"📄 **Expediente PDF Actualizado:** [Ver en Google Drive]({st.session_state['act_pdf_final_link']})")
                st.download_button(
                    label=f"📩 Descargar {st.session_state.get('act_pdf_final_nombre', 'Expediente.pdf')}",
                    data=st.session_state.get('act_pdf_final_bytes', b''),
                    file_name=st.session_state.get('act_pdf_final_nombre', 'Expediente.pdf'),
                    mime="application/pdf",
                    key="btn_descarga_act_hist"
                )
                res_sig_act = st.session_state.get('act_res_sigersol')
                if res_sig_act:
                    if res_sig_act.get('ok'):
                        accion_res = res_sig_act.get('accion', 'actualizaron')
                        n_filas = res_sig_act.get('filas_afectadas', 0)
                        p_nom = res_sig_act.get('pestana', '')
                        if accion_res == 'sin_cambios':
                            st.info(f"🌐 **SIGERSOL:** Los registros correspondientes a este certificado ya se encontraban actualizados en la pestaña **`{p_nom}`**.")
                        else:
                            st.success(f"🌐 **SIGERSOL:** Se {accion_res} exitosamente **{n_filas} fila(s)** en la pestaña **`{p_nom}`** de Sigersol.")
                    else:
                        st.warning(f"⚠️ **SIGERSOL (Aviso):** No se pudo sincronizar automáticamente con Sigersol: {res_sig_act.get('error', 'Error desconocido')}")

                st.info("💡 **Nota sobre el visor web de Google Drive:** Si haces clic en 'Ver en Google Drive' y aún observas la versión anterior, se debe a la caché del visor de Google Drive. Puedes presionar **Ctrl + Shift + R** en el visor de Drive o descargar el archivo con el botón superior para comprobar los cambios de inmediato.")
                if st.session_state.get('act_nuevo_cert_preview'):
                    with st.expander("🔍 Ver texto detectado en el nuevo certificado (Página 1)", expanded=True):
                        try:
                            from pypdf import PdfReader
                            r_pv = PdfReader(io.BytesIO(st.session_state['act_nuevo_cert_preview']))
                            st.text(r_pv.pages[0].extract_text()[:900])
                        except Exception as e_pv:
                            st.caption(f"No se pudo renderizar texto previo: {e_pv}")

    with tab_manual:
        st.markdown("### ⚡ Herramienta Rápida de Sustitución (Archivos Locales)")
        st.caption("Útil si tienes el PDF del expediente y el nuevo certificado guardados en tu computadora y no requieres consultar Drive.")
        
        c_m1, c_m2 = st.columns(2)
        with c_m1:
            pdf_expediente_file = st.file_uploader("1. Sube el PDF consolidado existente (con las guías):", type=["pdf"], key="up_pdf_expediente_manual")
        with c_m2:
            nuevo_cert_file = st.file_uploader("2. Sube el nuevo Certificado (.docx o .pdf):", type=["docx", "pdf"], key="up_nuevo_cert_manual")

        if pdf_expediente_file and nuevo_cert_file:
            nombre_descarga = pdf_expediente_file.name
            ya_swap = bool(st.session_state.get('manual_swap_bytes'))
            col_b1, col_b2 = st.columns([3, 1])
            with col_b1:
                btn_swap = st.button("✅ Expediente Ya Sustituido" if ya_swap else "⚡ Sustituir Certificado y Generar PDF", type="primary", disabled=ya_swap, key="btn_swap_manual")
            with col_b2:
                if ya_swap and st.button("🔄 Nuevo Ensamblado", key="btn_reset_swap_manual"):
                    for k in ['manual_swap_bytes', 'manual_swap_name', 'manual_res_sigersol']:
                        st.session_state.pop(k, None)
                    st.rerun()

            if btn_swap:
                with st.spinner("Ensamblando nuevo expediente..."):
                    try:
                        c_docx_bytes = None
                        if nuevo_cert_file.name.lower().endswith('.pdf'):
                            c_bytes = nuevo_cert_file.getvalue()
                        else:
                            c_docx_bytes = nuevo_cert_file.getvalue()
                            c_bytes = convertir_docx_a_pdf(c_docx_bytes)

                        pdf_res = sustituir_certificado_en_pdf(pdf_expediente_file.getvalue(), c_bytes, num_paginas_reemplazar=1)
                        st.session_state['manual_swap_bytes'] = pdf_res
                        st.session_state['manual_swap_name'] = nombre_descarga

                        # Sincronización automática con Sigersol si se cargó Word o PDF
                        try:
                            from src.services.sigersol_service import sincronizar_certificado_editado_a_sigersol
                            res_sig_man = sincronizar_certificado_editado_a_sigersol(
                                docx_bytes=c_docx_bytes,
                                pdf_bytes=c_bytes
                            )
                            st.session_state['manual_res_sigersol'] = res_sig_man
                            if res_sig_man.get('ok'):
                                st.toast(f"✅ Sigersol: {res_sig_man.get('accion')} {res_sig_man.get('filas_afectadas', 0)} fila(s)")
                        except Exception as e_sm:
                            print(f"Aviso Sigersol manual: {e_sm}")
                            st.session_state['manual_res_sigersol'] = {"ok": False, "error": str(e_sm)}

                        st.success("✅ ¡Expediente ensamblado con éxito! La página 1 fue reemplazada y las guías se mantuvieron intactas.")
                    except Exception as err_m:
                        st.error(f"❌ Error procesando archivos: {err_m}")

            if st.session_state.get('manual_swap_bytes'):
                st.download_button(
                    label=f"📩 Descargar {st.session_state.get('manual_swap_name', 'Expediente.pdf')}",
                    data=st.session_state.get('manual_swap_bytes', b''),
                    file_name=st.session_state.get('manual_swap_name', 'Expediente.pdf'),
                    mime="application/pdf",
                    key="btn_descarga_swap_manual"
                )
                res_sig_man = st.session_state.get('manual_res_sigersol')
                if res_sig_man and res_sig_man.get('ok'):
                    st.success(f"🌐 **SIGERSOL:** Se {res_sig_man.get('accion', 'actualizaron')} exitosamente **{res_sig_man.get('filas_afectadas', 0)} fila(s)** en la pestaña **`{res_sig_man.get('pestana', '')}`**.")

            if st.session_state.get('manual_swap_bytes'):
                st.download_button(
                    label=f"📩 Descargar {st.session_state.get('manual_swap_name', 'Expediente.pdf')}",
                    data=st.session_state.get('manual_swap_bytes', b''),
                    file_name=st.session_state.get('manual_swap_name', 'Expediente.pdf'),
                    mime="application/pdf",
                    key="btn_descarga_swap_manual"
                )

elif modulo_actual == "📑 Expedientes Petramás":
    with st.sidebar:
        st.info("⚙️ Controles Petramás")
        if st.button("Limpiar Sesión Petramás", key="btn_limpiar_sesion_petramas", use_container_width=True):
            llaves_petramas = [
                'inversa_datos_extraidos', 'inversa_pdf_bytes', 'inversa_resultado_desglose',
                'resultado_ensamblado', 'constancia_analizada_data', 'constancia_analizada_bytes'
            ]
            for k in llaves_petramas:
                st.session_state.pop(k, None)
            st.toast("Memoria temporal de Petramás limpiada.")
            st.rerun()

    from src.modules.petramas import render_petramas
    render_petramas()

elif modulo_actual == "🏢 Sigersol":
    with st.sidebar:
        st.info("⚙️ Controles de Sigersol")
        if st.button("Limpiar Sesión Activa", use_container_width=True):
            llaves_protegidas = ['repo', 'usuario_rol', 'usuario_email', 'metricas_exitosos', 'metricas_errores']
            for k in list(st.session_state.keys()):
                if k not in llaves_protegidas:
                    del st.session_state[k]
            st.session_state.uploader_key = st.session_state.get('uploader_key', 0) + 1
            st.rerun()

        if st.session_state.get('usuario_rol') == 'Admin':
            st.divider()
            with st.expander("🛠️ Admin Tools"):
                st.warning("Controles Elevados")
                st.markdown("### 📊 Rendimiento de Sesión")
                col1, col2 = st.columns(2)
                col1.metric(label="Certificados", value=st.session_state.get('metricas_exitosos', 0), delta="Esta sesión")
                col2.metric(label="Errores", value=st.session_state.get('metricas_errores', 0), delta="Alertas", delta_color="inverse")
                st.divider()
                if st.button("Forzar Purga de Caché GCP", use_container_width=True):
                    st.cache_data.clear()
                    st.success("Toda la Memoria RAM del entorno purgó Sheets y Drive.")
                    
    from src.modules.sigersol import render_sigersol
    render_sigersol()