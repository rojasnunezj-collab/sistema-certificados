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
    leer_sheet_seguro,
    obtener_catalogo_guias, buscar_guias_repositorio, descargar_guias_drive, actualizar_bitacora_guias, buscar_actualizar_guia,
    buscar_guias_asociadas_para_unir, descargar_archivo_drive_por_id_o_nombre, unir_tres_documentos_pdf,
    subir_pdf_a_drive, actualizar_link_pdf_historial, extraer_id_drive,
    obtener_catalogo_servicios_por_categoria,
    convertir_docx_a_pdf, buscar_datos_certificado_en_historial,
    sobrescribir_o_subir_pdf_drive, registrar_edicion_en_historial,
    obtener_info_revision_documento_drive
)

from src.config.settings import PLANTILLAS, CARPETAS_DESTINO # <-- Añade esto
from src.utils.document_utils import inyectar_tabla_en_docx, sustituir_certificado_en_pdf
from src.utils.format_utils import (
    limpiar_monto, formato_inteligente, normalizar_fecha, 
    limpiar_descripcion, formatear_guia, obtener_fin_de_mes,
    formato_nompropio
)
from docxtpl import DocxTemplate
import urllib.parse
import requests

def mostrar_login_google():
    """Genera la URL OAuth puramente REST, sin librerías que obligen a PKCE"""
    client_id = str(st.secrets["gcp_oauth"]["client_id"]).strip()
    redirect_uri = str(st.secrets["gcp_oauth"]["redirect_uri"]).strip()
    
    # Construimos el auth URL limpio
    params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": "openid https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/userinfo.profile",
        "prompt": "consent",
        "access_type": "offline"
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
    """Atrapa el callback ?code= e invoca un POST explícito, sorteando librerías problemáticas."""
    if 'code' in st.query_params:
        try:
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
            res = requests.post(token_url, data=data).json()
            
            if "error" in res:
                st.error(f"Rechazo en Fase Token: {res.get('error_description', res['error'])}")
                st.stop()
                
            access_token = res["access_token"]
            
            # Recuperar perfil de Usuario
            usr_resp = requests.get(
                "https://www.googleapis.com/oauth2/v1/userinfo",
                headers={"Authorization": f"Bearer {access_token}"}
            ).json()
            
            st.session_state['usuario_email'] = usr_resp.get('email', '').strip().lower()
            
            # Limpiar la URL y forzar login
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"Falla crítica procesando autenticación REST: {e}")
            st.stop()

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

# 5. Adquisición Exitosas: Inyectar el Rol en el hilo continuo
st.session_state['usuario_rol'] = bd_usuarios[correo_actual].get('Rol', 'User')
st.session_state['usuario_nombre'] = bd_usuarios[correo_actual].get('Nombre', correo_actual)

with st.sidebar:
    st.info(f"👤 Conectado como: **{st.session_state['usuario_nombre']}**")
    st.divider()

if 'datos_extraidos' not in st.session_state:
    st.session_state.datos_extraidos = None

if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

modulo_actual = st.radio("Módulo", ["📄 Generador de Certificados", "🔄 Actualizar Expediente", "🏢 Sigersol"], horizontal=True, label_visibility="collapsed")

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
                
        if repositorio_masivo and st.session_state.get('repo_tipo_flujo'):
            tipo_flujo_auto = st.session_state['repo_tipo_flujo']
            idx_repo = opciones_finales.index(tipo_flujo_auto) if tipo_flujo_auto in opciones_finales else 0
            tipo_flujo = st.selectbox(
                "Tipo de Certificado", 
                options=opciones_finales, 
                index=idx_repo, 
                disabled=True, 
                help="Definido automáticamente desde la Columna J de las guías en el repositorio"
            )
            tipo_flujo = tipo_flujo_auto
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
                    d = procesar_guia_ia_vertex(arc.read())
                    if d:
                        if not grl: grl = d 
                        s, f, p = d.get('serie','S/N'), d.get('fecha',''), d.get('vehiculo','')
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
                cliente_crudo_tmp = grl.get('cliente') or grl.get('razon_social') or grl.get('empresa') or ""
                cliente_upper = str(cliente_crudo_tmp).upper().strip()
                
                if ("PROSEMBRA" in cliente_upper or "VILLACURI" in cliente_upper.replace(" ", "")) and str(tipo_flujo).upper().strip() == "COMERCIALIZACIÓN":
                    st.info("📅 REGLA ESPECIAL (HOY)")
                    f_calc = (datetime.utcnow() - timedelta(hours=5)).strftime("%d/%m/%Y")
                else:
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
                opciones_clientes = [""] + list(diccionario_clientes.keys())
                v_cli = st.selectbox("Cliente (Desde Base de Datos)", options=opciones_clientes)
                
                ruc_encontrado = diccionario_clientes.get(v_cli, "") if v_cli else ""
                v_ruc_c = st.text_input("RUC Cliente", value=ruc_encontrado)
            else:
                from src.services.google_service import obtener_clientes_desde_sheets
                diccionario_clientes = obtener_clientes_desde_sheets()
                
                v_cli = st.text_input("Cliente (Extraído)", value=cliente_ocr)
                
                # Autocompletado robusto: Si hace match con BD, usa su RUC. Si no, usa el extraído crudo de IA
                ruc_calculado = diccionario_clientes.get(str(v_cli).strip().upper(), ruc_crudo)
                v_ruc_c = st.text_input("RUC Cliente", value=ruc_calculado)
            
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
                v_items_df['guia_origen'] = v_items_df['guia_origen'].apply(lambda x: v_guia if str(x).strip() in ['', 'None', 'nan'] else x)
            if 'placa_origen' in v_items_df.columns:
                v_items_df['placa_origen'] = v_items_df['placa_origen'].apply(lambda x: v_placa if str(x).strip() in ['', 'None', 'nan'] else x)
        # --- FIX ORDENAMIENTO GLOBAL ANTES DE GENERAR ---
        if 'fecha_origen' in v_items_df.columns:
            def fecha_a_entero_gen(fecha_str):
                try:
                    p = str(fecha_str).strip().split('/')
                    if len(p) == 3: return int(f"{p[2]}{p[1]}{p[0]}")
                except: pass
                return 99999999 
            v_items_df['_llave_orden'] = v_items_df['fecha_origen'].apply(fecha_a_entero_gen)
            v_items_df = v_items_df.sort_values(by='_llave_orden', ascending=True).drop(columns=['_llave_orden']).reset_index(drop=True)

        guias_unicas_prev = [g for g in v_items_df['guia_origen'].unique() if str(g).strip() not in ['None', '', 'nan']]
        if len(guias_unicas_prev) > 1:
            modalidad_gen = st.radio("Modalidad de Generación", [
                "Agrupada (Fusionar las guías seleccionadas en 1 solo certificado)", 
                "Individual (Crear un certificado separado por cada guía seleccionada)"
            ])
        else:
            modalidad_gen = "Agrupada (Fusionar las guías seleccionadas en 1 solo certificado)"
        
        # EL BOTÓN SOLO APARECE AQUÍ, SI formulario_completo es VERDADERO
        if st.button("GENERAR CERTIFICADOS" if "Individual" in modalidad_gen else "GENERAR CERTIFICADO", type="primary"):
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
                        guias_unicas = [g for g in v_items_df['guia_origen'].unique() if str(g).strip() not in ['None', '', 'nan']]
                        if not guias_unicas: guias_unicas = ["S/N"]
                        
                        corr_actual_int = int(str(v_corr).strip() or 1)
                        progreso = st.progress(0)
                        from datetime import datetime, timedelta
                        
                        exitosos = []
                        fallidos = []
                        
                        for idx, archivo in enumerate(guias_unicas):
                            st.toast(f"Procesando guía {archivo}...")
                            try:
                                df_filtrado = v_items_df[v_items_df['guia_origen'] == archivo] if archivo != "S/N" else v_items_df
                                corr_str = f"{corr_actual_int:03d}"
                                
                                if es_modelo:
                                    from src.services.google_service import obtener_plantilla_drive
                                    fh = obtener_plantilla_drive(empresa_firma, tipo_flujo, drive)
                                    doc = DocxTemplate(fh)
                                else:
                                    try:
                                        tipo_flujo_limpio = str(tipo_flujo).strip()
                                        id_p = PLANTILLAS[empresa_firma][tipo_flujo_limpio]
                                    except KeyError:
                                        st.error(f"❌ Error Crítico: No se encontró la plantilla '{tipo_flujo_limpio}' para la empresa '{empresa_firma}' en la configuración base.")
                                        st.stop()
                                    
                                    f_meta = drive.files().get(fileId=id_p, fields='mimeType').execute()
                                    m_type = f_meta.get('mimeType', '')
                                    
                                    if m_type == 'application/vnd.google-apps.document':
                                        req = drive.files().export_media(fileId=id_p, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                                    else:
                                        req = drive.files().get_media(fileId=id_p)
                                        
                                    fh = io.BytesIO()
                                    from googleapiclient.http import MediaIoBaseDownload
                                    dl = MediaIoBaseDownload(fh, req)
                                    done = False
                                    while not done: _, done = dl.next_chunk()
                                    doc = DocxTemplate(io.BytesIO(fh.getvalue()))
                                    
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
                                final_bytes = inyectar_tabla_en_docx(io.BytesIO(buf_tpl.getvalue()), df_filtrado.to_dict('records'))
                                
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
                                    carpeta_exacta = CARPETAS_DESTINO[empresa_firma][tipo_flujo] 
                                    link_drive = subir_a_drive(final_bytes, nombre_archivo_final, tipo_flujo, carpeta_id=carpeta_exacta)
                                    val_cert = "COMERCIALIZACIÓN" if "Comercialización" in tipo_flujo else "FINAL"
                                    
                                link_final = link_drive if link_drive else "Error de Permisos en Drive"
                                fecha_registro = (datetime.utcnow() - timedelta(hours=5)).strftime("%d/%m/%Y")
                                datos_log = [fecha_registro, str(v_cli).strip().upper(), destino_final, corr_str, val_cert, str(archivo).upper(), "", link_final, "", ""]
                                registrar_en_control(datos_log)
                                
                                corr_actual_int += 1
                                exitosos.append(nombre_archivo_final)
                                st.session_state['metricas_exitosos'] += 1
                                
                                if link_drive:
                                    st.markdown(f"📄 **Certificado Generado:** [Ver Documento]({link_drive})")
                                    inyectar_estado_sheets_robusto(str(archivo).upper())
                                
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
                            
                        st.info("El lote de certificados individuales finalizó.")
                        st.stop()
                    if es_modelo:
                        from src.services.google_service import obtener_plantilla_drive
                        fh = obtener_plantilla_drive(empresa_firma, tipo_flujo, drive)
                        doc = DocxTemplate(fh)
                    else:
                        try:
                            tipo_flujo_limpio = str(tipo_flujo).strip()
                            id_p = PLANTILLAS[empresa_firma][tipo_flujo_limpio]
                        except KeyError:
                            st.error(f"❌ Error Crítico: No se encontró la plantilla '{tipo_flujo_limpio}' para la empresa '{empresa_firma}' en la configuración base.")
                            st.stop()
                        
                        f_meta = drive.files().get(fileId=id_p, fields='mimeType').execute()
                        m_type = f_meta.get('mimeType', '')
                        
                        if m_type == 'application/vnd.google-apps.document':
                            req = drive.files().export_media(fileId=id_p, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                        else:
                            req = drive.files().get_media(fileId=id_p)
                            
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
                        "FECHA_EMISION": (datetime.utcnow() - timedelta(hours=5)).strftime("%d/%m/%Y") if ("PROSEMBRA" in str(v_cli).upper().strip() or "VILLACURI" in str(v_cli).upper().replace(" ", "")) and ("COMERCIALIZACI" in str(tipo_flujo).upper().strip()) else v_fec_emis,
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
                    
                    st.session_state['msg_generado'] = True
                    st.session_state['msg_descargado'] = False
                    
                    st.balloons()
                    st.rerun()

                except Exception as e:
                    st.error(f"Error: {e}")
                    link_final = "Error"

    # --- 2. MOSTRAR DESCARGA Y REGISTRO (SOLO SI YA SE GENERÓ) ---
    if st.session_state.get('generado'):
        if st.session_state.get('msg_generado'):
            st.success("Certificado generado exitosamente en memoria.")
            
        nombre_safe = st.session_state.get('nombre_generado', 'Certificado')
        
        def confirmar_descarga():
            st.session_state['msg_descargado'] = True
            st.session_state['msg_generado'] = False

        st.download_button(
            label="📩 Descargar Certificado", 
            data=st.session_state.word_buffer, 
            file_name=f"{nombre_safe}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            on_click=confirmar_descarga
        )
        
        if st.session_state.get('msg_descargado'):
            st.success("Certificado descargado exitosamente.")

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
                    
                    # Le enviamos esa carpeta exacta a la función y usamos buffer
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
                            
                reg_res = registrar_en_control(datos_log)
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
                        carpeta_dest_safe = CARPETAS_DESTINO.get(empresa_firma, {}).get(tipo_flujo, '12PMJ1d-CSWo64m7aNQRQj2yGHFdp9B9S') if not es_modelo else '1LUErbILxjVHnzuHkdWaeAMI4HnLg1c7E'
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
                    
                    if st.button("🚀 Juntar Documentos y Guardar PDF Final", type="primary", key="btn_ejecutar_union"):
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
                                if link_pdf_drive:
                                    actualizar_link_pdf_historial(
                                        sht, 
                                        fila_hist, 
                                        link_pdf_drive, 
                                        correlativo=corr_c, 
                                        link_doc=link_drive_actual
                                    )
                                
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
            else:
                st.info("Flujo finalizado. El certificado Word quedó registrado en Google Drive y Sheets.")

elif modulo_actual == "🔄 Actualizar Expediente":
    st.title("🔄 Actualizar y Regenerar Expediente")
    st.info("💡 Modifica un certificado emitido, reemplázalo quirúrgicamente en el PDF consolidado conservando las guías originales intactas, y sincroniza los cambios en Google Drive y la hoja Historial.")

    tab_historial, tab_manual = st.tabs(["📑 Desde Historial (Automático con Drive)", "⚡ Reemplazo Directo (Carga de Archivos)"])

    drv, sht = obtener_servicios()

    with tab_historial:
        st.markdown("### 1. Localizar Certificado en Historial")
        c_busq1, c_busq2 = st.columns([3, 1])
        with c_busq1:
            corr_busqueda = st.text_input("Ingrese N° Correlativo:", placeholder="Ej: 045, 001...", key="input_corr_actualizar")
        with c_busq2:
            st.write("")
            st.write("")
            btn_buscar_corr = st.button("🔍 Buscar en Historial", type="primary", use_container_width=True, key="btn_buscar_historial")

        if btn_buscar_corr and corr_busqueda:
            with st.spinner("Buscando en Historial y resolviendo enlaces de Drive..."):
                hallados = buscar_datos_certificado_en_historial(sht, corr_busqueda, servicio_drive=drv)
                st.session_state['act_certificados_encontrados'] = hallados
                if not hallados:
                    st.warning(f"⚠️ No se encontró ningún certificado con el correlativo '{corr_busqueda}' en Historial.")
                else:
                    st.success(f"✅ Se encontraron {len(hallados)} registro(s) para el correlativo '{corr_busqueda}'.")

        hallados = st.session_state.get('act_certificados_encontrados', [])
        if hallados:
            if len(hallados) > 1:
                opciones_sel = [f"Fila {h['fila']} | {h['fecha']} | {h['empresa']} | {h['fundo']} | Guías: {h['guias']}" for h in hallados]
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

                # 1. Enlace a Word
                if link_doc_actual and str(link_doc_actual).startswith(('http://', 'https://')):
                    st.link_button("📄 Editar Word en Docs ↗", link_doc_actual, use_container_width=True)
                elif raw_doc:
                    st.markdown(f"📄 **Word:** `{raw_doc}`")
                else:
                    st.caption("📄 Word: Sin registro en Historial")

                # 2. Enlace a PDF
                if link_pdf_actual and str(link_pdf_actual).startswith(('http://', 'https://')):
                    st.link_button("📑 Ver PDF Unificado en Drive ↗", link_pdf_actual, use_container_width=True)
                elif raw_pdf:
                    st.markdown(f"📑 **PDF:** `{raw_pdf}`")
                else:
                    st.caption("📑 PDF: Sin registro en Historial")

                if obs_actual:
                    st.caption(f"📝 *Obs previa:* {obs_actual}")

            st.divider()
            st.markdown("### 3. Edición del Certificado y Verificación en la Nube")
            st.info("""
            📝 **Flujo recomendado para editar:**
            1. Haz clic en el botón superior **[📄 Editar Word en Docs ↗]** para abrir y modificar los datos (pesos, placas, correlativo, etc.) en Google Docs.
            2. En Google Docs, confirma que aparezca el icono de la nube con el check ✓ **'Guardado en Drive'**.
            3. Haz clic en el botón **'🔄 Recargar de la Nube y Verificar Cambios'** a continuación para limpiar la memoria/caché y comprobar que Google Drive ya tiene tu nueva versión antes de regenerar.
            """)

            doc_target_actual = cert_sel.get('link_doc') or cert_sel.get('raw_doc')
            c_v1, c_v2 = st.columns([2, 1])
            with c_v1:
                btn_refrescar_cloud = st.button("🔄 Recargar de la Nube y Verificar Cambios (Limpiar Caché)", type="secondary", use_container_width=True, key="btn_check_doc_cloud")
            with c_v2:
                btn_limpiar_sesion = st.button("🧹 Limpiar Sesión y Reiniciar", use_container_width=True, key="btn_reset_modulo_actualizar")

            if btn_limpiar_sesion:
                for k in ['act_certificados_encontrados', 'act_pdf_final_bytes', 'act_pdf_final_link', 'act_pdf_final_nombre', 'act_nuevo_cert_preview', 'act_doc_cloud_info']:
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
                        st.warning(f"⚠️ **Aviso de Sincronización:** En Google Drive, este archivo figura modificado el **{info_doc['modified_pe']} ({info_doc['tiempo_rel']})**.\n\nSi acabas de editar en Google Docs y no ves tus cambios reflejados, Google Docs puede tardar unos segundos en volcar el archivo a Drive. Asegúrate de presionar Enter o cerrar la pestaña de Google Docs, espera unos instantes y vuelve a presionar **[🔄 Recargar de la Nube y Verificar Cambios]**.")

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
            st.markdown("### 4. Regenerar y Publicar Expediente")

            # Resolver nombre original exacto del archivo PDF
            raw_pdf_val = str(cert_sel.get('raw_pdf', '')).strip()
            nombre_pdf_orig = ""

            if raw_pdf_val and raw_pdf_val.lower().endswith('.pdf') and not raw_pdf_val.startswith(('http://', 'https://')):
                nombre_pdf_orig = raw_pdf_val
            else:
                # Intentar leer el nombre actual directamente de Google Drive
                id_para_nombre = extraer_id_drive(cert_sel.get('link_pdf')) or extraer_id_drive(raw_pdf_val)
                if drv and id_para_nombre:
                    try:
                        meta_drv = drv.files().get(fileId=id_para_nombre, fields='name', supportsAllDrives=True).execute()
                        name_drv = meta_drv.get('name', '')
                        if name_drv and name_drv.lower().endswith('.pdf'):
                            nombre_pdf_orig = name_drv
                    except Exception:
                        pass

            # Si no se pudo obtener de Drive o Historial, reconstruir con el formato original estándar
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
                    "📌 Nombre del archivo PDF unificado:", 
                    value=nombre_pdf_orig, 
                    key="input_nombre_pdf_actualizado"
                )
            with c_act_col2:
                obs_adicional = st.text_input(
                    "Observación o motivo del cambio (opcional):", 
                    placeholder="Ej: Corrección de placa de vehículo", 
                    key="input_motivo_edicion"
                )

            if st.button("🚀 Regenerar Expediente y Actualizar Historial", type="primary", use_container_width=True, key="btn_ejecutar_actualizacion_hist"):
                doc_target = cert_sel.get('link_doc') or cert_sel.get('raw_doc')
                pdf_target = cert_sel.get('link_pdf') or cert_sel.get('raw_pdf')
                if not pdf_target:
                    st.error("❌ El registro seleccionado no tiene un enlace ni archivo de PDF unificado en Historial para reemplazar la carátula.")
                else:
                    with st.spinner("⏳ Limpiando memoria y procesando sustitución quirúrgica..."):
                        try:
                            # 0. Limpiar caché inmediatamente
                            st.cache_data.clear()

                            # 1. Obtener los bytes del nuevo certificado en PDF
                            if subida_local_doc:
                                nombre_subido = subida_local_doc.name.lower()
                                if nombre_subido.endswith('.pdf'):
                                    nuevo_cert_pdf_bytes = subida_local_doc.getvalue()
                                else:
                                    nuevo_cert_pdf_bytes = convertir_docx_a_pdf(subida_local_doc.getvalue())
                                st.toast("✅ Certificado cargado desde archivo local.")
                            else:
                                if not doc_target:
                                    raise Exception("No se encontró el archivo Word en Drive ni se subió un archivo local.")
                                doc_drive_io = descargar_archivo_drive_por_id_o_nombre(drv, doc_target)
                                if not doc_drive_io:
                                    raise Exception(f"No se pudo descargar el documento Word '{doc_target}' desde Google Drive.")
                                nuevo_cert_pdf_bytes = convertir_docx_a_pdf(doc_drive_io.getvalue())
                                st.toast("✅ Versión actualizada del Word descargada de Google Drive.")

                            if not nuevo_cert_pdf_bytes:
                                raise Exception("Falló la conversión del nuevo certificado a PDF.")

                            # 2. Descargar el PDF unificado actual de Drive
                            pdf_unido_io = descargar_archivo_drive_por_id_o_nombre(drv, pdf_target)
                            if not pdf_unido_io:
                                raise Exception(f"No se pudo descargar el PDF consolidado '{pdf_target}' desde Google Drive.")

                            pdf_unido_existente_bytes = pdf_unido_io.getvalue()

                            # 3. Sustituir quirúrgicamente la carátula conservando las guías
                            pdf_actualizado_bytes = sustituir_certificado_en_pdf(
                                pdf_unido_existente_bytes, 
                                nuevo_cert_pdf_bytes, 
                                num_paginas_reemplazar=1
                            )

                            # 4. Actualizar en Google Drive (sobreescritura in-place para conservar el mismo link público y nombre original)
                            file_id_pdf = extraer_id_drive(cert_sel.get('link_pdf')) if str(cert_sel.get('link_pdf', '')).startswith(('http://', 'https://')) else extraer_id_drive(cert_sel.get('raw_pdf'))
                            nombre_final_pdf = nombre_pdf_usuario.strip() if nombre_pdf_usuario else nombre_pdf_orig
                            if not nombre_final_pdf.lower().endswith('.pdf'):
                                nombre_final_pdf += '.pdf'
                            
                            nuevo_link_drive = sobrescribir_o_subir_pdf_drive(
                                drv, 
                                file_id_pdf, 
                                pdf_actualizado_bytes, 
                                nombre_archivo=nombre_final_pdf, 
                                tipo_flujo=cert_sel.get('tipo_cert', 'Comercialización')
                            )

                            # 5. Registrar auditoría en la pestaña 'Historial'
                            usuario_editor = st.session_state.get('usuario_email', 'Usuario')
                            link_para_historial = nuevo_link_drive or cert_sel.get('link_pdf') or cert_sel.get('raw_pdf')
                            registrar_edicion_en_historial(
                                sht, 
                                cert_sel['fila'], 
                                link_para_historial, 
                                usuario_editor=usuario_editor, 
                                obs_extra=obs_adicional
                            )

                            st.session_state['act_pdf_final_bytes'] = pdf_actualizado_bytes
                            st.session_state['act_pdf_final_link'] = link_para_historial
                            st.session_state['act_pdf_final_nombre'] = nombre_final_pdf
                            st.session_state['act_nuevo_cert_preview'] = nuevo_cert_pdf_bytes
                            
                            st.cache_data.clear()
                            st.success("✅ ¡Expediente actualizado exitosamente en Google Drive y registrado en la pestaña Historial!")
                            st.balloons()
                        except Exception as e_proc:
                            st.error(f"❌ Error durante la actualización del expediente: {e_proc}")

            if st.session_state.get('act_pdf_final_link'):
                st.markdown(f"📄 **Expediente PDF Actualizado:** [Ver en Google Drive]({st.session_state['act_pdf_final_link']})")
                st.download_button(
                    label=f"📩 Descargar {st.session_state.get('act_pdf_final_nombre', 'Expediente.pdf')}",
                    data=st.session_state.get('act_pdf_final_bytes', b''),
                    file_name=st.session_state.get('act_pdf_final_nombre', 'Expediente.pdf'),
                    mime="application/pdf",
                    key="btn_descarga_act_hist"
                )
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
            if st.button("⚡ Sustituir Certificado y Generar PDF", type="primary", key="btn_swap_manual"):
                with st.spinner("Ensamblando nuevo expediente..."):
                    try:
                        if nuevo_cert_file.name.lower().endswith('.pdf'):
                            c_bytes = nuevo_cert_file.getvalue()
                        else:
                            c_bytes = convertir_docx_a_pdf(nuevo_cert_file.getvalue())

                        pdf_res = sustituir_certificado_en_pdf(pdf_expediente_file.getvalue(), c_bytes, num_paginas_reemplazar=1)
                        st.session_state['manual_swap_bytes'] = pdf_res
                        st.session_state['manual_swap_name'] = nombre_descarga
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