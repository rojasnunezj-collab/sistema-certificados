# ====================================================================
# --- BLOQUE 0: Imports ---
# ====================================================================
import vertexai
from vertexai.generative_models import GenerativeModel, Part, GenerationConfig
from google.oauth2 import service_account
import os
import json
import streamlit as st
import warnings

# Suprimir explicitamente warnings de deprecación de Vertex AI para evitar KeyError: 'src' en Streamlit
warnings.filterwarnings("ignore", category=UserWarning, module="vertexai")

# ====================================================================
# --- BLOQUE 1: Función Principal y Variables Estáticas ---
# ====================================================================
def procesar_guia_ia_vertex(pdf_bytes):
    """
    Procesamiento Ultra-Resiliente con descubrimiento de modelos y multi-región.
    """
    PROJECT_ID = "sistemacertificados-485822"
    
    # ====================================================================
    # --- BLOQUE 2: Carga de Credenciales Inteligente (Nube o Local) ---
    # ====================================================================
    creds = None
    
    # 1. Intentar cargar desde los Secrets de Streamlit (Nube)
    if "google" in st.secrets:
        try:
            creds_info = dict(st.secrets["google"])
            creds = service_account.Credentials.from_service_account_info(creds_info)
        except Exception as e:
            st.error(f"Error cargando credenciales desde st.secrets: {e}")

    # 2. Si no hay secrets, intentar cargar desde archivo local (PC)
    if not creds:
        cred_path = next((p for p in ["secretoslocal.json", "secretos_local.json", "secretos.json"] if os.path.exists(p)), None)
        if cred_path:
            try:
                creds = service_account.Credentials.from_service_account_file(cred_path)
            except Exception as e:
                st.error(f"Error cargando archivo {cred_path}: {e}")

    # 3. Inicializar Vertex AI con las credenciales obtenidas
    try:
        vertexai.init(project=PROJECT_ID, location="us-central1", credentials=creds)
    except Exception as e:
        st.error(f"Error inicializando Vertex AI: {e}")

    # ====================================================================
    # --- BLOQUE 3: Configuración de Regiones y Modelos (Fallbacks) ---
    # ====================================================================
    # 2. Estrategia de búsqueda (Regiones y Modelos dinámicos)
    # us-central1 (estándar), us-west1 (estable), us-east4 (fallback común)
    regiones = ["us-central1", "us-west1", "us-east4", "southamerica-east1"]
    
    # Modelos detectados en este proyecto específico
    modelos_flash = ["gemini-2.0-flash-001", "gemini-2.5-flash", "gemini-1.5-flash-002", "gemini-1.5-flash-8b"]
    modelos_pro = ["gemini-2.5-pro", "gemini-3.1-pro-preview", "gemini-1.5-pro-002"]

   # ====================================================================
    # --- BLOQUE 4: Prompt del Generative Engine (Strict Extraction V4) ---
    # ====================================================================
    prompt = """
    INSTRUCCIÓN DE SISTEMA: Eres un extractor de datos OCR estricto. Tu única tarea es extraer datos del PDF adjunto y devolverlos ÚNICAMENTE en formato JSON válido. Tienes PROHIBIDO inventar datos, alucinar información o incluir texto fuera del JSON (como ```json o explicaciones).
    
    ESTRUCTURA JSON EXACTA Y REGLAS DE NEGOCIO OBLIGATORIAS:
    {
        "cliente": "Razón Social exacta del REMITENTE. Regla Estricta: NO extraer la empresa de transportes, NO extraer nombres de conductores.",
        "ruc_cliente": "Número de RUC del Remitente o Cliente Emisor.",
        "fecha": "dd/mm/yyyy", 
        "serie": "Serie-Numero completo de la guía. Ejemplo: T001-000000", 
        "vehiculo": "PLACA del vehículo. Busca en todo el documento. Obligatorio.", 
        
        # --- REGLA DE PARTIDA (ESTRICTA) ---
        # PASO 1: Extrae la dirección base de partida.
        # PASO 2: Busca exhaustivamente en "Observaciones" o resto del texto las palabras: "Fundo", "Planta", "Sede", "Sucursal" o "Predio". Identifica y extrae el nombre del fundo. Presta especial atención al fundo "LA ESPERANZA"; si aparece este nombre o algo similar, asígnalo estrictamente al campo Fundo.
        # PASO 3: Si encuentras alguna, concatena: "Dirección - [Palabra encontrada + Nombre]". (Ej: "Panamericana Sur Km 280 - Fundo La Esperanza").
        # PASO 4: Si y solo si NO hay ninguna de esas palabras en todo el PDF, concatena por defecto: "Dirección - Sede Principal".
        "punto_partida": "Valor final concatenado", 
        
        "punto_llegada": "Dirección Completa exacta de Llegada", 
        "destinatario": "Razón Social Completa del Destinatario", 
        
        "items": [
            {
                "desc": "Descripción literal del bien", 
                "cant": "Número", 
                "um": "Unidad de medida (KG, UNID, GLN)", 
                "peso": "Peso numérico explícito (o 0.00 si no existe)"
            }
        ]
    }
    """

    # ====================================================================
    # --- BLOQUE 5: Bucle Multi-Región y Ejecución de Modelos IA ---
    # ====================================================================
    pdf_part = Part.from_data(data=pdf_bytes, mime_type="application/pdf")

    # Bucle de Recuperación de Desastres
    errores_acumulados = []
    
    for region in regiones:
        try:
            vertexai.init(project=PROJECT_ID, location=region, credentials=creds)
            
            # 3. Intentar Modelos en esta región
            for m_name in modelos_flash + modelos_pro:
                try:
                    model = GenerativeModel(m_name)
                    response = model.generate_content(
                        [pdf_part, prompt],
                        generation_config=GenerationConfig(response_mime_type="application/json")
                    )
                    datos = json.loads(response.text)
                    
                    if datos.get("destinatario") or len(datos.get("vehiculo", "")) >= 3:
                        if region != "us-central1":
                            st.info(f"💡 Conectado exitosamente vía {region} con {m_name}")
                        return datos
                except Exception as e:
                    err_msg = str(e)
                    if "404" not in err_msg: # Si es otro error (ej. cuota), lo guardamos
                        errores_acumulados.append(f"{region}/{m_name}: {err_msg}")
                    continue
        except Exception as e:
            errores_acumulados.append(f"Init {region}: {str(e)}")
            continue

    # ====================================================================
    # --- BLOQUE 6: Manejo de Errores Globales y Feedback de Usuario ---
    # ====================================================================
    # Si llegamos aquí, nada funcionó
    st.error("❌ No se encontró ningún modelo de Gemini disponible en tu proyecto.")
    st.markdown(f"""
    **Causas probables:**
    1. **API no activada**: Aunque facturación esté lista, debes entrar a [Vertex AI Studio](https://console.cloud.google.com/vertex-ai/generative/multimodal/create/text?project={PROJECT_ID}) y dar clic en **'Habilitar'** si aparece.
    2. **Términos no aceptados**: Ve a [Model Garden](https://console.cloud.google.com/vertex-ai/model-garden?project={PROJECT_ID}), busca 'Gemini', haz clic en uno y verifica si pide 'Aceptar'.
    3. **Propagación**: Google puede tardar hasta 1 hora en activar IA en cuentas nuevas.
    
    **Último error detectado:** `{errores_acumulados[-1] if errores_acumulados else '404 Model Not Found'}`
    """)
    return None
