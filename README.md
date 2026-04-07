# Sistema de Certificados IA 🧠📄

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.30.0%2B-FF4B4B?logo=streamlit&logoColor=white)
![Vertex AI](https://img.shields.io/badge/Google_Cloud-Vertex_AI-4285F4?logo=google-cloud&logoColor=white)
![Google Workspace](https://img.shields.io/badge/Google_APIs-Drive_%26_Sheets-34A853?logo=google-drive&logoColor=white)

## 📌 Resumen Ejecutivo

El **Sistema de Certificados IA** es una aplicación web avanzada desarrollada en Python potenciada por Streamlit. Su propósito principal es eliminar la transcripción manual técnica y los errores inherentes al procesamiento burocrático de múltiples Guías de Remisión Electrónicas (GRE) en formato PDF. 

Utilizando Modelos de Lenguaje Grandes (LLMs) de la infraestructura de Google Cloud (Vertex AI con Gemini), el sistema es capaz de analizar cognitivamente documentos no estructurados. Extrae dinámicamente variables operacionales críticas (RUCs, pesos, lotes, procedencias) para compilar masivamente **Certificados de Operación** en formato Microsoft Word (`.docx`).

## ⚙️ Arquitectura del Sistema

La solución opera mediante un orquestador híbrido que fusiona interacciones de Frontend paramétrico, potencia de inferencia y alta disponibilidad en la nube:

1. **Ingesta Acoplada (Frontend/Cloud Input):** El usuario ingresa las guías físicas mediante cargas visuales tipo Drag-and-Drop, o directamente invoca el Repositorio Masivo donde la aplicación ejecuta transacciones de *HTTP Stream* contra Google Drive originadas de variables cruzadas en bases de Google Sheets.
2. **Motor Cognitivo (Vertex AI Pipeline):** Los `bytes` en memoria de los archivos PDF son procesados por Gemini Vision usando ingenierías de prompt especializadas, estructurando la data impura en un ecosistema manipulable (`JSON` -> `Pandas DataFrames`).
3. **Data Binding Dinámico:** Cruce sincrónico e inyección de diccionarios locales; vinculando las extracciones contra las bases maestras estocásticas (Clientela/Emisores) mediante validación estricta de nombres y normalización tipográfica.
4. **Procesamiento de Documentos (`docxtpl`):** El sistema escoge, manipula en la nube y descarga plantillas oficiales exclusivas. Acto seguido renderiza "tags" de contexto incrustando dinámicamente matrices de ítems transaccionales, creando un `.docx` listo para despacho en un buffer virtual purgado.
5. **Enrutamiento Inteligente y Registro:** El binario virtual terminado se inyecta en el *Data Lake* productivo (Workspace), derivando en carpetas segregadas bajo lógicas de Taller o Flujo Comercial ("Comercialización" vs. "Disposición Final"). El ID Cloud retornado es automáticamente embebido como hipervínculo en una bitácora final inmutable dentro de Sheets.

## ✨ Características Clave

- **Inteligencia Perceptiva (OCR Activo-Pasivo):** Transformación automática de PDFs genéricos hacia métricas transaccionales con control estadístico y limpiezas inteligentes de Unidades de Medida.
- **Repositorio Masivo Cascade UI:** Interfaz paralela incrustada conectada hacia flujos M2M con Google API; permite filtrar por *Empresa > Mes > Fundo* para indexar listados limpios sin latencia host-based (Direct memory download).
- **Control de Estados (State Management):** Persistencia resiliente entre refrescos de página controlando flujos de memoria (`st.session_state`), previniendo caídas críticas por re-renders de UI visuales y control en feedback atómico. 
- **Enrutamiento Paramétrico Dinámico:** Agrupación relacional y distribución algorítmica automatizada de los documentos recién concebidos a las Shared Drives pertinentes sin interacción humana secundaria.
- **Trazabilidad Pura:** Grabación automatizada de las métricas clave post-compilación conformando historiales auditables (`Bitácora`), provisto con inserciones temporales en batch para la repulsión de refactorizaciones accidentales (`✅ Nuevo: [Fecha]`).
- **Modos Flexibles (Bypass Engine):** Modalidad "Modelo de Pruebas" libre de riesgos con trazabilidad paralela; y modalidad de inserción puramente digital/manual para transacciones excepcionales ajenas al OCR.

## 🗂️ Requisitos Previos

Para ejecutar la platarforma en un entorno local o de producción, asegúrate de cumplir con:

- Entorno de **Python 3.10+**.
- Acceso administrativo a un proyecto dentro de **Google Cloud Console** que incluya activadas las librerías nativas (`Google Drive API`, `Google Sheets API`, `Vertex AI API`).
- Matriz de Permisos: Archivos OAUTH (`token.json`) firmados y revalidados, o llaves operacionales de cuentas de servicio con la métrica adecuada (`scopes`).
- Estructuración `secrets.toml`: Las variables secretas del LLM (ej. `GEMINI_API_KEY`) montadas de forma oculta en la raíz (`.streamlit/`).

## 🚀 Instalación y Despliegue

1. **Clonar el Repositorio de la Red:**
   ```bash
   git clone <URL_DEL_REPOSITORIO>
   cd Sistema_Certificados_IA
   ```

2. **(Opcional / Recomendado) Aislar con un Entorno Virtual:**
   ```bash
   python -m venv app_env
   # En Windows: app_env\Scripts\activate
   # En Mac/Linux: source app_env/bin/activate
   ```

3. **Inyectar las Dependencias de Producción:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Set Up Confidencial:**
   - Asegúrate de copiar tu respectivo archivo validador `token.json` o servicio de credencial GCP sobre el directorio principal.

5. **Inyección en Servidor Local:**
   ```bash
   streamlit run app.py
   ```

## 🏗️ Mapa Fundamental del Código Central

- `app.py`: Cerebro frontal. Almacena las directrices UI, control condicional e integra la lógica de orquestamiento final.
- `src/services/google_service.py`: Motor Input/Output remoto. Proporciona el túnel encriptado del código hacía APIs operacionales de Drive y Sheets.
- `src/services/vertex_service.py`: Enlace Neuronal. Conecta en backend puro a la terminal Vertex alimentando el esquema estricto (JSON output schema) para extracciones precisas.
- `src/utils/`: Bloques funcionales encapsulados dedicados a la manipulación tabular nativa (`INYECTORES DOCX`) e inteligencia de parseo de cadenas operativas.
- `src/config/settings.py`: Declarativo nativo para las IDs fijas (`Constantes`) referentes a URLs, ShareDrives referenciales, y jerga de subdominios.
