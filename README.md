# Sistema de Certificados IA 🧠📄

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.30.0%2B-FF4B4B?logo=streamlit&logoColor=white)
![Vertex AI](https://img.shields.io/badge/Google_Cloud-Vertex_AI-4285F4?logo=google-cloud&logoColor=white)
![Google Workspace](https://img.shields.io/badge/Google_APIs-Drive_%26_Sheets-34A853?logo=google-drive&logoColor=white)
![Security RBAC](https://img.shields.io/badge/Security-OAuth2.0_RBAC-black?logo=security&logoColor=white)

## 🎯 Resumen Ejecutivo

El **Sistema de Certificados IA** es una aplicación web avanzada desarrollada en Python potenciada por Streamlit. Su propósito principal es eliminar la transcripción manual técnica y los errores inherentes al procesamiento burocrático de múltiples Guías de Remisión Electrónicas (GRE) en formato PDF. 

Utilizando Modelos de Lenguaje Grandes (LLMs) de la infraestructura de Google Cloud (Vertex AI con Gemini 2.0 Flash), el sistema es capaz de analizar cognitivamente documentos no estructurados. Extrae dinámicamente variables operacionales críticas (RUCs, pesos, lotes, procedencias) para compilar masivamente **Certificados de Operación** en formato Microsoft Word (`.docx`).

La plataforma incorpora una férrea arquitectura de permisos y observabilidad corporativa, orquestando accesos y auditorías en tiempo real.

## 🏗️ Arquitectura del Sistema

La solución opera mediante un orquestador híbrido que fusiona interacciones de Frontend paramétrico, potencia de inferencia y alta disponibilidad en la nube:

1. **Gatekeeper (Identity & Access Guard):** Intercepción de estado total en puerto cero. Ejecuta logueo SSO delegitimando a entidades ajenas utilizando autenticación OAuth 2.0 puramente desacoplada con protocolos REST (Bypass Stateless PKCE). Filtra perfiles contra una base de roles (RBAC) viva en Google Sheets.
2. **Ingesta Acoplada (Frontend/Cloud Input):** El usuario ingresa las guías físicas mediante cargas visuales tipo Drag-and-Drop, o directamente invoca el Repositorio Masivo donde la aplicación ejecuta transacciones de *HTTP Stream* contra Google Drive originadas de variables cruzadas.
3. **Motor Cognitivo (Vertex AI Pipeline):** Los `bytes` en memoria de los archivos PDF son procesados por Gemini Vision usando ingenierías de prompt especializadas, estructurando la data impura en un ecosistema manipulable (`JSON` -> `Pandas DataFrames`).
4. **Data Binding Dinámico:** Cruce sincrónico e inyección de diccionarios locales; vinculando las extracciones contra las bases maestras estocásticas (Clientela/Emisores) mediante validación estricta de nombres y normalización tipográfica.
5. **Procesamiento de Documentos (`docxtpl`):** El sistema escoge, manipula en la nube y descarga plantillas oficiales exclusivas. Acto seguido renderiza "tags" de contexto incrustando dinámicamente matrices de ítems transaccionales.
6. **Enrutamiento Inteligente y Auditoría:** El binario virtual terminado se inyecta en el *Data Lake* productivo, derivando en carpetas segregadas bajo lógicas algorítmicas de Taller corporativo. Finalmente, el motor central expide una señal inmutable hacia la gran matriz de **Auditoría**, calcando el accionar del usuario, volumen, modalidad y temporalidad.

## 🌟 Características Clave

- **Control de Acceso Basado en Roles (RBAC):** Flujo SSO con validación de Excel Maestro determinando privilegios administrativos de cada operario y bloqueando ingresos fantasma.
- **Auditoría Continua (Telemetry Logging):** Pistas de Auditoría exhaustivas (Audit Trails). El software registra silenciosamente horas, usuarios, modos de ingesta e índices de validación por cada botón de despacho.
- **Inteligencia Perceptiva (OCR Activo-Pasivo):** Transformación automática de PDFs genéricos hacia métricas transaccionales con control estadístico y limpiezas inteligentes de Unidades de Medida.
- **Repositorio Masivo Cascade UI:** Interfaz paralela incrustada conectada hacia flujos M2M con Google API; permite filtrar por *Empresa > Mes > Fundo* para indexar listados limpios sin latencia host-based.
- **Control de Estados (State Management Variable):** Persistencia resiliente entre refrescos de página controlando flujos de memoria (`st.session_state`), previniendo caídas críticas por re-renders de UI visuales y control en feedback atómico. 
- **Enrutamiento Paramétrico Dinámico:** Agrupación relacional y distribución algorítmica automatizada de los documentos recién concebidos a las Shared Drives pertinentes.
- **Modos Flexibles (Bypass Engine):** Modalidad "Modelo de Pruebas" libre de riesgos con trazabilidad paralela; y modalidad de inserción puramente digital/manual para transacciones excepcionales ajenas al OCR.

## 🔐 Requisitos Previos

Para ejecutar la platarforma en un entorno local o de producción en la red de Streamlit, asegúrate de cumplir con:

- Entorno de **Python 3.10+**.
- Acceso administrativo a un proyecto dentro de **Google Cloud Console** que incluya activadas las librerías nativas (`Google Drive API`, `Google Sheets API`, `Vertex AI API`).
- Matriz de Permisos Crossover: Llaves operacionales de **Cuentas de Servicio** (Para el Backend File-saving) con permisos otorgados explícitamente en modo Escritor en los Shared Drives.
- Infraestructura Secreta `secrets.toml`: Creación robusta de las variables secretas que contendrán parámetros OAuth, credenciales Web Client y el LLM API (`gcp_oauth`, `google`).

## 🚀 Instalación y Despliegue

1. **Clonar el Repositorio de la Red:**
   ```bash
   git clone <URL_DEL_REPOSITORIO>
   cd Sistema_Certificados_IA
   ```

2. **Aislar con un Entorno Virtual:**
   ```bash
   python -m venv app_env
   # En Windows: app_env\Scripts\activate
   # En Mac/Linux: source app_env/bin/activate
   ```

3. **Inyectar las Dependencias de Producción:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Inyección en Servidor Local o Despliegue CI/CD (`Streamlit Cloud`):**
   ```bash
   streamlit run app.py
   ```

## 🗺️ Mapa Fundamental del Código Central

- `app.py`: Cerebro frontal Gatekeeper. Almacena directrices UI, orquesta autenticación en capa Base, controla flujos modales y coordina auditorías de cierre.
- `src/services/google_service.py`: Motor Input/Output + Auth remoto. Proporciona túneles encriptados hacia bases RBAC, subidas de PDFs/Docs y conectores de Drive M2M.
- `src/services/vertex_service.py`: Enlace Neuronal. Conecta en backend puro a la terminal Vertex alimentando el esquema estricto (JSON output schema) para extracciones precisas.
- `src/utils/`: Bloques funcionales encapsulados dedicados a la manipulación tabular nativa (`INYECTORES DOCX`) e inteligencia de parseo de cadenas operativas.
- `src/config/settings.py`: Declarativo nativo para las IDs fijas (`Constantes`) referentes a URLs, bases de datos de seguridad y jerga taxonómica.
