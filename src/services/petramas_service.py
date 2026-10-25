# ====================================================================
# --- SERVICIO CENTRAL: GESTIÓN DE EXPEDIENTES PETRAMÁS Y SUSTENTOS ---
# ====================================================================
import os
import io
import re
import gc
import json
import tempfile
import asyncio
from datetime import datetime, timezone, timedelta
from typing import Dict, List, Any, Optional, Tuple, Union

import pandas as pd
from pypdf import PdfReader, PdfWriter
from google import genai
from google.genai import types
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaFileUpload

from src.config.settings import (
    ID_SHEET_CONTROL,
    ID_SHEET_PETRAMAS,
    SHEET_CERTIFICADOS_PETRAMAS,
    SHEET_SUSTENTOS_CAMPO,
    CARPETAS_PETRAMAS,
    EMPRESAS_EORS,
    GCP_PROJECT
)
from src.services.google_service import obtener_servicios

# Zona horaria de Perú (UTC-5)
TZ_PE = timezone(timedelta(hours=-5))

EMPRESAS_VALIDAS = list(EMPRESAS_EORS.keys())

def normalizar_empresa_eors(texto: Optional[str]) -> str:
    """Normaliza el nombre de la empresa EO-RS a una de las 3 oficiales del sistema."""
    if not texto:
        return "EPMI S.A.C."
    t = str(texto).upper()
    if "ALHUAY" in t:
        return "TRANSPORT & SERVICE ALHUAY S.A.C."
    elif "INECOVE" in t:
        return "INECOVE S.A.C."
    elif "EPMI" in t:
        return "EPMI S.A.C."
    return "EPMI S.A.C."

def obtener_hoja_historial_empresa(empresa: str) -> str:
    """Retorna la pestaña de Historial correspondiente según la EO-RS."""
    emp_norm = normalizar_empresa_eors(empresa)
    return EMPRESAS_EORS.get(emp_norm, {}).get("hoja_historial", "Historial")

def clean_json_response(text: str) -> str:
    """Limpia los bloques Markdown ```json ... ``` devueltos por el LLM."""
    if not text:
        return "{}"
    t = text.strip()
    if t.startswith("```"):
        lines = t.splitlines()
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        t = "\n".join(lines).strip()
    return t

# ====================================================================
# --- CLIENTE GENAI / VERTEX AI RESILIENTE ---
# ====================================================================
def _obtener_cliente_genai():
    """Inicializa cliente google-genai para Vertex AI o Google AI Studio con credenciales del proyecto."""
    import streamlit as st
    from google.oauth2 import service_account

    creds = None
    project_id = GCP_PROJECT or "sistemacertificados-485822"
    location = "us-central1"

    # 1. Intentar cargar desde Secrets de Streamlit
    try:
        if "google" in st.secrets:
            info = dict(st.secrets["google"])
            creds = service_account.Credentials.from_service_account_info(info)
            creds = creds.with_scopes(['https://www.googleapis.com/auth/cloud-platform'])
    except Exception:
        pass

    # 2. Cargar archivo local de credenciales
    if not creds:
        for p in ["secretoslocal.json", "secretos_local.json", "secretos.json", "credenciales_lia.json"]:
            if os.path.exists(p):
                try:
                    creds = service_account.Credentials.from_service_account_file(p)
                    creds = creds.with_scopes(['https://www.googleapis.com/auth/cloud-platform'])
                    break
                except Exception:
                    continue

    client = genai.Client(vertexai=True, project=project_id, location=location, credentials=creds)
    return client

def _invocar_gemini_multimodal(partes: List[Any], prompt: str, is_json: bool = True) -> str:
    """Ejecuta inferencia multimodal con Gemini probando modelos resilientes."""
    client = _obtener_cliente_genai()
    modelos = ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-1.5-flash-002", "gemini-2.0-flash", "gemini-1.5-pro"]
    
    config = types.GenerateContentConfig(
        temperature=0.1,
        top_p=0.1,
        max_output_tokens=8192,
        response_mime_type="application/json" if is_json else "text/plain"
    )

    last_error = None
    for model_name in modelos:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=partes + [prompt],
                config=config
            )
            if response and response.text:
                return response.text
        except Exception as e:
            last_error = e
            continue

    raise RuntimeError(f"Error invocando Gemini multimodal en modelos {modelos}: {last_error}")

# ====================================================================
# --- 1. CLASIFICADOR INTELIGENTE DE DOCUMENTOS EN PDF ---
# ====================================================================
def detectar_tipo_documento(file_source: Union[str, bytes, io.BytesIO]) -> str:
    """
    Inspecciona texto y metadatos de las primeras páginas del PDF:
    - EXPEDIENTE: Certificado emitido + Petramás + Manifiestos/Boletas (>= 3 págs).
    - PETRAMAS: Constancia oficial de disposición final de Petramás.
    - SUSTENTO: Escaneos de campo (boletas de pesaje, manifiestos MINAM, guías).
    - FACTURA: Comprobante de pago o factura electrónica.
    - GUIA: Guía de remisión electrónica o de transportista aislada.
    """
    try:
        if isinstance(file_source, (bytes, bytearray)):
            stream = io.BytesIO(file_source)
        elif isinstance(file_source, str):
            stream = open(file_source, "rb")
        else:
            stream = file_source

        reader = PdfReader(stream)
        total_pags = len(reader.pages)
        texto = ""

        # Leer hasta las primeras 5 páginas para clasificación rápida
        for i in range(min(5, total_pags)):
            extraido = reader.pages[i].extract_text()
            if extraido:
                texto += " " + extraido.upper()

        if isinstance(file_source, str):
            stream.close()

        if not texto.strip():
            return "DESCONOCIDO"

        tiene_cert_epmi = any(k in texto for k in [
            "CERTIFICADO DE RECOLECCION", "CERTIFICADO DE RECOLECCIÓN",
            "CERT-SER-", "CERT-COM-", "OPERADOR DE RESIDUOS SOLIDOS", "EO-RS"
        ])
        tiene_petramas = ("PETRAMAS" in texto or "PETRAMÁS" in texto) and any(k in texto for k in [
            "CONSTANCIA DE DISPOSICION", "CONSTANCIA DE DISPOSICIÓN", "DISPOSICION FINAL", "HUAYCOLORO"
        ])
        tiene_sustentos = any(k in texto for k in [
            "BOLETA DE PESAJE", "MANIFIESTO DE MANEJO", "MANIFIESTO", "RELLENO DE SEGURIDAD", "PESO BRUTO"
        ])

        # Si incluye Certificado de la empresa Y constancia o sustentos en un solo PDF multi-página:
        if total_pags >= 3 and tiene_cert_epmi and (tiene_petramas or tiene_sustentos):
            return "EXPEDIENTE"

        score_petramas = 0
        if "PETRAMAS" in texto or "PETRAMÁS" in texto: score_petramas += 3
        if "CONSTANCIA DE DISPOSICION" in texto or "CONSTANCIA DE DISPOSICIÓN" in texto: score_petramas += 5
        if "CONSTANCIA N" in texto or "CONSTANCIA NRO" in texto: score_petramas += 2

        score_sustento = 0
        if "BOLETA DE PESAJE" in texto or "HUAYCOLORO" in texto: score_sustento += 5
        if "MANIFIESTO DE MANEJO" in texto or "MANIFIESTO" in texto: score_sustento += 5
        if "PESO BRUTO (KG)" in texto or "PESO TARA (KG)" in texto or "PESO NETO (KG)" in texto: score_sustento += 4

        score_factura = 0
        if "FACTURA ELECTRONICA" in texto or "FACTURA ELECTRÓNICA" in texto: score_factura += 5
        if "OP. GRAVADA" in texto or "IMPORTE TOTAL" in texto: score_factura += 2

        score_guia = 0
        if "GUIA DE REMISION" in texto or "GUÍA DE REMISIÓN" in texto: score_guia += 4
        if "PUNTO DE PARTIDA" in texto and "PUNTO DE LLEGADA" in texto: score_guia += 3

        scores = {
            "PETRAMAS": score_petramas,
            "SUSTENTO": score_sustento,
            "FACTURA": score_factura,
            "GUIA": score_guia
        }
        max_tipo = max(scores, key=scores.get)
        if scores[max_tipo] >= 3:
            return max_tipo
        return "DESCONOCIDO"
    except Exception:
        return "DESCONOCIDO"

# ====================================================================
# --- 2. MÓDULO DE INGENIERÍA INVERSA (ANALIZAR EXPEDIENTES COMPLETOS) ---
# ====================================================================
PROMPT_INGENIERIA_INVERSA = """Eres un Ingeniero Senior experto en auditoría documental de residuos sólidos en Perú.
Se te proporciona un archivo PDF que corresponde a un EXPEDIENTE COMPLETO / FINAL UNIFICADO de disposición final.
Tu misión es hacer INGENIERÍA INVERSA del documento: identificar cada una de las secciones, las páginas que ocupa cada sección (números de página 1-based, enteros), y extraer con precisión los datos clave.

El expediente completo contiene hasta 5 partes en este orden:
1. Certificado emitido por la EO-RS (EPMI S.A.C., INECOVE S.A.C. o TRANSPORT & SERVICE ALHUAY S.A.C.).
2. Constancia de Disposición Final emitida por PETRAMÁS S.A.C. (Relleno Huaycoloro).
3. Manifiestos de Manejo de Residuos Peligrosos/No Peligrosos (formato oficial MINAM con casillas y firmas).
4. Boletas de Pesaje de balanza Petramás (Huaycoloro con pesos bruto, tara y neto, tickets rectangulares).
5. Guías de Remisión Remitente y de Transporte (con sellos y cargos de recepción).

Extrae los datos en formato JSON estricto con esta estructura exacta:
{
  "empresa_eors": "EPMI S.A.C." o "INECOVE S.A.C." o "TRANSPORT & SERVICE ALHUAY S.A.C.",
  "correlativo_cert": "039",
  "numero_cert_completo": "CERT-SER-039-2026",
  "fecha_emision_cert": "24/08/2026",
  "cliente_generador": "CORPORACION AGROLATINA S.A.C.",
  "sede": "FUNDO DON PEPE",
  "fecha_servicio": "21/08/2026",
  "placa": "F8C-700",
  "guia_remision": "EG07-1956",
  "guia_transporte": "EG03-00000437",
  
  "caratula_paginas": [1],

  "petramas": {
    "existe": true,
    "paginas": [2, 3],
    "numero_constancia": "308793-26",
    "fecha_disposicion": "21/08/2026",
    "peso_total_kg": 5190.0,
    "boletas": ["1889936", "1889940", "1889941", "1889942", "1889943"]
  },

  "sustentos": {
    "paginas_manifiestos": [4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
    "paginas_boletas": [14, 15, 16, 17, 18],
    "paginas_guias": [19, 20],
    "boletas_pesaje": ["1889936", "1889940", "1889941", "1889942", "1889943"]
  }
}

Reglas obligatorias:
- Las listas de páginas deben ser enteros 1-based (la primera página del PDF es 1).
- En "boletas", extrae únicamente números de boleta/ticket limpios sin texto.
- En "correlativo_cert", si dice "CERT-SER-039-2026", extrae "039".
- Si no existe alguna sección, coloca la lista vacía [] y existe: false.
- Responde ÚNICAMENTE con el objeto JSON válido, sin explicaciones.
"""

def analizar_expediente_completo_inversa(file_source: Union[str, bytes]) -> Dict[str, Any]:
    """Analiza con Gemini Multimodal un expediente completo para desglosar sus 5 secciones."""
    if isinstance(file_source, (bytes, bytearray)):
        file_bytes = bytes(file_source)
    else:
        with open(file_source, "rb") as f:
            file_bytes = f.read()

    reader = PdfReader(io.BytesIO(file_bytes))
    total_paginas = len(reader.pages)

    part = types.Part.from_bytes(data=file_bytes, mime_type="application/pdf")
    resp_text = _invocar_gemini_multimodal([part], PROMPT_INGENIERIA_INVERSA, is_json=True)
    clean_txt = clean_json_response(resp_text)
    datos = json.loads(clean_txt)

    datos["total_paginas"] = total_paginas
    datos["empresa_eors"] = normalizar_empresa_eors(datos.get("empresa_eors", ""))

    p_man = datos.get("sustentos", {}).get("paginas_manifiestos", [])
    p_bol = datos.get("sustentos", {}).get("paginas_boletas", [])
    p_gui = datos.get("sustentos", {}).get("paginas_guias", [])
    datos["sustentos"]["rangos_texto"] = f"M:{_formato_rango_paginas(p_man)} | B:{_formato_rango_paginas(p_bol)} | G:{_formato_rango_paginas(p_gui)}"
    return datos

def _formato_rango_paginas(lista_paginas: List[int]) -> str:
    """Convierte [4, 5, 6, 7, 8] en '4-8'."""
    if not lista_paginas:
        return "Ninguna"
    sorted_p = sorted(list(set(lista_paginas)))
    if len(sorted_p) == 1:
        return str(sorted_p[0])
    if sorted_p[-1] - sorted_p[0] == len(sorted_p) - 1:
        return f"{sorted_p[0]}-{sorted_p[-1]}"
    return ", ".join(str(p) for p in sorted_p)

# ====================================================================
# --- 3. LECTURA DE CONSTANCIA PETRAMÁS AISLADA ---
# ====================================================================
PROMPT_ANALISIS_PETRAMAS = """Eres un asistente experto en auditoría ambiental de residuos sólidos en Perú.
Analiza la CONSTANCIA DE DISPOSICIÓN FINAL DE RESIDUOS emitida por PETRAMÁS S.A.C. adjunta.
Extrae la información en formato JSON estricto con los siguientes campos:

{
  "empresa_eors": "EPMI S.A.C." o "INECOVE S.A.C." o "TRANSPORT & SERVICE ALHUAY S.A.C.",
  "numero_constancia": "308793-26",
  "fecha_emision": "28/08/2026",
  "generador": "CORPORACION AGROLATINA S.A.C.",
  "sede": "FUNDO DON PEPE",
  "fecha_disposicion": "21/08/2026",
  "boletas": ["1889936", "1889940", "1889941", "1889942", "1889943"],
  "peso_total_kg": 5190.0,
  "total_viajes": 5
}

Reglas:
- Fechas en formato DD/MM/YYYY.
- "boletas": lista de strings numéricos extraídos de la columna 'Nro. Boleta de Pesaje'.
- "peso_total_kg": número flotante. Si viene en TM, conviértelo a KG multiplicando por 1000.
- Devuelve ÚNICAMENTE el JSON válido.
"""

def analizar_constancia_petramas(file_source: Union[str, bytes]) -> Dict[str, Any]:
    """Extrae datos estructurados de una constancia de Petramás con Gemini."""
    if isinstance(file_source, (bytes, bytearray)):
        file_bytes = bytes(file_source)
    else:
        with open(file_source, "rb") as f:
            file_bytes = f.read()

    part = types.Part.from_bytes(data=file_bytes, mime_type="application/pdf")
    resp_text = _invocar_gemini_multimodal([part], PROMPT_ANALISIS_PETRAMAS, is_json=True)
    clean_txt = clean_json_response(resp_text)
    datos = json.loads(clean_txt)
    datos["empresa_eors"] = normalizar_empresa_eors(datos.get("empresa_eors", ""))
    return datos

# ====================================================================
# --- 4. LECTURA DE SUSTENTOS DE CAMPO AISLADOS ---
# ====================================================================
PROMPT_ANALISIS_SUSTENTO = """Eres un asistente experto en fiscalización de residuos sólidos.
Analiza el documento PDF adjunto que contiene sustentos de campo (Manifiestos, Boletas de pesaje de Huaycoloro y/o Guías).
Determina las páginas que ocupa cada documento y extrae los metadatos esenciales en formato JSON estricto:

{
  "empresa_eors": "EPMI S.A.C." o "INECOVE S.A.C." o "TRANSPORT & SERVICE ALHUAY S.A.C.",
  "generador": "CORPORACION AGROLATINA S.A.C.",
  "sede": "FUNDO DON PEPE",
  "fecha_servicio": "21/08/2026",
  "placa": "F8C-700",
  "guia_remision": "EG07-1956",
  "boletas": ["1889936", "1889940", "1889941", "1889942", "1889943"],
  "paginas_manifiestos": [1, 2, 3],
  "paginas_boletas": [4, 5],
  "paginas_guias": [6],
  "tiene_manifiestos": true,
  "tiene_boletas": true,
  "tiene_guias": true
}

Reglas:
- Páginas 1-based.
- Boletas: lista de strings numéricos de los tickets de balanza de Huaycoloro.
- Devuelve ÚNICAMENTE el JSON válido.
"""

def analizar_sustento_campo(file_source: Union[str, bytes]) -> Dict[str, Any]:
    """Extrae datos estructurados y rangos de páginas de sustentos de campo con Gemini."""
    if isinstance(file_source, (bytes, bytearray)):
        file_bytes = bytes(file_source)
    else:
        with open(file_source, "rb") as f:
            file_bytes = f.read()

    reader = PdfReader(io.BytesIO(file_bytes))
    total_pags = len(reader.pages)

    part = types.Part.from_bytes(data=file_bytes, mime_type="application/pdf")
    resp_text = _invocar_gemini_multimodal([part], PROMPT_ANALISIS_SUSTENTO, is_json=True)
    clean_txt = clean_json_response(resp_text)
    datos = json.loads(clean_txt)

    datos["total_paginas"] = total_pags
    datos["empresa_eors"] = normalizar_empresa_eors(datos.get("empresa_eors", ""))

    p_man = datos.get("paginas_manifiestos", [])
    p_bol = datos.get("paginas_boletas", [])
    p_gui = datos.get("paginas_guias", [])

    tiene_m = bool(p_man) or datos.get("tiene_manifiestos", False)
    tiene_b = bool(p_bol) or datos.get("tiene_boletas", False)
    tiene_g = bool(p_gui) or datos.get("tiene_guias", False)

    datos["tiene_manifiestos"] = tiene_m
    datos["tiene_boletas"] = tiene_b
    datos["tiene_guias"] = tiene_g

    if tiene_m and tiene_b:
        estado = "🟢 SUSTENTOS COMPLETOS"
    elif tiene_m and not tiene_b:
        estado = "🟡 PENDIENTE BOLETAS"
    elif tiene_b and not tiene_m:
        estado = "🟡 PENDIENTE MANIFIESTOS"
    else:
        estado = "🟡 PENDIENTE SUSTENTOS"

    datos["estado_documentacion"] = estado
    datos["rangos_texto"] = f"M:{_formato_rango_paginas(p_man)} | B:{_formato_rango_paginas(p_bol)} | G:{_formato_rango_paginas(p_gui)}"
    return datos

# ====================================================================
# --- 5. SUBIDA A GOOGLE DRIVE POR EMPRESA Y TIPO ---
# ====================================================================
def subir_pdf_petramas_drive(
    archivo_o_bytes: Union[str, bytes],
    tipo_carpeta: str,
    empresa: str,
    nombre_destino: str
) -> Dict[str, str]:
    """
    Sube un archivo PDF a la subcarpeta adecuada según:
    - tipo_carpeta: 'certificados', 'sustentos', 'expedientes'
    - empresa: 'EPMI S.A.C.', 'INECOVE S.A.C.', 'TRANSPORT & SERVICE ALHUAY S.A.C.'
    Retorna dict con 'id' y 'link' (webViewLink).
    """
    drv, _ = obtener_servicios()
    if not drv:
        raise ConnectionError("No se pudo obtener el servicio de Google Drive.")

    empresa_norm = normalizar_empresa_eors(empresa)
    folder_id = CARPETAS_PETRAMAS.get(tipo_carpeta, {}).get(empresa_norm)
    if not folder_id:
        folder_id = CARPETAS_PETRAMAS.get(tipo_carpeta, {}).get("EPMI S.A.C.")

    body = {
        'name': nombre_destino,
        'parents': [folder_id]
    }

    if isinstance(archivo_o_bytes, (bytes, bytearray)):
        media = MediaIoBaseUpload(io.BytesIO(archivo_o_bytes), mimetype='application/pdf', resumable=True)
    else:
        media = MediaFileUpload(archivo_o_bytes, mimetype='application/pdf', resumable=True)

    archivo_subido = drv.files().create(body=body, media_body=media, fields='id, webViewLink').execute()
    return {
        "id": archivo_subido.get('id', ''),
        "link": archivo_subido.get('webViewLink', '')
    }

# ====================================================================
# --- 6. REGISTRO EN GOOGLE SHEETS (PETRAMÁS Y SUSTENTOS) ---
# ====================================================================
def guardar_constancia_petramas_en_sheet(datos: Dict[str, Any], link_drive: str) -> Tuple[str, int]:
    """
    Registra o actualiza en la pestaña 'Certificados_Petramas' (11 columnas):
    A: Fecha Registro | B: Empresa EO-RS | C: N° Constancia | D: Fecha Disposición |
    E: Generador | F: Sede | G: N° Boletas | H: Peso Total (kg) | I: Enlace Drive PDF |
    J: Certificado Ligado | K: Estado ('REGISTRADO_DISPONIBLE' o 'VINCULADO_EXPEDIENTE')
    """
    _, sht = obtener_servicios()
    if not sht:
        raise ConnectionError("No se pudo conectar a Google Sheets.")

    num_constancia = str(datos.get("numero_constancia", "")).strip().upper()
    empresa = normalizar_empresa_eors(datos.get("empresa_eors", "EPMI S.A.C."))
    timestamp = datetime.now(TZ_PE).strftime("%d/%m/%Y %H:%M")
    boletas_str = ", ".join(str(b).strip() for b in datos.get("boletas", []) if str(b).strip())

    cert_ligado = str(datos.get("certificado_ligado", "")).strip()
    estado = "VINCULADO_EXPEDIENTE" if cert_ligado else "REGISTRADO_DISPONIBLE"
    link_formula = f'=HYPERLINK("{link_drive}"; "{num_constancia}")' if link_drive else num_constancia

    row_data = [
        timestamp,
        empresa,
        num_constancia,
        str(datos.get("fecha_disposicion", "")).strip(),
        str(datos.get("generador", "")).strip().upper(),
        str(datos.get("sede", "")).strip().upper(),
        boletas_str,
        datos.get("peso_total_kg", 0.0),
        link_formula,
        cert_ligado,
        estado
    ]

    # Leer constancias existentes para saber si se actualiza o se inserta
    r = sht.spreadsheets().values().get(
        spreadsheetId=ID_SHEET_PETRAMAS,
        range=f"'{SHEET_CERTIFICADOS_PETRAMAS}'!C:C"
    ).execute()
    col_c = [str(x[0]).strip().upper() if x else "" for x in r.get('values', [])]

    if num_constancia and num_constancia in col_c and num_constancia not in ["", "S/N"]:
        row_idx = col_c.index(num_constancia) + 1
        sht.spreadsheets().values().update(
            spreadsheetId=ID_SHEET_PETRAMAS,
            range=f"'{SHEET_CERTIFICADOS_PETRAMAS}'!A{row_idx}:K{row_idx}",
            valueInputOption="USER_ENTERED",
            body={"values": [row_data]}
        ).execute()
        return "updated", row_idx
    else:
        sht.spreadsheets().values().append(
            spreadsheetId=ID_SHEET_PETRAMAS,
            range=f"'{SHEET_CERTIFICADOS_PETRAMAS}'!A:K",
            valueInputOption="USER_ENTERED",
            body={"values": [row_data]}
        ).execute()
        row_idx = len(col_c) + 1
        return "appended", row_idx

def guardar_sustento_campo_en_sheet(datos: Dict[str, Any], link_drive: str) -> Tuple[str, int]:
    """
    Registra o actualiza en la pestaña 'Sustentos_Campo' (17 columnas):
    A: Fecha Registro | B: Empresa EO-RS | C: Generador | D: Sede | E: Fecha Servicio |
    F: Placa | G: N° Guía Remisión | H: N° Boletas Pesaje | I: Tiene Manifiestos |
    J: Tiene Boletas | K: Tiene Guías | L: Rangos / Páginas | M: Enlace Drive Escaneo |
    N: Certificado Petramás Vinculado | O: Certificado Ligado | P: Estado Documentación |
    Q: Enlace Expediente Final
    """
    _, sht = obtener_servicios()
    if not sht:
        raise ConnectionError("No se pudo conectar a Google Sheets.")

    empresa = normalizar_empresa_eors(datos.get("empresa_eors", "EPMI S.A.C."))
    generador = str(datos.get("generador", "")).strip().upper()
    sede = str(datos.get("sede", "")).strip().upper()
    fecha = str(datos.get("fecha_servicio", "")).strip()
    placa = str(datos.get("placa", "")).strip().upper()
    guia = str(datos.get("guia_remision", "")).strip().upper()

    boletas_nuevas = set(str(b).strip() for b in datos.get("boletas", []) if str(b).strip())

    # Leer filas existentes para hacer matching por Boletas o (Generador+Fecha+Placa)
    res = sht.spreadsheets().values().get(
        spreadsheetId=ID_SHEET_PETRAMAS,
        range=f"'{SHEET_SUSTENTOS_CAMPO}'!A:Q"
    ).execute()
    filas = res.get('values', [])

    row_match_idx = None
    existing_row = None

    if len(filas) > 1:
        for idx, r in enumerate(filas[1:], start=2):
            r_gen = str(r[2] if len(r) > 2 else "").strip().upper()
            r_fec = str(r[4] if len(r) > 4 else "").strip()
            r_plc = str(r[5] if len(r) > 5 else "").strip().upper()
            r_bol_raw = str(r[7] if len(r) > 7 else "").strip()
            r_boletas = set(b.strip() for b in r_bol_raw.split(",") if b.strip())

            # Coincidencia por boletas de pesaje
            if boletas_nuevas and (boletas_nuevas & r_boletas):
                row_match_idx = idx
                existing_row = r
                break
            # Coincidencia secundaria: Generador + Fecha + Placa
            elif r_gen and r_fec and r_plc and (r_gen == generador and r_fec == fecha and r_plc == placa):
                row_match_idx = idx
                existing_row = r
                break

    tiene_m = bool(datos.get("tiene_manifiestos", False))
    tiene_b = bool(datos.get("tiene_boletas", False))
    tiene_g = bool(datos.get("tiene_guias", False))
    boletas_combinadas = list(boletas_nuevas)
    cert_petramas_vinculado = str(datos.get("cert_petramas_vinculado", "")).strip()
    cert_epmi_vinculado = str(datos.get("cert_epmi_vinculado", "")).strip()
    expediente_link = str(datos.get("enlace_expediente_final", "")).strip()

    if existing_row:
        prev_m = "SÍ" in str(existing_row[8] if len(existing_row) > 8 else "").upper()
        prev_b = "SÍ" in str(existing_row[9] if len(existing_row) > 9 else "").upper()
        prev_g = "SÍ" in str(existing_row[10] if len(existing_row) > 10 else "").upper()
        tiene_m = tiene_m or prev_m
        tiene_b = tiene_b or prev_b
        tiene_g = tiene_g or prev_g

        prev_bol_raw = str(existing_row[7] if len(existing_row) > 7 else "").strip()
        prev_boletas = set(b.strip() for b in prev_bol_raw.split(",") if b.strip())
        boletas_combinadas = sorted(list(boletas_nuevas | prev_boletas))

        cert_petramas_vinculado = cert_petramas_vinculado or str(existing_row[13] if len(existing_row) > 13 else "")
        cert_epmi_vinculado = cert_epmi_vinculado or str(existing_row[14] if len(existing_row) > 14 else "")
        expediente_link = expediente_link or str(existing_row[16] if len(existing_row) > 16 else "")

    # Auto-cruzar con Constancia Petramás por boletas si no está vinculada aún
    if not cert_petramas_vinculado and boletas_combinadas:
        try:
            res_p = sht.spreadsheets().values().get(
                spreadsheetId=ID_SHEET_PETRAMAS,
                range=f"'{SHEET_CERTIFICADOS_PETRAMAS}'!A:K"
            ).execute()
            filas_p = res_p.get('values', [])
            if len(filas_p) > 1:
                for pr in filas_p[1:]:
                    pr_const = pr[2] if len(pr) > 2 else ""
                    pr_bol_raw = pr[6] if len(pr) > 6 else ""
                    pr_boletas = set(b.strip() for b in pr_bol_raw.split(",") if b.strip())
                    if set(boletas_combinadas) & pr_boletas:
                        cert_petramas_vinculado = pr_const
                        break
        except Exception:
            pass

    # Regla estricta de estados en columna P
    if expediente_link or (cert_epmi_vinculado and tiene_m and tiene_b):
        estado_final = "🟢 EXPEDIENTE COMPLETO"
    elif tiene_m and tiene_b:
        estado_final = "🟢 SUSTENTOS COMPLETOS"
    elif tiene_m and not tiene_b:
        estado_final = "🟡 PENDIENTE BOLETAS"
    elif tiene_b and not tiene_m:
        estado_final = "🟡 PENDIENTE MANIFIESTOS"
    else:
        estado_final = "🟡 PENDIENTE SUSTENTOS"

    timestamp = datetime.now(TZ_PE).strftime("%d/%m/%Y %H:%M")
    nombre_link = f"SCAN-{placa}-{fecha.replace('/', '')}" if placa and fecha else "SCAN-SUSTENTO"
    link_formula = f'=HYPERLINK("{link_drive}"; "{nombre_link}")' if link_drive else ""

    row_data = [
        timestamp,
        empresa,
        generador,
        sede,
        fecha,
        placa,
        guia,
        ", ".join(boletas_combinadas),
        "SÍ" if tiene_m else "NO",
        "SÍ" if tiene_b else "NO",
        "SÍ" if tiene_g else "NO",
        datos.get("rangos_texto", ""),
        link_formula,
        cert_petramas_vinculado,
        cert_epmi_vinculado,
        estado_final,
        expediente_link
    ]

    if row_match_idx:
        sht.spreadsheets().values().update(
            spreadsheetId=ID_SHEET_PETRAMAS,
            range=f"'{SHEET_SUSTENTOS_CAMPO}'!A{row_match_idx}:Q{row_match_idx}",
            valueInputOption="USER_ENTERED",
            body={"values": [row_data]}
        ).execute()
        return "updated", row_match_idx
    else:
        sht.spreadsheets().values().append(
            spreadsheetId=ID_SHEET_PETRAMAS,
            range=f"'{SHEET_SUSTENTOS_CAMPO}'!A:Q",
            valueInputOption="USER_ENTERED",
            body={"values": [row_data]}
        ).execute()
        row_idx = len(filas) + 1
        return "appended", row_idx

# ====================================================================
# --- 7. DESGLOSE FÍSICO Y REGISTRO COMPLETO POR INGENIERÍA INVERSA ---
# ====================================================================
def desglosar_y_registrar_expediente_inversa(
    file_source: Union[str, bytes],
    datos: Dict[str, Any],
    usuario: str = "Ingeniería Inversa"
) -> Dict[str, Any]:
    """
    Desglosa físicamente el expediente PDF en sus partes con pypdf,
    sube cada sub-PDF a su subcarpeta correspondiente en Google Drive y
    actualiza atómicamente Certificados_Petramas, Sustentos_Campo y el Historial respectivo.
    """
    temp_dir = tempfile.mkdtemp(prefix="petramas_inversa_")
    try:
        if isinstance(file_source, (bytes, bytearray)):
            file_bytes = bytes(file_source)
        else:
            with open(file_source, "rb") as f:
                file_bytes = f.read()

        reader = PdfReader(io.BytesIO(file_bytes))
        total_paginas = len(reader.pages)

        empresa = normalizar_empresa_eors(datos.get("empresa_eors", "EPMI S.A.C."))
        correlativo = str(datos.get("correlativo_cert", "000")).strip()
        generador = str(datos.get("cliente_generador", "")).strip().upper()
        sede = str(datos.get("sede", "")).strip().upper()
        fecha_servicio = str(datos.get("fecha_servicio", "")).strip()
        placa = str(datos.get("placa", "")).strip().upper()
        num_cert = str(datos.get("numero_cert_completo", f"CERT-SER-{correlativo}")).strip()

        # 1. Subir expediente final unificado a 03_Expedientes_Finales/<Empresa>/
        nombre_expediente = f"EXPEDIENTE-FINAL-{empresa[:4].strip()}-{correlativo}.pdf"
        res_exp = subir_pdf_petramas_drive(file_bytes, "expedientes", empresa, nombre_expediente)
        link_expediente = res_exp.get("link", "")

        # 2. Desglosar y subir páginas de Petramás si existen
        pet_data = datos.get("petramas", {})
        pags_pet = pet_data.get("paginas", [])
        num_constancia = str(pet_data.get("numero_constancia", "")).strip()
        link_petramas = ""

        if pet_data.get("existe") and pags_pet:
            writer_pet = PdfWriter()
            for p_idx in pags_pet:
                if 1 <= p_idx <= total_paginas:
                    writer_pet.add_page(reader.pages[p_idx - 1])
            
            buf_pet = io.BytesIO()
            writer_pet.write(buf_pet)
            pet_bytes = buf_pet.getvalue()
            buf_pet.close()

            nombre_pet = f"CONSTANCIA-PETRAMAS-{num_constancia or correlativo}.pdf"
            res_pet = subir_pdf_petramas_drive(pet_bytes, "certificados", empresa, nombre_pet)
            link_petramas = res_pet.get("link", "")

            # Guardar en Certificados_Petramas
            pet_payload = {
                "empresa_eors": empresa,
                "numero_constancia": num_constancia,
                "fecha_disposicion": pet_data.get("fecha_disposicion", fecha_servicio),
                "generador": generador,
                "sede": sede,
                "boletas": pet_data.get("boletas", []),
                "peso_total_kg": pet_data.get("peso_total_kg", 0.0),
                "certificado_ligado": num_cert
            }
            guardar_constancia_petramas_en_sheet(pet_payload, link_petramas)

        # 3. Desglosar y subir páginas de Sustentos de Campo (Manifiestos + Boletas + Guías)
        sus_data = datos.get("sustentos", {})
        pags_sus = sorted(list(set(
            sus_data.get("paginas_manifiestos", []) +
            sus_data.get("paginas_boletas", []) +
            sus_data.get("paginas_guias", [])
        )))
        link_sustento = ""

        if pags_sus:
            writer_sus = PdfWriter()
            for p_idx in pags_sus:
                if 1 <= p_idx <= total_paginas:
                    writer_sus.add_page(reader.pages[p_idx - 1])

            buf_sus = io.BytesIO()
            writer_sus.write(buf_sus)
            sus_bytes = buf_sus.getvalue()
            buf_sus.close()

            fec_clean = fecha_servicio.replace("/", "")
            nombre_sus = f"SUSTENTO-{placa}-{fec_clean or correlativo}.pdf"
            res_sus = subir_pdf_petramas_drive(sus_bytes, "sustentos", empresa, nombre_sus)
            link_sustento = res_sus.get("link", "")

            # Guardar en Sustentos_Campo
            sus_payload = {
                "empresa_eors": empresa,
                "generador": generador,
                "sede": sede,
                "fecha_servicio": fecha_servicio,
                "placa": placa,
                "guia_remision": datos.get("guia_remision", ""),
                "boletas": sus_data.get("boletas_pesaje", pet_data.get("boletas", [])),
                "tiene_manifiestos": bool(sus_data.get("paginas_manifiestos")),
                "tiene_boletas": bool(sus_data.get("paginas_boletas")),
                "tiene_guias": bool(sus_data.get("paginas_guias")),
                "rangos_texto": sus_data.get("rangos_texto", ""),
                "cert_petramas_vinculado": num_constancia,
                "cert_epmi_vinculado": num_cert,
                "enlace_expediente_final": link_expediente
            }
            guardar_sustento_campo_en_sheet(sus_payload, link_sustento)

        # 4. Actualizar o Registrar en la pestaña de Historial de la EO-RS
        hoja_historial = obtener_hoja_historial_empresa(empresa)
        _, sht = obtener_servicios()
        
        # Buscar correlativo en Columna D de la hoja Historial
        r_hist = sht.spreadsheets().values().get(
            spreadsheetId=ID_SHEET_CONTROL,
            range=f"'{hoja_historial}'!D:D"
        ).execute()
        col_corr = [str(x[0]).strip() if x else "" for x in r_hist.get('values', [])]
        corr_clean = correlativo.lstrip('0')

        matched_row = None
        for i, c in enumerate(col_corr[1:], start=2):
            if c.lstrip('0') == corr_clean and corr_clean:
                matched_row = i
                break

        obs_texto = f"Expediente Completo (Petramás - {total_paginas} págs) - {usuario}"
        val_link = f'=HYPERLINK("{link_expediente}"; "{nombre_expediente}")'

        if matched_row:
            sht.spreadsheets().values().update(
                spreadsheetId=ID_SHEET_CONTROL,
                range=f"'{hoja_historial}'!I{matched_row}:J{matched_row}",
                valueInputOption="USER_ENTERED",
                body={"values": [[val_link, obs_texto]]}
            ).execute()
        else:
            fecha_emis = datos.get("fecha_emision_cert", datetime.now(TZ_PE).strftime("%d/%m/%Y"))
            guia_completa = datos.get("guia_remision", "")
            if datos.get("guia_transporte"):
                guia_completa = f"{guia_completa}, {datos.get('guia_transporte')}" if guia_completa else datos.get("guia_transporte")

            mes_nombre = ""
            try:
                partes_fec = fecha_emis.split("/")
                if len(partes_fec) >= 2:
                    mes_num = int(partes_fec[1])
                    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre", "noviembre", "diciembre"]
                    if 1 <= mes_num <= 12:
                        mes_nombre = meses[mes_num - 1]
            except Exception:
                mes_nombre = ""

            row_hist = [
                fecha_emis,
                generador,
                sede,
                correlativo,
                num_cert,
                guia_completa,
                "", # Col G: Link guía
                "", # Col H: Link Word
                val_link, # Col I: Link PDF Expediente
                obs_texto, # Col J: Obs
                mes_nombre # Col K: Mes
            ]
            sht.spreadsheets().values().append(
                spreadsheetId=ID_SHEET_CONTROL,
                range=f"'{hoja_historial}'!A:K",
                valueInputOption="USER_ENTERED",
                body={"values": [row_hist]}
            ).execute()
            matched_row = len(col_corr) + 1

        return {
            "status": "success",
            "empresa": empresa,
            "correlativo": correlativo,
            "hoja_historial": hoja_historial,
            "fila_historial": matched_row,
            "link_expediente": link_expediente,
            "link_petramas": link_petramas,
            "link_sustento": link_sustento,
            "total_paginas": total_paginas
        }
    finally:
        # Liberar memoria y archivos temporales de disco de manera estricta
        gc.collect()
        try:
            for arch in os.listdir(temp_dir):
                os.remove(os.path.join(temp_dir, arch))
            os.rmdir(temp_dir)
        except Exception:
            pass

# ====================================================================
# --- 8. ENSAMBLADOR DE EXPEDIENTE FINAL EN 5 PARTES SECUENCIALES ---
# ====================================================================
def ensamblar_expediente_disposicion_final(
    empresa: str,
    correlativo: str,
    cert_caratula_pdf_bytes: bytes,
    petramas_pdf_bytes: Optional[bytes] = None,
    sustento_pdf_bytes: Optional[bytes] = None,
    paginas_manifiestos: Optional[List[int]] = None,
    paginas_boletas: Optional[List[int]] = None,
    paginas_guias: Optional[List[int]] = None,
    orden_secciones: Optional[List[str]] = None,
    usuario_editor: str = "Ensamblador Petramás"
) -> Dict[str, Any]:
    """
    Fusiona quirúrgicamente los documentos en las 5 partes secuenciales obligatorias:
    1. Certificado Emitido
    2. Constancia de Petramás
    3. Manifiestos de Manejo de Residuos
    4. Boletas de Pesaje
    5. Guías de Remisión y Transporte
    Sube a 03_Expedientes_Finales/<Empresa>/ y sincroniza con Google Sheets.
    """
    empresa_norm = normalizar_empresa_eors(empresa)
    orden = orden_secciones or ["caratula", "petramas", "manifiestos", "boletas", "guias"]

    writer = PdfWriter()
    reader_caratula = PdfReader(io.BytesIO(cert_caratula_pdf_bytes))
    paginas_dict = {
        "caratula": list(reader_caratula.pages),
        "petramas": [],
        "manifiestos": [],
        "boletas": [],
        "guias": []
    }

    if petramas_pdf_bytes:
        r_pet = PdfReader(io.BytesIO(petramas_pdf_bytes))
        paginas_dict["petramas"] = list(r_pet.pages)

    if sustento_pdf_bytes:
        r_sus = PdfReader(io.BytesIO(sustento_pdf_bytes))
        total_p = len(r_sus.pages)

        if paginas_manifiestos:
            for p in paginas_manifiestos:
                if 1 <= p <= total_p:
                    paginas_dict["manifiestos"].append(r_sus.pages[p - 1])

        if paginas_boletas:
            for p in paginas_boletas:
                if 1 <= p <= total_p:
                    paginas_dict["boletas"].append(r_sus.pages[p - 1])

        if paginas_guias:
            for p in paginas_guias:
                if 1 <= p <= total_p:
                    paginas_dict["guias"].append(r_sus.pages[p - 1])

    # Ensamblar en el orden estricto solicitado
    for sec in orden:
        pags = paginas_dict.get(sec, [])
        for p in pags:
            writer.add_page(p)

    buf_out = io.BytesIO()
    writer.write(buf_out)
    pdf_unificado_bytes = buf_out.getvalue()
    total_pags = len(writer.pages)
    buf_out.close()

    # Subir a Google Drive
    nombre_pdf = f"EXPEDIENTE-FINAL-{empresa_norm[:4].strip()}-{correlativo}.pdf"
    res_drive = subir_pdf_petramas_drive(pdf_unificado_bytes, "expedientes", empresa_norm, nombre_pdf)
    link_pdf = res_drive.get("link", "")

    # Actualizar pestaña Historial correspondiente
    hoja_historial = obtener_hoja_historial_empresa(empresa_norm)
    _, sht = obtener_servicios()

    try:
        r_hist = sht.spreadsheets().values().get(
            spreadsheetId=ID_SHEET_CONTROL,
            range=f"'{hoja_historial}'!D:D"
        ).execute()
        col_corr = [str(x[0]).strip() if x else "" for x in r_hist.get('values', [])]
        corr_clean = str(correlativo).strip().lstrip('0')

        matched_row = None
        for i, c in enumerate(col_corr[1:], start=2):
            if c.lstrip('0') == corr_clean and corr_clean:
                matched_row = i
                break

        obs_texto = f"Expediente Disposición Final ({total_pags} págs) - {usuario_editor}"
        val_link = f'=HYPERLINK("{link_pdf}"; "{nombre_pdf}")'

        if matched_row:
            sht.spreadsheets().values().update(
                spreadsheetId=ID_SHEET_CONTROL,
                range=f"'{hoja_historial}'!I{matched_row}:J{matched_row}",
                valueInputOption="USER_ENTERED",
                body={"values": [[val_link, obs_texto]]}
            ).execute()
    except Exception as e_hist:
        print(f"Aviso actualizando Historial: {e_hist}")

    # Liberación de memoria
    del writer
    del reader_caratula
    gc.collect()

    return {
        "status": "success",
        "nombre_archivo": nombre_pdf,
        "enlace_drive": link_pdf,
        "total_paginas": total_pags,
        "hoja_historial": hoja_historial,
        "pdf_bytes": pdf_unificado_bytes
    }

# ====================================================================
# --- 9. CONSULTAS Y ENDPOINT DE CONTROL DE SUSTENTOS ---
# ====================================================================
def obtener_resumen_sustentos_pendientes(empresa: Optional[str] = None) -> Dict[str, Any]:
    """
    Retorna el inventario de sustentos agrupados por estado para el panel o endpoint:
    GET /api/petramas/pendientes
    """
    _, sht = obtener_servicios()
    if not sht:
        return {"pendientes": [], "completos": [], "todos": []}

    try:
        res = sht.spreadsheets().values().get(
            spreadsheetId=ID_SHEET_PETRAMAS,
            range=f"'{SHEET_SUSTENTOS_CAMPO}'!A:Q"
        ).execute()
        filas = res.get('values', [])
        if len(filas) <= 1:
            return {"pendientes": [], "completos": [], "todos": []}

        pendientes = []
        completos = []
        todos = []

        empresa_filtro = normalizar_empresa_eors(empresa) if empresa and empresa != "Todas" else None

        for idx, r in enumerate(filas[1:], start=2):
            emp_r = r[1] if len(r) > 1 else ""
            if empresa_filtro and normalizar_empresa_eors(emp_r) != empresa_filtro:
                continue

            estado = r[15] if len(r) > 15 else ""
            item = {
                "fila": idx,
                "fecha_registro": r[0] if len(r) > 0 else "",
                "empresa": emp_r,
                "generador": r[2] if len(r) > 2 else "",
                "sede": r[3] if len(r) > 3 else "",
                "fecha_servicio": r[4] if len(r) > 4 else "",
                "placa": r[5] if len(r) > 5 else "",
                "guia": r[6] if len(r) > 6 else "",
                "boletas": r[7] if len(r) > 7 else "",
                "tiene_manifiestos": r[8] if len(r) > 8 else "NO",
                "tiene_boletas": r[9] if len(r) > 9 else "NO",
                "tiene_guias": r[10] if len(r) > 10 else "NO",
                "rangos": r[11] if len(r) > 11 else "",
                "enlace_drive": r[12] if len(r) > 12 else "",
                "cert_petramas": r[13] if len(r) > 13 else "",
                "cert_epmi": r[14] if len(r) > 14 else "",
                "estado": estado,
                "enlace_expediente": r[16] if len(r) > 16 else ""
            }
            todos.append(item)
            if "PENDIENTE" in estado.upper():
                pendientes.append(item)
            elif "COMPLETO" in estado.upper() and not item.get("cert_epmi"):
                completos.append(item)

        return {
            "pendientes": pendientes,
            "completos": completos,
            "todos": todos
        }
    except Exception as e:
        print(f"Error consultando sustentos: {e}")
        return {"pendientes": [], "completos": [], "todos": []}

def obtener_constancias_petramas_registradas(empresa: Optional[str] = None) -> List[Dict[str, Any]]:
    """Consulta todas las constancias registradas en 'Certificados_Petramas'."""
    _, sht = obtener_servicios()
    if not sht:
        return []

    try:
        res = sht.spreadsheets().values().get(
            spreadsheetId=ID_SHEET_PETRAMAS,
            range=f"'{SHEET_CERTIFICADOS_PETRAMAS}'!A:K"
        ).execute()
        filas = res.get('values', [])
        if len(filas) <= 1:
            return []

        empresa_filtro = normalizar_empresa_eors(empresa) if empresa and empresa != "Todas" else None
        constancias = []

        for idx, r in enumerate(filas[1:], start=2):
            emp_r = r[1] if len(r) > 1 else ""
            if empresa_filtro and normalizar_empresa_eors(emp_r) != empresa_filtro:
                continue

            constancias.append({
                "fila": idx,
                "fecha_registro": r[0] if len(r) > 0 else "",
                "empresa": emp_r,
                "numero_constancia": r[2] if len(r) > 2 else "",
                "fecha_disposicion": r[3] if len(r) > 3 else "",
                "generador": r[4] if len(r) > 4 else "",
                "sede": r[5] if len(r) > 5 else "",
                "boletas": r[6] if len(r) > 6 else "",
                "peso_total_kg": r[7] if len(r) > 7 else 0.0,
                "enlace_drive": r[8] if len(r) > 8 else "",
                "certificado_ligado": r[9] if len(r) > 9 else "",
                "estado": r[10] if len(r) > 10 else "REGISTRADO_DISPONIBLE"
            })
        return constancias
    except Exception as e:
        print(f"Error consultando constancias Petramás: {e}")
        return []
