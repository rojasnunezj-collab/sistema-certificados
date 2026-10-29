import re
import os
from datetime import datetime
import pandas as pd
from src.config.settings import ID_SHEET_SIGERSOL
from src.services.google_service import obtener_servicios

MESES_NOMBRE = {
    "01": "ENERO", "02": "FEBRERO", "03": "MARZO", "04": "ABRIL",
    "05": "MAYO", "06": "JUNIO", "07": "JULIO", "08": "AGOSTO",
    "09": "SETIEMBRE", "10": "OCTUBRE", "11": "NOVIEMBRE", "12": "DICIEMBRE",
    "1": "ENERO", "2": "FEBRERO", "3": "MARZO", "4": "ABRIL",
    "5": "MAYO", "6": "JUNIO", "7": "JULIO", "8": "AGOSTO",
    "9": "SETIEMBRE"
}

def normalizar_texto_clave(s: str) -> str:
    """Elimina acentos, espacios y símbolos para coincidencia exacta de columnas."""
    t = str(s).upper().strip()
    t = re.sub(r'[ÁÀÄÂ]', 'A', t)
    t = re.sub(r'[ÉÈËÊ]', 'E', t)
    t = re.sub(r'[ÍÌÏÎ]', 'I', t)
    t = re.sub(r'[ÓÒÖÔ]', 'O', t)
    t = re.sub(r'[ÚÙÜÛ]', 'U', t)
    t = re.sub(r'[^A-Z0-9]', '', t)
    return t

def extraer_mes_y_trimestre(fecha_str: str) -> tuple[str, str]:
    """Extrae el nombre del mes y la clave del trimestre a partir de una fecha DD/MM/YYYY."""
    try:
        partes = str(fecha_str).strip().split('/')
        if len(partes) >= 2:
            mes_num = partes[1].zfill(2)
            nombre_mes = MESES_NOMBRE.get(mes_num, "DESCONOCIDO")
            m_int = int(mes_num)
            if 1 <= m_int <= 3:
                trim = "ENE-FEB-MAR"
            elif 4 <= m_int <= 6:
                trim = "ABR-MAY-JUN"
            elif 7 <= m_int <= 9:
                trim = "JUL-AGO-SET"
            else:
                trim = "OCT-NOV-DIC"
            return nombre_mes, trim
    except Exception:
        pass
    
    # Fallback fecha actual
    mes_num = datetime.now().strftime("%m")
    nombre_mes = MESES_NOMBRE.get(mes_num, "ENERO")
    m_int = int(mes_num)
    if 1 <= m_int <= 3: trim = "ENE-FEB-MAR"
    elif 4 <= m_int <= 6: trim = "ABR-MAY-JUN"
    elif 7 <= m_int <= 9: trim = "JUL-AGO-SET"
    else: trim = "OCT-NOV-DIC"
    return nombre_mes, trim

def clasificar_basilea(descripcion: str, tipo_flujo: str = "", tipo_residuo: str = "") -> tuple[str, bool]:
    """
    Determina el código de Basilea y si es Peligroso (True) o No Peligroso (False).
    """
    desc = str(descripcion).upper().strip()
    flujo = str(tipo_flujo).upper()
    res = str(tipo_residuo).upper()
    
    es_peligroso_declarado = ("PELIGROSO" in res and "NO PELIGROSO" not in res)
    
    if any(k in desc for k in ["CARTON", "PAPEL", "ARCHIVADOR"]):
        return ("B3020", False)
    if any(k in desc for k in ["PLASTICO", "PLÁSTICO", "FILM", "PET", "STRETCH", "POLIETILENO", "PVC"]):
        return ("B3010", False)
    if any(k in desc for k in ["CHATARRA", "METAL", "HIERRO", "FIERRO", "COBRE", "ALUMINIO", "BRONCE", "ACERO"]):
        return ("B1010", False)
    if any(k in desc for k in ["BATERIA", "BATERÍA", "PILA", "ACUMULADOR"]):
        return ("A1170", True)
    if any(k in desc for k in ["ACEITE", "HIDROCARBURO", "COMBUSTIBLE", "GRASA", "LUBRICANTE", "BORRA"]):
        return ("A4060", True)
    if any(k in desc for k in ["BIOCONTAMINADO", "MEDICO", "MÉDICO", "PUNZOCORTANTE", "HOSPITALARIO"]):
        return ("A4020", True)
    if any(k in desc for k in ["PINTURA", "SOLVENTE", "QUIMICO", "QUÍMICO", "ACIDO", "ÁCIDO", "REACTIVO"]):
        return ("A4070", True)
    if any(k in desc for k in ["MADERA", "PALLET", "PARIHUELA"]):
        return ("B3050", False)
    if any(k in desc for k in ["VIDRIO"]):
        return ("B2020", False)
    if any(k in desc for k in ["LLANTA", "NEUMATICO", "NEUMÁTICO", "CAUCHO"]):
        return ("B3140", False)
    if any(k in desc for k in ["GENERAL", "COMUN", "COMÚN", "NO REAPROVECHABLE", "DOMESTICO", "DOMÉSTICO"]):
        return ("RSMNM", False)
        
    if es_peligroso_declarado:
        return ("A4020", True)
    elif "COMERCIALIZA" in flujo:
        return ("B3010", False)
    else:
        return ("RSMNM", False)

from src.utils.format_utils import (
    limpiar_monto,
    limpiar_guia_o_factura,
    normalizar_fecha,
    limpiar_descripcion
)

def limpiar_peso_a_toneladas(peso_val) -> float:
    """Convierte peso (en Kg) a Toneladas dividiendo entre 1000."""
    try:
        val_kg = limpiar_monto(peso_val)
        return round(val_kg / 1000.0, 3)
    except Exception:
        return 0.0

def formatear_decimal_es(num_float: float) -> str:
    """Formatea el float con coma decimal para compatibilidad con la hoja de cálculo."""
    if num_float == 0.0:
        return "0,00"
    s = f"{num_float:.3f}".rstrip('0').rstrip('.')
    if '.' not in s and ',' not in s:
        s += ",00"
    return s.replace('.', ',')

def asegurar_pestana_trimestre(sheets_api, spreadsheet_id: str, nombre_pestana: str) -> str:
    """
    Verifica si existe la pestaña trimestral en la hoja destino.
    Si no existe, la crea duplicando una de plantilla base (o agregándola con headers).
    """
    try:
        meta = sheets_api.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        pestanas = meta.get('sheets', [])
        
        nombre_norm = normalizar_texto_clave(nombre_pestana)
        for s in pestanas:
            p_title = s['properties']['title']
            if normalizar_texto_clave(p_title) == nombre_norm:
                return p_title  # Retorna el título exacto existente
                
        # Si no existe, buscamos una hoja modelo para duplicar formato y cabecera
        sheet_id_modelo = None
        sufijo_buscar = "COM" if "COM" in nombre_norm else "SERV"
        for s in pestanas:
            t = normalizar_texto_clave(s['properties']['title'])
            if sufijo_buscar in t and ("OCTNOVDIC" in t or "JULAGOSET" in t or "ENEFEBMAR" in t):
                sheet_id_modelo = s['properties']['sheetId']
                break
                
        if sheet_id_modelo is not None:
            # Duplicar hoja modelo
            req = {
                "duplicateSheet": {
                    "sourceSheetId": sheet_id_modelo,
                    "newSheetName": nombre_pestana
                }
            }
            res = sheets_api.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={"requests": [req]}
            ).execute()
            new_id = res['replies'][0]['duplicateSheet']['properties']['sheetId']
            # Limpiar contenido a partir de la fila 6 para que quede solo el header
            limpieza = {
                "updateCells": {
                    "range": {
                        "sheetId": new_id,
                        "startRowIndex": 5, # Fila 6 (0-based 5)
                        "startColumnIndex": 0
                    },
                    "fields": "userEnteredValue"
                }
            }
            sheets_api.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={"requests": [limpieza]}
            ).execute()
            return nombre_pestana
        else:
            # Crear hoja en blanco con encabezados estándar
            req = {"addSheet": {"properties": {"title": nombre_pestana}}}
            sheets_api.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={"requests": [req]}
            ).execute()
            headers = [
                ["ITEM", "FECHA", "MES", "EMPRESA GENERADORA", "RUC ", "DIRECCION DEL GENERADOR", 
                 "N° GUIA", "N° FACTURA", "DESCRIPCION DE RESIDUOS", "COD. BASILEA", 
                 "Peligroso", "No Peligroso", "CANTIDAD (Tn)", "CANT. TOTAL", "PLACA(S)", 
                 "DESTINO FINAL ", "REGISTRO EO.RS"]
            ]
            sheets_api.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"'{nombre_pestana}'!A5:Q5",
                valueInputOption="USER_ENTERED",
                body={"values": headers}
            ).execute()
            return nombre_pestana
    except Exception as e:
        print(f"[!] Error asegurando pestaña '{nombre_pestana}': {e}")
        return nombre_pestana

def mapear_columnas_tabla_docx(hdr_cells: list) -> dict:
    """Mapea los encabezados de la tabla del documento Word a los nombres de campos internos."""
    col_map = {}
    for idx, cell_text in enumerate(hdr_cells):
        t = normalizar_texto_clave(cell_text)
        if 'FECHA' in t and 'fecha_origen' not in col_map:
            col_map['fecha_origen'] = idx
        elif any(k in t for k in ['PLACA', 'VEHICULO']) and 'placa_origen' not in col_map:
            col_map['placa_origen'] = idx
        elif 'GUIA' in t and 'guia_origen' not in col_map:
            col_map['guia_origen'] = idx
        elif any(k in t for k in ['DESC', 'RESIDUO', 'MATERIAL', 'DETALLE']) and 'desc' not in col_map:
            col_map['desc'] = idx
        elif 'CANT' in t and 'cant' not in col_map:
            col_map['cant'] = idx
        elif any(k in t for k in ['UM', 'MEDIDA', 'UNIDAD']) and 'um' not in col_map:
            col_map['um'] = idx
        elif any(k in t for k in ['PESO', 'KILO', 'TN', 'TON']) and 'peso' not in col_map:
            col_map['peso'] = idx
    return col_map

def extraer_datos_de_docx_para_sigersol(docx_bytes: bytes) -> tuple[pd.DataFrame, dict]:
    """
    Extrae las filas de la tabla de residuos y metadatos desde un archivo .docx modificado.
    """
    from docx import Document
    import io
    import re
    from src.utils.format_utils import limpiar_guia_o_factura, normalizar_fecha

    doc = Document(io.BytesIO(docx_bytes))
    df_items = pd.DataFrame()
    cab_extraida = {}

    for table in doc.tables:
        if len(table.rows) < 2:
            continue
        hdr_cells = [c.text.strip() for c in table.rows[0].cells]
        col_map = mapear_columnas_tabla_docx(hdr_cells)

        if 'desc' in col_map and ('guia_origen' in col_map or 'peso' in col_map or 'placa_origen' in col_map):
            filas_items = []
            for row in table.rows[1:]:
                vals = [c.text.strip() for c in row.cells]
                if not any(vals):
                    continue
                d_val = vals[col_map['desc']].upper() if 'desc' in col_map and col_map['desc'] < len(vals) else ''
                if any(x in d_val for x in ['TOTAL', 'SUMA', 'SUBTOTAL']):
                    continue
                item = {
                    'fecha_origen': normalizar_fecha(vals[col_map['fecha_origen']]) if 'fecha_origen' in col_map and col_map['fecha_origen'] < len(vals) else '',
                    'placa_origen': vals[col_map['placa_origen']].upper() if 'placa_origen' in col_map and col_map['placa_origen'] < len(vals) else '',
                    'guia_origen': limpiar_guia_o_factura(vals[col_map['guia_origen']]) if 'guia_origen' in col_map and col_map['guia_origen'] < len(vals) else '',
                    'desc': d_val,
                    'cant': vals[col_map['cant']] if 'cant' in col_map and col_map['cant'] < len(vals) else '1',
                    'um': vals[col_map['um']].upper() if 'um' in col_map and col_map['um'] < len(vals) else 'KG',
                    'peso': vals[col_map['peso']] if 'peso' in col_map and col_map['peso'] < len(vals) else '0'
                }
                filas_items.append(item)
            if filas_items:
                df_items = pd.DataFrame(filas_items)
                break

    texto_completo = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    match_ruc = re.search(r'\b(10|20)\d{9}\b', texto_completo)
    if match_ruc:
        cab_extraida['ruc_cliente'] = match_ruc.group(0)

    match_fec = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', texto_completo)
    if match_fec:
        cab_extraida['fecha'] = match_fec.group(1)

    match_corr = re.search(r'(?:CERT-|N°\s*|NRO\.?\s*)(\d{3,6})', texto_completo, re.IGNORECASE)
    if match_corr:
        cab_extraida['correlativo'] = match_corr.group(1)

    return df_items, cab_extraida

def extraer_datos_de_pdf_para_sigersol(pdf_bytes: bytes) -> tuple[pd.DataFrame, dict]:
    """
    Extrae texto y datos tabulares básicos desde un archivo PDF de certificado.
    """
    import io
    import re
    from pypdf import PdfReader
    from src.utils.format_utils import limpiar_guia_o_factura, normalizar_fecha

    df_items = pd.DataFrame()
    cab_extraida = {}

    if not pdf_bytes:
        return df_items, cab_extraida

    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        if not reader.pages:
            return df_items, cab_extraida

        texto_p0 = reader.pages[0].extract_text() or ""
        match_ruc = re.search(r'\b(10|20)\d{9}\b', texto_p0)
        if match_ruc:
            cab_extraida['ruc_cliente'] = match_ruc.group(0)

        match_fec = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', texto_p0)
        if match_fec:
            cab_extraida['fecha'] = match_fec.group(1)

        match_corr = re.search(r'(?:CERT-|N°\s*|NRO\.?\s*)(\d{3,6})', texto_p0, re.IGNORECASE)
        if match_corr:
            cab_extraida['correlativo'] = match_corr.group(1)

        guias_encontradas = re.findall(r'\b([A-Z0-9]{1,4}-\d{3,8})\b', texto_p0)
        if guias_encontradas:
            cab_extraida['guias'] = ", ".join(dict.fromkeys(guias_encontradas))

        lineas = [l.strip() for l in texto_p0.split('\n') if l.strip()]
        filas_items = []
        for l in lineas:
            m_g = re.search(r'\b([A-Z0-9]{1,4}-\d{3,8})\b', l)
            m_f = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', l)
            m_nums = re.findall(r'\b(\d+(?:[.,]\d+)?)\b', l)
            if (m_g or m_f) and m_nums and not any(k in l.upper() for k in ['CERTIFICADO', 'OPERACION', 'RUC', 'TOTAL']):
                guia = limpiar_guia_o_factura(m_g.group(1)) if m_g else ''
                fecha = normalizar_fecha(m_f.group(1)) if m_f else cab_extraida.get('fecha', '')
                peso_candidato = m_nums[-1].replace(',', '.')
                filas_items.append({
                    'fecha_origen': fecha,
                    'placa_origen': '',
                    'guia_origen': guia,
                    'desc': 'RESIDUOS DECLARADOS',
                    'cant': '1',
                    'um': 'KG',
                    'peso': peso_candidato
                })
        if filas_items:
            df_items = pd.DataFrame(filas_items)
    except Exception as e:
        print(f"[Aviso] No se pudo parsear texto de PDF para Sigersol: {e}")

    return df_items, cab_extraida

def migrar_datos_certificado_a_sigersol(
    df_items: pd.DataFrame,
    cabecera: dict,
    tipo_flujo: str,
    emisor_info: dict = None,
    actualizar_si_existe: bool = True,
    guias_anteriores: list = None
) -> dict:
    """
    Migra las filas de residuos de la tabla del certificado hacia el Google Sheet de Sigersol (3 trimestre).
    
    Parámetros:
    - df_items: DataFrame con las filas de la tabla interactiva (fecha_origen, desc, cant, peso, etc.)
    - cabecera: Diccionario con datos del certificado (fecha, cliente, ruc_cliente, partida, guia, placa, etc.)
    - tipo_flujo: 'Comercialización' o 'Disposición Final / Servicio'
    - emisor_info: Diccionario con datos de la empresa emisora (nombre, ruc, registro)
    
    Retorna:
    - dict con {'ok': bool, 'filas_insertadas': int, 'pestana': str, 'error': str|None}
    """
    if df_items is None or df_items.empty:
        return {"ok": False, "filas_insertadas": 0, "pestana": "", "error": "Tabla de residuos vacía"}

    emisor_info = emisor_info or {}
    
    try:
        _, sheets_api = obtener_servicios()
        if not sheets_api:
            return {"ok": False, "filas_insertadas": 0, "pestana": "", "error": "No se pudo autenticar con Google Sheets"}

        spreadsheet_id = ID_SHEET_SIGERSOL

        # 1. Determinar Fecha Representativa y Trimestre
        fecha_repr = str(cabecera.get('fecha', '')).strip()
        if not fecha_repr and 'fecha_origen' in df_items.columns:
            f_val = str(df_items['fecha_origen'].iloc[0]).strip()
            if f_val and f_val.lower() not in ['none', 'nan', '']:
                fecha_repr = f_val
                
        if not fecha_repr:
            fecha_repr = datetime.now().strftime("%d/%m/%Y")

        mes_nombre_rep, trim_rep = extraer_mes_y_trimestre(fecha_repr)
        sufijo = "(COM)" if "comercializa" in str(tipo_flujo).lower() else "(SERV)"
        nombre_pestana_objetivo = f"{trim_rep} {sufijo}"

        # 2. Asegurar que la pestaña exista
        pestana_real = asegurar_pestana_trimestre(sheets_api, spreadsheet_id, nombre_pestana_objetivo)

        # 3. Leer estructura actual de la pestaña
        res_vals = sheets_api.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f"'{pestana_real}'!A1:Z500"
        ).execute()
        todas_filas = res_vals.get('values', [])

        # Localizar fila de encabezados
        hdr_idx = None
        for i, fila in enumerate(todas_filas[:10]):
            fila_txt = [normalizar_texto_clave(c) for c in fila]
            if any('FECHA' in c for c in fila_txt) and any('GENERADORA' in c for c in fila_txt):
                hdr_idx = i
                break

        if hdr_idx is None:
            # Si no se detectó, se asume fila 5 (índice 4)
            hdr_idx = 4
            headers_lista = [
                "ITEM", "FECHA", "MES", "EMPRESA GENERADORA", "RUC ", "DIRECCION DEL GENERADOR", 
                "N° GUIA", "N° FACTURA", "DESCRIPCION DE RESIDUOS", "COD. BASILEA", 
                "Peligroso", "No Peligroso", "CANTIDAD (Tn)", "CANT. TOTAL", "PLACA(S)", 
                "DESTINO FINAL ", "REGISTRO EO.RS"
            ]
        else:
            headers_lista = todas_filas[hdr_idx]

        hdr_map = {normalizar_texto_clave(h): idx for idx, h in enumerate(headers_lista)}
        ancho_fila = max(len(headers_lista), 17)

        # 4. Encontrar primera fila libre
        fila_inicio_num = len(todas_filas) + 1
        for i in range(hdr_idx + 1, len(todas_filas)):
            fila_actual = todas_filas[i]
            # Si tanto la columna de FECHA como la de GENERADORA están vacías, esa es la fila libre
            col_f = hdr_map.get('FECHA', 1)
            col_g = hdr_map.get('EMPRESAGENERADORA', 3)
            val_f = fila_actual[col_f].strip() if len(fila_actual) > col_f else ""
            val_g = fila_actual[col_g].strip() if len(fila_actual) > col_g else ""
            if not val_f and not val_g:
                fila_inicio_num = i + 1
                break

        if fila_inicio_num <= (hdr_idx + 1):
            fila_inicio_num = hdr_idx + 2

        # Contador de ítems existentes
        conteo_items_previos = 0
        for i in range(hdr_idx + 1, fila_inicio_num - 1):
            if i < len(todas_filas) and len(todas_filas[i]) > 0 and todas_filas[i][0].strip():
                conteo_items_previos += 1

        # 5. Calcular Toneladas Totales del Lote / Certificado
        lista_pesos_tn = []
        for _, r in df_items.iterrows():
            p_val = r.get('peso', 0)
            tn = limpiar_peso_a_toneladas(p_val)
            lista_pesos_tn.append(tn)
        total_tn = sum(lista_pesos_tn)
        total_tn_str = formatear_decimal_es(total_tn)

        # 6. Construir filas normalizadas para Google Sheets
        filas_para_insertar = []
        empresa_generadora = str(cabecera.get('cliente', '')).strip().upper()
        ruc_generador = str(cabecera.get('ruc_cliente', '')).strip()
        direccion_base = str(cabecera.get('partida', '')).strip()
        guia_cabecera = str(cabecera.get('guia', '')).strip().upper()
        placa_cabecera = str(cabecera.get('placa', '')).strip().upper()
        tipo_residuo_cab = str(cabecera.get('tipo_residuo', '')).strip().upper()
        
        destino_final_nombre = str(emisor_info.get('nombre', '')).strip().upper()
        registro_eors = str(emisor_info.get('registro', '')).strip()

        for idx_fila, (_, item) in enumerate(df_items.iterrows()):
            # FECHA & MES
            f_item = str(item.get('fecha_origen', '')).strip()
            fecha_final = f_item if f_item and f_item.lower() not in ['none', 'nan', ''] else fecha_repr
            mes_item, _ = extraer_mes_y_trimestre(fecha_final)
            
            # GUIA
            g_item = str(item.get('guia_origen', '')).strip().upper()
            guia_final = g_item if g_item and g_item.lower() not in ['none', 'nan', ''] else guia_cabecera
            
            # PLACA
            p_item = str(item.get('placa_origen', '')).strip().upper()
            placa_final = p_item if p_item and p_item.lower() not in ['none', 'nan', ''] else placa_cabecera

            # DIRECCIÓN
            d_item = str(item.get('ocr_partida', '')).strip()
            dir_final = d_item if d_item and d_item.lower() not in ['none', 'nan', ''] else direccion_base

            # DESCRIPCIÓN & BASILEA
            desc_item = str(item.get('desc', '') or item.get('descripcion', '')).strip().upper()
            cod_basilea, es_peligroso = clasificar_basilea(desc_item, tipo_flujo, tipo_residuo_cab)
            
            val_peligroso = "X" if es_peligroso else ""
            val_no_peligroso = "" if es_peligroso else "X"

            # CANTIDAD (Tn)
            peso_tn = lista_pesos_tn[idx_fila]
            peso_tn_str = formatear_decimal_es(peso_tn)

            # Matriz vacía para la fila
            fila_datos = [""] * ancho_fila

            def asignar(clave_norm, valor):
                pos = hdr_map.get(clave_norm)
                if pos is not None and pos < ancho_fila:
                    fila_datos[pos] = str(valor)

            # Mapeo exacto ordenado a los encabezados
            asignar('ITEM', conteo_items_previos + idx_fila + 1)
            asignar('FECHA', fecha_final)
            asignar('MES', mes_item)
            asignar('EMPRESAGENERADORA', empresa_generadora)
            asignar('RUC', ruc_generador)
            asignar('DIRECCIONDELGENERADOR', dir_final)
            asignar('NGUIA', guia_final)
            asignar('NFACTURA', "")
            asignar('DESCRIPCIONDERESIDUOS', desc_item)
            asignar('CODBASILEA', cod_basilea)
            asignar('PELIGROSO', val_peligroso)
            asignar('NOPELIGROSO', val_no_peligroso)
            asignar('CANTIDADTN', peso_tn_str)
            asignar('CANTTOTAL', total_tn_str)
            asignar('PLACAS', placa_final)
            asignar('DESTINOFINAL', destino_final_nombre)
            asignar('REGISTROEORS', registro_eors)

            filas_para_insertar.append(fila_datos)

        # 6.1. Detección y Sincronización de Filas Existentes (para evitar duplicados o actualizar si fue editado)
        pos_guia = hdr_map.get('NGUIA')
        pos_empresa = hdr_map.get('EMPRESAGENERADORA')
        pos_desc = hdr_map.get('DESCRIPCIONDERESIDUOS')
        pos_fecha = hdr_map.get('FECHA')
        pos_item = hdr_map.get('ITEM')

        guias_busqueda = set()
        def _agregar_guias(val):
            if not val:
                return
            if isinstance(val, (list, set, tuple)):
                for item in val:
                    _agregar_guias(item)
                return
            for parte in str(val).split(','):
                p = parte.strip().upper()
                if p and p not in ['S/N', '-', 'NAN', 'NONE']:
                    guias_busqueda.add(p)

        if guias_anteriores:
            _agregar_guias(guias_anteriores)
        if 'guia_origen' in df_items.columns:
            _agregar_guias(df_items['guia_origen'].unique())
        g_cab = cabecera.get('guia')
        if g_cab:
            _agregar_guias(g_cab)

        guias_busqueda_clean = {limpiar_guia_o_factura(g) for g in guias_busqueda if limpiar_guia_o_factura(g)}

        filas_coincidentes_sheet = [] # Índices 0-based en todas_filas
        if len(todas_filas) > (hdr_idx + 1):
            empresa_norm = normalizar_texto_clave(empresa_generadora)
            for i in range(hdr_idx + 1, fila_inicio_num - 1):
                r = todas_filas[i]
                g_ex = r[pos_guia].strip().upper() if (pos_guia is not None and len(r) > pos_guia) else ""
                g_ex_clean = limpiar_guia_o_factura(g_ex) if g_ex else ""
                e_ex = normalizar_texto_clave(r[pos_empresa]) if (pos_empresa is not None and len(r) > pos_empresa) else ""
                
                # Match prioritario por número de guía (exacto o normalizado sin ceros a la izquierda)
                coincide_guia = False
                if g_ex and g_ex in guias_busqueda:
                    coincide_guia = True
                elif g_ex_clean and g_ex_clean in guias_busqueda_clean:
                    coincide_guia = True

                if coincide_guia:
                    filas_coincidentes_sheet.append(i)
                # Fallback por empresa y fecha si la guía era S/N
                elif not guias_busqueda and empresa_norm and e_ex and (empresa_norm in e_ex or e_ex in empresa_norm):
                    f_ex = r[pos_fecha].strip() if (pos_fecha is not None and len(r) > pos_fecha) else ""
                    if f_ex == fecha_repr:
                        filas_coincidentes_sheet.append(i)

        if filas_coincidentes_sheet:
            # 1. Comprobar si las filas existentes son 100% IDÉNTICAS (sin cambios)
            if len(filas_coincidentes_sheet) == len(filas_para_insertar):
                todas_identicas = True
                for idx_c, r_idx in enumerate(filas_coincidentes_sheet):
                    r_existente = todas_filas[r_idx]
                    r_nueva = filas_para_insertar[idx_c]
                    cols_criticas = [pos_fecha, pos_empresa, pos_guia, pos_desc, hdr_map.get('CANTIDADTN'), hdr_map.get('PLACAS')]
                    for c_pos in cols_criticas:
                        if c_pos is not None:
                            val_ex = r_existente[c_pos].strip().upper() if len(r_existente) > c_pos else ""
                            val_nu = r_nueva[c_pos].strip().upper() if len(r_nueva) > c_pos else ""
                            if val_ex != val_nu:
                                todas_identicas = False
                                break
                    if not todas_identicas:
                        break

                if todas_identicas:
                    print(f"[Sigersol] Se omitió inserción: las {len(filas_para_insertar)} filas ya existen y son idénticas en '{pestana_real}'.")
                    return {
                        "ok": True,
                        "filas_afectadas": len(filas_para_insertar),
                        "filas_insertadas": 0,
                        "pestana": pestana_real,
                        "fila_inicio": fila_inicio_num,
                        "accion": "sin_cambios",
                        "ya_existia": True,
                        "aviso": f"Los datos de esta guía ya se encuentran registrados en Sigersol ({pestana_real}) sin modificaciones.",
                        "error": None
                    }

            # 2. Si son diferentes y actualizar_si_existe está activo -> ACTUALIZACIÓN IN-PLACE
            if actualizar_si_existe:
                print(f"[Sigersol] Actualizando in-place {len(filas_para_insertar)} fila(s) editadas en la pestaña '{pestana_real}'.")
                for idx_u, r_idx in enumerate(filas_coincidentes_sheet):
                    if idx_u < len(filas_para_insertar):
                        if pos_item is not None and len(todas_filas[r_idx]) > pos_item and todas_filas[r_idx][pos_item].strip():
                            filas_para_insertar[idx_u][pos_item] = todas_filas[r_idx][pos_item].strip()
                        
                        rango_fila = f"'{pestana_real}'!A{r_idx + 1}"
                        sheets_api.spreadsheets().values().update(
                            spreadsheetId=spreadsheet_id,
                            range=rango_fila,
                            valueInputOption="USER_ENTERED",
                            body={"values": [filas_para_insertar[idx_u]]}
                        ).execute()

                if len(filas_para_insertar) > len(filas_coincidentes_sheet):
                    sobrantes = filas_para_insertar[len(filas_coincidentes_sheet):]
                    rango_extra = f"'{pestana_real}'!A{fila_inicio_num}"
                    sheets_api.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=rango_extra,
                        valueInputOption="USER_ENTERED",
                        body={"values": sobrantes}
                    ).execute()
                elif len(filas_para_insertar) < len(filas_coincidentes_sheet):
                    for r_del in filas_coincidentes_sheet[len(filas_para_insertar):]:
                        sheets_api.spreadsheets().values().clear(
                            spreadsheetId=spreadsheet_id,
                            range=f"'{pestana_real}'!A{r_del + 1}:Z{r_del + 1}"
                        ).execute()

                return {
                    "ok": True,
                    "filas_afectadas": len(filas_para_insertar),
                    "filas_insertadas": len(filas_para_insertar),
                    "pestana": pestana_real,
                    "fila_inicio": filas_coincidentes_sheet[0] + 1,
                    "accion": "actualizaron",
                    "ya_existia": False,
                    "error": None
                }

        # 7. Ejecutar inserción en Google Sheets (filas nuevas)
        rango_destino = f"'{pestana_real}'!A{fila_inicio_num}"
        body_update = {"values": filas_para_insertar}
        sheets_api.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=rango_destino,
            valueInputOption="USER_ENTERED",
            body=body_update
        ).execute()

        return {
            "ok": True,
            "filas_afectadas": len(filas_para_insertar),
            "filas_insertadas": len(filas_para_insertar),
            "pestana": pestana_real,
            "fila_inicio": fila_inicio_num,
            "accion": "insertaron",
            "error": None
        }

    except Exception as e:
        print(f"[!] Error crítico en migrar_datos_certificado_a_sigersol: {e}")
        return {
            "ok": False,
            "filas_insertadas": 0,
            "filas_afectadas": 0,
            "pestana": "",
            "error": str(e)
        }

def sincronizar_certificado_editado_a_sigersol(
    docx_bytes: bytes = None,
    pdf_bytes: bytes = None,
    cert_historial: dict = None,
    df_items_override: pd.DataFrame = None,
    cabecera_override: dict = None
) -> dict:
    """
    Sincroniza un certificado o expediente que fue editado/modificado hacia el Google Sheet de Sigersol.
    Extrae la tabla de residuos actualizada del .docx (o usa df_items_override) y actualiza
    in-place las filas existentes en Sigersol correspondientes a las guías de dicho certificado.
    """
    cert_historial = cert_historial or {}
    
    # 1. Extraer DataFrame de items
    df_items = None
    cab_extraida = {}
    
    if df_items_override is not None and not df_items_override.empty:
        df_items = df_items_override.copy()
    elif docx_bytes:
        df_items, cab_extraida = extraer_datos_de_docx_para_sigersol(docx_bytes)

    if (df_items is None or df_items.empty) and pdf_bytes:
        df_items, cab_pdf = extraer_datos_de_pdf_para_sigersol(pdf_bytes)
        if not cab_extraida:
            cab_extraida = cab_pdf

    # Si aún no tenemos items pero cert_historial tiene guias:
    if (df_items is None or df_items.empty) and cert_historial.get('guias'):
        raw_guias = cert_historial.get('guias', '')
        g_list = [g.strip().upper() for g in raw_guias.split(',') if g.strip()]
        if g_list:
            df_items = pd.DataFrame([{
                'fecha_origen': cert_historial.get('fecha', ''),
                'placa_origen': '',
                'guia_origen': g,
                'desc': 'RESIDUOS DECLARADOS',
                'cant': '1',
                'um': 'KG',
                'peso': '0'
            } for g in g_list])
        
    if df_items is None or df_items.empty:
        return {
            "ok": False,
            "filas_afectadas": 0,
            "error": "No se encontraron filas de residuos en el documento editado para sincronizar con Sigersol."
        }

    # 2. Armar cabecera con prioridad: cabecera_override > cab_extraida > cert_historial
    empresa = cert_historial.get('empresa') or cab_extraida.get('cliente', '')
    fecha = cert_historial.get('fecha') or cab_extraida.get('fecha', '')
    correlativo = cert_historial.get('correlativo', '')
    partida = cert_historial.get('fundo', '')
    tipo_flujo = cert_historial.get('tipo_cert') or 'Comercialización'
    
    cabecera = {
        "fecha": fecha,
        "cliente": empresa,
        "ruc_cliente": cab_extraida.get('ruc_cliente', '') or cert_historial.get('ruc_cliente', ''),
        "partida": partida,
        "guia": cert_historial.get('guias', ''),
        "correlativo": correlativo,
        "tipo_residuo": ""
    }
    if cabecera_override:
        cabecera.update(cabecera_override)

    # 3. Resolver datos de la empresa emisora (EO-RS)
    try:
        from src.services.google_service import obtener_configuracion_empresas, buscar_config_empresa
        cfgs = obtener_configuracion_empresas()
        hoja_hist = cert_historial.get('hoja_historial', 'Historial')
        cfg_emisor = None
        for emp_nom, c_val in cfgs.items():
            if c_val.get('hoja_historial') == hoja_hist:
                cfg_emisor = c_val
                break
        if not cfg_emisor:
            cfg_emisor = buscar_config_empresa(None, cfgs)
    except Exception:
        cfg_emisor = {}
        
    emisor_info = {
        "nombre": cfg_emisor.get('nombre', ''),
        "ruc": cfg_emisor.get('ruc', ''),
        "registro": cfg_emisor.get('registro', '')
    }

    # 4. Obtener lista de guías anteriores para localización
    raw_guias = cert_historial.get('guias', '')
    guias_anteriores = [g.strip().upper() for g in raw_guias.split(',') if g.strip()] if raw_guias else []

    # 5. Ejecutar migración / actualización in-place en Sigersol
    return migrar_datos_certificado_a_sigersol(
        df_items=df_items,
        cabecera=cabecera,
        tipo_flujo=tipo_flujo,
        emisor_info=emisor_info,
        actualizar_si_existe=True,
        guias_anteriores=guias_anteriores
    )
