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

from src.utils.format_utils import limpiar_monto

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

def migrar_datos_certificado_a_sigersol(
    df_items: pd.DataFrame,
    cabecera: dict,
    tipo_flujo: str,
    emisor_info: dict = None
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

        # 7. Ejecutar inserción en Google Sheets
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
            "filas_insertadas": len(filas_para_insertar),
            "pestana": pestana_real,
            "fila_inicio": fila_inicio_num,
            "error": None
        }

    except Exception as e:
        print(f"[!] Error crítico en migrar_datos_certificado_a_sigersol: {e}")
        return {
            "ok": False,
            "filas_insertadas": 0,
            "pestana": "",
            "error": str(e)
        }
