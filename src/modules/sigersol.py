import streamlit as st
import pandas as pd
from datetime import datetime
from googleapiclient.discovery import build

from src.config.settings import ID_SHEET_CONTROL
from src.services.google_service import obtener_servicios, descargar_guias_drive
from src.services.vertex_service import procesar_guia_ia_vertex
from src.utils.format_utils import limpiar_monto, formato_inteligente
import time

# ID del destino dado en el prompt
ID_DESTINO_SUNAT = "1EBS1S0-oHaCBKuHPgjxIkHMLp__TDYgj"

def leer_registro_guias():
    _, sheets = obtener_servicios()
    if not sheets: return pd.DataFrame()
    try:
        # Obtenemos todo el rango incluyendo la columna O (Sigersol)
        r = sheets.spreadsheets().values().get(spreadsheetId=ID_SHEET_CONTROL, range="'Registro_Guias'!A:P").execute()
        v = r.get('values', [])
        if not v or len(v) < 2: return pd.DataFrame()
        
        headers = v[0]
        # Aseguramos que todas las filas tengan la misma longitud que los headers
        filas = []
        for i, fila in enumerate(v[1:]):
            fila_completa = fila + [''] * (len(headers) - len(fila))
            fila_completa = fila_completa[:len(headers)]
            # Guardamos el índice real de Excel (i + 2) para poder hacer update luego
            fila_completa.append(i + 2)
            filas.append(fila_completa)
            
        headers_con_idx = headers + ['_excel_row_idx']
        df = pd.DataFrame(filas, columns=headers_con_idx)
        return df
    except Exception as e:
        st.error(f"Error leyendo Registro_Guias: {e}")
        return pd.DataFrame()

def actualizar_sigersol_origen(filas_excel_idx):
    _, sheets = obtener_servicios()
    if not sheets or not filas_excel_idx: return False
    try:
        data = []
        marca = f"✅ Sigersol {datetime.now().strftime('%d/%m/%Y')}"
        for idx in filas_excel_idx:
            # Columna O es la 15. A=1... O=15. En range es 'Registro_Guias'!O{idx}
            data.append({
                "range": f"'Registro_Guias'!O{idx}",
                "values": [[marca]]
            })
        body = {
            "valueInputOption": "USER_ENTERED",
            "data": data
        }
        sheets.spreadsheets().values().batchUpdate(spreadsheetId=ID_SHEET_CONTROL, body=body).execute()
        return True
    except Exception as e:
        st.error(f"Error actualizando origen: {e}")
        return False

def asegurar_pestana_destino(sheets, nombre_pestana):
    try:
        # Obtenemos metadata del spreadsheet destino
        meta = sheets.spreadsheets().get(spreadsheetId=ID_DESTINO_SUNAT).execute()
        pestanas_existentes = [s['properties']['title'] for s in meta.get('sheets', [])]
        
        if nombre_pestana not in pestanas_existentes:
            # Crear pestaña
            req = {
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": nombre_pestana
                            }
                        }
                    }
                ]
            }
            sheets.spreadsheets().batchUpdate(spreadsheetId=ID_DESTINO_SUNAT, body=req).execute()
            
            # Inicializar headers si es nueva
            headers = [
                ["", "FECHA", "MES", "EMPRESA GENERADORA", "RUC", "DIRECCION DEL GENERADOR", 
                 "N° GUIA", "", "DESCRIPCION DE RESIDUOS", "", "", "", "CANTIDAD (Tn)", "", "DESTINO FINAL"]
            ]
            sheets.spreadsheets().values().update(
                spreadsheetId=ID_DESTINO_SUNAT,
                range=f"'{nombre_pestana}'!A1:O1",
                valueInputOption="USER_ENTERED",
                body={"values": headers}
            ).execute()
            
        return True
    except Exception as e:
        st.error(f"Error al verificar/crear pestaña '{nombre_pestana}' en destino: {e}")
        return False

def migrar_a_sunat(df_seleccionados, nombre_pestana):
    _, sheets = obtener_servicios()
    if not sheets: return False
    
    if not asegurar_pestana_destino(sheets, nombre_pestana):
        return False
        
    try:
        valores_a_insertar = []
        for _, row in df_seleccionados.iterrows():
            # Mapeo estricto
            # Columna B: FECHA -> row['FECHA']
            # Columna C: MES -> row['MES']
            # Columna D: EMPRESA GENERADORA -> row['EMPRESA GENERADORA']
            # Columna E: RUC -> row['RUC']
            # Columna F: DIRECCION DEL GENERADOR -> row['DIRECCION DEL GENERADOR']
            # Columna G: N° GUIA -> row['N° GUIA (SUNAT)']
            # Columna H: (vacia)
            # Columna I: DESCRIPCION DE RESIDUOS -> row['DESCRIPCION DE RESIDUOS']
            # Columna J, K, L: (vacias)
            # Columna M: CANTIDAD (Tn) -> row['CANTIDAD (Tn)']
            # Columna N: (vacia)
            # Columna O: DESTINO FINAL -> row['DESTINO FINAL']
            
            fila = [
                "", # A
                str(row.get('FECHA', '')), # B
                str(row.get('MES', '')), # C
                str(row.get('EMPRESA GENERADORA', '')), # D
                str(row.get('RUC', '')), # E
                str(row.get('DIRECCION DEL GENERADOR', '')), # F
                str(row.get('N° GUIA (SUNAT)', '')), # G
                "", # H
                str(row.get('DESCRIPCION DE RESIDUOS', '')), # I
                "", # J
                "", # K
                "", # L
                str(row.get('CANTIDAD (Tn)', '')), # M
                "", # N
                str(row.get('DESTINO FINAL', '')) # O
            ]
            valores_a_insertar.append(fila)
            
        if valores_a_insertar:
            sheets.spreadsheets().values().append(
                spreadsheetId=ID_DESTINO_SUNAT,
                range=f"'{nombre_pestana}'!A:O",
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={"values": valores_a_insertar}
            ).execute()
            return True
    except Exception as e:
        st.error(f"Error migrando datos a SUNAT: {e}")
        return False
    return False

def render_sigersol():
    st.header("🏢 Módulo Sigersol - Declaración SUNAT")
    st.info("Gestión y migración de guías certificadas hacia la base de datos de SUNAT.")
    
    # --- 1. Filtros ---
    c1, c2, c3 = st.columns(3)
    
    # Cargamos el dataframe base
    with st.spinner("Cargando Registro de Guías..."):
        df_base = leer_registro_guias()
        
    if df_base.empty:
        st.warning("No se pudieron cargar los registros. Verifica la conexión a la base de datos.")
        return
        
    # Identificar columnas. Asumimos la estructura estandar si coinciden los nombres.
    col_mes = 'Mes' if 'Mes' in df_base.columns else df_base.columns[13] if len(df_base.columns) > 13 else None
    col_empresa = 'Empresa Principal' if 'Empresa Principal' in df_base.columns else df_base.columns[5] if len(df_base.columns) > 5 else None
    col_sigersol = 'Sigersol' if 'Sigersol' in df_base.columns else df_base.columns[14] if len(df_base.columns) > 14 else None
    col_fecha = 'Fecha' if 'Fecha' in df_base.columns else df_base.columns[0]
    col_guia_original = 'N° Guía' if 'N° Guía' in df_base.columns else 'N Gua' if 'N Gua' in df_base.columns else df_base.columns[1]
    col_guia_ligada = 'Guía ligada' if 'Guía ligada' in df_base.columns else 'Gua ligada' if 'Gua ligada' in df_base.columns else df_base.columns[2]
    col_guia_hecha = 'Guia hecha' if 'Guia hecha' in df_base.columns else df_base.columns[8] if len(df_base.columns) > 8 else None
    
    # Llenamos opciones
    opciones_meses = ["Todos"] + sorted([str(m) for m in df_base[col_mes].unique() if str(m).strip()]) if col_mes else ["Todos"]
    opciones_empresas = ["Todas"] + sorted([str(e) for e in df_base[col_empresa].unique() if str(e).strip()]) if col_empresa else ["Todas"]
    
    mes_sel = c1.selectbox("Mes de Regularización", options=opciones_meses)
    empresa_sel = c2.selectbox("Empresa Principal", options=opciones_empresas)
    tipo_op = c3.radio("Tipo de Operación", options=["Comercialización", "Disposición Final"])
    
    # --- 2. Lógica de Filtrado ---
    # Filtrar solo las que NO tengan "✅ Sigersol"
    if col_sigersol:
        mask_pendientes = ~df_base[col_sigersol].astype(str).str.contains("✅ Sigersol", case=False, na=False)
        df_filtrado = df_base[mask_pendientes].copy()
    else:
        df_filtrado = df_base.copy()
        
    if mes_sel != "Todos" and col_mes:
        df_filtrado = df_filtrado[df_filtrado[col_mes].astype(str) == mes_sel]
        
    if empresa_sel != "Todas" and col_empresa:
        df_filtrado = df_filtrado[df_filtrado[col_empresa].astype(str) == empresa_sel]
        
    if df_filtrado.empty:
        st.info("No hay guías pendientes de declarar para los filtros seleccionados.")
        return
        
    # --- 3. Preparar DataFrame Editable ---
    if 'sigersol_ia_cache' not in st.session_state:
        st.session_state.sigersol_ia_cache = {}
        
    # Botón de extracción con IA
    st.markdown("---")
    st.markdown("#### 🤖 Extracción Inteligente de Tonelajes")
    st.caption("Usa Vertex AI para leer las guías escaneadas y extraer las toneladas totales automáticamente.")
    
    if st.button("🧠 Auto-Completar Datos con IA (Vertex)", use_container_width=True):
        if col_guia_hecha:
            with st.spinner("Conectando con Google Drive y Vertex AI..."):
                drv, _ = obtener_servicios()
                if drv:
                    progreso = st.progress(0)
                    total_docs = len(df_filtrado)
                    
                    for i, (idx_df, row) in enumerate(df_filtrado.iterrows()):
                        row_excel_idx = row['_excel_row_idx']
                        nombre_archivo = str(row.get(col_guia_hecha, '')).strip()
                        
                        if nombre_archivo and nombre_archivo.lower() not in ['nan', 'none', '']:
                            archivos = descargar_guias_drive(drv, [nombre_archivo])
                            if archivos and len(archivos) > 0:
                                # Llamar a Vertex OCR
                                data_ia = procesar_guia_ia_vertex(archivos[0].read())
                                if data_ia and 'items' in data_ia:
                                    suma_peso = 0.0
                                    descripciones = []
                                    for item in data_ia['items']:
                                        peso_str = item.get('peso', '0')
                                        peso_float = float(formato_inteligente(limpiar_monto(peso_str)) or 0.0)
                                        suma_peso += peso_float
                                        
                                        desc = str(item.get('desc', '')).strip().upper()
                                        if desc:
                                            descripciones.append(desc)
                                    
                                    desc_final = " / ".join(set(descripciones)) if descripciones else "RESIDUOS SOLIDOS NO PELIGROSOS"
                                    
                                    st.session_state.sigersol_ia_cache[row_excel_idx] = {
                                        "cantidad": f"{suma_peso:.2f}",
                                        "descripcion": desc_final
                                    }
                        progreso.progress((i + 1) / total_docs)
                    
                    time.sleep(0.5)
                    progreso.empty()
                    st.success("✅ ¡Extracción completada! La tabla se ha actualizado.")
        else:
            st.error("No se encontró la columna de 'Guia hecha' para buscar los archivos.")
    
    st.markdown("---")
    
    # Convertimos a formato para el data_editor
    lista_editor = []
    
    # Obtener diccionarios para autocompletar RUC si es posible
    from src.services.google_service import obtener_datos_empresas_desde_sheets
    datos_empresas = obtener_datos_empresas_desde_sheets()
    
    for _, row in df_filtrado.iterrows():
        # Lógica de Guía: La información que va a SUNAT es directamente la de la Columna B
        guia_original = str(row.get(col_guia_original, ''))
        
        # Según la indicación, debemos dar como resultado la columna B
        guia_sunat = guia_original
        
        empresa_gen = str(row.get(col_empresa, ''))
        ruc_gen = datos_empresas.get(empresa_gen.upper(), {}).get('ruc', '')
        
        # Verificar si tenemos datos cacheados de la IA para esta fila
        row_idx = row['_excel_row_idx']
        cache_data = st.session_state.sigersol_ia_cache.get(row_idx, {})
        
        desc_default = cache_data.get("descripcion", "RESIDUOS SOLIDOS NO PELIGROSOS")
        cant_default = cache_data.get("cantidad", "0.00")
        
        lista_editor.append({
            "Migrar": False,
            "_excel_row_idx": row_idx,
            "FECHA": str(row.get(col_fecha, '')),
            "MES": str(row.get(col_mes, '')),
            "EMPRESA GENERADORA": empresa_gen,
            "RUC": ruc_gen,
            "DIRECCION DEL GENERADOR": str(row.get('Destinatario/Remitente', '')), # Origen/Dirección
            "N° GUIA (SUNAT)": guia_sunat,
            "DESCRIPCION DE RESIDUOS": desc_default,
            "CANTIDAD (Tn)": cant_default,
            "DESTINO FINAL": str(row.get('Fundo/Planta', ''))
        })
        
    df_editor = pd.DataFrame(lista_editor)
    
    st.markdown("### Selecciona las filas a migrar y completa los datos faltantes")
    st.caption("Asegúrate de llenar la Cantidad (Tn) y validar la Descripción de Residuos.")
    
    # Configuramos el editor
    edited_df = st.data_editor(
        df_editor,
        column_config={
            "Migrar": st.column_config.CheckboxColumn("Migrar", default=False),
            "_excel_row_idx": None, # Ocultamos la columna del ID
        },
        disabled=["_excel_row_idx", "FECHA", "MES", "N° GUIA (SUNAT)"], # Bloqueamos campos que no deberian alterarse
        use_container_width=True,
        hide_index=True
    )
    
    # --- 4. Acción de Migración ---
    filas_a_migrar = edited_df[edited_df["Migrar"] == True]
    
    if len(filas_a_migrar) > 0:
        st.info(f"Has seleccionado {len(filas_a_migrar)} filas para migrar a SUNAT.")
        
        # Nombre de la pestaña dinámico
        emp_clean = str(empresa_sel).strip().upper().replace(" ", "_") if empresa_sel != "Todas" else "VARIOS"
        if tipo_op == "Comercialización":
            nombre_pestana = f"COM-{emp_clean}"
        else:
            nombre_pestana = f"SERV-{emp_clean}"
            
        st.write(f"📂 Pestaña destino: **{nombre_pestana}**")
        
        if st.button("🚀 Migrar a Sigersol", type="primary", use_container_width=True):
            with st.spinner(f"Migrando {len(filas_a_migrar)} registros a SUNAT..."):
                exito_migracion = migrar_a_sunat(filas_a_migrar, nombre_pestana)
                
                if exito_migracion:
                    filas_idx = filas_a_migrar['_excel_row_idx'].tolist()
                    exito_update = actualizar_sigersol_origen(filas_idx)
                    
                    if exito_update:
                        st.success(f"✅ ¡Éxito! Se migraron las guías y se marcó la hoja de origen.")
                        st.balloons()
                        # Limpiar cache para forzar recarga
                        st.cache_data.clear()
                        # Usar st.rerun pero con delay si es necesario, aquí no lo haremos para que vean el success
                    else:
                        st.warning("Se migraron a SUNAT, pero falló la actualización del check en la hoja de origen.")
                else:
                    st.error("Error al migrar los datos a la hoja destino de SUNAT.")

