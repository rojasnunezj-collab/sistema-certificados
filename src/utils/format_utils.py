# ====================================================================
# --- BLOQUE 0: Imports ---
# ====================================================================
import re
from datetime import datetime, timedelta

# ====================================================================
# --- BLOQUE 1: Funciones de Limpieza Numérica y Formato Monetario ---
# ====================================================================
def limpiar_monto(valor):
    """
    Convierte string a float.
    Maneja formato europeo/latino intercambiando comas por puntos.
    """
    if not valor: return 0.0
    s = str(valor).strip()
    
    s = s.replace(',', '.')
    
    if s.count('.') > 1:
        parts = s.split('.')
        s = "".join(parts[:-1]) + '.' + parts[-1]
    
    s = re.sub(r'[^\d.]', '', s)
    try:
        return float(s)
    except:
        return 0.0

def formato_inteligente(valor):
    """
    Formatea números: 100.0 -> "100", 3580.50 -> "3580.5"
    """
    try:
        f = float(valor)
        if f.is_integer():
            return f"{int(f)}"
        else:
            return f"{f}"
    except:
        return str(valor)

# ====================================================================
# --- BLOQUE 2: Operaciones con Fechas y Formato de Textos ---
# ====================================================================
def obtener_fin_de_mes(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, "%d/%m/%Y")
        next_month = dt.replace(day=28) + timedelta(days=4)
        res = next_month - timedelta(days=next_month.day)
        return res.strftime("%d/%m/%Y")
    except: return fecha_str

def limpiar_descripcion(texto):
    if not texto: return ""
    return re.sub(r'VEN\s*-\s*AMB\s*-\s*', '', str(texto).strip(), flags=re.IGNORECASE).strip()

def formato_nompropio(texto):
    return str(texto).strip().title() if texto else ""

def normalizar_fecha(fecha_str):
    if not fecha_str: return datetime.now().strftime("%d/%m/%Y")
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
        try: return datetime.strptime(fecha_str.strip(), fmt).strftime("%d/%m/%Y")
        except: continue
    return fecha_str 

def formatear_guia(serie_str):
    if not serie_str or '-' not in str(serie_str): return serie_str
    try:
        p = str(serie_str).split('-')
        if len(p) == 2: return f"{p[0].strip()}-{str(int(p[1].strip()))}"
    except: pass
    return serie_str


# ====================================================================
# --- BLOQUE 3: Detección y Validación de Placa vs Guía ---
# ====================================================================
def es_formato_guia(valor):
    """
    Determina si un texto corresponde al patrón característico de una Guía de Remisión (SUNAT).
    Ejemplos típicos:
      - GRE: T001-00012345, EG01-1234, T002-1, E001-12, V001-123, GR01-456
      - Físicas: 001-0012345, 0001-1234, 002-123456
    """
    if not valor:
        return False
    s = str(valor).strip().upper()
    if s in ['', 'NONE', 'NAN', 'S/N', 'SIN DATOS']:
        return False

    if '-' in s:
        partes = s.split('-', 1)
        serie = partes[0].strip()
        correlativo = partes[1].strip()

        # Series típicas de Guía Electrónica (GRE)
        if re.match(r'^(T\d{2,3}|EG\d{2}|E\d{3}|V\d{3}|GR\d{2}|[A-Z]\d{3})$', serie) and correlativo.isdigit():
            return True

        # Series físicas numéricas (ej. 001-0012345, 0001-1234)
        if re.match(r'^\d{3,4}$', serie) and correlativo.isdigit():
            return True

        # Series alfanuméricas con correlativo numérico extendido
        if re.match(r'^[A-Z0-9]{3,4}$', serie) and correlativo.isdigit() and len(correlativo) >= 4:
            return True

    return False


def es_formato_placa(valor):
    """
    Determina si un texto corresponde al patrón característico de una Placa vehicular (SUNARP/MTC).
    Ejemplos típicos:
      - ABC-123, P2M-843, F4X-901, C8A-712, D6U-819
      - ABC123, P2M843
      - Remolques / Especiales: TC-1234, Z1-1234
    """
    if not valor:
        return False
    s = str(valor).strip().upper()
    if s in ['', 'NONE', 'NAN', 'S/N', 'SIN DATOS']:
        return False

    s_limpio = s.replace('-', '').replace(' ', '')

    # La longitud habitual de placa peruana es de 5 o 6 caracteres sin guion
    if len(s_limpio) not in [5, 6]:
        return False

    # Descartar series típicas de guías electrónicas
    if re.match(r'^(T\d{3}|EG\d{2}|E\d{3}|V\d{3}|GR\d{2})$', s_limpio[:4]):
        return False

    # Formato estándar de placa (combinación de letras y números)
    if re.match(r'^[A-Z0-9]{3}[A-Z0-9]{3}$', s_limpio):
        tiene_letras = bool(re.search(r'[A-Z]', s_limpio))
        tiene_numeros = bool(re.search(r'\d', s_limpio))
        if tiene_letras and tiene_numeros:
            return True

    # Formato de remolque: 2 letras + 4 números (ej. TC1234)
    if re.match(r'^[A-Z]{2}\d{4}$', s_limpio):
        return True

    return False


def verificar_consistencia_guia_placa(guia, placa):
    """
    Evalúa la coherencia del par Guía y Placa.
    Retorna una tupla: (estado, mensaje_detalle)
    Estados posibles:
      - 'OK': Datos correctos o no concluyentes (sin evidencia de error).
      - 'INVERTIDOS': La Guía tiene formato de Placa Y la Placa tiene formato de Guía.
      - 'PLACA_COMO_GUIA': El campo Guía contiene una Placa vehicular.
      - 'GUIA_COMO_PLACA': El campo Placa contiene una Guía de remisión.
    """
    guia_str = str(guia).strip().upper() if guia else ""
    placa_str = str(placa).strip().upper() if placa else ""

    if not guia_str and not placa_str:
        return 'OK', ""

    es_g_guia = es_formato_guia(guia_str)
    es_g_placa = es_formato_placa(guia_str)

    es_p_guia = es_formato_guia(placa_str)
    es_p_placa = es_formato_placa(placa_str)

    # Caso 1: Ambos están invertidos
    if es_g_placa and es_p_guia:
        return 'INVERTIDOS', f"Campos invertidos: 'Guía' contiene '{guia_str}' (formato de Placa) y 'Placa' contiene '{placa_str}' (formato de Guía)."

    # Caso 2: Se colocó una placa en la guía
    if es_g_placa and not es_g_guia:
        return 'PLACA_COMO_GUIA', f"El campo 'Guía' contiene '{guia_str}', el cual tiene formato de Placa vehicular."

    # Caso 3: Se colocó una guía en la placa
    if es_p_guia and not es_p_placa:
        return 'GUIA_COMO_PLACA', f"El campo 'Placa' contiene '{placa_str}', el cual tiene formato de Guía de Remisión."

    return 'OK', ""


def validar_inconsistencias_df(df_items):
    """
    Inspecciona todas las filas del DataFrame buscando inconsistencias entre 'guia_origen' y 'placa_origen'.
    Retorna una lista de diccionarios con los errores detectados por fila.
    """
    errores = []
    if df_items is None or df_items.empty:
        return errores

    for idx, row in df_items.iterrows():
        g = str(row.get('guia_origen', '')).strip()
        p = str(row.get('placa_origen', '')).strip()
        estado, detalle = verificar_consistencia_guia_placa(g, p)
        if estado != 'OK':
            errores.append({
                'fila': idx + 1,
                'guia': g,
                'placa': p,
                'tipo': estado,
                'detalle': detalle
            })
    return errores

