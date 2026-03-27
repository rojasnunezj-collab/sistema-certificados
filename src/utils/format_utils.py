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
