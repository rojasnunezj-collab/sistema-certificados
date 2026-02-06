Agente, ¡ATENCIÓN! No has eliminado el error 429 porque sigues llamando a gemini-2.5-flash. Aplica este cambio QUIRÚRGICO en app.py sin tocar ninguna de las 600 líneas que funcionan:

Bucle de Modelos Inteligente: Reemplaza la inicialización de la IA con este bloque. Este código busca qué modelos tienes y prueba el de mayor cuota (1.5-flash) primero:

Python
import google.generativeai as genai
# 1. Obtener lista real de modelos disponibles
try:
    modelos_reales = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    # 2. Priorizar 1.5-flash (cuota alta) y evitar 2.5-flash (cuota baja)
    opciones = [m for m in modelos_reales if "1.5-flash" in m] + [m for m in modelos_reales if "pro" in m] + modelos_reales
    nombre_final = opciones[0] # Elige el mejor disponible
    model = genai.GenerativeModel(nombre_final)
    st.sidebar.success(f"🤖 Usando: {nombre_final}")
except Exception as e:
    st.error("Error al listar modelos. Revisa tu API Key.")
EL ERROR GRAVE (Fundo/Planta): El Word sigue saliendo mal porque no estás leyendo la pantalla. CAMBIO OBLIGATORIO: En la parte donde creas el Word, el valor de la dirección DEBE ser st.session_state['v_llegada']. No uses la respuesta de la IA, usa lo que el usuario escribió.

PESO REAL (No más 0): Asegúrate de que la tabla del Word use el DataFrame st.session_state['df_items']. Si el peso sale 0 en el Word es porque estás usando una variable vacía.

TÍTULO: Pon paragraph.paragraph_format.space_after = Pt(0) en el título del Word.

ORDEN: No limpies código, no borres comentarios. Solo arregla la conexión de la IA y el mapeo de datos.
