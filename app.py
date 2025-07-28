import streamlit as st
import json
import requests # Importar la librería requests

# Título de la aplicación
st.set_page_config(page_title="Asistente de Fórmulas Excel con IA", layout="centered")

st.title("📊 Asistente de Fórmulas Excel con IA")
st.markdown("¡Hola! Soy tu asistente personal para Excel. Describe tu problema y te ayudaré a encontrar la fórmula y la estructura que necesitas.")

# Área de texto para que el usuario ingrese su problema
user_problem = st.text_area(
    "Describe tu problema en Excel (ej: 'Necesito sumar los valores de la columna B si la columna A contiene \"Ventas\" y la columna C es mayor a 100.')",
    height=150,
    placeholder="Ejemplo: Quiero encontrar el valor máximo en la columna D para las filas donde la columna E sea 'Activo'."
)

# Botón para enviar la consulta a la IA
if st.button("Obtener Solución de Excel"):
    if user_problem:
        st.info("Generando solución... Por favor, espera.")

        # Construir el prompt para la IA
        prompt = f"""
        Eres un experto en Excel y en la creación de fórmulas. El usuario te proporcionará una descripción de un problema o una necesidad en Excel.
        Tu tarea es proporcionar una solución completa que incluya:
        1.  **Estructuración de la Solución**: Una breve explicación de cómo abordar el problema en Excel (ej: qué columnas usar, si se necesita una tabla auxiliar, etc.).
        2.  **Fórmula de Excel**: La fórmula o fórmulas exactas que el usuario puede copiar y pegar, con ejemplos claros de rangos (ej: A1:A10, B:B).
        3.  **Explicación de la Fórmula**: Una descripción detallada de cada parte de la fórmula y cómo funciona.

        Asegúrate de que la respuesta sea clara, concisa y directamente aplicable.
        Si es necesario, puedes sugerir el uso de tablas o rangos con nombre para mejorar la legibilidad.

        Problema del usuario:
        "{user_problem}"

        Formato de respuesta deseado:
        **Estructuración de la Solución:**
        [Tu explicación de la estructura]

        **Fórmula de Excel:**
        ```excel
        [Tu fórmula aquí]
        ```

        **Explicación de la Fórmula:**
        [Tu explicación detallada de la fórmula]
        """

        try:
            # Llamada a la API de Gemini para generar la respuesta
            chatHistory = []
            chatHistory.append({"role": "user", "parts": [{"text": prompt}]}) # Usar .append para listas de Python
            payload = {"contents": chatHistory}
            # La clave API se proporciona en tiempo de ejecución por el entorno de Canvas.
            # El error "PERMISSION_DENIED" sugiere un problema con la configuración del entorno
            # o los permisos de la clave API, no con el código en sí.
            apiKey = ""
            apiUrl = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={apiKey}"

            # Realizar la llamada HTTP síncrona usando requests
            response = requests.post(
                apiUrl,
                headers={'Content-Type': 'application/json'},
                data=json.dumps(payload)
            )
            result = response.json() # Obtener el JSON de la respuesta

            # Verificar si la respuesta es exitosa y contiene contenido
            if result.get("candidates") and len(result["candidates"]) > 0 and \
               result["candidates"][0].get("content") and result["candidates"][0]["content"].get("parts") and \
               len(result["candidates"][0]["content"]["parts"]) > 0:
                ai_response_text = result["candidates"][0]["content"]["parts"][0]["text"]
                st.subheader("💡 Solución Propuesta:")
                st.markdown(ai_response_text)
            else:
                st.error("Lo siento, no pude generar una solución en este momento. Por favor, intenta de nuevo.")
                st.json(result) # Para depuración, mostrar la respuesta completa si falla

        except Exception as e:
            st.error(f"Ocurrió un error al comunicarse con la IA: {e}")
            st.warning("Asegúrate de que tu conexión a internet sea estable y que la API de Gemini esté accesible.")
    else:
        st.warning("Por favor, describe tu problema en Excel antes de enviar.")

st.markdown("---")
st.markdown("Este asistente utiliza inteligencia artificial para ayudarte con tus tareas de Excel.")
