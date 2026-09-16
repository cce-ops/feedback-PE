import os
import time
import streamlit as st

# -----------------------------------------------------------------------------
# IMPORTACIÓN DE LIBRERÍAS DE LECTURA DE ARCHIVOS
# -----------------------------------------------------------------------------
try:
    import pypdf
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

try:
    import docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# SDK Oficial de Google GenAI
try:
    from google import genai
    from google.genai import types
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y DIRECTORIOS
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Plataforma Integrada de Feedback Académico",
    page_icon="🎓",
    layout="wide"
)

CARPETA_MATERIALES = "materiales"
FORMATOS_SOPORTADOS = ["pdf", "pptx", "xlsx", "xls", "csv", "docx", "txt", "py", "m"]

# -----------------------------------------------------------------------------
# FUNCIONES PARA EXTRACCIÓN DE TEXTO SEGÚN EL TIPO DE ARCHIVO
# -----------------------------------------------------------------------------
def procesar_archivo_path_o_bytes(file_obj, filename_hint=""):
    """
    Procesa un archivo indicando la ruta en disco (string) o un objeto de subida (UploadedFile/BytesIO).
    Soporta documentos, hojas de cálculo, presentaciones y archivos de código (Python, MATLAB).
    """
    nombre = (filename_hint if filename_hint else getattr(file_obj, "name", str(file_obj))).lower()
    
    if nombre.endswith(".pdf"):
        if not HAS_PYPDF:
            return "Error: Falta la librería 'pypdf'. Ejecuta 'pip install pypdf'."
        try:
            reader = pypdf.PdfReader(file_obj)
            texto = ""
            for i, page in enumerate(reader.pages):
                txt = page.extract_text()
                if txt:
                    texto += f"\n--- PÁGINA {i+1} ---\n" + txt
            return texto
        except Exception as e:
            return f"Error al leer PDF: {str(e)}"

    elif nombre.endswith(".pptx"):
        if not HAS_PPTX:
            return "Error: Falta la librería 'python-pptx'. Ejecuta 'pip install python-pptx'."
        try:
            prs = Presentation(file_obj)
            texto = ""
            for i, slide in enumerate(prs.slides):
                texto += f"\n--- DIAPOSITIVA {i+1} ---\n"
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        texto += shape.text + "\n"
            return texto
        except Exception as e:
            return f"Error al leer PowerPoint: {str(e)}"

    elif nombre.endswith((".xlsx", ".xls")):
        if not HAS_PANDAS:
            return "Error: Falta la librería 'pandas' o 'openpyxl'. Ejecuta 'pip install pandas openpyxl'."
        try:
            excel_file = pd.ExcelFile(file_obj)
            texto = ""
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                texto += f"\n--- HOJA DE CÁLCULO: {sheet_name} ---\n"
                texto += df.to_string(index=False) + "\n"
            return texto
        except Exception as e:
            return f"Error al leer archivo de Excel: {str(e)}"

    elif nombre.endswith(".csv"):
        if not HAS_PANDAS:
            return "Error: Falta la librería 'pandas'. Ejecuta 'pip install pandas'."
        try:
            df = pd.read_csv(file_obj)
            return "--- ARCHIVO CSV ---\n" + df.to_string(index=False)
        except Exception as e:
            return f"Error al leer CSV: {str(e)}"

    elif nombre.endswith(".docx"):
        if not HAS_DOCX:
            return "Error: Falta la librería 'python-docx'. Ejecuta 'pip install python-docx'."
        try:
            doc = docx.Document(file_obj)
            return "\n".join([p.text for p in doc.paragraphs if p.text])
        except Exception as e:
            return f"Error al leer Word: {str(e)}"

    elif nombre.endswith(".py"):
        try:
            if isinstance(file_obj, str):
                with open(file_obj, "r", encoding="utf-8") as f:
                    contenido = f.read()
            else:
                contenido = file_obj.read().decode("utf-8")
            return f"--- CÓDIGO FUENTE PYTHON ({nombre}) ---\n" + contenido
        except Exception as e:
            return f"Error al leer código Python: {str(e)}"

    elif nombre.endswith(".m"):
        try:
            if isinstance(file_obj, str):
                with open(file_obj, "r", encoding="utf-8", errors="ignore") as f:
                    contenido = f.read()
            else:
                contenido = file_obj.read().decode("utf-8", errors="ignore")
            return f"--- CÓDIGO FUENTE MATLAB ({nombre}) ---\n" + contenido
        except Exception as e:
            return f"Error al leer script de MATLAB: {str(e)}"

    elif nombre.endswith(".txt"):
        try:
            if isinstance(file_obj, str):
                with open(file_obj, "r", encoding="utf-8") as f:
                    return f.read()
            else:
                return file_obj.read().decode("utf-8")
        except Exception as e:
            return f"Error al leer TXT: {str(e)}"

    return "Formato de archivo no soportado."

def cargar_materiales_predeterminados():
    """Escanea la carpeta local 'materiales' y extrae el texto automáticamente al iniciar la app."""
    base_conocimiento = {}
    if os.path.exists(CARPETA_MATERIALES):
        archivos = os.listdir(CARPETA_MATERIALES)
        for nombre in archivos:
            extension = nombre.split(".")[-1].lower()
            if extension in FORMATOS_SOPORTADOS:
                ruta_completa = os.path.join(CARPETA_MATERIALES, nombre)
                texto = procesar_archivo_path_o_bytes(ruta_completa, filename_hint=nombre)
                base_conocimiento[nombre] = texto
    return base_conocimiento

# -----------------------------------------------------------------------------
# INICIALIZACIÓN DE ESTADO (SESSION STATE)
# -----------------------------------------------------------------------------
if "base_conocimiento" not in st.session_state:
    st.session_state.base_conocimiento = cargar_materiales_predeterminados()

if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "assistant", "content": "¡Hola! Soy tu tutor orientador. Dime en qué sección, script de Python/MATLAB o archivo de tu entrega tienes dudas y te indicaré de forma directa qué debes revisar y cómo corregirlo."}
    ]

def obtener_contexto_asignatura():
    """Compila todo el conocimiento cargado en un bloque de texto formateado."""
    if not st.session_state.base_conocimiento:
        return "No hay materiales subidos por el profesor. Se evaluará con criterios generales."
    
    contexto = "=== MATERIALES Y CONOCIMIENTO OFICIAL DE LA ASIGNATURA ===\n\n"
    for nombre_file, contenido in st.session_state.base_conocimiento.items():
        contexto += f"--- ARCHIVO FUENTE: {nombre_file} ---\n{contenido}\n\n"
    return contexto

def consultar_gemini(prompt_sistema, prompt_usuario, api_key, modelo_seleccionado):
    """Realiza la llamada a la API de Gemini con el modelo especificado."""
    if not HAS_GEMINI:
        return "Error: El paquete 'google-genai' no está instalado. Ejecuta 'pip install google-genai'."
    if not api_key:
        return "⚠️ Por favor, introduce tu API Key de Gemini en el menú lateral."
    
    try:
        client = genai.Client(api_key=api_key)
        contenido_completo = f"{prompt_sistema}\n\nENTRADA DEL USUARIO:\n{prompt_usuario}"
        
        response = client.models.generate_content(
            model=modelo_seleccionado,
            contents=contenido_completo,
            config=types.GenerateContentConfig(temperature=0.3)
        )
        return response.text
    except Exception as e:
        error_str = str(e)
        if "429" in error_str or "RESOURCE_EXHAUSTED" in error_str:
            return f"⚠️ El modelo **{modelo_seleccionado}** está saturado o has alcanzado el límite de cuota. Prueba seleccionando un modelo alternativo como **gemini-3.5-flash-lite** o **gemini-3.1-flash-lite** en el menú lateral."
        return f"Error al comunicarse con la API de Gemini ({modelo_seleccionado}): {error_str}"

# -----------------------------------------------------------------------------
# BARRA LATERAL: CONFIGURACIÓN, MODELOS Y TUTOR ORIENTADOR
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Configuración del Motor")
    api_key_env = os.environ.get("GEMINI_API_KEY", "")
    api_key = st.text_input("Gemini API Key:", value=api_key_env, type="password", help="Obtén tu clave en Google AI Studio.")

    # SELECTOR DE MODELOS
    modelo_seleccionado = st.selectbox(
        "Selecciona el Modelo de IA:",
        options=[
            "gemini-3.6-flash",
            "gemini-3.5-flash",
            "gemini-3.5-flash-lite",
            "gemini-3.1-flash-lite"
        ],
        index=0,
        help="Si un modelo reporta saturación, selecciona una versión Flash-Lite para mayor capacidad y rapidez de respuesta."
    )

    st.markdown("---")
    st.header("🤖 Tutor Orientador")
    st.caption("Resuelve tus dudas sobre código, teoría o ejercicios prácticos.")

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.write(msg["content"])

    if user_prompt := st.chat_input("Plantea tu duda sobre la asignatura/tarea..."):
        st.session_state.messages.append({"role": "user", "content": user_prompt})
        with st.chat_message("user"):
            st.write(user_prompt)

        with st.chat_message("assistant"):
            with st.spinner("Analizando tu consulta con los materiales..."):
                contexto_mat = obtener_contexto_asignatura()
                
                system_prompt_orientador = f"""
                Eres un tutor docente universitario experto, claro, directo y didáctico.
                Tu objetivo es señalar con precisión los errores u omisiones del estudiante en sus memorias o scripts de código (Python/MATLAB) y explicarle cómo corregirlos.

                REGLAS DE ACTUACIÓN:
                1. SÉ DIRECTO Y ESPECÍFICO: Señala claramente qué concepto, fórmula, script, función o parámetro está equivocado o falta.
                2. EXPLICA EL POR QUÉ: Basándote en el conocimiento de la asignatura, explica por qué está mal o incompleto.
                3. GUÍA SIN REGALAR LA SOLUCIÓN FINAL: Explica la metodología o los pasos exactos para solucionarlo, pero deja que el alumno redacte el texto final o escriba las líneas de código finales.
                
                CONOCIMIENTO DE LA ASIGNATURA DISPONIBLE:
                {contexto_mat}
                """
                
                if api_key:
                    respuesta_bot = consultar_gemini(system_prompt_orientador, user_prompt, api_key, modelo_seleccionado)
                else:
                    respuesta_bot = f"🤖 [Modo Simulación - {modelo_seleccionado}]: Respecto a tu consulta '{user_prompt[:30]}...', según el tema cargado en la asignatura, debes revisar la estructura del código ya que omite la definición de parámetros requerida."

                st.write(respuesta_bot)
                st.session_state.messages.append({"role": "assistant", "content": respuesta_bot})

# -----------------------------------------------------------------------------
# CUERPO PRINCIPAL
# -----------------------------------------------------------------------------
st.title("🎓 Plataforma Integrada de Retroalimentación Académica")
st.markdown("Evaluación automática basada en los materiales oficiales de la asignatura.")

tab_rubrica, tab_evaluador, tab_profesor = st.tabs([
    "📋 Paso 1: Lista de Cotejo / Rúbrica",
    "📤 Paso 2: Entrega y Feedback Automático",
    "📚 Materiales del Curso (Consulta)"
])

# -----------------------------------------------------------------------------
# TAB 1: RÚBRICA INTERACTIVA
# -----------------------------------------------------------------------------
with tab_rubrica:
    st.subheader("Verificación Previa a la Entrega")
    st.write("Confirma la revisión de los puntos clave antes del envío final:")

    col1, col2 = st.columns(2)
    with col1:
        req1 = st.checkbox("Estructura adecuada según el tipo de entrega (Informe, Excel, PPTX o Código Python/MATLAB).")
        req2 = st.checkbox("Uso correcto de datos, funciones, algoritmos o terminología explicada en clase.")
        req3 = st.checkbox("Citas bibliográficas o fuentes de datos/código correctamente referenciadas.")
    with col2:
        req4 = st.checkbox("Cumplimiento de los requisitos de la guía docente o enunciado del ejercicio.")
        req5 = st.checkbox("Revisión de ortografía, comentarios en el código, etiquetas o formato general.")

    if req1 and req2 and req3 and req4 and req5:
        st.success("✅ ¡Lista de verificación completada! Procede a la entrega de tus archivos en el Paso 2.")

# -----------------------------------------------------------------------------
# TAB 2: EVALUADOR DE ENTREGABLES (SOPORTE MULTIARCHIVOS, CÓDIGO Y BULLET POINTS)
# -----------------------------------------------------------------------------
with tab_evaluador:
    st.subheader("Envío de Tareas / Archivos del Proyecto")
    st.write("Puedes subir **varios archivos a la vez** (ej. memoria PDF, código en Python `.py`, scripts de MATLAB `.m`, tablas Excel y presentación PPTX).")
    
    if st.session_state.base_conocimiento:
        st.caption(f"📚 **Materiales del curso precargados activos:** {len(st.session_state.base_conocimiento)} archivo(s)")
    else:
        st.warning("⚠️ No se encontraron materiales en la carpeta 'materiales'. La evaluación usará criterios generales.")

    archivos_alumno = st.file_uploader(
        "Sube tus entregables (PDF, PPTX, Excel, CSV, Word, Python .py, MATLAB .m o TXT):", 
        type=FORMATOS_SOPORTADOS,
        accept_multiple_files=True,
        key="alumno_uploader"
    )
    texto_manual = st.text_area("O añade comentarios/explicaciones adicionales para la entrega:", height=100)

    if st.button("🚀 Analizar Tareas y Generar Feedback", type="primary"):
        contenido_estudiante = ""
        nombres_archivos = []
        
        if archivos_alumno:
            for arch in archivos_alumno:
                txt_arch = procesar_archivo_path_o_bytes(arch)
                contenido_estudiante += f"\n=== ENTREGABLE DEL ALUMNO: {arch.name} ===\n" + txt_arch + "\n\n"
                nombres_archivos.append(arch.name)

        if texto_manual.strip():
            contenido_estudiante += "\n=== NOTAS/TEXTO ADICIONAL DEL ALUMNO ===\n" + texto_manual.strip()

        if not contenido_estudiante.strip():
            st.error("Por favor, sube al menos un archivo o escribe comentarios antes de analizar.")
        else:
            with st.spinner(f"Analizando entregables con {modelo_seleccionado}..."):
                contexto_docente = obtener_contexto_asignatura()
                
                system_prompt_evaluador = f"""
                Eres un evaluador académico experto para esta asignatura.
                Tu tarea es analizar el conjunto de entregables presentados por el estudiante (que pueden incluir documentos, diapositivas, hojas de cálculo o código fuente en Python/MATLAB) y generar un informe de feedback detallado y estructurado.
                
                MATERIALES DE REFERENCIA DE LA ASIGNATURA:
                {contexto_docente}
                
                INSTRUCCIONES DE EVALUACIÓN:
                1. Revisa de forma conjunta todos los entregables subidos por el estudiante.
                2. Compara el trabajo entregado (documentos y/o código) directamente con los conceptos, metodologías y soluciones de los materiales oficiales de la asignatura.
                3. REGLA DE ORO DE FIABILIDAD: Evalúa el trabajo del alumno basándote ÚNICAMENTE en la información explícita de los MATERIALES DE REFERENCIA. Si el trabajo aborda un tema o concepto que NO está en los materiales, indica textualmente: 'Este punto no puede ser verificado con los materiales actuales del curso'.
                4. Estructura el feedback en formato Markdown obligatoriamente con las siguientes secciones:
                   - **Resumen Ejecutivo:** (Valoración general en 2 o 3 frases).
                   - **Archivos Evaluados:** (Lista de entregables analizados).
                   - **Calificación Global Estimada:** (Nota de 0 a 10 con justificación).
                   - **Fortalezas Destacadas:** (Puntos fuertes identificados en documentación o código).
                   - **Análisis Detallado por Criterios:** (Evaluación punto por punto de teoría y código si aplica).
                   - **📌 RESUMEN DE CORRECCIONES Y AÑADIDOS NECESARIOS:** (Un listado exclusivo en BULLET POINTS (* o -) estructurado con todo lo que el estudiante debe corregir o añadir obligatoriamente en su informe o código antes de la entrega final).
                """
                
                if api_key:
                    resultado_feedback = consultar_gemini(system_prompt_evaluador, contenido_estudiante, api_key, modelo_seleccionado)
                else:
                    resultado_feedback = f"""
### 📊 Resumen Ejecutivo
Se han analizado los **{len(nombres_archivos)} entregables** subidos por el estudiante frente a la base de conocimiento oficial del curso.

### 📁 Archivos Evaluados
{chr(10).join([f'- `{nombre}`' for nombre in nombres_archivos]) if nombres_archivos else '- Texto introducido manualmente'}

### 🔢 Calificación Global Estimada
**8.2 / 10**

### ✅ Fortalezas Destacadas
- **Alineación con el temario:** El contenido y código demuestran comprensión de los conceptos clave integrados en la asignatura.
- **Variedad de Formatos:** Presentación clara combinando informe y scripts de programación.

### 📋 Análisis Detallado por Criterios
- **Metodología y Código:** Correcta aplicación general, aunque falta comentar las funciones principales en el script.
- **Precisión Numérica/Algorítmica:** Los resultados calculados coinciden con los datos planteados en el enunciado.

---

### 📌 RESUMEN DE CORRECCIONES Y AÑADIDOS NECESARIOS
*   **Corregir el script (`.py`/`.m`):** Ajustar los valores de los parámetros según lo especificado en el Tema 3.
*   **Añadir documentación:** Incluir comentarios descriptivos en las funciones principales del código.
*   **Formato en informe:** Etiquetar correctamente las figuras y gráficas generadas antes de la entrega final.

*(Nota: Añade tu Gemini API Key en el menú lateral para obtener la evaluación generada en tiempo real).*
                    """

            st.markdown("---")
            st.markdown("## 📊 Informe de Retroalimentación Automática")
            st.markdown(resultado_feedback)

# -----------------------------------------------------------------------------
# TAB 3: VISOR DE MATERIALES DE LA ASIGNATURA
# -----------------------------------------------------------------------------
with tab_profesor:
    st.subheader("📚 Materiales Oficiales del Curso")
    st.write("Los siguientes archivos provienen de la carpeta `/materiales` del repositorio de la asignatura y son utilizados por la IA como fuente oficial de conocimiento.")

    if st.session_state.base_conocimiento:
        for doc_name, text_content in st.session_state.base_conocimiento.items():
            with st.expander(f"📄 Documento / Código: {doc_name}"):
                st.caption(f"Caracteres procesados: {len(text_content)}")
                st.text(text_content[:600] + ("..." if len(text_content) > 600 else ""))
        
        if st.button("🔄 Recargar materiales desde la carpeta '/materiales'"):
            st.session_state.base_conocimiento = cargar_materiales_predeterminados()
            st.rerun()
    else:
        st.info("No hay materiales cargados en la carpeta '/materiales'. Para añadir apuntes o guías de código, súbelos directamente a la carpeta 'materiales/' en tu repositorio de GitHub.")