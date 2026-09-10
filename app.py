import streamlit as st
from docx import Document
from io import BytesIO
import math
import PyPDF2
import json
import re
from huggingface_hub import InferenceClient

# --- CONFIGURACIÓN DE PÁGINA Y ESTADOS ---
st.set_page_config(page_title="Agente Sustanciador - RPM (Open Source)", layout="wide")

if 'ibl' not in st.session_state: st.session_state.ibl = 1500000.0
if 'smmlv' not in st.session_state: st.session_state.smmlv = 1300000.0
if 'semanas' not in st.session_state: st.session_state.semanas = 1300
if 'resumen_peticion' not in st.session_state: st.session_state.resumen_peticion = ""

st.title("⚖️ Agente Sustanciador Inteligente - RPM")
st.markdown("Carga la petición del ciudadano. El agente (potenciado por IA Open Source) extraerá las variables y proyectará la motivación jurídica.")

# --- FUNCIONES DE LECTURA DE ARCHIVOS ---
def leer_archivo(archivo):
    texto = ""
    if archivo.name.endswith('.pdf'):
        lector = PyPDF2.PdfReader(archivo)
        for pagina in lector.pages:
            texto += pagina.extract_text() + "\n"
    elif archivo.name.endswith('.docx'):
        doc = Document(archivo)
        for parrafo in doc.paragraphs:
            texto += parrafo.text + "\n"
    elif archivo.name.endswith('.txt'):
        texto = archivo.read().decode('utf-8')
    return texto

# --- FUNCIÓN DE INTERPRETACIÓN CON IA OPEN SOURCE (Hugging Face) ---
def analizar_texto_peticion(texto, api_key):
    # Utilizamos Mixtral-8x7B, un modelo de código abierto de alto rendimiento
    cliente = InferenceClient(model="mistralai/Mixtral-8x7B-Instruct-v0.1", token=api_key)
    
    # El prompt usa las etiquetas [INST] recomendadas para modelos Mistral/Mixtral
    prompt = f"""[INST] Eres un sustanciador experto en pensiones. Lee la siguiente petición de un ciudadano.
    Extrae estrictamente los siguientes datos en formato JSON. Si no encuentras un dato exacto, estima el más lógico según el texto. 
    Tu respuesta debe ser ÚNICA Y EXCLUSIVAMENTE el JSON, sin texto de introducción ni conclusiones.
    
    Formato JSON requerido:
    {{
        "semanas_cotizadas": <int>,
        "ingreso_base_liquidacion_ibl": <float>,
        "smmlv_ano_causacion": <float>,
        "resumen_hechos": "<Breve resumen de lo que pide en 2 lineas>"
    }}
    
    Petición del ciudadano:
    {texto}
    [/INST]"""
    
    # Generar la respuesta limitando la temperatura para respuestas más precisas y predecibles
    respuesta = cliente.text_generation(prompt, max_new_tokens=500, temperature=0.1)
    
    # Expresión regular para "atrapar" solo el JSON, en caso de que el modelo hable texto adicional
    match = re.search(r'\{.*\}', respuesta, re.DOTALL)
    if match:
        texto_json = match.group(0)
        return json.loads(texto_json)
    else:
        raise ValueError("El modelo no devolvió un formato JSON válido.")

# --- FUNCIONES DE CÁLCULO ---
def calcular_tasa_reemplazo(ibl, smmlv, semanas_totales):
    s = ibl / smmlv
    porcentaje_base = max(65.50 - (0.50 * s), 55.5)
    semanas_adicionales = max(0, semanas_totales - 1300)
    grupos_de_50 = math.floor(semanas_adicionales / 50)
    puntos_adicionales = min(grupos_de_50 * 1.5, 15.0)
    tasa_final = min(porcentaje_base + puntos_adicionales, 80.0)
    return s, porcentaje_base, grupos_de_50, puntos_adicionales, tasa_final

# --- FUNCIÓN DE GENERACIÓN DE WORD ---
def generar_word(texto_motivacion):
    doc = Document()
    doc.add_heading('MOTIVACIÓN DEL ACTO ADMINISTRATIVO', 0)
    for parrafo in texto_motivacion.split('\n'):
        if parrafo.strip():
            doc.add_paragraph(parrafo.strip())
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ DE USUARIO ---
st.sidebar.header("⚙️ Configuración del Agente")
api_key = st.sidebar.text_input("Hugging Face API Token:", type="password", help="Genera un token gratuito en huggingface.co/settings/tokens")

# 1. ZONA DE CARGA DE DOCUMENTOS
st.subheader("1. Análisis de Petición")
archivo_cargado = st.file_uploader("Adjuntar solicitud del peticionario (PDF, Word, TXT)", type=["pdf", "docx", "txt"])

if archivo_cargado is not None:
    if st.button("🧠 Leer e Interpretar Petición"):
        if not api_key:
            st.warning("⚠️ Ingresa el Token de Hugging Face en el menú lateral.")
        else:
            with st.spinner('El agente está leyendo el documento con IA Open Source...'):
                try:
                    texto_peticion = leer_archivo(archivo_cargado)
                    datos_extraidos = analizar_texto_peticion(texto_peticion, api_key)
                    
                    st.session_state.ibl = float(datos_extraidos.get("ingreso_base_liquidacion_ibl", 1500000.0))
                    st.session_state.smmlv = float(datos_extraidos.get("smmlv_ano_causacion", 1300000.0))
                    st.session_state.semanas = int(datos_extraidos.get("semanas_cotizadas", 1300))
                    st.session_state.resumen_peticion = datos_extraidos.get("resumen_hechos", "")
                    
                    st.success("Lectura completada. Datos precargados en el formulario.")
                    if st.session_state.resumen_peticion:
                        st.info(f"**Resumen interpretado:** {st.session_state.resumen_peticion}")
                        
                except Exception as e:
                    st.error(f"Error al procesar el documento. Asegúrate de que el Token sea correcto. Detalles: {e}")

st.divider()

# 2. ZONA DE EDICIÓN Y LIQUIDACIÓN
st.subheader("2. Variables de Liquidación (Editables)")
col1, col2, col3 = st.columns(3)
with col1:
    ibl_input = st.number_input("Ingreso Base de Liquidación (IBL):", min_value=0.0, value=st.session_state.ibl, step=100000.0)
with col2:
    smmlv_input = st.number_input("SMMLV del año de causación:", min_value=0.0, value=st.session_state.smmlv, step=10000.0)
with col3:
    semanas_input = st.number_input("Total Semanas Cotizadas:", min_value=0, value=st.session_state.semanas, step=1)

st.divider()

# 3. ZONA DE GENERACIÓN DE MOTIVACIÓN
st.subheader("3. Generación de Acto Administrativo")
if st.button("⚖️ Proyectar Motivación Jurídica"):
    s, p_base, grupos, p_add, t_final = calcular_tasa_reemplazo(ibl_input, smmlv_input, semanas_input)
    
    texto_resumen = f"Que al revisar la petición allegada, se advierte que el objeto central de la solicitud radica en: {st.session_state.resumen_peticion}\n\n" if st.session_state.resumen_peticion else ""
    
    motivacion = f"""CONSIDERANDO:

{texto_resumen}Que de conformidad con el artículo 33 de la Ley 100 de 1993, modificado por el artículo 9 de la Ley 797 de 2003, para tener derecho a la Pensión de Vejez es necesario acreditar las edades establecidas en la norma y un mínimo de 1.300 semanas de cotización.

Que el(la) afiliado(a) acredita un total de {semanas_input} semanas cotizadas al Sistema General de Pensiones, contabilizadas de acuerdo con el Parágrafo 2 del artículo 33 de la Ley 100 de 1993, aplicando la regla de conversión de 51,42 semanas por año, cumpliendo con el requisito de densidad exigido.

Que en cumplimiento del artículo 21 de la Ley 100 de 1993, se determinó el Ingreso Base de Liquidación (IBL) en la suma de ${ibl_input:,.2f} COP.

LIQUIDACIÓN DE LA TASA DE REEMPLAZO (Monto de la Pensión):
De conformidad con el artículo 34 de la Ley 100 de 1993, modificado por el artículo 10 de la Ley 797 de 2003, el monto mensual de la pensión se determina mediante una fórmula decreciente y la suma de puntos adicionales por semanas extra cotizadas. El cálculo se sustenta así:

1. Proporción del IBL respecto al salario mínimo (s):
Se divide el IBL (${ibl_input:,.2f}) entre el SMMLV del año respectivo (${smmlv_input:,.2f}), arrojando un factor (s) de {s:.2f} salarios mínimos.

2. Porcentaje Inicial:
Aplicando la fórmula legal r = 65.50 - 0.50s:
r = 65.50 - (0.50 * {s:.2f}) = {p_base:.2f}% (Se respeta el límite inferior legal del 55.5%).

3. Puntos adicionales por semanas excedentes:
El afiliado acreditó {semanas_input} semanas. Al restar las 1.300 semanas mínimas exigidas, se obtiene un excedente de {max(0, semanas_input - 1300)} semanas. 
La norma establece un incremento del 1.5% por cada 50 semanas adicionales (con un tope estricto de 15 puntos, equivalentes a 500 semanas). 
El afiliado cuenta con {grupos} bloque(s) completo(s) de 50 semanas.
Incremento adicional = {grupos} * 1.5% = {p_add:.2f}%.

4. Tasa de Reemplazo Definitiva:
Sumando el porcentaje inicial ({p_base:.2f}%) y los puntos adicionales ({p_add:.2f}%), se establece una tasa de reemplazo total del {t_final:.2f}% (Aplicando el tope máximo normativo del 80%).

DESCUENTOS DE LEY EN SALUD:
En consecuencia, el valor de la mesada pensional corresponderá al {t_final:.2f}% del IBL, quedando sujeta a los descuentos de Ley en materia de salud con cargo al pensionado (Art. 143 Ley 100 de 1993). Dado el monto liquidado, se aplicará el porcentaje de descuento conforme a lo regulado en el Artículo 1 de la Ley 2018 de 2020.
"""
    
    st.text_area("Vista previa del Acto Administrativo:", motivacion, height=500)
    
    word_file = generar_word(motivacion)
    st.download_button(
        label="📄 Descargar Motivación en Word (.docx)",
        data=word_file,
        file_name="Resolucion_Motivada_RPM.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
