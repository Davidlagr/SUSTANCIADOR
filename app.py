import streamlit as st
from docx import Document
from io import BytesIO
import math
import PyPDF2
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Agente Sustanciador RPM (Autónomo)", layout="wide")

if 'ibl' not in st.session_state: st.session_state.ibl = 1500000.0
if 'smmlv' not in st.session_state: st.session_state.smmlv = 1300000.0
if 'semanas' not in st.session_state: st.session_state.semanas = 1300
if 'texto_documento' not in st.session_state: st.session_state.texto_documento = ""

st.title("⚖️ Agente Sustanciador - RPM (Versión Autónoma)")
st.markdown("Herramienta gratuita y privada. Carga el documento, el sistema extraerá el texto, buscará los datos clave mediante patrones y proyectará la motivación jurídica sin usar APIs externas.")

# --- FUNCIONES DE LECTURA DE ARCHIVOS ---
def leer_archivo(archivo):
    texto = ""
    try:
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
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
    return texto

# --- FUNCIÓN DE EXTRACCIÓN MEDIANTE PATRONES (REGEX) ---
def extraer_datos_locales(texto):
    datos = {"semanas": 1300, "ibl": 1500000.0}
    
    # Buscar patrones de semanas (ej. "1350 semanas", "1.420 semanas", "1,200 semanas")
    match_semanas = re.search(r'((?:\d{1,3}[.,]?\d{3})|\d{3,4})\s*semanas', texto, re.IGNORECASE)
    if match_semanas:
        # Limpiar puntos y comas del número encontrado
        num_limpio = re.sub(r'[.,]', '', match_semanas.group(1))
        if num_limpio.isdigit():
            datos["semanas"] = int(num_limpio)
            
    # Buscar patrones que parezcan un IBL (ej. "$ 1.500.000", "IBL de 2.300.000")
    match_ibl = re.search(r'(?:IBL|ingreso base|promedio).*?\$?\s*((?:\d{1,3}[.,]?)+(?:\d{3}))', texto, re.IGNORECASE)
    if match_ibl:
        num_limpio = re.sub(r'[.,]', '', match_ibl.group(1))
        if num_limpio.isdigit():
            datos["ibl"] = float(num_limpio)
            
    return datos

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
col_izq, col_der = st.columns([1, 1])

with col_izq:
    st.subheader("📄 1. Análisis de Expediente")
    archivo_cargado = st.file_uploader("Adjuntar petición (PDF, Word, TXT)", type=["pdf", "docx", "txt"])
    
    if archivo_cargado is not None:
        if st.button("🔍 Extraer Datos Localmente"):
            with st.spinner('Extrayendo texto y buscando variables...'):
                texto_peticion = leer_archivo(archivo_cargado)
                st.session_state.texto_documento = texto_peticion
                
                # Ejecutar motor de reglas
                datos = extraer_datos_locales(texto_peticion)
                st.session_state.semanas = datos["semanas"]
                st.session_state.ibl = datos["ibl"]
                st.success("Extracción completada. Revisa los datos en el panel derecho.")
                
    if st.session_state.texto_documento:
        with st.expander("Ver texto extraído del documento", expanded=True):
            st.text_area("Texto sin formato:", st.session_state.texto_documento, height=300)

with col_der:
    st.subheader("⚙️ 2. Liquidación y Formulación")
    st.info("Verifica y ajusta las variables extraídas antes de generar el acto.")
    
    ibl_input = st.number_input("Ingreso Base de Liquidación (IBL):", min_value=0.0, value=st.session_state.ibl, step=100000.0)
    smmlv_input = st.number_input("SMMLV del año de causación:", min_value=0.0, value=st.session_state.smmlv, step=10000.0)
    semanas_input = st.number_input("Total Semanas Cotizadas:", min_value=0, value=st.session_state.semanas, step=1)
    
    st.divider()
    
    if st.button("⚖️ Proyectar Motivación Jurídica", use_container_width=True):
        s, p_base, grupos, p_add, t_final = calcular_tasa_reemplazo(ibl_input, smmlv_input, semanas_input)
        
        motivacion = f"""CONSIDERANDO:

Que de conformidad con el artículo 33 de la Ley 100 de 1993, modificado por el artículo 9 de la Ley 797 de 2003, para tener derecho a la Pensión de Vejez es necesario acreditar las edades establecidas en la norma y un mínimo de 1.300 semanas de cotización.

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
En consecuencia, el valor de la mesada pensional corresponderá al {t_final:.2f}% del IBL, quedando sujeta a los descuentos de Ley en materia de salud con cargo al pensionado (Art. 143 Ley 100 de 1993 y Art. 1 Ley 2018 de 2020).
"""
        
        st.text_area("Vista previa del Acto Administrativo:", motivacion, height=400)
        
        word_file = generar_word(motivacion)
        st.download_button(
            label="📄 Descargar Motivación en Word (.docx)",
            data=word_file,
            file_name="Resolucion_Motivada_RPM.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
