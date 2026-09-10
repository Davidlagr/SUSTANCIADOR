import streamlit as st
from docx import Document
from io import BytesIO
import math
import PyPDF2
from datetime import datetime
import re

# LangChain Imports
from langchain_huggingface import HuggingFaceEndpoint
from langchain_core.prompts import PromptTemplate

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Agente Sustanciador Híbrido", layout="wide")

# Inicialización de estado de variables extraídas (Token-Free)
vars_keys = ['nombre', 'cedula', 'ibl', 'smmlv', 'semanas', 'edad', 'genero', 'motivacion_generada']
for key in vars_keys:
    if key not in st.session_state:
        st.session_state[key] = "" if key in ['nombre', 'cedula', 'motivacion_generada'] else 0

if 'ibl' not in st.session_state or st.session_state.ibl == 0: st.session_state.ibl = 1500000.0
if 'smmlv' not in st.session_state or st.session_state.smmlv == 0: st.session_state.smmlv = 1300000.0
if 'genero' not in st.session_state or st.session_state.genero == 0: st.session_state.genero = "Femenino"

st.title("⚖️ Agente Sustanciador - Arquitectura Híbrida")
st.markdown("1. Extracción gratuita mediante motor local. | 2. Cálculos matemáticos de precisión. | 3. Redacción jurídica profunda vía IA (LangChain).")

# --- 1. MÓDULO DE LECTURA Y EXTRACCIÓN (TOKEN-FREE) ---
def extraer_texto(archivo):
    texto = ""
    try:
        if archivo.name.endswith('.pdf'):
            lector = PyPDF2.PdfReader(archivo)
            for pagina in lector.pages:
                if pagina.extract_text(): texto += pagina.extract_text() + "\n"
        elif archivo.name.endswith('.docx'):
            doc = Document(archivo)
            for parrafo in doc.paragraphs: texto += parrafo.text + "\n"
        elif archivo.name.endswith('.txt'):
            texto = archivo.read().decode('utf-8')
    except Exception as e:
        st.error(f"Error al leer {archivo.name}: {e}")
    return texto

def procesar_historia_laboral(texto):
    # Buscar semanas totales
    match_hl = re.search(r'TOTAL SEMANAS COTIZADAS[\s:]*([\d.,]+)', texto, re.IGNORECASE)
    if match_hl:
        num = match_hl.group(1).replace('.', '').replace(',', '.')
        st.session_state.semanas = int(float(num))
    
    # Buscar fecha nacimiento para calcular edad
    match_nac = re.search(r'Nacimiento[\s:]*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    if match_nac:
        try:
            fnac = datetime.strptime(match_nac.group(1), '%d/%m/%Y')
            st.session_state.edad = datetime.now().year - fnac.year
        except: pass
        
    # Buscar Nombre y Cédula
    match_ced = re.search(r'(?:Número de Documento|C\.?C\.?)[\s:]*([\d.,]+)', texto, re.IGNORECASE)
    if match_ced: st.session_state.cedula = match_ced.group(1).replace('.', '').strip()
    
    match_nom = re.search(r'Nombre[\s:]*([A-Za-zÑñÁÉÍÓÚáéíóú\s]+)', texto, re.IGNORECASE)
    if match_nom:
        n = match_nom.group(1).strip()
        if len(n) > 3 and "Dirección" not in n: st.session_state.nombre = n

def procesar_resolucion(texto):
    match_ibl = re.search(r'(?:IBL|Liquidación[\s\w]*de)[\s:\$]*([\d.,]{6,})', texto, re.IGNORECASE)
    if match_ibl:
        num = match_ibl.group(1).replace(',', '').replace('.', '')
        if num.isdigit(): st.session_state.ibl = float(num)
        
    match_res = re.search(r'([\d.,]+)\s*semanas', texto, re.IGNORECASE)
    if match_res and st.session_state.semanas == 0:
        num = match_res.group(1).replace(',', '').replace('.', '')
        if num.isdigit(): st.session_state.semanas = int(num)

def procesar_peticion(texto):
    match_edad = re.search(r'(?:edad\s*de\s*)?(\d{2})\s*años', texto, re.IGNORECASE)
    if match_edad and st.session_state.edad == 0:
        st.session_state.edad = int(match_edad.group(1))

# --- 2. CÁLCULO MATEMÁTICO DE REGLAS (TOKEN-FREE) ---
def calcular_derecho_pensional(edad, semanas, genero, ibl, smmlv):
    # Reglas extraídas del documento legal Colpensiones
    edad_req = 57 if genero == "Femenino" else 62
    cumple_edad = edad >= edad_req
    cumple_semanas = semanas >= 1300
    
    s = ibl / smmlv if smmlv > 0 else 0
    porcentaje_base = max(65.50 - (0.50 * s), 55.5)
    semanas_adicionales = max(0, semanas - 1300)
    grupos_de_50 = math.floor(semanas_adicionales / 50)
    puntos_adicionales = min(grupos_de_50 * 1.5, 15.0) # Tope 15 puntos (Art 34 Ley 100/ Ley 797)
    tasa_final = min(porcentaje_base + puntos_adicionales, 80.0) # Tope 80%
    
    # Descuento de Salud
    mesada_calculada = ibl * (tasa_final / 100)
    if mesada_calculada <= smmlv: desc_salud = "4%"
    elif mesada_calculada <= (smmlv * 2): desc_salud = "10%"
    else: desc_salud = "12%"
    
    return {
        "cumple_derecho": cumple_edad and cumple_semanas,
        "faltante_semanas": 1300 - semanas if not cumple_semanas else 0,
        "tasa_final": tasa_final,
        "desc_salud": desc_salud,
        "mesada": mesada_calculada
    }

# --- 3. ANÁLISIS A FONDO Y REDACCIÓN VÍA LANGCHAIN (CONSUME TOKENS) ---
def redactar_acto_langchain(datos_solicitante, calculos, api_key):
    llm = HuggingFaceEndpoint(
        repo_id="mistralai/Mixtral-8x7B-Instruct-v0.1",
        huggingfacehub_api_token=api_key,
        temperature=0.2,
        max_new_tokens=1024
    )
    
    plantilla = """[INST] Eres un sustanciador experto de Colpensiones. Redacta la sección "CONSIDERANDO" de una resolución administrativa en Colombia (Ley 100 de 1993 y Ley 797 de 2003). 
    
    REGLAS ESTRICTAS DE REDACCIÓN:
    - Inicia directamente con el texto legal. No saludes ni hagas introducciones.
    - Si cumple el derecho, fundamenta el reconocimiento, detalla el cálculo de la tasa de reemplazo y menciona el descuento de salud aplicable.
    - Si NO cumple el derecho, redacta una negativa empática, explicando claramente cuántas semanas le faltan y mencionando la alternativa de la Indemnización Sustitutiva de Vejez.
    
    DATOS DEL PETICIONARIO A INCLUIR:
    Nombre: {nombre}
    Cédula: {cedula}
    Género: {genero}
    Edad Actual: {edad} años
    Semanas Cotizadas Validadas: {semanas}
    IBL Calculado: ${ibl}
    
    RESULTADO DEL ANÁLISIS TÉCNICO (OBLIGATORIO APLICAR):
    ¿Cumple el derecho?: {cumple_derecho}
    Semanas faltantes (si aplica): {faltante_semanas}
    Tasa de Reemplazo Final: {tasa_final}%
    Descuento de Salud aplicable (Ley 2018 de 2020): {desc_salud}
    [/INST]"""
    
    prompt = PromptTemplate(
        template=plantilla,
        input_variables=["nombre", "cedula", "genero", "edad", "semanas", "ibl", "cumple_derecho", "faltante_semanas", "tasa_final", "desc_salud"]
    )
    
    cadena = prompt | llm
    
    parametros = {
        "nombre": datos_solicitante['nombre'],
        "cedula": datos_solicitante['cedula'],
        "genero": datos_solicitante['genero'],
        "edad": str(datos_solicitante['edad']),
        "semanas": str(datos_solicitante['semanas']),
        "ibl": f"{datos_solicitante['ibl']:,.0f}",
        "cumple_derecho": "SÍ" if calculos['cumple_derecho'] else "NO",
        "faltante_semanas": str(calculos['faltante_semanas']),
        "tasa_final": f"{calculos['tasa_final']:.2f}",
        "desc_salud": calculos['desc_salud']
    }
    
    return cadena.invoke(parametros)

def generar_word(texto_motivacion):
    doc = Document()
    doc.add_heading('MOTIVACIÓN DEL ACTO ADMINISTRATIVO', 0)
    for parrafo in texto_motivacion.split('\n'):
        if parrafo.strip(): doc.add_paragraph(parrafo.strip())
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ DE USUARIO ---
st.sidebar.header("⚙️ Análisis LangChain")
api_key = st.sidebar.text_input("HuggingFace API Token:", type="password", help="Solo se utilizará para la redacción final, ahorrando tokens.")

col1, col2 = st.columns([1, 1.2])

with col1:
    st.subheader("📂 1. Módulos de Carga de Documentos")
    st.info("La extracción local mediante patrones es gratuita e inmediata.")
    
    arc_hl = st.file_uploader("A. Historia Laboral (PDF)", type=["pdf", "txt"], key="hl")
    arc_res = st.file_uploader("B. Resoluciones Previas (PDF, Word)", type=["pdf", "docx", "txt"], key="res")
    arc_pet = st.file_uploader("C. Petición del Ciudadano (PDF, Word)", type=["pdf", "docx", "txt"], key="pet")
    
    if st.button("🔍 Extraer Datos del Solicitante (Local / Sin Tokens)", use_container_width=True):
        with st.spinner("Leyendo documentos..."):
            if arc_hl: procesar_historia_laboral(extraer_texto(arc_hl))
            if arc_res: procesar_resolucion(extraer_texto(arc_res))
            if arc_pet: procesar_peticion(extraer_texto(arc_pet))
            st.success("Extracción completada. Revise los datos consolidados a la derecha.")

with col2:
    st.subheader("👤 2. Perfil del Solicitante y Variables")
    
    c_nom, c_ced = st.columns(2)
    with c_nom:
        st.session_state.nombre = st.text_input("Nombre Completo:", value=st.session_state.nombre)
    with c_ced:
        st.session_state.cedula = st.text_input("Cédula:", value=st.session_state.cedula)
        
    c_gen, c_edad, c_sem = st.columns(3)
    with c_gen:
        st.session_state.genero = st.selectbox("Género:", ["Femenino", "Masculino"], index=0 if st.session_state.genero == "Femenino" else 1)
    with c_edad:
        st.session_state.edad = st.number_input("Edad:", min_value=0, value=st.session_state.edad)
    with c_sem:
        st.session_state.semanas = st.number_input("Semanas Validadas:", min_value=0, value=st.session_state.semanas)
        
    c_ibl, c_smmlv = st.columns(2)
    with c_ibl:
        st.session_state.ibl = st.number_input("IBL Calculado ($):", min_value=0.0, value=float(st.session_state.ibl), step=10000.0)
    with c_smmlv:
        st.session_state.smmlv = st.number_input("SMMLV aplicable ($):", min_value=0.0, value=float(st.session_state.smmlv), step=10000.0)
        
    st.divider()
    
    if st.button("⚖️ Generar Análisis y Motivación Jurídica (LangChain)", type="primary", use_container_width=True):
        if not api_key:
            st.warning("⚠️ Requiere el Token de HuggingFace en la barra lateral para redactar el documento.")
        else:
            with st.spinner("La IA está redactando la motivación basándose en las Reglas de Colpensiones..."):
                # Ejecutar cálculo de reglas duro (Python)
                calculos = calcular_derecho_pensional(
                    st.session_state.edad, st.session_state.semanas, 
                    st.session_state.genero, st.session_state.ibl, st.session_state.smmlv
                )
                
                # Ejecutar redacción (LLM)
                datos_sol = {
                    "nombre": st.session_state.nombre,
                    "cedula": st.session_state.cedula,
                    "genero": st.session_state.genero,
                    "edad": st.session_state.edad,
                    "semanas": st.session_state.semanas,
                    "ibl": st.session_state.ibl
                }
                
                try:
                    resultado_ia = redactar_acto_langchain(datos_sol, calculos, api_key)
                    st.session_state.motivacion_generada = resultado_ia
                except Exception as e:
                    st.error(f"Error en la IA: {e}")

    if st.session_state.motivacion_generada:
        st.text_area("Vista previa de la Resolución:", st.session_state.motivacion_generada, height=350)
        
        doc_word = generar_word(st.session_state.motivacion_generada)
        st.download_button(
            label="📄 Descargar Motivación en Word (.docx)",
            data=doc_word,
            file_name=f"Resolucion_{st.session_state.cedula}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
