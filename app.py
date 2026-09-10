import streamlit as st
from docx import Document
from io import BytesIO
import math
import PyPDF2
from datetime import datetime
import json

# LangChain Imports
from langchain_huggingface import HuggingFaceEndpoint
from langchain_core.prompts import PromptTemplate
from langchain_core.output_parsers import JsonOutputParser
from langchain_core.pydantic_v1 import BaseModel, Field

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Agente Sustanciador RPM - LangChain", layout="wide")

if 'ibl' not in st.session_state: st.session_state.ibl = 1500000.0
if 'smmlv' not in st.session_state: st.session_state.smmlv = 1300000.0
if 'semanas' not in st.session_state: st.session_state.semanas = 0
if 'edad' not in st.session_state: st.session_state.edad = 0
if 'texto_documento' not in st.session_state: st.session_state.texto_documento = ""
if 'datos_capturados' not in st.session_state: st.session_state.datos_capturados = False

st.title("⚖️ Agente Sustanciador - Motor LangChain")
st.markdown("Análisis semántico avanzado de Historias Laborales y Resoluciones mediante cadenas de LangChain y parseo estructurado.")

# --- LECTURA DE MÚLTIPLES ARCHIVOS ---
def leer_multiples_archivos(lista_archivos):
    texto_total = ""
    for archivo in lista_archivos:
        texto_total += f"\n\n--- INICIO DOCUMENTO: {archivo.name} ---\n\n"
        try:
            if archivo.name.endswith('.pdf'):
                lector = PyPDF2.PdfReader(archivo)
                for pagina in lector.pages:
                    if pagina.extract_text():
                        texto_total += pagina.extract_text() + "\n"
            elif archivo.name.endswith('.docx'):
                doc = Document(archivo)
                for parrafo in doc.paragraphs:
                    texto_total += parrafo.text + "\n"
            elif archivo.name.endswith('.txt'):
                texto_total += archivo.read().decode('utf-8') + "\n"
        except Exception as e:
            st.error(f"Error al leer el archivo {archivo.name}: {e}")
    return texto_total

# --- ESTRUCTURA DE DATOS ESPERADA (PYDANTIC) ---
class DatosPension(BaseModel):
    semanas: int = Field(description="Total de semanas cotizadas encontradas en el documento")
    ibl: float = Field(description="Ingreso Base de Liquidación (IBL) en formato numérico sin símbolos")
    edad: int = Field(description="Edad del peticionario. Si hay fecha de nacimiento, calcular edad al año actual")

# --- CADENA DE ANÁLISIS LANGCHAIN ---
def analizar_con_langchain(texto_documento, api_key):
    # Inicializar el modelo Open Source (Mixtral)
    llm = HuggingFaceEndpoint(
        repo_id="mistralai/Mixtral-8x7B-Instruct-v0.1",
        huggingfacehub_api_token=api_key,
        temperature=0.1,
        max_new_tokens=512
    )
    
    # Configurar el Parser para asegurar salida JSON
    parser = JsonOutputParser(pydantic_object=DatosPension)
    
    # Crear el Prompt Template con las instrucciones de formato inyectadas
    prompt = PromptTemplate(
        template="""Eres un experto jurídico analizando expedientes pensionales.
        Extrae la siguiente información del documento proporcionado.
        
        {format_instructions}
        
        DOCUMENTO:
        {texto}
        """,
        input_variables=["texto"],
        partial_variables={"format_instructions": parser.get_format_instructions()},
    )
    
    # Ensamblar la cadena (Chain)
    cadena = prompt | llm | parser
    
    # Ejecutar la cadena
    try:
        # Limitamos el texto para no exceder la ventana de contexto del modelo si es muy largo
        texto_truncado = texto_documento[:15000] 
        resultado = cadena.invoke({"texto": texto_truncado})
        return resultado
    except Exception as e:
        raise Exception(f"Fallo en la extracción de LangChain: {e}")

# --- FUNCIONES DE CÁLCULO Y WORD ---
def calcular_tasa_reemplazo(ibl, smmlv, semanas_totales):
    s = ibl / smmlv if smmlv > 0 else 0
    porcentaje_base = max(65.50 - (0.50 * s), 55.5)
    semanas_adicionales = max(0, semanas_totales - 1300)
    grupos_de_50 = math.floor(semanas_adicionales / 50)
    puntos_adicionales = min(grupos_de_50 * 1.5, 15.0)
    tasa_final = min(porcentaje_base + puntos_adicionales, 80.0)
    return s, porcentaje_base, grupos_de_50, puntos_adicionales, tasa_final

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
st.sidebar.header("⚙️ Motor LangChain")
api_key = st.sidebar.text_input("HuggingFace API Token:", type="password", help="Requerido para la inferencia semántica del LLM")

col_izq, col_der = st.columns([1, 1.2])

with col_izq:
    st.subheader("📄 1. Análisis de Expediente")
    
    archivos_cargados = st.file_uploader(
        "Adjuntar expedientes (PDF, Word, TXT)", 
        type=["pdf", "docx", "txt"], 
        accept_multiple_files=True
    )
    
    if archivos_cargados:
        if st.button("🧠 Ejecutar Cadena LangChain", use_container_width=True):
            if not api_key:
                st.warning("Ingrese su Token de HuggingFace en el panel lateral.")
            else:
                with st.spinner(f'Ejecuting LangChain Pipeline...'):
                    texto_consolidado = leer_multiples_archivos(archivos_cargados)
                    st.session_state.texto_documento = texto_consolidado
                    
                    try:
                        datos = analizar_con_langchain(texto_consolidado, api_key)
                        
                        st.session_state.semanas = datos.get("semanas", 0)
                        st.session_state.ibl = datos.get("ibl", 1500000.0)
                        st.session_state.edad = datos.get("edad", 60)
                        st.session_state.datos_capturados = True
                        st.success("✅ Extracción Semántica Completada.")
                    except Exception as e:
                        st.error(str(e))
                
    if st.session_state.datos_capturados:
        c1, c2, c3 = st.columns(3)
        c1.metric("Semanas", f"{st.session_state.semanas}")
        c2.metric("IBL", f"${st.session_state.ibl:,.0f}")
        c3.metric("Edad", f"{st.session_state.edad} años")
                
    if st.session_state.texto_documento:
        with st.expander("Ver texto consolidado de los documentos", expanded=False):
            st.text_area("Texto extraído:", st.session_state.texto_documento, height=250)

with col_der:
    st.subheader("⚙️ 2. Variables y Decisión")
    
    col_a, col_b = st.columns(2)
    with col_a:
        genero = st.selectbox("Género del Peticionario:", ["Femenino", "Masculino"])
        edad_input = st.number_input("Edad verificada:", min_value=0, value=st.session_state.edad, step=1)
    with col_b:
        semanas_input = st.number_input("Semanas a reconocer:", min_value=0, value=st.session_state.semanas, step=1)
        
    st.divider()
    ibl_input = st.number_input("Ingreso Base de Liquidación (IBL):", min_value=0.0, value=float(st.session_state.ibl), step=10000.0)
    smmlv_input = st.number_input("SMMLV aplicable:", min_value=0.0, value=st.session_state.smmlv, step=10000.0)
    
    st.divider()
    
    if st.button("⚖️ Evaluar y Proyectar Acto Administrativo", use_container_width=True):
        
        edad_requerida = 57 if genero == "Femenino" else 62
        cumple_edad = edad_input >= edad_requerida
        cumple_semanas = semanas_input >= 1300
        
        if cumple_edad and cumple_semanas:
            st.success("🟢 ESTADO: CUMPLE REQUISITOS. Generando reconocimiento.")
            s, p_base, grupos, p_add, t_final = calcular_tasa_reemplazo(ibl_input, smmlv_input, semanas_input)
            
            motivacion = f"""CONSIDERANDO:

Que de conformidad con el artículo 33 de la Ley 100 de 1993, modificado por el artículo 9 de la Ley 797 de 2003, para tener derecho a la Pensión de Vejez es necesario acreditar las edades establecidas en la norma y un mínimo de 1.300 semanas de cotización.

Que el(la) afiliado(a) cuenta con {edad_input} años de edad, cumpliendo con la edad exigida por la normatividad vigente para el género {genero.lower()} ({edad_requerida} años).

Que revisada la historia laboral y demás documentos allegados, el(la) afiliado(a) acredita un total de {semanas_input} semanas cotizadas al Sistema General de Pensiones, cumpliendo con el requisito de densidad exigido.

Que en cumplimiento del artículo 21 de la Ley 100 de 1993, se determinó el Ingreso Base de Liquidación (IBL) en la suma de ${ibl_input:,.2f} COP.

LIQUIDACIÓN DE LA TASA DE REEMPLAZO:
1. Proporción del IBL respecto al salario mínimo: {s:.2f} SMMLV.
2. Porcentaje Inicial: {p_base:.2f}%.
3. Puntos adicionales ({grupos} grupos de 50 semanas extra): {p_add:.2f}%.
4. Tasa de Reemplazo Definitiva: {t_final:.2f}%.

En consecuencia, el valor de la mesada pensional corresponderá al {t_final:.2f}% del IBL, quedando sujeta a los descuentos de Ley en materia de salud con cargo al pensionado (Art. 143 Ley 100 de 1993 y Art. 1 Ley 2018 de 2020).
"""
        else:
            st.error("🔴 ESTADO: NO CUMPLE REQUISITOS. Generando acto denegatorio en lenguaje claro.")
            faltante_semanas = 1300 - semanas_input if not cumple_semanas else 0
            
            texto_edad = f"Actualmente, usted tiene {edad_input} años. Como la ley exige {edad_requerida} años para su género, usted **{'sí' if cumple_edad else 'aún no'}** cumple este requisito."
            texto_semanas = f"Usted cuenta con {semanas_input} semanas cotizadas. Le hacen falta {faltante_semanas} semanas para cumplir la meta de 1.300." if not cumple_semanas else f"Usted cuenta con {semanas_input} semanas cotizadas, cumpliendo este requisito."
            
            motivacion = f"""CONSIDERANDO:

Para nuestra entidad es fundamental brindarle total claridad sobre su situación pensional. Hemos revisado detalladamente su historia laboral para determinar si en este momento es posible reconocer su pensión de vejez.

De acuerdo con la Ley 797 de 2003, para acceder a la pensión de vejez se debe cumplir con dos condiciones al mismo tiempo:
1. Tener la edad requerida ({edad_requerida} años).
2. Haber cotizado un mínimo de 1.300 semanas.

Su situación actual tras la auditoría de sus documentos es la siguiente:
- Requisito de Edad: {texto_edad}
- Requisito de Semanas: {texto_semanas}

Al no reunirse la totalidad de los requisitos de forma simultánea, nos resulta jurídicamente imposible aprobar su pensión de vejez en este instante. 

Alternativas:
{'1. Puede continuar cotizando al sistema hasta completar las semanas faltantes.' if not cumple_semanas else ''}
2. Solicitar la Indemnización Sustitutiva de Vejez: Si declara su imposibilidad de seguir cotizando, y teniendo en cuenta su edad, tiene derecho a solicitar la devolución de sus aportes (Art. 37 de la Ley 100 de 1993).

Por lo expuesto, se hace necesario negar el reconocimiento de la pensión de vejez en esta oportunidad.
"""
        
        st.text_area("Vista previa del Acto Administrativo:", motivacion, height=350)
        word_file = generar_word(motivacion)
        st.download_button(
            label="📄 Descargar Motivación en Word (.docx)",
            data=word_file,
            file_name="Proyecto_Resolucion.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
