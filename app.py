import streamlit as st
from docx import Document
from io import BytesIO
import math
import PyPDF2
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Agente Sustanciador RPM (Autónomo)", layout="wide")

# Inicializar estado de variables
if 'ibl' not in st.session_state: st.session_state.ibl = 1500000.0
if 'smmlv' not in st.session_state: st.session_state.smmlv = 1300000.0
if 'semanas' not in st.session_state: st.session_state.semanas = 1150
if 'texto_documento' not in st.session_state: st.session_state.texto_documento = ""

st.title("⚖️ Agente Sustanciador - RPM (Versión Lenguaje Claro)")
st.markdown("Genera motivaciones aprobatorias rigurosas o resoluciones denegatorias empáticas y pedagógicas, orientadas al ciudadano.")

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
    datos = {"semanas": 1150, "ibl": 1500000.0}
    match_semanas = re.search(r'((?:\d{1,3}[.,]?\d{3})|\d{3,4})\s*semanas', texto, re.IGNORECASE)
    if match_semanas:
        num_limpio = re.sub(r'[.,]', '', match_semanas.group(1))
        if num_limpio.isdigit(): datos["semanas"] = int(num_limpio)
            
    match_ibl = re.search(r'(?:IBL|ingreso base|promedio).*?\$?\s*((?:\d{1,3}[.,]?)+(?:\d{3}))', texto, re.IGNORECASE)
    if match_ibl:
        num_limpio = re.sub(r'[.,]', '', match_ibl.group(1))
        if num_limpio.isdigit(): datos["ibl"] = float(num_limpio)
            
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
                datos = extraer_datos_locales(texto_peticion)
                st.session_state.semanas = datos["semanas"]
                st.session_state.ibl = datos["ibl"]
                st.success("Extracción completada. Revisa los datos en el panel derecho.")
                
    if st.session_state.texto_documento:
        with st.expander("Ver texto extraído del documento", expanded=False):
            st.text_area("Texto sin formato:", st.session_state.texto_documento, height=200)

with col_der:
    st.subheader("⚙️ 2. Estructuración de la Decisión")
    
    sentido_decision = st.radio("Sentido del Acto Administrativo:", 
                                ["Reconocer Pensión (Cumple Requisitos)", "Negar Pensión (Lenguaje Claro y Empático)"],
                                horizontal=False)
    st.divider()
    
    # Variables comunes
    semanas_input = st.number_input("Total Semanas Cotizadas:", min_value=0, value=st.session_state.semanas, step=1)
    
    if "Reconocer" in sentido_decision:
        ibl_input = st.number_input("Ingreso Base de Liquidación (IBL):", min_value=0.0, value=st.session_state.ibl, step=100000.0)
        smmlv_input = st.number_input("SMMLV del año de causación:", min_value=0.0, value=st.session_state.smmlv, step=10000.0)
    
    if "Negar" in sentido_decision:
        genero = st.selectbox("Género del Peticionario:", ["Femenino", "Masculino"])
        edad_input = st.number_input("Edad actual del peticionario:", min_value=0, value=60, step=1)
        edad_requerida = 57 if genero == "Femenino" else 62
        
    st.divider()
    
    if st.button("⚖️ Proyectar Motivación", use_container_width=True):
        
        if "Reconocer" in sentido_decision:
            s, p_base, grupos, p_add, t_final = calcular_tasa_reemplazo(ibl_input, smmlv_input, semanas_input)
            
            motivacion = f"""CONSIDERANDO:

Que de conformidad con el artículo 33 de la Ley 100 de 1993, modificado por el artículo 9 de la Ley 797 de 2003, para tener derecho a la Pensión de Vejez es necesario acreditar las edades establecidas en la norma y un mínimo de 1.300 semanas de cotización.

Que el(la) afiliado(a) acredita un total de {semanas_input} semanas cotizadas al Sistema General de Pensiones, contabilizadas de acuerdo con el Parágrafo 2 del artículo 33 de la Ley 100 de 1993, aplicando la regla de conversión de 51,42 semanas por año, cumpliendo con el requisito de densidad exigido.

Que en cumplimiento del artículo 21 de la Ley 100 de 1993, se determinó el Ingreso Base de Liquidación (IBL) en la suma de ${ibl_input:,.2f} COP.

LIQUIDACIÓN DE LA TASA DE REEMPLAZO (Monto de la Pensión):
1. Proporción del IBL respecto al salario mínimo (s): Se divide el IBL (${ibl_input:,.2f}) entre el SMMLV respectivo (${smmlv_input:,.2f}), arrojando un factor de {s:.2f} salarios mínimos.
2. Porcentaje Inicial: r = 65.50 - (0.50 * {s:.2f}) = {p_base:.2f}%.
3. Puntos adicionales: El afiliado cuenta con {grupos} bloque(s) completo(s) de 50 semanas adicionales a las 1.300. Incremento = {p_add:.2f}%.
4. Tasa de Reemplazo Definitiva: {t_final:.2f}% (Aplicando los topes normativos).

En consecuencia, el valor de la mesada pensional corresponderá al {t_final:.2f}% del IBL, quedando sujeta a los descuentos de Ley en materia de salud con cargo al pensionado (Art. 143 Ley 100 de 1993 y Art. 1 Ley 2018 de 2020).
"""
        else:
            # Lógica para redactar la negativa en lenguaje claro
            faltante_semanas = 1300 - semanas_input
            cumple_edad = edad_input >= edad_requerida
            
            texto_edad = f"Actualmente, usted tiene {edad_input} años. Como la ley exige {edad_requerida} años para el género {genero.lower()}, usted **{'sí' if cumple_edad else 'aún no'}** cumple con el requisito de edad."
            
            motivacion = f"""CONSIDERANDO:

Para nuestra entidad es fundamental brindarle total claridad sobre su situación pensional y darle respuesta a su solicitud de manera transparente y comprensible.

Entendemos el esfuerzo y la dedicación que representa cada semana cotizada a lo largo de su vida laboral. Por ello, hemos revisado detalladamente su historia laboral para determinar si en este momento es posible reconocer su pensión de vejez.

¿Cuáles son los requisitos que exige la Ley?
De acuerdo con la Ley 797 de 2003, para que cualquier ciudadano en Colombia acceda a la pensión de vejez en el Régimen de Prima Media, debe cumplir obligatoriamente con dos condiciones al mismo tiempo:
1. Tener la edad requerida ({edad_requerida} años para el caso del género {genero.lower()}).
2. Haber cotizado un mínimo de 1.300 semanas.

Su situación actual:
Al realizar el conteo matemático de sus aportes, validamos lo siguiente:
- Requisito de Edad: {texto_edad}
- Requisito de Semanas: Usted cuenta con {semanas_input} semanas cotizadas válidas en el sistema.

¿Por qué no es posible acceder a la pensión en este momento?
Dado que la ley nos exige un mínimo de 1.300 semanas para otorgar el derecho, y usted cuenta con {semanas_input} semanas, le hacen falta {faltante_semanas} semanas para cumplir la meta legal. Como entidad pública, debemos aplicar la norma de manera estricta, lo que nos imposibilita jurídicamente aprobar su pensión de vejez en este instante.

¿Qué alternativas tiene a su disposición?
Queremos acompañarlo(a) en este proceso. Al no cumplir con las semanas, usted tiene las siguientes opciones:

1. Continuar cotizando: Si está dentro de sus posibilidades laborales o económicas (como trabajador dependiente o independiente), puede seguir realizando aportes al sistema hasta completar las {faltante_semanas} semanas que le faltan para consolidar su derecho a una pensión vitalicia.

2. Solicitar la Indemnización Sustitutiva de Vejez: Si usted declara su imposibilidad absoluta de seguir cotizando, y teniendo en cuenta que {'ya cumplió' if cumple_edad else 'una vez cumpla'} la edad de {edad_requerida} años, tiene derecho a solicitar la devolución de sus aportes. Esta figura legal (Art. 37 de la Ley 100 de 1993) le permite recibir en un único pago el valor ajustado de los saldos cotizados durante su vida.

Por lo anterior expuesto, y con el propósito de garantizar el debido proceso y la legalidad, se hace necesario negar el reconocimiento de la pensión de vejez, dejando a salvo sus derechos para que opte por las alternativas mencionadas.
"""
        
        st.text_area("Vista previa del Acto Administrativo:", motivacion, height=450)
        
        word_file = generar_word(motivacion)
        st.download_button(
            label="📄 Descargar Motivación en Word (.docx)",
            data=word_file,
            file_name="Resolucion_RPM.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
