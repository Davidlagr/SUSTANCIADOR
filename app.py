import streamlit as st
from docx import Document
from io import BytesIO
import math
import PyPDF2
import re
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Agente Sustanciador RPM", layout="wide")

# Inicializar estado de variables
if 'ibl' not in st.session_state: st.session_state.ibl = 1500000.0
if 'smmlv' not in st.session_state: st.session_state.smmlv = 1300000.0
if 'semanas' not in st.session_state: st.session_state.semanas = 0
if 'edad' not in st.session_state: st.session_state.edad = 0
if 'texto_documento' not in st.session_state: st.session_state.texto_documento = ""
if 'datos_capturados' not in st.session_state: st.session_state.datos_capturados = False

st.title("⚖️ Agente Sustanciador - Análisis Multinorma (Resoluciones e Historia Laboral)")
st.markdown("Extrae datos automáticamente de Historias Laborales y Resoluciones, mostrándolos en un panel de verificación antes de proyectar el acto administrativo.")

# --- FUNCIONES DE LECTURA DE ARCHIVOS ---
def leer_archivo(archivo):
    texto = ""
    try:
        if archivo.name.endswith('.pdf'):
            lector = PyPDF2.PdfReader(archivo)
            for pagina in lector.pages:
                if pagina.extract_text():
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

# --- FUNCIÓN DE EXTRACCIÓN AVANZADA (REGEX) ---
def extraer_datos_locales(texto):
    datos = {"semanas": 0, "ibl": 0.0, "edad": 0}
    
    # 1. Extraer Semanas (Prioriza Historia Laboral, luego Resoluciones)
    # Busca "TOTAL SEMANAS COTIZADAS: 1.385,14" 
    match_hl = re.search(r'TOTAL SEMANAS COTIZADAS[\s:]*([\d.,]+)', texto, re.IGNORECASE)
    if match_hl:
        num_limpio = match_hl.group(1).replace('.', '').replace(',', '.')
        datos["semanas"] = int(float(num_limpio))
    else:
        # Busca en resoluciones: "1,424 semanas"
        match_res = re.search(r'([\d.,]+)\s*semanas', texto, re.IGNORECASE)
        if match_res:
            num_limpio = match_res.group(1).replace(',', '').replace('.', '')
            if num_limpio.isdigit(): datos["semanas"] = int(num_limpio)
            
    # 2. Extraer IBL (Busca patrones como "IBL:8,133,378" o "Liquidación de $8,133,378")
    match_ibl = re.search(r'(?:IBL|Liquidación[\s\w]*de)[\s:\$]*([\d.,]{6,})', texto, re.IGNORECASE)
    if match_ibl:
        num_limpio = match_ibl.group(1).replace(',', '').replace('.', '')
        if num_limpio.isdigit(): datos["ibl"] = float(num_limpio)
            
    # 3. Extraer Edad (Busca "67 años de edad" o calcula desde "Fecha de Nacimiento: 22/01/1974")
    match_nacimiento = re.search(r'Nacimiento[\s:]*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    if match_nacimiento:
        try:
            fecha_nac = datetime.strptime(match_nacimiento.group(1), '%d/%m/%Y')
            edad_calculada = datetime.now().year - fecha_nac.year
            datos["edad"] = edad_calculada
        except:
            pass
    else:
        match_edad = re.search(r'(?:edad\s*de\s*)?(\d{2})\s*años', texto, re.IGNORECASE)
        if match_edad:
            datos["edad"] = int(match_edad.group(1))
            
    return datos

# --- FUNCIONES DE CÁLCULO ---
def calcular_tasa_reemplazo(ibl, smmlv, semanas_totales):
    s = ibl / smmlv if smmlv > 0 else 0
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
col_izq, col_der = st.columns([1, 1.2])

with col_izq:
    st.subheader("📄 1. Análisis de Expediente")
    archivo_cargado = st.file_uploader("Adjuntar Historia Laboral o Resolución (PDF, Word, TXT)", type=["pdf", "docx", "txt"])
    
    if archivo_cargado is not None:
        if st.button("🔍 Extraer y Mostrar Datos", use_container_width=True):
            with st.spinner('Procesando documento y aplicando reglas de extracción...'):
                texto_peticion = leer_archivo(archivo_cargado)
                st.session_state.texto_documento = texto_peticion
                datos = extraer_datos_locales(texto_peticion)
                
                # Guardar en sesión
                st.session_state.semanas = datos["semanas"]
                st.session_state.ibl = datos["ibl"] if datos["ibl"] > 0 else 1500000.0
                st.session_state.edad = datos["edad"] if datos["edad"] > 0 else 60
                st.session_state.datos_capturados = True
                
    if st.session_state.datos_capturados:
        st.success("✅ Lectura Finalizada. Verifique los datos encontrados:")
        # Panel visual de datos capturados
        c1, c2, c3 = st.columns(3)
        c1.metric("Semanas Encontradas", f"{st.session_state.semanas}")
        c2.metric("IBL Encontrado", f"${st.session_state.ibl:,.0f}")
        c3.metric("Edad Calculada", f"{st.session_state.edad} años")
                
    if st.session_state.texto_documento:
        with st.expander("Ver texto plano del documento", expanded=False):
            st.text_area("Texto extraído:", st.session_state.texto_documento, height=250)

with col_der:
    st.subheader("⚙️ 2. Variables y Decisión")
    st.info("Puede ajustar manualmente los datos extraídos si el documento original presentaba errores de formato.")
    
    col_a, col_b = st.columns(2)
    with col_a:
        genero = st.selectbox("Género del Peticionario:", ["Femenino", "Masculino"])
        edad_input = st.number_input("Edad verificada:", min_value=0, value=st.session_state.edad, step=1)
    with col_b:
        semanas_input = st.number_input("Semanas a reconocer:", min_value=0, value=st.session_state.semanas, step=1)
        
    st.divider()
    ibl_input = st.number_input("Ingreso Base de Liquidación (IBL):", min_value=0.0, value=st.session_state.ibl, step=10000.0)
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
