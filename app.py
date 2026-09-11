import streamlit as st
from docx import Document
from io import BytesIO
import math
import PyPDF2
from datetime import datetime
import re

# Hugging Face Native Client
from huggingface_hub import InferenceClient

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Agente Sustanciador Híbrido", layout="wide")

# Inicialización de estado
vars_keys = ['nombre', 'cedula', 'ibl', 'smmlv', 'semanas', 'edad', 'genero', 'motivacion_generada', 'hf_token']
for key in vars_keys:
    if key not in st.session_state:
        st.session_state[key] = "" if key in ['nombre', 'cedula', 'motivacion_generada', 'hf_token'] else 0

if 'ibl' not in st.session_state or st.session_state.ibl == 0: st.session_state.ibl = 1500000.0
if 'smmlv' not in st.session_state or st.session_state.smmlv == 0: st.session_state.smmlv = 1300000.0
if 'genero' not in st.session_state or st.session_state.genero == 0: st.session_state.genero = "Femenino"

st.title("⚖️ Agente Sustanciador - Arquitectura Híbrida")
st.markdown("1. Extracción gratuita omnidireccional. | 2. Cálculos matemáticos. | 3. Redacción jurídica (Hugging Face Conversacional).")

# --- 1. MÓDULO DE LECTURA Y EXTRACCIÓN UNIFICADA (TOKEN-FREE) ---
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

def extraer_datos_consolidados(texto):
    # 1. Buscar Cédula
    match_ced = re.search(r'(?:CC\s*No\.?|C\.C\.?|Documento)[:\s]*([\d.,]+)', texto, re.IGNORECASE)
    if match_ced:
        st.session_state.cedula = match_ced.group(1).replace('.', '').replace(',', '').strip()
        
    # 2. Buscar Nombre
    match_nom = re.search(r'(?:señor(?:a)?|Nombre)[:\s]+([A-ZÑÁÉÍÓÚ\s]{5,})(?:,|-|\n|identificado)', texto)
    if match_nom:
        n = match_nom.group(1).strip()
        if "DIRECCION" not in n.upper(): st.session_state.nombre = n
            
    # 3. Determinar Género por contexto
    if re.search(r'\bseñora\b|\bmujer\b|\bfemenino\b', texto, re.IGNORECASE):
        st.session_state.genero = "Femenino"
    elif re.search(r'\bseñor\b|\bhombre\b|\bmasculino\b', texto, re.IGNORECASE):
        st.session_state.genero = "Masculino"
        
    # 4. Buscar Edad o Fecha de Nacimiento
    match_nac = re.search(r'Nacimiento[\s:]*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    if match_nac:
        try:
            fnac = datetime.strptime(match_nac.group(1), '%d/%m/%Y')
            st.session_state.edad = datetime.now().year - fnac.year
        except: pass
    else:
        match_edad = re.search(r'(\d{2})\s*años', texto, re.IGNORECASE)
        if match_edad: st.session_state.edad = int(match_edad.group(1))

    # 5. Buscar Semanas
    match_sem_hl = re.search(r'TOTAL SEMANAS COTIZADAS[\s:]*([\d.,]+)', texto, re.IGNORECASE)
    if match_sem_hl:
        num = match_sem_hl.group(1).replace('.', '').replace(',', '.')
        st.session_state.semanas = int(float(num))
    else:
        match_sem_res = re.search(r'([\d.,]+)\s*semanas', texto, re.IGNORECASE)
        if match_sem_res:
            num = match_sem_res.group(1).replace(',', '').replace('.', '')
            if num.isdigit(): st.session_state.semanas = int(num)
            
    # 6. Buscar IBL
    match_ibl = re.search(r'(?:IBL|Ingreso [B|b]ase de [L|l]iquidación)[^\$]*\$?\s*([\d.,]{6,})', texto)
    if match_ibl:
        num = match_ibl.group(1).replace(',', '').replace('.', '')
        if num.isdigit(): st.session_state.ibl = float(num)

# --- 2. CÁLCULO MATEMÁTICO DE REGLAS (MODIFICADO PARA EXPLICAR LA FÓRMULA) ---
def calcular_derecho_pensional(edad, semanas, genero, ibl, smmlv):
    edad_req = 57 if genero == "Femenino" else 62
    cumple_edad = edad >= edad_req
    cumple_semanas = semanas >= 1300
    
    # Desglose matemático para enviarlo a la IA
    s = ibl / smmlv if smmlv > 0 else 0
    porcentaje_base = max(65.50 - (0.50 * s), 55.5)
    semanas_adicionales = max(0, semanas - 1300)
    grupos_de_50 = math.floor(semanas_adicionales / 50)
    puntos_adicionales = min(grupos_de_50 * 1.5, 15.0) 
    tasa_final = min(porcentaje_base + puntos_adicionales, 80.0) 
    
    mesada_calculada = ibl * (tasa_final / 100)
    
    if mesada_calculada <= smmlv: desc_salud = "4%"
    elif mesada_calculada <= (smmlv * 2): desc_salud = "10%"
    else: desc_salud = "12%"
    
    return {
        "cumple_derecho": cumple_edad and cumple_semanas,
        "faltante_semanas": 1300 - semanas if not cumple_semanas else 0,
        "s_smmlv": s,
        "porcentaje_base": porcentaje_base,
        "grupos_de_50": grupos_de_50,
        "puntos_adicionales": puntos_adicionales,
        "tasa_final": tasa_final,
        "desc_salud": desc_salud,
        "mesada": mesada_calculada
    }

# --- 3. ANÁLISIS A FONDO Y REDACCIÓN VÍA IA CONVERSACIONAL (PROMPT MEJORADO) ---
def redactar_acto_ia(datos_solicitante, calculos, api_key):
    try:
        cliente = InferenceClient(model="Qwen/Qwen2.5-72B-Instruct", token=api_key)
        
        # PROMPT DE DISEÑO LEGAL EXPLICATIVO
        mensajes = [
            {
                "role": "system",
                "content": """Eres un abogado sustanciador de Colpensiones experto en Legal Design. Redacta la sección "CONSIDERANDO" de una resolución administrativa en Colombia.
Tu tono debe ser respetuoso, claro, jurídico y sumamente didáctico para que el ciudadano entienda perfectamente el origen de sus cálculos y no tenga dudas que motiven recursos de apelación.

ESTRUCTURA OBLIGATORIA DEL DOCUMENTO:
1. ANÁLISIS DE REQUISITOS: Cita los requisitos del Art. 33 de la Ley 100 de 1993 (modificado por Ley 797 de 2003) e indica claramente si el ciudadano los cumple.
2. CÁLCULO DE LA TASA DE REEMPLAZO (Si cumple): Explica la fórmula del Art. 34 de la Ley 100 (modificado por Ley 797 de 2003). Desarrolla de forma pedagógica cómo el Ingreso Base de Liquidación (IBL) determina el "Porcentaje Base", y cómo las semanas adicionales a 1300 otorgan "Puntos Adicionales" (hasta el tope normativo).
3. DESCUENTOS DE LEY EN SALUD: Explica obligatoriamente por qué se debe descontar salud, citando expresamente el Art. 143 de la Ley 100 de 1993, y justifica el porcentaje a descontar basándote en los rangos de la Ley 2018 de 2020.
4. NEGATIVA Y ALTERNATIVA (Si NO cumple): Explica empáticamente cuántas semanas le faltan y menciona la opción de la Indemnización Sustitutiva de Vejez si no puede seguir cotizando.
"""
            },
            {
                "role": "user",
                "content": f"""DATOS DEL PETICIONARIO:
Nombre: {datos_solicitante['nombre']}
Cédula: {datos_solicitante['cedula']}
Género: {datos_solicitante['genero']}
Edad: {datos_solicitante['edad']} años
Semanas Cotizadas: {datos_solicitante['semanas']}
IBL Calculado: ${datos_solicitante['ibl']:,.0f} COP

CÁLCULOS MATEMÁTICOS PARA DESGLOSAR Y EXPLICAR EN EL TEXTO:
¿Cumple el derecho?: {"SÍ" if calculos['cumple_derecho'] else "NO"}
Semanas faltantes: {calculos['faltante_semanas']}

DESGLOSE DE FÓRMULA (Solo úsalo si cumple el derecho):
- Proporción IBL vs Salario Mínimo (s): {calculos['s_smmlv']:.2f} salarios mínimos.
- Porcentaje Base calculado: {calculos['porcentaje_base']:.2f}%
- Grupos de 50 semanas adicionales a las 1300: {calculos['grupos_de_50']} grupos.
- Puntos Adicionales ganados: {calculos['puntos_adicionales']:.2f}%
- Tasa de Reemplazo Definitiva (Tope max 80%): {calculos['tasa_final']:.2f}%

DESCUENTO A EXPLICAR:
- Porcentaje a descontar en salud: {calculos['desc_salud']} (Justificar según valor final de la mesada)."""
            }
        ]
        
        # Le permitimos más tokens (2048) para asegurar que el texto explicativo no se corte
        respuesta = cliente.chat_completion(
            messages=mensajes,
            max_tokens=2048,
            temperature=0.3
        )
        
        return respuesta.choices[0].message.content
        
    except Exception as e:
        raise Exception(f"Fallo de conexión con el modelo. Detalle técnico: {str(e)}")

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
st.sidebar.header("⚙️ Configuración del Agente")

token_input = st.sidebar.text_input("Hugging Face API Token:", type="password", help="Pega tu token de Hugging Face aquí.")

if st.sidebar.button("Activar Token", use_container_width=True):
    if token_input.strip():
        st.session_state.hf_token = token_input.strip()
        st.sidebar.success("✅ Token activado para esta sesión.")
    else:
        st.sidebar.error("⚠️ Ingrese un token válido.")

if st.session_state.hf_token:
    st.sidebar.info("Conexión con Hugging Face Activa 🟢")

col1, col2 = st.columns([1, 1.2])

with col1:
    st.subheader("📂 1. Módulos de Carga de Documentos")
    st.info("Sube cualquier documento. El motor unificado extraerá todo automáticamente.")
    
    arc_hl = st.file_uploader("A. Historia Laboral (PDF)", type=["pdf", "txt"], key="hl")
    arc_res = st.file_uploader("B. Resoluciones Previas (PDF, Word)", type=["pdf", "docx", "txt"], key="res")
    arc_pet = st.file_uploader("C. Petición del Ciudadano (PDF, Word)", type=["pdf", "docx", "txt"], key="pet")
    
    if st.button("🔍 Extraer Datos del Solicitante", use_container_width=True):
        with st.spinner("Escaneando documentos..."):
            texto_total = ""
            if arc_hl: texto_total += extraer_texto(arc_hl) + "\n"
            if arc_res: texto_total += extraer_texto(arc_res) + "\n"
            if arc_pet: texto_total += extraer_texto(arc_pet) + "\n"
            
            if texto_total.strip():
                extraer_datos_consolidados(texto_total)
                st.success("Extracción completada. Revise los datos en el panel derecho.")
            else:
                st.warning("No se cargó ningún documento para analizar.")

with col2:
    st.subheader("👤 2. Perfil del Solicitante y Variables")
    
    c_nom, c_ced = st.columns(2)
    with c_nom:
        st.session_state.nombre = st.text_input("Nombre Completo:", value=st.session_state.nombre)
    with c_ced:
        st.session_state.cedula = st.text_input("Cédula:", value=st.session_state.cedula)
        
    c_gen, c_edad, c_sem = st.columns(3)
    with c_gen:
        idx_gen = 0 if st.session_state.genero == "Femenino" else 1
        st.session_state.genero = st.selectbox("Género:", ["Femenino", "Masculino"], index=idx_gen)
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
    
    if st.button("⚖️ Generar Análisis y Motivación Jurídica (IA)", type="primary", use_container_width=True):
        if not st.session_state.hf_token:
            st.error("⚠️ Falta el Token. Por favor, ingresa y activa tu token de Hugging Face en el panel lateral.")
        else:
            with st.spinner("La IA está redactando la motivación basándose en las Reglas de Colpensiones..."):
                calculos = calcular_derecho_pensional(
                    st.session_state.edad, st.session_state.semanas, 
                    st.session_state.genero, st.session_state.ibl, st.session_state.smmlv
                )
                
                datos_sol = {
                    "nombre": st.session_state.nombre,
                    "cedula": st.session_state.cedula,
                    "genero": st.session_state.genero,
                    "edad": st.session_state.edad,
                    "semanas": st.session_state.semanas,
                    "ibl": st.session_state.ibl
                }
                
                try:
                    resultado_ia = redactar_acto_ia(datos_sol, calculos, st.session_state.hf_token)
                    st.session_state.motivacion_generada = resultado_ia
                except Exception as e:
                    st.error(str(e))

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
