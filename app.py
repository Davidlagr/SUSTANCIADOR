import streamlit as st
import docx
from docx import Document
import io
import pandas as pd

st.set_page_config(page_title="Asistente de Resoluciones - Liquidador Pensional Pro", layout="wide")

st.title("⚖️ Asistente de Sustanciación de Resoluciones")
st.subheader("Módulo de Actos Administrativos - Prestaciones Económicas")

# 1. Datos del Asegurado / Solicitante
st.header("1. Datos del Asegurado / Solicitante")
col1, col2 = st.columns(2)
with col1:
    nombres = st.text_input("Nombres")
    apellidos = st.text_input("Apellidos")
with col2:
    identificacion = st.text_input("Número de Identificación")
    nacimiento = st.date_input("Fecha de Nacimiento", format="DD/MM/YYYY")

# 2. Antecedentes
st.header("2. Antecedentes")
antecedentes = st.text_area("Pegue aquí los antecedentes del trámite (Radicados, historias laborales, solicitudes previas):", height=150)

# 3. Banco de Motivaciones (Consideraciones Legales)
st.header("3. Banco de Motivaciones Legales")
banco_motivaciones = {
    "Sustitución Pensional - Marco General (Ley 797/2003)": "Que el artículo 47 de la citada Ley 100 de 1993, modificado por el artículo 13 de la Ley 797 de 2003 establece como beneficiarios de la pensión de sobrevivientes: a) En forma vitalicia, el cónyuge o la compañera o compañero permanente supérstite, siempre y cuando dicho beneficiario, a la fecha del fallecimiento del causante, tenga 30 o más años de edad...",
    "Sustitución Pensional - Hijo Inválido (Dependencia)": "Que frente a la dependencia económica de los hijos inválidos, el Memorando OAL-001-2022 del 13 de enero de 2022 emitido por la Oficina Asesora de Asuntos Legales, establece que para acreditar la dependencia económica del hijo inválido NO se requiere probar la carencia total y absoluta de medios económicos, existiendo subordinación cuando la persona requiera total o parcialmente de los ingresos de otra para cubrir sus necesidades básicas...",
    "Pensión de Vejez - Fórmula Decreciente": "Que para establecer el monto de la pensión, se aplicará la fórmula decreciente estipulada en el artículo 10 de la Ley 797 de 2003, que modificó el artículo 34 de la Ley 100 de 1993, determinando la tasa de reemplazo según el número de salarios mínimos del IBL y las semanas adicionales cotizadas...",
    "Aplicación de IPC (Mesada)": "Que el valor de la mesada será reajustado al momento del pago, según el Índice de Precios al Consumidor certificado por el DANE, de acuerdo con lo establecido por el artículo 14 de la Ley 100 de 1993."
}

motivaciones_seleccionadas = st.multiselect(
    "Seleccione las motivaciones a incluir en las CONSIDERACIONES:",
    options=list(banco_motivaciones.keys())
)

texto_motivaciones = "\n\n".join([banco_motivaciones[m] for m in motivaciones_seleccionadas])
if texto_motivaciones:
    st.info("Motivaciones seleccionadas preparadas para el acto administrativo.")

# 4. Fórmula Decreciente (Ley 797 de 2003)
st.header("4. Cálculo y Explicación de Fórmula Decreciente (Ley 797)")
col3, col4, col5 = st.columns(3)
with col3:
    semanas = st.number_input("Total Semanas Cotizadas", min_value=0.0, value=1300.0, step=1.0)
with col4:
    ibl = st.number_input("IBL Calculado ($)", min_value=0.0, value=1300000.0, step=10000.0)
with col5:
    smlmv = st.number_input("SMLMV del año base ($)", min_value=1.0, value=1300000.0, step=10000.0)

generar_formula = st.checkbox("Generar anexo explicativo de la fórmula decreciente")
explicacion_formula = ""

if generar_formula and smlmv > 0:
    # Lógica de la fórmula
    s = ibl / smlmv
    r = 65.5 - (0.5 * s)
    r = max(55.0, min(r, 65.5)) # Límites de r entre 55 y 65.5
    
    semanas_adicionales = max(0, semanas - 1300)
    bloques_50 = int(semanas_adicionales // 50)
    incremento = bloques_50 * 1.5
    
    tasa_final = r + incremento
    tasa_final = min(80.0, tasa_final) # Límite máximo 80%
    
    mesada = ibl * (tasa_final / 100.0)
    mesada_final = max(smlmv, mesada) # Mínimo vital
    
    explicacion_formula = f"""Para la liquidación de la prestación, se aplica la fórmula establecida en la Ley 797 de 2003, artículo 10 (r = 65.5 - 0.5s):
    
- Ingreso Base de Liquidación (IBL): ${ibl:,.2f}
- Salarios Mínimos del IBL (s): {s:.2f} SMLMV
- Tasa Base de Reemplazo (r): {r:.2f}%
- Total Semanas Reconocidas: {semanas} (Semanas adicionales a 1300: {semanas_adicionales})
- Incremento por semanas adicionales (1.5% por cada 50 semanas): {incremento:.2f}%
- Tasa de Reemplazo Final Aplicada: {tasa_final:.2f}% (Máximo 80%)
- Valor de la Mesada Pensional Resultante: ${mesada_final:,.2f}
"""
    st.text_area("Vista previa de la fórmula a inyectar:", value=explicacion_formula, height=200)

# 5. Generación del Acto Administrativo
st.header("5. Plantilla y Generación de la Resolución")
uploaded_file = st.file_uploader("Suba la plantilla en Word (.docx)", type="docx")

if st.button("Generar Resolución Estructurada", type="primary"):
    if uploaded_file is not None:
        # Cargar el documento
        doc = Document(uploaded_file)
        
        # Diccionario de reemplazo
        reemplazos = {
            "{{NOMBRES}}": nombres.upper(),
            "{{APELLIDOS}}": apellidos.upper(),
            "{{IDENTIFICACION}}": identificacion,
            "{{NACIMIENTO}}": str(nacimiento),
            "{{ANTECEDENTES}}": antecedentes,
            "{{MOTIVACIONES}}": texto_motivaciones,
            "{{FORMULA}}": explicacion_formula
        }
        
        # Iterar sobre párrafos y reemplazar
        for p in doc.paragraphs:
            for key, value in reemplazos.items():
                if key in p.text:
                    p.text = p.text.replace(key, value)
                    
        # Iterar sobre tablas (si la plantilla tiene tablas de historia laboral)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in reemplazos.items():
                        if key in cell.text:
                            cell.text = cell.text.replace(key, value)

        # Guardar en memoria para descarga
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success("¡Acto administrativo generado exitosamente!")
        st.download_button(
            label="📥 Descargar Resolución Sustanciada (.docx)",
            data=buffer,
            file_name=f"Resolucion_{identificacion}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("Por favor, suba una plantilla en formato .docx para proceder.")
