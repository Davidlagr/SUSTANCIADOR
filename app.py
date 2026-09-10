import streamlit as st
from docx import Document
from io import BytesIO
import math

# Configuración de la página
st.set_page_config(page_title="Agente Sustanciador - RPM", layout="wide")

st.title("Agente Sustanciador: Proyección de Prestaciones Económicas")
st.markdown("Generador de motivaciones jurídicas fundamentadas en la Ley 100 de 1993 y Ley 797 de 2003.")

# --- Funciones de Cálculo ---
def calcular_tasa_reemplazo(ibl, smmlv, semanas_totales):
    # 1. Calcular número de salarios mínimos (s)
    s = ibl / smmlv
    
    # 2. Fórmula base: r = 65.50 - 0.50 * s
    porcentaje_base = 65.50 - (0.50 * s)
    if porcentaje_base < 55.5:
        porcentaje_base = 55.5
        
    # 3. Cálculo de puntos adicionales por semanas
    semanas_adicionales = max(0, semanas_totales - 1300)
    grupos_de_50 = math.floor(semanas_adicionales / 50)
    
    # Tope legal de 15 puntos (máximo 500 semanas adicionales)
    puntos_adicionales = min(grupos_de_50 * 1.5, 15.0)
    
    # 4. Tasa final con tope del 80%
    tasa_final = min(porcentaje_base + puntos_adicionales, 80.0)
    
    return s, porcentaje_base, grupos_de_50, puntos_adicionales, tasa_final

# --- Función para generar el documento Word (.docx) ---
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

# --- Interfaz de Usuario (Inputs) ---
st.sidebar.header("Datos del Expediente")
tipo_prestacion = st.sidebar.selectbox("Seleccione el trámite:", ["Pensión de Vejez (Ley 797/2003)"])

if tipo_prestacion == "Pensión de Vejez (Ley 797/2003)":
    st.sidebar.subheader("Variables de Liquidación")
    ibl_input = st.sidebar.number_input("Ingreso Base de Liquidación (IBL) en COP:", min_value=0.0, value=1500000.0, step=100000.0)
    smmlv_input = st.sidebar.number_input("SMMLV del año de causación en COP:", min_value=0.0, value=1300000.0, step=10000.0)
    semanas_input = st.sidebar.number_input("Total Semanas Cotizadas:", min_value=0, value=1300, step=1)
    
    st.subheader("Simulación y Motivación Jurídica")
    
    if st.button("Generar Motivación"):
        # Ejecutar cálculos
        s, p_base, grupos, p_add, t_final = calcular_tasa_reemplazo(ibl_input, smmlv_input, semanas_input)
        
        # Estructurar el texto motivacional
        motivacion = f"""
CONSIDERANDO:

Que de conformidad con el artículo 33 de la Ley 100 de 1993, modificado por el artículo 9 de la Ley 797 de 2003, para tener derecho a la Pensión de Vejez es necesario acreditar las edades establecidas en la norma y un mínimo de 1.300 semanas de cotización.

Que el(la) afiliado(a) acredita un total de {semanas_input} semanas cotizadas al Sistema General de Pensiones, contabilizadas de acuerdo con el Parágrafo 2 del artículo 33 de la Ley 100 de 1993, cumpliendo con el requisito de densidad de semanas.

Que en cumplimiento del artículo 21 de la Ley 100 de 1993, se determinó el Ingreso Base de Liquidación (IBL) en la suma de ${ibl_input:,.2f} COP.

LIQUIDACIÓN DE LA TASA DE REEMPLAZO (Monto de la Pensión):
De conformidad con el artículo 34 de la Ley 100 de 1993, modificado por el artículo 10 de la Ley 797 de 2003, el monto mensual de la pensión se determina mediante una fórmula decreciente y la suma de puntos adicionales por semanas extra cotizadas. El cálculo se sustenta así:

1. Proporción del IBL respecto al salario mínimo (s):
Se divide el IBL (${ibl_input:,.2f}) entre el SMMLV del año respectivo (${smmlv_input:,.2f}), arrojando un factor (s) de {s:.2f} salarios mínimos.

2. Porcentaje Inicial:
Aplicando la fórmula legal r = 65.50 - 0.50s:
r = 65.50 - (0.50 * {s:.2f}) = {p_base:.2f}%

3. Puntos adicionales por semanas excedentes:
El afiliado acreditó {semanas_input} semanas. Al restar las 1.300 semanas mínimas exigidas, se obtiene un excedente de {max(0, semanas_input - 1300)} semanas. 
La norma establece un incremento del 1.5% por cada 50 semanas adicionales (hasta un máximo de 15 puntos, equivalentes a 500 semanas). 
El afiliado cuenta con {grupos} bloque(s) completo(s) de 50 semanas.
Incremento adicional = {grupos} * 1.5% = {p_add:.2f}%.

4. Tasa de Reemplazo Definitiva:
Sumando el porcentaje inicial ({p_base:.2f}%) y los puntos adicionales ({p_add:.2f}%), se establece una tasa de reemplazo total del {t_final:.2f}% (Recordando que la ley impone un tope máximo del 80%).

En consecuencia, el valor de la mesada pensional corresponderá al {t_final:.2f}% del IBL, quedando sujeta a los descuentos de Ley en materia de salud (Art. 143 Ley 100/93 y Art. 1 Ley 2018/2020) correspondientes según el rango de la mesada final.
"""
        
        st.text_area("Vista previa de la Motivación:", motivacion, height=450)
        
        # Botón de descarga Word
        word_file = generar_word(motivacion)
        st.download_button(
            label="📄 Descargar Motivación en Word (.docx)",
            data=word_file,
            file_name="Motivacion_Resolucion_RPM.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
