import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Importaciones de tus módulos locales (Deben estar en el mismo repo de GitHub)
from data_processor import extraer_tabla_cruda, limpiar_y_estandarizar, aplicar_regla_simultaneidad
from logic import LiquidadorPension

st.set_page_config(page_title="Agente Sustanciador RPM", layout="wide", page_icon="⚖️")

# ==========================================
# 1. MANEJO DE ESTADO PARA STREAMLIT (FUNDAMENTAL)
# ==========================================
if 'df_crudo' not in st.session_state: 
    st.session_state.df_crudo = None
if 'df_final' not in st.session_state: 
    st.session_state.df_final = None
if 'datos_procesados' not in st.session_state:
    st.session_state.datos_procesados = False

# Función auxiliar para requisitos (tomada de tu app 1)[cite: 1]
def get_requisitos_estatus(genero, fecha_estatus, fecha_cumple_edad=None):
    edad_req = 62 if genero == "Masculino" else 57
    if genero == "Masculino":
        semanas_req = 1300
        nota = "Aplica regla general Ley 797/2003 (1300 semanas)."
    else:
        if pd.isna(fecha_estatus):
            anio_actual = datetime.now().year
            anio_cumple = fecha_cumple_edad.year if fecha_cumple_edad else anio_actual
            anio_proy = max(anio_actual, anio_cumple)
            anio_proy = max(2026, anio_proy)
            semanas_req = 1250 if anio_proy == 2026 else max(1000, 1300 - (50 + ((anio_proy - 2026) * 25)))
            nota = f"No consolida estatus. Proyección de {semanas_req} semanas (Sentencia C-197/23)."
        else:
            anio = fecha_estatus.year
            if anio < 2026:
                semanas_req = 1300
                nota = f"Consolidó estatus en {anio}. Exigencia de 1300 semanas."
            else:
                semanas_req = 1250 if anio == 2026 else max(1000, 1300 - (50 + ((anio - 2026) * 25)))
                nota = f"Consolidó estatus en {anio}. Aplica disminución progresiva (Sentencia C-197/23): {semanas_req} semanas."
    return edad_req, semanas_req, nota

# ==========================================
# 2. GENERADOR DEL ACTO ADMINISTRATIVO (RESOLUCIÓN)
# ==========================================
def generar_resolucion_word(datos_afi, liq, req, estatus_cumplido):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    # TIPO DE RESOLUCIÓN
    tipo_resolucion = "RECONOCE" if estatus_cumplido else "NIEGA"
    
    # ENCABEZADO
    p_enc = doc.add_paragraph()
    p_enc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_enc.add_run("ADMINISTRADORA COLOMBIANA DE PENSIONES - COLPENSIONES\n").bold = True
    p_enc.add_run("DIRECCIÓN DE PRESTACIONES ECONÓMICAS\n").bold = True
    p_enc.add_run(f"RESOLUCIÓN NÚMERO _________ DE {datetime.now().year}\n").bold = True
    p_enc.add_run(f"( {datetime.now().strftime('%d/%m/%Y')} )\n\n")
    p_enc.add_run(f"Por la cual se {tipo_resolucion} una Pensión de Vejez a favor de {datos_afi['nombre'].upper()}").bold = True

    # ANTECEDENTES
    doc.add_heading('1. ANTECEDENTES (HECHOS)', level=1)
    doc.add_paragraph(f"Que el(la) señor(a) {datos_afi['nombre'].upper()}, identificado(a) con cédula de ciudadanía No. {datos_afi['cedula']}, elevó solicitud de reconocimiento de Pensión de Vejez ante esta Administradora.")
    doc.add_paragraph(f"Que verificada la historia laboral, se constata que cuenta con un total de {liq['semanas']:,.2f} semanas cotizadas.")

    # CONSIDERACIONES
    doc.add_heading('2. CONSIDERACIONES JURÍDICAS', level=1)
    doc.add_paragraph("Que el artículo 33 de la Ley 100 de 1993, modificado por el artículo 9 de la Ley 797 de 2003, establece que para tener derecho a la Pensión de Vejez se deben reunir los requisitos de edad y semanas.")
    
    if "C-197/23" in req['nota']:
        doc.add_paragraph(f"Que en virtud de la Sentencia C-197 de 2023 de la Corte Constitucional, para el año de causación aplica un requisito de {req['semanas']} semanas.")

    if estatus_cumplido:
        doc.add_paragraph(f"Que el(la) solicitante acreditó el cumplimiento de {req['edad']} años de edad y superó el requisito de densidad de semanas, consolidando el derecho pensional.")
        
        f_data = liq['formula_tasa']
        doc.add_heading('LIQUIDACIÓN DE LA PRESTACIÓN:', level=2)
        doc.add_paragraph(f"Que en cumplimiento del artículo 21 de la Ley 100 de 1993, se determinó como Ingreso Base de Liquidación (IBL) más favorable el correspondiente a {liq['origen_ibl']}, por valor de ${liq['ibl']:,.0f} COP.")
        
        # Integración de la lógica de tu app.py original[cite: 2]
        doc.add_paragraph("La tasa de reemplazo (Art. 34 Ley 100/1993) se establece así:")
        doc.add_paragraph(
            f"1. Proporción IBL sobre SMMLV (s): {f_data['s']:.2f}\n"
            f"2. Porcentaje base [65.50 - (0.50 * s)]: {f_data['tasa_base']:.2f}%\n"
            f"3. Incremento por {f_data['semanas_adicionales']:.2f} semanas adicionales: {f_data['incremento']:.2f}%\n"
            f"4. TASA DE REEMPLAZO FINAL: {f_data['tasa_final']:.2f}%"
        )
        doc.add_paragraph(f"Que en consecuencia, el valor de la mesada pensional asciende a ${liq['mesada']:,.0f} COP mensuales.")
    else:
        doc.add_paragraph(f"Que, analizado el caso concreto, el(la) solicitante NO CUMPLE con los requisitos exigidos por la norma. A la fecha acredita {liq['semanas']:,.2f} semanas de las {req['semanas']} exigidas, por lo cual no hay lugar a reconocer la prestación económica solicitada.")

    # PARTE RESOLUTIVA
    doc.add_heading('RESUELVE:', level=1)
    if estatus_cumplido:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: RECONOCER Y ORDENAR EL PAGO de una Pensión de Vejez a favor de {datos_afi['nombre'].upper()}, en cuantía de ${liq['mesada']:,.0f} COP mensuales.")
        doc.add_paragraph("ARTÍCULO SEGUNDO: DESCUENTOS DE LEY. La mesada pensional estará sujeta a los descuentos obligatorios en materia de salud en los porcentajes de ley.")
    else:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: NEGAR el reconocimiento de la Pensión de Vejez a {datos_afi['nombre'].upper()}, por no cumplir con la densidad de semanas requeridas.")

    doc.add_paragraph("ARTÍCULO FINAL: RECURSOS. Contra la presente Resolución proceden los recursos de reposición y apelación (Ley 1437 de 2011).")
    doc.add_paragraph("\nNOTIFÍQUESE Y CÚMPLASE\n\nFirma Autorizada\nDirección de Prestaciones Económicas\nColpensiones")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 3. INTERFAZ DE STREAMLIT (UI)
# ==========================================
st.title("⚖️ Agente Sustanciador: Actos Administrativos RPM")

with st.sidebar:
    st.header("1. Datos del Expediente")
    cedula = st.text_input("Cédula")
    nombre = st.text_input("Nombre Completo")
    genero = st.radio("Género", ["Masculino", "Femenino"])
    fecha_nac = st.date_input("Nacimiento", min_value=date(1900,1,1), max_value=date.today())
    
    st.header("2. Cargar Historia Laboral")
    uploaded_file = st.file_uploader("Archivo PDF (Historia Laboral)", type="pdf")
    
    if st.button("🔄 Reiniciar Análisis"):
        st.session_state.df_crudo = None
        st.session_state.df_final = None
        st.session_state.datos_procesados = False
        st.rerun()

# ==========================================
# 4. FLUJO DE PROCESAMIENTO
# ==========================================
if uploaded_file and not st.session_state.datos_procesados:
    if st.session_state.df_crudo is None:
        st.session_state.df_crudo = extraer_tabla_cruda(uploaded_file)
    
    if st.session_state.df_crudo is not None:
        df = st.session_state.df_crudo
        st.write("### Mapeo de Columnas de Historia Laboral")
        cols = df.columns.tolist()
        c1, c2, c3, c4 = st.columns(4)
        cd = c1.selectbox("Desde", cols, index=2 if len(cols)>2 else 0)
        ch = c2.selectbox("Hasta", cols, index=3 if len(cols)>3 else 0)
        ci = c3.selectbox("IBC", cols, index=4 if len(cols)>4 else 0)
        cs = c4.selectbox("Semanas", cols, index=len(cols)-1)
        
        if st.button("Aplicar Reglas Normativas y Sustanciar"):
            clean = limpiar_y_estandarizar(df, cd, ch, ci, cs)
            if not clean.empty:
                st.session_state.df_final = aplicar_regla_simultaneidad(clean)
                st.session_state.datos_procesados = True
                st.rerun()

# ==========================================
# 5. RESULTADOS Y GENERACIÓN DE RESOLUCIÓN
# ==========================================
if st.session_state.datos_procesados and st.session_state.df_final is not None:
    df_f = st.session_state.df_final
    liq = LiquidadorPension(df_f, genero, fecha_nac)
    
    fechas = liq.determinar_fechas_clave()
    ibl_10, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "ultimos_10")
    ibl_vida, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "toda_vida")
    ibl_def = max(ibl_10, ibl_vida)
    origen = "Últimos 10 Años" if ibl_10 >= ibl_vida else "Toda la Vida"
    
    total_sem = df_f['Semanas'].sum()
    edad_req, sem_req, nota_req = get_requisitos_estatus(genero, fechas['fecha_estatus'], fechas['fecha_cumple_edad'])
    
    # Se pasa SMLMV estático 2026 para el ejemplo, pero en tu clase esto debería llamarse dinámicamente
    mesada, tasa, info = liq.calcular_tasa_reemplazo_797(ibl_def, total_sem, datetime.now().year, True)
    
    # Simulando el desglosar_formula_tasa[cite: 1] para la resolución
    s_val = ibl_def / 1750905.0 # SMLMV quemado como en app(1).py
    t_base = max(55.0, min(65.5 - (0.5 * s_val), 65.5))
    sem_ad = max(0, total_sem - sem_req)
    bloques = int(sem_ad // 50)
    
    datos_afi = {"nombre": nombre, "cedula": cedula, "fecha_nac": fecha_nac.strftime('%d/%m/%Y')}
    liq_data = {
        "semanas": total_sem, "ibl": ibl_def, "origen_ibl": origen, "mesada": mesada,
        "formula_tasa": {"s": s_val, "tasa_base": t_base, "semanas_adicionales": sem_ad, "bloques": bloques, "incremento": bloques * 1.5, "tasa_final": tasa}
    }
    req_data = {"edad": edad_req, "semanas": sem_req, "nota": nota_req}

    # Dashboard 
    st.success("Análisis Jurídico Completo")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Cumple Edad", "Sí" if fechas['fecha_cumple_edad'] <= datetime.now().date() else "No")
    c2.metric("Semanas Totales", f"{total_sem:,.2f}")
    c3.metric("IBL Aplicado", f"${ibl_def:,.0f}")
    
    estado_texto = "RECONOCIMIENTO" if fechas['tiene_estatus'] else "NEGACIÓN"
    color_estado = "normal" if fechas['tiene_estatus'] else "inverse"
    c4.metric("Sentido del Acto", estado_texto, delta_color=color_estado)

    # Word Builder
    docx_buffer = generar_resolucion_word(datos_afi, liq_data, req_data, fechas['tiene_estatus'])
    
    st.download_button(
        label=f"📥 Descargar Proyecto de Resolución ({estado_texto})",
        data=docx_buffer,
        file_name=f"Resolucion_{estado_texto}_{cedula}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary"
    )
