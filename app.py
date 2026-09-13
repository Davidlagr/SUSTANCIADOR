import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==========================================
# IMPORTACIONES REALES DE TUS MÓDULOS
# ==========================================
from data_processor import extraer_tabla_cruda, limpiar_y_estandarizar, aplicar_regla_simultaneidad
from logic import LiquidadorPension

st.set_page_config(page_title="Sustanciador Pro - Colpensiones", layout="wide", page_icon="⚖️")

# ==========================================
# MANEJO DE ESTADO EN SESIÓN
# ==========================================
if 'df_crudo' not in st.session_state: st.session_state.df_crudo = None
if 'df_final' not in st.session_state: st.session_state.df_final = None
if 'analisis_completado' not in st.session_state: st.session_state.analisis_completado = False

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
# GENERADOR DEL ACTO ADMINISTRATIVO
# ==========================================
def generar_resolucion_word(datos_afi, liq, req, estatus_cumplido, archivos_cargados):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    tipo_resolucion = "RECONOCE" if estatus_cumplido else "NIEGA"
    
    p_enc = doc.add_paragraph()
    p_enc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_enc.add_run("ADMINISTRADORA COLOMBIANA DE PENSIONES - COLPENSIONES\n").bold = True
    p_enc.add_run("DIRECCIÓN DE PRESTACIONES ECONÓMICAS\n").bold = True
    p_enc.add_run(f"RESOLUCIÓN NÚMERO _________ DE {datetime.now().year}\n").bold = True
    p_enc.add_run(f"( {datetime.now().strftime('%d/%m/%Y')} )\n\n")
    p_enc.add_run(f"Por la cual se {tipo_resolucion} una Pensión de Vejez a favor de {datos_afi['nombre'].upper()}").bold = True

    doc.add_heading('1. ANTECEDENTES (HECHOS)', level=1)
    
    # Inyección dinámica de documentos base
    texto_peticion = "radicó petición formal" if archivos_cargados['peticion'] else "elevó solicitud"
    doc.add_paragraph(f"Que el(la) señor(a) {datos_afi['nombre'].upper()}, identificado(a) con cédula No. {datos_afi['cedula']}, {texto_peticion} ante esta Administradora para el estudio de su prestación económica.")
    
    if archivos_cargados['resoluciones'] > 0:
        doc.add_paragraph(f"Que obran en el expediente {archivos_cargados['resoluciones']} acto(s) administrativo(s) previo(s) proferido(s) por la entidad, los cuales fueron objeto de análisis integral para el presente pronunciamiento.")
    
    doc.add_paragraph(f"Que procesada y analizada la Historia Laboral aportada, se constata que el(la) afiliado(a) cuenta con un total de {liq['semanas']:,.2f} semanas cotizadas válidas al Sistema General de Pensiones.")

    doc.add_heading('2. CONSIDERACIONES JURÍDICAS Y CASO CONCRETO', level=1)
    doc.add_paragraph("Que el artículo 33 de la Ley 100 de 1993, modificado por el art. 9 de la Ley 797 de 2003, establece los requisitos concurrentes de edad y densidad de semanas.")

    if "C-197/23" in req['nota']:
        doc.add_paragraph(f"Que en aplicación irrestricta de la Sentencia C-197 de 2023 proferida por la Corte Constitucional, el requisito de semanas exigido para la causación del derecho en el presente caso es de {req['semanas']} semanas.")

    if estatus_cumplido:
        doc.add_paragraph(f"Que el(la) solicitante acreditó el cumplimiento de {req['edad']} años de edad y superó el umbral de semanas exigidas, consolidando legítimamente el derecho pensional.")
        f_data = liq['formula_tasa']
        
        doc.add_heading('LIQUIDACIÓN TÉCNICA DE LA PRESTACIÓN:', level=2)
        doc.add_paragraph(f"Que en cumplimiento del mandato contenido en el artículo 21 de la Ley 100 de 1993, y aplicando el principio de favorabilidad, se determinó como Ingreso Base de Liquidación (IBL) el correspondiente a {liq['origen_ibl']}, por un valor indexado de ${liq['ibl']:,.0f} COP.")
        
        doc.add_paragraph("De conformidad con el artículo 34 de la Ley 100 de 1993, la tasa de reemplazo se liquida mediante la siguiente fórmula:")
        doc.add_paragraph(
            f"1. Proporción del IBL respecto al SMMLV (s): {f_data['s']:.2f}\n"
            f"2. Porcentaje base [65.50 - (0.50 * s)]: {f_data['tasa_base']:.2f}%\n"
            f"3. Incremento por {f_data['semanas_adicionales']:.2f} semanas adicionales ({f_data['bloques']} bloque(s)): +{f_data['incremento']:.2f}%\n"
            f"4. TASA DE REEMPLAZO FINAL DEFINITIVA: {f_data['tasa_final']:.2f}%"
        )
        doc.add_paragraph(f"En consecuencia, el valor de la mesada pensional asciende a la suma de ${liq['mesada']:,.0f} COP mensuales.")
    else:
        doc.add_paragraph(f"Que, descendiendo al caso concreto, el(la) solicitante NO CUMPLE con los requisitos exigidos. Si bien acredita la edad requerida, a la fecha de corte cuenta únicamente con {liq['semanas']:,.2f} semanas, siendo insuficientes frente a las {req['semanas']} semanas exigidas por el ordenamiento.")

    doc.add_heading('RESUELVE:', level=1)
    if estatus_cumplido:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: RECONOCER Y ORDENAR EL PAGO de una Pensión de Vejez a favor de {datos_afi['nombre'].upper()}, identificado(a) con C.C. {datos_afi['cedula']}, en cuantía de ${liq['mesada']:,.0f} COP mensuales.")
        doc.add_paragraph("ARTÍCULO SEGUNDO: DESCUENTOS DE LEY. La presente prestación estará sujeta a los descuentos obligatorios para aportes al Sistema General de Seguridad Social en Salud.")
    else:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: NEGAR el reconocimiento de la Pensión de Vejez a favor de {datos_afi['nombre'].upper()}, identificado(a) con C.C. {datos_afi['cedula']}, por las razones expuestas en la parte motiva del presente acto.")

    doc.add_paragraph("ARTÍCULO FINAL: RECURSOS. Contra la presente Resolución proceden los recursos de reposición y en subsidio el de apelación de conformidad con lo reglado en el Código de Procedimiento Administrativo y de lo Contencioso Administrativo (Ley 1437 de 2011).")
    doc.add_paragraph("\nNOTIFÍQUESE Y CÚMPLASE\n\n\nFirma Autorizada\nDirección de Prestaciones Económicas\nAdministradora Colombiana de Pensiones - Colpensiones")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# INTERFAZ Y FLUJO DE EJECUCIÓN 
# ==========================================
st.title("⚖️ Agente Sustanciador: Actos Administrativos RPM")
st.markdown("Motor analítico integrado con lector de Historia Laboral y parametrización de actos.")

# --- SECCIÓN 1: DATOS Y CARGA DOCUMENTAL ---
with st.sidebar:
    st.header("1. Datos del Ciudadano")
    cedula = st.text_input("Número de Cédula")
    nombre = st.text_input("Nombre Completo")
    genero = st.radio("Género", ["Masculino", "Femenino"])
    fecha_nac = st.date_input("Fecha de Nacimiento", min_value=date(1900,1,1), max_value=date.today())
    
    st.divider()
    st.header("2. Expediente Electrónico")
    hl_file = st.file_uploader("1. Historia Laboral (Obligatorio)", type="pdf", help="Cargue el extracto emitido por Colpensiones")
    resoluciones_files = st.file_uploader("2. Histórico de Resoluciones", type="pdf", accept_multiple_files=True, help="Actos administrativos previos que deban revocarse o mencionarse")
    peticion_file = st.file_uploader("3. Petición del Ciudadano", type="pdf", help="Escrito libre, derecho de petición o formulario")
    
    st.divider()
    if st.button("🔄 Reiniciar Análisis", use_container_width=True):
        st.session_state.df_crudo = None
        st.session_state.df_final = None
        st.session_state.analisis_completado = False
        st.rerun()

# --- SECCIÓN 2: PROCESAMIENTO DE HISTORIA LABORAL ---
if hl_file and not st.session_state.analisis_completado:
    st.info("Extrayendo matrices de la Historia Laboral usando el motor de data_processor.py...")
    
    if st.session_state.df_crudo is None:
        st.session_state.df_crudo = extraer_tabla_cruda(hl_file)
    
    if st.session_state.df_crudo is not None and not st.session_state.df_crudo.empty:
        df = st.session_state.df_crudo
        st.write("### Mapeo de Variables Jurídicas")
        cols = df.columns.tolist()
        c1, c2, c3, c4 = st.columns(4)
        cd = c1.selectbox("Columna 'Desde'", cols, index=2 if len(cols)>2 else 0)
        ch = c2.selectbox("Columna 'Hasta'", cols, index=3 if len(cols)>3 else 0)
        ci = c3.selectbox("Columna 'IBC'", cols, index=4 if len(cols)>4 else 0)
        cs = c4.selectbox("Columna 'Semanas'", cols, index=len(cols)-1)
        
        if st.button("Ejecutar Saneamiento y Aplicar Simultaneidad", type="primary"):
            clean = limpiar_y_estandarizar(df, cd, ch, ci, cs)
            if not clean.empty:
                st.session_state.df_final = aplicar_regla_simultaneidad(clean)
                st.session_state.analisis_completado = True
                st.rerun()
            else:
                st.error("Error validando las columnas tras la limpieza.")

# --- SECCIÓN 3: SUBSUNCIÓN NORMATIVA Y PROYECCIÓN ---
if st.session_state.analisis_completado and st.session_state.df_final is not None:
    df_f = st.session_state.df_final
    
    # 1. Instanciamos tu clase original
    liq = LiquidadorPension(df_f, genero, fecha_nac)
    fechas = liq.determinar_fechas_clave()
    
    # 2. Análisis del IBL más favorable
    ibl_10, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "ultimos_10")
    ibl_vida, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "toda_vida")
    ibl_def = max(ibl_10, ibl_vida)
    origen_ibl = "Últimos 10 Años" if ibl_10 >= ibl_vida else "Toda la Vida"
    
    # 3. Determinación de la Tasa
    total_sem = df_f['Semanas'].sum()
    edad_req, sem_req, nota_req = get_requisitos_estatus(genero, fechas['fecha_estatus'], fechas['fecha_cumple_edad'])
    
    # Asumiendo que SMLMV y tope se pasan como en app (1).py
    mesada, tasa, info = liq.calcular_tasa_reemplazo_797(ibl_def, total_sem, datetime.now().year, True)
    
    # 4. Estructuración de datos para el Word
    smlmv_ref = 1750905.0 # Mantenemos la base 2026 de app(1)
    s = ibl_def / smlmv_ref if smlmv_ref > 0 else 1
    tasa_base = max(55.0, min(65.5 - (0.5 * s), 65.5))
    sem_adicionales = max(0, total_sem - sem_req)
    bloques = int(sem_adicionales // 50)
    incremento = min(bloques * 1.5, 15.0)

    datos_afi = {"nombre": nombre, "cedula": cedula, "fecha_nac": fecha_nac.strftime('%d/%m/%Y')}
    liq_data = {
        "semanas": total_sem, "ibl": ibl_def, "origen_ibl": origen_ibl, "mesada": mesada,
        "formula_tasa": {"s": s, "tasa_base": tasa_base, "semanas_adicionales": sem_adicionales, "bloques": bloques, "incremento": incremento, "tasa_final": tasa}
    }
    req_data = {"edad": edad_req, "semanas": sem_req, "nota": nota_req}
    archivos_cargados = {"resoluciones": len(resoluciones_files) if resoluciones_files else 0, "peticion": True if peticion_file else False}

    st.success("✅ Análisis Jurídico y Financiero Completado")
    
    # Dashboard de Resultados
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Cumple Edad", "Sí" if fechas['fecha_cumple_edad'] <= datetime.now().date() else "No")
    c2.metric("Semanas Válidas", f"{total_sem:,.2f}")
    c3.metric("IBL Favorable", f"${ibl_def:,.0f}", origen_ibl)
    
    estado_texto = "RECONOCIMIENTO" if fechas['tiene_estatus'] else "NEGACIÓN"
    c4.metric("Sentido del Acto", estado_texto, delta_color="normal" if fechas['tiene_estatus'] else "inverse")

    # Generación Documental
    st.divider()
    docx_buffer = generar_resolucion_word(datos_afi, liq_data, req_data, fechas['tiene_estatus'], archivos_cargados)
    
    st.download_button(
        label=f"📥 Descargar Resolución de Fondo ({estado_texto})",
        data=docx_buffer,
        file_name=f"Resolucion_{estado_texto}_{cedula}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True
    )
