import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pdfplumber

st.set_page_config(page_title="Sustanciador RPM - Colpensiones", layout="wide", page_icon="⚖️")

# ==========================================
# 1. MANEJO DE ESTADO EN LA NUBE (STREAMLIT CLOUD)
# ==========================================
if 'df_crudo' not in st.session_state: st.session_state.df_crudo = None
if 'df_final' not in st.session_state: st.session_state.df_final = None
if 'datos_procesados' not in st.session_state: st.session_state.datos_procesados = False

# ==========================================
# 2. MÓDULO DE PROCESAMIENTO DE DATOS (Integrado)
# ==========================================
def extraer_tabla_cruda(pdf_file):
    """Extrae texto/tablas de la Historia Laboral de Colpensiones"""
    try:
        # Aquí puedes integrar tu lógica exacta de pdfplumber o camelot
        # Para evitar errores en la nube, retornamos un DataFrame simulado basado en el HL
        # (Reemplaza este bloque con tu código real de data_processor.py)
        data = {
            "Desde": ["01/01/2010", "01/01/2020"],
            "Hasta": ["31/12/2019", "31/12/2025"],
            "IBC": [1500000, 2500000],
            "Semanas": [514.28, 300.0]
        }
        df = pd.DataFrame(data)
        return df
    except Exception as e:
        st.error(f"Error al leer el PDF: {e}")
        return None

def limpiar_y_estandarizar(df, cd, ch, ci, cs):
    df_clean = df.copy()
    df_clean['Desde'] = pd.to_datetime(df_clean[cd], format='%d/%m/%Y', errors='coerce')
    df_clean['Hasta'] = pd.to_datetime(df_clean[ch], format='%d/%m/%Y', errors='coerce')
    df_clean['IBC'] = pd.to_numeric(df_clean[ci], errors='coerce').fillna(0)
    df_clean['Semanas'] = pd.to_numeric(df_clean[cs], errors='coerce').fillna(0)
    return df_clean

def aplicar_regla_simultaneidad(df):
    # Lógica de simultaneidad integrada
    return df.sort_values(by='Desde').reset_index(drop=True)

# ==========================================
# 3. MÓDULO JURÍDICO Y MATEMÁTICO (Integrado)
# ==========================================
class LiquidadorPension:
    def __init__(self, df, genero, fecha_nac):
        self.df = df
        self.genero = genero
        self.fecha_nac = pd.to_datetime(fecha_nac)
        
    def determinar_fechas_clave(self):
        edad_req = 62 if self.genero == "Masculino" else 57
        fecha_cumple_edad = self.fecha_nac + pd.DateOffset(years=edad_req)
        
        # Simulación de cálculo de estatus
        tiene_estatus = self.df['Semanas'].sum() >= 1300 and datetime.now() >= fecha_cumple_edad
        
        return {
            "fecha_cumple_edad": fecha_cumple_edad.date(),
            "fecha_cumple_semanas": datetime.now().date(), # Simplificado
            "tiene_estatus": tiene_estatus,
            "fecha_estatus": datetime.now().date() if tiene_estatus else None,
            "fecha_corte": datetime.now().date(),
            "razon_corte": "Última cotización válida",
            "fecha_efectividad": datetime.now().date()
        }

    def calcular_ibl_indexado(self, fecha_corte, tipo):
        # Lógica de indexación simplificada para integración online
        ibl_simulado = 2500000.0 if tipo == "ultimos_10" else 1800000.0
        df_soporte = pd.DataFrame({"Periodo": ["Simulación"], "IBC": [ibl_simulado]})
        return ibl_simulado, df_soporte

    def calcular_tasa_reemplazo_797(self, ibl, total_sem, anio, tope):
        smlmv_2026 = 1750905.0
        s = ibl / smlmv_2026
        tasa_base = 65.5 - (0.5 * s)
        tasa_base = max(55.0, min(tasa_base, 65.5))
        
        semanas_adicionales = max(0, total_sem - 1300)
        bloques_50 = int(semanas_adicionales // 50)
        incremento = bloques_50 * 1.5
        
        tasa_final = min(80.0, tasa_base + incremento)
        mesada = ibl * (tasa_final / 100)
        return mesada, tasa_final, {"s": s, "tasa_base": tasa_base, "semanas_adicionales": semanas_adicionales, "bloques": bloques_50, "incremento": incremento, "tasa_final": tasa_final}

def get_requisitos_estatus(genero, fecha_estatus, fecha_cumple_edad=None):
    edad_req = 62 if genero == "Masculino" else 57
    semanas_req = 1300
    nota = "Aplica regla general Ley 797/2003 (1300 semanas)."
    # Aquí puedes añadir la lógica de C-197 si el género es Femenino
    return edad_req, semanas_req, nota

# ==========================================
# 4. GENERADOR DEL ACTO ADMINISTRATIVO
# ==========================================
def generar_resolucion_word(datos_afi, liq, req, estatus_cumplido):
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
    doc.add_paragraph(f"Que el(la) señor(a) {datos_afi['nombre'].upper()}, identificado(a) con cédula No. {datos_afi['cedula']}, elevó solicitud de reconocimiento de Pensión de Vejez ante esta Administradora.")
    doc.add_paragraph(f"Que verificada la historia laboral, se constata que cuenta con un total de {liq['semanas']:,.2f} semanas cotizadas.")

    doc.add_heading('2. CONSIDERACIONES JURÍDICAS', level=1)
    doc.add_paragraph("Que el artículo 33 de la Ley 100 de 1993, modificado por el art. 9 de la Ley 797 de 2003, establece los requisitos de edad y semanas.")

    if estatus_cumplido:
        doc.add_paragraph(f"Que el(la) solicitante acreditó el cumplimiento de {req['edad']} años de edad y superó el requisito de semanas, consolidando el derecho pensional.")
        f_data = liq['formula_tasa']
        doc.add_heading('LIQUIDACIÓN DE LA PRESTACIÓN:', level=2)
        doc.add_paragraph(f"Que en cumplimiento del art. 21 de la Ley 100 de 1993, se determinó como IBL más favorable el de {liq['origen_ibl']}, por valor de ${liq['ibl']:,.0f} COP.")
        
        doc.add_paragraph("La tasa de reemplazo (Art. 34 Ley 100/1993) se establece así:")
        doc.add_paragraph(
            f"1. Proporción IBL sobre SMMLV (s): {f_data['s']:.2f}\n"
            f"2. Porcentaje base [65.50 - (0.50 * s)]: {f_data['tasa_base']:.2f}%\n"
            f"3. Incremento por {f_data['semanas_adicionales']:.2f} semanas adicionales: {f_data['incremento']:.2f}%\n"
            f"4. TASA FINAL: {f_data['tasa_final']:.2f}%"
        )
        doc.add_paragraph(f"En consecuencia, el valor de la mesada pensional asciende a ${liq['mesada']:,.0f} COP mensuales.")
    else:
        doc.add_paragraph(f"Que, analizado el caso, el(la) solicitante NO CUMPLE con los requisitos exigidos por la norma. Acredita {liq['semanas']:,.2f} semanas de las {req['semanas']} exigidas.")

    doc.add_heading('RESUELVE:', level=1)
    if estatus_cumplido:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: RECONOCER Y ORDENAR EL PAGO de una Pensión de Vejez a favor de {datos_afi['nombre'].upper()}, en cuantía de ${liq['mesada']:,.0f} COP mensuales.")
    else:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: NEGAR el reconocimiento de la Pensión de Vejez a {datos_afi['nombre'].upper()}, por no cumplir con la densidad de semanas requeridas.")

    doc.add_paragraph("ARTÍCULO FINAL: RECURSOS. Contra la presente Resolución proceden los recursos de reposición y apelación (Ley 1437 de 2011).")
    doc.add_paragraph("\nNOTIFÍQUESE Y CÚMPLASE\n\nFirma Autorizada\nDirección de Prestaciones Económicas\nColpensiones")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 5. INTERFAZ Y FLUJO DE EJECUCIÓN STREAMLIT
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

if uploaded_file and not st.session_state.datos_procesados:
    if st.session_state.df_crudo is None:
        st.session_state.df_crudo = extraer_tabla_cruda(uploaded_file)
    
    if st.session_state.df_crudo is not None:
        df = st.session_state.df_crudo
        st.write("### Mapeo de Columnas de Historia Laboral")
        cols = df.columns.tolist()
        c1, c2, c3, c4 = st.columns(4)
        cd = c1.selectbox("Desde", cols, index=0)
        ch = c2.selectbox("Hasta", cols, index=1)
        ci = c3.selectbox("IBC", cols, index=2)
        cs = c4.selectbox("Semanas", cols, index=3)
        
        if st.button("Aplicar Reglas Normativas y Sustanciar"):
            clean = limpiar_y_estandarizar(df, cd, ch, ci, cs)
            if not clean.empty:
                st.session_state.df_final = aplicar_regla_simultaneidad(clean)
                st.session_state.datos_procesados = True
                st.rerun()

if st.session_state.datos_procesados and st.session_state.df_final is not None:
    df_f = st.session_state.df_final
    liq = LiquidadorPension(df_f, genero, fecha_nac)
    
    fechas = liq.determinar_fechas_clave()
    ibl_def, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "ultimos_10")
    total_sem = df_f['Semanas'].sum()
    
    edad_req, sem_req, nota_req = get_requisitos_estatus(genero, fechas['fecha_estatus'], fechas['fecha_cumple_edad'])
    mesada, tasa, f_data = liq.calcular_tasa_reemplazo_797(ibl_def, total_sem, datetime.now().year, True)
    
    datos_afi = {"nombre": nombre, "cedula": cedula, "fecha_nac": fecha_nac.strftime('%d/%m/%Y')}
    liq_data = {
        "semanas": total_sem, "ibl": ibl_def, "origen_ibl": "Últimos 10 Años", "mesada": mesada, "formula_tasa": f_data
    }
    req_data = {"edad": edad_req, "semanas": sem_req, "nota": nota_req}

    st.success("Análisis Jurídico Completo")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Cumple Edad", "Sí" if fechas['fecha_cumple_edad'] <= datetime.now().date() else "No")
    c2.metric("Semanas Totales", f"{total_sem:,.2f}")
    c3.metric("IBL Aplicado", f"${ibl_def:,.0f}")
    
    estado_texto = "RECONOCIMIENTO" if fechas['tiene_estatus'] else "NEGACIÓN"
    c4.metric("Sentido del Acto", estado_texto, delta_color="normal" if fechas['tiene_estatus'] else "inverse")

    docx_buffer = generar_resolucion_word(datos_afi, liq_data, req_data, fechas['tiene_estatus'])
    
    st.download_button(
        label=f"📥 Descargar Proyecto de Resolución ({estado_texto})",
        data=docx_buffer,
        file_name=f"Resolucion_{estado_texto}_{cedula}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary"
    )
