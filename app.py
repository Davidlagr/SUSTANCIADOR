import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==========================================
# IMPORTACIÓN DE MÓDULOS DEL REPOSITORIO
# ==========================================
# Estos archivos (data_processor.py y logic.py) deben estar en el mismo repositorio de GitHub
try:
    from data_processor import extraer_tabla_cruda, limpiar_y_estandarizar, aplicar_regla_simultaneidad
    from logic import LiquidadorPension
    MODULOS_CARGADOS = True
except ImportError:
    MODULOS_CARGADOS = False

st.set_page_config(page_title="Sustanciador RPM", layout="wide", page_icon="⚖️")

# ==========================================
# GENERADOR DINÁMICO DEL ACTO ADMINISTRATIVO
# ==========================================
def generar_resolucion_word(datos_afiliado, liq_data, req_data, cumple_requisitos):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    # 1. ENCABEZADO
    p_encabezado = doc.add_paragraph()
    p_encabezado.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_encabezado.add_run("ADMINISTRADORA COLOMBIANA DE PENSIONES - COLPENSIONES\n").bold = True
    p_encabezado.add_run("DIRECCIÓN DE PRESTACIONES ECONÓMICAS\n").bold = True
    p_encabezado.add_run(f"RESOLUCIÓN NÚMERO _________ DE {datetime.now().year}\n").bold = True
    p_encabezado.add_run(f"( {datetime.now().strftime('%d/%m/%Y')} )\n\n")
    
    if cumple_requisitos:
        p_encabezado.add_run(f"Por la cual se reconoce una Pensión de Vejez a favor de {datos_afiliado['nombre'].upper()}").bold = True
    else:
        p_encabezado.add_run(f"Por la cual se NIEGA una Pensión de Vejez a favor de {datos_afiliado['nombre'].upper()}").bold = True

    # 2. ANTECEDENTES
    doc.add_heading('1. ANTECEDENTES', level=1)
    doc.add_paragraph(f"Que el(la) señor(a) {datos_afiliado['nombre'].upper()}, identificado(a) con cédula de ciudadanía No. {datos_afiliado['cedula']}, elevó solicitud de reconocimiento de Pensión de Vejez.")
    doc.add_paragraph(f"Que revisada la historia laboral válida y aplicadas las reglas de simultaneidad, se constata que cuenta con un total de {liq_data['semanas']:,.2f} semanas cotizadas al Sistema General de Pensiones.")

    # 3. CONSIDERACIONES
    doc.add_heading('2. CONSIDERACIONES', level=1)
    doc.add_paragraph("Que el artículo 33 de la Ley 100 de 1993, modificado por el artículo 9 de la Ley 797 de 2003, exige para la Pensión de Vejez el cumplimiento de una edad mínima y 1.300 semanas de cotización.")
    
    if "Sentencia C-197/23" in req_data['nota']:
        doc.add_paragraph(f"Que en virtud de la Sentencia C-197 de 2023 de la Corte Constitucional, se aplica la disminución progresiva del número de semanas exigidas, siendo aplicable un requisito de {req_data['semanas_exigidas']} semanas.")

    if cumple_requisitos:
        doc.add_paragraph(f"Que el(la) solicitante nació el {datos_afiliado['fecha_nac'].strftime('%d/%m/%Y')}, cumpliendo la edad requerida ({req_data['edad_exigida']} años) y acreditando la densidad de semanas exigidas.")
        
        # LIQUIDACIÓN (Solo si cumple)
        doc.add_heading('LIQUIDACIÓN DE LA PRESTACIÓN:', level=2)
        f_data = liq_data['formula_tasa']
        doc.add_paragraph(f"Que en estricto cumplimiento del artículo 21 de la Ley 100 de 1993, el Ingreso Base de Liquidación (IBL) más favorable corresponde a {liq_data['origen_ibl']}, por valor de ${liq_data['ibl']:,.0f} COP.")
        doc.add_paragraph("De conformidad con el artículo 34 de la Ley 100 de 1993 (modificado por la Ley 797 de 2003):")
        doc.add_paragraph(
            f"1. Proporción del IBL respecto al SMMLV (s): {f_data['s']:.2f}\n"
            f"2. Porcentaje base [65.50 - (0.50 * s)]: {f_data['tasa_base']:.2f}%\n"
            f"3. Puntos por {f_data['semanas_adicionales']:.2f} semanas excedentes: {f_data['incremento']:.2f}%\n"
            f"4. Tasa de reemplazo definitiva: {f_data['tasa_final']:.2f}%"
        )
        doc.add_paragraph(f"En consecuencia, el valor de la mesada pensional asciende a ${liq_data['mesada']:,.0f} COP mensuales.")
    else:
        doc.add_paragraph(f"Que el(la) solicitante nació el {datos_afiliado['fecha_nac'].strftime('%d/%m/%Y')}. Si bien acredita la edad de {req_data['edad_exigida']} años (o está en vías de cumplirla), cuenta únicamente con {liq_data['semanas']:,.2f} semanas cotizadas, de las {req_data['semanas_exigidas']} exigidas por la normatividad vigente.")
        doc.add_paragraph("En consecuencia, al no reunir la densidad de semanas requeridas, resulta jurídicamente inviable acceder al reconocimiento pensional, dejando a salvo el derecho a solicitar la Indemnización Sustitutiva si declara la imposibilidad de seguir cotizando.")

    # 4. PARTE RESOLUTIVA
    doc.add_heading('RESUELVE:', level=1)
    if cumple_requisitos:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: RECONOCER Y ORDENAR EL PAGO de una Pensión de Vejez a favor de {datos_afiliado['nombre'].upper()}, en cuantía de ${liq_data['mesada']:,.0f} COP mensuales.")
        doc.add_paragraph("ARTÍCULO SEGUNDO: DESCUENTOS DE LEY. La mesada pensional estará sujeta a los descuentos obligatorios en materia de salud.")
    else:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: NEGAR el reconocimiento de la Pensión de Vejez a favor de {datos_afiliado['nombre'].upper()}, por no acreditar el requisito de semanas cotizadas.")

    doc.add_paragraph("ARTÍCULO TERCERO: RECURSOS. Contra la presente Resolución proceden los recursos de reposición y apelación en los términos de la Ley 1437 de 2011 (CPACA).")
    doc.add_paragraph("\n\nNOTIFÍQUESE Y CÚMPLASE\n\nFirma Autorizada\nDirección de Prestaciones Económicas\nColpensiones")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# INTERFAZ Y MANEJO DE ESTADO EN STREAMLIT
# ==========================================
if 'df_final' not in st.session_state:
    st.session_state.df_final = None
if 'liq_data' not in st.session_state:
    st.session_state.liq_data = None
if 'req_data' not in st.session_state:
    st.session_state.req_data = None

with st.sidebar:
    st.header("1. Datos del Afiliado")
    cedula = st.text_input("Cédula")
    nombre = st.text_input("Nombre Completo")
    genero = st.radio("Género", ["Masculino", "Femenino"])
    fecha_nac = st.date_input("Fecha de Nacimiento", value=date(1960, 1, 1), min_value=date(1900,1,1), max_value=datetime.now().date())
    
    st.header("2. Historia Laboral")
    hl_file = st.file_uploader("Cargar HL (PDF)", type="pdf")
    
    if st.button("🔄 Reiniciar Trámite"):
        st.session_state.df_final = None
        st.session_state.liq_data = None
        st.rerun()

st.title("⚖️ Sustanciador Colpensiones RPM")

if not MODULOS_CARGADOS:
    st.warning("⚠️ Módulos `data_processor.py` y `logic.py` no encontrados. Ejecutando en modo maqueta (Mockup). Asegúrate de subirlos a GitHub.")

if hl_file and nombre and cedula:
    if st.session_state.df_final is None:
        with st.spinner("Procesando Historia Laboral..."):
            if MODULOS_CARGADOS:
                # Flujo real invocando tus funciones
                df_crudo = extraer_tabla_cruda(hl_file)
                # Asumiendo índices por defecto o puedes añadir selects como en app(1).py
                cols = df_crudo.columns.tolist()
                df_clean = limpiar_y_estandarizar(df_crudo, cols[2], cols[3], cols[4], cols[-1])
                st.session_state.df_final = aplicar_regla_simultaneidad(df_clean)
            else:
                # Simulador si no están los módulos cargados aún
                st.session_state.df_final = pd.DataFrame({'Semanas': [1350]})
    
    if st.session_state.df_final is not None and st.session_state.liq_data is None:
        # Ejecución de la Liquidación Jurídica
        if MODULOS_CARGADOS:
            liq = LiquidadorPension(st.session_state.df_final, genero, fecha_nac)
            fechas_clave = liq.determinar_fechas_clave()
            ibl_10, _ = liq.calcular_ibl_indexado(fechas_clave['fecha_corte'], "ultimos_10")
            ibl_vida, _ = liq.calcular_ibl_indexado(fechas_clave['fecha_corte'], "toda_vida")
            
            ibl_def = max(ibl_10, ibl_vida)
            origen_ibl = "Últimos 10 Años" if ibl_10 >= ibl_vida else "Toda la Vida"
            total_sem = st.session_state.df_final['Semanas'].sum()
            
            # Aquí deberías invocar get_requisitos_estatus (debes importarla o integrarla a logic.py)
            # Simplificamos para el ejemplo estructural:
            edad_req = 62 if genero == "Masculino" else 57
            semanas_req = 1300 # Ajustar dinámicamente según C-197 si es mujer en tu lógica
            
            mesada, tasa, formula = liq.calcular_tasa_reemplazo_797(ibl_def, total_sem, datetime.now().year, True)
            
            st.session_state.req_data = {"edad_exigida": edad_req, "semanas_exigidas": semanas_req, "nota": "Regla general Ley 797"}
            st.session_state.liq_data = {
                "semanas": total_sem, "ibl": ibl_def, "origen_ibl": origen_ibl, 
                "mesada": mesada, "formula_tasa": formula
            }
        else:
            # Datos simulados para testing de la UI
            st.session_state.req_data = {"edad_exigida": 62, "semanas_exigidas": 1300, "nota": ""}
            st.session_state.liq_data = {
                "semanas": 1250, "ibl": 2500000, "origen_ibl": "Últimos 10 Años", "mesada": 1625000,
                "formula_tasa": {"s": 1.9, "tasa_base": 64.5, "semanas_adicionales": 0, "incremento": 0, "tasa_final": 64.5}
            }

    # ==========================================
    # DASHBOARD Y DESCARGA
    # ==========================================
    if st.session_state.liq_data:
        req = st.session_state.req_data
        liq = st.session_state.liq_data
        
        # Calcular si cumple requisitos
        edad_actual = (datetime.now().date() - fecha_nac).days // 365
        cumple_edad = edad_actual >= req['edad_exigida']
        cumple_semanas = liq['semanas'] >= req['semanas_exigidas']
        derecho_consolidado = cumple_edad and cumple_semanas

        st.subheader("📋 Calificación del Derecho")
        col1, col2, col3 = st.columns(3)
        
        estado_edad = "✅ Sí" if cumple_edad else "❌ No"
        estado_sem = "✅ Sí" if cumple_semanas else "❌ No"
        estado_der = "🟢 RECONOCER" if derecho_consolidado else "🔴 NEGAR"
        
        col1.metric(f"Edad ({req['edad_exigida']} req)", f"{edad_actual} años", estado_edad)
        col2.metric(f"Semanas ({req['semanas_exigidas']} req)", f"{liq['semanas']:.2f}", estado_sem)
        col3.metric("Sentido del Acto Administrativo", estado_der)

        st.divider()
        st.markdown("### 📝 Proyección del Acto Administrativo")
        
        datos_afi = {"nombre": nombre, "cedula": cedula, "fecha_nac": fecha_nac}
        docx_buffer = generar_resolucion_word(datos_afi, liq, req, derecho_consolidado)
        
        estado_archivo = "Reconocimiento" if derecho_consolidado else "Negacion"
        
        st.download_button(
            label="📥 Descargar Resolución Pensional (Word)",
            data=docx_buffer,
            file_name=f"Resolucion_{estado_archivo}_{cedula}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
else:
    st.info("Carga el expediente en la barra lateral para iniciar la sustanciación.")
