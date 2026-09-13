import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from data_processor import extraer_tabla_cruda, limpiar_y_estandarizar, aplicar_regla_simultaneidad, extraer_datos_basicos
from logic import LiquidadorPension

st.set_page_config(page_title="Sustanciador Pro - Colpensiones", layout="wide", page_icon="⚖️")

# ==========================================
# MANEJO DE ESTADO EN SESIÓN
# ==========================================
if 'datos_basicos' not in st.session_state:
    st.session_state.datos_basicos = {"cedula": "", "nombre": "", "fecha_nac": date(1975, 1, 1), "genero": "Masculino"}
if 'df_crudo' not in st.session_state: st.session_state.df_crudo = None
if 'df_final' not in st.session_state: st.session_state.df_final = None
if 'liq_resultados' not in st.session_state: st.session_state.liq_resultados = None
if 'archivos_cargados' not in st.session_state: st.session_state.archivos_cargados = {"resoluciones": 0, "peticion": False}

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

def generar_resolucion_word(datos_afi, liq, req, estatus_cumplido, archivos):
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
    texto_peticion = "radicó petición formal" if archivos['peticion'] else "elevó solicitud"
    doc.add_paragraph(f"Que el(la) señor(a) {datos_afi['nombre'].upper()}, identificado(a) con cédula No. {datos_afi['cedula']}, {texto_peticion} ante esta Administradora para el estudio de su prestación económica.")
    
    if archivos['resoluciones'] > 0:
        doc.add_paragraph(f"Que obran en el expediente {archivos['resoluciones']} acto(s) administrativo(s) previo(s) proferido(s) por la entidad, los cuales fueron objeto de análisis integral.")
    
    doc.add_paragraph(f"Que procesada la Historia Laboral aportada, se constata que cuenta con {liq['semanas']:,.2f} semanas cotizadas válidas.")

    doc.add_heading('2. CONSIDERACIONES JURÍDICAS Y CASO CONCRETO', level=1)
    doc.add_paragraph("Que el artículo 33 de la Ley 100 de 1993, modificado por el art. 9 de la Ley 797 de 2003, establece los requisitos concurrentes de edad y densidad de semanas.")

    if "C-197/23" in req['nota']:
        doc.add_paragraph(f"Que en aplicación de la Sentencia C-197 de 2023 de la Corte Constitucional, el requisito de semanas exigido para la causación del derecho es de {req['semanas']} semanas.")

    if estatus_cumplido:
        doc.add_paragraph(f"Que el(la) solicitante acreditó el cumplimiento de {req['edad']} años de edad y superó el umbral de semanas exigidas, consolidando el derecho pensional.")
        f_data = liq['formula_tasa']
        
        doc.add_heading('LIQUIDACIÓN TÉCNICA DE LA PRESTACIÓN:', level=2)
        doc.add_paragraph(f"Que en cumplimiento del artículo 21 de la Ley 100 de 1993, se determinó como IBL más favorable el correspondiente a {liq['origen_ibl']}, indexado por valor de ${liq['ibl']:,.0f} COP.")
        
        doc.add_paragraph("Tasa de reemplazo (Art. 34 Ley 100 de 1993):")
        doc.add_paragraph(
            f"1. Proporción IBL respecto al SMMLV (s): {f_data['s']:.2f}\n"
            f"2. Porcentaje base [65.50 - (0.50 * s)]: {f_data['tasa_base']:.2f}%\n"
            f"3. Incremento por semanas adicionales ({f_data['bloques']} bloque(s)): +{f_data['incremento']:.2f}%\n"
            f"4. TASA DE REEMPLAZO FINAL DEFINITIVA: {f_data['tasa_final']:.2f}%"
        )
        doc.add_paragraph(f"En consecuencia, el valor de la mesada pensional asciende a ${liq['mesada']:,.0f} COP mensuales.")
    else:
        doc.add_paragraph(f"Que el(la) solicitante NO CUMPLE con los requisitos exigidos. A la fecha de corte cuenta únicamente con {liq['semanas']:,.2f} semanas, siendo insuficientes frente a las {req['semanas']} semanas exigidas por el ordenamiento.")

    doc.add_heading('RESUELVE:', level=1)
    if estatus_cumplido:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: RECONOCER Y ORDENAR EL PAGO de una Pensión de Vejez a favor de {datos_afi['nombre'].upper()}, C.C. {datos_afi['cedula']}, en cuantía de ${liq['mesada']:,.0f} COP mensuales.")
    else:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: NEGAR el reconocimiento de la Pensión de Vejez a favor de {datos_afi['nombre'].upper()}, C.C. {datos_afi['cedula']}, por las razones expuestas en la parte motiva.")

    doc.add_paragraph("ARTÍCULO FINAL: RECURSOS. Contra la presente Resolución proceden los recursos de ley.")
    doc.add_paragraph("\nNOTIFÍQUESE Y CÚMPLASE\n\n\nFirma Autorizada\nDirección de Prestaciones Económicas\nColpensiones")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# INTERFAZ PRINCIPAL (TABS)
# ==========================================
st.title("⚖️ Agente Sustanciador: RPM Colpensiones")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/4/4e/Colpensiones_logo.png", width=200)
    st.write("Panel de Control")
    if st.button("🔄 Reiniciar Expediente", use_container_width=True, type="primary"):
        st.session_state.datos_basicos = {"cedula": "", "nombre": "", "fecha_nac": date(1975, 1, 1), "genero": "Masculino"}
        st.session_state.df_crudo = None
        st.session_state.df_final = None
        st.session_state.liq_resultados = None
        st.session_state.archivos_cargados = {"resoluciones": 0, "peticion": False}
        st.rerun()

# --- MÓDULOS DEL SUSTANCIADOR ---
mod1, mod2, mod3, mod4 = st.tabs([
    "Módulo 1: Datos Básicos", 
    "Módulo 2: Historia Laboral", 
    "Módulo 3: Liquidación", 
    "Módulo 4: Sustanciación"
])

# ------------------------------------------
# MÓDULO 1: DATOS BÁSICOS
# ------------------------------------------
with mod1:
    st.header("📂 Carga del Expediente y Datos Biográficos")
    
    col_docs1, col_docs2 = st.columns(2)
    with col_docs1:
        hl_file = st.file_uploader("1. Historia Laboral (PDF)", type="pdf")
        if hl_file and st.session_state.df_crudo is None:
            hl_file.seek(0)
            st.session_state.datos_basicos = extraer_datos_basicos(hl_file)
            hl_file.seek(0)
            st.session_state.df_crudo = extraer_tabla_cruda(hl_file)
            st.rerun()
            
    with col_docs2:
        resoluciones_files = st.file_uploader("2. Resoluciones Previas", type="pdf", accept_multiple_files=True)
        peticion_file = st.file_uploader("3. Petición (Opcional)", type="pdf")
        
        # Guardar en sesión estado de archivos
        st.session_state.archivos_cargados['resoluciones'] = len(resoluciones_files) if resoluciones_files else 0
        st.session_state.archivos_cargados['peticion'] = True if peticion_file else False

    st.divider()
    st.subheader("Datos del Asegurado (Extraídos/Editables)")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.datos_basicos["cedula"] = st.text_input("Cédula", value=st.session_state.datos_basicos["cedula"])
        st.session_state.datos_basicos["nombre"] = st.text_input("Nombre Completo", value=st.session_state.datos_basicos["nombre"])
    with c2:
        st.session_state.datos_basicos["fecha_nac"] = st.date_input("Fecha de Nacimiento", value=st.session_state.datos_basicos["fecha_nac"], min_value=date(1900,1,1), max_value=date.today())
        idx_gen = 0 if st.session_state.datos_basicos["genero"] == "Masculino" else 1
        st.session_state.datos_basicos["genero"] = st.radio("Género", ["Masculino", "Femenino"], index=idx_gen, horizontal=True)

# ------------------------------------------
# MÓDULO 2: HISTORIA LABORAL
# ------------------------------------------
with mod2:
    st.header("📋 Procesamiento de la Historia Laboral")
    st.info(f"Asegurado: {st.session_state.datos_basicos['nombre']} | C.C. {st.session_state.datos_basicos['cedula']}")
    
    if st.session_state.df_crudo is not None and not st.session_state.df_crudo.empty:
        df = st.session_state.df_crudo
        st.write("Configure el mapeo de columnas antes de sanear:")
        cols = df.columns.tolist()
        c1, c2, c3, c4 = st.columns(4)
        cd = c1.selectbox("Columna 'Desde'", cols, index=2 if len(cols)>2 else 0)
        ch = c2.selectbox("Columna 'Hasta'", cols, index=3 if len(cols)>3 else 0)
        ci = c3.selectbox("Columna 'IBC'", cols, index=4 if len(cols)>4 else 0)
        cs = c4.selectbox("Columna 'Semanas'", cols, index=len(cols)-1)
        
        if st.button("Ejecutar Saneamiento y Regla de Simultaneidad"):
            clean = limpiar_y_estandarizar(df, cd, ch, ci, cs)
            if not clean.empty:
                st.session_state.df_final = aplicar_regla_simultaneidad(clean)
                st.success("Historia Laboral procesada exitosamente. Continúe al Módulo 3.")
            else:
                st.error("Error validando columnas.")
                
        if st.session_state.df_final is not None:
            st.write("### Base Consolidada de Cotizaciones")
            st.dataframe(st.session_state.df_final.style.format({'IBC': "${:,.0f}", 'Semanas': "{:.2f}"}), use_container_width=True)
    else:
        st.warning("Debe cargar una Historia Laboral válida en el Módulo 1.")

# ------------------------------------------
# MÓDULO 3: LIQUIDACIÓN
# ------------------------------------------
with mod3:
    st.header("🧮 Liquidación y Derechos")
    st.info(f"Asegurado: {st.session_state.datos_basicos['nombre']} | Fecha de Nacimiento: {st.session_state.datos_basicos['fecha_nac'].strftime('%d/%m/%Y')}")
    
    if st.session_state.df_final is not None:
        if st.button("Calcular Derecho y Liquidar Prestación", type="primary"):
            df_f = st.session_state.df_final
            genero = st.session_state.datos_basicos["genero"]
            fecha_nac = st.session_state.datos_basicos["fecha_nac"]
            
            liq = LiquidadorPension(df_f, genero, fecha_nac)
            fechas = liq.determinar_fechas_clave()
            
            ibl_10, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "ultimos_10")
            ibl_vida, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "toda_vida")
            ibl_def = max(ibl_10, ibl_vida)
            origen_ibl = "Últimos 10 Años" if ibl_10 >= ibl_vida else "Toda la Vida"
            
            total_sem = df_f['Semanas'].sum()
            edad_req, sem_req, nota_req = get_requisitos_estatus(genero, fechas['fecha_estatus'], fechas['fecha_cumple_edad'])
            mesada, tasa, info = liq.calcular_tasa_reemplazo_797(ibl_def, total_sem, datetime.now().year, True)
            
            smlmv_ref = 1750905.0
            s = ibl_def / smlmv_ref if smlmv_ref > 0 else 1
            tasa_base = max(55.0, min(65.5 - (0.5 * s), 65.5))
            sem_adicionales = max(0, total_sem - sem_req)
            bloques = int(sem_adicionales // 50)
            incremento = min(bloques * 1.5, 15.0)

            # Corrección del Timestamp issue aplicada aquí
            cumple_edad = pd.to_datetime(fechas['fecha_cumple_edad']).date() <= date.today()
            cumple_sem = total_sem >= sem_req
            reconoce = cumple_edad and cumple_sem

            st.session_state.liq_resultados = {
                "semanas": total_sem, "ibl": ibl_def, "origen_ibl": origen_ibl, "mesada": mesada,
                "formula_tasa": {"s": s, "tasa_base": tasa_base, "semanas_adicionales": sem_adicionales, "bloques": bloques, "incremento": incremento, "tasa_final": tasa},
                "edad_req": edad_req, "sem_req": sem_req, "nota_req": nota_req, "reconoce": reconoce,
                "cumple_edad": cumple_edad
            }
            
        if st.session_state.liq_resultados is not None:
            r = st.session_state.liq_resultados
            if r['reconoce']:
                st.success("✅ PROCEDE EL RECONOCIMIENTO")
                c1, c2 = st.columns(2)
                c1.metric("IBL Aplicado", f"${r['ibl']:,.0f}", r['origen_ibl'])
                c2.metric("Tasa de Reemplazo", f"{r['formula_tasa']['tasa_final']:.2f}%")
                st.metric("MESADA PROYECTADA", f"${r['mesada']:,.0f}")
            else:
                st.error("❌ NO ACREDITA REQUISITOS DE LEY")
                st.write(f"- Semanas actuales: **{r['semanas']:,.2f}** / Semanas exigidas: **{r['sem_req']}**")
                st.write(f"- Edad cumplida: **{'Sí' if r['cumple_edad'] else 'No'}**")
    else:
        st.warning("Debe procesar la Historia Laboral en el Módulo 2.")

# ------------------------------------------
# MÓDULO 4: RESULTADO VISUAL Y DESCARGA
# ------------------------------------------
with mod4:
    st.header("📄 Sustanciación Final")
    
    if st.session_state.liq_resultados is not None:
        r = st.session_state.liq_resultados
        estado_texto = "RECONOCIMIENTO" if r['reconoce'] else "NEGACIÓN"
        
        st.subheader("Resumen Ejecutivo")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Cédula", st.session_state.datos_basicos["cedula"])
        c2.metric("Semanas Válidas", f"{r['semanas']:,.2f}")
        c3.metric("IBL", f"${r['ibl']:,.0f}" if r['reconoce'] else "N/A")
        c4.metric("Sentido del Acto", estado_texto, delta_color="normal" if r['reconoce'] else "inverse")

        st.divider()
        st.write("Generando acto administrativo con las consideraciones jurídicas correspondientes...")
        
        req_data = {"edad": r['edad_req'], "semanas": r['sem_req'], "nota": r['nota_req']}
        docx_buffer = generar_resolucion_word(
            st.session_state.datos_basicos, 
            r, 
            req_data, 
            r['reconoce'], 
            st.session_state.archivos_cargados
        )
        
        st.download_button(
            label=f"📥 DESCARGAR RESOLUCIÓN ({estado_texto})",
            data=docx_buffer,
            file_name=f"Resolucion_{estado_texto}_{st.session_state.datos_basicos['cedula']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    else:
        st.warning("Complete la Liquidación en el Módulo 3 para generar el acto administrativo.")
