import streamlit as st
import pandas as pd
import re
from datetime import datetime, date
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pdfplumber

# Soporte OCR para lectura de imágenes en el Módulo 5
try:
    from PIL import Image
    import pytesseract
    OCR_DISPONIBLE = True
except ImportError:
    OCR_DISPONIBLE = False

from data_processor import extraer_tabla_cruda, limpiar_y_estandarizar, aplicar_regla_simultaneidad, extraer_datos_basicos, extraer_datos_peticion
from logic import LiquidadorPension

st.set_page_config(page_title="Sustanciador Pro - Colpensiones", layout="wide", page_icon="⚖️")

# ==========================================
# MANEJO DE ESTADO EN SESIÓN
# ==========================================
if 'datos_basicos' not in st.session_state: st.session_state.datos_basicos = {"cedula": "", "nombre": "", "fecha_nac": date(1975, 1, 1), "genero": "Masculino"}
if 'df_crudo' not in st.session_state: st.session_state.df_crudo = None
if 'df_final' not in st.session_state: st.session_state.df_final = None
if 'liq_resultados' not in st.session_state: st.session_state.liq_resultados = None
if 'archivos_cargados' not in st.session_state: st.session_state.archivos_cargados = {"resoluciones": 0, "peticion": False}
if 'peticiones_texto' not in st.session_state: st.session_state.peticiones_texto = "Reconocimiento de Pensión de Vejez."

# Estados exclusivos del Módulo 5 (Autónomo)
if 'banco_motivas' not in st.session_state: st.session_state.banco_motivas = pd.DataFrame(columns=["Categoria", "Palabras_Clave", "Texto_Motiva"])
if 'mod5_texto_capturado' not in st.session_state: st.session_state.mod5_texto_capturado = ""
if 'mod5_motivas_encontradas' not in st.session_state: st.session_state.mod5_motivas_encontradas = []

# --- REQUISITOS EN TIEMPO REAL ---
def get_requisitos_estatus(genero, fecha_estatus, fecha_cumple_edad=None):
    anio_actual = datetime.now().year
    edad_req = 62 if genero == "Masculino" else 57
    
    if genero == "Masculino":
        semanas_req = 1300
        nota = "Aplica regla general Ley 797/2003 (1300 semanas)."
    else:
        anio_eval = anio_actual if pd.isna(fecha_estatus) else fecha_estatus.year
        if anio_eval < 2026:
            semanas_req = 1300
            nota = f"Exigencia de 1300 semanas (causación en {anio_eval})."
        elif anio_eval == 2026:
            semanas_req = 1250
            nota = f"Aplica disminución Sentencia C-197/23 para el año {anio_eval}: 1250 semanas."
        else:
            descenso = 50 + ((anio_eval - 2026) * 25)
            semanas_req = max(1000, 1300 - descenso)
            nota = f"Aplica disminución progresiva (Sentencia C-197/23) para el año {anio_eval}: {semanas_req} semanas."
            
    return edad_req, semanas_req, nota

# ==========================================
# GENERADORES DE DOCUMENTOS WORD
# ==========================================
def generar_resolucion_word(datos_afi, liq, req, estatus_cumplido, archivos, peticiones_texto, df_historia, motivas_dinamicas):
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

    # --- 1. ANTECEDENTES ---
    doc.add_heading('1. ANTECEDENTES (HECHOS)', level=1)
    texto_peticion = "radicó petición formal" if archivos['peticion'] else "elevó solicitud"
    
    doc.add_paragraph(f"Que el(la) señor(a) {datos_afi['nombre'].upper()}, identificado(a) con cédula No. {datos_afi['cedula']}, {texto_peticion} ante esta Administradora manifestando expresamente las siguientes pretensiones:")
    
    p_pet = doc.add_paragraph()
    p_pet.add_run(f'"{peticiones_texto}"').italic = True
    
    if archivos['resoluciones'] > 0:
        doc.add_paragraph(f"Que obran en el expediente {archivos['resoluciones']} acto(s) administrativo(s) previo(s) proferido(s) por la entidad, los cuales fueron objeto de análisis integral.")
    
    # --- 2. CONSIDERACIONES Y CASO CONCRETO ---
    doc.add_heading('2. CONSIDERACIONES JURÍDICAS Y CASO CONCRETO', level=1)
    doc.add_paragraph("Que el artículo 33 de la Ley 100 de 1993, modificado por el artículo 9 de la Ley 797 de 2003, establece que para tener derecho a la Pensión de Vejez, el afiliado debe reunir concurrentemente los requisitos de edad y de semanas cotizadas al Sistema General de Seguridad Social en Pensiones.")

    if "C-197/23" in req['nota']:
        doc.add_paragraph(f"Que de acuerdo con la Sentencia C-197 de 2023 proferida por la Corte Constitucional, para el enfoque de género se aplica una disminución progresiva en las semanas exigidas, siendo el requisito normativo de {req['semanas']} semanas para la causación del derecho evaluado a la fecha.")

    doc.add_heading('ANÁLISIS DE LA HISTORIA LABORAL:', level=2)
    doc.add_paragraph(f"Que, evaluada la trayectoria laboral del(la) señor(a) {datos_afi['nombre'].upper()}, se evidencia la siguiente densidad de aportes al Sistema General de Pensiones. En cumplimiento del artículo 18 de la Ley 100 de 1993, se aplicó la regla de simultaneidad, garantizando que para periodos concurrentes de cotización se acumulen los Ingresos Base de Cotización (IBC) sin que los días superen los 30 días mensuales o 360 anuales:")

    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(['Periodo', 'Desde', 'Hasta', 'IBC Sumado', 'Semanas']):
        hdr_cells[i].text = header_text
        hdr_cells[i].paragraphs[0].runs[0].bold = True

    for index, row in df_historia.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['Periodo'])
        row_cells[1].text = row['Desde'].strftime('%d/%m/%Y') if pd.notnull(row['Desde']) else ""
        row_cells[2].text = row['Hasta'].strftime('%d/%m/%Y') if pd.notnull(row['Hasta']) else ""
        row_cells[3].text = f"${row['IBC']:,.0f}"
        row_cells[4].text = f"{row['Semanas']:.2f}"
        
    doc.add_paragraph(f"\nTotal de semanas consolidadas y válidas: {liq['semanas']:,.2f} semanas.")

    if estatus_cumplido:
        doc.add_paragraph(f"Que, descendiendo al caso concreto, el(la) solicitante acreditó el cumplimiento de {req['edad']} años de edad y superó el umbral de semanas exigidas, consolidando el derecho pensional de manera satisfactoria.")
        doc.add_heading('LIQUIDACIÓN TÉCNICA DE LA PRESTACIÓN:', level=2)
        doc.add_paragraph(f"Se determina como IBL más favorable el correspondiente a {liq['origen_ibl']}, por un valor definitivo de ${liq['ibl']:,.0f} COP.")
        f_data = liq['formula_tasa']
        doc.add_paragraph(
            f"TASA DE REEMPLAZO:\n"
            f"1. Proporción del IBL respecto al SMMLV (s): {f_data['s']:.2f} salarios mínimos.\n"
            f"2. Aplicación de la fórmula base: {f_data['tasa_base']:.2f}%\n"
            f"3. Semanas adicionales cotizadas: {f_data['semanas_adicionales']:.2f} semanas ({f_data['bloques']} bloques de 50).\n"
            f"4. Incremento legal acumulado: +{f_data['incremento']:.2f}%\n"
            f"5. TASA DE REEMPLAZO FINAL: {f_data['tasa_final']:.2f}%\n"
            f"El valor de la mesada pensional asciende a ${liq['mesada']:,.0f} COP mensuales."
        )
    else:
        doc.add_paragraph("Que, descendiendo al caso concreto, el(la) solicitante NO CUMPLE con la totalidad de los requisitos exigidos por la norma para acceder a la Pensión de Vejez.")

    doc.add_heading('ANÁLISIS DE PETICIONES Y MOTIVAS ADICIONALES:', level=2)
    if motivas_dinamicas:
        for motiva in motivas_dinamicas: doc.add_paragraph(motiva['Texto'])
    else:
        doc.add_paragraph("Respecto a las demás pretensiones accesorias contenidas en el escrito petitorio, se despachan conforme al ordenamiento legal vigente.")

    doc.add_heading('RESUELVE:', level=1)
    if estatus_cumplido:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: RECONOCER Y ORDENAR EL PAGO de una Pensión de Vejez a favor de {datos_afi['nombre'].upper()}, C.C. {datos_afi['cedula']}, en cuantía de ${liq['mesada']:,.0f} COP mensuales.")
    else:
        doc.add_paragraph(f"ARTÍCULO PRIMERO: NEGAR el reconocimiento de la Pensión de Vejez a {datos_afi['nombre'].upper()}, C.C. {datos_afi['cedula']}.")

    doc.add_paragraph("\nNOTIFÍQUESE Y CÚMPLASE\n\nFirma Autorizada\nColpensiones")
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generar_word_solo_motivas(lista_motivas):
    """Genera un archivo Word independiente solo con las motivas resultantes del Módulo 5"""
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    
    doc.add_heading('ANÁLISIS DE PETICIONES Y MOTIVAS JURÍDICAS', level=1)
    doc.add_paragraph("A continuación, se exponen los argumentos jurídicos aplicables a las pretensiones analizadas:\n")
    
    for m in lista_motivas:
        doc.add_heading(m['Categoria'], level=2)
        doc.add_paragraph(m['Texto'])
        
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
    if st.button("🔄 Reiniciar Expediente (Módulos 1-4)", use_container_width=True, type="primary"):
        st.session_state.datos_basicos = {"cedula": "", "nombre": "", "fecha_nac": date(1975, 1, 1), "genero": "Masculino"}
        st.session_state.df_crudo = None
        st.session_state.df_final = None
        st.session_state.liq_resultados = None
        st.session_state.archivos_cargados = {"resoluciones": 0, "peticion": False}
        st.session_state.peticiones_texto = "Reconocimiento de Pensión de Vejez."
        st.rerun()

mod1, mod2, mod3, mod4, mod5 = st.tabs([
    "Módulo 1: Datos Básicos", 
    "Módulo 2: Historia Laboral", 
    "Módulo 3: Liquidación", 
    "Módulo 4: Sustanciación",
    "Módulo 5: Análisis de Motivas (Autónomo)"
])

# ------------------------------------------
# MÓDULOS 1, 2, 3 y 4 (Mantienen su lógica original)
# ------------------------------------------
with mod1:
    st.header("📂 Carga del Expediente y Datos Biográficos")
    col_docs1, col_docs2 = st.columns(2)
    with col_docs1: hl_file = st.file_uploader("1. Historia Laboral (PDF)", type="pdf", key="hl_mod1")
    with col_docs2:
        resoluciones_files = st.file_uploader("2. Resoluciones Previas", type="pdf", accept_multiple_files=True, key="res_mod1")
        peticion_file = st.file_uploader("3. Petición del Ciudadano (Solo extracción básica de datos)", type="pdf", key="pet_mod1")
        
    st.session_state.archivos_cargados['resoluciones'] = len(resoluciones_files) if resoluciones_files else 0
    st.session_state.archivos_cargados['peticion'] = True if peticion_file else False

    if st.button("Procesar Archivos de Expediente", type="primary"):
        datos_temp = st.session_state.datos_basicos.copy()
        if hl_file:
            hl_file.seek(0)
            datos_temp = extraer_datos_basicos(hl_file)
            hl_file.seek(0)
            st.session_state.df_crudo = extraer_tabla_cruda(hl_file)
        if peticion_file:
            peticion_file.seek(0)
            datos_temp, pet_ext = extraer_datos_peticion(peticion_file, datos_temp)
            if pet_ext != "Reconocimiento de Pensión de Vejez.":
                st.session_state.peticiones_texto = pet_ext
        st.session_state.datos_basicos = datos_temp
        st.success("Expediente procesado con éxito.")
        st.rerun()

    st.subheader("Datos del Asegurado")
    c1, c2 = st.columns(2)
    st.session_state.datos_basicos["cedula"] = c1.text_input("Cédula", value=st.session_state.datos_basicos["cedula"])
    st.session_state.datos_basicos["nombre"] = c1.text_input("Nombre Completo", value=st.session_state.datos_basicos["nombre"])
    st.session_state.datos_basicos["fecha_nac"] = c2.date_input("Fecha Nacimiento", value=st.session_state.datos_basicos["fecha_nac"])
    idx_gen = 0 if st.session_state.datos_basicos["genero"] == "Masculino" else 1
    st.session_state.datos_basicos["genero"] = c2.radio("Género", ["Masculino", "Femenino"], index=idx_gen, horizontal=True)

with mod2:
    st.header("📋 Procesamiento de la Historia Laboral")
    if st.session_state.df_crudo is not None and not st.session_state.df_crudo.empty:
        df = st.session_state.df_crudo
        cols = df.columns.tolist()
        c1, c2, c3, c4 = st.columns(4)
        cd = c1.selectbox("Columna 'Desde'", cols, index=2 if len(cols)>2 else 0)
        ch = c2.selectbox("Columna 'Hasta'", cols, index=3 if len(cols)>3 else 0)
        ci = c3.selectbox("Columna 'IBC'", cols, index=4 if len(cols)>4 else 0)
        cs = c4.selectbox("Columna 'Semanas'", cols, index=len(cols)-1)
        if st.button("Ejecutar Saneamiento"):
            clean = limpiar_y_estandarizar(df, cd, ch, ci, cs)
            if not clean.empty:
                st.session_state.df_final = aplicar_regla_simultaneidad(clean)
                st.success("Historia Laboral procesada.")
        if st.session_state.df_final is not None:
            st.dataframe(st.session_state.df_final.style.format({'IBC': "${:,.0f}", 'Semanas': "{:.2f}"}), use_container_width=True)
    else: st.warning("Cargue Historia Laboral en el Módulo 1.")

with mod3:
    st.header("🧮 Liquidación y Derechos")
    if st.session_state.df_final is not None:
        if st.button("Calcular Derecho", type="primary"):
            liq = LiquidadorPension(st.session_state.df_final, st.session_state.datos_basicos["genero"], st.session_state.datos_basicos["fecha_nac"])
            fechas = liq.determinar_fechas_clave()
            ibl_10, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "ultimos_10")
            ibl_vida, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "toda_vida")
            ibl_def = max(ibl_10, ibl_vida)
            origen_ibl = "Últimos 10 Años" if ibl_10 >= ibl_vida else "Toda la Vida"
            total_sem = st.session_state.df_final['Semanas'].sum()
            edad_req, sem_req, nota_req = get_requisitos_estatus(st.session_state.datos_basicos["genero"], fechas['fecha_estatus'], fechas['fecha_cumple_edad'])
            mesada, tasa, _ = liq.calcular_tasa_reemplazo_797(ibl_def, total_sem, datetime.now().year, True)
            
            s = ibl_def / 1750905.0 if ibl_def > 0 else 1
            tasa_base = max(55.0, min(65.5 - (0.5 * s), 65.5))
            sem_adicionales = max(0, total_sem - sem_req)
            bloques = int(sem_adicionales // 50)
            
            cumple_edad = pd.to_datetime(fechas['fecha_cumple_edad']).date() <= date.today()
            cumple_sem = total_sem >= sem_req
            st.session_state.liq_resultados = {
                "semanas": total_sem, "ibl": ibl_def, "origen_ibl": origen_ibl, "mesada": mesada,
                "formula_tasa": {"s": s, "tasa_base": tasa_base, "semanas_adicionales": sem_adicionales, "bloques": bloques, "incremento": min(bloques*1.5, 15.0), "tasa_final": tasa},
                "edad_req": edad_req, "sem_req": sem_req, "nota_req": nota_req, "reconoce": cumple_edad and cumple_sem,
                "cumple_edad": cumple_edad, "cumple_sem": cumple_sem
            }
        if st.session_state.liq_resultados is not None:
            r = st.session_state.liq_resultados
            st.success("✅ PROCEDE" if r['reconoce'] else "❌ NO PROCEDE")
            st.write(f"IBL: ${r['ibl']:,.0f} | Tasa: {r['formula_tasa']['tasa_final']:.2f}% | Semanas: {r['semanas']:,.2f}")
    else: st.warning("Complete Módulo 2.")

with mod4:
    st.header("📄 Sustanciación Final (Resolución Completa)")
    if st.session_state.liq_resultados is not None:
        r = st.session_state.liq_resultados
        docx_buffer = generar_resolucion_word(
            st.session_state.datos_basicos, r, {"edad": r['edad_req'], "semanas": r['sem_req'], "nota": r['nota_req']},
            r['reconoce'], st.session_state.archivos_cargados, st.session_state.peticiones_texto,
            st.session_state.df_final, st.session_state.mod5_motivas_encontradas
        )
        st.download_button("📥 DESCARGAR RESOLUCIÓN", data=docx_buffer, file_name="Resolucion_Completa.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
    else: st.warning("Complete Liquidación.")

# ------------------------------------------
# MÓDULO 5: ANÁLISIS DE MOTIVAS (TOTALMENTE AUTÓNOMO)
# ------------------------------------------
with mod5:
    st.header("🔍 Asistente Autónomo de Motivas Jurídicas")
    st.info("Este módulo es independiente de la liquidación de fondo. Permite leer solicitudes, buscar motivas, explicar fórmulas y generar un Word con los argumentos.")
    
    t_analisis, t_formula, t_banco = st.tabs(["📄 1. Análisis de Solicitud", "🧮 2. Fórmula Decreciente (Ley 797)", "⚙️ 3. Banco de Datos"])
    
    # --------------------------------
    # PESTAÑA 1: ANÁLISIS Y LECTURA
    # --------------------------------
    with t_analisis:
        st.subheader("Carga y Lectura de la Petición")
        archivo_mod5 = st.file_uploader("Sube el escrito ciudadano (Soporta PDF, DOCX, PNG, JPG)", type=['pdf', 'docx', 'png', 'jpg', 'jpeg'])
        
        if st.button("Leer Documento Cargado"):
            texto_extraido = ""
            if archivo_mod5:
                nombre = archivo_mod5.name.lower()
                try:
                    if nombre.endswith('pdf'):
                        with pdfplumber.open(archivo_mod5) as pdf:
                            texto_extraido = "\n".join([page.extract_text() or "" for page in pdf.pages])
                    elif nombre.endswith('docx'):
                        doc_mod5 = Document(archivo_mod5)
                        texto_extraido = "\n".join([p.text for p in doc_mod5.paragraphs])
                    elif nombre.endswith(('png', 'jpg', 'jpeg')):
                        if OCR_DISPONIBLE:
                            img = Image.open(archivo_mod5)
                            texto_extraido = pytesseract.image_to_string(img, lang='spa')
                        else:
                            st.error("Librerías OCR no detectadas. Instala pytesseract para leer imágenes.")
                except Exception as e:
                    st.error(f"Error al leer el archivo: {e}")
                
                if texto_extraido:
                    st.session_state.mod5_texto_capturado = texto_extraido.strip()
                    st.success("Texto extraído correctamente. Puedes editarlo en la caja inferior.")
            else:
                st.warning("Por favor carga un archivo primero.")

        st.write("### Texto de la Petición Capturada")
        st.caption("Edita, agrega o elimina libremente el contenido antes de buscar las motivas.")
        
        # Área de texto interactiva vinculada a Session State
        st.session_state.mod5_texto_capturado = st.text_area(
            "Cuerpo de la Petición", 
            value=st.session_state.mod5_texto_capturado, 
            height=200
        )
        
        if st.button("Buscar Motivas Aplicables al Caso", type="primary"):
            peticion_lower = st.session_state.mod5_texto_capturado.lower()
            encontradas = []
            
            if not st.session_state.banco_motivas.empty:
                for _, fila in st.session_state.banco_motivas.iterrows():
                    claves = [p.strip().lower() for p in str(fila['Palabras_Clave']).split(',')]
                    if any(clave in peticion_lower for clave in claves if clave):
                        encontradas.append({"Categoria": fila['Categoria'], "Texto": fila['Texto_Motiva']})
                        
                st.session_state.mod5_motivas_encontradas = encontradas
                if not encontradas: st.warning("No hubo coincidencias. Intenta agregar palabras clave al texto.")
            else:
                st.error("El banco de motivas está vacío. Cárgalo en la pestaña 'Banco de Datos'.")

        # Edición de las motivas encontradas y generación de Word independiente
        if st.session_state.mod5_motivas_encontradas:
            st.divider()
            st.subheader(f"Motivas Encontradas ({len(st.session_state.mod5_motivas_encontradas)})")
            
            # Permitir edición de cada motiva
            for i, m in enumerate(st.session_state.mod5_motivas_encontradas):
                nuevo_texto = st.text_area(f"Motiva: {m['Categoria']}", value=m['Texto'], height=120, key=f"edit_mot_{i}")
                st.session_state.mod5_motivas_encontradas[i]['Texto'] = nuevo_texto
                
            # Botón para descargar el Word solo con estas motivas
            word_motivas_buffer = generar_word_solo_motivas(st.session_state.mod5_motivas_encontradas)
            st.download_button(
                label="📥 EXPORTAR MOTIVAS A WORD",
                data=word_motivas_buffer,
                file_name="Motivas_Autonomas_RPM.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="secondary",
                use_container_width=True
            )

    # --------------------------------
    # PESTAÑA 2: FÓRMULA DECRECIENTE
    # --------------------------------
    with t_formula:
        st.subheader("Generador de Motiva: Fórmula Decreciente (Ley 797/2003)")
        st.write("Ingrese los parámetros para generar la explicación jurídica y matemática al ciudadano.")
        
        c_form1, c_form2 = st.columns(2)
        with c_form1:
            val_ibl = st.number_input("Ingreso Base de Liquidación (IBL)", min_value=0.0, value=2500000.0, step=100000.0)
            val_semanas = st.number_input("Total Semanas Cotizadas", min_value=0.0, value=1850.0, step=10.0)
        with c_form2:
            val_smmlv = st.number_input("SMMLV al momento del reconocimiento", min_value=1.0, value=1750905.0, step=50000.0)
            val_req = st.number_input("Semanas Mínimas Exigidas (Ej. 1300 o 1250)", min_value=1000, value=1300, step=10)

        if st.button("Calcular y Generar Motiva", type="primary"):
            s = val_ibl / val_smmlv
            t_base = max(55.0, min(65.5 - (0.5 * s), 65.5))
            sem_extra = max(0, val_semanas - val_req)
            bloques_extra = int(sem_extra // 50)
            incremento = bloques_extra * 1.5
            t_total = min(80.0, t_base + incremento) # Tope legal
            
            explicacion = (
                f"Respecto a la aplicación de la tasa de reemplazo, esta Administradora procedió conforme al artículo 34 de la Ley 100 de 1993, modificado por el artículo 10 de la Ley 797 de 2003, aplicando la fórmula decreciente de la siguiente manera:\n\n"
                f"1. Relación IBL/SMMLV (s): ${val_ibl:,.0f} / ${val_smmlv:,.0f} = {s:.2f} salarios mínimos.\n"
                f"2. Tasa base [65.5 - (0.5 x {s:.2f})]: {t_base:.2f}%.\n"
                f"3. Semanas adicionales cotizadas sobre las {val_req} mínimas de ley: {sem_extra:.2f} semanas, lo que equivale a {bloques_extra} bloque(s) completo(s) de 50 semanas.\n"
                f"4. Incremento legal por semanas adicionales ({bloques_extra} bloques x 1.5%): +{incremento:.2f}%.\n"
                f"5. Tasa de reemplazo final calculada: {t_base + incremento:.2f}%, la cual, en cumplimiento estricto del inciso normativo que establece que 'el monto máximo de la pensión no podrá ser superior al 80% del ingreso base de liquidación', se ajusta al tope de {t_total:.2f}%.\n\n"
                f"Por tanto, la liquidación de la prestación económica se encuentra ajustada a los mandatos imperativos de la Ley 797 de 2003."
            )
            
            # Se muestra en un text_area para que el usuario pueda copiarla o editarla
            st.session_state.texto_formula_generada = st.text_area("Explicación Matemática de la Fórmula", value=explicacion, height=280)

    # --------------------------------
    # PESTAÑA 3: BANCO DE DATOS
    # --------------------------------
    with t_banco:
        st.subheader("Cargar o Actualizar Banco de Motivas")
        archivo_banco = st.file_uploader("Sube el archivo CSV del banco", type=["csv"], key="csv_mod5")
        if archivo_banco:
            try:
                df_banco = pd.read_csv(archivo_banco, sep=';')
                st.session_state.banco_motivas = df_banco
                st.success(f"Banco actualizado: {len(df_banco)} motivas en memoria.")
            except Exception as e:
                st.error(f"Error leyendo el archivo CSV: {e}")
                
        st.dataframe(st.session_state.banco_motivas, use_container_width=True, height=250)
