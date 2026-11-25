import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pdfplumber

try:
    from PIL import Image
    import pytesseract
    OCR_DISPONIBLE = True
except ImportError:
    OCR_DISPONIBLE = False

from data_processor import (
    extraer_tabla_cruda,
    limpiar_y_estandarizar,
    aplicar_regla_simultaneidad,
    extraer_datos_basicos,
    extraer_datos_peticion,
    identificar_ultimo_empleador_publico
)
from logic import LiquidadorPension

st.set_page_config(page_title="Sustanciador Pro - Colpensiones", layout="wide", page_icon="⚖️")

# ==========================================
# MANEJO DE ESTADO EN SESIÓN
# ==========================================
if 'datos_basicos' not in st.session_state:
    st.session_state.datos_basicos = {"cedula": "", "nombre": "", "fecha_nac": date(1975, 1, 1), "genero": "Masculino"}
if 'df_crudo' not in st.session_state:
    st.session_state.df_crudo = None
if 'df_final' not in st.session_state:
    st.session_state.df_final = None
if 'info_empleador' not in st.session_state:
    st.session_state.info_empleador = None
if 'liq_resultados' not in st.session_state:
    st.session_state.liq_resultados = None
if 'retro_resultados' not in st.session_state:
    st.session_state.retro_resultados = None
if 'archivos_cargados' not in st.session_state:
    st.session_state.archivos_cargados = {"resoluciones": 0, "peticion": False}
if 'peticiones_texto' not in st.session_state:
    st.session_state.peticiones_texto = "Reconocimiento de Pensión de Vejez."

# Estados exclusivos del Módulo 5 (Autónomo)
if 'banco_motivas' not in st.session_state:
    st.session_state.banco_motivas = pd.DataFrame(columns=["Categoria", "Palabras_Clave", "Texto_Motiva"])
if 'mod5_texto_capturado' not in st.session_state:
    st.session_state.mod5_texto_capturado = ""
if 'mod5_motivas_encontradas' not in st.session_state:
    st.session_state.mod5_motivas_encontradas = []


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
def generar_resolucion_word(
    datos_afi, liq, req, estatus_cumplido, archivos,
    peticiones_texto, df_historia, motivas_dinamicas,
    info_empleador=None, retro=None
):
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
    doc.add_paragraph(
        f"Que el(la) señor(a) {datos_afi['nombre'].upper()}, identificado(a) con cédula No. {datos_afi['cedula']}, "
        f"{texto_peticion} ante esta Administradora manifestando expresamente las siguientes pretensiones:"
    )
    p_pet = doc.add_paragraph()
    p_pet.add_run(f'"{peticiones_texto}"').italic = True

    if archivos['resoluciones'] > 0:
        doc.add_paragraph(
            f"Que obran en el expediente {archivos['resoluciones']} acto(s) administrativo(s) previo(s) "
            f"proferido(s) por la entidad, los cuales fueron objeto de análisis integral."
        )

    # --- 2. CONSIDERACIONES Y CASO CONCRETO ---
    doc.add_heading('2. CONSIDERACIONES JURÍDICAS Y CASO CONCRETO', level=1)
    doc.add_paragraph(
        "Que el artículo 33 de la Ley 100 de 1993, modificado por el artículo 9 de la Ley 797 de 2003, "
        "establece que para tener derecho a la Pensión de Vejez, el afiliado debe reunir concurrentemente "
        "los requisitos de edad y de semanas cotizadas al Sistema General de Seguridad Social en Pensiones."
    )

    if "C-197/23" in req['nota']:
        doc.add_paragraph(
            f"Que de acuerdo con la Sentencia C-197 de 2023 proferida por la Corte Constitucional, "
            f"para el enfoque de género se aplica una disminución progresiva en las semanas exigidas, "
            f"siendo el requisito normativo de {req['semanas']} semanas para la causación del derecho evaluado."
        )

    # Historia laboral
    doc.add_heading('ANÁLISIS DE LA HISTORIA LABORAL:', level=2)
    doc.add_paragraph(
        f"Que, evaluada la trayectoria laboral del(la) señor(a) {datos_afi['nombre'].upper()}, se evidencia "
        f"la siguiente densidad de aportes al Sistema General de Pensiones. En cumplimiento del artículo 18 "
        f"de la Ley 100 de 1993, se aplicó la regla de simultaneidad, garantizando que para periodos concurrentes "
        f"se acumulen los Ingresos Base de Cotización (IBC) sin exceder 30 días mensuales o 360 anuales:"
    )

    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(['Periodo', 'Desde', 'Hasta', 'IBC Sumado', 'Semanas']):
        hdr_cells[i].text = header_text
        hdr_cells[i].paragraphs[0].runs[0].bold = True

    for _, row in df_historia.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['Periodo'])
        row_cells[1].text = row['Desde'].strftime('%d/%m/%Y') if pd.notnull(row['Desde']) else ""
        row_cells[2].text = row['Hasta'].strftime('%d/%m/%Y') if pd.notnull(row['Hasta']) else ""
        row_cells[3].text = f"${row['IBC']:,.0f}"
        row_cells[4].text = f"{row['Semanas']:.2f}"

    doc.add_paragraph(f"\nTotal de semanas consolidadas y válidas: {liq['semanas']:,.2f} semanas.")

    if info_empleador:
        doc.add_paragraph(
            f"Último aportante registrado en historia laboral: {info_empleador['razon_social']} "
            f"(NIT/ID: {info_empleador['nit']}) — Naturaleza detectada: "
            f"{'ENTIDAD PÚBLICA / SEM >= 90%' if info_empleador['es_publico'] else 'SECTOR PRIVADO / INDEPENDIENTE'}."
        )

    if estatus_cumplido:
        doc.add_paragraph(
            f"Que, descendiendo al caso concreto, el(la) solicitante acreditó el cumplimiento de {req['edad']} años "
            f"de edad y superó el umbral de semanas exigidas ({req['semanas']} semanas), consolidando el estatus pensional."
        )

        doc.add_heading('LIQUIDACIÓN TÉCNICA DE LA PRESTACIÓN:', level=2)
        doc.add_paragraph(
            "De conformidad con el artículo 21 de la Ley 100 de 1993, el IBL se calculó indexando los salarios "
            "con base en el IPC certificado por el DANE, evaluando ambos escenarios legales:"
        )
        doc.add_paragraph(f"• IBL indexado de los últimos 10 años cotizados: ${liq['ibl_10']:,.0f} COP.")
        doc.add_paragraph(f"• IBL indexado de toda la vida laboral: ${liq['ibl_vida']:,.0f} COP.")
        doc.add_paragraph(
            f"Se determina como IBL más favorable el correspondiente a {liq['origen_ibl']}, "
            f"por un valor definitivo de ${liq['ibl']:,.0f} COP (año base {liq['anio_base']})."
        )

        f_data = liq['formula_tasa']
        doc.add_paragraph("\nTASA DE REEMPLAZO (FÓRMULA DECRECIENTE LEY 797 DE 2003):")
        doc.add_paragraph(
            "Conforme al artículo 34 de la Ley 100 de 1993, modificado por el artículo 10 de la Ley 797 de 2003, "
            "la tasa de reemplazo se calcula mediante r = 65.50 - 0.50s, con incrementos del 1.5% por cada 50 semanas "
            "adicionales al mínimo exigido, sujeto al límite legal de 15 puntos porcentuales adicionales y tope máximo del 80%:"
        )
        doc.add_paragraph(
            f"1. SMMLV de referencia ({liq['anio_base']}): ${f_data['smmlv_ref']:,.0f} COP.\n"
            f"2. Proporción del IBL respecto al SMMLV (s): {f_data['s']:.2f} salarios mínimos.\n"
            f"3. Fórmula base [65.50 - (0.50 x {f_data['s']:.2f})]: {f_data['tasa_base']:.2f}%.\n"
            f"4. Semanas adicionales computables: {f_data['semanas_adicionales']:.2f} semanas ({f_data['bloques']} bloque(s) de 50).\n"
            f"5. Incremento legal acumulado (máximo 15 puntos): +{f_data['incremento']:.2f}%.\n"
            f"6. TASA DE REEMPLAZO FINAL DEFINITIVA: {f_data['tasa_final']:.2f}%."
        )
        doc.add_paragraph(
            f"LIQUIDACIÓN INICIAL: Mesada pensional inicial equivalente a ${liq['mesada']:,.0f} COP "
            f"para el año {liq['anio_base']}."
        )

        # --- EFECTIVIDAD, PRESCRIPCIÓN Y RETROACTIVO ---
        if retro:
            doc.add_heading('EFECTIVIDAD, PRESCRIPCIÓN Y LIQUIDACIÓN DE RETROACTIVO:', level=2)
            doc.add_paragraph(retro['nota_efectividad'])

            if retro['aplica_prescripcion']:
                doc.add_paragraph(
                    f"PRESCRIPCIÓN TRIENAL (Arts. 488 CST y 151 CPTSS): Dado que la fecha de efectividad "
                    f"({retro['fecha_efectividad'].strftime('%d/%m/%Y')}) es anterior en más de tres (3) años a la fecha "
                    f"de radicación de la reclamación escrita, las mesadas causadas con anterioridad al "
                    f"{retro['fecha_prescripcion'].strftime('%d/%m/%Y')} se encuentran prescritas. "
                    f"En consecuencia, el pago del retroactivo rige a partir del {retro['fecha_inicio_pago'].strftime('%d/%m/%Y')}."
                )

            if retro['genera_retroactivo']:
                doc.add_paragraph(
                    f"Reajustando anualmente la mesada conforme al artículo 14 de la Ley 100 de 1993, "
                    f"aplicando la mesada adicional de noviembre (Art. 50 Ley 100/1993) "
                    f"{'y mesada 14 (Acto Legislativo 01/2005)' if retro['aplica_mesada_14'] else '(sin lugar a mesada 14 conforme al Acto Legislativo 01/2005)'}, "
                    f"y descontando los aportes imperativos en salud (Art. 143 Ley 100/1993 y Ley 2018/2020), "
                    f"se consolida el siguiente resumen de retroactivo pensional:"
                )
                doc.add_paragraph(
                    f"• Mesada actualizada vigente al ingreso en nómina: ${retro['mesada_actualizada']:,.0f} COP.\n"
                    f"• Total Retroactivo Mesadas Ordinarias: ${retro['bruto_ordinarias']:,.0f} COP.\n"
                    f"• Total Mesadas Adicionales (Sin descuento de salud): ${retro['bruto_adicionales']:,.0f} COP.\n"
                    f"• Total Bruto Retroactivo: ${retro['total_bruto']:,.0f} COP.\n"
                    f"• Menos Descuentos Obligatorios en Salud: -${retro['total_salud']:,.0f} COP.\n"
                    f"• VALOR NETO A GIRAR POR RETROACTIVO: ${retro['total_neto']:,.0f} COP."
                )
            else:
                doc.add_paragraph(
                    f"En virtud de lo expuesto, no se genera pago de retroactivo pensional. "
                    f"La mesada actualizada proyectada para el ingreso en nómina asciende a ${retro['mesada_actualizada']:,.0f} COP."
                )
    else:
        doc.add_paragraph(
            "Que, descendiendo al caso concreto, el(la) solicitante NO CUMPLE con la totalidad de los requisitos "
            "exigidos por la norma para acceder a la Pensión de Vejez."
        )
        if liq['cumple_sem'] and not liq['cumple_edad']:
            doc.add_paragraph(
                f"Si bien la historia laboral evidencia {liq['semanas']:,.2f} semanas cotizadas, "
                f"el(la) asegurado(a) NO ACREDITA el requisito de edad ({req['edad']} años para el género {datos_afi['genero'].lower()})."
            )
        elif liq['cumple_edad'] and not liq['cumple_sem']:
            doc.add_paragraph(
                f"Si bien acredita el requisito de edad ({req['edad']} años), cuenta únicamente con "
                f"{liq['semanas']:,.2f} semanas frente a las {req['semanas']} semanas exigidas por ley."
            )
        else:
            doc.add_paragraph(
                f"El(la) afiliado(a) NO ACREDITA el requisito de edad ({req['edad']} años) ni el requisito "
                f"de semanas cotizadas ({req['semanas']} exigidas), contando únicamente con {liq['semanas']:,.2f} semanas."
            )

    # --- ANÁLISIS DE PETICIONES Y MOTIVAS ADICIONALES ---
    doc.add_heading('ANÁLISIS DE PETICIONES Y MOTIVAS ADICIONALES:', level=2)
    if motivas_dinamicas:
        for motiva in motivas_dinamicas:
            doc.add_paragraph(motiva['Texto'])
    else:
        if estatus_cumplido:
            doc.add_paragraph(
                "Respecto a las demás pretensiones accesorias contenidas en el escrito petitorio, "
                "la Administradora precisa que se despachan conforme a estricto derecho y al ordenamiento vigente."
            )
        else:
            doc.add_paragraph(
                "Al no cumplirse los presupuestos normativos de causación para acceder a la prestación principal, "
                "las pretensiones accesorias corren la misma suerte y son despachadas desfavorablemente."
            )

    # --- 3. RESUELVE ---
    doc.add_heading('RESUELVE:', level=1)
    if estatus_cumplido:
        mesada_resuelve = retro['mesada_actualizada'] if retro else liq['mesada']
        if retro and retro['condicionado_retiro_publico']:
            doc.add_paragraph(
                f"ARTÍCULO PRIMERO: RECONOCER una Pensión de Vejez a favor de {datos_afi['nombre'].upper()}, "
                f"C.C. {datos_afi['cedula']}, en cuantía de ${mesada_resuelve:,.0f} COP mensuales, dejando en suspenso "
                f"su inclusión en nómina y disfrute hasta tanto acredite el acto administrativo de retiro definitivo "
                f"del servicio público (Art. 128 C.P. y Ley 1821 de 2016)."
            )
        else:
            texto_efec = (
                f" con efectividad a partir del {retro['fecha_efectividad'].strftime('%d/%m/%Y')}"
                if retro else ""
            )
            doc.add_paragraph(
                f"ARTÍCULO PRIMERO: RECONOCER Y ORDENAR EL PAGO de una Pensión de Vejez a favor de "
                f"{datos_afi['nombre'].upper()}, C.C. {datos_afi['cedula']}, en cuantía actualizada de "
                f"${mesada_resuelve:,.0f} COP mensuales{texto_efec}."
            )
            if retro and retro['genera_retroactivo']:
                doc.add_paragraph(
                    f"ARTÍCULO SEGUNDO: ORDENAR el pago de un retroactivo pensional neto de ${retro['total_neto']:,.0f} COP "
                    f"(liquidado desde el {retro['fecha_inicio_pago'].strftime('%d/%m/%Y')}), previo descuento legal en salud "
                    f"por valor de ${retro['total_salud']:,.0f} COP con destino a la EPS respectiva."
                )
    else:
        doc.add_paragraph(
            f"ARTÍCULO PRIMERO: NEGAR el reconocimiento de la Pensión de Vejez y despachar desfavorablemente "
            f"las peticiones elevadas por {datos_afi['nombre'].upper()}, C.C. {datos_afi['cedula']}, "
            f"por las razones expuestas en la parte motiva."
        )

    doc.add_paragraph(
        "ARTÍCULO FINAL: RECURSOS. Contra la presente Resolución proceden los recursos de ley "
        "conforme a lo establecido en la Ley 1437 de 2011 (CPACA)."
    )
    doc.add_paragraph("\nNOTIFÍQUESE Y CÚMPLASE\n\n\nFirma Autorizada\nDirección de Prestaciones Económicas\nColpensiones")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def generar_word_solo_motivas(lista_motivas):
    """Genera un archivo Word independiente solo con las motivas resultantes del Módulo 5."""
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
    st.write("Panel de Control")
    if st.button("🔄 Reiniciar Expediente (Módulos 1-4)", use_container_width=True, type="primary"):
        st.session_state.datos_basicos = {"cedula": "", "nombre": "", "fecha_nac": date(1975, 1, 1), "genero": "Masculino"}
        st.session_state.df_crudo = None
        st.session_state.df_final = None
        st.session_state.info_empleador = None
        st.session_state.liq_resultados = None
        st.session_state.retro_resultados = None
        st.session_state.archivos_cargados = {"resoluciones": 0, "peticion": False}
        st.session_state.peticiones_texto = "Reconocimiento de Pensión de Vejez."
        st.rerun()

mod1, mod2, mod3, mod4, mod5 = st.tabs([
    "Módulo 1: Datos Básicos",
    "Módulo 2: Historia Laboral",
    "Módulo 3: Liquidación y Retroactivo",
    "Módulo 4: Sustanciación",
    "Módulo 5: Análisis de Motivas (Autónomo)"
])

# ------------------------------------------
# MÓDULO 1: DATOS BÁSICOS
# ------------------------------------------
with mod1:
    st.header("📂 Carga del Expediente y Datos Biográficos")
    col_docs1, col_docs2 = st.columns(2)
    with col_docs1:
        hl_file = st.file_uploader("1. Historia Laboral (PDF)", type="pdf", key="hl_mod1")
    with col_docs2:
        resoluciones_files = st.file_uploader("2. Resoluciones Previas", type="pdf", accept_multiple_files=True, key="res_mod1")
        peticion_file = st.file_uploader("3. Petición del Ciudadano", type="pdf", key="pet_mod1")

    st.session_state.archivos_cargados['resoluciones'] = len(resoluciones_files) if resoluciones_files else 0
    st.session_state.archivos_cargados['peticion'] = True if peticion_file else False

    if st.button("Procesar Archivos para Extraer Datos Básicos y Peticiones", type="primary"):
        datos_temp = {"cedula": "", "nombre": "", "fecha_nac": date(1975, 1, 1), "genero": "Masculino"}
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
        st.success("Cruce de datos completado.")
        st.rerun()

    st.divider()
    st.subheader("Datos del Asegurado (Extraídos/Editables)")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.datos_basicos["cedula"] = st.text_input("Cédula", value=st.session_state.datos_basicos["cedula"])
        st.session_state.datos_basicos["nombre"] = st.text_input("Nombre Completo", value=st.session_state.datos_basicos["nombre"])
    with c2:
        st.session_state.datos_basicos["fecha_nac"] = st.date_input(
            "Fecha de Nacimiento",
            value=st.session_state.datos_basicos["fecha_nac"],
            min_value=date(1900, 1, 1),
            max_value=date.today()
        )
        idx_gen = 0 if st.session_state.datos_basicos["genero"] == "Masculino" else 1
        st.session_state.datos_basicos["genero"] = st.radio("Género", ["Masculino", "Femenino"], index=idx_gen, horizontal=True)

    st.divider()
    st.subheader("Peticiones a Resolver de Fondo")
    st.session_state.peticiones_texto = st.text_area("Pretensiones capturadas:", value=st.session_state.peticiones_texto, height=120)

# ------------------------------------------
# MÓDULO 2: HISTORIA LABORAL + DETECTOR EMPLEADOR PÚBLICO
# ------------------------------------------
with mod2:
    st.header("📋 Procesamiento de la Historia Laboral")
    st.info(f"Asegurado: {st.session_state.datos_basicos['nombre']} | C.C. {st.session_state.datos_basicos['cedula']}")

    if st.session_state.df_crudo is not None and not st.session_state.df_crudo.empty:
        df = st.session_state.df_crudo
        cols = df.columns.tolist()

        def idx_col(nombre_esperado, fallback):
            return cols.index(nombre_esperado) if nombre_esperado in cols else min(fallback, len(cols) - 1)

        with st.expander("👁️ Ver tabla cruda extraída del PDF (verificación de NIT y Razón Social)", expanded=False):
            st.dataframe(df, use_container_width=True)

        st.write("**Mapeo de columnas detectadas:**")
        r1c1, r1c2 = st.columns(2)
        c_nit = r1c1.selectbox("Columna 'NIT / ID Aportante'", cols, index=idx_col("Identificación Aportante", 0))
        c_ap = r1c2.selectbox("Columna 'Nombre / Razón Social Aportante'", cols, index=idx_col("Nombre o Razón Social", 1))

        c1, c2, c3, c4 = st.columns(4)
        cd = c1.selectbox("Columna 'Desde'", cols, index=idx_col("Desde", 2))
        ch = c2.selectbox("Columna 'Hasta'", cols, index=idx_col("Hasta", 3))
        ci = c3.selectbox("Columna 'IBC'", cols, index=idx_col("Último Salario", 4))
        cs = c4.selectbox("Columna 'Semanas'", cols, index=idx_col("Total", len(cols) - 1))

        if st.button("Ejecutar Saneamiento, Simultaneidad y Validación de Empleador", type="primary"):
            clean = limpiar_y_estandarizar(df, cd, ch, ci, cs, col_nit=c_nit, col_aportante=c_ap)
            if not clean.empty:
                st.session_state.info_empleador = identificar_ultimo_empleador_publico(clean)
                st.session_state.df_final = aplicar_regla_simultaneidad(clean)
                st.success("Historia Laboral procesada exitosamente. Continúe al Módulo 3.")

        if st.session_state.info_empleador:
            emp = st.session_state.info_empleador
            if emp["es_publico"]:
                st.warning(
                    f"🏛️ **ALERTA SECTOR PÚBLICO DETECTADA:** Último aportante **{emp['razon_social']}** "
                    f"(NIT/ID: {emp['nit']}) hasta {emp['ultima_fecha']}.\n\n"
                    f"Motivo: {emp['motivo']}. Se activará la regla de retiro definitivo del servicio (Art. 128 C.P.) en el Módulo 3."
                )
            else:
                st.info(
                    f"🏢 **Último Aportante:** {emp['razon_social']} (NIT/ID: {emp['nit']}) "
                    f"hasta {emp['ultima_fecha']} — Sector Privado / Independiente."
                )

        if st.session_state.df_final is not None:
            st.write("### Base Consolidada de Cotizaciones")
            df_vista = st.session_state.df_final.rename(columns={
                "NIT_Aportante": "Identificación Aportante",
                "Aportante": "Nombre o Razón Social"
            })
            st.dataframe(
                df_vista.style.format({'IBC': "${:,.0f}", 'Semanas': "{:.2f}"}),
                use_container_width=True
            )
    else:
        st.warning("Debe cargar una Historia Laboral válida en el Módulo 1.")

# ------------------------------------------
# MÓDULO 3: LIQUIDACIÓN, EFECTIVIDAD Y RETROACTIVO
# ------------------------------------------
with mod3:
    st.header("🧮 Liquidación de Prestación y Retroactivo")
    st.info(f"Asegurado: {st.session_state.datos_basicos['nombre']}")

    if st.session_state.df_final is not None:
        emp = st.session_state.info_empleador
        opciones_vinc = [
            "Dependiente Privado (Con retiro / Retiro tácito)",
            "Independiente / Subsidiado",
            "Servidor Público",
            "Inactivo antes de cumplir edad",
            "Activo (Sin novedad de retiro)"
        ]
        idx_default_vinc = 2 if (emp and emp["es_publico"]) else 0

        st.subheader("Parámetros de Efectividad y Retroactivo")
        p1, p2, p3 = st.columns(3)
        with p1:
            tipo_vinculacion = st.selectbox("Tipo de Vinculación / Escenario", opciones_vinc, index=idx_default_vinc)
        with p2:
            fecha_radicacion = st.date_input("Fecha de Radicación de la Solicitud", value=date.today())
        with p3:
            # Por defecto primer día del mes siguiente
            hoy = date.today()
            prox_mes_1 = date(hoy.year + (1 if hoy.month == 12 else 0), 1 if hoy.month == 12 else hoy.month + 1, 1)
            fecha_inclusion_nomina = st.date_input("Fecha Estimada Inclusión en Nómina", value=prox_mes_1)

        r1, r2 = st.columns(2)
        with r1:
            tiene_retiro_formal = st.checkbox(
                "Acredita novedad de retiro en PILA o acto administrativo de retiro definitivo",
                value=False
            )
        with r2:
            ultima_fecha_hl = st.session_state.df_final['Hasta'].max().date()
            fecha_retiro_formal = st.date_input(
                "Fecha de Retiro Formal / Acto de Retiro",
                value=ultima_fecha_hl,
                disabled=not tiene_retiro_formal
            )

        if st.button("Calcular Derecho, Liquidación y Retroactivo", type="primary"):
            df_f = st.session_state.df_final
            genero = st.session_state.datos_basicos["genero"]
            fecha_nac = st.session_state.datos_basicos["fecha_nac"]

            liq = LiquidadorPension(df_f, genero, fecha_nac)
            fechas = liq.determinar_fechas_clave()

            anio_base = fechas['fecha_corte'].year
            ibl_10, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "ultimos_10")
            ibl_vida, _ = liq.calcular_ibl_indexado(fechas['fecha_corte'], "toda_vida")
            ibl_def = max(ibl_10, ibl_vida)
            origen_ibl = "Últimos 10 Años" if ibl_10 >= ibl_vida else "Toda la Vida"

            total_sem = df_f['Semanas'].sum()
            edad_req, sem_req, nota_req = get_requisitos_estatus(genero, fechas['fecha_estatus'], fechas['fecha_cumple_edad'])
            mesada, tasa, f_detalle = liq.calcular_tasa_reemplazo_797(ibl_def, total_sem, anio_base, True)

            cumple_edad = pd.to_datetime(fechas['fecha_cumple_edad']).date() <= date.today()
            cumple_sem = total_sem >= sem_req
            reconoce = cumple_edad and cumple_sem

            st.session_state.liq_resultados = {
                "semanas": total_sem,
                "ibl": ibl_def,
                "origen_ibl": origen_ibl,
                "mesada": mesada,
                "anio_base": anio_base,
                "ibl_10": ibl_10,
                "ibl_vida": ibl_vida,
                "formula_tasa": f_detalle,
                "edad_req": edad_req,
                "sem_req": sem_req,
                "nota_req": nota_req,
                "reconoce": reconoce,
                "cumple_edad": cumple_edad,
                "cumple_sem": cumple_sem,
                "fechas": fechas
            }

            if reconoce:
                st.session_state.retro_resultados = liq.calcular_retroactivo(
                    mesada_base=mesada,
                    anio_base=anio_base,
                    fecha_estatus=fechas['fecha_estatus'],
                    fecha_ultima_cot=fechas['ultima_cotizacion'],
                    fecha_radicacion=fecha_radicacion,
                    fecha_inclusion_nomina=fecha_inclusion_nomina,
                    tipo_vinculacion=tipo_vinculacion,
                    tiene_retiro_formal=tiene_retiro_formal,
                    fecha_retiro_formal=fecha_retiro_formal if tiene_retiro_formal else None
                )
            else:
                st.session_state.retro_resultados = None

        if st.session_state.liq_resultados is not None:
            r = st.session_state.liq_resultados
            retro = st.session_state.retro_resultados

            if r['reconoce']:
                st.success("✅ PROCEDE EL RECONOCIMIENTO")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric(f"IBL ({r['anio_base']})", f"${r['ibl']:,.0f}", r['origen_ibl'])
                c2.metric("Tasa de Reemplazo", f"{r['formula_tasa']['tasa_final']:.2f}%")
                c3.metric(f"Mesada Inicial ({r['anio_base']})", f"${r['mesada']:,.0f}")
                if retro:
                    c4.metric("Mesada Actualizada", f"${retro['mesada_actualizada']:,.0f}")

                if retro:
                    st.divider()
                    st.subheader("📅 Estudio de Efectividad y Retroactivo")
                    if retro["condicionado_retiro_publico"]:
                        st.warning(f"🏛️ {retro['nota_efectividad']}")
                    else:
                        st.info(retro['nota_efectividad'])

                    if retro["aplica_prescripcion"]:
                        st.warning(
                            f"⚠️ **Prescripción Trienal Aplicada:** Efectividad del {retro['fecha_efectividad'].strftime('%d/%m/%Y')}, "
                            f"pero pago liquidado a partir del **{retro['fecha_inicio_pago'].strftime('%d/%m/%Y')}** (Art. 488 CST / 151 CPTSS)."
                        )

                    if retro["genera_retroactivo"]:
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("Bruto Ordinarias", f"${retro['bruto_ordinarias']:,.0f}")
                        m2.metric("Mesadas Adicionales", f"${retro['bruto_adicionales']:,.0f}")
                        m3.metric("Descuento Salud", f"-${retro['total_salud']:,.0f}")
                        m4.metric("RETROACTIVO NETO", f"${retro['total_neto']:,.0f}")

                        with st.expander("Ver liquidación detallada mes a mes del retroactivo"):
                            st.dataframe(
                                retro["df_detalle"].style.format({
                                    'Mesada Vigente': "${:,.0f}",
                                    'Valor Ordinario': "${:,.0f}",
                                    'Valor Adicional': "${:,.0f}",
                                    '% Salud': "{:.0f}%",
                                    'Descuento Salud': "${:,.0f}",
                                    'Neto a Pagar': "${:,.0f}"
                                }),
                                use_container_width=True
                            )
                    else:
                        st.info("No se liquida valor de retroactivo para el escenario seleccionado.")
            else:
                st.error("❌ NO ACREDITA REQUISITOS DE LEY")
                st.write(f"- Edad cumplida: **{'Sí' if r['cumple_edad'] else 'No'}** (Requisito: {r['edad_req']} años)")
                st.write(f"- Semanas actuales: **{r['semanas']:,.2f}** / Semanas exigidas: **{r['sem_req']}**")
    else:
        st.warning("Debe procesar la Historia Laboral en el Módulo 2.")

# ------------------------------------------
# MÓDULO 4: RESULTADO VISUAL Y DESCARGA
# ------------------------------------------
with mod4:
    st.header("📄 Sustanciación Final (Resolución Completa)")

    if st.session_state.liq_resultados is not None:
        r = st.session_state.liq_resultados
        retro = st.session_state.retro_resultados
        estado_texto = "RECONOCIMIENTO" if r['reconoce'] else "NEGACIÓN"

        st.subheader("Resumen Ejecutivo")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Cédula", st.session_state.datos_basicos["cedula"])
        c2.metric("Semanas Válidas", f"{r['semanas']:,.2f}")
        c3.metric(
            "Retroactivo Neto",
            f"${retro['total_neto']:,.0f}" if (r['reconoce'] and retro and retro['genera_retroactivo']) else "$0"
        )
        c4.metric("Sentido del Acto", estado_texto, delta_color="normal" if r['reconoce'] else "inverse")

        st.divider()
        req_data = {"edad": r['edad_req'], "semanas": r['sem_req'], "nota": r['nota_req']}
        datos_afi = {
            "nombre": st.session_state.datos_basicos["nombre"],
            "cedula": st.session_state.datos_basicos["cedula"],
            "fecha_nac": st.session_state.datos_basicos["fecha_nac"].strftime('%d/%m/%Y'),
            "genero": st.session_state.datos_basicos["genero"]
        }

        docx_buffer = generar_resolucion_word(
            datos_afi,
            r,
            req_data,
            r['reconoce'],
            st.session_state.archivos_cargados,
            st.session_state.peticiones_texto,
            st.session_state.df_final,
            st.session_state.mod5_motivas_encontradas,
            info_empleador=st.session_state.info_empleador,
            retro=retro
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

# ------------------------------------------
# MÓDULO 5: ANÁLISIS DE MOTIVAS (TOTALMENTE AUTÓNOMO)
# ------------------------------------------
with mod5:
    st.header("🔍 Asistente Autónomo de Motivas Jurídicas")
    st.info(
        "Este módulo opera de forma independiente a la liquidación de fondo. Permite leer solicitudes "
        "(PDF, Word o imagen), buscar motivas, explicar la fórmula decreciente y generar un Word autónomo."
    )

    t_analisis, t_formula, t_banco = st.tabs([
        "📄 1. Análisis de Solicitud",
        "🧮 2. Fórmula Decreciente (Ley 797)",
        "⚙️ 3. Banco de Datos"
    ])

    # PESTAÑA 1: ANÁLISIS Y LECTURA
    with t_analisis:
        st.subheader("Carga y Lectura de la Petición")
        archivo_mod5 = st.file_uploader(
            "Sube el escrito ciudadano (Soporta PDF, DOCX, PNG, JPG)",
            type=['pdf', 'docx', 'png', 'jpg', 'jpeg']
        )

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
                            st.error("Librerías OCR no detectadas. Verifica requirements.txt y packages.txt.")
                except Exception as e:
                    st.error(f"Error al leer el archivo: {e}")

                if texto_extraido:
                    st.session_state.mod5_texto_capturado = texto_extraido.strip()
                    st.success("Texto extraído correctamente. Puedes editarlo en el cuadro inferior.")
            else:
                st.warning("Por favor carga un archivo primero.")

        st.write("### Texto de la Petición Capturada")
        st.caption("Edita, agrega o elimina libremente el contenido antes de buscar las motivas.")
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
                if not encontradas:
                    st.warning("No hubo coincidencias. Intenta ajustar el texto o verificar las palabras clave del banco.")
            else:
                st.error("El banco de motivas está vacío. Cárgalo en la pestaña 'Banco de Datos'.")

        if st.session_state.mod5_motivas_encontradas:
            st.divider()
            st.subheader(f"Motivas Encontradas ({len(st.session_state.mod5_motivas_encontradas)})")

            for i, m in enumerate(st.session_state.mod5_motivas_encontradas):
                nuevo_texto = st.text_area(f"Motiva: {m['Categoria']}", value=m['Texto'], height=120, key=f"edit_mot_{i}")
                st.session_state.mod5_motivas_encontradas[i]['Texto'] = nuevo_texto

            word_motivas_buffer = generar_word_solo_motivas(st.session_state.mod5_motivas_encontradas)
            st.download_button(
                label="📥 EXPORTAR MOTIVAS A WORD",
                data=word_motivas_buffer,
                file_name="Motivas_Autonomas_RPM.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="secondary",
                use_container_width=True
            )

    # PESTAÑA 2: FÓRMULA DECRECIENTE (CON TOPE DE 15 PUNTOS)
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
            sem_extra = max(0.0, val_semanas - val_req)
            bloques_extra = int(sem_extra // 50)

            incremento_bruto = bloques_extra * 1.5
            incremento_legal = min(incremento_bruto, 15.0)
            t_total = min(80.0, t_base + incremento_legal)

            texto_limite = ""
            if incremento_bruto > 15.0:
                texto_limite = " (Valor limitado a 15.00%, en razón a que la fórmula legal solo contempla un aumento máximo de 15 puntos)"

            explicacion = (
                f"Respecto a la aplicación de la tasa de reemplazo, esta Administradora procedió conforme al artículo 34 "
                f"de la Ley 100 de 1993, modificado por el artículo 10 de la Ley 797 de 2003, aplicando la fórmula decreciente "
                f"de la siguiente manera:\n\n"
                f"1. Relación IBL/SMMLV (s): ${val_ibl:,.0f} / ${val_smmlv:,.0f} = {s:.2f} salarios mínimos.\n"
                f"2. Tasa base [65.5 - (0.5 x {s:.2f})]: {t_base:.2f}%.\n"
                f"3. Semanas adicionales cotizadas sobre las {val_req} mínimas de ley: {sem_extra:.2f} semanas, "
                f"lo que equivale a {bloques_extra} bloque(s) completo(s) de 50 semanas.\n"
                f"4. Incremento legal por semanas adicionales ({bloques_extra} bloques x 1.5%): +{incremento_legal:.2f}%{texto_limite}.\n"
                f"5. Tasa de reemplazo final calculada: {t_base + incremento_legal:.2f}%, la cual, en estricto cumplimiento "
                f"del tope normativo que determina que 'el monto máximo de la pensión no podrá ser superior al 80% del ingreso "
                f"base de liquidación', se fija en {t_total:.2f}%.\n\n"
                f"Por tanto, la liquidación de la prestación económica se encuentra ajustada a los mandatos imperativos de la "
                f"Ley 797 de 2003, la cual establece de forma perentoria que la sumatoria de puntos porcentuales por semanas "
                f"adicionales tiene un tope máximo legal de 15 puntos."
            )

            st.session_state.texto_formula_generada = st.text_area("Explicación Matemática de la Fórmula", value=explicacion, height=300)

    # PESTAÑA 3: BANCO DE DATOS
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
