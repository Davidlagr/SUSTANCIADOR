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
    identificar_ultimo_empleador_publico,
    extraer_totales_oficiales_pdf
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
if 'totales_pdf' not in st.session_state:
    st.session_state.totales_pdf = None
if 'liq_resultados' not in st.session_state:
    st.session_state.liq_resultados = None
if 'retro_resultados' not in st.session_state:
    st.session_state.retro_resultados = None
if 'archivos_cargados' not in st.session_state:
    st.session_state.archivos_cargados = {"resoluciones": 0, "peticion": False}
if 'peticiones_texto' not in st.session_state:
    st.session_state.peticiones_texto = "Reconocimiento de Pensión de Vejez."
if 'exp_version' not in st.session_state:
    st.session_state.exp_version = 0

# Estados exclusivos del Módulo 5 (Autónomo)
if 'banco_motivas' not in st.session_state:
    st.session_state.banco_motivas = pd.DataFrame(columns=["Categoria", "Palabras_Clave", "Texto_Motiva"])
if 'mod5_texto_capturado' not in st.session_state:
    st.session_state.mod5_texto_capturado = ""
if 'mod5_motivas_encontradas' not in st.session_state:
    st.session_state.mod5_motivas_encontradas = []
if 'texto_formula_generada' not in st.session_state:
    st.session_state.texto_formula_generada = ""


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

    # Historia laboral (7 columnas con ID Aportante y Nombre / Razón Social)
    doc.add_heading('ANÁLISIS DE LA HISTORIA LABORAL:', level=2)
    doc.add_paragraph(
        f"Que, evaluada la trayectoria laboral del(la) señor(a) {datos_afi['nombre'].upper()}, se evidencia "
        f"la siguiente densidad de aportes al Sistema General de Pensiones. En cumplimiento del artículo 18 "
        f"de la Ley 100 de 1993, se aplicó la regla de simultaneidad, garantizando que para periodos concurrentes "
        f"se acumulen los Ingresos Base de Cotización (IBC) sin exceder 30 días mensuales o 360 anuales:"
    )

    table = doc.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    encabezados_hl = ['Periodo', 'ID Aportante', 'Nombre / Razón Social', 'Desde', 'Hasta', 'IBC Sumado', 'Semanas']
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(encabezados_hl):
        hdr_cells[i].text = ""
        run_h = hdr_cells[i].paragraphs[0].add_run(header_text)
        run_h.bold = True
        run_h.font.name = 'Arial'
        run_h.font.size = Pt(8.5)

    for _, row in df_historia.iterrows():
        row_cells = table.add_row().cells
        valores_fila = [
            str(row['Periodo']),
            str(row.get('NIT_Aportante', '')),
            str(row.get('Aportante', '')),
            row['Desde'].strftime('%d/%m/%Y') if pd.notnull(row['Desde']) else "",
            row['Hasta'].strftime('%d/%m/%Y') if pd.notnull(row['Hasta']) else "",
            f"${row['IBC']:,.0f}",
            f"{row['Semanas']:.2f}"
        ]
        for i, val in enumerate(valores_fila):
            row_cells[i].text = ""
            run_c = row_cells[i].paragraphs[0].add_run(val)
            run_c.font.name = 'Arial'
            run_c.font.size = Pt(8.5)

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
