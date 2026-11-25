import streamlit as st
import pandas as pd
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
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


# ==========================================
# UTILIDADES DE EDAD Y REQUISITOS EN TIEMPO REAL
# ==========================================
def calcular_edad_actual(fecha_nac):
    """Calcula la edad exacta actual en años y meses."""
    if not fecha_nac:
        return 0, 0, "No registrada"
    f_nac = pd.to_datetime(fecha_nac).date()
    hoy = date.today()
    if f_nac > hoy:
        return 0, 0, "Fecha inválida"
    rd = relativedelta(hoy, f_nac)
    texto = f"{rd.years} años y {rd.months} mes(es)" if rd.months > 0 else f"{rd.years} años"
    return rd.years, rd.months, texto


def _fmt_fecha(val, vacio="No acredita"):
    if val is None or pd.isna(val):
        return vacio
    return pd.to_datetime(val).strftime('%d/%m/%Y')


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


def _resumen_efectividad_texto(reconoce, retro):
    if not reconoce:
        return "No aplica (Prestación negada por no reunir requisitos de ley)."
    if not retro:
        return "Pendiente de parametrización de efectividad."
    if retro.get("condicionado_retiro_publico"):
        return "Condicionada a la acreditación del acto administrativo de retiro definitivo del servicio público (Art. 128 C.P.)."
    f_efec = _fmt_fecha(retro.get("fecha_efectividad"), "Inclusión en nómina")
    if retro.get("genera_retroactivo"):
        f_ini = _fmt_fecha(retro.get("fecha_inicio_pago"), f_efec)
        if retro.get("aplica_prescripcion"):
            return f"Efectividad: {f_efec} | Inicio de pago por prescripción trienal: {f_ini}"
        return f"Efectividad y pago a partir del {f_efec}"
    return f"A partir de la inclusión en nómina de pensionados ({f_efec})"


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
    fechas = liq.get("fechas", {})
    estatus_texto = "CONSOLIDADO" if estatus_cumplido else "NO CONSOLIDADO"
    fecha_est_str = _fmt_fecha(fechas.get("fecha_estatus"), "No consolidado") if estatus_cumplido else "No consolidado"
    efectividad_str = _resumen_efectividad_texto(estatus_cumplido, retro)

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
        f"Que el(la) señor(a) {datos_afi['nombre'].upper()}, identificado(a) con cédula de ciudadanía No. {datos_afi['cedula']}, "
        f"de género {datos_afi['genero'].lower()}, nacido(a) el {datos_afi['fecha_nac']} (con edad actual de {datos_afi['edad_texto']}), "
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

    # --- ESTUDIO DE ESTATUS PENSIONAL ---
    doc.add_heading('ESTUDIO DE DATOS BIOGRÁFICOS Y ESTATUS PENSIONAL:', level=2)
    doc.add_paragraph(
        f"• Fecha de Nacimiento: {datos_afi['fecha_nac']} | Edad Actual: {datos_afi['edad_texto']} "
        f"(Edad mínima exigida: {req['edad']} años — {'ACREDITA' if liq['cumple_edad'] else 'NO ACREDITA'}).\n"
        f"• Fecha de cumplimiento de edad mínima: {_fmt_fecha(fechas.get('fecha_cumple_edad'))}.\n"
        f"• Semanas válidas consolidadas: {liq['semanas']:,.2f} semanas "
        f"(Semanas mínimas exigidas: {req['semanas']} semanas — {'ACREDITA' if liq['cumple_sem'] else 'NO ACREDITA'}).\n"
        f"• Fecha de cumplimiento de semanas mínimas: {_fmt_fecha(fechas.get('fecha_cumple_semanas'), 'No acredita aún')}.\n"
        f"• Estatus Pensional: {estatus_texto} (Fecha de consolidación: {fecha_est_str}).\n"
        f"• Fecha de Corte / Indexación aplicada: {_fmt_fecha(fechas.get('fecha_corte'))} ({fechas.get('razon_corte', '')}).\n"
        f"• Efectividad Otorgada: {efectividad_str}"
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
        naturaleza_emp = "ENTIDAD PÚBLICA / SEM >= 90%" if info_empleador['es_publico'] else "SECTOR PRIVADO / INDEPENDIENTE"
        doc.add_paragraph(
            f"Último aportante registrado en historia laboral: {info_empleador['razon_social']} "
            f"(NIT/ID: {info_empleador['nit']}) — Naturaleza detectada: {naturaleza_emp}."
        )

    if estatus_cumplido:
        doc.add_paragraph(
            f"Que, descendiendo al caso concreto, el(la) solicitante acreditó el cumplimiento de {req['edad']} años "
            f"de edad (edad actual: {datos_afi['edad_texto']}) y superó el umbral de semanas exigidas ({req['semanas']} semanas), "
            f"consolidando el estatus pensional el {fecha_est_str}."
        )

        # --- ESTUDIO COMPARATIVO DE IBL ---
        doc.add_heading('ESTUDIO TÉCNICO DEL INGRESO BASE DE LIQUIDACIÓN (IBL):', level=2)
        doc.add_paragraph(
            f"De conformidad con lo dispuesto en el artículo 21 de la Ley 100 de 1993, el Ingreso Base de Liquidación (IBL) "
            f"se calculó indexando anualmente los Ingresos Base de Cotización (IBC) con base en el Índice de Precios al "
            f"Consumidor (IPC) certificado por el DANE hasta la fecha de corte ({_fmt_fecha(fechas.get('fecha_corte'))}). "
            f"En aplicación del principio de favorabilidad, se liquidaron ambos escenarios previstos en la norma:"
        )

        df_10 = liq.get("df_ibl_10", pd.DataFrame())
        df_vida = liq.get("df_ibl_vida", pd.DataFrame())
        rango_10 = f"{_fmt_fecha(df_10['Desde'].min())} al {_fmt_fecha(df_10['Hasta'].max())}" if not df_10.empty else "N/A"
        rango_vida = f"{_fmt_fecha(df_vida['Desde'].min())} al {_fmt_fecha(df_vida['Hasta'].max())}" if not df_vida.empty else "N/A"

        tbl_ibl = doc.add_table(rows=1, cols=5)
        tbl_ibl.style = 'Table Grid'
        hdr_ibl = ['Escenario Legal (Art. 21 Ley 100/93)', 'Periodos', 'Rango Evaluado', 'IBL Indexado', 'Selección']
        for i, h_txt in enumerate(hdr_ibl):
            tbl_ibl.rows[0].cells[i].text = ""
            r_h = tbl_ibl.rows[0].cells[i].paragraphs[0].add_run(h_txt)
            r_h.bold = True
            r_h.font.name = 'Arial'
            r_h.font.size = Pt(9)

        filas_ibl = [
            (
                "Últimos 10 Años Cotizados",
                str(len(df_10)),
                rango_10,
                f"${liq['ibl_10']:,.0f}",
                "APLICADO (Más favorable)" if liq['origen_ibl'] == "Últimos 10 Años" else "No aplicado"
            ),
            (
                "Toda la Vida Laboral",
                str(len(df_vida)),
                rango_vida,
                f"${liq['ibl_vida']:,.0f}",
                "APLICADO (Más favorable)" if liq['origen_ibl'] == "Toda la Vida" else "No aplicado"
            )
        ]
        for f_ibl in filas_ibl:
            rc = tbl_ibl.add_row().cells
            for i, val in enumerate(f_ibl):
                rc[i].text = ""
                r_c = rc[i].paragraphs[0].add_run(val)
                r_c.font.name = 'Arial'
                r_c.font.size = Pt(9)

        doc.add_paragraph(
            f"\nEn virtud del estudio comparativo anterior, se adopta como IBL definitivo el correspondiente a "
            f"{liq['origen_ibl']}, por un valor de ${liq['ibl']:,.0f} COP (año base {liq['anio_base']})."
        )

        # --- CÁLCULO PASO A PASO DE LA TASA DE REEMPLAZO ---
        f_data = liq['formula_tasa']
        doc.add_heading('CÁLCULO DE LA TASA DE REEMPLAZO (FÓRMULA DECRECIENTE LEY 797 DE 2003):', level=2)
        doc.add_paragraph(
            "Conforme al artículo 34 de la Ley 100 de 1993, modificado por el artículo 10 de la Ley 797 de 2003, "
            "la tasa de reemplazo se determina mediante la fórmula decreciente r = 65.50 - 0.50s (donde 's' equivale al "
            "número de SMMLV contenidos en el IBL), adicionando un 1.5% por cada 50 semanas cotizadas por encima del "
            "mínimo legal, con un límite normativo de 15 puntos porcentuales adicionales y un tope máximo total del 80%:"
        )
        doc.add_paragraph(
            f"1. SMMLV de referencia ({liq['anio_base']}): ${f_data['smmlv_ref']:,.0f} COP.\n"
            f"2. IBL definitivo aplicado ({liq['origen_ibl']}): ${liq['ibl']:,.0f} COP.\n"
            f"3. Proporción del IBL frente al SMMLV (s = IBL / SMMLV): {f_data['s']:.2f} salarios mínimos.\n"
            f"4. Tasa base inicial [65.50 - (0.50 x {f_data['s']:.2f})]: {f_data['tasa_base']:.2f}%.\n"
            f"5. Semanas válidas totales: {liq['semanas']:,.2f} | Semanas computables (tope 1.800): {f_data.get('semanas_usadas', liq['semanas']):,.2f} | Semanas mínimas exigidas: {f_data.get('semanas_minimas_exigidas', req['semanas'])}.\n"
            f"6. Semanas adicionales sobre el mínimo legal: {f_data['semanas_adicionales']:.2f} semanas, equivalentes a {f_data['bloques']} bloque(s) completo(s) de 50 semanas.\n"
            f"7. Incremento legal acumulado ({f_data['bloques']} x 1.5%, tope máximo 15 puntos): +{f_data['incremento']:.2f}%.\n"
            f"8. TASA DE REEMPLAZO FINAL DEFINITIVA (tope legal 80%): {f_data['tasa_final']:.2f}%."
        )
        doc.add_paragraph(
            f"LIQUIDACIÓN INICIAL DE LA MESADA: Aplicando la tasa de reemplazo del {f_data['tasa_final']:.2f}% sobre el "
            f"IBL de ${liq['ibl']:,.0f} COP (garantizando la pensión mínima de 1 SMMLV), se determina una mesada inicial "
            f"de ${liq['mesada']:,.0f} COP mensuales para el año base {liq['anio_base']}."
        )

        # --- EFECTIVIDAD, PRESCRIPCIÓN Y RETROACTIVO ---
        if retro:
            doc.add_heading('EFECTIVIDAD OTORGADA, PRESCRIPCIÓN Y LIQUIDACIÓN DE RETROACTIVO:', level=2)
            doc.add_paragraph(f"EFECTIVIDAD OTORGADA: {efectividad_str}")
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
                nota_m14 = (
                    "y mesada 14 (Acto Legislativo 01/2005)"
                    if retro['aplica_mesada_14']
                    else "(sin lugar a mesada 14 conforme al Acto Legislativo 01/2005)"
                )
                doc.add_paragraph(
                    f"Reajustando anualmente la mesada conforme al artículo 14 de la Ley 100 de 1993, "
                    f"aplicando la mesada adicional de noviembre (Art. 50 Ley 100/1993) {nota_m14}, "
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
            f"Que, descendiendo al caso concreto, el(la) solicitante (nacido(a) el {datos_afi['fecha_nac']}, "
            f"edad actual {datos_afi['edad_texto']}) NO CUMPLE con la totalidad de los requisitos exigidos "
            f"por la norma para consolidar el estatus de Pensión de Vejez."
        )
        if liq['cumple_sem'] and not liq['cumple_edad']:
            doc.add_paragraph(
                f"Si bien la historia laboral evidencia {liq['semanas']:,.2f} semanas cotizadas, "
                f"el(la) asegurado(a) NO ACREDITA el requisito de edad ({req['edad']} años para el género {datos_afi['genero'].lower()}), "
                f"cumpliéndolo únicamente hasta el {_fmt_fecha(fechas.get('fecha_cumple_edad'))}."
            )
        elif liq['cumple_edad'] and not liq['cumple_sem']:
            doc.add_paragraph(
                f"Si bien acredita el requisito de edad ({datos_afi['edad_texto']}, frente a los {req['edad']} años exigidos), "
                f"cuenta únicamente con {liq['semanas']:,.2f} semanas frente a las {req['semanas']} semanas exigidas por ley."
            )
        else:
            doc.add_paragraph(
                f"El(la) afiliado(a) NO ACREDITA el requisito de edad ({datos_afi['edad_texto']} frente a {req['edad']} años) "
                f"ni el requisito de semanas cotizadas ({req['semanas']} exigidas), contando únicamente con {liq['semanas']:,.2f} semanas."
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
                f"C.C. {datos_afi['cedula']}, nacido(a) el {datos_afi['fecha_nac']} ({datos_afi['edad_texto']}), "
                f"con estatus consolidado el {fecha_est_str}, liquidada sobre un IBL de ${liq['ibl']:,.0f} COP "
                f"({liq['origen_ibl']}) y una tasa de reemplazo del {liq['formula_tasa']['tasa_final']:.2f}%, "
                f"en cuantía actualizada de ${mesada_resuelve:,.0f} COP mensuales, dejando en suspenso su inclusión "
                f"en nómina y disfrute hasta tanto acredite el acto administrativo de retiro definitivo del servicio "
                f"público (Art. 128 C.P. y Ley 1821 de 2016)."
            )
        else:
            texto_efec = (
                f" con efectividad a partir del {retro['fecha_efectividad'].strftime('%d/%m/%Y')}"
                if retro else ""
            )
            doc.add_paragraph(
                f"ARTÍCULO PRIMERO: RECONOCER Y ORDENAR EL PAGO de una Pensión de Vejez a favor de "
                f"{datos_afi['nombre'].upper()}, C.C. {datos_afi['cedula']}, nacido(a) el {datos_afi['fecha_nac']} "
                f"({datos_afi['edad_texto']}), con estatus consolidado el {fecha_est_str}, liquidada sobre un IBL "
                f"de ${liq['ibl']:,.0f} COP ({liq['origen_ibl']}) y una tasa de reemplazo del "
                f"{liq['formula_tasa']['tasa_final']:.2f}%, en cuantía actualizada de ${mesada_resuelve:,.0f} COP "
                f"mensuales{texto_efec}."
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
            f"nacido(a) el {datos_afi['fecha_nac']} ({datos_afi['edad_texto']}), por no acreditar los requisitos "
            f"de ley conforme a la parte motiva."
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
        st.session_state.totales_pdf = None
        st.session_state.liq_resultados = None
        st.session_state.retro_resultados = None
        st.session_state.archivos_cargados = {"resoluciones": 0, "peticion": False}
        st.session_state.peticiones_texto = "Reconocimiento de Pensión de Vejez."
        st.session_state.exp_version = st.session_state.get("exp_version", 0) + 1
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
            hl_file.seek(0)
            st.session_state.totales_pdf = extraer_totales_oficiales_pdf(hl_file)
            st.session_state.df_final = None
            st.session_state.info_empleador = None
            st.session_state.liq_resultados = None
            st.session_state.retro_resultados = None

        if peticion_file:
            peticion_file.seek(0)
            datos_temp, pet_ext = extraer_datos_peticion(peticion_file, datos_temp)
            if pet_ext != "Reconocimiento de Pensión de Vejez.":
                st.session_state.peticiones_texto = pet_ext

        st.session_state.datos_basicos = datos_temp
        st.session_state.exp_version = st.session_state.get("exp_version", 0) + 1
        st.success("Cruce de datos completado.")
        st.rerun()

    st.divider()
    st.subheader("Datos del Asegurado (Extraídos/Editables)")

    v = st.session_state.get("exp_version", 0)
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.datos_basicos["cedula"] = st.text_input(
            "Cédula",
            value=st.session_state.datos_basicos["cedula"],
            key=f"ced_{v}"
        )
        st.session_state.datos_basicos["nombre"] = st.text_input(
            "Nombre Completo",
            value=st.session_state.datos_basicos["nombre"],
            key=f"nom_{v}"
        )
    with c2:
        st.session_state.datos_basicos["fecha_nac"] = st.date_input(
            "Fecha de Nacimiento",
            value=st.session_state.datos_basicos["fecha_nac"],
            min_value=date(1900, 1, 1),
            max_value=date.today(),
            key=f"fnac_{v}"
        )
        idx_gen = 0 if st.session_state.datos_basicos["genero"] == "Masculino" else 1
        st.session_state.datos_basicos["genero"] = st.radio(
            "Género",
            ["Masculino", "Femenino"],
            index=idx_gen,
            horizontal=True,
            key=f"gen_{v}"
        )

    edad_anios, _, edad_texto = calcular_edad_actual(st.session_state.datos_basicos["fecha_nac"])
    edad_min_ref = 62 if st.session_state.datos_basicos["genero"] == "Masculino" else 57
    st.info(
        f"🎂 **Fecha de Nacimiento:** {st.session_state.datos_basicos['fecha_nac'].strftime('%d/%m/%Y')}  |  "
        f"**Edad Actual:** {edad_texto}  |  "
        f"**Edad Mínima de Ley ({st.session_state.datos_basicos['genero']}):** {edad_min_ref} años "
        f"({'✅ Cumplida' if edad_anios >= edad_min_ref else '⏳ Pendiente'})"
    )

    st.divider()
    st.subheader("Peticiones a Resolver de Fondo")
    st.session_state.peticiones_texto = st.text_area(
        "Pretensiones capturadas:",
        value=st.session_state.peticiones_texto,
        height=120,
        key=f"pet_txt_{v}"
    )

# ------------------------------------------
# MÓDULO 2: HISTORIA LABORAL + DETECTOR EMPLEADOR PÚBLICO + AUDITORÍA
# ------------------------------------------
with mod2:
    st.header("📋 Procesamiento de la Historia Laboral")
    _, _, edad_txt_m2 = calcular_edad_actual(st.session_state.datos_basicos["fecha_nac"])
    st.info(
        f"Asegurado: **{st.session_state.datos_basicos['nombre']}** | "
        f"C.C. **{st.session_state.datos_basicos['cedula']}** | "
        f"Nacimiento: **{st.session_state.datos_basicos['fecha_nac'].strftime('%d/%m/%Y')}** ({edad_txt_m2})"
    )

    if st.session_state.df_crudo is not None and not st.session_state.df_crudo.empty:
        df = st.session_state.df_crudo
        cols = df.columns.tolist()

        def idx_col(nombre_esperado, fallback):
            return cols.index(nombre_esperado) if nombre_esperado in cols else min(fallback, len(cols) - 1)

        with st.expander("👁️ Ver tabla cruda extraída del PDF (Parte General)", expanded=False):
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

            sem_capturadas = st.session_state.df_final['Semanas'].sum()
            tot_of = st.session_state.totales_pdf or {}
            sem_pdf_26 = tot_of.get("total_26")

            if sem_pdf_26 is not None:
                diff = abs(sem_capturadas - sem_pdf_26)
                c_aud1, c_aud2, c_aud3 = st.columns(3)
                c_aud1.metric("Semanas Capturadas (Parte General)", f"{sem_capturadas:,.2f}")
                c_aud2.metric("Total Oficial PDF [26]", f"{sem_pdf_26:,.2f}")
                c_aud3.metric("Diferencia (Redondeo)", f"{sem_capturadas - sem_pdf_26:+.2f}")

                if diff <= 1.5:
                    st.success("✅ Auditoría aprobada: La captura coincide con la Parte General oficial de Colpensiones.")
                else:
                    st.error(
                        f"⚠️ Alerta de auditoría: Diferencia de {diff:.2f} semanas frente al campo [26] del PDF. "
                        f"Revise la tabla cruda desplegable antes de liquidar."
                    )

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
# MÓDULO 3: LIQUIDACIÓN, ESTUDIO IBL, TASA DE REEMPLAZO, ESTATUS Y RETROACTIVO
# ------------------------------------------
with mod3:
    st.header("🧮 Estudio de Fondo: Estatus, IBL, Tasa de Reemplazo y Retroactivo")
    _, _, edad_txt_m3 = calcular_edad_actual(st.session_state.datos_basicos["fecha_nac"])
    st.info(
        f"Asegurado: **{st.session_state.datos_basicos['nombre']}** | "
        f"C.C. **{st.session_state.datos_basicos['cedula']}** | "
        f"Fecha Nacimiento: **{st.session_state.datos_basicos['fecha_nac'].strftime('%d/%m/%Y')}** | "
        f"Edad Actual: **{edad_txt_m3}** | "
        f"Género: **{st.session_state.datos_basicos['genero']}**"
    )

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

        if st.button("Calcular Derecho, Estudio Técnico y Retroactivo", type="primary"):
            df_f = st.session_state.df_final
            genero = st.session_state.datos_basicos["genero"]
            fecha_nac = st.session_state.datos_basicos["fecha_nac"]

            liq = LiquidadorPension(df_f, genero, fecha_nac)
            fechas = liq.determinar_fechas_clave()

            anio_base = fechas['fecha_corte'].year
            ibl_10, df_ibl_10 = liq.calcular_ibl_indexado(fechas['fecha_corte'], "ultimos_10")
            ibl_vida, df_ibl_vida = liq.calcular_ibl_indexado(fechas['fecha_corte'], "toda_vida")
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
                "df_ibl_10": df_ibl_10,
                "df_ibl_vida": df_ibl_vida,
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
            fechas = r['fechas']
            f_tasa = r['formula_tasa']

            # --- 1. ESTATUS PENSIONAL Y EFECTIVIDAD ---
            st.divider()
            if r['reconoce']:
                st.success("✅ PROCEDE EL RECONOCIMIENTO — ESTATUS PENSIONAL CONSOLIDADO")
            else:
                st.error("❌ NO ACREDITA REQUISITOS DE LEY — ESTATUS NO CONSOLIDADO")

            st.subheader("1. Estatus Pensional, Edad y Efectividad Otorgada")
            e1, e2, e3, e4 = st.columns(4)
            e1.metric(
                "Edad Actual",
                edad_txt_m3,
                f"Nac: {st.session_state.datos_basicos['fecha_nac'].strftime('%d/%m/%Y')}"
            )
            e2.metric(
                "Estatus Pensional",
                "CONSOLIDADO" if r['reconoce'] else "NO CONSOLIDADO",
                _fmt_fecha(fechas.get('fecha_estatus'), "Pendiente") if r['reconoce'] else "Faltan requisitos"
            )
            e3.metric(
                "Semanas Válidas",
                f"{r['semanas']:,.2f}",
                f"Exigidas: {r['sem_req']}"
            )
            e4.metric(
                "Efectividad Otorgada",
                _fmt_fecha(retro['fecha_efectividad'], "Inclusión nómina") if (r['reconoce'] and retro and not retro['condicionado_retiro_publico']) else ("Retiro público" if (retro and retro['condicionado_retiro_publico']) else "No aplica"),
                "Con prescripción" if (retro and retro['aplica_prescripcion']) else ("Disfrute directo" if r['reconoce'] else "Negación")
            )

            df_estatus = pd.DataFrame([
                {"Concepto": "Fecha de Nacimiento / Edad Actual", "Detalle": f"{st.session_state.datos_basicos['fecha_nac'].strftime('%d/%m/%Y')} ({edad_txt_m3})"},
                {"Concepto": "Requisito de Edad y Cumplimiento", "Detalle": f"{r['edad_req']} años — Fecha cumplimiento: {_fmt_fecha(fechas.get('fecha_cumple_edad'))} ({'Cumplida' if r['cumple_edad'] else 'Pendiente'})"},
                {"Concepto": "Requisito de Semanas y Cumplimiento", "Detalle": f"{r['sem_req']} semanas ({r['nota_req']}) — Fecha cumplimiento: {_fmt_fecha(fechas.get('fecha_cumple_semanas'), 'No acredita aún')}"},
                {"Concepto": "Consolidación del Estatus Pensional", "Detalle": f"{'CONSOLIDADO el ' + _fmt_fecha(fechas.get('fecha_estatus')) if r['reconoce'] else 'NO CONSOLIDADO'}"},
                {"Concepto": "Fecha de Corte para Indexación (IPC)", "Detalle": f"{_fmt_fecha(fechas.get('fecha_corte'))} — {fechas.get('razon_corte', '')}"},
                {"Concepto": "Efectividad Otorgada", "Detalle": _resumen_efectividad_texto(r['reconoce'], retro)}
            ])
            st.table(df_estatus)

            # --- 2. ESTUDIO COMPARATIVO DEL IBL ---
            st.divider()
            st.subheader("2. Estudio del Cálculo del IBL (Últimos 10 Años vs. Toda la Vida Laboral)")
            if not r['reconoce']:
                st.caption("ℹ️ Proyección técnica informativa (no genera derecho al no acreditarse los requisitos de ley).")

            df_10 = r.get("df_ibl_10", pd.DataFrame())
            df_vida = r.get("df_ibl_vida", pd.DataFrame())
            rango_10 = f"{_fmt_fecha(df_10['Desde'].min())} al {_fmt_fecha(df_10['Hasta'].max())}" if not df_10.empty else "Sin registros"
            rango_vida = f"{_fmt_fecha(df_vida['Desde'].min())} al {_fmt_fecha(df_vida['Hasta'].max())}" if not df_vida.empty else "Sin registros"

            i1, i2, i3 = st.columns(3)
            i1.metric(
                "IBL Últimos 10 Años",
                f"${r['ibl_10']:,.0f}",
                "✅ Más favorable" if r['origen_ibl'] == "Últimos 10 Años" else f"{len(df_10)} periodos"
            )
            i2.metric(
                "IBL Toda la Vida Laboral",
                f"${r['ibl_vida']:,.0f}",
                "✅ Más favorable" if r['origen_ibl'] == "Toda la Vida" else f"{len(df_vida)} periodos"
            )
            i3.metric(
                f"IBL Definitivo ({r['anio_base']})",
                f"${r['ibl']:,.0f}",
                f"Origen: {r['origen_ibl']}"
            )

            df_comp_ibl = pd.DataFrame([
                {
                    "Escenario (Art. 21 Ley 100/93)": "Últimos 10 Años Cotizados",
                    "Periodos": len(df_10),
                    "Rango Evaluado": rango_10,
                    "IBL Indexado": r['ibl_10'],
                    "Resultado": "✅ APLICADO (Más favorable)" if r['origen_ibl'] == "Últimos 10 Años" else "Descartado"
                },
                {
                    "Escenario (Art. 21 Ley 100/93)": "Toda la Vida Laboral",
                    "Periodos": len(df_vida),
                    "Rango Evaluado": rango_vida,
                    "IBL Indexado": r['ibl_vida'],
                    "Resultado": "✅ APLICADO (Más favorable)" if r['origen_ibl'] == "Toda la Vida" else "Descartado"
                }
            ])
            st.dataframe(
                df_comp_ibl.style.format({"IBL Indexado": "${:,.0f}"}),
                use_container_width=True,
                hide_index=True
            )

            col_det1, col_det2 = st.columns(2)
            with col_det1:
                with st.expander(f"🔎 Ver detalle de indexación IPC — Últimos 10 Años ({len(df_10)} periodos)"):
                    if not df_10.empty:
                        st.dataframe(
                            df_10.style.format({
                                'IBC_Historico': "${:,.0f}",
                                'Factor_IPC': "{:.4f}",
                                'IBC_Actualizado': "${:,.0f}",
                                'Semanas': "{:.2f}"
                            }),
                            use_container_width=True
                        )
            with col_det2:
                with st.expander(f"🔎 Ver detalle de indexación IPC — Toda la Vida Laboral ({len(df_vida)} periodos)"):
                    if not df_vida.empty:
                        st.dataframe(
                            df_vida.style.format({
                                'IBC_Historico': "${:,.0f}",
                                'Factor_IPC': "{:.4f}",
                                'IBC_Actualizado': "${:,.0f}",
                                'Semanas': "{:.2f}"
                            }),
                            use_container_width=True
                        )

            # --- 3. CÁLCULO PASO A PASO DE LA TASA DE REEMPLAZO ---
            st.divider()
            st.subheader("3. Detalle del Cálculo de la Tasa de Reemplazo (Ley 797 de 2003)")
            t1, t2, t3, t4 = st.columns(4)
            t1.metric("Relación s (IBL / SMMLV)", f"{f_tasa['s']:.2f} SMMLV")
            t2.metric("Tasa Base (65.5 - 0.5s)", f"{f_tasa['tasa_base']:.2f}%")
            t3.metric("Incremento Sem. Extra", f"+{f_tasa['incremento']:.2f}%", f"{f_tasa['bloques']} bloque(s) de 50")
            t4.metric("Tasa de Reemplazo Final", f"{f_tasa['tasa_final']:.2f}%", f"Mesada ({r['anio_base']}): ${r['mesada']:,.0f}")

            df_pasos_tasa = pd.DataFrame([
                {"Paso": "1. SMMLV de referencia", "Fórmula / Criterio": f"Año base de liquidación ({r['anio_base']})", "Resultado": f"${f_tasa['smmlv_ref']:,.0f}"},
                {"Paso": "2. Proporción IBL / SMMLV (s)", "Fórmula / Criterio": f"${r['ibl']:,.0f} / ${f_tasa['smmlv_ref']:,.0f}", "Resultado": f"{f_tasa['s']:.2f} salarios mínimos"},
                {"Paso": "3. Tasa base inicial (r)", "Fórmula / Criterio": f"65.50 - (0.50 × {f_tasa['s']:.2f}) [Piso 55% / Techo 65.5%]", "Resultado": f"{f_tasa['tasa_base']:.2f}%"},
                {"Paso": "4. Semanas computables vs. mínimas", "Fórmula / Criterio": f"Válidas: {r['semanas']:,.2f} | Tope 1.800: {f_tasa.get('semanas_usadas', r['semanas']):,.2f} | Mínimo legal: {f_tasa.get('semanas_minimas_exigidas', r['sem_req'])}", "Resultado": f"{f_tasa['semanas_adicionales']:.2f} semanas extra"},
                {"Paso": "5. Bloques completos de 50 semanas", "Fórmula / Criterio": f"Parte entera({f_tasa['semanas_adicionales']:.2f} / 50)", "Resultado": f"{f_tasa['bloques']} bloque(s)"},
                {"Paso": "6. Incremento porcentual acumulado", "Fórmula / Criterio": f"{f_tasa['bloques']} × 1.50% (límite normativo: máx. 15 puntos)", "Resultado": f"+{f_tasa['incremento']:.2f}%"},
                {"Paso": "7. Tasa de Reemplazo Definitiva", "Fórmula / Criterio": f"{f_tasa['tasa_base']:.2f}% + {f_tasa['incremento']:.2f}% (tope legal máximo: 80%)", "Resultado": f"{f_tasa['tasa_final']:.2f}%"},
                {"Paso": f"8. Mesada Inicial ({r['anio_base']})", "Fórmula / Criterio": f"${r['ibl']:,.0f} × {f_tasa['tasa_final']:.2f}% (garantía mínima 1 SMMLV)", "Resultado": f"${r['mesada']:,.0f}"}
            ])
            st.table(df_pasos_tasa)

            # --- 4. RETROACTIVO (SI RECONOCE) ---
            if r['reconoce'] and retro:
                st.divider()
                st.subheader("4. Estudio de Efectividad, Prescripción y Retroactivo")
                if retro["condicionado_retiro_publico"]:
                    st.warning(f"🏛️ {retro['nota_efectividad']}")
                else:
                    st.info(retro['nota_efectividad'])

                if retro["aplica_prescripcion"]:
                    st.warning(
                        f"⚠️ **Prescripción Trienal Aplicada:** Efectividad del {retro['fecha_efectividad'].strftime('%d/%m/%Y')}, "
                        f"pero pago liquidado a partir del **{retro['fecha_inicio_pago'].strftime('%d/%m/%Y')}** (Art. 488 CST / 151 CPTSS)."
                    )

                c_mes1, c_mes2 = st.columns(2)
                c_mes1.metric(f"Mesada Inicial ({r['anio_base']})", f"${r['mesada']:,.0f}")
                c_mes2.metric("Mesada Actualizada al Ingreso en Nómina", f"${retro['mesada_actualizada']:,.0f}")

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
        st.warning("Debe procesar la Historia Laboral en el Módulo 2.")

# ------------------------------------------
# MÓDULO 4: RESULTADO VISUAL Y DESCARGA
# ------------------------------------------
with mod4:
    st.header("📄 Sustanciación Final (Resolución Completa)")

    if st.session_state.liq_resultados is not None:
        r = st.session_state.liq_resultados
        retro = st.session_state.retro_resultados
        fechas = r['fechas']
        f_tasa = r['formula_tasa']
        estado_texto = "RECONOCIMIENTO" if r['reconoce'] else "NEGACIÓN"
        _, _, edad_txt_m4 = calcular_edad_actual(st.session_state.datos_basicos["fecha_nac"])
        efectividad_m4 = _resumen_efectividad_texto(r['reconoce'], retro)

        st.subheader("Resumen Ejecutivo del Acto Administrativo")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Asegurado / C.C.", st.session_state.datos_basicos["cedula"], st.session_state.datos_basicos["nombre"])
        c2.metric("Edad Actual", edad_txt_m4, f"Nac: {st.session_state.datos_basicos['fecha_nac'].strftime('%d/%m/%Y')}")
        c3.metric("Estatus Pensional", "CONSOLIDADO" if r['reconoce'] else "NO CONSOLIDADO", _fmt_fecha(fechas.get('fecha_estatus'), "N/A") if r['reconoce'] else "Sin consolidar")
        c4.metric("Sentido del Acto", estado_texto, delta_color="normal" if r['reconoce'] else "inverse")

        c5, c6, c7, c8 = st.columns(4)
        c5.metric("IBL 10 Últimos Años", f"${r['ibl_10']:,.0f}")
        c6.metric("IBL Toda la Vida", f"${r['ibl_vida']:,.0f}")
        c7.metric("Tasa de Reemplazo", f"{f_tasa['tasa_final']:.2f}%", f"IBL usado: {r['origen_ibl']}")
        c8.metric(
            "Retroactivo Neto",
            f"${retro['total_neto']:,.0f}" if (r['reconoce'] and retro and retro['genera_retroactivo']) else "$0"
        )

        st.info(f"**Efectividad Otorgada:** {efectividad_m4}")

        with st.expander("👁️ Vista Previa del Estudio Técnico que se incluirá en la Resolución Word", expanded=True):
            st.markdown(
                f"**1. Datos Biográficos y Estatus:**\n"
                f"- **Nombre:** {st.session_state.datos_basicos['nombre'].upper()} | **C.C.:** {st.session_state.datos_basicos['cedula']} | **Género:** {st.session_state.datos_basicos['genero']}\n"
                f"- **Fecha de Nacimiento:** {st.session_state.datos_basicos['fecha_nac'].strftime('%d/%m/%Y')} | **Edad Actual:** {edad_txt_m4}\n"
                f"- **Semanas Válidas:** {r['semanas']:,.2f} / Exigidas: {r['sem_req']} ({r['nota_req']})\n"
                f"- **Estatus Pensional:** {'CONSOLIDADO el ' + _fmt_fecha(fechas.get('fecha_estatus')) if r['reconoce'] else 'NO CONSOLIDADO'}\n"
                f"- **Efectividad Otorgada:** {efectividad_m4}\n\n"
                f"**2. Estudio Comparativo del IBL (Año base {r['anio_base']}):**\n"
                f"- **Últimos 10 Años:** ${r['ibl_10']:,.0f} COP\n"
                f"- **Toda la Vida Laboral:** ${r['ibl_vida']:,.0f} COP\n"
                f"- **IBL Definitivo Más Favorable ({r['origen_ibl']}):** ${r['ibl']:,.0f} COP\n\n"
                f"**3. Cálculo de la Tasa de Reemplazo (Ley 797 de 2003):**\n"
                f"- **SMMLV ({r['anio_base']}):** ${f_tasa['smmlv_ref']:,.0f} | **Proporción s:** {f_tasa['s']:.2f}\n"
                f"- **Tasa Base (65.50 - 0.50s):** {f_tasa['tasa_base']:.2f}% | **Semanas Extra:** {f_tasa['semanas_adicionales']:.2f} ({f_tasa['bloques']} bloques × 1.5% = +{f_tasa['incremento']:.2f}%)\n"
                f"- **Tasa Final Aplicada:** {f_tasa['tasa_final']:.2f}% | **Mesada Inicial ({r['anio_base']}):** ${r['mesada']:,.0f} COP"
            )

        st.divider()
        req_data = {"edad": r['edad_req'], "semanas": r['sem_req'], "nota": r['nota_req']}
        datos_afi = {
            "nombre": st.session_state.datos_basicos["nombre"],
            "cedula": st.session_state.datos_basicos["cedula"],
            "fecha_nac": st.session_state.datos_basicos["fecha_nac"].strftime('%d/%m/%Y'),
            "edad_texto": edad_txt_m4,
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