import pdfplumber
import pandas as pd
import re
from datetime import datetime, date

try:
    import pytesseract
    OCR_DISPONIBLE = True
except ImportError:
    OCR_DISPONIBLE = False

# ==========================================
# CATÁLOGO DE ENTIDADES PÚBLICAS Y SEM (>= 90% CAPITAL ESTATAL)
# ==========================================
ENTIDADES_PUBLICAS_NITS = {
    "899999068": "Ecopetrol S.A.",
    "800194208": "Gensa S.A. E.S.P.",
    "812000370": "Urrá S.A. E.S.P.",
    "811000740": "Isagen (Histórico Público)",
    "800037800": "Banco Agrario de Colombia S.A.",
    "800096329": "Findeter S.A.",
    "800149923": "Bancóldex S.A.",
    "800116398": "Finagro",
    "860011153": "Positiva Compañía de Seguros S.A.",
    "860002400": "Previsora S.A. Compañía de Seguros",
    "860525148": "Fiduciaria La Previsora S.A. (Fiduprevisora)",
    "800178148": "Fiducoldex",
    "899999284": "Fondo Nacional del Ahorro (FNA)",
    "899999035": "ICETEX",
    "860042945": "Central de Inversiones S.A. (CISA)",
    "900062917": "Servicios Postales Nacionales S.A. (4-72)",
    "900002583": "RTVC - Sistema de Medios Públicos",
    "899999044": "Industria Militar (Indumil)",
    "899999143": "Satena S.A.",
    "899999270": "CIAC S.A.",
    "901038962": "Empresa Metro de Bogotá S.A.",
    "899999094": "Empresa de Acueducto y Alcantarillado de Bogotá (EAAB)",
    "890904996": "Empresas Públicas de Medellín (EPM)",
    "890399003": "EMCALI EICE ESP",
    "899999115": "ETB S.A. ESP",
    "900336004": "Colpensiones",
    "900505060": "Coljuegos",
    "899999002": "Vecol S.A.",
    "860007887": "Artesanías de Colombia S.A.",
    "830001113": "Imprenta Nacional de Colombia",
    "899999001": "Nación / Entidad del Orden Nacional",
    "899999061": "Distrito Capital de Bogotá",
    "899999028": "SENA",
    "899999239": "ICBF",
    "800197268": "DIAN",
}

PATRONES_ENTIDAD_PUBLICA = [
    r"\bMINISTERIO\b", r"\bDEPARTAMENTO\s+ADMINISTRATIVO\b", r"\bSUPERINTENDENCIA\b",
    r"\bUNIDAD\s+ADMINISTRATIVA\s+ESPECIAL\b", r"\bGOBERNACION\b", r"\bALCALDIA\b",
    r"\bMUNICIPIO\s+DE\b", r"\bDISTRITO\s+(?:CAPITAL|ESPECIAL|TURISTICO)\b",
    r"\bSECRETARIA\s+DE\b", r"\bCONCEJO\s+(?:MUNICIPAL|DISTRITAL|DE)\b",
    r"\bASAMBLEA\s+DEPARTAMENTAL\b", r"\bCONTRALORIA\b", r"\bPERSONERIA\b",
    r"\bPROCURADURIA\b", r"\bDEFENSORIA\s+DEL\s+PUEBLO\b", r"\bFISCALIA\s+GENERAL\b",
    r"\bRAMA\s+JUDICIAL\b", r"\bCONSEJO\s+SUPERIOR\s+DE\s+LA\s+JUDICATURA\b",
    r"\bREGISTRADURIA\b", r"\bEMPRESA\s+SOCIAL\s+DEL\s+ESTADO\b", r"\bE\.?S\.?E\.?\b",
    r"\bHOSPITAL\s+(?:DEPARTAMENTAL|MUNICIPAL|UNIVERSITARIO|REGIONAL|LOCAL|SIQUIATRICO|PSIQUIATRICO)\b",
    r"\bUNIVERSIDAD\s+(?:NACIONAL|DISTRITAL|DE\s+ANTIOQUIA|DEL\s+VALLE|INDUSTRIAL\s+DE\s+SANTANDER|PEDAGOGICA|DEL\s+ATLANTICO|DE\s+CALDAS|DEL\s+CAUCA|DE\s+CARTAGENA|DEL\s+TOLIMA|DE\s+NARIÑO|DE\s+CORDOBA|SURCOLOMBIANA|DE\s+CUNDINAMARCA|TECNOLOGICA)\b",
    r"\bSENA\b", r"\bSERVICIO\s+NACIONAL\s+DE\s+APRENDIZAJE\b", r"\bICBF\b",
    r"\bBIENESTAR\s+FAMILIAR\b", r"\bPOLICIA\s+NACIONAL\b", r"\bEJERCITO\s+NACIONAL\b",
    r"\bARMADA\s+NACIONAL\b", r"\bFUERZA\s+AEREA\b", r"\bFUERZA\s+AEROESPACIAL\b",
    r"\bINPEC\b", r"\bCREMIL\b", r"\bCASUR\b", r"\bCOLPENSIONES\b",
    r"\bADMINISTRADORA\s+COLOMBIANA\s+DE\s+PENSIONES\b", r"\bUGPP\b",
    r"\bDIAN\b", r"\bECOPETROL\b", r"\bBANCO\s+AGRARIO\b", r"\bFINDETER\b",
    r"\bBANCOLDEX\b", r"\bFINAGRO\b", r"\bPOSITIVA\b", r"\bPREVISORA\b",
    r"\bFIDUPREVISORA\b", r"\bFIDUCOLDEX\b", r"\bFONDO\s+NACIONAL\s+DEL\s+AHORRO\b",
    r"\bICETEX\b", r"\bCISA\b", r"\bSERVICIOS\s+POSTALES\s+NACIONALES\b", r"\bRTVC\b",
    r"\bINDUMIL\b", r"\bSATENA\b", r"\bCOTECMAR\b", r"\bCIAC\b", r"\bINVIAS\b",
    r"\bAEROCIVIL\b", r"\bMETRO\s+DE\s+(?:BOGOTA|MEDELLIN)\b", r"\bACUEDUCTO\s+Y\s+ALCANTARILLADO\b",
    r"\bEMPRESAS\s+PUBLICAS\s+DE\b", r"\bEMCALI\b", r"\bEICE\b", r"\bCOLJUEGOS\b",
    r"\bVECOL\b", r"\bARTESANIAS\s+DE\s+COLOMBIA\b", r"\bIMPRENTA\s+NACIONAL\b",
    r"\bGENSA\b", r"\bURRA\s+S\.?A\.?\b", r"\bCEDENAR\b", r"\bELECTROHUILA\b",
    r"\bELECTROCAQUETA\b", r"\bDISPAC\b", r"\bEBSA\b", r"\bTELEANTIOQUIA\b",
    r"\bTELECARIBE\b", r"\bTELEPACIFICO\b", r"\bCANAL\s+CAPITAL\b", r"\bCANAL\s+TRECE\b",
    r"\bAGENCIA\s+NACIONAL\b", r"\bINSTITUTO\s+(?:NACIONAL|COLOMBIANO|DEPARTAMENTAL|DISTRITAL)\b",
    r"\bCORPORACION\s+AUTONOMA\s+REGIONAL\b", r"\bCAJANAL\b", r"\bCAPRECOM\b"
]

COLUMNAS_RESUMEN = [
    "Identificación Aportante",
    "Nombre o Razón Social",
    "Desde",
    "Hasta",
    "Último Salario",
    "Semanas",
    "Lic",
    "Sim",
    "Total"
]

# Marcadores duros que delimitan el FIN de la Parte General (Resumen) e INICIO de la Parte Detallada
MARCADORES_FIN_PARTE_GENERAL = [
    "RESUMEN TIEMPO PÚBLICO SIMULTÁNEO",
    "RESUMEN TIEMPO PUBLICO SIMULTANEO",
    "[22]DESDE",
    "[22] DESDE",
    "[26]TOTAL SEMANAS",
    "[26] TOTAL SEMANAS",
    "DETALLE DE PAGOS EFECTUADOS ANTERIORES A 1995",
    "DETALLE DE PAGOS EFECTUADOS A PARTIR DE 1995",
    "DETALLE DE PERIODOS REPORTADOS POR ENTIDADES",
    "[27] IDENTIFICACIÓN",
    "[27] IDENTIFICACION",
    "[34]",
    "[47]",
    "LECTURA DEL REPORTE DE LA HISTORIA LABORAL"
]


# ==========================================
# UTILIDADES DE LIMPIEZA Y VALIDACIÓN
# ==========================================
def _limpiar_texto_colpensiones(texto):
    if not texto:
        return ""
    t = str(texto)
    t = t.replace("Ã‘", "Ñ").replace("Ã±", "ñ")
    t = re.sub(r'([A-Z])ÃO\b', r'\1ÑO', t)
    t = re.sub(r'([A-Z])Ã([AEIOU])', r'\1Ñ\2', t)
    return re.sub(r'\s+', ' ', t).strip()


def _limpiar_id_aportante(val):
    return re.sub(r'\D', '', str(val or ''))


def _clean_num(val):
    """Convierte valores monetarios y decimales en formato colombiano/Colpensiones a float."""
    if val is None or str(val).lower() == 'none':
        return 0.0
    v = re.sub(r'[^\d\.,]', '', str(val))
    if not v:
        return 0.0

    if ',' in v and '.' in v:
        v = v.replace('.', '').replace(',', '.')
    elif v.count('.') > 1:
        v = v.replace('.', '')
    elif ',' in v:
        if len(v.split(',')[-1]) <= 2:
            v = v.replace(',', '.')
        else:
            v = v.replace(',', '')
    elif '.' in v:
        if len(v.split('.')[-1]) == 3:
            v = v.replace('.', '')

    try:
        return float(v)
    except Exception:
        return 0.0


def _separar_id_y_nombre(texto):
    t = _limpiar_texto_colpensiones(texto)
    if not t:
        return "(Sin ID)", ""

    m = re.match(
        r'^(?:(?:NIT|NI|CC|CE|TI)[\s:\.\-]*)?(\d[\d\.\-]{3,15})(?:\s+|(?=[A-ZÑÁÉÍÓÚ]))(.*)$',
        t,
        re.IGNORECASE
    )
    if m:
        nit = _limpiar_id_aportante(m.group(1))
        nom = _limpiar_texto_colpensiones(m.group(2))
        if nit:
            return nit, nom

    m2 = re.match(r'^(\d{2}\s+\d{4}\s+\d{3,6})\s+(.*)$', t)
    if m2:
        return _limpiar_id_aportante(m2.group(1)), _limpiar_texto_colpensiones(m2.group(2))

    return "(Sin ID)", t


def _extraer_texto_pagina_con_ocr(page):
    text = page.extract_text(layout=False) or ""
    if len(text.strip()) >= 40:
        return text, False

    if OCR_DISPONIBLE:
        try:
            pil_img = page.to_image(resolution=250).original
            ocr_text = pytesseract.image_to_string(pil_img, lang='spa', config='--psm 6')
            return ocr_text or "", True
        except Exception:
            return text, False
    return text, False


def _validar_fila_resumen(nit, nom, f_desde, f_hasta, sal_str, sem_str, lic_str, sim_str, tot_str):
    """
    Aplica parámetros estrictos de auditoría sobre cada fila de la Parte General:
    1. Fechas válidas (Hasta >= Desde).
    2. Los 4 campos finales son numéricos no negativos (excluye filas del Detalle).
    3. Identidad aritmética: Total == Semanas - Lic - Sim (autocorrige fallos OCR como '414' -> '4,14').
    4. Límite físico real del periodo: Total <= ((Hasta - Desde + 1) / 7) + 0.5 (sin destruir licencias).
    """
    d1 = pd.to_datetime(f_desde, dayfirst=True, errors='coerce')
    d2 = pd.to_datetime(f_hasta, dayfirst=True, errors='coerce')
    if pd.isna(d1) or pd.isna(d2) or d2 < d1:
        return None

    # Rechazar cualquier valor negativo explícito (propio de licencias en el Detalle Pre-1995)
    for campo in (sem_str, lic_str, sim_str, tot_str):
        c_str = str(campo).strip()
        if c_str.startswith("-") or not re.fullmatch(r'[\d\.,]+', c_str):
            return None

    sem_v = _clean_num(sem_str)
    lic_v = _clean_num(lic_str)
    sim_v = _clean_num(sim_str)
    tot_v = _clean_num(tot_str)

    max_fisico = ((d2 - d1).days + 1) / 7.0 + 0.5

    # Si Semanas supera el máximo físico del rango, es una fila ajena o corrupta
    if sem_v > max_fisico + 1.0:
        return None

    # Identidad aritmética de Colpensiones: [9] Total = [6] Semanas - [7] Lic - [8] Sim
    esperado = max(0.0, sem_v - lic_v - sim_v)
    if abs(esperado - tot_v) > 0.15:
        tot_v = round(esperado, 2)
        tot_str = f"{tot_v:.2f}".replace(".", ",")

    if tot_v > max_fisico:
        tot_v = round(min(esperado, max_fisico), 2)
        tot_str = f"{tot_v:.2f}".replace(".", ",")

    nit_clean = _limpiar_id_aportante(nit)
    return [
        nit_clean if nit_clean else "(Sin ID)",
        _limpiar_texto_colpensiones(nom),
        f_desde,
        f_hasta,
        str(sal_str).strip(),
        str(sem_str).strip(),
        str(lic_str).strip(),
        str(sim_str).strip(),
        str(tot_str).strip()
    ]


def _extraer_catalogo_nombres_detalle(full_text):
    """
    Lee de forma pasiva los nombres completos de aportantes en el Detalle Post-1995 ([34]/[35])
    únicamente para completar razones sociales truncadas a 20 caracteres en el Resumen.
    Nunca agrega ni modifica periodos o semanas.
    """
    mapa = {}

    def registrar(nit_raw, nom_raw):
        nit = _limpiar_id_aportante(nit_raw)
        if not (5 <= len(nit) <= 12):
            return
        nom = _limpiar_texto_colpensiones(nom_raw)
        nom = re.sub(r'\b(?:SI|NO)\b(?:\s+(?:19|20)\d{4})?', '', nom)
        nom = re.sub(r'\b(?:19|20)\d{2}(?:0[1-9]|1[0-2])\b', '', nom)
        nom = re.sub(r'\s+', ' ', nom).strip(" -|")
        up = nom.upper()
        if not nom or any(h in up for h in [
            "NOMBRE O RAZON", "NOMBRE O RAZÓN", "NO REGISTRA",
            "IDENTIFICACION", "IDENTIFICACIÓN", "APORTANTE", "SIN NOMBRE"
        ]):
            return
        if len(nom) > len(mapa.get(nit, "")):
            mapa[nit] = nom

    upper_full = full_text.upper()
    idx_det = upper_full.find("DETALLE DE PAGOS EFECTUADOS A PARTIR DE 1995")
    if idx_det == -1:
        return mapa
    texto_det = full_text[idx_det:]

    pat_det = re.compile(
        r'(?:^|\n)\s*(\d{5,12})\s+([A-ZÑÁÉÍÓÚÃ0-9\.\,\-\s/&]{3,80}?)\s+(?:SI|NO|S1|N0)\s+(?:19|20)\d{2}(?:0[1-9]|1[0-2])\b'
    )
    for m in pat_det.finditer(texto_det):
        registrar(m.group(1), m.group(2))

    pat_det_alt = re.compile(
        r'(?:^|\n)\s*(\d{5,12})\s+([A-ZÑÁÉÍÓÚÃ0-9\.\,\-\s/&]{3,80}?)\s+(?:19|20)\d{2}(?:0[1-9]|1[0-2])\s+(?:SI|NO)\b'
    )
    for m in pat_det_alt.finditer(texto_det):
        registrar(m.group(1), m.group(2))

    return mapa


# ==========================================
# AUDITORÍA CONTRA TOTALES OFICIALES DEL PDF ([10], [21], [25], [26])
# ==========================================
def extraer_totales_oficiales_pdf(archivo_pdf):
    """
    Extrae los totales de control impresos por Colpensiones al final de la Parte General:
    [10] Cotizadas, [21] Tiempos Públicos Reportados, [25] Simultáneas, [26] Total Semanas.
    """
    totales = {
        "cotizadas_10": None,
        "publicas_21": None,
        "simultaneas_25": None,
        "total_26": None
    }
    try:
        paginas = []
        with pdfplumber.open(archivo_pdf) as pdf:
            for page in pdf.pages:
                txt, _ = _extraer_texto_pagina_con_ocr(page)
                paginas.append(txt)
                if "DETALLE DE PAGOS EFECTUADOS A PARTIR DE 1995" in txt.upper() or "[34]" in txt:
                    break
        full_text = "\n".join(paginas)

        m10 = re.search(r'\[10\]\s*TOTAL\s+SEMANAS\s+COTIZADAS\s*:?[ \t]*\n?[ \t]*([\d\.,]+)', full_text, re.IGNORECASE)
        if m10:
            totales["cotizadas_10"] = _clean_num(m10.group(1))

        m21 = re.search(r'\[21\]\s*TOTAL\s+SEMANAS\s+REPORTADAS\s*:?[ \t]*\n?[ \t]*([\d\.,]+)', full_text, re.IGNORECASE)
        if not m21:
            m21 = re.search(r'\n[ \t]*([\d\.,]+)[ \t]*\n[ \t]*\[21\]\s*TOTAL\s+SEMANAS\s+REPORTADAS', full_text, re.IGNORECASE)
            if m21:
                prev_chunk = full_text[max(0, m21.start() - 60):m21.start()].upper()
                if "NO REGISTRA" in prev_chunk or "[10]" in prev_chunk:
                    m21 = None
        if m21:
            totales["publicas_21"] = _clean_num(m21.group(1))

        m26 = re.search(r'\[26\]\s*TOTAL\s+SEMANAS[^\n]*\)[ \t:]*\n?[ \t]*([\d\.,]+)', full_text, re.IGNORECASE)
        if not m26:
            m26 = re.search(r'\b([\d\.,]+)[ \t]*\n[ \t]*\[26\]\s*TOTAL\s+SEMANAS', full_text, re.IGNORECASE)
        if m26:
            totales["total_26"] = _clean_num(m26.group(1))

        m25 = re.search(r'\[25\]\s*TOTAL\s+SEMANAS\s+SIMULT[ÁA]NEAS\s*:?[ \t]*([\d\.,]+)', full_text, re.IGNORECASE)
        if m25:
            totales["simultaneas_25"] = _clean_num(m25.group(1))
    except Exception:
        pass
    return totales


# ==========================================
# DETECTOR DE EMPLEADOR PÚBLICO Y SEM >= 90%
# ==========================================
def clasificar_empleador_publico(nit_raw, nombre_raw):
    nit_limpio = _limpiar_id_aportante(nit_raw)
    nit_9 = nit_limpio[:9] if len(nit_limpio) >= 9 else nit_limpio

    nombre_norm = _limpiar_texto_colpensiones(nombre_raw).upper()
    for orig, repl in [('Á', 'A'), ('É', 'E'), ('Í', 'I'), ('Ó', 'O'), ('Ú', 'U')]:
        nombre_norm = nombre_norm.replace(orig, repl)

    if nit_9 and nit_9 in ENTIDADES_PUBLICAS_NITS:
        entidad = ENTIDADES_PUBLICAS_NITS[nit_9]
        return True, f"Coincidencia por NIT ({nit_9}): {entidad}"

    for pat in PATRONES_ENTIDAD_PUBLICA:
        match = re.search(pat, nombre_norm)
        if match:
            hallazgo = match.group(0)
            return True, f"Coincidencia por denominación institucional ('{hallazgo}')"

    return False, "Empleador del sector privado o cotizante independiente"


def identificar_ultimo_empleador_publico(df_limpio):
    resultado_base = {
        "es_publico": False,
        "nit": "No identificado",
        "razon_social": "No identificado",
        "ultima_fecha": None,
        "motivo": "Sin registros para evaluar"
    }
    if df_limpio is None or df_limpio.empty:
        return resultado_base

    fecha_max = df_limpio['Hasta'].max()
    ultimos = df_limpio[df_limpio['Hasta'] == fecha_max]

    for _, row in ultimos.iterrows():
        nit_val = str(row.get('NIT_Aportante', '')).strip()
        nom_val = str(row.get('Aportante', '')).strip()
        es_pub, motivo = clasificar_empleador_publico(nit_val, nom_val)
        if es_pub:
            return {
                "es_publico": True,
                "nit": nit_val if nit_val else "No registrado",
                "razon_social": nom_val if nom_val else "Entidad Pública",
                "ultima_fecha": fecha_max.date() if pd.notnull(fecha_max) else None,
                "motivo": motivo
            }

    ultimo = ultimos.iloc[-1]
    return {
        "es_publico": False,
        "nit": str(ultimo.get('NIT_Aportante', 'No registrado')).strip(),
        "razon_social": str(ultimo.get('Aportante', 'No identificado')).strip(),
        "ultima_fecha": fecha_max.date() if pd.notnull(fecha_max) else None,
        "motivo": "Empleador del sector privado o cotizante independiente"
    }


# ==========================================
# EXTRACCIÓN DE DATOS BÁSICOS Y PETICIÓN
# ==========================================
def _es_nombre_persona_valido(candidato):
    c = _limpiar_texto_colpensiones(candidato)
    if not c or len(c) < 5 or len(c) > 80 or re.search(r'\d', c):
        return False

    upper = c.upper()
    prohibidos = [
        "CEDULA", "CÉDULA", "CIUDADANIA", "CIUDADANÍA", "DIRECCION", "DIRECCIÓN",
        "RAZON SOCIAL", "RAZÓN SOCIAL", "APORTANTE", "FECHA", "ESTADO", "ACTIVO",
        "COTIZANTE", "INACTIVO", "CORREO", "UBICACION", "UBICACIÓN", "COLPENSIONES",
        "INFORMACION", "INFORMACIÓN", "AFILIADO", "DOCUMENTO", "REPORTE", "SEMANAS",
        "TIPO DE", "NUMERO DE", "NÚMERO DE", "PERIODO", "ACTUALIZADO", "RESUMEN"
    ]
    if any(p in upper for p in prohibidos):
        return False

    palabras = [p for p in re.split(r'\s+', c) if re.fullmatch(r'[A-Za-zÑñÁÉÍÓÚáéíóúÜü\.]+', p)]
    return len(palabras) >= 2 and len(palabras) == len(c.split())


def extraer_datos_basicos(pdf_file):
    datos = {
        "cedula": "",
        "nombre": "",
        "fecha_nac": date(1975, 1, 1),
        "genero": "Masculino"
    }
    try:
        with pdfplumber.open(pdf_file) as pdf:
            if len(pdf.pages) > 0:
                p0 = pdf.pages[0]
                text_p0, _ = _extraer_texto_pagina_con_ocr(p0)

                try:
                    tablas_p0 = p0.extract_tables() or []
                    for tbl in tablas_p0[:3]:
                        for fila in tbl:
                            celdas_raw = [str(c or "").strip() for c in fila]
                            for idx_c, val in enumerate(celdas_raw):
                                if not val:
                                    continue
                                val_up = val.upper()
                                sig_val = next((celdas_raw[j] for j in range(idx_c + 1, len(celdas_raw)) if celdas_raw[j]), "")

                                if ("NÚMERO DE DOCUMENTO" in val_up or "NUMERO DE DOCUMENTO" in val_up) and sig_val:
                                    m_ced = re.search(r'\b(\d{5,12})\b', sig_val.replace(".", ""))
                                    if m_ced and not datos["cedula"]:
                                        datos["cedula"] = m_ced.group(1)
                                    if "NOMBRE" in val_up and not datos["nombre"]:
                                        for ln in sig_val.splitlines():
                                            if _es_nombre_persona_valido(ln):
                                                datos["nombre"] = _limpiar_texto_colpensiones(ln).upper()
                                                break

                                if (
                                    re.search(r'\bNOMBRES?(?:\s+Y\s+APELLIDOS?)?\s*:?$', val_up)
                                    and "RAZON" not in val_up
                                    and "RAZÓN" not in val_up
                                    and sig_val
                                    and not datos["nombre"]
                                ):
                                    primera_linea = sig_val.splitlines()[0].strip()
                                    if _es_nombre_persona_valido(primera_linea):
                                        datos["nombre"] = _limpiar_texto_colpensiones(primera_linea).upper()

                                if "FECHA DE NACIMIENTO" in val_up and sig_val:
                                    m_f = re.search(r'(\d{2}/\d{2}/\d{4})', sig_val)
                                    if m_f:
                                        datos["fecha_nac"] = datetime.strptime(m_f.group(1), "%d/%m/%Y").date()

                                if ("SEXO" in val_up or "GÉNERO" in val_up or "GENERO" in val_up) and sig_val:
                                    if sig_val.upper().startswith("F"):
                                        datos["genero"] = "Femenino"
                                    elif sig_val.upper().startswith("M"):
                                        datos["genero"] = "Masculino"
                except Exception:
                    pass

                if text_p0:
                    upper_t = text_p0.upper()
                    corte = upper_t.find("RESUMEN DE SEMANAS COTIZADAS")
                    header_txt = text_p0[:corte] if corte != -1 else text_p0[:2000]

                    if not datos["cedula"]:
                        m_ced = re.search(r'(?:N[úu]mero\s+de\s+Documento|C\.C\.)[\s:\|\n]*([\d\.]{5,15})', header_txt, re.IGNORECASE)
                        if m_ced:
                            datos["cedula"] = m_ced.group(1).replace(".", "").strip()
                        else:
                            m_ced2 = re.search(r'C[ée]dula\s+de\s+Ciudadan[íi]a[\s\S]{0,120}?\b(\d{5,12})\b', header_txt, re.IGNORECASE)
                            if m_ced2:
                                datos["cedula"] = m_ced2.group(1)

                    if not datos["nombre"]:
                        m_nom1 = re.search(r'(?:Nombres\s+y\s+Apellidos|Nombre)\s*[:\|]\s*([^\n\|]+)', header_txt, re.IGNORECASE)
                        if m_nom1:
                            bruto = re.split(
                                r'\s+(?:Fecha|Estado|Direcci[óo]n|Correo|Ubicaci[óo]n|Sexo|Tipo|N[úu]mero|Tel[ée]fono)\b',
                                m_nom1.group(1),
                                maxsplit=1,
                                flags=re.IGNORECASE
                            )[0].strip()
                            if _es_nombre_persona_valido(bruto):
                                datos["nombre"] = _limpiar_texto_colpensiones(bruto).upper()

                    if not datos["nombre"] and datos["cedula"]:
                        m_nom2 = re.search(rf'{re.escape(datos["cedula"])}\s*\n\s*([^\n]+)', header_txt)
                        if m_nom2:
                            cand = re.split(
                                r'\s+(?:Fecha|Estado|Direcci[óo]n|Correo|Ubicaci[óo]n)\b',
                                m_nom2.group(1),
                                maxsplit=1,
                                flags=re.IGNORECASE
                            )[0].strip()
                            if _es_nombre_persona_valido(cand):
                                datos["nombre"] = _limpiar_texto_colpensiones(cand).upper()

                    m_fecha = re.search(r'(?:Fecha\s+de\s+Nacimiento|Nacimiento)[\s:\|\n]*(\d{2}/\d{2}/\d{4})', header_txt, re.IGNORECASE)
                    if m_fecha:
                        try:
                            datos["fecha_nac"] = datetime.strptime(m_fecha.group(1), "%d/%m/%Y").date()
                        except Exception:
                            pass
                    else:
                        # En ciertos PDFs nativos las fechas aparecen antes que las etiquetas (1ra fecha = nacimiento)
                        fechas_hdr = re.findall(r'\b(\d{2}/\d{2}/\d{4})\b', header_txt)
                        if fechas_hdr:
                            try:
                                f_cand = datetime.strptime(fechas_hdr[0], "%d/%m/%Y").date()
                                if date(1920, 1, 1) <= f_cand <= date(2005, 12, 31):
                                    datos["fecha_nac"] = f_cand
                            except Exception:
                                pass

                    m_sexo = re.search(r'(?:Sexo|G[ée]nero)[\s:\|]*([MF]|Masculino|Femenino)', header_txt, re.IGNORECASE)
                    if m_sexo:
                        datos["genero"] = "Femenino" if m_sexo.group(1).upper().startswith("F") else "Masculino"

                # Respaldo en encabezado de página 2 (ej. "45427704C BEATRIZ ISABEL SERRANO GULFO")
                if not datos["nombre"] and datos["cedula"] and len(pdf.pages) > 1:
                    text_p1, _ = _extraer_texto_pagina_con_ocr(pdf.pages[1])
                    if text_p1:
                        top_p1 = text_p1[:800]
                        m_misma = re.search(rf'{re.escape(datos["cedula"])}[A-Z]?\s+([^\n]+)', top_p1)
                        m_antes = re.search(rf'([^\n]+)\n\s*[A-Z]?\s*\n?\s*{re.escape(datos["cedula"])}\b', top_p1)
                        m_despues = re.search(rf'{re.escape(datos["cedula"])}\s*\n\s*([^\n]+)', top_p1)
                        for m_cand in (m_misma, m_antes, m_despues):
                            if m_cand and _es_nombre_persona_valido(m_cand.group(1)):
                                datos["nombre"] = _limpiar_texto_colpensiones(m_cand.group(1)).upper()
                                break
    except Exception:
        pass
    return datos


def extraer_datos_peticion(pdf_file, datos_actuales):
    peticiones = "Reconocimiento de Pensión de Vejez."
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            for page in pdf.pages[:3]:
                page_text, _ = _extraer_texto_pagina_con_ocr(page)
                if page_text:
                    text += page_text + "\n"

            if text:
                if not datos_actuales.get("cedula"):
                    m_ced = re.search(r'(?:C\.C\.|cédula.*?No\.|Ciudadanía:\s*)([\d\.]+)', text, re.IGNORECASE)
                    if m_ced:
                        datos_actuales["cedula"] = m_ced.group(1).replace(".", "").strip()

                if not datos_actuales.get("nombre"):
                    m_nom = re.search(r'(?:Yo[,]?\s+|Peticionari[oa]:\s*)([A-ZÑÁÉÍÓÚa-zñáéíóú\s]+)(?:[,]?\s+identificad|Cédula)', text, re.IGNORECASE)
                    if m_nom:
                        datos_actuales["nombre"] = m_nom.group(1).strip().upper()

                m_gen = re.search(r'identificad(a|o)\s+con', text, re.IGNORECASE)
                if m_gen:
                    datos_actuales["genero"] = "Femenino" if m_gen.group(1).lower() == 'a' else "Masculino"

                meses = {
                    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
                    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
                }
                m_nac = re.search(r'nac[íi]\s+el\s+(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(\d{4})', text, re.IGNORECASE)
                if m_nac:
                    dia, mes_str, anio = int(m_nac.group(1)), m_nac.group(2).lower(), int(m_nac.group(3))
                    if mes_str in meses:
                        datos_actuales["fecha_nac"] = date(anio, meses[mes_str], dia)

                match_solicitud = re.search(
                    r'(?i)(?:solicito|peticion(?:es)?|pretension(?:es)?|solicitud(?:es)?)[\s\n:]+(.*?)(?:\n\s*(?:hechos|fundamentos|pruebas|anexos|notificaciones|derecho)|\Z)',
                    text,
                    re.DOTALL
                )
                if match_solicitud:
                    ext = re.sub(r'\n+', ' ', match_solicitud.group(1).strip())
                    if len(ext) > 10:
                        peticiones = ext
    except Exception:
        pass
    return datos_actuales, peticiones


# ==========================================
# EXTRACCIÓN ESTRICTA DE LA PARTE GENERAL (RESUMEN [1-9] Y PÚBLICOS [12-20])
# ==========================================
def extraer_tabla_cruda(archivo_pdf):
    """
    Limita estrictamente la captura a la PARTE GENERAL de la Historia Laboral Unificada:
      - Resumen de Semanas Cotizadas por Empleador ([1] a [9])
      - Resumen de Tiempos Públicos no Cotizados a Colpensiones ([12] a [20])
    Bloquea totalmente la Parte Detallada ([27], [34], [47], Lectura) y evita duplicaciones.
    """
    paginas_texto = []
    with pdfplumber.open(archivo_pdf) as pdf:
        for page in pdf.pages:
            page_text, _ = _extraer_texto_pagina_con_ocr(page)
            paginas_texto.append(page_text)

    full_text = "\n".join(paginas_texto)

    # 1. Delimitar estrictamente el bloque de texto correspondiente a la Parte General
    chunks_general = []
    inicio_encontrado = False

    for p_txt in paginas_texto:
        up = p_txt.upper()

        # Si la página ya es exclusivamente de Parte Detallada o Instructivo, detener lectura de inmediato
        es_solo_detalle = (
            any(m in up for m in [
                "DETALLE DE PAGOS EFECTUADOS ANTERIORES A 1995",
                "DETALLE DE PAGOS EFECTUADOS A PARTIR DE 1995",
                "DETALLE DE PERIODOS REPORTADOS POR ENTIDADES",
                "LECTURA DEL REPORTE DE LA HISTORIA LABORAL"
            ])
            and "RESUMEN DE SEMANAS COTIZADAS POR EMPLEADOR" not in up
            and "RESUMEN DE TIEMPOS PÚBLICOS NO COTIZADOS" not in up
            and "RESUMEN DE TIEMPOS PUBLICOS NO COTIZADOS" not in up
        )
        if es_solo_detalle:
            break

        bloque = p_txt
        if not inicio_encontrado:
            pos_ini = up.find("RESUMEN DE SEMANAS COTIZADAS POR EMPLEADOR")
            if pos_ini == -1:
                pos_ini = up.find("[1]IDENTIFICACI")
            if pos_ini == -1:
                pos_ini = up.find("IDENTIFICACIÓN APORTANTE")
            if pos_ini != -1:
                bloque = p_txt[pos_ini:]
                inicio_encontrado = True
            else:
                continue

        # Cortar en cuanto aparezca cualquier marcador de fin de Parte General dentro de la página
        up_b = bloque.upper()
        corte = -1
        for mf in MARCADORES_FIN_PARTE_GENERAL:
            pos = up_b.find(mf)
            if pos != -1 and (corte == -1 or pos < corte):
                corte = pos

        if corte != -1:
            chunks_general.append(bloque[:corte])
            break
        else:
            chunks_general.append(bloque)

    texto_general = "\n".join(chunks_general)
    if not texto_general.strip():
        return pd.DataFrame()

    # 2. Parsear líneas de la Parte General con reconstrucción multilínea y validación de 9 columnas
    filas_validas = []
    vistos = set()
    linea_pendiente = ""
    seccion_actual = "Cotizadas"

    for raw_ln in texto_general.splitlines():
        s = raw_ln.replace("|", " ").strip()
        if not s:
            continue
        up = s.upper()

        if "RESUMEN DE TIEMPOS PÚBLICOS NO COTIZADOS" in up or "RESUMEN DE TIEMPOS PUBLICOS NO COTIZADOS" in up:
            seccion_actual = "Publicas"
            linea_pendiente = ""
            continue

        # Saltar encabezados, totales y pies de página
        if any(skip in up for skip in [
            "[1]IDENTIFICAC", "[12]IDENTIFICAC", "IMPRESO POR INTERNET",
            "PERIODO DE INFORME", "ACTUALIZADO A:", "COLPENSIONES NIT",
            "REPORTE DE SEMANAS COTIZADAS", "[10] TOTAL SEMANAS",
            "[11] SEMANAS COTIZADAS", "[21]TOTAL SEMANAS", "NO REGISTRA INFORMACIÓN",
            "NO REGISTRA INFORMACION", "EL SIGUIENTE RESUMEN", "APORTANTE [2]"
        ]):
            linea_pendiente = ""
            continue

        fechas = re.findall(r'\d{2}/\d{2}/\d{4}', s)
        if len(fechas) >= 2:
            # Reconstruir fila si el nombre del empleador se partió en dos líneas (ej. Tiempos Públicos)
            if not re.match(r'^\d{5,15}\b', s) and linea_pendiente:
                s = f"{linea_pendiente} {s}"
            linea_pendiente = ""

            try:
                p_antes, resto_1 = s.split(fechas[0], 1)
                _, p_despues = resto_1.split(fechas[1], 1)
                nit, nom = _separar_id_y_nombre(p_antes)

                # Unir "$ 497.000" -> "$497.000" sin alterar "4.245.000$ 4,29"
                p_despues = re.sub(r'(?:^|(?<=\s))\$\s+([\d\.\,]+)', r'$\1', p_despues.strip())
                tokens = [t for t in re.split(r'\s+', p_despues) if t]

                # Exigir los 5 campos exactos de la Parte General tras Hasta: Salario, Semanas, Lic, Sim, Total
                if len(tokens) < 5:
                    continue

                fila_ok = _validar_fila_resumen(
                    nit, nom, fechas[0], fechas[1],
                    tokens[0], tokens[1], tokens[2], tokens[3], tokens[4]
                )
                if not fila_ok:
                    continue

                # Desduplicación exacta de registro
                clave = (fila_ok[0], fila_ok[2], fila_ok[3], fila_ok[4], fila_ok[8], seccion_actual)
                if clave in vistos:
                    continue
                vistos.add(clave)
                filas_validas.append(fila_ok)
            except Exception:
                continue
        else:
            # Posible primera línea de una razón social multilínea ("890200500 EMPRESA SOCIAL DEL ESTADO...")
            if re.match(r'^\d{5,15}\s+[A-ZÑÁÉÍÓÚ]', s):
                linea_pendiente = s
            else:
                linea_pendiente = ""

    if not filas_validas:
        return pd.DataFrame()

    # 3. Completar razones sociales truncadas usando únicamente el catálogo de nombres del Detalle Post-1995
    catalogo_nombres = _extraer_catalogo_nombres_detalle(full_text)
    for fila in filas_validas:
        nit_clean = fila[0]
        nom_actual = fila[1]
        if nit_clean and nit_clean != "(Sin ID)":
            nom_det = catalogo_nombres.get(nit_clean, "")
            if len(nom_det) > len(nom_actual):
                fila[1] = nom_det

    return pd.DataFrame(filas_validas, columns=COLUMNAS_RESUMEN)


# ==========================================
# SANEAMIENTO Y CONSOLIDACIÓN (SIN INFLAR LICENCIAS NI BORRAR MESES PARTIDOS)
# ==========================================
def limpiar_y_estandarizar(df_crudo, col_desde, col_hasta, col_ibc, col_semanas, col_nit=None, col_aportante=None):
    datos = []
    tiene_col_sim = "Sim" in df_crudo.columns
    tiene_col_lic = "Lic" in df_crudo.columns
    usa_columna_total = (str(col_semanas).strip().lower() == "total")

    for _, row in df_crudo.iterrows():
        try:
            raw_desde = str(row[col_desde])
            raw_hasta = str(row[col_hasta])
            raw_ibc = str(row[col_ibc])
            raw_semanas = str(row[col_semanas])
            raw_nit = _limpiar_id_aportante(row[col_nit]) if col_nit and col_nit in row else ""
            raw_aportante = _limpiar_texto_colpensiones(row[col_aportante]) if col_aportante and col_aportante in row else "No especificado"

            match_d = re.search(r'\d{2}/\d{2}/\d{4}', raw_desde)
            match_h = re.search(r'\d{2}/\d{2}/\d{4}', raw_hasta)
            if not match_d or not match_h:
                continue

            desde = pd.to_datetime(match_d.group(0), dayfirst=True, errors='coerce')
            hasta = pd.to_datetime(match_h.group(0), dayfirst=True, errors='coerce')
            if pd.isna(desde) or pd.isna(hasta) or hasta < desde:
                continue

            ibc = _clean_num(raw_ibc)
            semanas_leidas = _clean_num(raw_semanas)
            sim_leida = _clean_num(row["Sim"]) if tiene_col_sim else 0.0
            lic_leida = _clean_num(row["Lic"]) if tiene_col_lic else 0.0
            max_fisico = ((hasta - desde).days + 1) / 7.0 + 0.5

            semanas_final = semanas_leidas
            es_simultaneo_cero = (semanas_leidas == 0.0 and sim_leida > 0.0)

            # Solo corregir si excede el máximo físico del periodo (nunca sobrescribir periodos multianuales válidos)
            if semanas_final > max_fisico:
                semanas_final = max(0.0, round(max_fisico - 0.5 - lic_leida - sim_leida, 2))
            elif not es_simultaneo_cero and semanas_final <= 0.0:
                semanas_final = max(0.0, round(((hasta - desde).days + 1) / 7.0 - lic_leida - sim_leida, 2))

            if semanas_final > 0 or (es_simultaneo_cero and ibc > 0):
                datos.append({
                    "NIT_Aportante": raw_nit if raw_nit else "(Sin ID)",
                    "Aportante": raw_aportante if raw_aportante else "No especificado",
                    "Desde": desde,
                    "Hasta": hasta,
                    "IBC": ibc,
                    "Semanas": semanas_final,
                    "_usa_total": usa_columna_total
                })
        except Exception:
            continue

    df = pd.DataFrame(datos)
    return df.sort_values('Desde') if not df.empty else df


def _unir_unicos(serie, vacio="(Sin ID)"):
    vals = []
    for v in serie:
        s = str(v).strip()
        if s and s not in ("(Sin ID)", "No especificado", "nan", "None") and s not in vals:
            vals.append(s)
    return " / ".join(vals) if vals else vacio


def aplicar_regla_simultaneidad(df):
    """
    Agrupa por periodo mensual de inicio sumando los IBC concurrentes.
    Si se usa la columna oficial [9] Total (donde Colpensiones ya descontó [7] Lic y [8] Sim),
    se suman las fracciones del mes (ej. 10 días + 18 días = 4,00 semanas) sin duplicar simultáneos.
    """
    if df.empty:
        return df
    df_tmp = df.copy()
    usa_total = bool(df_tmp['_usa_total'].all()) if '_usa_total' in df_tmp.columns else True
    df_tmp['Periodo'] = df_tmp['Desde'].dt.to_period('M')

    agg_dict = {
        'NIT_Aportante': lambda s: _unir_unicos(s, "(Sin ID)"),
        'Aportante': lambda s: _unir_unicos(s, "No especificado"),
        'Desde': 'min',
        'Hasta': 'max',
        'IBC': 'sum',
        'Semanas': 'sum' if usa_total else 'max'
    }
    res = (
        df_tmp.groupby('Periodo')
        .agg(agg_dict)
        .reset_index()
        .sort_values('Periodo')
    )
    return res

