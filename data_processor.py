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
    r"\bHOSPITAL\s+(?:DEPARTAMENTAL|MUNICIPAL|UNIVERSITARIO|REGIONAL|LOCAL)\b",
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


# ==========================================
# UTILIDADES DE LECTURA Y LIMPIEZA COLPENSIONES
# ==========================================
def _limpiar_texto_colpensiones(texto):
    """Normaliza espacios y corrige artefactos típicos de codificación en reportes Colpensiones."""
    if not texto:
        return ""
    t = str(texto)
    t = t.replace("Ã‘", "Ñ").replace("Ã±", "ñ")
    t = re.sub(r'([A-Z])ÃO\b', r'\1ÑO', t)
    t = re.sub(r'([A-Z])Ã([AEIOU])', r'\1Ñ\2', t)
    return re.sub(r'\s+', ' ', t).strip()


def _limpiar_id_aportante(val):
    """Extrae únicamente dígitos del NIT/Cédula del aportante."""
    return re.sub(r'\D', '', str(val or ''))


def _extraer_texto_pagina_con_ocr(page):
    """Extrae texto nativo de una página PDF; si viene como imagen, aplica OCR automáticamente."""
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


def _extraer_catalogo_nombres_detalle(full_text, tablas_detalle=None):
    """
    Construye un mapa {ID_Aportante: Nombre_Completo} leyendo el Detalle Post-1995 ([34]/[35])
    y Tiempos Públicos ([47]/[48]) para completar los nombres truncados a 20 chars del Resumen.
    """
    mapa = {}

    def registrar(nit_raw, nom_raw):
        nit = _limpiar_id_aportante(nit_raw)
        if not (5 <= len(nit) <= 12):
            return
        nom = _limpiar_texto_colpensiones(nom_raw)
        # Eliminar restos de columnas contiguas ([36] RA: SI/NO, [37] Periodo: YYYYMM)
        nom = re.sub(r'\b(?:SI|NO)\b(?:\s+(?:19|20)\d{4})?', '', nom)
        nom = re.sub(r'\b(?:19|20)\d{2}(?:0[1-9]|1[0-2])\b', '', nom)
        nom = re.sub(r'\s+', ' ', nom).strip(" -|")
        upper_nom = nom.upper()
        if not nom or any(h in upper_nom for h in [
            "NOMBRE O RAZON", "NOMBRE O RAZÓN", "NO REGISTRA",
            "IDENTIFICACION", "IDENTIFICACIÓN", "APORTANTE"
        ]):
            return
        if len(nom) > len(mapa.get(nit, "")):
            mapa[nit] = nom

    # 1. Desde tablas nativas del Detalle
    if tablas_detalle:
        for tabla in tablas_detalle:
            for fila in tabla:
                if not fila or len(fila) < 2:
                    continue
                c0 = str(fila[0] or "").strip()
                c1 = str(fila[1] or "").strip()
                nits = [x.strip() for x in c0.splitlines() if re.fullmatch(r'[\d\.\-]{5,15}', x.strip())]
                if len(nits) == 1:
                    registrar(nits[0], c1.replace('\n', ' '))
                elif len(nits) > 1 and len(set(nits)) == 1:
                    lineas_nom = [ln.strip() for ln in c1.splitlines() if ln.strip()]
                    if len(lineas_nom) == len(nits):
                        registrar(nits[0], lineas_nom[0])
                    elif len(lineas_nom) > 0 and len(lineas_nom) % len(nits) == 0:
                        paso = len(lineas_nom) // len(nits)
                        registrar(nits[0], " ".join(lineas_nom[:paso]))

    # 2. Respaldo desde texto plano / OCR del Detalle Post-1995
    upper_full = full_text.upper()
    idx_det = upper_full.find("DETALLE DE PAGOS EFECTUADOS A PARTIR DE 1995")
    texto_det = full_text[idx_det:] if idx_det != -1 else full_text

    # En texto plano/OCR una fila del detalle puede ocupar 1 o 2 líneas antes de SI/NO + YYYYMM
    pat_det = re.compile(
        r'(?:^|\n)\s*(\d{5,12})\s+([A-ZÑÁÉÍÓÚÃ0-9\.\,\-\s/&]{3,80}?)\s+(?:SI|NO|S1|N0)\s+(?:19|20)\d{2}(?:0[1-9]|1[0-2])\b'
    )
    for m in pat_det.finditer(texto_det):
        registrar(m.group(1), m.group(2))

    # Variación donde el orden en OCR invierte Periodo y SI/NO (ej. "... 202206 SI")
    pat_det_alt = re.compile(
        r'(?:^|\n)\s*(\d{5,12})\s+([A-ZÑÁÉÍÓÚÃ0-9\.\,\-\s/&]{3,80}?)\s+(?:19|20)\d{2}(?:0[1-9]|1[0-2])\s+(?:SI|NO)\b'
    )
    for m in pat_det_alt.finditer(texto_det):
        registrar(m.group(1), m.group(2))

    return mapa


def _desdoblar_fila_tabla_resumen(fila):
    """
    Normaliza una fila extraída por pdfplumber.extract_tables() del Resumen Colpensiones.
    Soporta columnas vacías intermedias (como ocurre entre [5] Salario y [6] Semanas en pág. 1)
    y celdas multilínea donde dos periodos vienen agrupados con '\n'.
    """
    celdas = [str(c).strip() if c is not None else "" for c in fila]
    # Eliminar columnas completamente vacías
    celdas_no_vacias = [c for c in celdas if c != ""]
    if len(celdas_no_vacias) < 6:
        return []

    # Ubicar las dos celdas que contienen fechas DD/MM/YYYY
    regex_fecha = re.compile(r'\d{2}/\d{2}/\d{4}')
    idx_fechas = [i for i, c in enumerate(celdas_no_vacias) if regex_fecha.search(c)]
    if len(idx_fechas) < 2:
        return []

    i_desde, i_hasta = idx_fechas[0], idx_fechas[1]
    if i_desde < 1:
        return []

    # Desdoblar si hay múltiples fechas en la misma celda (filas pegadas por '\n')
    fechas_desde = regex_fecha.findall(celdas_no_vacias[i_desde])
    fechas_hasta = regex_fecha.findall(celdas_no_vacias[i_hasta])
    n_sub = min(len(fechas_desde), len(fechas_hasta))
    if n_sub == 0:
        return []

    def dividir_campo(texto, n):
        partes = [p.strip() for p in texto.splitlines() if p.strip()]
        if len(partes) == n:
            return partes
        if n == 1:
            return [" ".join(partes)]
        # Rellenar con el último valor si había un solo elemento para varias subfilas
        if len(partes) == 1:
            return partes * n
        return (partes + [partes[-1]] * n)[:n]

    if i_desde == 1:
        # Caso donde ID y Nombre quedaron fusionados en la celda 0
        bloques_0 = dividir_campo(celdas_no_vacias[0], n_sub)
        ids_list, noms_list = [], []
        for b in bloques_0:
            m_sep = re.match(r'^([\d\.\-]{5,15})\s+(.*)$', b)
            if m_sep:
                ids_list.append(m_sep.group(1))
                noms_list.append(m_sep.group(2))
            else:
                ids_list.append("")
                noms_list.append(b)
    else:
        ids_list = dividir_campo(celdas_no_vacias[0], n_sub)
        noms_list = dividir_campo(" ".join(celdas_no_vacias[1:i_desde]), n_sub)

    resto = celdas_no_vacias[i_hasta + 1:]
    # Se esperan 5 campos después de Hasta: [5] Salario, [6] Semanas, [7] Lic, [8] Sim, [9] Total
    salarios = dividir_campo(resto[0], n_sub) if len(resto) > 0 else ["0"] * n_sub
    semanas = dividir_campo(resto[1], n_sub) if len(resto) > 1 else ["0"] * n_sub
    lics = dividir_campo(resto[2], n_sub) if len(resto) > 2 else ["0"] * n_sub
    sims = dividir_campo(resto[3], n_sub) if len(resto) > 3 else ["0"] * n_sub
    totales = dividir_campo(resto[4], n_sub) if len(resto) > 4 else semanas

    resultados = []
    for k in range(n_sub):
        nit_k = _limpiar_id_aportante(ids_list[k])
        nom_k = _limpiar_texto_colpensiones(noms_list[k])
        resultados.append([
            nit_k if nit_k else "(Sin ID)",
            nom_k,
            fechas_desde[k],
            fechas_hasta[k],
            salarios[k],
            semanas[k],
            lics[k],
            sims[k],
            totales[k]
        ])
    return resultados


def _parsear_linea_resumen_texto(linea):
    """
    Parsea una línea de texto/OCR del formato Colpensiones Resumen:
    [1] ID  [2] Nombre o Razón Social  [3] Desde  [4] Hasta  [5] Salario  [6] Semanas  [7] Lic  [8] Sim  [9] Total
    """
    linea = linea.replace("|", " ").strip()
    if not linea:
        return None

    # Descartar encabezados o pies de página
    upper = linea.upper()
    if any(skip in upper for skip in [
        "IDENTIFICACION APORTANTE", "IDENTIFICACIÓN APORTANTE",
        "IMPRESO POR INTERNET", "PERIODO DE INFORME", "ACTUALIZADO A:"
    ]):
        return None

    pat_completo = re.compile(
        r'^\s*(\d[\d\.\-]{4,14})\s+'                    # [1] Identificación Aportante
        r'(.+?)\s+'                                      # [2] Nombre o Razón Social
        r'(\d{2}/\d{2}/\d{4})\s+'                        # [3] Desde
        r'(\d{2}/\d{2}/\d{4})\s+'                        # [4] Hasta
        r'(\$?\s*[\d\.\,]+\s*\$?)\s+'                    # [5] Último Salario ($ antes o después)
        r'([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s*$'  # [6]Sem [7]Lic [8]Sim [9]Total
    )
    m = pat_completo.match(linea)
    if m:
        g = list(m.groups())
        g[0] = _limpiar_id_aportante(g[0])
        g[1] = _limpiar_texto_colpensiones(g[1])
        return g

    # Respaldo flexible si el OCR omitió alguna columna numérica final o el signo $
    fechas = re.findall(r'\d{2}/\d{2}/\d{4}', linea)
    if len(fechas) >= 2:
        try:
            p_antes, resto_1 = linea.split(fechas[0], 1)
            _, p_despues = resto_1.split(fechas[1], 1)
            p_antes = p_antes.strip()
            m_id = re.match(r'^([\d\.\-]{5,15})\s+(.*)$', p_antes)
            if m_id:
                nit = _limpiar_id_aportante(m_id.group(1))
                nom = _limpiar_texto_colpensiones(m_id.group(2))
            else:
                nit = "(Sin ID)"
                nom = _limpiar_texto_colpensiones(p_antes)

            # Unir "$ 497.000" -> "$497.000" antes de dividir por espacios
            p_despues = re.sub(r'\$\s+([\d\.\,]+)', r'$\1', p_despues.strip())
            tokens = [t for t in re.split(r'\s+', p_despues) if t]
            if len(tokens) >= 2:
                salario = tokens[0]
                sem = tokens[1] if len(tokens) > 1 else "0"
                lic = tokens[2] if len(tokens) > 2 else "0"
                sim = tokens[3] if len(tokens) > 3 else "0"
                tot = tokens[4] if len(tokens) > 4 else sem
                return [nit, nom, fechas[0], fechas[1], salario, sem, lic, sim, tot]
        except Exception:
            return None

    return None


# ==========================================
# DETECTOR DE EMPLEADOR PÚBLICO Y SEM >= 90%
# ==========================================
def clasificar_empleador_publico(nit_raw, nombre_raw):
    """Evalúa si un empleador pertenece al sector público o SEM >= 90% estatal."""
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
    """Identifica el aportante más reciente en la historia laboral y verifica si es público."""
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

    # Si hubo simultaneidad en el último periodo, priorizar si alguno es entidad pública
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
def extraer_datos_basicos(pdf_file):
    """
    Extrae cédula, nombre, fecha de nacimiento y sexo (si existe)
    del encabezado del Reporte de Semanas Cotizadas de Colpensiones (nativo u OCR).
    """
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
                text, _ = _extraer_texto_pagina_con_ocr(p0)

                # 1. Intentar desde tabla de cabecera si es PDF nativo
                try:
                    tablas_p0 = p0.extract_tables() or []
                    for tbl in tablas_p0[:2]:
                        for fila in tbl:
                            fila_txt = " ".join([str(c or "") for c in fila])
                            if "Número de Documento" in fila_txt or "Fecha de Nacimiento" in fila_txt:
                                celdas = [str(c or "").strip() for c in fila if c and str(c).strip()]
                                for idx_c, val in enumerate(celdas):
                                    if "Número de Documento" in val and idx_c + 1 < len(celdas):
                                        m_ced = re.search(r'\b(\d{6,12})\b', celdas[idx_c + 1])
                                        if m_ced:
                                            datos["cedula"] = m_ced.group(1)
                                        lineas_der = [ln.strip() for ln in celdas[idx_c + 1].splitlines() if ln.strip()]
                                        for ln in lineas_der:
                                            if re.fullmatch(r'[A-ZÑÁÉÍÓÚ\s]{5,}', ln) and "CÉDULA" not in ln:
                                                datos["nombre"] = _limpiar_texto_colpensiones(ln)
                                                break
                                    if "Fecha de Nacimiento" in val and idx_c + 1 < len(celdas):
                                        m_f = re.search(r'(\d{2}/\d{2}/\d{4})', celdas[idx_c + 1])
                                        if m_f:
                                            datos["fecha_nac"] = datetime.strptime(m_f.group(1), "%d/%m/%Y").date()
                except Exception:
                    pass

                # 2. Respaldo / complemento sobre texto plano o OCR
                if text:
                    if not datos["cedula"]:
                        m_ced = re.search(r'(?:Número de Documento|Documento|C\.C\.)[\s:\n|]*([\d\.]{6,15})', text, re.IGNORECASE)
                        if m_ced:
                            datos["cedula"] = m_ced.group(1).replace(".", "").strip()
                        else:
                            # En el formato Colpensiones, el número aparece justo debajo de "Cédula de Ciudadanía"
                            m_ced2 = re.search(r'C[ée]dula\s+de\s+Ciudadan[íi]a[\s\S]{0,80}?\b(\d{6,12})\b', text, re.IGNORECASE)
                            if m_ced2:
                                datos["cedula"] = m_ced2.group(1)

                    if not datos["nombre"]:
                        m_nom = re.search(r'(?:Nombres y Apellidos|Nombre)[\s:\|]+([A-ZÑÁÉÍÓÚ\s]{5,60})(?=\n|Dirección|Fecha|Sexo|Tipo)', text, re.IGNORECASE)
                        if m_nom:
                            cand = _limpiar_texto_colpensiones(m_nom.group(1))
                            if not any(x in cand.upper() for x in ["DIRECCION", "RAZON SOCIAL", "APORTANTE"]):
                                datos["nombre"] = cand
                        if not datos["nombre"] and datos["cedula"]:
                            # En el formato Colpensiones el nombre completo figura justo después del número de cédula
                            m_nom2 = re.search(rf'{re.escape(datos["cedula"])}\s*\n\s*([A-ZÑÁÉÍÓÚ]{{2,}}(?:\s+[A-ZÑÁÉÍÓÚ]{{2,}}){{1,5}})', text)
                            if m_nom2:
                                datos["nombre"] = _limpiar_texto_colpensiones(m_nom2.group(1))

                    m_fecha = re.search(r'(?:Fecha de Nacimiento|Nacimiento)[\s:\n|]*(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
                    if m_fecha:
                        try:
                            datos["fecha_nac"] = datetime.strptime(m_fecha.group(1), "%d/%m/%Y").date()
                        except Exception:
                            pass

                    m_sexo = re.search(r'(?:Sexo)[\s:]*([MF])', text, re.IGNORECASE)
                    if m_sexo:
                        datos["genero"] = "Femenino" if m_sexo.group(1).upper() == 'F' else "Masculino"
    except Exception:
        pass
    return datos


def extraer_datos_peticion(pdf_file, datos_actuales):
    """Rescata datos básicos faltantes desde la petición y extrae el bloque de pretensiones."""
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
# EXTRACCIÓN DE TABLA RESUMEN + ENRIQUECIMIENTO DE NOMBRES
# ==========================================
def extraer_tabla_cruda(archivo_pdf):
    """
    Extrae el Resumen de Semanas Cotizadas por Empleador (y Tiempos Públicos si existen),
    separando limpiamente [1] Identificación Aportante y [2] Nombre o Razón Social,
    y completando los nombres truncados usando [34]/[35] del Detalle Post-1995.
    """
    filas_resumen = []
    tablas_detalle = []
    full_text_paginas = []
    en_resumen = True

    with pdfplumber.open(archivo_pdf) as pdf:
        for page in pdf.pages:
            page_text, uso_ocr = _extraer_texto_pagina_con_ocr(page)
            full_text_paginas.append(page_text)
            upper_page = page_text.upper()

            es_pagina_detalle = (
                "DETALLE DE PAGOS EFECTUADOS A PARTIR DE 1995" in upper_page
                or "[34]" in page_text
                or "LECTURA DEL REPORTE DE LA HISTORIA LABORAL" in upper_page
            )
            es_pagina_resumen = (
                "RESUMEN DE SEMANAS COTIZADAS POR EMPLEADOR" in upper_page
                or "[1]IDENTIFICACION" in upper_page
                or "[1]IDENTIFICACIÓN" in upper_page
                or "RESUMEN DE TIEMPOS PÚBLICOS NO COTIZADOS" in upper_page
                or "RESUMEN DE TIEMPOS PUBLICOS NO COTIZADOS" in upper_page
            )

            if es_pagina_detalle and not es_pagina_resumen:
                en_resumen = False

            # 1. Intentar extracción por tablas nativas (solo si no requirió OCR)
            if not uso_ocr:
                tablas = page.extract_tables() or []
                for tbl in tablas:
                    if not tbl:
                        continue
                    encabezado_tbl = " ".join([str(c or "") for c in tbl[0]]).upper()
                    if "[34]" in encabezado_tbl or "[47]" in encabezado_tbl or "REFERENCIA DE" in encabezado_tbl:
                        tablas_detalle.append(tbl)
                        continue

                    if en_resumen or es_pagina_resumen:
                        for fila in tbl:
                            subfilas = _desdoblar_fila_tabla_resumen(fila)
                            filas_resumen.extend(subfilas)

    full_text = "\n".join(full_text_paginas)

    # 2. Respaldo por texto estructurado / OCR si no se extrajeron filas por tablas nativas
    if not filas_resumen:
        upper_full = full_text.upper()
        idx_ini = upper_full.find("RESUMEN DE SEMANAS COTIZADAS POR EMPLEADOR")
        if idx_ini == -1:
            idx_ini = upper_full.find("IDENTIFICACIÓN APORTANTE")
        if idx_ini == -1:
            idx_ini = upper_full.find("IDENTIFICACION APORTANTE")

        # Detener antes del bloque de Detalle o de Simultáneo/Lectura
        marcadores_fin = [
            "RESUMEN TIEMPO PÚBLICO SIMULTÁNEO",
            "RESUMEN TIEMPO PUBLICO SIMULTANEO",
            "DETALLE DE PAGOS EFECTUADOS ANTERIORES",
            "DETALLE DE PAGOS EFECTUADOS A PARTIR DE 1995",
            "LECTURA DEL REPORTE DE LA HISTORIA LABORAL"
        ]
        idx_fin = -1
        for mf in marcadores_fin:
            pos = upper_full.find(mf, max(0, idx_ini))
            if pos != -1 and (idx_fin == -1 or pos < idx_fin):
                idx_fin = pos

        if idx_ini != -1:
            texto_resumen = full_text[idx_ini:idx_fin] if idx_fin != -1 else full_text[idx_ini:]
        else:
            texto_resumen = full_text

        for linea in texto_resumen.splitlines():
            parsed = _parsear_linea_resumen_texto(linea)
            if parsed:
                filas_resumen.append(parsed)

    if not filas_resumen:
        return pd.DataFrame()

    # 3. Enriquecer nombres truncados del Resumen con los nombres completos del Detalle ([34]/[35])
    catalogo_nombres = _extraer_catalogo_nombres_detalle(full_text, tablas_detalle=tablas_detalle)

    for fila in filas_resumen:
        nit_clean = _limpiar_id_aportante(fila[0])
        nom_actual = _limpiar_texto_colpensiones(fila[1])
        if nit_clean:
            fila[0] = nit_clean
            nom_det = catalogo_nombres.get(nit_clean, "")
            if len(nom_det) > len(nom_actual):
                fila[1] = nom_det
            else:
                fila[1] = nom_actual

    df_out = pd.DataFrame(filas_resumen, columns=COLUMNAS_RESUMEN)
    return df_out


# ==========================================
# SANEAMIENTO Y SIMULTANEIDAD
# ==========================================
def limpiar_y_estandarizar(df_crudo, col_desde, col_hasta, col_ibc, col_semanas, col_nit=None, col_aportante=None):
    datos = []
    tiene_col_sim = "Sim" in df_crudo.columns

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
            if pd.isna(desde) or pd.isna(hasta):
                continue

            def clean_num(val):
                if not val or str(val).lower() == 'none':
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

            ibc = clean_num(raw_ibc)
            semanas_leidas = clean_num(raw_semanas)
            sim_leida = clean_num(row["Sim"]) if tiene_col_sim else 0.0
            semanas_final = semanas_leidas

            # Solo recalcular si no es un periodo totalmente simultáneo (Total = 0 con Sim > 0)
            es_simultaneo_cero = (semanas_leidas == 0.0 and sim_leida > 0.0)
            if not es_simultaneo_cero and (semanas_leidas <= 0.1 or semanas_leidas > 60):
                dias_calculados = (hasta - desde).days + 1
                semanas_final = dias_calculados / 7.0 if 0 < dias_calculados < 12000 else 0.0

            if semanas_final > 0 or (es_simultaneo_cero and ibc > 0):
                datos.append({
                    "NIT_Aportante": raw_nit if raw_nit else "(Sin ID)",
                    "Aportante": raw_aportante if raw_aportante else "No especificado",
                    "Desde": desde,
                    "Hasta": hasta,
                    "IBC": ibc,
                    "Semanas": semanas_final
                })
        except Exception:
            continue

    df = pd.DataFrame(datos)
    return df.sort_values('Desde') if not df.empty else df


def aplicar_regla_simultaneidad(df):
    if df.empty:
        return df
    df_tmp = df.copy()
    df_tmp['Periodo'] = df_tmp['Desde'].dt.to_period('M')
    return (
        df_tmp.groupby('Periodo')
        .agg({'IBC': 'sum', 'Semanas': 'max', 'Desde': 'min', 'Hasta': 'max'})
        .reset_index()
        .sort_values('Periodo')
    )
