import pdfplumber
import pandas as pd
import re
from datetime import datetime, date

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
    r"\bINPEC\b", r"\bCREMIL\b", r"\bCASUR\b", r"\bCOLPENSIONES\b", r"\bUGPP\b",
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


def clasificar_empleador_publico(nit_raw, nombre_raw):
    """Evalúa si un empleador pertenece al sector público o SEM >= 90% estatal."""
    nit_limpio = re.sub(r'\D', '', str(nit_raw or ''))
    nit_9 = nit_limpio[:9] if len(nit_limpio) >= 9 else nit_limpio

    nombre_norm = str(nombre_raw or '').upper().strip()
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

    # Si hubo simultaneidad en el último periodo, priorizar si alguno de los empleadores es público
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


def extraer_datos_basicos(pdf_file):
    """Extrae cédula, nombre, fecha de nacimiento y sexo del encabezado de la Historia Laboral."""
    datos = {
        "cedula": "",
        "nombre": "",
        "fecha_nac": date(1975, 1, 1),
        "genero": "Masculino"
    }
    try:
        with pdfplumber.open(pdf_file) as pdf:
            if len(pdf.pages) > 0:
                text = pdf.pages[0].extract_text()
                if text:
                    match_cedula = re.search(r'(?:Número de Documento|Documento|C\.C\.)[\s:]*([\d\.]+)', text, re.IGNORECASE)
                    match_nombre = re.search(r'(?:Nombres y Apellidos|Nombre)[\s:]*([A-ZÑÁÉÍÓÚ\s]+)(?=\n|Fecha|Sexo|Tipo)', text, re.IGNORECASE)
                    match_fecha = re.search(r'(?:Fecha de Nacimiento|Nacimiento)[\s:]*(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
                    match_sexo = re.search(r'(?:Sexo)[\s:]*([MF])', text, re.IGNORECASE)

                    if match_cedula:
                        datos["cedula"] = match_cedula.group(1).replace(".", "").strip()
                    if match_nombre:
                        datos["nombre"] = re.sub(r'\s+', ' ', match_nombre.group(1)).strip()
                    if match_fecha:
                        try:
                            datos["fecha_nac"] = datetime.strptime(match_fecha.group(1), "%d/%m/%Y").date()
                        except Exception:
                            pass
                    if match_sexo:
                        datos["genero"] = "Femenino" if match_sexo.group(1).upper() == 'F' else "Masculino"
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
                page_text = page.extract_text()
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


def extraer_tabla_cruda(archivo_pdf):
    filas_crudas = []
    with pdfplumber.open(archivo_pdf) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text() or ""
            full_text += "\n" + text

    marcador_inicio = "RESUMEN DE SEMANAS COTIZADAS POR EMPLEADOR"
    marcador_fin = "DETALLE DE PAGOS EFECTUADOS ANTERIORES"
    idx_inicio = full_text.find(marcador_inicio)
    idx_fin = full_text.find(marcador_fin)
    texto_a_procesar = full_text

    if idx_inicio != -1:
        if idx_fin != -1:
            texto_a_procesar = full_text[idx_inicio:idx_fin]
        else:
            texto_a_procesar = full_text[idx_inicio:]
    else:
        idx_alt = full_text.find("Identificación Aportante")
        if idx_alt != -1:
            texto_a_procesar = full_text[idx_alt:]

    lineas = texto_a_procesar.split('\n')
    regex_fecha = re.compile(r'\d{2}/\d{2}/\d{4}')

    for linea in lineas:
        linea = linea.strip()
        if not linea or not regex_fecha.search(linea):
            continue

        if '","' in linea:
            token_sep = "||SEP||"
            linea_temp = linea.replace('","', token_sep).strip('"')
            partes = [p.strip() for p in linea_temp.split(token_sep)]
            filas_crudas.append(partes)
        else:
            fechas = regex_fecha.findall(linea)
            if len(fechas) >= 2:
                try:
                    split_1 = linea.split(fechas[0], 1)
                    encabezado_fila = split_1[0].strip()
                    # Separar NIT inicial si viene pegado al nombre del aportante
                    m_nit = re.match(r'^([\d\.\-]{6,15})\s+(.*)$', encabezado_fila)
                    if m_nit:
                        id_ap = m_nit.group(1).strip()
                        nom_ap = m_nit.group(2).strip()
                    else:
                        id_ap = "(Sin ID)"
                        nom_ap = encabezado_fila

                    split_2 = split_1[1].split(fechas[1], 1)
                    p3 = split_2[1].strip()
                    valores = [v for v in re.split(r'\s+', p3) if v]
                    filas_crudas.append([id_ap, nom_ap, fechas[0], fechas[1]] + valores)
                except Exception:
                    filas_crudas.append(linea.split())
            else:
                filas_crudas.append(linea.split())

    if not filas_crudas:
        return pd.DataFrame()
    max_cols = max(len(f) for f in filas_crudas)
    header = [f"Columna {i}" for i in range(max_cols)]
    datos_norm = [f + [None] * (max_cols - len(f)) for f in filas_crudas]
    return pd.DataFrame(datos_norm, columns=header)


def limpiar_y_estandarizar(df_crudo, col_desde, col_hasta, col_ibc, col_semanas, col_nit=None, col_aportante=None):
    datos = []
    for _, row in df_crudo.iterrows():
        try:
            raw_desde = str(row[col_desde])
            raw_hasta = str(row[col_hasta])
            raw_ibc = str(row[col_ibc])
            raw_semanas = str(row[col_semanas])
            raw_nit = str(row[col_nit]).strip() if col_nit and col_nit in row else ""
            raw_aportante = str(row[col_aportante]).strip() if col_aportante and col_aportante in row else "No especificado"

            match_d = re.search(r'\d{2}/\d{2}/\d{4}', raw_desde)
            match_h = re.search(r'\d{2}/\d{2}/\d{4}', raw_hasta)
            if not match_d or not match_h:
                continue

            desde = pd.to_datetime(match_d.group(0), dayfirst=True, errors='coerce')
            hasta = pd.to_datetime(match_h.group(0), dayfirst=True, errors='coerce')
            if pd.isna(desde) or pd.isna(hasta):
                continue

            def clean_num(val):
                if not val or val.lower() == 'none':
                    return 0.0
                v = re.sub(r'[^\d\.,]', '', val)
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
            semanas_final = semanas_leidas

            if semanas_leidas <= 0.1 or semanas_leidas > 55:
                dias_calculados = (hasta - desde).days + 1
                semanas_final = dias_calculados / 7 if 0 < dias_calculados < 12000 else 0

            if semanas_final > 0:
                datos.append({
                    "NIT_Aportante": raw_nit,
                    "Aportante": raw_aportante,
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
