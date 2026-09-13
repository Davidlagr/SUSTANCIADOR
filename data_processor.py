import pdfplumber
import pandas as pd
import re
from datetime import datetime, date

def extraer_datos_basicos(pdf_file):
    """
    Extrae cédula, nombre, fecha de nacimiento y sexo 
    del encabezado de la Historia Laboral de Colpensiones.
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
                text = pdf.pages[0].extract_text()
                if text:
                    match_cedula = re.search(r'(?:Número de Documento|Documento|C\.C\.)[\s:]*([\d\.]+)', text, re.IGNORECASE)
                    match_nombre = re.search(r'(?:Nombres y Apellidos|Nombre)[\s:]*([A-ZÑÁÉÍÓÚ\s]+)(?=\n|Fecha|Sexo|Tipo)', text, re.IGNORECASE)
                    match_fecha = re.search(r'(?:Fecha de Nacimiento|Nacimiento)[\s:]*(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
                    match_sexo = re.search(r'(?:Sexo)[\s:]*([MF])', text, re.IGNORECASE)

                    if match_cedula: datos["cedula"] = match_cedula.group(1).replace(".", "").strip()
                    if match_nombre: datos["nombre"] = re.sub(r'\s+', ' ', match_nombre.group(1)).strip()
                    if match_fecha:
                        try: datos["fecha_nac"] = datetime.strptime(match_fecha.group(1), "%d/%m/%Y").date()
                        except: pass
                    if match_sexo: datos["genero"] = "Femenino" if match_sexo.group(1).upper() == 'F' else "Masculino"
    except: pass
    return datos

def extraer_datos_peticion(pdf_file, datos_actuales):
    """
    Rescata datos básicos faltantes desde la petición del ciudadano
    y extrae el bloque exacto de pretensiones o solicitudes para la resolución.
    """
    peticiones = "Reconocimiento de Pensión de Vejez."
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            # Leer las primeras páginas donde suele estar la petición
            for page in pdf.pages[:3]:
                page_text = page.extract_text()
                if page_text: text += page_text + "\n"
                
            if text:
                # 1. Rescate de Datos Básicos faltantes
                if not datos_actuales.get("cedula") or datos_actuales["cedula"] == "":
                    match_cedula = re.search(r'(?:C\.C\.|cédula de ciudadanía.*?N[o\.]?|identificado.*?con.*?N[o\.]?)[\s:]*([\d\.]+)', text, re.IGNORECASE)
                    if match_cedula: datos_actuales["cedula"] = match_cedula.group(1).replace(".", "").strip()
                    
                if not datos_actuales.get("nombre") or datos_actuales["nombre"] == "":
                    match_nombre = re.search(r'(?:Yo[,]?\s+)(.*?)(?:[,]?\s+identificado|[,]?\s+mayor de edad)', text, re.IGNORECASE)
                    if match_nombre: datos_actuales["nombre"] = match_nombre.group(1).strip().upper()

                # 2. Extracción de Peticiones a resolver de fondo
                match_solicitud = re.search(r'(?i)(?:solicito|peticion(?:es)?|pretension(?:es)?|solicitud(?:es)?)[\s\n:]+(.*?)(?:\n\s*(?:hechos|fundamentos|pruebas|anexos|notificaciones|derecho)|\Z)', text, re.DOTALL)
                if match_solicitud:
                    ext = match_solicitud.group(1).strip()
                    ext = re.sub(r'\n+', ' ', ext)
                    if len(ext) > 10: 
                        peticiones = ext
    except: pass
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
        if idx_fin != -1: texto_a_procesar = full_text[idx_inicio:idx_fin]
        else: texto_a_procesar = full_text[idx_inicio:]
    else:
        idx_alt = full_text.find("Identificación Aportante")
        if idx_alt != -1: texto_a_procesar = full_text[idx_alt:]

    lineas = texto_a_procesar.split('\n')
    regex_fecha = re.compile(r'\d{2}/\d{2}/\d{4}')
    
    for linea in lineas:
        linea = linea.strip()
        if not linea: continue
        if not regex_fecha.search(linea): continue
            
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
                    p1 = split_1[0].strip() 
                    split_2 = split_1[1].split(fechas[1], 1)
                    p3 = split_2[1].strip() 
                    valores = re.split(r'\s+', p3)
                    valores = [v for v in valores if v]
                    filas_crudas.append(["(Sin ID)", p1, fechas[0], fechas[1]] + valores)
                except: filas_crudas.append(linea.split())
            else: filas_crudas.append(linea.split())

    if not filas_crudas: return pd.DataFrame()
    max_cols = max(len(f) for f in filas_crudas)
    header = [f"Columna {i}" for i in range(max_cols)]
    datos_norm = [f + [None]*(max_cols-len(f)) for f in filas_crudas]
    return pd.DataFrame(datos_norm, columns=header)

def limpiar_y_estandarizar(df_crudo, col_desde, col_hasta, col_ibc, col_semanas):
    datos = []
    for idx, row in df_crudo.iterrows():
        try:
            raw_desde = str(row[col_desde])
            raw_hasta = str(row[col_hasta])
            raw_ibc = str(row[col_ibc])
            raw_semanas = str(row[col_semanas])
            
            match_d = re.search(r'\d{2}/\d{2}/\d{4}', raw_desde)
            match_h = re.search(r'\d{2}/\d{2}/\d{4}', raw_hasta)
            if not match_d or not match_h: continue
            
            desde = pd.to_datetime(match_d.group(0), dayfirst=True, errors='coerce')
            hasta = pd.to_datetime(match_h.group(0), dayfirst=True, errors='coerce')
            if pd.isna(desde) or pd.isna(hasta): continue
            
            def clean_num(val):
                if not val or val.lower() == 'none': return 0.0
                v = re.sub(r'[^\d\.,]', '', val)
                if not v: return 0.0
                if ',' in v and '.' in v: v = v.replace('.','').replace(',','.')
                elif v.count('.') > 1: v = v.replace('.','')
                elif ',' in v: 
                    if len(v.split(',')[-1])==2: v = v.replace(',','.')
                    else: v = v.replace(',','')
                try: return float(v)
                except: return 0.0

            ibc = clean_num(raw_ibc)
            semanas_leidas = clean_num(raw_semanas)
            semanas_final = semanas_leidas
            
            recalcular = False
            if semanas_leidas <= 0.1: recalcular = True
            elif semanas_leidas > 55: recalcular = True 
            
            if recalcular:
                dias_calculados = (hasta - desde).days + 1
                if 0 < dias_calculados < 12000: semanas_final = dias_calculados / 7
                else: semanas_final = 0
            
            if semanas_final > 0:
                datos.append({"Desde": desde, "Hasta": hasta, "IBC": ibc, "Semanas": semanas_final, "Aportante": "Manual"})
        except Exception as e: continue
            
    df = pd.DataFrame(datos)
    return df.sort_values('Desde') if not df.empty else df

def aplicar_regla_simultaneidad(df):
    if df.empty: return df
    df['Periodo'] = df['Desde'].dt.to_period('M')
    return df.groupby('Periodo').agg({'IBC': 'sum', 'Semanas': 'max', 'Desde': 'min', 'Hasta': 'max'}).reset_index().sort_values('Periodo')
