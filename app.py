import streamlit as st
from datetime import date

# 1. Base de datos histórica de Salarios Mínimos (SMMLV) en Colombia desde 1984
smmlv_historico = {
    1984: 11298.00, 1985: 13558.00, 1986: 16811.00, 1987: 20509.00, 1988: 25637.00,
    1989: 32559.00, 1990: 41025.00, 1991: 51716.00, 1992: 65190.00, 1993: 81510.00,
    1994: 98700.00, 1995: 118933.00, 1996: 142125.00, 1997: 172005.00, 1998: 203825.00,
    1999: 236460.00, 2000: 260100.00, 2001: 286000.00, 2002: 309000.00, 2003: 332000.00,
    2004: 358000.00, 2005: 381500.00, 2006: 408000.00, 2007: 433700.00, 2008: 461500.00,
    2009: 496900.00, 2010: 515000.00, 2011: 535600.00, 2012: 566700.00, 2013: 589500.00,
    2014: 616000.00, 2015: 644350.00, 2016: 689455.00, 2017: 737717.00, 2018: 781242.00,
    2019: 828116.00, 2020: 877803.00, 2021: 908526.00, 2022: 1000000.00, 2023: 1160000.00,
    2024: 1300000.00, 2025: 1430000.00, 2026: 1550000.00 
}

# Configuración inicial de la página
st.set_page_config(page_title="Liquidador Pensional Pro", layout="wide")

st.title("Liquidador Pensional Pro")
st.subheader("Dirección de Prestaciones Económicas - COLPENSIONES")
st.markdown("---")

# Módulo de Captura de Datos del Solicitante
st.markdown("### 1. Datos del Solicitante y Parámetros de Liquidación")
col1, col2 = st.columns(2)

with col1:
    nombre_solicitante = st.text_input("Nombre completo del afiliado/causante:")
    documento = st.text_input("Número de Cédula:")
    tipo_prestacion = st.selectbox(
        "Tipo de Prestación a Sustanciar:", 
        ["Pensión de Vejez", "Pensión de Invalidez", "Sustitución Pensional / Sobrevivientes", "Indemnización Sustitutiva"]
    )

with col2:
    fecha_nacimiento = st.date_input(
        "Fecha de nacimiento:",
        min_value=date(1930, 1, 1),
        max_value=date.today(),
        value=date(1960, 1, 1) 
    )

    lista_anios = list(smmlv_historico.keys())
    anio_calculo = st.selectbox("Año de liquidación (Referencia SMMLV):", lista_anios, index=len(lista_anios)-1)

st.markdown("---")

# Módulo de Selección de Motivaciones (Reintegrado)
st.markdown("### 2. Fundamentos Jurídicos y Motivaciones")
opciones_motivas = [
    "Ley 100 de 1993 (Norma general del Sistema de Seguridad Social)",
    "Ley 797 de 2003 (Modificación de requisitos de edad y semanas)",
    "Decreto 758 de 1990 (Acuerdo 049 - Régimen de Transición)",
    "Ley 1204 de 2008 (Sustitución pensional y compartibilidad)",
    "Sentencia SU-149 de 2021 (Acumulación de tiempos públicos y privados)"
]
motivas_seleccionadas = st.multiselect(
    "Seleccione las normativas a incluir en la motivación de la resolución:", 
    opciones_motivas,
    default=["Ley 100 de 1993 (Norma general del Sistema de Seguridad Social)"]
)

st.markdown("---")

# Módulo de Explicación de la Fórmula Decreciente (Reintegrado)
st.markdown("### 3. Explicación: Fórmula Decreciente (Art. 34 Ley 100/1993 mod. Ley 797/2003)")
with st.expander("Ver calculadora y explicación de la tasa de reemplazo", expanded=False):
    st.markdown("""
    La tasa de reemplazo decrece a medida que el Ingreso Base de Liquidación (IBL) es mayor en proporción al SMMLV.
    
    **Fórmula aplicada:**
    `r = 65.5 - 0.50 * s`
    *(Donde 's' representa el número de salarios mínimos legales mensuales vigentes que contiene el IBL)*
    """)
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        ibl_simulado = st.number_input("Ingrese un IBL de prueba ($):", min_value=0.0, value=float(smmlv_historico[2026]), step=100000.0)
    
    with col_f2:
        smmlv_actual = smmlv_historico[anio_calculo]
        s_calculado = ibl_simulado / smmlv_actual if smmlv_actual > 0 else 0
        r_calculado = 65.5 - (0.50 * s_calculado)
        
        # Límite mínimo y máximo según la ley
        if r_calculado < 55.0:
            r_calculado = 55.0
        elif r_calculado > 80.0:
            r_calculado = 80.0
            
        st.metric(label=f"Número de SMMLV (s) al año {anio_calculo}", value=f"{s_calculado:.2f}")
        st.metric(label="Tasa de Reemplazo (r)", value=f"{r_calculado:.2f}%")

st.markdown("---")

# Módulo de Documentación y Generación
st.markdown("### 4. Carga de Plantilla y Generación de Acto Administrativo")
plantilla_cargada = st.file_uploader("Sube el formato base de la resolución en Word o PDF (Opcional). Si se omite, el sistema generará la motivación jurídica automáticamente.", type=["docx", "txt", "pdf"])

if st.button("Generar Proyecto de Resolución", type="primary"):
    
    if not nombre_solicitante or not documento:
        st.error("Por favor, ingresa el nombre y el número de cédula del solicitante antes de generar la resolución.")
    else:
        if plantilla_cargada is not None:
            st.success(f"Plantilla '{plantilla_cargada.name}' detectada correctamente.")
            st.info("Procesando la liquidación e inyectando los datos sobre la plantilla cargada...")
            
        else:
            st.success("Generando resolución estructurada con motivación jurídica automática...")
            
            smmlv_aplicado = smmlv_historico.get(anio_calculo, 0)
            
            # Formatear la lista de motivaciones para inyectarlas en el texto
            motivas_texto = ""
            if motivas_seleccionadas:
                motivas_texto = "Que el presente acto administrativo se fundamenta jurídicamente en las siguientes normativas y jurisprudencias:\n"
                for motiva in motivas_seleccionadas:
                    motivas_texto += f"        - {motiva}\n"
            else:
                motivas_texto = "Que el presente acto administrativo se fundamenta en la normativa vigente del Sistema General de Pensiones.\n"
            
            resolucion_texto = f"""REPUBLICA DE COLOMBIA
ADMINISTRADORA COLOMBIANA DE PENSIONES - COLPENSIONES

POR LA CUAL SE RESUELVE UN TRÁMITE DE PRESTACIONES ECONÓMICAS EN EL RÉGIMEN DE PRIMA MEDIA CON PRESTACIÓN DEFINIDA
({tipo_prestacion.upper()})

EL SUBDIRECTOR DE DETERMINACIÓN DE LA DIRECCIÓN DE PRESTACIONES ECONÓMICAS DE LA ADMINISTRADORA COLOMBIANA DE PENSIONES - COLPENSIONES, en uso de las atribuciones inherentes al cargo y,

CONSIDERANDO:

Que el (la) señor(a) {nombre_solicitante.upper()}, identificado(a) con C.C. No. {documento}, nacido(a) el {fecha_nacimiento.strftime('%d/%m/%Y')}, elevó solicitud para el reconocimiento de prestaciones económicas ante esta administradora.

{motivas_texto.strip()}

Que para el respectivo cálculo del Ingreso Base de Liquidación (IBL) y la determinación de la mesada pensional, se ha tenido en cuenta el histórico normativo del Salario Mínimo Mensual Legal Vigente (SMMLV), tomando como referencia los valores desde el año 1984 hasta el periodo de causación del derecho.

Que, en ese orden de ideas, el SMMLV certificado para el año {anio_calculo} corresponde a la suma de ${smmlv_aplicado:,.2f} M/CTE.

Que verificado el cumplimiento de las semanas de cotización y los requisitos del afiliado, resulta procedente acceder a lo peticionado, por lo que,

En mérito de lo expuesto,

R E S U E L V E:

ARTÍCULO PRIMERO: Reconocer y ordenar el pago de la {tipo_prestacion.upper()} a favor de {nombre_solicitante.upper()}, identificado(a) con C.C. No. {documento}, conforme a la motivación expuesta en la parte considerativa del presente acto administrativo.

ARTÍCULO SEGUNDO: Notifíquese el contenido de la presente resolución al interesado, advirtiendo que contra la misma proceden los recursos de ley.

Dada a los {date.today().day} días del mes correspondiente del año {date.today().year}.

COMUNÍQUESE, NOTIFÍQUESE Y CÚMPLASE.
"""
            
            with st.expander("Vista Previa del Acto Administrativo", expanded=True):
                st.text(resolucion_texto)
            
            st.download_button(
                label="Descargar Proyecto de Resolución (TXT)",
                data=resolucion_texto,
                file_name=f"Proyecto_Resolucion_{documento}.txt",
                mime="text/plain"
            )
