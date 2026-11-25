import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta


class LiquidadorPension:
    def __init__(self, historia_laboral, genero, fecha_nacimiento):
        self.df = historia_laboral
        self.genero = genero
        self.fecha_nacimiento = pd.to_datetime(fecha_nacimiento)
        self.fecha_actual = datetime.now()

        # IPC HISTÓRICO ANUAL (1967 - 2026)
        self.ipc_historico = {
            1967: 8.3, 1968: 6.8, 1969: 7.7, 1970: 6.7, 1971: 11.6, 1972: 13.6,
            1973: 20.3, 1974: 24.3, 1975: 23.3, 1976: 20.5, 1977: 33.7, 1978: 17.8,
            1979: 24.8, 1980: 26.5, 1981: 27.5, 1982: 24.5, 1983: 19.8, 1984: 16.3,
            1985: 24.0, 1986: 18.9, 1987: 23.3, 1988: 28.1, 1989: 25.8, 1990: 29.1,
            1991: 30.4, 1992: 27.0, 1993: 22.6, 1994: 22.6, 1995: 19.5, 1996: 21.6,
            1997: 17.7, 1998: 16.7, 1999: 9.2, 2000: 8.8, 2001: 7.7, 2002: 7.0,
            2003: 6.5, 2004: 5.5, 2005: 4.9, 2006: 4.5, 2007: 5.7, 2008: 7.7,
            2009: 2.0, 2010: 3.2, 2011: 3.7, 2012: 2.4, 2013: 1.9, 2014: 3.7,
            2015: 6.8, 2016: 5.8, 2017: 4.1, 2018: 3.2, 2019: 3.8, 2020: 1.6,
            2021: 5.6, 2022: 13.1, 2023: 9.3, 2024: 5.0, 2025: 4.0, 2026: 3.5
        }

        # SMMLV HISTÓRICO COLOMBIA (2000 - 2026)
        self.smmlv_historico = {
            2000: 260100.0, 2001: 286000.0, 2002: 309000.0, 2003: 332000.0,
            2004: 358000.0, 2005: 381500.0, 2006: 408000.0, 2007: 433700.0,
            2008: 461500.0, 2009: 496900.0, 2010: 515000.0, 2011: 535600.0,
            2012: 566700.0, 2013: 589500.0, 2014: 616000.0, 2015: 644350.0,
            2016: 689455.0, 2017: 737717.0, 2018: 781242.0, 2019: 828116.0,
            2020: 877803.0, 2021: 908526.0, 2022: 1000000.0, 2023: 1160000.0,
            2024: 1300000.0, 2025: 1423500.0, 2026: 1750905.0
        }

    def obtener_smmlv(self, anio):
        if anio in self.smmlv_historico:
            return self.smmlv_historico[anio]
        if anio > max(self.smmlv_historico.keys()):
            return self.smmlv_historico[max(self.smmlv_historico.keys())]
        return self.smmlv_historico[min(self.smmlv_historico.keys())]

    def obtener_factor_ipc(self, fecha_inicio, fecha_corte):
        """Calcula inflación acumulada hasta una fecha de corte específica."""
        anio_inicio = fecha_inicio.year
        anio_fin = fecha_corte.year

        min_anio = min(self.ipc_historico.keys())
        max_anio = max(self.ipc_historico.keys())

        if anio_inicio < min_anio:
            anio_inicio = min_anio
        if anio_fin > max_anio:
            anio_fin = max_anio

        factor = 1.0
        for anio in range(anio_inicio, anio_fin):
            if anio in self.ipc_historico:
                factor *= (1 + (self.ipc_historico[anio] / 100.0))
        return factor

    def determinar_fechas_clave(self):
        """
        Determina Fecha Cumplimiento Edad, Fecha Cumplimiento Semanas,
        Fecha Estatus y Fecha de Indexación (Corte), aplicando la Sentencia C-197/23.
        """
        req_edad = 62 if self.genero == "Masculino" else 57
        fecha_cumple_edad = self.fecha_nacimiento + relativedelta(years=req_edad)

        df_sort = self.df.sort_values('Hasta')
        acumulado = 0
        fecha_cumple_semanas = None

        for _, row in df_sort.iterrows():
            acumulado += row['Semanas']
            current_date = row['Hasta']

            estatus_teorico = max(fecha_cumple_edad, current_date)
            anio_estatus_teorico = estatus_teorico.year

            if self.genero == "Masculino":
                req_sem_dinamico = 1300
            else:
                if anio_estatus_teorico < 2026:
                    req_sem_dinamico = 1300
                elif anio_estatus_teorico == 2026:
                    req_sem_dinamico = 1250
                else:
                    descenso = 50 + ((anio_estatus_teorico - 2026) * 25)
                    req_sem_dinamico = max(1000, 1300 - descenso)

            if acumulado >= req_sem_dinamico and fecha_cumple_semanas is None:
                fecha_cumple_semanas = current_date
                break

        tiene_estatus = (fecha_cumple_semanas is not None)
        fecha_estatus = max(fecha_cumple_edad, fecha_cumple_semanas) if tiene_estatus else None

        ultima_cotizacion = df_sort['Hasta'].max() if not df_sort.empty else self.fecha_actual
        if not tiene_estatus:
            fecha_corte = self.fecha_actual
            razon_corte = "Año de Estudio (No acredita estatus)"
        else:
            cotizaciones_posteriores = df_sort[df_sort['Hasta'] > (fecha_estatus + timedelta(days=30))]
            if cotizaciones_posteriores.empty:
                fecha_corte = fecha_estatus
                razon_corte = "Fecha de Estatus (Sin semanas posteriores)"
            else:
                fecha_corte = ultima_cotizacion
                razon_corte = "Última Cotización (Con semanas posteriores al estatus)"

        fecha_efectividad = fecha_corte + timedelta(days=1)

        return {
            "fecha_cumple_edad": fecha_cumple_edad,
            "fecha_cumple_semanas": fecha_cumple_semanas,
            "fecha_estatus": fecha_estatus,
            "tiene_estatus": tiene_estatus,
            "fecha_corte": fecha_corte,
            "razon_corte": razon_corte,
            "fecha_efectividad": fecha_efectividad,
            "ultima_cotizacion": ultima_cotizacion
        }

    def calcular_ibl_indexado(self, fecha_corte_personalizada=None, metodo="toda_vida"):
        if self.df.empty:
            return 0.0, pd.DataFrame()

        f_corte = fecha_corte_personalizada if fecha_corte_personalizada else self.fecha_actual
        df_calc = self.df.copy()

        if metodo == "ultimos_10":
            fecha_max_cot = df_calc['Hasta'].max()
            fecha_inicio_10 = fecha_max_cot - relativedelta(years=10)
            df_calc = df_calc[df_calc['Hasta'] >= fecha_inicio_10]

        detalles = []
        for _, row in df_calc.iterrows():
            ibc_hist = max(0.0, row['IBC'])
            factor = self.obtener_factor_ipc(row['Hasta'], f_corte)
            ibc_act = ibc_hist * factor

            detalles.append({
                'Desde': row['Desde'],
                'Hasta': row['Hasta'],
                'IBC_Historico': ibc_hist,
                'Factor_IPC': factor,
                'IBC_Actualizado': ibc_act,
                'Semanas': row['Semanas']
            })

        df_detalles = pd.DataFrame(detalles)
        if df_detalles.empty:
            return 0.0, pd.DataFrame()

        return df_detalles['IBC_Actualizado'].mean(), df_detalles

    def calcular_tasa_reemplazo_797(self, ibl, semanas, anio_pension, limitar_semanas_cotizadas=True, smmlv=None):
        if ibl <= 0:
            return 0.0, 0.0, {}

        smmlv_ref = smmlv if smmlv else self.obtener_smmlv(anio_pension)
        s = ibl / smmlv_ref
        r_inicial = max(55.0, min(65.5 - (0.5 * s), 65.5))

        semanas_minimas = 1300
        if self.genero == 'Femenino':
            if anio_pension < 2026:
                semanas_minimas = 1300
            elif anio_pension == 2026:
                semanas_minimas = 1250
            else:
                diff = anio_pension - 2026
                semanas_minimas = max(1000, 1250 - (25 * diff))

        semanas_computables = min(semanas, 1800) if limitar_semanas_cotizadas else semanas

        semanas_extra = max(0.0, semanas_computables - semanas_minimas)
        bloques = int(semanas_extra // 50)
        # Tope normativo estricto: máximo 15 puntos porcentuales adicionales
        puntos_extra = min(bloques * 1.5, 15.0)

        tasa = max(0.0, min(r_inicial + puntos_extra, 80.0))
        mesada = max(ibl * (tasa / 100.0), smmlv_ref)

        detalle = {
            "smmlv_ref": smmlv_ref,
            "s": s,
            "tasa_base": r_inicial,
            "semanas_usadas": semanas_computables,
            "semanas_adicionales": semanas_extra,
            "bloques": bloques,
            "incremento": puntos_extra,
            "tasa_final": tasa,
            "semanas_minimas_exigidas": semanas_minimas
        }
        return mesada, tasa, detalle

    def calcular_retroactivo(
        self,
        mesada_base,
        anio_base,
        fecha_estatus,
        fecha_ultima_cot,
        fecha_radicacion,
        fecha_inclusion_nomina,
        tipo_vinculacion,
        tiene_retiro_formal=False,
        fecha_retiro_formal=None
    ):
        """
        Calcula efectividad, prescripción trienal (Art. 488 CST / 151 CPTSS),
        reajuste anual de mesadas (Art. 14 Ley 100/93), mesada 13/14 y salud (Ley 2018/2020).
        """
        def to_d(val):
            if val is None or pd.isna(val):
                return None
            return pd.to_datetime(val).date()

        f_est = to_d(fecha_estatus)
        f_ult = to_d(fecha_ultima_cot)
        f_rad = to_d(fecha_radicacion)
        f_inc = to_d(fecha_inclusion_nomina)
        f_ret = to_d(fecha_retiro_formal)

        if f_est is None:
            f_est = f_ult if f_ult else date.today()

        # 1. Regla de efectividad según vinculación
        genera_retroactivo = True
        condicionado_retiro_publico = False

        if tipo_vinculacion == "Servidor Público":
            if tiene_retiro_formal and f_ret:
                f_efec = f_est if f_ret < f_est else (f_ret + timedelta(days=1))
                nota_efec = (
                    f"Servidor público con acto administrativo de retiro acreditado al {f_ret.strftime('%d/%m/%Y')}. "
                    f"Se reconoce efectividad a partir del {f_efec.strftime('%d/%m/%Y')} (Art. 128 C.P., Ley 1821 de 2016 y Art. 19 Ley 344 de 1996)."
                )
            else:
                f_efec = f_inc
                genera_retroactivo = False
                condicionado_retiro_publico = True
                nota_efec = (
                    "Último empleador identificado como entidad pública / SEM >= 90% estatal sin acto administrativo de retiro definitivo. "
                    "Por prohibición constitucional de percibir doble asignación del tesoro público (Art. 128 C.P. y Art. 19 Ley 344 de 1996), "
                    "el disfrute e inclusión en nómina queda condicionado a la acreditación del retiro definitivo del servicio."
                )
        elif tipo_vinculacion == "Activo (Sin novedad de retiro)":
            f_efec = f_inc
            genera_retroactivo = False
            nota_efec = (
                "Afiliado activo sin novedad de retiro en PILA. De conformidad con los artículos 13 y 35 del Acuerdo 049 de 1990 "
                "y la Sentencia C-1037 de 2003, el disfrute opera a partir de la inclusión en nómina de pensionados."
            )
        elif tipo_vinculacion == "Inactivo antes de cumplir edad":
            f_efec = f_est
            nota_efec = (
                f"Afiliado inactivo con anterioridad al cumplimiento de la edad mínima. "
                f"La efectividad rige a partir de la consolidación del estatus el {f_efec.strftime('%d/%m/%Y')} (Art. 33 Ley 100/1993)."
            )
        else:
            base_ret = f_ret if (tiene_retiro_formal and f_ret) else f_ult
            if base_ret is None or base_ret < f_est:
                f_efec = f_est
            else:
                f_efec = base_ret + timedelta(days=1)
            nota_efec = (
                f"Se reconoce efectividad a partir del {f_efec.strftime('%d/%m/%Y')} (día siguiente a la última cotización / retiro acreditado), "
                f"conforme a los artículos 13 y 35 del Acuerdo 049 de 1990."
            )

        # 2. Prescripción trienal (Art. 488 CST y Art. 151 CPTSS)
        f_presc = f_rad - relativedelta(years=3)
        aplica_presc = genera_retroactivo and (f_efec < f_presc)
        f_inicio_pago = max(f_efec, f_presc) if genera_retroactivo else f_inc

        # 3. Evaluación de Mesada 14 (Acto Legislativo 01 de 2005)
        smmlv_caus = self.obtener_smmlv(f_est.year)
        aplica_m14 = (f_est < date(2005, 7, 25)) or (
            f_est < date(2011, 7, 31) and mesada_base <= (3 * smmlv_caus)
        )

        # 4. Proyección anual de mesadas por IPC / SMMLV (Art. 14 Ley 100 de 1993)
        mesadas_por_anio = {anio_base: float(mesada_base)}
        max_y = max(f_inc.year, anio_base)
        min_y = min(f_efec.year, f_inicio_pago.year, anio_base)

        for y in range(anio_base + 1, max_y + 1):
            prev_m = mesadas_por_anio[y - 1]
            prev_smmlv = self.obtener_smmlv(y - 1)
            curr_smmlv = self.obtener_smmlv(y)
            ipc_prev = self.ipc_historico.get(y - 1, 4.0)
            if abs(prev_m - prev_smmlv) < 10.0:
                nueva_m = float(curr_smmlv)
            else:
                nueva_m = max(prev_m * (1.0 + ipc_prev / 100.0), float(curr_smmlv))
            mesadas_por_anio[y] = nueva_m

        for y in range(anio_base - 1, min_y - 1, -1):
            next_m = mesadas_por_anio[y + 1]
            ipc_y = self.ipc_historico.get(y, 4.0)
            curr_smmlv = self.obtener_smmlv(y)
            mesadas_por_anio[y] = max(next_m / (1.0 + ipc_y / 100.0), float(curr_smmlv))

        mesada_actualizada = mesadas_por_anio.get(f_inc.year, float(mesada_base))

        if not genera_retroactivo or f_inicio_pago >= f_inc:
            return {
                "genera_retroactivo": False,
                "condicionado_retiro_publico": condicionado_retiro_publico,
                "fecha_efectividad": f_efec,
                "fecha_inicio_pago": f_inicio_pago,
                "fecha_prescripcion": f_presc,
                "aplica_prescripcion": aplica_presc,
                "aplica_mesada_14": aplica_m14,
                "nota_efectividad": nota_efec,
                "mesada_actualizada": mesada_actualizada,
                "df_detalle": pd.DataFrame(),
                "bruto_ordinarias": 0.0,
                "bruto_adicionales": 0.0,
                "total_bruto": 0.0,
                "total_salud": 0.0,
                "total_neto": 0.0
            }

        # 5. Liquidación mensual comercial (meses de 30 días)
        f_fin = f_inc - timedelta(days=1)
        filas = []
        cur_y, cur_m = f_inicio_pago.year, f_inicio_pago.month

        while (cur_y, cur_m) <= (f_fin.year, f_fin.month):
            d_ini = min(f_inicio_pago.day, 30) if (cur_y == f_inicio_pago.year and cur_m == f_inicio_pago.month) else 1
            d_fin = min(f_fin.day, 30) if (cur_y == f_fin.year and cur_m == f_fin.month) else 30

            if cur_y == f_fin.year and cur_m == f_fin.month:
                prox_mes = date(cur_y + (1 if cur_m == 12 else 0), 1 if cur_m == 12 else cur_m + 1, 1)
                ultimo_dia_real = (prox_mes - timedelta(days=1)).day
                if f_fin.day == ultimo_dia_real:
                    d_fin = 30

            dias = max(0, d_fin - d_ini + 1)
            m_anio = mesadas_por_anio[cur_y]
            smmlv_anio = self.obtener_smmlv(cur_y)
            val_ord = m_anio * (dias / 30.0)

            # Descuento en salud (Art. 143 Ley 100/1993 y Ley 2018/2020)
            ratio = m_anio / smmlv_anio
            if ratio <= 1.001:
                pct_salud = 0.04
            elif ratio <= 2.0:
                pct_salud = 0.10
            else:
                pct_salud = 0.12
            desc_salud = val_ord * pct_salud

            # Mesada 13 (Noviembre) y Mesada 14 (Junio si aplica) - Sin descuento de salud
            val_adic = 0.0
            tipo_adic = "-"
            if cur_m == 11 and dias > 0:
                val_adic = m_anio
                tipo_adic = "Mesada 13"
            elif cur_m == 6 and aplica_m14 and dias > 0:
                val_adic = m_anio
                tipo_adic = "Mesada 14"

            neto = (val_ord - desc_salud) + val_adic
            filas.append({
                "Año": cur_y,
                "Mes": f"{cur_y}-{cur_m:02d}",
                "Días": dias,
                "Mesada Vigente": m_anio,
                "Valor Ordinario": val_ord,
                "Adicional": tipo_adic,
                "Valor Adicional": val_adic,
                "% Salud": pct_salud * 100.0,
                "Descuento Salud": desc_salud,
                "Neto a Pagar": neto
            })

            if cur_m == 12:
                cur_y += 1
                cur_m = 1
            else:
                cur_m += 1

        df_det = pd.DataFrame(filas)
        bruto_ord = float(df_det["Valor Ordinario"].sum())
        bruto_adic = float(df_det["Valor Adicional"].sum())
        tot_salud = float(df_det["Descuento Salud"].sum())

        return {
            "genera_retroactivo": True,
            "condicionado_retiro_publico": condicionado_retiro_publico,
            "fecha_efectividad": f_efec,
            "fecha_inicio_pago": f_inicio_pago,
            "fecha_prescripcion": f_presc,
            "aplica_prescripcion": aplica_presc,
            "aplica_mesada_14": aplica_m14,
            "nota_efectividad": nota_efec,
            "mesada_actualizada": mesada_actualizada,
            "df_detalle": df_det,
            "bruto_ordinarias": bruto_ord,
            "bruto_adicionales": bruto_adic,
            "total_bruto": bruto_ord + bruto_adic,
            "total_salud": tot_salud,
            "total_neto": float(df_det["Neto a Pagar"].sum())
        }
