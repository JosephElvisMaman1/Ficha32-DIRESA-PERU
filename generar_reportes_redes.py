# -*- coding: utf-8 -*-
"""Generador de Informes Técnicos por Red de Salud — Ficha N.° 32 — corte del mes configurado + plan de cumplimiento.

El periodo (mes de corte, año) y la DIRESA activa se toman de ficha32_modelo (M.ETQ / M.DIRESA).

Formato institucional heredado de los informes del Mes 7 (banner, Arial, texto en negro, tono técnico y constructivo,
sin firmas). Todas las cifras provienen de ficha32_modelo.construir() (única fuente de verdad).
"""
import os
import sys
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

import ficha32_modelo as M
import ficha32_graficos as G
from docx_utils import *  # noqa
from docx_utils import _run

OUT = os.path.join(M.BASE, "words")
os.makedirs(OUT, exist_ok=True)

ESPERADO = M.MES_CORTE / M.MESES_AÑO * 100
SEM_TXT = {"VERDE": "Verde", "AMARILLO": "Amar.", "ROJO": "Rojo"}
ORDEN = {"ROJO": 0, "AMARILLO": 1, "VERDE": 2}


def nivel_texto(df):
    vc = df["NIVEL"].value_counts()
    trozos = [f"{vc.get(n, 0)} de categoría {n}" for n in ["I-1", "I-2", "I-3", "I-4", "II-1", "II-2"] if vc.get(n, 0)]
    return "(" + (", ".join(trozos[:-1]) + " y " + trozos[-1] if len(trozos) > 1 else trozos[0]) + ")"


def nombres_pts(sub, col="PUNTAJE"):
    return [f"{r['NOMBRE']} ({r[col]})" for _, r in sub.iterrows()]


def pct(a, b):
    return round(a / b * 100, 1) if b else 0.0


def construir_informe(doc, prov, disp, slug, df, dfb, figs):
    total = len(df)
    hosp_nombre = M.HOSPITAL_REF.get(prov, "")
    tiene_hosp = dfb is not None and len(dfb) > 0
    cnt = {k: int((df["SEMAFORO"] == k).sum()) for k in ("VERDE", "AMARILLO", "ROJO")}
    cnt7 = {k: int((df["SEMAFORO_ANT"] == k).sum()) for k in ("VERDE", "AMARILLO", "ROJO")}
    mean7, mean8 = df["PUNTAJE_ANT"].mean(), df["PUNTAJE"].mean()
    med8 = df["PUNTAJE"].median()
    sumas = {s: dict(r7=int(df[f"{s} ANT"].sum()), r8=int(df[f"{s} REALIZADAS"].sum()),
                     m=int(df[f"{s} META"].sum()), mes=int(df[f"{s} MES{M.MES_CORTE}"].sum())) for s in M.SERV}
    cero = {s: int((df[f"{s} REALIZADAS"] == 0).sum()) for s in M.SERV}
    cumplen = {s: int((df[f"{s} REALIZADAS"] >= df[f"{s} META"]).sum()) for s in M.SERV}
    sin_avance = int((df["DELTA_PTS"] == 0).sum())
    mejoran = int((df["DELTA_PTS"] > 0).sum())
    suben_sem = int((df["SEMAFORO"].map(ORDEN) > df["SEMAFORO_ANT"].map(ORDEN)).sum())
    esc = {c: df[c].mean() for c in ("PUNTAJE_PROY_OCT", "PUNTAJE_DIC_TEND", "PUNTAJE_DIC_AGO", "PUNTAJE_PLAN_BASE",
                                     "PUNTAJE_PLAN")}
    catn = {c: int((df["CATEGORIA"] == c).sum()) for c in M.CATEGORIAS}
    df = df.copy()
    df["SEM_DIC_TEND"] = df["PUNTAJE_DIC_TEND"].apply(M.clasificar)

    # ------------------------------------------------------------------ página / pie
    for s in doc.sections:
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)
        p_ft = s.footer.paragraphs[0]
        p_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_ft = p_ft.add_run(f"Coordinación de Telemedicina — DIRESA {M.DIRESA.title()} | Red de Salud {disp} (Ficha N.° 32 — {M.ETQ['MES']})")
        r_ft.font.name = "Arial"
        r_ft.font.size = Pt(8)
        r_ft.font.color.rgb = COLOR_DARK_GRAY

    # ------------------------------------------------------------------ encabezado
    tbl_hdr = doc.add_table(rows=1, cols=2)
    tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_hdr.autofit = False
    cl, cr = tbl_hdr.cell(0, 0), tbl_hdr.cell(0, 1)
    fijar_anchos(tbl_hdr, [4.6, 2.4])
    set_cell_background(cl, "F1F5F9")
    set_cell_background(cr, "E2E8F0")
    set_cell_margins(cl, 140, 140, 180, 140)
    set_cell_margins(cr, 140, 140, 140, 140)
    pl = cl.paragraphs[0]
    pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(BANNER_PATH):
        pl.add_run().add_picture(BANNER_PATH, width=Inches(4.3))
    else:
        _run(pl, f"GOBIERNO REGIONAL {M.DIRESA}\nDIRECCIÓN REGIONAL DE SALUD {M.DIRESA}\nCoordinación de Telemedicina", 9, bold=True)
    pr = cr.paragraphs[0]
    pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _run(pr, f"INFORME TÉCNICO {disp}\n", 9, bold=True)
    _run(pr, f"Monitoreo Situacional y Plan de Cumplimiento\nCorte: {M.ETQ['MES']} / Año {M.ETQ['ANIO_TXT']}", 8.5, color=COLOR_DARK_GRAY)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    pt = doc.add_paragraph()
    pt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pt.paragraph_format.space_before = Pt(12)
    pt.paragraph_format.space_after = Pt(4)
    _run(pt, "INFORME TÉCNICO DE MONITOREO SITUACIONAL EN TELEMEDICINA\n"
             f"RED DE SALUD {disp} — CORTE {M.ETQ['MES'].upper()} Y PLAN DE CUMPLIMIENTO {M.ETQ['PLAN_LARGO'].upper()}", 14, bold=True)
    ps = doc.add_paragraph()
    ps.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ps.paragraph_format.space_after = Pt(14)
    _run(ps, f"Seguimiento Operativo de los {total} Establecimientos de Primer Nivel"
             + (f" y del {hosp_nombre}" if tiene_hosp else "") +
             f"\nFicha Técnica N.° 32 (Convenios de Gestión MINSA {M.ETQ['ANIO_TXT']}) | Plataforma Oficial https://ficha32-ditel.web.app/",
         9.5, italic=True, color=COLOR_DARK_GRAY)

    add_callout_box(
        doc, f"PRESENTACIÓN Y PROPÓSITO DEL INFORME TÉCNICO — RED DE SALUD {disp}",
        f"El presente informe ha sido elaborado por la Coordinación de Telemedicina para la Dirección de la Red de Salud {disp}, "
        f"los jefes de Microred y los responsables de cada establecimiento. Presenta el estado de avance de los {total} "
        f"establecimientos de primer nivel" + (" (que integra territorialmente las Microredes de las provincias de Huancané y Moho)" if prov == "HUANCANE" else "") + (f" y del {hosp_nombre}" if tiene_hosp else "") +
        f" con la información oficial al {M.ETQ['MES']} (enero–{M.ETQ['MES_LARGO'].lower()} {M.ETQ['ANIO_TXT']}), la comparación con el Mes {M.MES_CORTE - 1} y una ruta concreta para "
        f"orientar el trabajo de {M.ETQ['PLAN_LARGO'].lower()}.\n\n"
        "El propósito es netamente constructivo y orientador: identificar qué establecimientos están cerca de su meta anual, "
        "cuáles pueden mejorar el escenario de la red con pocas atenciones adicionales y qué estrategias permitirán avanzar a "
        "aquellos con mayor brecha, aun cuando no alcancen la meta en este periodo.",
        bg_hex="F8FAFC", border_color="475569")

    # ------------------------------------------------------------------ 1. Marco metodológico
    add_heading_1(doc, "1. MARCO METODOLÓGICO Y REGLAS DE EVALUACIÓN (FICHA N.° 32)")
    add_body_p(doc,
               "El monitoreo se rige por la Ficha Técnica N.° 32 («Tasa de uso de los servicios de telemedicina») del MINSA. "
               f"Las metas son ANUALES (evaluación final: 1 de enero al 31 de diciembre de {M.ETQ['ANIO_TXT']}) y las atenciones que muestra la "
               f"plataforma son ACUMULADAS desde enero hasta el mes de corte. Se verificó con los reportes del Mes {M.MES_CORTE - 1} y del {M.ETQ['MES']}: "
               f"las metas no varían entre meses y las atenciones acumuladas nunca disminuyen. Al cierre de {M.ETQ['MES_LARGO'].lower()} han transcurrido "
               f"{M.MES_CORTE} de {M.MESES_AÑO} meses, por lo que el avance lineal esperado es de {ESPERADO:.1f} % de la meta anual.",
               bold_prefix="Metas anuales y producción acumulada: ")
    add_tabla(doc, ["Servicio Evaluado", "Código", "Peso Oficial", "Fórmula de Meta ANUAL (Primer Nivel)", "Códigos CPMS"],
              [["Teleinterconsulta", "TI", "60 puntos", "máx(24, redondeo[(120/10000) × Población])", "99499.11 y 99499.12"],
               ["Teleconsulta", "TC", "10 puntos", "máx(1, redondeo[(10/10000) × Población])", "99499.01 y 99499.03"],
               ["Telemonitoreo", "TM", "30 puntos", "máx(24, redondeo[(72/10000) × Población])",
                "99499.10 con CIE-10 del Anexo 1 y tipo de diagnóstico R"]],
              [1.2, 0.5, 0.85, 2.7, 1.75], ["L", "C", "C", "L", "L"], font=8.3, bold_cols=(2,), pad=80)
    add_body_p(doc,
               f"Para los {total} establecimientos del primer nivel en {disp}:\n"
               "Puntaje de Logro = [min(TI, MetaTI)/MetaTI × 60] + [min(TC, MetaTC)/MetaTC × 10] + [min(TM, MetaTM)/MetaTM × 30]\n"
               "Cumplimiento oficial (%) = (Puntaje − 10) / (100 − 10) × 100 (valor umbral 10; logro esperado 100).\n\n"
               "Semáforo de gestión: Verde (60.0 a 100 pts, cumplimiento destacado); Amarillo (30.0 a 59.9 pts, avance regular); "
               "Rojo (0.0 a 29.9 pts, prioridad de acompañamiento). El tope por servicio evita que la sobreproducción de una "
               "modalidad compense la inactividad de otra." +
               ("\n\nPara el hospital consultor (logro B) solo se evalúa la teleinterconsulta (peso 100) con meta "
                "TI = máx(100, redondeo[(40/10000) × atenciones totales HIS acumuladas]); esta meta crece conforme avanza el año."
                if tiene_hosp else ""),
               bold_prefix="Fórmula y clasificación semafórica: ")

    # ------------------------------------------------------------------ 2. Diagnóstico
    add_heading_1(doc, f"2. DIAGNÓSTICO SITUACIONAL AL {M.ETQ['MES'].upper()} DE LA RED DE SALUD {disp}")
    add_body_p(doc,
               f"La red comprende {total} establecimientos de primer nivel {nivel_texto(df)}" + (" (35 en la provincia de Huancané y 16 en la provincia de Moho)" if prov == "HUANCANE" else "") + f". El Puntaje de Logro promedio pasó de "
               f"{n1(mean7)} puntos (Mes {M.MES_CORTE - 1}) a {n1(mean8)} puntos ({M.ETQ['MES']}), con una mediana de {n1(med8)} puntos. "
               f"{mejoran} establecimientos ({pct(mejoran, total)} %) mejoraron su puntaje en el mes y {suben_sem} subieron de semáforo; "
               f"{sin_avance} no registraron producción adicional entre ambos cortes.",
               bold_prefix="Balance operativo: ")
    filas = []
    fills = []
    desc = {"ROJO": "Prioridad de acompañamiento (<30 pts)", "AMARILLO": "Avance regular (30 a 59.9 pts)",
            "VERDE": "Cumplimiento destacado (≥60 pts)"}
    for k in ("ROJO", "AMARILLO", "VERDE"):
        dlt = cnt[k] - cnt7[k]
        filas.append([k, desc[k], f"{cnt7[k]}", f"{cnt[k]}", f"{dlt:+d}", f"{pct(cnt[k], total)} %"])
        fills.append({0: SEM_BG[k]})
    add_tabla(doc, ["Semáforo", "Nivel de Avance", f"IPRESS Mes {M.MES_CORTE - 1}", f"IPRESS {M.ETQ['MES']}", "Variación", f"% del Total ({M.ETQ['MES']})"],
              filas, [0.9, 2.2, 0.9, 0.9, 0.8, 1.2], ["C", "L", "C", "C", "C", "C"], font=8.5, fills=fills,
              bold_cols=(0, 3), pad=80)
    add_figura(doc, figs["semaforo"], f"Figura 1: Semáforo de gestión de las {total} IPRESS de la Red de Salud {disp} — Mes {M.MES_CORTE - 1} vs {M.ETQ['MES']}")

    add_heading_2(doc, "Avance acumulado por servicio frente a la meta anual")
    filas = []
    for s in M.SERV:
        d = sumas[s]
        filas.append([f"{M.NOMBRE_SERV[s]} ({s})", n0(d["r7"]), n0(d["mes"]), n0(d["r8"]), n0(d["m"]),
                      f"{pct(d['r8'], d['m'])} %", f"{cumplen[s]} de {total}"])
    add_tabla(doc, ["Servicio", f"Acumulado Mes {M.MES_CORTE - 1}", f"Producción del {M.ETQ['MES']}", f"Acumulado {M.ETQ['MES']}", "Meta anual",
                    f"% de la meta (esperado {ESPERADO:.1f} %)", "IPRESS con meta anual cumplida"],
              filas, [1.5, 0.85, 0.9, 0.85, 0.75, 1.1, 1.05], ["L", "C", "C", "C", "C", "C", "C"], font=8.3, pad=80)
    add_figura(doc, figs["avance"], f"Figura 2: Avance acumulado (enero–{M.ETQ['MES_LARGO'].lower()}) frente a la meta anual por servicio — Red de Salud {disp}")

    ti, tc, tm = sumas["TI"], sumas["TC"], sumas["TM"]
    add_heading_2(doc, f"Aspectos Clave del Diagnóstico en {disp}")
    add_body_p(doc,
               f"La red acumula {n0(ti['r8'])} de {n0(ti['m'])} teleinterconsultas programadas para el año ({pct(ti['r8'], ti['m'])} %), "
               f"frente a {ESPERADO:.1f} % de avance lineal esperado; en {M.ETQ['MES_LARGO'].lower()} se registraron {n0(ti['mes'])}. "
               f"{cero['TI']} de {total} establecimientos ({pct(cero['TI'], total)} %) aún no registran ninguna TI en el año. "
               "Como la TI aporta 60 de los 100 puntos, es el componente que más puede mover el semáforo.",
               bold_prefix="1. Teleinterconsulta (TI - Peso 60 %): ")
    add_body_p(doc,
               f"El telemonitoreo acumula {n0(tm['r8'])} atenciones frente a una meta anual de {n0(tm['m'])} ({pct(tm['r8'], tm['m'])} %); "
               f"{cumplen['TM']} establecimientos ya alcanzaron su meta anual de TM y {cero['TM']} ({pct(cero['TM'], total)} %) todavía "
               "no registran telemonitoreo. Es el servicio de ejecución más autónoma de la IPRESS (personal de enfermería, obstetricia, "
               "psicología o medicina) y no depende de un hospital consultor.",
               bold_prefix="2. Telemonitoreo (TM - Peso 30 %): ")
    add_body_p(doc,
               f"La teleconsulta acumula {n0(tc['r8'])} atenciones frente a una meta anual de {n0(tc['m'])} ({pct(tc['r8'], tc['m'])} %); "
               f"{cumplen['TC']} establecimientos ya cumplieron su meta anual de TC y {cero['TC']} ({pct(cero['TC'], total)} %) no registran "
               "ninguna. En muchos establecimientos la meta anual de TC es de apenas 1 a 3 atenciones: una sola teleconsulta puede aportar "
               "hasta 10 puntos.",
               bold_prefix="3. Teleconsulta (TC - Peso 10 %): ")
    top_mov = df.sort_values("DELTA_PTS", ascending=False).head(3)
    top_mov = top_mov[top_mov["DELTA_PTS"] > 0]
    if len(top_mov):
        add_body_p(doc, "; ".join(f"{r['NOMBRE']} (+{r['DELTA_PTS']} pts: {r['PUNTAJE_ANT']} → {r['PUNTAJE']})"
                                  for _, r in top_mov.iterrows()) + ".",
                   bold_prefix="4. Mayor avance del mes: ")

    # Acreditación en SUSALUD RENIPRESS (Cartera de Servicios y Telesalud)
    subio_doc = int((df.get("VARIACION_RESOLUCION", pd.Series(dtype=str)) == "SUBIÓ").sum())
    con_res_num = int((df.get("NRO_RESOLUCION", pd.Series(dtype=str)).fillna("—") != "—").sum())
    con_tele = int(df.get("TELESALUD_ROL", pd.Series(dtype=str)).isin(["CONSULTOR", "CONSULTANTE", "AMBOS"]).sum())
    add_body_p(doc,
               f"En la plataforma oficial de SUSALUD RENIPRESS, {con_tele} de los {total} establecimientos ({pct(con_tele, total)} %) "
               f"cuentan con modalidad de Telesalud registrada (consultor/consultante). Respecto a la Cartera de Servicios formal, "
               f"{con_res_num} disponen de número de resolución registrado, de los cuales {subio_doc} ({pct(subio_doc, total)} %) "
               f"tienen el documento digital de resolución efectivamente cargado en el sistema («SUBIÓ»). Se recomienda a la Red gestionar "
               f"la regularización documental de los establecimientos pendientes.",
               bold_prefix="5. Estado de Cartera de Servicios y Telesalud (SUSALUD): ")

    # ------------------------------------------------------------------ 3. Matriz
    add_heading_1(doc, f"3. MATRIZ DETALLADA DE EVALUACIÓN: CADA ESTABLECIMIENTO DE {disp}")
    add_body_p(doc,
               f"Se presenta la matriz completa de los {total} establecimientos de primer nivel, ordenada de mayor a menor Puntaje de "
               f"Logro al {M.ETQ['MES']}. R/M = atenciones acumuladas (enero–{M.ETQ['MES_LARGO'].lower()}) / meta anual. La columna «Pts M{M.MES_CORTE - 1}→M{M.MES_CORTE}» muestra la evolución "
               "del puntaje entre ambos cortes.", bold_prefix="Matriz Nominal de IPRESS (Primer Nivel): ")
    ms = df.sort_values(["PUNTAJE", "PUNTAJE_ANT"], ascending=False).reset_index(drop=True)
    filas, fills = [], []
    for i, r in ms.iterrows():
        filas.append([i + 1, r["NIVEL"], r["NOMBRE"], n0(r["DENOMINADOR"]),
                      f"{int(r['TI REALIZADAS'])}/{int(r['TI META'])}", f"{int(r['TC REALIZADAS'])}/{int(r['TC META'])}",
                      f"{int(r['TM REALIZADAS'])}/{int(r['TM META'])}", f"{r['PUNTAJE_ANT']} → {r['PUNTAJE']}",
                      f"{SEM_TXT[r['SEMAFORO']]} ({r['PUNTAJE']})"])
        fills.append({8: SEM_BG[r["SEMAFORO"]]})
    add_tabla(doc, ["N.°", "Cat.", "Establecimiento / IPRESS", "Población", "TI (R/M)", "TC (R/M)", "TM (R/M)",
                    f"Pts M{M.MES_CORTE - 1}→M{M.MES_CORTE}", f"Semáforo ({M.ETQ['MES']})"],
              filas, [0.3, 0.4, 1.75, 0.65, 0.7, 0.65, 0.7, 0.9, 0.95], ["C", "C", "L", "C", "C", "C", "C", "C", "C"],
              font=7.6, fills=fills, bold_cols=(0, 8))
    add_figura(doc, figs["ranking"], f"Figura 3: Ranking de IPRESS por Puntaje de Logro al {M.ETQ['MES']} (rombo = puntaje del Mes {M.MES_CORTE - 1}) — Red de Salud {disp}")

    # ------------------------------------------------------------------ 4. Grupos
    add_heading_1(doc, f"4. ANÁLISIS POR NIVELES DE AVANCE EN {disp}")
    gv = ms[ms["SEMAFORO"] == "VERDE"]
    ga = ms[ms["SEMAFORO"] == "AMARILLO"]
    gr = ms[ms["SEMAFORO"] == "ROJO"]
    gr0 = gr[gr["PUNTAJE"] == 0]
    grp = gr[gr["PUNTAJE"] > 0]
    add_heading_2(doc, f"Establecimientos Destacados (Semáforo Verde — {len(gv)} IPRESS)")
    if len(gv):
        add_body_p(doc, "; ".join(nombres_pts(gv)) + ". Combinan producción en las tres modalidades y constituyen la referencia "
                        "de buenas prácticas para el resto de la red.", bold_prefix="Buenas Prácticas en la Red: ")
    else:
        add_body_p(doc, f"En el corte al {M.ETQ['MES']} ningún establecimiento alcanza aún el Semáforo Verde; los del grupo amarillo "
                        "son los candidatos más cercanos.", bold_prefix="Situación del Grupo Destacado: ")
    add_heading_2(doc, f"Establecimientos en Avance Regular (Semáforo Amarillo — {len(ga)} IPRESS)")
    if len(ga):
        add_body_p(doc, "; ".join(nombres_pts(ga)) + ". Su característica común es un cumplimiento aceptable en telemonitoreo y/o "
                        "teleconsulta con producción todavía incipiente en teleinterconsulta.",
                   bold_prefix="Potencial Inmediato de Crecimiento: ")
    else:
        add_body_p(doc, "Ningún establecimiento se ubica en el rango amarillo en este corte.", bold_prefix="Situación del Grupo Intermedio: ")
    add_heading_2(doc, f"Establecimientos Prioritarios para Acompañamiento (Semáforo Rojo — {len(gr)} IPRESS)")
    if len(gr):
        txt = "Se ubican entre 0.0 y 29.9 puntos. Conviene distinguir dos realidades:\n"
        if len(grp):
            txt += (f"1. Con avance inicial ({len(grp)} IPRESS): " + "; ".join(nombres_pts(grp)) +
                    ". Muestran voluntad operativa y requieren soporte para activar interconsultas.\n")
        if len(gr0):
            txt += (f"2. Sin producción registrada ({len(gr0)} IPRESS con 0 puntos): " +
                    ", ".join(gr0["NOMBRE"]) + ". Se recomienda verificar si la inactividad responde a falta de capacitación, "
                    "conectividad, ausencia de responsable o demoras en la digitación HIS.")
        add_body_p(doc, txt, bold_prefix="Focalización del Soporte Operativo: ")
    else:
        add_body_p(doc, "Ningún establecimiento está en rojo en este corte.", bold_prefix="Situación del Grupo Prioritario: ")

    # ------------------------------------------------------------------ 5. Hospital
    sec = 5
    if tiene_hosp:
        for _, h in dfb.iterrows():
            nom_hosp = h["NOMBRE"] if h["NOMBRE"] else (hosp_nombre or "Hospital Consultor")
            add_heading_1(doc, f"5. ROL CONSULTOR: {nom_hosp.upper()} (LOGRO B)")
            add_body_p(doc,
                       f"El hospital consultor cumple un rol clave de la red {disp}: su meta TI es dinámica y depende de sus atenciones "
                       f"HIS acumuladas ({n0(h['DENOMINADOR'])} al {M.ETQ['MES']}, frente a {n0(h['DEN ANT'])} al Mes {M.MES_CORTE - 1}). Con el ritmo actual, se "
                       f"proyecta una meta de {n0(h['META_PROY_DIC'])} TI al cierre de {M.MES_TXT[M.MESES_AÑO].lower()}.",
                       bold_prefix="Evaluación Institucional del Hospital: ")
            add_tabla(doc, ["Hospital", "Cat.", "Atenciones HIS (acum.)", f"Meta TI {M.ETQ['MES']}", f"TI Mes {M.MES_CORTE - 1}", f"TI {M.ETQ['MES']}", f"Pts M{M.MES_CORTE - 1}→M{M.MES_CORTE}",
                            f"Meta TI proy. {M.MES_ABR[M.MESES_AÑO]}"],
                      [[h["NOMBRE"], h["NIVEL"], n0(h["DENOMINADOR"]), n0(h["TI META"]), n0(h["TI ANT"]),
                        n0(h["TI REALIZADAS"]), f"{h['PUNTAJE_ANT']} → {h['PUNTAJE']}", n0(h["META_PROY_DIC"])]],
                      [1.7, 0.4, 0.9, 0.7, 0.6, 0.6, 0.9, 0.8], ["L", "C", "C", "C", "C", "C", "C", "C"], font=8,
                      fills=[{6: SEM_BG[h["SEMAFORO"]]}], bold_cols=(0, 6), pad=80)
            if h["PUNTAJE"] >= 100:
                add_body_p(doc,
                           f"El hospital ya supera su meta con {n0(h['TI REALIZADAS'])} TI (meta {n0(h['TI META'])}) y su proyección a {M.MES_TXT[M.MESES_AÑO].lower()} "
                           f"mantiene el cumplimiento ({n0(h['TI_PROY_DIC_TEND'])} TI frente a {n0(h['META_PROY_DIC'])}). Su capacidad consultora "
                           f"es un activo para la red: los {total} establecimientos de primer nivel acumulan {n0(ti['r8'])} TI frente a {n0(ti['m'])} "
                           "de meta, por lo que existe espacio para canalizar más interconsultas hacia el hospital.",
                           bold_prefix="Oportunidad de Articulación Red-Hospital: ")
            else:
                add_body_p(doc,
                           f"Para alcanzar su meta proyectada de {M.MES_TXT[M.MESES_AÑO].lower()} ({n0(h['META_PROY_DIC'])} TI), partiendo de {n0(h['TI_PROY_OCT'])} TI "
                           f"proyectadas a fin de {M.ETQ['HASTA'].lower()}, se requerirían {n0(h['BRECHA_DIC'])} TI en {M.ETQ['PLAN_LARGO'].lower()} (≈{n0(h['REQ_MES'])} por mes). "
                           f"Con un esfuerzo exigente sostenido (≈{n0(h['CAP_MES'])} TI/mes, el mayor entre su ritmo de {M.ETQ['MES_LARGO'].lower()} y el doble de su "
                           f"promedio enero–{M.ETQ['MES_LARGO'].lower()}) el hospital alcanzaría ≈{h['PUNTAJE_PLAN']} puntos al cierre del año. "
                           "La articulación con los establecimientos de la red (derivar TI en lugar de referencias físicas) y la programación "
                           "de turnos de especialidad son las palancas principales.",
                           bold_prefix="Ruta de Mejora del Hospital: ")
        sec = 6

    # ------------------------------------------------------------------ 6. PLAN
    add_heading_1(doc, f"{sec}. PLAN DE CUMPLIMIENTO {M.ETQ['PLAN_LARGO'].upper()} {M.ETQ['ANIO_TXT']} — RED DE SALUD {disp}")
    add_heading_2(doc, f"{sec}.1 Premisas y método de proyección")
    add_callout_box(
        doc, "SUPUESTOS DEL PLAN (TRANSPARENCIA METODOLÓGICA)",
        f"• Corte oficial disponible: {M.ETQ['MES']} ({M.ETQ['MES_LARGO'].lower()}). Los meses hasta fin de {M.ETQ['HASTA'].lower()} se PROYECTAN al ritmo promedio "
        f"mensual de {M.ETQ['RANGO'].lower()} de cada establecimiento y servicio (escenario conservador). Al publicarse reportes posteriores, "
        "el modelo se actualiza reemplazando la proyección por la cifra oficial.\n"
        f"• Brecha de {M.ETQ['PLAN']} = meta anual − atenciones proyectadas a fin de {M.ETQ['HASTA'].lower()} (por servicio, sin sobreproducción útil).\n"
        "• Capacidad realista mensual de cada IPRESS = el mayor entre 12 atenciones, el 5 % de su meta anual total y su rendimiento mensual ya demostrado.\n"
        "• Plan base (exigible): establecimientos CERCA y MODERADO completan su meta anual; los de BRECHA ALTA duplican su ritmo y, como piso de "
        "arranque, aportan un tercio de su capacidad mensual, priorizando el servicio que más puntos rinde por atención.\n"
        f"• Plan ampliado (techo operativo): todos usan su capacidad completa durante {M.ETQ['PLAN']}.",
        bg_hex="F8FAFC", border_color="475569")
    add_tabla(doc, ["Categoría", "Criterio (esfuerzo requerido ÷ capacidad mensual)", "N.° IPRESS", f"Atenciones a registrar en {M.ETQ['PLAN']} (plan base)"],
              [[c, d, str(catn[c]), n0(df.loc[df['CATEGORIA'] == c, ['TI OBJ_ND', 'TC OBJ_ND', 'TM OBJ_ND']].sum().sum())]
               for c, d in zip(M.CATEGORIAS, ["Meta anual asegurada (brecha 0)", "Hasta 1/3 de su capacidad: meta al alcance",
                                              "Entre 1/3 y 1 vez su capacidad: requiere esfuerzo sostenido",
                                              "Más de 1 vez su capacidad: ruta escalonada de avance"])],
              [1.1, 3.3, 0.8, 1.8], ["C", "L", "C", "C"], font=8.3, fills=[{0: CAT_BG[c]} for c in M.CATEGORIAS],
              bold_cols=(0,), pad=70)
    add_figura(doc, figs["categorias"], f"Figura 4: Distribución de las IPRESS según esfuerzo requerido para cumplir la meta — Red de Salud {disp}")

    add_heading_2(doc, f"{sec}.2 Escenarios a {M.MES_TXT[M.MESES_AÑO].lower()} de {M.ETQ['ANIO_TXT']}")
    s_tend = {k: int((df["SEM_DIC_TEND"] == k).sum()) for k in ("VERDE", "AMARILLO", "ROJO")}
    s_base = {k: int((df["SEMAFORO_PLAN_BASE"] == k).sum()) for k in ("VERDE", "AMARILLO", "ROJO")}
    s_amp = {k: int((df["SEMAFORO_PLAN"] == k).sum()) for k in ("VERDE", "AMARILLO", "ROJO")}
    s_oct = {k: int((df["SEMAFORO_OCT"] == k).sum()) for k in ("VERDE", "AMARILLO", "ROJO")}

    def cum(p):
        return f"{max(0, (p - 10) / 90 * 100):.1f} %"
    filas = [
        [f"{M.ETQ['MES']} (oficial)", n1(mean8), cum(mean8), f"{cnt['VERDE']} / {cnt['AMARILLO']} / {cnt['ROJO']}"],
        [f"Proyección fin de {M.ETQ['HASTA'].lower()} (ritmo {M.ETQ['RANGO'].lower()})", n1(esc["PUNTAJE_PROY_OCT"]), cum(esc["PUNTAJE_PROY_OCT"]),
         f"{s_oct['VERDE']} / {s_oct['AMARILLO']} / {s_oct['ROJO']}"],
        [f"{M.MES_TXT[M.MESES_AÑO]} — sigue la tendencia (ritmo {M.ETQ['RANGO'].lower()})", n1(esc["PUNTAJE_DIC_TEND"]), cum(esc["PUNTAJE_DIC_TEND"]),
         f"{s_tend['VERDE']} / {s_tend['AMARILLO']} / {s_tend['ROJO']}"],
        [f"{M.MES_TXT[M.MESES_AÑO]} — mantiene el ritmo de {M.ETQ['MES_LARGO'].lower()}", n1(esc["PUNTAJE_DIC_AGO"]), cum(esc["PUNTAJE_DIC_AGO"]), "—"],
        [f"{M.MES_TXT[M.MESES_AÑO]} — PLAN BASE (exigible)", n1(esc["PUNTAJE_PLAN_BASE"]), cum(esc["PUNTAJE_PLAN_BASE"]),
         f"{s_base['VERDE']} / {s_base['AMARILLO']} / {s_base['ROJO']}"],
        [f"{M.MES_TXT[M.MESES_AÑO]} — PLAN AMPLIADO (techo operativo)", n1(esc["PUNTAJE_PLAN"]), cum(esc["PUNTAJE_PLAN"]),
         f"{s_amp['VERDE']} / {s_amp['AMARILLO']} / {s_amp['ROJO']}"],
    ]
    add_tabla(doc, ["Escenario", "Puntaje promedio de la red", "Cumplimiento oficial", "IPRESS Verde / Amarillo / Rojo"],
              filas, [3.0, 1.2, 1.2, 1.6], ["L", "C", "C", "C"], font=8.5,
              fills=[None, None, None, None, {0: HEX_GREEN_BG}, {0: HEX_GREEN_BG}], bold_cols=(1,), pad=80)
    add_body_p(doc,
               f"Si cada establecimiento sostiene su ritmo actual, la red cerraría el año con ≈{n1(esc['PUNTAJE_DIC_TEND'])} puntos promedio. "
               f"El plan base eleva ese resultado a ≈{n1(esc['PUNTAJE_PLAN_BASE'])} puntos y el techo operativo a ≈{n1(esc['PUNTAJE_PLAN'])}. "
               "La meta de 100 puntos para todos los establecimientos no es exigible en el corto plazo, pero la brecha puede reducirse de manera "
               "significativa si se concentra el esfuerzo según las categorías del plan.",
               bold_prefix="Lectura de los escenarios: ")
    add_figura(doc, figs["proyeccion"], f"Figura 5: Proyección del Puntaje de Logro promedio a {M.MES_TXT[M.MESES_AÑO].lower()} — Red de Salud {disp}")
    add_figura(doc, figs["semaforo_proy"], f"Figura 6: Composición del semáforo proyectada según escenario — Red de Salud {disp}")

    # Hitos de seguimiento
    ps_inter = M.pts_hasta(df, {s: np.minimum(df[f"{s} META"], df[f"{s} PROY_OCT"] + df[f"{s} OBJ_ND"] / M.MESES_PLAN) for s in M.SERV})
    hit = [(f"Fin de {M.ETQ['HASTA'].lower()} (proyección previa)", {s: df[f"{s} PROY_OCT"].sum() for s in M.SERV}, esc["PUNTAJE_PROY_OCT"])]
    if M.MESES_PLAN > 1:
        hito_txt = f"Mes intermedio del plan" if M.MESES_PLAN > 2 else f"Fin de {M.ETQ['PLAN_INI_TXT'].lower()}"
        hit.append((f"{hito_txt} (hito del plan base)",
                    {s: (df[f"{s} PROY_OCT"] + df[f"{s} OBJ_ND"] / M.MESES_PLAN).sum() for s in M.SERV}, ps_inter.mean()))
    hit.append((f"Fin de {M.ETQ['PLAN_FIN_TXT'].lower()} (cierre del plan base)", {s: df[f"{s} BASE_DIC"].sum() for s in M.SERV}, esc["PUNTAJE_PLAN_BASE"]))
    add_heading_2(doc, f"{sec}.3 Hitos de seguimiento mensual de la red (plan base)")
    add_tabla(doc, ["Corte", "TI acumuladas", "TC acumuladas", "TM acumuladas", "Puntaje promedio"],
              [[h[0], n0(h[1]["TI"]), n0(h[1]["TC"]), n0(h[1]["TM"]), n1(h[2])] for h in hit],
              [2.6, 1.1, 1.1, 1.1, 1.1], ["L", "C", "C", "C", "C"], font=8.5, bold_cols=(4,), pad=80)
    add_body_p(doc, "Estos hitos son acumulados de la red (suma de los establecimientos con su tope por meta) y sirven para verificar cada semana "
                    "en https://ficha32-ditel.web.app/ si la red se mantiene sobre la trayectoria prevista.",
               bold_prefix="Uso de los hitos: ")

    # --- 6.4 Meta al alcance
    cerca_mod = df[df["CATEGORIA"].isin(["CUMPLE", "CERCA", "MODERADO"])].copy()
    cerca_mod["_o"] = cerca_mod["CATEGORIA"].map({"CUMPLE": 0, "CERCA": 1, "MODERADO": 2})
    cerca_mod = cerca_mod.sort_values(["_o", "REQ_MES_TOTAL"])
    add_heading_2(doc, f"{sec}.4 IPRESS con la meta anual al alcance ({len(cerca_mod)} establecimientos)")
    if len(cerca_mod):
        add_body_p(doc,
                   f"Estos establecimientos pueden cumplir el 100 % de su meta anual (100 puntos) con la producción indicada en {M.MES_TXT[M.MESES_AÑO - 1].lower()} y {M.MES_TXT[M.MESES_AÑO].lower()}. "
                   "Son la prioridad para elevar el semáforo de la red con el menor esfuerzo. La producción se expresa como atenciones totales "
                   f"a registrar entre el 1 de {M.MES_TXT[M.MESES_AÑO - 1]} y el 31 de {M.MES_TXT[M.MESES_AÑO]} (ya considerando la proyección de setiembre y octubre).",
                   bold_prefix="Meta al alcance: ")
        filas, fills = [], []
        for _, r in cerca_mod.iterrows():
            filas.append([r["NOMBRE"], r["CATEGORIA"], r["PUNTAJE"], r["PUNTAJE_PROY_OCT"], int(r["TI OBJ_ND"]),
                          int(r["TC OBJ_ND"]), int(r["TM OBJ_ND"]), r["PUNTAJE_PLAN_BASE"]])
            fills.append({1: CAT_BG[r["CATEGORIA"]], 7: SEM_BG[r["SEMAFORO_PLAN_BASE"]]})
        add_tabla(doc, ["Establecimiento / IPRESS", "Categoría", f"Pts {M.ETQ['MES']}", f"Pts proy. {M.ETQ['PROY']}", f"TI {M.ETQ['PLAN']}", f"TC {M.ETQ['PLAN']}", f"TM {M.ETQ['PLAN']}",
                        f"Pts {M.MES_ABR[M.MESES_AÑO]} (plan)"],
                  filas, [2.0, 0.9, 0.6, 0.65, 0.6, 0.6, 0.6, 0.75], ["L", "C", "C", "C", "C", "C", "C", "C"], font=7.8, fills=fills,
                  bold_cols=(7,))
    else:
        add_body_p(doc, "En esta red ningún establecimiento se encuentra en las categorías CUMPLE, CERCA o MODERADO; el plan se concentra "
                        "en la ruta escalonada de avance descrita a continuación.", bold_prefix="Meta al alcance: ")

    # --- 6.5 Palanca
    sec_pal = f"{sec}.5"
    add_heading_2(doc, f"{sec_pal} IPRESS palanca: las que más ayudan a subir el escenario de la red")
    pal = df[df["CATEGORIA"] != "CUMPLE"].copy()
    pal["GANA"] = pal["PUNTAJE_PLAN_BASE"] - pal["PUNTAJE_PROY_OCT"]
    pal["AT"] = pal["OBJ_MES_TOTAL"]
    pal["SUBE"] = pal["SEMAFORO_PLAN_BASE"].map(ORDEN) > pal["SEMAFORO_OCT"].map(ORDEN)
    pal_top = pal.sort_values(["GANA"], ascending=False).head(10)
    suben_plan = int(pal["SUBE"].sum())
    add_body_p(doc,
               f"Con el plan base, {suben_plan} establecimientos cambiarían de semáforo hacia un nivel superior entre la proyección de {M.ETQ['HASTA'].lower()} y "
               f"{M.MES_TXT[M.MESES_AÑO].lower()}. A continuación se listan los {len(pal_top)} que más puntaje aportan al promedio de la red.",
               bold_prefix="Efecto multiplicador: ")
    filas, fills = [], []
    for _, r in pal_top.iterrows():
        filas.append([r["NOMBRE"], r["CATEGORIA"], r["PUNTAJE_PROY_OCT"], r["PUNTAJE_PLAN_BASE"], f"+{r['GANA']:.1f}",
                      int(r["TI OBJ_ND"] + r["TC OBJ_ND"] + r["TM OBJ_ND"]), f"{SEM_TXT[r['SEMAFORO_OCT']]} → {SEM_TXT[r['SEMAFORO_PLAN_BASE']]}"])
        fills.append({1: CAT_BG[r["CATEGORIA"]], 6: SEM_BG[r["SEMAFORO_PLAN_BASE"]]})
    add_tabla(doc, ["Establecimiento / IPRESS", "Categoría", f"Pts proy. {M.ETQ['PROY']}", f"Pts {M.MES_ABR[M.MESES_AÑO]} (plan)", "Ganancia", f"Atenciones {M.ETQ['PLAN']}",
                    f"Semáforo {M.ETQ['PROY']} → {M.MES_ABR[M.MESES_AÑO]}"],
              filas, [2.0, 0.9, 0.7, 0.7, 0.7, 0.9, 1.1], ["L", "C", "C", "C", "C", "C", "C"], font=7.8, fills=fills, bold_cols=(4,))
    add_figura(doc, figs["palanca"], f"Figura 7: IPRESS con mayor ganancia de puntaje con el plan base (barra gris: proyección {M.ETQ['HASTA'].lower()}) — Red de Salud {disp}")

    # --- 6.6 Brecha alta
    ba = df[df["CATEGORIA"] == "BRECHA ALTA"].copy()
    ba["_o"] = ba["RUTA"].map({"Ruta a Verde": 0, "Ruta a Amarillo": 1, "Arranque": 2})
    ba = ba.sort_values(["_o", "PUNTAJE"], ascending=[True, False])
    add_heading_2(doc, f"{sec}.6 Ruta escalonada para IPRESS con brecha alta ({len(ba)} establecimientos)")
    if len(ba):
        rc = ba["RUTA"].value_counts()
        add_body_p(doc,
                   f"Estos {len(ba)} establecimientos necesitan una producción en {M.ETQ['PLAN']} mayor a su capacidad mensual actual para cerrar el 100 % de la "
                   f"meta anual. Probablemente no la cumplan este año, pero pueden avanzar de forma sustancial: {rc.get('Ruta a Verde', 0)} tienen "
                   f"una Ruta a Verde (con su capacidad completa alcanzarían ≥60 pts), {rc.get('Ruta a Amarillo', 0)} una Ruta a Amarillo (≥30 pts) y "
                   f"{rc.get('Arranque', 0)} requieren primero un arranque. La tabla indica la producción exigible del plan base (piso de avance) y el "
                   f"puntaje de {M.MES_TXT[M.MESES_AÑO].lower()} bajo el plan base y bajo el techo operativo.",
                   bold_prefix="Estrategia de avance progresivo: ")
        filas, fills = [], []
        for _, r in ba.iterrows():
            filas.append([r["NOMBRE"], r["RUTA"].replace("Ruta a ", "→ "), r["PUNTAJE"], int(r["TI OBJ_ND"]), int(r["TC OBJ_ND"]),
                          int(r["TM OBJ_ND"]), r["PUNTAJE_PLAN_BASE"], r["PUNTAJE_PLAN"]])
            fills.append({1: {"Ruta a Verde": HEX_GREEN_BG, "Ruta a Amarillo": HEX_YELLOW_BG, "Arranque": HEX_RED_BG}[r["RUTA"]],
                          6: SEM_BG[r["SEMAFORO_PLAN_BASE"]], 7: SEM_BG[r["SEMAFORO_PLAN"]]})
        add_tabla(doc, ["Establecimiento / IPRESS", "Ruta", f"Pts {M.ETQ['MES']}", f"TI {M.ETQ['PLAN']}", f"TC {M.ETQ['PLAN']}", f"TM {M.ETQ['PLAN']}",
                        f"Pts {M.MES_ABR[M.MESES_AÑO]} (plan base)", f"Pts {M.MES_ABR[M.MESES_AÑO]} (techo)"],
                  filas, [2.0, 0.8, 0.6, 0.6, 0.6, 0.6, 0.85, 0.75], ["L", "C", "C", "C", "C", "C", "C", "C"], font=7.8, fills=fills,
                  bold_cols=(6,))
    else:
        add_body_p(doc, "Ningún establecimiento de la red presenta brecha alta.", bold_prefix="Estrategia de avance progresivo: ")

    # --- 6.7 Estrategias
    add_heading_2(doc, f"{sec}.7 Estrategias operativas por categoría")
    hosp_txt = hosp_nombre or "el hospital de referencia de la red"
    add_callout_box(
        doc, "PAUTA 1: CERRAR LA META EN LAS IPRESS CERCA Y MODERADO",
        f"• Entregar a cada jefe de establecimiento la cifra exacta de TI, TC y TM por registrar en {M.MES_TXT[M.MESES_AÑO - 1].lower()} y {M.MES_TXT[M.MESES_AÑO].lower()} (tablas {sec}.4 y {sec}.6).\n"
        f"• Programar un cronograma de teleinterconsultas con {hosp_txt} y asignar turnos fijos semanales (idealmente 2 por semana por IPRESS).\n"
        "• Priorizar el servicio con más puntos por atención en cada establecimiento: una sola teleconsulta puede valer hasta 10 puntos cuando la meta anual de TC es 1.\n"
        "• Realizar un corte semanal en https://ficha32-ditel.web.app/ y reforzar de inmediato a quienes no avancen.",
        bg_hex="F8FAFC", border_color="475569")
    add_callout_box(
        doc, "PAUTA 2: ESCALERA DE TRES PASOS PARA LAS IPRESS CON BRECHA ALTA",
        f"• Paso 1 (primeras dos semanas de {M.MES_TXT[M.MESES_AÑO - 1].lower()}): registrar la primera teleconsulta y un bloque semanal de telemonitoreo (enfermería, obstetricia, "
        "psicología o medicina) a pacientes con hipertensión, diabetes, anemia, salud mental y violencia familiar (Anexo 1 de la Ficha 32).\n"
        f"• Paso 2 (mediados de {M.MES_TXT[M.MESES_AÑO - 1].lower()}): coordinar con {hosp_txt} la primera tanda de teleinterconsultas, mediante una sesión de apertura con el hospital "
        "y una lista de pacientes candidatos de la consulta externa.\n"
        f"• Paso 3 ({M.MES_TXT[M.MESES_AÑO].lower()}): consolidar el flujo semanal y ordenar la digitación del mes para que toda la producción figure en el corte final del año.\n"
        "• Meta mínima de avance: alcanzar al menos el piso del plan base indicado en la tabla; si hay capacidad, apuntar al techo operativo.",
        bg_hex="F8FAFC", border_color="475569")
    add_callout_box(
        doc, "PAUTA 3: DIAGNÓSTICO RÁPIDO Y SOPORTE PARA LAS IPRESS SIN PRODUCCIÓN",
        "• Verificar en la primera semana si la inactividad se debe a falta de responsable de Telesalud, conectividad, equipamiento, falta de credenciales en Teleatiendo o desconocimiento de flujos.\n"
        "• Designar un responsable de Telesalud por establecimiento, gestionar sus accesos a Teleatiendo y vincularlo con el gestor de la microred.\n"
        "• Programar una sesión virtual breve de inducción sobre el flujo en Teleatiendo (citas sincrónicas/asincrónicas, interconsultas y telemonitoreo con diagnósticos del Anexo 1) asegurando que registren directamente allí.\n"
        f"• Acompañamiento prioritario de la Coordinación de Telemedicina a las IPRESS que permanezcan en cero al cierre de {M.MES_TXT[M.MESES_AÑO - 1].lower()}.",
        bg_hex="F8FAFC", border_color="475569")
    add_callout_box(
        doc, "PAUTA 4: REGISTRO EXCLUSIVO EN TELEATIENDO (NO DUPLICAR EN HIS MANUAL NI FUA)",
        "• REGISTRAR ÚNICA Y DIRECTAMENTE EN LA PLATAFORMA TELEATIENDO: Las atenciones de telemedicina deben ser ingresadas exclusivamente a través de Teleatiendo (lote TLH). Esta plataforma se encuentra interoperada e integrada automáticamente con el sistema HIS-MINSA y la generación del FUA (SIS), por lo que ya NO es necesario registrar la atención de forma manual en el HIS.\n"
        "• CRÍTICO — EVITAR DUPLICIDAD: Si una atención se registra en Teleatiendo y paralelamente se digita en una hoja HIS física/manual, los algoritmos de control de calidad y validación ministerial detectan duplicidad de registro para el mismo paciente, fecha y código, procediendo a ANULAR O CANCELAR AMBOS REGISTROS. Por tanto, se prohíbe la doble digitación.\n"
        "• MONITOREO Y VERIFICACIÓN CONTINUA: Cada jefatura de IPRESS y responsable de Telesalud debe verificar semanalmente el reflejo de sus atenciones en la plataforma oficial https://ficha32-ditel.web.app/ para garantizar que toda la producción cargada en Teleatiendo sume efectivamente al avance del indicador sin rechazos.",
        bg_hex="F8FAFC", border_color="475569")

    # --- 6.8 Cronograma
    p_ini = M.ETQ['PLAN_INI_TXT'].lower()
    p_fin = M.ETQ['PLAN_FIN_TXT'].lower()
    add_heading_2(doc, f"{sec}.8 Cronograma sugerido y responsables")
    add_tabla(doc, ["Período", "Actividad", "Responsable", "Producto verificable"],
              [[f"1.ª semana de {p_ini}", "Presentar este plan en la Red y en cada Microred; designar responsable de Telesalud por IPRESS; capacitar en el registro exclusivo en Teleatiendo (cero digitación manual en HIS para evitar duplicidad).",
                "Dirección de Red / Jefes de Microred", "Lista de responsables y usuarios Teleatiendo activos"],
               [f"2.ª semana de {p_ini}", "Activar teleconsulta y telemonitoreo directamente en Teleatiendo en IPRESS en arranque; programar cronograma de TI con el hospital consultor.",
                "Responsable de Telesalud de la IPRESS", "Primeras atenciones en Teleatiendo y TI programadas"],
               [f"Semanas del plan ({M.ETQ['PLAN']})", "Ejecución del plan; corte semanal en https://ficha32-ditel.web.app/; verificación de interoperabilidad Teleatiendo-HIS y refuerzo a IPRESS sin avance.",
                "Microred / Coordinación de Telemedicina", f"Hitos de avance (sección {sec}.3)"],
               [f"Avance hacia {p_fin}", "Cierre de brechas de las IPRESS CERCA y MODERADO; consolidar la ruta de avance de las IPRESS de brecha alta mediante Teleatiendo.",
                "Jefaturas de IPRESS", "IPRESS con meta cumplida"],
               [f"Auditoría fin de {p_fin}", "Auditoría de consistencia: confirmar que no existan atenciones duplicadas o rechazadas y validar cierre de citas en Teleatiendo.",
                "Responsable de Telesalud / Estadística", f"Producción de {p_fin} cargada en Teleatiendo"],
               [f"Cierre y balance {M.ETQ['ANIO_TXT']}", "Verificación del corte final oficial (1 de enero al 31 de diciembre) y balance de cumplimiento de la Ficha 32.",
                "Coordinación de Telemedicina / Red", "Informe de cierre del indicador"]],
              [1.3, 3.0, 1.5, 1.2], ["L", "L", "L", "L"], font=8, pad=70)
    add_body_p(doc, "Indicadores de seguimiento semanal: atenciones de TI, TC y TM registradas en la semana; número de IPRESS con producción cero; "
                    "puntaje promedio de la red frente al hito del mes; IPRESS que cambiaron de semáforo.",
               bold_prefix="Seguimiento: ")



def construir_informe_regional(doc, A, B, figs):
    """Genera el Informe Ejecutivo Regional para la Dirección Regional de Salud."""
    total = len(A)
    poblacion = int(A["DENOMINADOR"].sum())
    cnt = {k: int((A["SEMAFORO"] == k).sum()) for k in ("VERDE", "AMARILLO", "ROJO")}
    cnt_ant = {k: int((A["SEMAFORO_ANT"] == k).sum()) for k in ("VERDE", "AMARILLO", "ROJO")}
    mean_ant, mean_act = A["PUNTAJE_ANT"].mean(), A["PUNTAJE"].mean()
    cump_oficial = max(0.0, (mean_act - 10) / 90 * 100)

    # Configuración de página
    for s in doc.sections:
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)
        p_ft = s.footer.paragraphs[0]
        p_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_ft = p_ft.add_run(f"Coordinación Regional de Telemedicina — DIRESA {M.DIRESA.title()} | Balance Regional (Ficha N.° 32 — {M.ETQ['MES']})")
        r_ft.font.name = "Arial"
        r_ft.font.size = Pt(8)
        r_ft.font.color.rgb = COLOR_DARK_GRAY

    # Encabezado institucional
    tbl_hdr = doc.add_table(rows=1, cols=2)
    tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_hdr.autofit = False
    cl, cr = tbl_hdr.cell(0, 0), tbl_hdr.cell(0, 1)
    fijar_anchos(tbl_hdr, [4.6, 2.4])
    set_cell_background(cl, "F1F5F9")
    set_cell_background(cr, "E2E8F0")
    set_cell_margins(cl, 140, 140, 180, 140)
    set_cell_margins(cr, 140, 140, 140, 140)
    pl = cl.paragraphs[0]
    pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(BANNER_PATH):
        pl.add_run().add_picture(BANNER_PATH, width=Inches(4.3))
    else:
        _run(pl, f"GOBIERNO REGIONAL {M.DIRESA}\nDIRECCIÓN REGIONAL DE SALUD {M.DIRESA}\nCoordinación de Telemedicina", 9, bold=True)
    pr = cr.paragraphs[0]
    pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _run(pr, f"INFORME EJECUTIVO REGIONAL\n", 9, bold=True)
    _run(pr, f"Balance de Gestión y Plan de Cumplimiento\nCorte: {M.ETQ['MES']} / Año {M.ETQ['ANIO_TXT']}", 8.5, color=COLOR_DARK_GRAY)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    pt = doc.add_paragraph()
    pt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pt.paragraph_format.space_before = Pt(12)
    pt.paragraph_format.space_after = Pt(4)
    _run(pt, f"INFORME EJECUTIVO REGIONAL DE TELEMEDICINA\nDIRESA {M.DIRESA.upper()} — CORTE {M.ETQ['MES'].upper()} Y PLAN {M.ETQ['PLAN_LARGO'].upper()}", 14, bold=True)
    ps = doc.add_paragraph()
    ps.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ps.paragraph_format.space_after = Pt(14)
    _run(ps, f"Monitoreo Situacional Integral de las {len(M.redes(A))} Redes de Salud y {total} IPRESS del Primer Nivel", 10.5, italic=True, color=COLOR_DARK_GRAY)

    # 1. Resumen Ejecutivo
    add_heading_1(doc, "1. RESUMEN EJECUTIVO REGIONAL")
    etq_ant = f"Mes {M.MES_CORTE - 1}" if M.MES_CORTE > 1 else "Línea base"
    callout_txt = (
        f"• Padrón regional evaluado: {total} establecimientos de primer nivel (población asignada: {n0(poblacion)} habitantes) y {len(B)} hospitales/institutos consultores.\n"
        f"• Puntaje de Logro promedio regional: {mean_act:.1f} puntos sobre 100 (frente a {mean_ant:.1f} en {etq_ant}, variación de {mean_act - mean_ant:+.1f} puntos).\n"
        f"• Cumplimiento oficial del puntaje regional: {cump_oficial:.1f} % según fórmula MINSA ((Puntaje − 10) / 90).\n"
        f"• Semáforo regional: {cnt['VERDE']} IPRESS en Verde ({pct(cnt['VERDE'], total):.1f} %), {cnt['AMARILLO']} en Amarillo ({pct(cnt['AMARILLO'], total):.1f} %) y {cnt['ROJO']} en Rojo ({pct(cnt['ROJO'], total):.1f} %).\n"
        f"• Brecha hacia fin de año: Con el plan base propuesto para {M.ETQ['PLAN_LARGO'].lower()}, la región alcanzaría un promedio de {A['PUNTAJE_PLAN_BASE'].mean():.1f} puntos, logrando {int((A['SEMAFORO_PLAN_BASE'] == 'VERDE').sum())} IPRESS en Verde."
    )
    add_callout_box(doc, f"DIRESA {M.DIRESA.title()}: INDICADORES CLAVE ({M.ETQ['MES']})", callout_txt)
    add_tabla(doc, ["Semáforo de Gestión", f"N.° IPRESS {etq_ant}", f"N.° IPRESS {M.ETQ['MES']}", f"% IPRESS {M.ETQ['MES']}", f"Proy. Plan {M.ETQ['PLAN']}"],
              [["Verde (≥ 60 pts)", str(cnt_ant["VERDE"]), str(cnt["VERDE"]), f"{pct(cnt['VERDE'], total):.1f} %", str(int((A['SEMAFORO_PLAN_BASE'] == 'VERDE').sum()))],
               ["Amarillo (30–59.9 pts)", str(cnt_ant["AMARILLO"]), str(cnt["AMARILLO"]), f"{pct(cnt['AMARILLO'], total):.1f} %", str(int((A['SEMAFORO_PLAN_BASE'] == 'AMARILLO').sum()))],
               ["Rojo (< 30 pts)", str(cnt_ant["ROJO"]), str(cnt["ROJO"]), f"{pct(cnt['ROJO'], total):.1f} %", str(int((A['SEMAFORO_PLAN_BASE'] == 'ROJO').sum()))]],
              [2.2, 1.2, 1.2, 1.2, 1.2], ["L", "C", "C", "C", "C"], font=8.5,
              fills=[{0: HEX_GREEN_BG}, {0: HEX_YELLOW_BG}, {0: HEX_RED_BG}], bold_cols=(0, 2))
    if "semaforo" in figs:
        add_figura(doc, figs["semaforo"], f"Figura 1: Semáforo comparativo regional ({etq_ant} vs {M.ETQ['MES']}) — DIRESA {M.DIRESA.title()}")

    # 2. Desempeño por Red de Salud
    add_heading_1(doc, "2. EVALUACIÓN Y RANKING POR RED DE SALUD")
    filas_redes = []
    fills_redes = []
    for prov, slug, disp in M.redes(A):
        sub = A[A["RED_SALUD"] == prov]
        p_ant = sub["PUNTAJE_ANT"].mean()
        p_act = sub["PUNTAJE"].mean()
        d_p = p_act - p_ant
        sem_act = M.clasificar(p_act)
        ti_pct = sub["TI REALIZADAS"].sum() / max(sub["TI META"].sum(), 1) * 100
        tc_pct = sub["TC REALIZADAS"].sum() / max(sub["TC META"].sum(), 1) * 100
        tm_pct = sub["TM REALIZADAS"].sum() / max(sub["TM META"].sum(), 1) * 100
        p_plan = sub["PUNTAJE_PLAN_BASE"].mean()
        filas_redes.append([disp, str(len(sub)), f"{p_ant:.1f}", f"{p_act:.1f}", f"{d_p:+.1f}",
                           f"{ti_pct:.0f}%", f"{tc_pct:.0f}%", f"{tm_pct:.0f}%", f"{p_plan:.1f}"])
        fills_redes.append({3: SEM_BG[sem_act]})

    add_tabla(doc, ["Red de Salud", "IPRESS", f"Pts {etq_ant}", f"Pts {M.ETQ['MES']}", "Δ Pts", "% TI", "% TC", "% TM", f"Plan {M.ETQ['PLAN']}"],
              filas_redes, [1.9, 0.6, 0.7, 0.7, 0.7, 0.6, 0.6, 0.6, 0.8],
              ["L", "C", "C", "C", "C", "C", "C", "C", "C"], font=8, fills=fills_redes, bold_cols=(0, 3, 8), pad=60)
    if "redes" in figs and figs["redes"]:
        add_figura(doc, figs["redes"], f"Figura 2: Comparativo del Puntaje de Logro promedio por Red de Salud — DIRESA {M.DIRESA.title()}")

    # 3. Avance por Servicio
    add_heading_1(doc, "3. AVANCE REGIONAL POR SERVICIO (TI, TC, TM)")
    add_body_p(doc, f"El indicador ministerial evalúa de forma diferenciada tres prestaciones con pesos ponderados específicos: "
                    f"Teleinterconsulta (TI, peso {M.PESO['TI']} %), Teleconsulta (TC, peso {M.PESO['TC']} %) y Telemonitoreo (TM, peso {M.PESO['TM']} %).")
    sum_serv = {s: (int(A[f"{s} REALIZADAS"].sum()), int(A[f"{s} META"].sum())) for s in M.SERV}
    add_tabla(doc, ["Servicio", "Peso", f"Producción {M.ETQ['MES']}", "Meta Anual Regional", "% de la Meta", f"Esperado a {M.ETQ['MES_LARGO'].lower()}"],
              [[f"{M.NOMBRE_SERV[s]} ({s})", f"{M.PESO[s]} %", n0(sum_serv[s][0]), n0(sum_serv[s][1]),
                f"{pct(sum_serv[s][0], sum_serv[s][1]):.1f} %", f"{(M.MES_CORTE / M.MESES_AÑO * 100):.1f} %"] for s in M.SERV],
              [2.2, 0.8, 1.1, 1.1, 1.0, 1.0], ["L", "C", "C", "C", "C", "C"], font=8.5, bold_cols=(0, 4))
    if "avance" in figs:
        add_figura(doc, figs["avance"], f"Figura 3: Avance acumulado vs meta anual regional por servicio — DIRESA {M.DIRESA.title()}")

    # 4. Hospitales Consultores (Reporte B)
    if len(B):
        add_heading_1(doc, "4. EVALUACIÓN DE HOSPITALES E INSTITUTOS CONSULTORES (LOGRO B)")
        add_body_p(doc, "Los hospitales e institutos de segundo y tercer nivel tienen como meta el 100 % de su indicador basado en Teleinterconsultas realizadas. "
                        "Su meta anual es dinámica y crece con sus atenciones HIS acumuladas.")
        filas_b = []
        fills_b = []
        for _, h in B.iterrows():
            filas_b.append([h["NOMBRE"], h["NIVEL"], h["RED_SALUD"], n0(h["DENOMINADOR"]), n0(h["TI META"]),
                            n0(h["TI REALIZADAS"]), f"{h['PUNTAJE']:.1f}", n0(h["META_PROY_DIC"])])
            fills_b.append({6: SEM_BG[h["SEMAFORO"]]})
        add_tabla(doc, ["Hospital / Instituto", "Cat.", "Red / Prov.", "Atenciones HIS", f"Meta TI {M.ETQ['MES']}", f"TI Realiz.", "Puntaje", f"Meta Proy. {M.MES_ABR[M.MESES_AÑO]}"],
                  filas_b, [2.3, 0.5, 1.2, 0.9, 0.7, 0.7, 0.7, 0.8],
                  ["L", "C", "L", "C", "C", "C", "C", "C"], font=8, fills=fills_b, bold_cols=(0, 6), pad=60)

    # 5. Plan de Cumplimiento Regional
    add_heading_1(doc, f"5. PLAN DE CUMPLIMIENTO REGIONAL {M.ETQ['PLAN_LARGO'].upper()} {M.ETQ['ANIO_TXT']}")
    catn = {c: int((A["CATEGORIA"] == c).sum()) for c in M.CATEGORIAS}
    total_req_nd = int(A[['TI OBJ_ND', 'TC OBJ_ND', 'TM OBJ_ND']].sum().sum())
    callout_plan = (
        f"• Capacidad operativa mensual estimada para la DIRESA: {int(A['CAP_MES'].sum()):,} atenciones/mes entre las {total} IPRESS.\n"
        f"• Meta de producción consolidada para {M.ETQ['PLAN']}: {total_req_nd:,} atenciones en total (plan base) para asegurar la máxima cantidad de IPRESS en Verde.\n"
        f"• IPRESS priorizadas: {catn['CERCA']} IPRESS en condición CERCA y {catn['MODERADO']} en MODERADO que requieren acompañamiento directo para cerrar su meta anual."
    )
    add_callout_box(doc, f"COMPROMISO OPERATIVO REGIONAL: {M.ETQ['PLAN_LARGO'].upper()}", callout_plan)
    add_tabla(doc, ["Categoría", "Criterio de Esfuerzo", "N.° IPRESS", "% Región", f"Atenciones Plan {M.ETQ['PLAN']}"],
              [[c, desc, str(catn[c]), f"{pct(catn[c], total):.1f} %",
                n0(A.loc[A['CATEGORIA'] == c, ['TI OBJ_ND', 'TC OBJ_ND', 'TM OBJ_ND']].sum().sum())]
               for c, desc in zip(M.CATEGORIAS, ["Meta anual asegurada (brecha 0)", "Requerimiento ≤ 1/3 de su capacidad mensual",
                                                 "Requerimiento entre 1/3 y 1 vez su capacidad mensual",
                                                 "Brecha alta: ruta escalonada de avance"])],
              [1.3, 2.7, 0.8, 0.8, 1.4], ["C", "L", "C", "C", "C"], font=8.3,
              fills=[{0: CAT_BG[c]} for c in M.CATEGORIAS], bold_cols=(0, 2, 4), pad=70)

    add_callout_box(
        doc, "DIRECTRIZ CRÍTICA DE REGISTRO MINISTERIAL (TOLERANCIA CERO A DUPLICIDAD)",
        "1. REGISTRO EXCLUSIVO EN TELEATIENDO: Cada atención de telemedicina debe ser ingresada únicamente en la plataforma Teleatiendo (lote TLH), la cual interopera automáticamente con el sistema HIS-MINSA y genera el FUA para el SIS.\n"
        "2. PROHIBICIÓN TOTAL DE DUPLICIDAD EN HOJAS FÍSICAS: Si el personal registra la atención en Teleatiendo y paralelamente la anota en hojas HIS manuales, el sistema ministerial detecta duplicidad de registro y ANULA AUTOMÁTICAMENTE AMBOS REGISTROS, eliminando la atención del cómputo oficial.\n"
        "3. MONITOREO SEMANAL OBLIGATORIO: Las Direcciones de Red y Coordinaciones de Telemedicina deben auditar semanalmente la plataforma oficial https://ficha32-ditel.web.app/ para corroborar que la producción ingresada figure sin observaciones.",
        bg_hex="F8FAFC", border_color="475569"
    )


def main(solo=None, incluir_regional=True):
    A, B = M.construir()
    redes = M.redes(A)
    archivos = []

    # 1. Informe Regional Ejecutivo (si no se filtró por red individual y está habilitado)
    es_filtro_red = bool(M.FILTRO.get("red") or (solo and not any("REGIONAL" in str(s).upper() for s in solo)))
    if incluir_regional and not es_filtro_red:
        print(f"Generando Informe Regional: DIRESA {M.DIRESA} ...")
        figs_reg = G.graficos_red(A, f"DIRESA {M.DIRESA.title()} ({len(A)} IPRESS)", "REGIONAL")
        try:
            figs_reg["redes"] = G.fig_redes(A)
        except Exception:
            figs_reg["redes"] = None
        doc_reg = Document()
        construir_informe_regional(doc_reg, A, B, figs_reg)
        out_reg = os.path.join(OUT, f"Reporte_Ficha32_REGIONAL_Mes{M.MES_CORTE}.docx")
        try:
            doc_reg.save(out_reg)
        except PermissionError:
            out_reg = out_reg.replace(".docx", "_nuevo.docx")
            doc_reg.save(out_reg)
        print(f"  OK: {out_reg} (DIRESA {M.DIRESA} — {len(A)} IPRESS)")
        archivos.append(out_reg)

    # 2. Informes por Red de Salud
    for prov, slug, disp in redes:
        if solo and slug not in solo and prov not in solo:
            continue
        print(f"Generando informe: {disp} ...")
        df = A[A["RED_SALUD"] == prov].copy()
        dfb = B[B["RED_SALUD"] == prov].copy()
        figs = G.graficos_red(df, f"Red de Salud {disp}", slug)
        doc = Document()
        construir_informe(doc, prov, disp, slug, df, dfb if len(dfb) else None, figs)
        out = os.path.join(OUT, f"Reporte_Ficha32_{slug}_Mes{M.MES_CORTE}.docx")
        try:
            doc.save(out)
        except PermissionError:
            out = out.replace(".docx", "_nuevo.docx")
            doc.save(out)
        print(f"  OK: {out} ({len(df)} IPRESS)")
        archivos.append(out)

    return archivos


if __name__ == "__main__":
    main(sys.argv[1:] or None)


