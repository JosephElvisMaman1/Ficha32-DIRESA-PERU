# -*- coding: utf-8 -*-
"""Libro Excel — Ficha N.° 32 (Telemedicina): datos, fórmulas, gráficos nativos y plan de cumplimiento.

El periodo (mes de corte, año) y la DIRESA activa se toman de ficha32_modelo (M.ETQ / M.DIRESA).
"""
import os
import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as COL

import ficha32_modelo as M

OUT = os.path.join(M.BASE, f"Ficha32_Telemedicina_Mes{M.MES_CORTE}_{M.ANIO}.xlsx")
ESPERADO = M.MES_CORTE / M.MESES_AÑO

SEM_COL = {"VERDE": ("#DCFCE7", "#166534"), "AMARILLO": ("#FEF3C7", "#92400E"), "ROJO": ("#FEE2E2", "#991B1B")}
CAT_COL = {"CUMPLE": "#DCFCE7", "CERCA": "#ECFCCB", "MODERADO": "#FEF3C7", "BRECHA ALTA": "#FEE2E2"}


class _Omitida:
    """Hoja ficticia: absorbe todas las escrituras.

    Permite omitir una parte del libro (Plan, Palanca, Supuestos, Hospitales…)
    sin tener que recorrer y condicionar cada línea que la escribe: las llamadas
    write/autofilter/format no hacen nada y el gráfico que se «inserte» en ella
    se descarta al guardar.
    """

    def __getattr__(self, _nombre):
        return lambda *a, **k: None


def build(serie=None, plan=True, hospitales=True, redes=True, evo_desde=None, out_path=None, cartera=True):
    """Genera el libro Excel. `serie` (opcional) es {mes: (A, B, ctx)} para la hoja Evolución.

    Interruptores de contenido (todos activos por defecto):
      * plan       → parte de planificación: hojas Plan, Palanca y Supuestos y las
                     columnas/escenarios de futuro de Resumen, Redes, Hospitales y
                     las hojas por red. Sin él solo queda la situación actual.
      * hospitales → la hoja Hospitales (Reporte B).
      * redes      → la hoja Redes (resumen por red) y las hojas Red_* (detalle por red).
      * evo_desde  → primer mes mostrado en la hoja Evolución (por defecto: mes 1).
      * out_path   → ruta de destino del archivo (opcional).
    """
    if isinstance(serie, str) and out_path is None:
        out_path = serie
        serie = None
    A, B = M.construir()
    A = A.sort_values(["RED_SALUD", "PUNTAJE", "PUNTAJE_ANT"], ascending=[True, False, False]).reset_index(drop=True)
    def _resolver_ruta_segura(p):
        if not os.path.exists(p):
            return p
        try:
            with open(p, "a+b"):
                pass
            return p
        except Exception:
            import time
            b, e = os.path.splitext(p)
            alt = f"{b}_{time.strftime('%H%M%S')}{e}"
            print(f"[aviso] El archivo '{os.path.basename(p)}' está en uso (abierto en Excel). Se guardará en: '{os.path.basename(alt)}'")
            return alt

    ruta_salida = _resolver_ruta_segura(out_path or OUT)
    os.makedirs(os.path.dirname(os.path.abspath(ruta_salida)), exist_ok=True)
    wb = xlsxwriter.Workbook(ruta_salida)
    wb.set_calc_mode("auto")
    F = {}
    base = {"font_name": "Arial", "font_size": 9}
    F["t"] = wb.add_format({**base, "bold": True, "font_size": 15, "font_color": "#0F172A"})
    F["st"] = wb.add_format({**base, "italic": True, "font_color": "#475569"})
    F["h"] = wb.add_format({**base, "bold": True, "bg_color": "#1E3A5F", "font_color": "#FFFFFF", "border": 1,
                            "border_color": "#94A3B8", "align": "center", "valign": "vcenter", "text_wrap": True})
    F["h2"] = wb.add_format({**base, "bold": True, "bg_color": "#E2E8F0", "border": 1, "border_color": "#94A3B8",
                             "align": "center", "valign": "vcenter", "text_wrap": True})
    F["txt"] = wb.add_format({**base, "border": 1, "border_color": "#CBD5E1"})
    F["rol_consultor"] = wb.add_format({**base, "bold": True, "font_color": "#1E40AF", "bg_color": "#DBEAFE", "align": "center", "border": 1, "border_color": "#CBD5E1"})
    F["rol_consultante"] = wb.add_format({**base, "bold": True, "font_color": "#065F46", "bg_color": "#D1FAE5", "align": "center", "border": 1, "border_color": "#CBD5E1"})
    F["rol_ambos"] = wb.add_format({**base, "bold": True, "font_color": "#5B21B6", "bg_color": "#EDE9FE", "align": "center", "border": 1, "border_color": "#CBD5E1"})
    F["rol_ninguno"] = wb.add_format({**base, "font_color": "#64748B", "align": "center", "border": 1, "border_color": "#CBD5E1"})
    F["si_conv"] = wb.add_format({**base, "bold": True, "font_color": "#166534", "bg_color": "#DCFCE7", "align": "center", "border": 1, "border_color": "#CBD5E1"})
    F["no_conv"] = wb.add_format({**base, "font_color": "#64748B", "align": "center", "border": 1, "border_color": "#CBD5E1"})
    F["txtw"] = wb.add_format({**base, "border": 1, "border_color": "#CBD5E1", "text_wrap": True, "valign": "top"})
    F["c"] = wb.add_format({**base, "border": 1, "border_color": "#CBD5E1", "align": "center"})
    F["int"] = wb.add_format({**base, "border": 1, "border_color": "#CBD5E1", "num_format": "#,##0", "align": "center"})
    F["d1"] = wb.add_format({**base, "border": 1, "border_color": "#CBD5E1", "num_format": "0.0", "align": "center"})
    F["d2"] = wb.add_format({**base, "border": 1, "border_color": "#CBD5E1", "num_format": "0.00", "align": "center"})
    F["pct"] = wb.add_format({**base, "border": 1, "border_color": "#CBD5E1", "num_format": "0.0%", "align": "center"})
    F["dlt"] = wb.add_format({**base, "border": 1, "border_color": "#CBD5E1", "num_format": "+0.0;-0.0;0.0", "align": "center"})
    F["b"] = wb.add_format({**base, "bold": True})
    F["bd1"] = wb.add_format({**base, "bold": True, "border": 1, "border_color": "#CBD5E1", "num_format": "0.0",
                              "align": "center", "bg_color": "#F1F5F9"})
    F["bint"] = wb.add_format({**base, "bold": True, "border": 1, "border_color": "#CBD5E1", "num_format": "#,##0",
                               "align": "center", "bg_color": "#F1F5F9"})
    F["bpct"] = wb.add_format({**base, "bold": True, "border": 1, "border_color": "#CBD5E1", "num_format": "0.0%",
                               "align": "center", "bg_color": "#F1F5F9"})
    F["btxt"] = wb.add_format({**base, "bold": True, "border": 1, "border_color": "#CBD5E1", "bg_color": "#F1F5F9"})
    F["note"] = wb.add_format({**base, "font_color": "#475569", "text_wrap": True, "valign": "top"})
    F["kpi_l"] = wb.add_format({**base, "bold": True, "bg_color": "#F1F5F9", "border": 1, "border_color": "#CBD5E1"})
    F["kpi_v"] = wb.add_format({**base, "bold": True, "font_size": 11, "border": 1, "border_color": "#CBD5E1",
                                "align": "center", "num_format": "#,##0.0"})
    F["kpi_i"] = wb.add_format({**base, "bold": True, "font_size": 11, "border": 1, "border_color": "#CBD5E1",
                                "align": "center", "num_format": "#,##0"})
    F["kpi_p"] = wb.add_format({**base, "bold": True, "font_size": 11, "border": 1, "border_color": "#CBD5E1",
                                "align": "center", "num_format": "0.0%"})
    F["subio"] = wb.add_format({"font_name": "Calibri", "font_size": 9, "bold": True, "bg_color": "#DCFCE7",
                                "font_color": "#166534", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#E2E8F0"})
    F["nosubio"] = wb.add_format({"font_name": "Calibri", "font_size": 9, "bold": False, "bg_color": "#FEF3C7",
                                  "font_color": "#92400E", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#E2E8F0"})
    sem_fmt = {k: wb.add_format({"bg_color": v[0], "font_color": v[1], "bold": True}) for k, v in SEM_COL.items()}
    cat_fmt = {k: wb.add_format({"bg_color": v}) for k, v in CAT_COL.items()}

    def cf_sem(ws, rng):
        for k, f in sem_fmt.items():
            ws.conditional_format(rng, {"type": "cell", "criteria": "==", "value": f'"{k}"', "format": f})

    def cf_cat(ws, rng):
        for k, f in cat_fmt.items():
            ws.conditional_format(rng, {"type": "cell", "criteria": "==", "value": f'"{k}"', "format": f})

    # ---------------------------------------------------------------- orden de hojas
    ws_res = wb.add_worksheet("Resumen")
    ws_red = wb.add_worksheet("Redes") if redes else _Omitida()
    ws_evo = wb.add_worksheet("Evolución") if serie else None
    ws_ip = wb.add_worksheet("IPRESS")
    ws_pl = wb.add_worksheet(f"Plan {M.ETQ['PLAN'].replace('–', '-')}") if plan else _Omitida()
    ws_pa = wb.add_worksheet("Palanca") if plan else _Omitida()
    ws_ho = wb.add_worksheet("Hospitales") if hospitales and len(B) else _Omitida()
    ws_su = wb.add_worksheet("Supuestos") if plan else _Omitida()
    for w, c in ((ws_res, "#1E3A5F"), (ws_red, "#1E3A5F"), (ws_evo, "#0EA5E9"), (ws_ip, "#0EA5E9"), (ws_pl, "#10B981"),
                 (ws_pa, "#10B981"), (ws_ho, "#8B5CF6"), (ws_su, "#64748B")):
        if w is not None:
            w.set_tab_color(c)

    n = len(A)
    last = n  # última fila de datos (1-based) en IPRESS / Plan  (cabecera en fila 1)

    # ================================================================= IPRESS
    ip_cols = [("Red de Salud", 20), ("Cat.", 6), ("Código", 10), ("Establecimiento / IPRESS", 42), ("Población", 10),
               (f"TI acum. Mes {M.MES_CORTE - 1}", 8), (f"TI acum. {M.ETQ['MES']}", 8), ("TI Meta anual", 8),
               (f"TC acum. Mes {M.MES_CORTE - 1}", 8), (f"TC acum. {M.ETQ['MES']}", 8), ("TC Meta anual", 8),
               (f"TM acum. Mes {M.MES_CORTE - 1}", 8), (f"TM acum. {M.ETQ['MES']}", 8), ("TM Meta anual", 8),
               (f"Puntaje Mes {M.MES_CORTE - 1}", 9), (f"Puntaje {M.ETQ['MES']}", 9), ("Δ Puntaje", 8), (f"Cumplim. oficial {M.ETQ['MES']}", 10),
               (f"Semáforo Mes {M.MES_CORTE - 1}", 11), (f"Semáforo {M.ETQ['MES']}", 11), ("% meta TI", 8), ("% meta TC", 8), ("% meta TM", 8),
               (f"Prod. {M.ETQ['MES']} TI", 8), (f"Prod. {M.ETQ['MES']} TC", 8), (f"Prod. {M.ETQ['MES']} TM", 8),
               ("Rol Telesalud (SUSALUD)", 18), ("Nro. Resolución Cartera", 28), ("Fecha Res. Cartera", 14), ("¿Subió Res.?", 12), ("Cat. SUSALUD (RENIPRESS)", 16), ("¿Cambió Categoría?", 20)]
    for j, (h, w) in enumerate(ip_cols):
        ws_ip.write(0, j, h, F["h"])
        ws_ip.set_column(j, j, w)
    ws_ip.set_row(0, 42)
    for i, r in A.iterrows():
        x = i + 1
        e = x + 1  # fila excel
        ws_ip.write(x, 0, r["RED_SALUD"], F["txt"])
        ws_ip.write(x, 1, r["NIVEL"], F["c"])
        ws_ip.write(x, 2, r["CODIGO"], F["c"])
        ws_ip.write(x, 3, r["NOMBRE"], F["txt"])
        ws_ip.write_number(x, 4, r["DENOMINADOR"], F["int"])
        vals = [r["TI ANT"], r["TI REALIZADAS"], r["TI META"], r["TC ANT"], r["TC REALIZADAS"], r["TC META"],
                r["TM ANT"], r["TM REALIZADAS"], r["TM META"]]
        for k, v in enumerate(vals):
            ws_ip.write_number(x, 5 + k, v, F["int"])
        ws_ip.write_formula(x, 14, f"=ROUND(60*MIN(F{e},H{e})/H{e}+10*MIN(I{e},K{e})/K{e}+30*MIN(L{e},N{e})/N{e},1)",
                            F["d1"], r["PUNTAJE_ANT"])
        ws_ip.write_formula(x, 15, f"=ROUND(60*MIN(G{e},H{e})/H{e}+10*MIN(J{e},K{e})/K{e}+30*MIN(M{e},N{e})/N{e},1)",
                            F["d1"], r["PUNTAJE"])
        ws_ip.write_formula(x, 16, f"=P{e}-O{e}", F["dlt"], round(r["PUNTAJE"] - r["PUNTAJE_ANT"], 1))
        ws_ip.write_formula(x, 17, f"=IF(P{e}>10,(P{e}-10)/90,0)", F["pct"], r["CUMPL_OFICIAL"] / 100)
        ws_ip.write_formula(x, 18, f'=IF(O{e}>=60,"VERDE",IF(O{e}>=30,"AMARILLO","ROJO"))', F["c"], r["SEMAFORO_ANT"])
        ws_ip.write_formula(x, 19, f'=IF(P{e}>=60,"VERDE",IF(P{e}>=30,"AMARILLO","ROJO"))', F["c"], r["SEMAFORO"])
        ws_ip.write_formula(x, 20, f"=G{e}/H{e}", F["pct"], r["TI REALIZADAS"] / r["TI META"])
        ws_ip.write_formula(x, 21, f"=J{e}/K{e}", F["pct"], r["TC REALIZADAS"] / r["TC META"])
        ws_ip.write_formula(x, 22, f"=M{e}/N{e}", F["pct"], r["TM REALIZADAS"] / r["TM META"])
        ws_ip.write_formula(x, 23, f"=G{e}-F{e}", F["int"], r[f"TI MES{M.MES_CORTE}"])
        ws_ip.write_formula(x, 24, f"=J{e}-I{e}", F["int"], r[f"TC MES{M.MES_CORTE}"])
        ws_ip.write_formula(x, 25, f"=M{e}-L{e}", F["int"], r[f"TM MES{M.MES_CORTE}"])
        ws_ip.write(x, 26, r.get("TELESALUD_ROL", "NO REGISTRADO"), F["c"])
        ws_ip.write(x, 27, r.get("NRO_RESOLUCION", "—"), F["c"])
        ws_ip.write(x, 28, r.get("FECHA_RESOLUCION", "—"), F["c"])
        sub_var = r.get("VARIACION_RESOLUCION", "NO SUBIÓ")
        ws_ip.write(x, 29, sub_var, F["c"])
        cat_su = r.get("CATEGORIA_SUSALUD") or "—"
        est_cat = r.get("ESTADO_CATEGORIA") or "—"
        ws_ip.write(x, 30, cat_su, F["c"])
        fmt_ip_est = F["subio"] if est_cat == "✓ IGUAL" else (F["nosubio"] if str(est_cat).startswith("⚠") else F["c"])
        ws_ip.write(x, 31, est_cat, fmt_ip_est)
    ws_ip.freeze_panes(1, 4)
    ws_ip.autofilter(0, 0, n, len(ip_cols) - 1)
    cf_sem(ws_ip, f"S2:T{last + 1}")
    ws_ip.conditional_format(f"P2:P{last + 1}", {"type": "data_bar", "bar_color": "#60A5FA", "bar_solid": True,
                                                 "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 100})
    ws_ip.conditional_format(f"Q2:Q{last + 1}", {"type": "cell", "criteria": ">", "value": 0,
                                                 "format": wb.add_format({"font_color": "#166534", "bold": True})})
    rng_ip = lambda col: f"IPRESS!${col}$2:${col}${last + 1}"  # noqa

    # ================================================================= PLAN
    pl_cols = [("Red de Salud", 20), ("Cat.", 6), ("Establecimiento / IPRESS", 42), (f"Puntaje {M.ETQ['MES']}", 9),
               (f"Ritmo TI /mes ({M.ETQ['RANGO'].lower()})", 9), ("Ritmo TC /mes", 9), ("Ritmo TM /mes", 9),
               (f"Proy. fin {M.ETQ['PROY']} TI", 9), (f"Proy. fin {M.ETQ['PROY']} TC", 9), (f"Proy. fin {M.ETQ['PROY']} TM", 9), (f"Puntaje proy. {M.ETQ['PROY']}", 9),
               (f"Brecha {M.ETQ['PLAN']} TI", 9), (f"Brecha {M.ETQ['PLAN']} TC", 9), (f"Brecha {M.ETQ['PLAN']} TM", 9),
               ("Req. total /mes (100 %)", 10), ("Capacidad realista /mes", 10), ("Esfuerzo (req ÷ cap)", 9),
               ("Categoría", 13), ("Ruta", 15),
               (f"PLAN BASE TI {M.ETQ['PLAN']}", 10), (f"PLAN BASE TC {M.ETQ['PLAN']}", 10), (f"PLAN BASE TM {M.ETQ['PLAN']}", 10),
               (f"Puntaje {M.MES_ABR[M.MESES_AÑO]} (plan base)", 10), (f"Semáforo {M.MES_ABR[M.MESES_AÑO]} (plan base)", 11), (f"Puntaje {M.MES_ABR[M.MESES_AÑO]} (techo operativo)", 11),
               (f"Puntaje {M.MES_ABR[M.MESES_AÑO]} (tendencia)", 10), (f"Puntaje {M.MES_ABR[M.MESES_AÑO]} (ritmo {M.ETQ['MES_LARGO'].lower()})", 10), ("Acción sugerida", 90)]
    for j, (h, w) in enumerate(pl_cols):
        ws_pl.write(0, j, h, F["h"])
        ws_pl.set_column(j, j, w)
    ws_pl.set_row(0, 54)
    for i, r in A.iterrows():
        x = i + 1
        e = x + 1
        row = [r["RED_SALUD"], r["NIVEL"], r["NOMBRE"]]
        for k, v in enumerate(row):
            ws_pl.write(x, k, v, F["txt"] if k != 1 else F["c"])
        ws_pl.write_number(x, 3, r["PUNTAJE"], F["d1"])
        for k, s in enumerate(M.SERV):
            ws_pl.write_number(x, 4 + k, round(r[f"{s} RITMO_PROM"], 2), F["d2"])
            ws_pl.write_number(x, 7 + k, r[f"{s} PROY_OCT"], F["int"])
            ws_pl.write_number(x, 11 + k, r[f"{s} BRECHA_ND"], F["int"])
            ws_pl.write_number(x, 19 + k, r[f"{s} OBJ_ND"], F["int"])
        ws_pl.write_number(x, 10, r["PUNTAJE_PROY_OCT"], F["d1"])
        ws_pl.write_formula(x, 14, f"=ROUNDUP(SUM(L{e}:N{e})/{M.MESES_PLAN},0)", F["int"], r["REQ_MES_TOTAL"])
        ws_pl.write_number(x, 15, r["CAP_MES"], F["int"])
        ws_pl.write_formula(x, 16, f"=O{e}/P{e}", F["d2"], r["RATIO_ESFUERZO"])
        ws_pl.write(x, 17, r["CATEGORIA"], F["c"])
        ws_pl.write(x, 18, r["RUTA"], F["c"])
        ws_pl.write_number(x, 22, r["PUNTAJE_PLAN_BASE"], F["d1"])
        ws_pl.write(x, 23, r["SEMAFORO_PLAN_BASE"], F["c"])
        ws_pl.write_number(x, 24, r["PUNTAJE_PLAN"], F["d1"])
        ws_pl.write_number(x, 25, r["PUNTAJE_DIC_TEND"], F["d1"])
        ws_pl.write_number(x, 26, r["PUNTAJE_DIC_AGO"], F["d1"])
        ws_pl.write(x, 27, r["ACCION"], F["txt"])
    ws_pl.freeze_panes(1, 3)
    ws_pl.autofilter(0, 0, n, len(pl_cols) - 1)
    cf_sem(ws_pl, f"X2:X{last + 1}")
    cf_cat(ws_pl, f"R2:R{last + 1}")
    rng_pl = lambda col: f"'Plan {M.ETQ['PLAN'].replace('–', '-')}'!${col}$2:${col}${last + 1}"  # noqa

    # ================================================================= REDES
    tit_red = f"Comparativo por Red de Salud — Ficha N.° 32 — Mes {M.MES_CORTE - 1} vs {M.ETQ['MES']}"
    sub_red = ("Fórmulas vinculadas a la hoja IPRESS. Puntaje promedio = media simple del "
               "Puntaje de Logro de las IPRESS de primer nivel.")
    if plan:
        tit_red += f" y proyección a {M.MES_TXT[M.MESES_AÑO].lower()}"
        sub_red = (f"Fórmulas vinculadas a las hojas IPRESS y Plan {M.ETQ['PLAN'].replace('–', '-')}. "
                   "Puntaje promedio = media simple del Puntaje de Logro de las IPRESS de primer nivel.")
    ws_red.write(0, 0, tit_red, F["t"])
    ws_red.write(1, 0, sub_red, F["st"])
    rh = ["Red de Salud", "N.° IPRESS", f"Pts prom. Mes {M.MES_CORTE - 1}", f"Pts prom. {M.ETQ['MES']}", "Δ Pts", f"Verde Mes {M.MES_CORTE - 1}", f"Amarillo Mes {M.MES_CORTE - 1}", f"Rojo Mes {M.MES_CORTE - 1}",
          f"Verde {M.ETQ['MES']}", f"Amarillo {M.ETQ['MES']}", f"Rojo {M.ETQ['MES']}", "% meta TI", "% meta TC", "% meta TM",
          f"Pts {M.MES_ABR[M.MESES_AÑO]} (tendencia)", f"Pts {M.MES_ABR[M.MESES_AÑO]} (plan base)", f"Pts {M.MES_ABR[M.MESES_AÑO]} (techo)", "CUMPLE", "CERCA", "MODERADO", "BRECHA ALTA",
          f"Verde {M.MES_ABR[M.MESES_AÑO]} (plan base)"]
    if not plan:
        rh = rh[:14]            # solo situación actual: sin columnas de planificación
    for j, h in enumerate(rh):
        ws_red.write(3, j, h, F["h"])
    ws_red.set_row(3, 40)
    ws_red.set_column(0, 0, 24)
    ws_red.set_column(1, len(rh) - 1, 10)
    filas_red = [(p, d) for p, _, d in M.redes(A)]
    r0 = 4
    for i, (prov, disp) in enumerate(filas_red):
        x = r0 + i
        e = x + 1
        sub = A[A["RED_SALUD"] == prov]
        ws_red.write(x, 0, prov, F["txt"])
        ws_red.write_formula(x, 1, f'=COUNTIF({rng_ip("A")},$A{e})', F["int"], len(sub))
        ws_red.write_formula(x, 2, f'=AVERAGEIFS({rng_ip("O")},{rng_ip("A")},$A{e})', F["d1"], sub["PUNTAJE_ANT"].mean())
        ws_red.write_formula(x, 3, f'=AVERAGEIFS({rng_ip("P")},{rng_ip("A")},$A{e})', F["d1"], sub["PUNTAJE"].mean())
        ws_red.write_formula(x, 4, f"=D{e}-C{e}", F["dlt"], sub["PUNTAJE"].mean() - sub["PUNTAJE_ANT"].mean())
        for k, s in enumerate(("VERDE", "AMARILLO", "ROJO")):
            ws_red.write_formula(x, 5 + k, f'=COUNTIFS({rng_ip("A")},$A{e},{rng_ip("S")},"{s}")', F["int"],
                                 int((sub["SEMAFORO_ANT"] == s).sum()))
            ws_red.write_formula(x, 8 + k, f'=COUNTIFS({rng_ip("A")},$A{e},{rng_ip("T")},"{s}")', F["int"],
                                 int((sub["SEMAFORO"] == s).sum()))
        for k, (s, cr, cm) in enumerate((("TI", "G", "H"), ("TC", "J", "K"), ("TM", "M", "N"))):
            ws_red.write_formula(x, 11 + k, f'=SUMIFS({rng_ip(cr)},{rng_ip("A")},$A{e})/SUMIFS({rng_ip(cm)},{rng_ip("A")},$A{e})',
                                 F["pct"], sub[f"{s} REALIZADAS"].sum() / sub[f"{s} META"].sum())
        if plan:
            ws_red.write_formula(x, 14, f'=AVERAGEIFS({rng_pl("Z")},{rng_pl("A")},$A{e})', F["d1"], sub["PUNTAJE_DIC_TEND"].mean())
            ws_red.write_formula(x, 15, f'=AVERAGEIFS({rng_pl("W")},{rng_pl("A")},$A{e})', F["d1"], sub["PUNTAJE_PLAN_BASE"].mean())
            ws_red.write_formula(x, 16, f'=AVERAGEIFS({rng_pl("Y")},{rng_pl("A")},$A{e})', F["d1"], sub["PUNTAJE_PLAN"].mean())
            for k, c in enumerate(M.CATEGORIAS):
                ws_red.write_formula(x, 17 + k, f'=COUNTIFS({rng_pl("A")},$A{e},{rng_pl("R")},"{c}")', F["int"],
                                     int((sub["CATEGORIA"] == c).sum()))
            ws_red.write_formula(x, 21, f'=COUNTIFS({rng_pl("A")},$A{e},{rng_pl("X")},"VERDE")', F["int"],
                                 int((sub["SEMAFORO_PLAN_BASE"] == "VERDE").sum()))
    # total
    xt = r0 + len(filas_red)
    e1, e2 = r0 + 1, xt
    ws_red.write(xt, 0, "TOTAL / PROMEDIO REGIONAL", F["btxt"])
    ws_red.write_formula(xt, 1, f"=SUM(B{e1}:B{e2})", F["bint"], n)
    ws_red.write_formula(xt, 2, f'=AVERAGE({rng_ip("O")})', F["bd1"], A["PUNTAJE_ANT"].mean())
    ws_red.write_formula(xt, 3, f'=AVERAGE({rng_ip("P")})', F["bd1"], A["PUNTAJE"].mean())
    ws_red.write_formula(xt, 4, f"=D{xt + 1}-C{xt + 1}", F["bd1"], A["PUNTAJE"].mean() - A["PUNTAJE_ANT"].mean())
    for j in range(5, 11):
        ws_red.write_formula(xt, j, f"=SUM({COL(j)}{e1}:{COL(j)}{e2})", F["bint"],
                             [int((A["SEMAFORO_ANT"] == s).sum()) for s in ("VERDE", "AMARILLO", "ROJO")][j - 5] if j < 8
                             else [int((A["SEMAFORO"] == s).sum()) for s in ("VERDE", "AMARILLO", "ROJO")][j - 8])
    for k, (s, cr, cm) in enumerate((("TI", "G", "H"), ("TC", "J", "K"), ("TM", "M", "N"))):
        ws_red.write_formula(xt, 11 + k, f'=SUM({rng_ip(cr)})/SUM({rng_ip(cm)})', F["bpct"],
                             A[f"{s} REALIZADAS"].sum() / A[f"{s} META"].sum())
    if plan:
        ws_red.write_formula(xt, 14, f'=AVERAGE({rng_pl("Z")})', F["bd1"], A["PUNTAJE_DIC_TEND"].mean())
        ws_red.write_formula(xt, 15, f'=AVERAGE({rng_pl("W")})', F["bd1"], A["PUNTAJE_PLAN_BASE"].mean())
        ws_red.write_formula(xt, 16, f'=AVERAGE({rng_pl("Y")})', F["bd1"], A["PUNTAJE_PLAN"].mean())
        for k, c in enumerate(M.CATEGORIAS):
            ws_red.write_formula(xt, 17 + k, f"=SUM({COL(17 + k)}{e1}:{COL(17 + k)}{e2})", F["bint"], int((A["CATEGORIA"] == c).sum()))
        ws_red.write_formula(xt, 21, f"=SUM(V{e1}:V{e2})", F["bint"], int((A["SEMAFORO_PLAN_BASE"] == "VERDE").sum()))
    ws_red.freeze_panes(4, 1)
    ws_red.conditional_format(f"D{e1}:D{e2}", {"type": "3_color_scale", "min_color": "#FCA5A5", "mid_color": "#FDE68A",
                                               "max_color": "#86EFAC"})
    # gráficos Redes
    ch = wb.add_chart({"type": "bar"})
    ser_red = [(2, f"Mes {M.MES_CORTE - 1}", "#CBD5E1"), (3, M.ETQ['MES'], "#60A5FA")]
    if plan:
        ser_red.append((15, f"{M.MES_ABR[M.MESES_AÑO]}: plan base", "#34D399"))
    for col, nombre, color in ser_red:
        ch.add_series({"name": nombre, "categories": ["Redes", r0, 0, xt - 1, 0], "values": ["Redes", r0, col, xt - 1, col],
                       "fill": {"color": color}, "border": {"color": "#475569"}, "data_labels": {"value": True, "num_format": "0.0", "font": {"size": 7}}})
    ch.set_title({"name": "Puntaje de Logro promedio por red", "name_font": {"size": 11}})
    ch.set_x_axis({"min": 0, "max": 100, "major_gridlines": {"visible": True, "line": {"color": "#E2E8F0"}}})
    ch.set_y_axis({"reverse": True, "num_font": {"size": 8}})
    ch.set_legend({"position": "bottom"})
    ch.set_size({"width": 760, "height": 520})
    ws_red.insert_chart(xt + 3, 0, ch)
    ch2 = wb.add_chart({"type": "column", "subtype": "percent_stacked"})
    for col, nombre, color in ((8, "Verde", "#86EFAC"), (9, "Amarillo", "#FDE047"), (10, "Rojo", "#FCA5A5")):
        ch2.add_series({"name": nombre, "categories": ["Redes", r0, 0, xt - 1, 0], "values": ["Redes", r0, col, xt - 1, col],
                        "fill": {"color": color}, "border": {"color": "#475569"}})
    ch2.set_title({"name": f"Composición del semáforo por red ({M.ETQ['MES']})", "name_font": {"size": 11}})
    ch2.set_x_axis({"num_font": {"size": 8, "rotation": -45}})
    ch2.set_legend({"position": "bottom"})
    ch2.set_size({"width": 760, "height": 520})
    ws_red.insert_chart(xt + 3, 9, ch2)
    if plan:
        ch3 = wb.add_chart({"type": "column", "subtype": "stacked"})
        for col, nombre, color in ((17, "CUMPLE", "#86EFAC"), (18, "CERCA", "#BEF264"), (19, "MODERADO", "#FDE047"), (20, "BRECHA ALTA", "#FCA5A5")):
            ch3.add_series({"name": nombre, "categories": ["Redes", r0, 0, xt - 1, 0], "values": ["Redes", r0, col, xt - 1, col],
                            "fill": {"color": color}, "border": {"color": "#475569"}})
        ch3.set_title({"name": f"IPRESS por categoría del plan {M.ETQ['PLAN']}", "name_font": {"size": 11}})
        ch3.set_x_axis({"num_font": {"size": 8, "rotation": -45}})
        ch3.set_legend({"position": "bottom"})
        ch3.set_size({"width": 760, "height": 420})
        ws_red.insert_chart(xt + 31, 0, ch3)
    ch4 = wb.add_chart({"type": "column"})
    for col, nombre, color in ((11, "TI", "#60A5FA"), (12, "TC", "#A78BFA"), (13, "TM", "#34D399")):
        ch4.add_series({"name": f"% meta {nombre}", "categories": ["Redes", r0, 0, xt - 1, 0], "values": ["Redes", r0, col, xt - 1, col],
                        "fill": {"color": color}, "border": {"color": "#475569"}})
    ch4.set_title({"name": f"Avance acumulado vs meta anual por servicio (esperado a {M.ETQ['MES_LARGO'].lower()}: {ESPERADO * 100:.1f} %)", "name_font": {"size": 11}})
    ch4.set_y_axis({"num_format": "0%", "min": 0})
    ch4.set_x_axis({"num_font": {"size": 8, "rotation": -45}})
    ch4.set_legend({"position": "bottom"})
    ch4.set_size({"width": 760, "height": 420})
    ws_red.insert_chart(xt + 31, 9, ch4)

    # ================================================================= RESUMEN
    ws_res.hide_gridlines(2)
    ws_res.set_column(0, 0, 44)
    ws_res.set_column(1, 4, 15)
    ws_res.write(0, 0, f"DIRESA {M.DIRESA} — Ficha N.° 32 Telemedicina — Corte {M.ETQ['MES']} ({M.ETQ['RANGO'].lower()} {M.ETQ['ANIO_TXT']})"
                       + (f" y Plan {M.ETQ['PLAN']}" if plan else ""), F["t"])
    ws_res.write(1, 0, f"Coordinación de Telemedicina | Fuente: reportes DITEL del Mes {M.MES_CORTE - 1} y {M.ETQ['MES']} (https://ficha32-ditel.web.app/) | Metas anuales, atenciones acumuladas", F["st"])
    co_cfg = M.CM.get_metas_config().get("cumplimiento_oficial", {})
    base_restar = int(co_cfg.get("base_restar", 10))
    divisor = int(co_cfg.get("divisor", 90))

    kp = [("IPRESS de primer nivel evaluadas (Reporte A)", f'=COUNTA({rng_ip("D")})', n, "kpi_i"),
          (f"Puntaje de Logro promedio — Mes {M.MES_CORTE - 1}", f'=AVERAGE({rng_ip("O")})', A["PUNTAJE_ANT"].mean(), "kpi_v"),
          (f"Puntaje de Logro promedio — {M.ETQ['MES']}", f'=AVERAGE({rng_ip("P")})', A["PUNTAJE"].mean(), "kpi_v"),
          (f"Cumplimiento oficial del puntaje promedio {M.ETQ['MES']} ((P−{base_restar})/{divisor})", f'=MAX(0,(B7-{base_restar})/{divisor})', max(0, (A["PUNTAJE"].mean() - base_restar) / divisor), "kpi_p"),
          (f"IPRESS en Verde ({M.ETQ['MES']})", f'=COUNTIF({rng_ip("T")},"VERDE")', int((A["SEMAFORO"] == "VERDE").sum()), "kpi_i"),
          (f"IPRESS en Amarillo ({M.ETQ['MES']})", f'=COUNTIF({rng_ip("T")},"AMARILLO")', int((A["SEMAFORO"] == "AMARILLO").sum()), "kpi_i"),
          (f"IPRESS en Rojo ({M.ETQ['MES']})", f'=COUNTIF({rng_ip("T")},"ROJO")', int((A["SEMAFORO"] == "ROJO").sum()), "kpi_i"),
          ("IPRESS con 0 TI en el año", f'=COUNTIF({rng_ip("G")},0)', int((A["TI REALIZADAS"] == 0).sum()), "kpi_i"),
          ("IPRESS con 0 TM en el año", f'=COUNTIF({rng_ip("M")},0)', int((A["TM REALIZADAS"] == 0).sum()), "kpi_i"),
          (f"Avance lineal esperado a {M.ETQ['MES_LARGO'].lower()} ({M.MES_CORTE}/{M.MESES_AÑO})", f"={M.MES_CORTE}/{M.MESES_AÑO}", ESPERADO, "kpi_p")]
    ws_res.write(3, 0, "INDICADORES CLAVE", F["h"])
    ws_res.write(3, 1, "Valor", F["h"])
    for i, (lab, fm, val, f) in enumerate(kp):
        ws_res.write(4 + i, 0, lab, F["kpi_l"])
        ws_res.write_formula(4 + i, 1, fm, F[f], val)
    # Avance por servicio
    ra = 4 + len(kp) + 1
    ws_res.write(ra, 0, "AVANCE ACUMULADO POR SERVICIO", F["h"])
    for j, h in enumerate([f"Acumulado {M.ETQ['MES']}", "Meta anual", "% de la meta", "Esperado"]):
        ws_res.write(ra, 1 + j, h, F["h"])
    for i, (s, cr, cm) in enumerate((("TI", "G", "H"), ("TC", "J", "K"), ("TM", "M", "N"))):
        x = ra + 1 + i
        ws_res.write(x, 0, f"{M.NOMBRE_SERV[s]} ({s}) — peso {M.PESO[s]}", F["kpi_l"])
        ws_res.write_formula(x, 1, f'=SUM({rng_ip(cr)})', F["int"], A[f"{s} REALIZADAS"].sum())
        ws_res.write_formula(x, 2, f'=SUM({rng_ip(cm)})', F["int"], A[f"{s} META"].sum())
        ws_res.write_formula(x, 3, f"=B{x + 1}/C{x + 1}", F["pct"], A[f"{s} REALIZADAS"].sum() / A[f"{s} META"].sum())
        ws_res.write_formula(x, 4, f"={M.MES_CORTE}/{M.MESES_AÑO}", F["pct"], ESPERADO)
    # Escenarios (solo con la parte de planificación)
    if plan:
        re_ = ra + 6
        ws_res.write(re_, 0, "ESCENARIOS DE PUNTAJE PROMEDIO", F["h"])
        for j, h in enumerate(["Puntaje promedio", "Cumplim. oficial", "IPRESS Verde", "IPRESS Rojo"]):
            ws_res.write(re_, 1 + j, h, F["h"])
        A["_sem_tend"] = A["PUNTAJE_DIC_TEND"].apply(M.clasificar)
        esc = [(f"Mes {M.MES_CORTE - 1} (oficial)", rng_ip("O"), "PUNTAJE_ANT", "SEMAFORO_ANT", rng_ip("S")),
               (f"{M.ETQ['MES']} (oficial)", rng_ip("P"), "PUNTAJE", "SEMAFORO", rng_ip("T")),
               (f"Proyección fin de {M.ETQ['HASTA'].lower()} (ritmo {M.ETQ['RANGO'].lower()})", rng_pl("K"), "PUNTAJE_PROY_OCT", "SEMAFORO_OCT", None),
               (f"{M.MES_TXT[M.MESES_AÑO]} — sigue la tendencia", rng_pl("Z"), "PUNTAJE_DIC_TEND", "_sem_tend", None),
               (f"{M.MES_TXT[M.MESES_AÑO]} — mantiene ritmo de {M.ETQ['MES_LARGO'].lower()}", rng_pl("AA"), "PUNTAJE_DIC_AGO", None, None),
               (f"{M.MES_TXT[M.MESES_AÑO]} — PLAN BASE (exigible)", rng_pl("W"), "PUNTAJE_PLAN_BASE", "SEMAFORO_PLAN_BASE", rng_pl("X")),
               (f"{M.MES_TXT[M.MESES_AÑO]} — PLAN AMPLIADO (techo operativo)", rng_pl("Y"), "PUNTAJE_PLAN", "SEMAFORO_PLAN", None)]
        for i, (lab, rg, col, scol, srng) in enumerate(esc):
            x = re_ + 1 + i
            ws_res.write(x, 0, lab, F["kpi_l"])
            ws_res.write_formula(x, 1, f"=AVERAGE({rg})", F["d1"], A[col].mean())
            ws_res.write_formula(x, 2, f"=MAX(0,(B{x + 1}-10)/90)", F["pct"], max(0, (A[col].mean() - 10) / 90))
            if scol:
                ws_res.write_number(x, 3, int((A[scol] == "VERDE").sum()), F["int"])
                ws_res.write_number(x, 4, int((A[scol] == "ROJO").sum()), F["int"])
            else:
                ws_res.write(x, 3, "—", F["c"])
                ws_res.write(x, 4, "—", F["c"])
        ce = re_ + len(esc) + 2
        ws_res.write(ce, 0, "IPRESS POR CATEGORÍA DEL PLAN", F["h"])
        ws_res.write(ce, 1, "N.° IPRESS", F["h"])
        ws_res.write(ce, 2, f"Atenciones {M.ETQ['PLAN']} (plan base)", F["h"])
        for i, c in enumerate(M.CATEGORIAS):
            x = ce + 1 + i
            ws_res.write(x, 0, c, F["kpi_l"])
            ws_res.write_formula(x, 1, f'=COUNTIF({rng_pl("R")},"{c}")', F["int"], int((A["CATEGORIA"] == c).sum()))
            ws_res.write_formula(x, 2, f'=SUMIF({rng_pl("R")},"{c}",{rng_pl("T")})+SUMIF({rng_pl("R")},"{c}",{rng_pl("U")})+SUMIF({rng_pl("R")},"{c}",{rng_pl("V")})',
                                 F["int"], int(A.loc[A["CATEGORIA"] == c, ["TI OBJ_ND", "TC OBJ_ND", "TM OBJ_ND"]].sum().sum()))
    ws_res.set_row(3, 20)
    # gráficos resumen
    if plan:
        c1 = wb.add_chart({"type": "column"})
        c1.add_series({"name": "Puntaje promedio", "categories": ["Resumen", re_ + 1, 0, re_ + len(esc), 0],
                       "values": ["Resumen", re_ + 1, 1, re_ + len(esc), 1], "gap": 60,
                       "points": [{"fill": {"color": c}} for c in ["#94A3B8", "#475569", "#93C5FD", "#FBBF24", "#FCD34D", "#34D399", "#059669"]],
                       "data_labels": {"value": True, "num_format": "0.0", "font": {"bold": True}}})
        c1.set_title({"name": f"Puntaje de Logro promedio regional: oficial y escenarios a {M.MES_TXT[M.MESES_AÑO].lower()}", "name_font": {"size": 11}})
        c1.set_y_axis({"min": 0, "max": 100, "major_gridlines": {"visible": True, "line": {"color": "#E2E8F0"}}})
        c1.set_x_axis({"num_font": {"size": 8}})
        c1.set_legend({"none": True})
        c1.set_size({"width": 700, "height": 380})
        ws_res.insert_chart(3, 6, c1)
    c2 = wb.add_chart({"type": "column"})
    for col, nombre, color in ((1, f"Acumulado {M.ETQ['MES']}", "#60A5FA"),):
        pass
    c2.add_series({"name": f"% de la meta anual ({M.ETQ['MES']})", "categories": ["Resumen", ra + 1, 0, ra + 3, 0],
                   "values": ["Resumen", ra + 1, 3, ra + 3, 3], "fill": {"color": "#60A5FA"}, "border": {"color": "#1D4ED8"},
                   "data_labels": {"value": True, "num_format": "0.0%", "font": {"bold": True}}})
    c2.add_series({"name": f"Avance lineal esperado ({ESPERADO * 100:.1f} %)", "categories": ["Resumen", ra + 1, 0, ra + 3, 0],
                   "values": ["Resumen", ra + 1, 4, ra + 3, 4], "type": "line",
                   "line": {"color": "#DC2626", "dash_type": "dash", "width": 2}})
    c2.set_title({"name": "Avance acumulado vs meta anual por servicio", "name_font": {"size": 11}})
    c2.set_y_axis({"num_format": "0%", "min": 0})
    c2.set_legend({"position": "bottom"})
    c2.set_size({"width": 700, "height": 340})
    ws_res.insert_chart(23, 6, c2)
    if plan:
        c3 = wb.add_chart({"type": "pie"})
        c3.add_series({"name": "Categorías", "categories": ["Resumen", ce + 1, 0, ce + 4, 0], "values": ["Resumen", ce + 1, 1, ce + 4, 1],
                       "points": [{"fill": {"color": c}} for c in ["#86EFAC", "#BEF264", "#FDE047", "#FCA5A5"]],
                       "data_labels": {"value": True, "percentage": True, "font": {"size": 9, "bold": True}}})
        c3.set_title({"name": f"IPRESS según esfuerzo requerido (plan {M.ETQ['PLAN']})", "name_font": {"size": 11}})
        c3.set_size({"width": 460, "height": 340})
        ws_res.insert_chart(ce + 6, 0, c3)
        c4 = wb.add_chart({"type": "column", "subtype": "stacked"})
        for nombre, col, color in (("Verde", "VERDE", "#86EFAC"), ("Amarillo", "AMARILLO", "#FDE047"), ("Rojo", "ROJO", "#FCA5A5")):
            pass
        ws_res.write(ce + 7 + 18, 0, "", F["b"])
        # tabla auxiliar semáforo por escenario (para gráfico apilado)
        aux = ce + 26
        ws_res.write(aux, 0, "SEMÁFORO POR ESCENARIO (N.° IPRESS)", F["h"])
        for j, h in enumerate(["Verde", "Amarillo", "Rojo"]):
            ws_res.write(aux, 1 + j, h, F["h"])
        esc_sem = [(f"Mes {M.MES_CORTE - 1}", "SEMAFORO_ANT"), (M.ETQ['MES'], "SEMAFORO"), (f"Proy. {M.ETQ['PROY']}", "SEMAFORO_OCT"), (f"{M.MES_ABR[M.MESES_AÑO]}: tendencia", "_sem_tend"),
                   (f"{M.MES_ABR[M.MESES_AÑO]}: plan base", "SEMAFORO_PLAN_BASE"), (f"{M.MES_ABR[M.MESES_AÑO]}: techo", "SEMAFORO_PLAN")]
        for i, (lab, col) in enumerate(esc_sem):
            ws_res.write(aux + 1 + i, 0, lab, F["kpi_l"])
            for j, s in enumerate(("VERDE", "AMARILLO", "ROJO")):
                ws_res.write_number(aux + 1 + i, 1 + j, int((A[col] == s).sum()), F["int"])
        for j, (nombre, color) in enumerate((("Verde", "#86EFAC"), ("Amarillo", "#FDE047"), ("Rojo", "#FCA5A5"))):
            c4.add_series({"name": nombre, "categories": ["Resumen", aux + 1, 0, aux + len(esc_sem), 0],
                           "values": ["Resumen", aux + 1, 1 + j, aux + len(esc_sem), 1 + j], "fill": {"color": color},
                           "border": {"color": "#475569"}, "data_labels": {"value": True, "font": {"size": 8}}})
        c4.set_title({"name": "Semáforo de las 371 IPRESS según escenario", "name_font": {"size": 11}})
        c4.set_legend({"position": "bottom"})
        c4.set_size({"width": 700, "height": 340})
        ws_res.insert_chart(41, 6, c4)

    # ================================================================= PALANCA
    ws_pa.write(0, 0, "IPRESS PALANCA — dónde pocas atenciones mueven más el escenario", F["t"])
    ws_pa.write(1, 0, f"Basado en la proyección a fin de {M.ETQ['HASTA'].lower()} (ritmo {M.ETQ['RANGO'].lower()}). Puntos por atención = peso del servicio ÷ meta anual del servicio.", F["st"])
    P = A[A["CATEGORIA"] != "CUMPLE"].copy()
    P["GANA"] = P["PUNTAJE_PLAN_BASE"] - P["PUNTAJE_PROY_OCT"]
    P["AT_PLAN"] = P["TI OBJ_ND"] + P["TC OBJ_ND"] + P["TM OBJ_ND"]
    P["EFIC"] = np.where(P["AT_PLAN"] > 0, P["GANA"] / P["AT_PLAN"], 0)
    # Tabla 1: mayor ganancia con el plan base
    t1 = P.sort_values("GANA", ascending=False).head(40)
    ws_pa.write(3, 0, f"A. 40 IPRESS con mayor ganancia de puntaje con el plan base ({M.ETQ['PLAN']})", F["b"])
    h1 = ["Red de Salud", "Cat.", "Establecimiento / IPRESS", "Categoría", f"Pts proy. {M.ETQ['PROY']}", f"Pts {M.MES_ABR[M.MESES_AÑO]} (plan base)", "Ganancia", "Atenciones plan", "Pts por atención", f"Semáforo {M.ETQ['PROY']}", f"Semáforo {M.MES_ABR[M.MESES_AÑO]}"]
    for j, h in enumerate(h1):
        ws_pa.write(4, j, h, F["h"])
    for i, (_, r) in enumerate(t1.iterrows()):
        x = 5 + i
        ws_pa.write(x, 0, r["RED_SALUD"], F["txt"]); ws_pa.write(x, 1, r["NIVEL"], F["c"]); ws_pa.write(x, 2, r["NOMBRE"], F["txt"])
        ws_pa.write(x, 3, r["CATEGORIA"], F["c"])
        ws_pa.write_number(x, 4, r["PUNTAJE_PROY_OCT"], F["d1"]); ws_pa.write_number(x, 5, r["PUNTAJE_PLAN_BASE"], F["d1"])
        ws_pa.write_formula(x, 6, f"=F{x + 1}-E{x + 1}", F["dlt"], r["GANA"])
        ws_pa.write_number(x, 7, r["AT_PLAN"], F["int"])
        ws_pa.write_formula(x, 8, f"=IF(H{x + 1}>0,G{x + 1}/H{x + 1},0)", F["d2"], r["EFIC"])
        ws_pa.write(x, 9, r["SEMAFORO_OCT"], F["c"]); ws_pa.write(x, 10, r["SEMAFORO_PLAN_BASE"], F["c"])
    cf_cat(ws_pa, f"D6:D{5 + len(t1)}")
    cf_sem(ws_pa, f"J6:K{5 + len(t1)}")
    # Tabla 2: victoria rápida teleconsulta
    q = A[(A["TC REALIZADAS"] == 0)].copy()
    q["PTS_TC"] = 10 / q["TC META"]
    q = q.sort_values("PTS_TC", ascending=False)
    c0 = 13
    ws_pa.write(3, c0, f"B. Victoria rápida: {len(q)} IPRESS sin ninguna teleconsulta — 1 TC aporta puntos inmediatos", F["b"])
    h2 = ["Red de Salud", "Establecimiento / IPRESS", "TC Meta anual", "Pts por 1 TC", f"Puntaje {M.ETQ['MES']}", "Puntaje con 1 TC"]
    for j, h in enumerate(h2):
        ws_pa.write(4, c0 + j, h, F["h"])
    for i, (_, r) in enumerate(q.head(60).iterrows()):
        x = 5 + i
        ws_pa.write(x, c0, r["RED_SALUD"], F["txt"]); ws_pa.write(x, c0 + 1, r["NOMBRE"], F["txt"])
        ws_pa.write_number(x, c0 + 2, r["TC META"], F["int"])
        ws_pa.write_formula(x, c0 + 3, f"=10/{COL(c0 + 2)}{x + 1}", F["d1"], r["PTS_TC"])
        ws_pa.write_number(x, c0 + 4, r["PUNTAJE"], F["d1"])
        ws_pa.write_formula(x, c0 + 5, f"={COL(c0 + 4)}{x + 1}+{COL(c0 + 3)}{x + 1}", F["d1"], r["PUNTAJE"] + r["PTS_TC"])
    # Tabla 3: a pocas atenciones del verde
    v = A[(A["SEMAFORO"] != "VERDE") & (A["AT_VERDE"] <= 12)].copy()
    v = v.sort_values("AT_VERDE")
    c1_ = 21
    ws_pa.write(3, c1_, f"C. {len(v)} IPRESS a ≤12 atenciones adicionales (desde proy. {M.ETQ['HASTA'].lower()}) de llegar a Verde (≥60 pts)", F["b"])
    h3 = ["Red de Salud", "Establecimiento / IPRESS", f"Puntaje {M.ETQ['MES']}", f"Pts proy. {M.ETQ['PROY']}", "Atenciones mín. para Verde", "TI", "TC", "TM"]
    for j, h in enumerate(h3):
        ws_pa.write(4, c1_ + j, h, F["h"])
    for i, (_, r) in enumerate(v.head(80).iterrows()):
        x = 5 + i
        det = r["AT_VERDE_DET"]
        ws_pa.write(x, c1_, r["RED_SALUD"], F["txt"]); ws_pa.write(x, c1_ + 1, r["NOMBRE"], F["txt"])
        ws_pa.write_number(x, c1_ + 2, r["PUNTAJE"], F["d1"]); ws_pa.write_number(x, c1_ + 3, r["PUNTAJE_PROY_OCT"], F["d1"])
        ws_pa.write_number(x, c1_ + 4, r["AT_VERDE"], F["int"])
        for k, s in enumerate(M.SERV):
            ws_pa.write_number(x, c1_ + 5 + k, det[s], F["int"])
    ws_pa.set_row(4, 40)
    for col, w in ((0, 20), (1, 6), (2, 36), (3, 13), (4, 10), (5, 10), (6, 10), (7, 10), (8, 10), (9, 11), (10, 11), (11, 3), (12, 3),
                   (13, 20), (14, 36), (15, 10), (16, 10), (17, 10), (18, 10), (19, 3), (20, 3), (21, 20), (22, 36), (23, 10), (24, 10),
                   (25, 12), (26, 6), (27, 6), (28, 6)):
        ws_pa.set_column(col, col, w)
    ws_pa.freeze_panes(5, 0)

    # ================================================================= HOSPITALES
    ws_ho.write(0, 0, "Hospitales consultores (Reporte B — Logro B): solo Teleinterconsulta, peso 100", F["t"])
    ws_ho.write(1, 0, f"Meta TI = máx(100, redondeo(40/10000 × atenciones HIS acumuladas)): es dinámica y crece con el año."
                      + (f" Proyección lineal de atenciones HIS a {M.MES_TXT[M.MESES_AÑO].lower()}." if plan else ""), F["st"])
    hh = ["Red", "Hospital", "Cat.", f"Atenciones HIS Mes {M.MES_CORTE - 1}", f"Atenciones HIS {M.ETQ['MES']}", f"Meta TI {M.ETQ['MES']}", f"TI Mes {M.MES_CORTE - 1}", f"TI {M.ETQ['MES']}",
          f"Pts Mes {M.MES_CORTE - 1}", f"Pts {M.ETQ['MES']}", "Semáforo", f"Meta TI proy. {M.MES_ABR[M.MESES_AÑO]}", f"TI proy. fin {M.ETQ['PROY']}", f"TI requeridas {M.ETQ['PLAN']}", "Req. /mes",
          "Esfuerzo exigente /mes", f"TI proy. {M.MES_ABR[M.MESES_AÑO]} (plan)", f"Pts {M.MES_ABR[M.MESES_AÑO]} (plan)",
          "Rol Telesalud (SUSALUD)", "Nro. Res. Cartera", "Fecha Res. Cartera", "¿Subió Res.?"]
    if not plan:
        hh = hh[:11]            # solo situación actual: sin proyección ni plan
    for j, h in enumerate(hh):
        ws_ho.write(3, j, h, F["h"])
    ws_ho.set_row(3, 42)
    for i, (_, r) in enumerate(B.iterrows()):
        x = 4 + i
        e = x + 1
        ws_ho.write(x, 0, r["RED_SALUD"], F["txt"]); ws_ho.write(x, 1, r["NOMBRE"], F["txt"]); ws_ho.write(x, 2, r["NIVEL"], F["c"])
        ws_ho.write_number(x, 3, r["DEN ANT"], F["int"]); ws_ho.write_number(x, 4, r["DENOMINADOR"], F["int"])
        ws_ho.write_number(x, 5, r["TI META"], F["int"]); ws_ho.write_number(x, 6, r["TI ANT"], F["int"]); ws_ho.write_number(x, 7, r["TI REALIZADAS"], F["int"])
        ws_ho.write_number(x, 8, r["PUNTAJE_ANT"], F["d1"])
        ws_ho.write_formula(x, 9, f"=ROUND(100*MIN(H{e},F{e})/F{e},1)", F["d1"], r["PUNTAJE"])
        ws_ho.write_formula(x, 10, f'=IF(J{e}>=60,"VERDE",IF(J{e}>=30,"AMARILLO","ROJO"))', F["c"], r["SEMAFORO"])
        if plan:
            ws_ho.write_formula(x, 11, f"=MAX(100,ROUND(40/10000*E{e}/{M.MES_CORTE}*{M.MESES_AÑO},0))", F["int"], r["META_PROY_DIC"])
            ws_ho.write_number(x, 12, r["TI_PROY_OCT"], F["int"])
            ws_ho.write_formula(x, 13, f"=MAX(0,L{e}-M{e})", F["int"], r["BRECHA_DIC"])
            ws_ho.write_formula(x, 14, f"=ROUNDUP(N{e}/{M.MESES_PLAN},0)", F["int"], r["REQ_MES"])
            ws_ho.write_number(x, 15, r["CAP_MES"], F["int"])
            ws_ho.write_formula(x, 16, f"=MIN(L{e},M{e}+{M.MESES_PLAN}*P{e})", F["int"], r["TI_PLAN_DIC"])
            ws_ho.write_formula(x, 17, f"=ROUND(100*Q{e}/L{e},1)", F["d1"], r["PUNTAJE_PLAN"])
        col_h_extra = 18 if plan else 11
        ws_ho.write(x, col_h_extra, r.get("TELESALUD_ROL", "NO REGISTRADO"), F["c"])
        ws_ho.write(x, col_h_extra + 1, r.get("NRO_RESOLUCION", "—"), F["c"])
        ws_ho.write(x, col_h_extra + 2, r.get("FECHA_RESOLUCION", "—"), F["c"])
        ws_ho.write(x, col_h_extra + 3, r.get("VARIACION_RESOLUCION", "NO SUBIÓ"), F["c"])
    num_hosp = len(B)
    if num_hosp > 0:
        cf_sem(ws_ho, f"K5:K{4 + num_hosp}")
    ws_ho.set_column(0, 0, 14); ws_ho.set_column(1, 1, 42); ws_ho.set_column(2, 2, 6); ws_ho.set_column(3, 17, 11)
    if num_hosp > 0:
        chh = wb.add_chart({"type": "column"})
        ser_ho = [(8, f"Pts Mes {M.MES_CORTE - 1}", "#CBD5E1"), (9, f"Pts {M.ETQ['MES']}", "#60A5FA")]
        if plan:
            ser_ho.append((17, f"Pts {M.MES_ABR[M.MESES_AÑO]} (plan)", "#34D399"))
        for col, nombre, color in ser_ho:
            chh.add_series({"name": nombre, "categories": ["Hospitales", 4, 1, 4 + num_hosp - 1, 1],
                            "values": ["Hospitales", 4, col, 4 + num_hosp - 1, col],
                            "fill": {"color": color}, "border": {"color": "#475569"},
                            "data_labels": {"value": True, "num_format": "0.0", "font": {"size": 8}}})
        chh.set_title({"name": "Puntaje de Logro de hospitales consultores", "name_font": {"size": 11}})
        chh.set_y_axis({"min": 0, "max": 110})
        chh.set_legend({"position": "bottom"})
        chh.set_size({"width": 760, "height": 360})
        chart_row = max(11, 4 + num_hosp + 2)
        ws_ho.insert_chart(chart_row, 0, chh)
        chh2 = wb.add_chart({"type": "column"})
        ser_ho2 = [(6, f"TI Mes {M.MES_CORTE - 1}", "#CBD5E1"), (7, f"TI {M.ETQ['MES']}", "#60A5FA"), (5, f"Meta TI {M.ETQ['MES']}", "#FCA5A5")]
        if plan:
            ser_ho2.append((11, f"Meta TI proy. {M.MES_ABR[M.MESES_AÑO]}", "#F87171"))
        for col, nombre, color in ser_ho2:
            chh2.add_series({"name": nombre, "categories": ["Hospitales", 4, 1, 4 + num_hosp - 1, 1],
                             "values": ["Hospitales", 4, col, 4 + num_hosp - 1, col],
                             "fill": {"color": color}, "border": {"color": "#475569"}})
        chh2.set_title({"name": f"Teleinterconsultas vs meta (Mes {M.MES_CORTE - 1}, {M.ETQ['MES']}"
                                + (f" y proyección de la meta a {M.MES_TXT[M.MESES_AÑO].lower()})" if plan else ")"),
                        "name_font": {"size": 11}})
        chh2.set_legend({"position": "bottom"})
        chh2.set_size({"width": 760, "height": 360})
        ws_ho.insert_chart(chart_row, 9, chh2)

    # ================================================================= SUPUESTOS
    ws_su.set_column(0, 0, 36)
    ws_su.set_column(1, 1, 120)
    ws_su.write(0, 0, "Supuestos, método y definiciones", F["t"])
    api_url = M.CM.get_api_url()
    filas = [
        ("Fuente de datos", f"Reportes DITEL de Mes {M.MES_CORTE - 1} y {M.ETQ['MES']}: Reporte A ({len(A)} IPRESS de primer nivel, logro A) y Reporte B ({len(B)} hospitales, logro B). Endpoint: {api_url}"),
        ("Metas anuales", f"Metas (TI, TC, TM) anuales de la Ficha Técnica N.° 32 ({M.ETQ['ANIO_TXT']}) en las {len(A)} IPRESS de primer nivel evaluadas."),
        ("Producción acumulada", f"Atenciones realizadas acumuladas ene→mes. Producción mensual calculada por diferencia frente al mes previo: {M.ETQ['MES']} − Mes {M.MES_CORTE - 1}."),
        ("Puntaje de Logro (A)", f"{M.PESO['TI']}·min(TI,MetaTI)/MetaTI + {M.PESO['TC']}·min(TC,MetaTC)/MetaTC + {M.PESO['TM']}·min(TM,MetaTM)/MetaTM. Cumplimiento oficial = (Puntaje − {base_restar})/{divisor}."),
        ("Puntaje de Logro (B)", "100·min(TI,MetaTI)/MetaTI; MetaTI = máx(100, redondeo(40/10000 × atenciones HIS acumuladas))."),
        (f"Proyección a {M.ETQ['HASTA'].lower()}", f"Proyección a fin de {M.ETQ['HASTA'].lower()}: acumulado {M.ETQ['MES']} + {M.MESES_PROY_PREVIOS} × ritmo promedio {M.ETQ['RANGO'].lower()} (conservador)."),
        (f"Brecha {M.ETQ['PLAN']}", f"Por servicio: máx(0, meta anual − proyección fin de {M.ETQ['HASTA'].lower()}). Requerimiento mensual total = suma de brechas ÷ {M.MESES_PLAN}, redondeado hacia arriba."),
        ("Capacidad realista /mes", f"Mayor entre {M.PISO_CAP_MES} atenciones, el {int(M.PCT_CAP_META * 100)} % de la meta anual total (TI+TC+TM) y el rendimiento mensual demostrado (suma de ritmos promedio {M.ETQ['RANGO'].lower()})."),
        ("Categorías", f"CUMPLE: brecha 0. CERCA: requerimiento ≤ 1/3 de la capacidad. MODERADO: ≤ 1 vez la capacidad. BRECHA ALTA: > 1 vez la capacidad. Ruta (solo brecha alta): según el puntaje que alcanzaría con el plan ampliado (≥60 → Ruta a Verde; ≥30 → Ruta a Amarillo; <30 → Arranque)."),
        ("PLAN BASE (exigible)", f"CUMPLE/CERCA/MODERADO completan su meta anual. BRECHA ALTA duplica su ritmo en {M.ETQ['PLAN']} y, como piso de arranque, aporta 1/3 de su capacidad mensual × {M.MESES_PLAN} meses; las atenciones se asignan primero al servicio que más puntos rinde por atención (peso ÷ meta)."),
        ("PLAN AMPLIADO (techo operativo)", f"Todas las IPRESS destinan su capacidad completa × {M.MESES_PLAN} meses, priorizando el servicio de mayor puntos por atención. Es un límite superior de referencia, no una meta exigible."),
        (f"Tendencia / ritmo de {M.ETQ['MES_LARGO'].lower()}", f"Escenarios de referencia sin acción adicional: continuar el ritmo promedio {M.ETQ['RANGO'].lower()} hasta {M.MES_TXT[M.MESES_AÑO].lower()} (tendencia) o la producción del {M.ETQ['MES']} (ritmo de {M.ETQ['MES_LARGO'].lower()})."),
        ("Hospitales (B)", f"Proyección lineal de las atenciones HIS a {M.MES_TXT[M.MESES_AÑO].lower()} (acumulado/{M.MES_CORTE} × {M.MESES_AÑO}). Esfuerzo exigente = máx(ritmo de {M.ETQ['MES_LARGO'].lower()}, 2 × promedio {M.ETQ['RANGO'].lower()})."),
        ("Códigos CPMS (Ficha 32)", "TI: 99499.11 y 99499.12 (LAB1/LAB2); TC: 99499.01 y 99499.03; TM: 99499.10 con CIE-10 del Anexo 1 y tipo de diagnóstico R."),
        ("Limitaciones", f"Modelo determinista basado en ritmos históricos. Las cifras previas a {M.ETQ['HASTA'].lower()} son proyecciones hasta contar con el reporte oficial."),
        ("DIRECTRIZ CRÍTICA DE REGISTRO", "REGISTRO EXCLUSIVO EN TELEATIENDO (NO DUPLICAR EN HIS): Cada atención de telemedicina debe registrarse ÚNICAMENTE en la plataforma Teleatiendo (lote TLH). Teleatiendo ya se encuentra conectada e interoperada con el HIS-MINSA y el FUA (SIS). Queda estrictamente prohibido registrar la misma atención en hojas HIS físicas/manuales, ya que la duplicidad anula o cancela ambos registros, perjudicando el indicador."),
        ("Hojas", f"Resumen (KPIs y gráficos) · Redes (comparativo con fórmulas) · IPRESS ({len(A)} IPRESS) · Plan {M.ETQ['PLAN'].replace('–', '-')} · Palanca · Hospitales ({len(B)} hospitales) · Red_* (vista por red con gráficos)."),
    ]
    for i, (a, b) in enumerate(filas):
        ws_su.write(2 + i, 0, a, F["kpi_l"])
        ws_su.write(2 + i, 1, b, F["txtw"])
        ws_su.set_row(2 + i, 40)

    # ================================================================= CARTERA SERVICIOS
    # Hoja oficial con datos de Cartera de Servicios desde SUSALUD RENIPRESS
    # Muestra todos los establecimientos, su rol oficial de Telesalud (Consultor/Consultante/Ambos)
    # y su estado de inclusión en el Convenio de Gestión Ficha 32.
    if cartera:
        ws_cartera = wb.add_worksheet("Cartera Servicios")
        ws_cartera.set_tab_color("#FBBF24")
    else:
        ws_cartera = _Omitida()

    df_car, resumen = M.obtener_cartera_renipress()
    cods_f32 = set(A["CODIGO"].dropna().unique())
    if len(B) and "CODIGO" in B.columns:
        cods_f32 |= set(B["CODIGO"].dropna().unique())

    # Filtrar Cartera de Servicios según el alcance del informe (si se seleccionó una red o eess específicos)
    alcance_txt = f"DIRESA {M.DIRESA}"
    filtro_red = M.FILTRO.get("red")
    filtro_cods = M.FILTRO.get("codigos") or []

    if filtro_red:
        pat_red = str(filtro_red).strip().upper()
        df_car = df_car[df_car["RED_SALUD"].astype(str).str.upper().str.contains(pat_red, regex=False, na=False)].reset_index(drop=True)
        alcance_txt = f"RED DE SALUD {filtro_red.upper()}"
    elif filtro_cods:
        cods_set = {str(c).strip().zfill(8) for c in filtro_cods}
        df_car = df_car[df_car["CODIGO_UNICO"].isin(cods_set)].reset_index(drop=True)
        alcance_txt = f"{len(df_car)} ESTABLECIMIENTOS SELECCIONADOS"

    # Formato adicional para resoluciones
    fmt_subio = F["subio"]
    fmt_nosubio = F["nosubio"]

    # Encabezados de la hoja Cartera Servicios (incluye Categorización, Cartera y Comparativa de Categorías)
    car_h = [
        "CÓDIGO ÚNICO", "ESTABLECIMIENTO / IPRESS", "RED DE SALUD", "PROVINCIA",
        "DISTRITO", "CATEGORÍA FICHA 32", "CATEGORÍA SUSALUD", "¿CAMBIÓ CATEGORÍA?",
        "ROL TELESALUD (SUSALUD)",
        "NRO. DOC. CATEGORIZACIÓN", "FECHA CATEGORIZACIÓN", "¿SUBIÓ RESOLUCIÓN CATEGORIZACIÓN?",
        "NRO. RESOLUCIÓN CARTERA", "FECHA RESOLUCIÓN CARTERA", "¿SUBIÓ RESOLUCIÓN CARTERA?",
        "EN CONVENIO FICHA 32", "ESTADO", "DIRECCIÓN"
    ]
    T0 = 0
    for j, h in enumerate(car_h):
        ws_cartera.write(T0, j, h, F["h"])
    ws_cartera.set_row(T0, 32)

    ws_cartera.set_column(0, 0, 14)   # Código
    ws_cartera.set_column(1, 1, 36)   # Nombre
    ws_cartera.set_column(2, 2, 18)   # Red
    ws_cartera.set_column(3, 3, 16)   # Provincia
    ws_cartera.set_column(4, 4, 16)   # Distrito
    ws_cartera.set_column(5, 5, 14)   # Categoría F32
    ws_cartera.set_column(6, 6, 14)   # Categoría SUSALUD
    ws_cartera.set_column(7, 7, 20)   # ¿Cambió Categoría?
    ws_cartera.set_column(8, 8, 22)   # Rol Telesalud
    ws_cartera.set_column(9, 9, 34)   # Nro. Doc Categorización
    ws_cartera.set_column(10, 10, 18) # Fecha Categorización
    ws_cartera.set_column(11, 11, 24) # ¿Subió Res. Categorización?
    ws_cartera.set_column(12, 12, 32) # Nro. Res. Cartera
    ws_cartera.set_column(13, 13, 18) # Fecha Res. Cartera
    ws_cartera.set_column(14, 14, 22) # ¿Subió Res. Cartera?
    ws_cartera.set_column(15, 15, 18) # En Convenio Ficha 32
    ws_cartera.set_column(16, 16, 12) # Estado
    ws_cartera.set_column(17, 17, 40) # Dirección

    # Ordenar: primero los que SÍ están en la Ficha 32, luego los que NO; dentro de cada bloque por Rol Telesalud y Nombre
    if not df_car.empty:
        orden_rol = {"AMBOS": 1, "CONSULTOR": 2, "CONSULTANTE": 3, "NO REGISTRADO": 4}
        df_car["_en_f32"] = df_car["CODIGO_UNICO"].apply(lambda c: 0 if c in cods_f32 else 1)
        df_car["_orden_rol"] = df_car["telemedicina"].map(orden_rol).fillna(5)
        df_car = df_car.sort_values(["_en_f32", "_orden_rol", "RED_SALUD", "nombre_ipress"]).reset_index(drop=True)

        dict_cat_f32 = dict(zip(A["CODIGO"], A["NIVEL"]))
        if len(B) and "CODIGO" in B.columns:
            dict_cat_f32.update(dict(zip(B["CODIGO"], B.get("NIVEL", pd.Series(["II-1"] * len(B))))))

        for i, (_, r) in enumerate(df_car.iterrows()):
            x = T0 + 1 + i
            cod = str(r.get("CODIGO_UNICO", r.get("id_ipress", ""))).zfill(8)
            rol_t = r.get("telemedicina", "NO REGISTRADO")
            en_f32 = "SÍ" if cod in cods_f32 else "NO"

            nro_cat = str(r.get("nro_categorizacion", "—") or "—")
            fec_cat = str(r.get("fecha_categorizacion", "—") or "—")
            sub_cat = str(r.get("subio_categorizacion", "NO SUBIÓ") or "NO SUBIÓ")

            nro_car = str(r.get("nro_cartera", "—") or "—")
            fec_car = str(r.get("fecha_cartera", "—") or "—")
            sub_car = str(r.get("subio_cartera", "NO SUBIÓ") or "NO SUBIÓ")

            # Categoría Ficha 32 vs SUSALUD
            cat_f32 = dict_cat_f32.get(cod, "—")
            cat_su = str(r.get("categoria_susalud") or r.get("clasificacion") or "").strip()
            if not cat_su or cat_su in ("—", "Sin Categoría"):
                cat_su = "—"

            if en_f32 == "NO":
                estado_cat = "— (Fuera de F32)"
                fmt_est = F["c"]
            elif cat_su == "—":
                estado_cat = "SIN DATOS SUSALUD"
                fmt_est = F["c"]
            elif cat_f32 == cat_su:
                estado_cat = "✓ IGUAL"
                fmt_est = fmt_subio
            else:
                estado_cat = f"⚠ VARIÓ ({cat_f32} → {cat_su})"
                fmt_est = fmt_nosubio

            ws_cartera.write(x, 0, cod, F["c"])
            ws_cartera.write(x, 1, r.get("nombre_ipress", ""), F["txt"])
            ws_cartera.write(x, 2, r.get("RED_SALUD", r.get("red", "")), F["c"])
            ws_cartera.write(x, 3, r.get("provincia", ""), F["txt"])
            ws_cartera.write(x, 4, r.get("distrito", ""), F["txt"])
            ws_cartera.write(x, 5, cat_f32, F["c"])
            ws_cartera.write(x, 6, cat_su, F["c"])
            ws_cartera.write(x, 7, estado_cat, fmt_est)

            # Formato según rol
            if rol_t == "CONSULTOR":
                fmt_rol = F.get("rol_consultor", F["c"])
            elif rol_t == "CONSULTANTE":
                fmt_rol = F.get("rol_consultante", F["c"])
            elif rol_t == "AMBOS":
                fmt_rol = F.get("rol_ambos", F["c"])
            else:
                fmt_rol = F.get("rol_ninguno", F["c"])
            ws_cartera.write(x, 8, rol_t, fmt_rol)

            ws_cartera.write(x, 9, nro_cat, F["c"])
            ws_cartera.write(x, 10, fec_cat, F["c"])
            ws_cartera.write(x, 11, sub_cat, fmt_subio if sub_cat == "SUBIÓ" else fmt_nosubio)

            ws_cartera.write(x, 12, nro_car, F["c"])
            ws_cartera.write(x, 13, fec_car, F["c"])
            ws_cartera.write(x, 14, sub_car, fmt_subio if sub_car == "SUBIÓ" else fmt_nosubio)

            ws_cartera.write(x, 15, en_f32, F.get("si_conv" if en_f32 == "SÍ" else "no_conv", F["c"]))
            ws_cartera.write(x, 16, r.get("estado", "ACTIVO"), F["c"])
            ws_cartera.write(x, 17, r.get("direccion", ""), F["txt"])

        ws_cartera.autofilter(T0, 0, T0 + len(df_car), len(car_h) - 1)
    else:
        ws_cartera.write(T0 + 1, 0, "No hay datos RENIPRESS disponibles para la región configurada", F["txt"])

    # Resumen ejecutivo inferior
    sr0 = T0 + 1 + (len(df_car) if not df_car.empty else 0) + 2
    ws_cartera.merge_range(sr0, 0, sr0, 4, f"RESUMEN EJECUTIVO — CATEGORIZACIÓN Y CARTERA DE SERVICIOS ({alcance_txt})", F["h"])
    ws_cartera.write(sr0 + 1, 0, f"Total IPRESS en el alcance ({alcance_txt}):", F["txt"])
    ws_cartera.write(sr0 + 1, 1, len(df_car), F["int"])

    n_en_f32 = sum(1 for _, r in df_car.iterrows() if str(r.get("CODIGO_UNICO", "")).zfill(8) in cods_f32) if not df_car.empty else 0
    n_no_f32 = len(df_car) - n_en_f32

    ws_cartera.write(sr0 + 2, 0, "  • IPRESS incluidas en Convenio Ficha 32 (prioridad en lista):", F["txt"])
    ws_cartera.write(sr0 + 2, 1, n_en_f32, F["int"])
    ws_cartera.write(sr0 + 3, 0, "  • IPRESS no conveniadas en Ficha 32 (resto de la red):", F["txt"])
    ws_cartera.write(sr0 + 3, 1, n_no_f32, F["int"])

    tot_cat_sub = sum(1 for _, r in df_car.iterrows() if r.get("subio_categorizacion") == "SUBIÓ") if not df_car.empty else 0
    tot_car_sub = sum(1 for _, r in df_car.iterrows() if r.get("subio_cartera") == "SUBIÓ") if not df_car.empty else 0
    tot_car_num = sum(1 for _, r in df_car.iterrows() if r.get("nro_cartera") and r.get("nro_cartera") != "—") if not df_car.empty else 0

    ws_cartera.write(sr0 + 4, 0, "IPRESS con Resolución de Categorización acreditada («SUBIÓ»):", F["txt"])
    ws_cartera.write(sr0 + 4, 1, tot_cat_sub, F["int"])
    ws_cartera.write(sr0 + 5, 0, "IPRESS con Resolución de Cartera de Servicios registrada:", F["txt"])
    ws_cartera.write(sr0 + 5, 1, tot_car_num, F["int"])
    ws_cartera.write(sr0 + 6, 0, "  • Resolución de Cartera de Servicios acreditada («SUBIÓ»):", F["txt"])
    ws_cartera.write(sr0 + 6, 1, tot_car_sub, F["int"])
    ws_cartera.write(sr0 + 7, 0, "  • Resolución de Cartera pendiente de subir documento («NO SUBIÓ»):", F["txt"])
    ws_cartera.write(sr0 + 7, 1, max(0, tot_car_num - tot_car_sub), F["int"])

    # Métricas de Telesalud
    con_tele_alcance = sum(1 for _, r in df_car.iterrows() if r.get("telemedicina") in ("CONSULTOR", "CONSULTANTE", "AMBOS")) if not df_car.empty else 0
    cons_alcance = sum(1 for _, r in df_car.iterrows() if r.get("telemedicina") == "CONSULTOR") if not df_car.empty else 0
    cont_alcance = sum(1 for _, r in df_car.iterrows() if r.get("telemedicina") == "CONSULTANTE") if not df_car.empty else 0
    ambos_alcance = sum(1 for _, r in df_car.iterrows() if r.get("telemedicina") == "AMBOS") if not df_car.empty else 0

    ws_cartera.write(sr0 + 8, 0, "IPRESS con Telesalud acreditadas en el alcance:", F["txt"])
    ws_cartera.write(sr0 + 8, 1, con_tele_alcance, F["int"])
    ws_cartera.write(sr0 + 9, 0, "  • Consultores (reciben teleinterconsultas):", F["txt"])
    ws_cartera.write(sr0 + 9, 1, cons_alcance, F["int"])
    ws_cartera.write(sr0 + 10, 0, "  • Consultantes (emiten teleinterconsultas):", F["txt"])
    ws_cartera.write(sr0 + 10, 1, cont_alcance, F["int"])
    ws_cartera.write(sr0 + 11, 0, "  • Ambos roles (mixtos):", F["txt"])
    ws_cartera.write(sr0 + 11, 1, ambos_alcance, F["int"])

    en_f32_con_tele = sum(1 for _, r in df_car.iterrows() if str(r.get("CODIGO_UNICO", "")).zfill(8) in cods_f32 and r.get("telemedicina") in ("CONSULTOR", "CONSULTANTE", "AMBOS")) if not df_car.empty else 0
    en_f32_sin_tele = max(0, n_en_f32 - en_f32_con_tele)
    pct_cob = round(100.0 * en_f32_con_tele / n_en_f32, 1) if n_en_f32 else 0.0

    ws_cartera.write(sr0 + 12, 0, "IPRESS en Convenio Ficha 32 con Telesalud acreditada:", F["txt"])
    ws_cartera.write(sr0 + 12, 1, en_f32_con_tele, F["int"])
    ws_cartera.write(sr0 + 13, 0, "IPRESS en Ficha 32 pendientes de acreditación Telesalud:", F["txt"])
    ws_cartera.write(sr0 + 13, 1, en_f32_sin_tele, F["int"])
    ws_cartera.write(sr0 + 14, 0, "Cobertura Telesalud acreditada en Ficha 32:", F["txt"])
    ws_cartera.write(sr0 + 14, 1, f"{pct_cob}%", F["c"])

    # ================================================================= HOJAS POR RED
    import re
    for prov, slug, disp in (M.redes(A) if redes else []):
        sub = A[A["RED_SALUD"] == prov].sort_values(["PUNTAJE", "PUNTAJE_ANT"], ascending=False).reset_index(drop=True)
        slug_seguro = re.sub(r"[\[\]:*?/\\]", "_", str(slug))
        nm = f"Red_{slug_seguro}"[:31]
        ws = wb.add_worksheet(nm)
        ws.set_tab_color("#94A3B8")
        k = len(sub)
        ws.write(0, 0, f"Red de Salud {disp} — Ficha N.° 32 — {M.ETQ['MES']}"
                       + (f" y plan {M.ETQ['PLAN']}" if plan else ""), F["t"])
        ws.write(1, 0, f"{k} IPRESS de primer nivel. Ordenado por Puntaje {M.ETQ['MES']}."
                       + (" Ver hoja Supuestos para el método de proyección." if plan else ""), F["st"])
        hc = ["N.°", "Cat.", "Establecimiento / IPRESS", "Población", "TI acum.", "TI meta", "TC acum.", "TC meta", "TM acum.", "TM meta",
              f"Pts Mes {M.MES_CORTE - 1}", f"Pts {M.ETQ['MES']}", "Δ Pts", f"Semáforo {M.ETQ['MES']}", "Categoría plan", "Ruta", f"TI {M.ETQ['PLAN']}", f"TC {M.ETQ['PLAN']}", f"TM {M.ETQ['PLAN']}",
              f"Pts {M.MES_ABR[M.MESES_AÑO]} (plan base)", f"Pts {M.MES_ABR[M.MESES_AÑO]} (techo)", "Acción sugerida"]
        if not plan:
            hc = hc[:14]        # solo situación actual: sin columnas de planificación
        T0 = 4
        for j, h in enumerate(hc):
            ws.write(T0, j, h, F["h"])
        ws.set_row(T0, 36)
        for i, r in sub.iterrows():
            x = T0 + 1 + i
            e = x + 1
            ws.write_number(x, 0, i + 1, F["c"]); ws.write(x, 1, r["NIVEL"], F["c"]); ws.write(x, 2, r["NOMBRE"], F["txt"])
            ws.write_number(x, 3, r["DENOMINADOR"], F["int"])
            for jj, s in enumerate(M.SERV):
                ws.write_number(x, 4 + 2 * jj, r[f"{s} REALIZADAS"], F["int"])
                ws.write_number(x, 5 + 2 * jj, r[f"{s} META"], F["int"])
            ws.write_number(x, 10, r["PUNTAJE_ANT"], F["d1"])
            ws.write_formula(x, 11, f"=ROUND(60*MIN(E{e},F{e})/F{e}+10*MIN(G{e},H{e})/H{e}+30*MIN(I{e},J{e})/J{e},1)", F["d1"], r["PUNTAJE"])
            ws.write_formula(x, 12, f"=L{e}-K{e}", F["dlt"], r["PUNTAJE"] - r["PUNTAJE_ANT"])
            ws.write_formula(x, 13, f'=IF(L{e}>=60,"VERDE",IF(L{e}>=30,"AMARILLO","ROJO"))', F["c"], r["SEMAFORO"])
            if plan:
                ws.write(x, 14, r["CATEGORIA"], F["c"]); ws.write(x, 15, r["RUTA"], F["c"])
                ws.write_number(x, 16, r["TI OBJ_ND"], F["int"]); ws.write_number(x, 17, r["TC OBJ_ND"], F["int"]); ws.write_number(x, 18, r["TM OBJ_ND"], F["int"])
                ws.write_number(x, 19, r["PUNTAJE_PLAN_BASE"], F["d1"]); ws.write_number(x, 20, r["PUNTAJE_PLAN"], F["d1"])
                ws.write(x, 21, r["ACCION"], F["txt"])
        TE = T0 + k  # última fila (0-based) de datos
        for col, w in zip(range(len(hc)), [5, 6, 40, 10, 8, 8, 8, 8, 8, 8, 9, 9, 8, 12, 13, 15, 9, 9, 9, 11, 11, 100]):
            ws.set_column(col, col, w)
        ws.freeze_panes(T0 + 1, 3)
        ws.autofilter(T0, 0, TE, len(hc) - 1)
        cf_sem(ws, f"N{T0 + 2}:N{TE + 1}")
        if plan:
            cf_cat(ws, f"O{T0 + 2}:O{TE + 1}")
        ws.conditional_format(f"L{T0 + 2}:L{TE + 1}", {"type": "data_bar", "bar_color": "#60A5FA", "bar_solid": True, "min_type": "num",
                                                       "min_value": 0, "max_type": "num", "max_value": 100})
        # panel derecho: resúmenes
        PC = 23  # columna X
        ws.set_column(PC, PC, 30)
        ws.set_column(PC + 1, PC + 4, 11)
        ws.write(T0, PC, "SEMÁFORO", F["h"]); ws.write(T0, PC + 1, f"Mes {M.MES_CORTE - 1}", F["h"]); ws.write(T0, PC + 2, M.ETQ['MES'], F["h"])
        for i, s in enumerate(("VERDE", "AMARILLO", "ROJO")):
            x = T0 + 1 + i
            ws.write(x, PC, s, F["kpi_l"])
            ws.write_number(x, PC + 1, int((sub["SEMAFORO_ANT"] == s).sum()), F["int"])
            ws.write_formula(x, PC + 2, f'=COUNTIF($N${T0 + 2}:$N${TE + 1},"{s}")', F["int"], int((sub["SEMAFORO"] == s).sum()))
        ws.write(T0 + 5, PC, "ESCENARIOS (puntaje promedio)", F["h"]); ws.write(T0 + 5, PC + 1, "Pts", F["h"])
        sc = [(f"Mes {M.MES_CORTE - 1}", f"=AVERAGE($K${T0 + 2}:$K${TE + 1})", sub["PUNTAJE_ANT"].mean()),
              (M.ETQ['MES'], f"=AVERAGE($L${T0 + 2}:$L${TE + 1})", sub["PUNTAJE"].mean())]
        if plan:
            sc += [(f"Proy. {M.ETQ['PROY']}", None, sub["PUNTAJE_PROY_OCT"].mean()),
                   (f"{M.MES_ABR[M.MESES_AÑO]}: tendencia", None, sub["PUNTAJE_DIC_TEND"].mean()),
                   (f"{M.MES_ABR[M.MESES_AÑO]}: plan base", f"=AVERAGE($T${T0 + 2}:$T${TE + 1})", sub["PUNTAJE_PLAN_BASE"].mean()),
                   (f"{M.MES_ABR[M.MESES_AÑO]}: techo", f"=AVERAGE($U${T0 + 2}:$U${TE + 1})", sub["PUNTAJE_PLAN"].mean())]
        for i, (lab, fm, val) in enumerate(sc):
            x = T0 + 6 + i
            ws.write(x, PC, lab, F["kpi_l"])
            if fm:
                ws.write_formula(x, PC + 1, fm, F["d1"], val)
            else:
                ws.write_number(x, PC + 1, round(val, 2), F["d1"])
        ws.write(T0 + 13, PC, "AVANCE POR SERVICIO", F["h"]); ws.write(T0 + 13, PC + 1, "% meta", F["h"]); ws.write(T0 + 13, PC + 2, "Esperado", F["h"])
        for i, (s, cr, cm) in enumerate((("TI", "E", "F"), ("TC", "G", "H"), ("TM", "I", "J"))):
            x = T0 + 14 + i
            ws.write(x, PC, f"{M.NOMBRE_SERV[s]} ({s})", F["kpi_l"])
            ws.write_formula(x, PC + 1, f"=SUM(${cr}${T0 + 2}:${cr}${TE + 1})/SUM(${cm}${T0 + 2}:${cm}${TE + 1})", F["pct"],
                             sub[f"{s} REALIZADAS"].sum() / sub[f"{s} META"].sum())
            ws.write_formula(x, PC + 2, f"={M.MES_CORTE}/{M.MESES_AÑO}", F["pct"], ESPERADO)
        if plan:
            ws.write(T0 + 18, PC, "CATEGORÍA DEL PLAN", F["h"]); ws.write(T0 + 18, PC + 1, "N.° IPRESS", F["h"])
            for i, c in enumerate(M.CATEGORIAS):
                x = T0 + 19 + i
                ws.write(x, PC, c, F["kpi_l"])
                ws.write_formula(x, PC + 1, f'=COUNTIF($O${T0 + 2}:$O${TE + 1},"{c}")', F["int"], int((sub["CATEGORIA"] == c).sum()))
        # gráficos
        g1 = wb.add_chart({"type": "column"})
        g1.add_series({"name": f"Mes {M.MES_CORTE - 1}", "categories": [nm, T0 + 1, PC, T0 + 3, PC], "values": [nm, T0 + 1, PC + 1, T0 + 3, PC + 1],
                       "points": [{"fill": {"color": "#D1FAE5"}}, {"fill": {"color": "#FEF9C7"}}, {"fill": {"color": "#FEE2E2"}}],
                       "border": {"color": "#475569"}, "data_labels": {"value": True}})
        g1.add_series({"name": M.ETQ['MES'], "categories": [nm, T0 + 1, PC, T0 + 3, PC], "values": [nm, T0 + 1, PC + 2, T0 + 3, PC + 2],
                       "points": [{"fill": {"color": "#86EFAC"}}, {"fill": {"color": "#FDE047"}}, {"fill": {"color": "#FCA5A5"}}],
                       "border": {"color": "#475569"}, "data_labels": {"value": True}})
        g1.set_title({"name": f"Semáforo Mes {M.MES_CORTE - 1} vs {M.ETQ['MES']} — {disp}", "name_font": {"size": 10}})
        g1.set_legend({"position": "bottom"})
        g1.set_size({"width": 470, "height": 280})
        ws.insert_chart(T0 + 24, PC, g1)
        if plan:
            g2 = wb.add_chart({"type": "column"})
            g2.add_series({"name": "Puntaje promedio", "categories": [nm, T0 + 6, PC, T0 + 11, PC], "values": [nm, T0 + 6, PC + 1, T0 + 11, PC + 1],
                           "points": [{"fill": {"color": c}} for c in ["#94A3B8", "#475569", "#93C5FD", "#FBBF24", "#34D399", "#059669"]],
                           "data_labels": {"value": True, "num_format": "0.0", "font": {"bold": True}}, "gap": 60})
            g2.set_title({"name": f"Proyección del puntaje promedio — {disp}", "name_font": {"size": 10}})
            g2.set_y_axis({"min": 0, "max": 100})
            g2.set_legend({"none": True})
            g2.set_size({"width": 470, "height": 280})
            ws.insert_chart(T0 + 24, PC + 5, g2)
        g3 = wb.add_chart({"type": "column"})
        g3.add_series({"name": "% de la meta anual", "categories": [nm, T0 + 14, PC, T0 + 16, PC], "values": [nm, T0 + 14, PC + 1, T0 + 16, PC + 1],
                       "fill": {"color": "#60A5FA"}, "border": {"color": "#1D4ED8"}, "data_labels": {"value": True, "num_format": "0.0%"}})
        g3.add_series({"name": f"Esperado a {M.ETQ['MES_LARGO'].lower()}", "categories": [nm, T0 + 14, PC, T0 + 16, PC], "values": [nm, T0 + 14, PC + 2, T0 + 16, PC + 2],
                       "type": "line", "line": {"color": "#DC2626", "dash_type": "dash", "width": 2}})
        g3.set_title({"name": f"Avance vs meta anual — {disp}", "name_font": {"size": 10}})
        g3.set_y_axis({"num_format": "0%", "min": 0})
        g3.set_legend({"position": "bottom"})
        g3.set_size({"width": 470, "height": 280})
        ws.insert_chart(T0 + 39, PC, g3)
        if plan:
            g4 = wb.add_chart({"type": "pie"})
            g4.add_series({"name": "Categorías", "categories": [nm, T0 + 19, PC, T0 + 22, PC], "values": [nm, T0 + 19, PC + 1, T0 + 22, PC + 1],
                           "points": [{"fill": {"color": c}} for c in ["#86EFAC", "#BEF264", "#FDE047", "#FCA5A5"]],
                           "data_labels": {"value": True, "font": {"size": 9, "bold": True}}})
            g4.set_title({"name": f"Categorías del plan — {disp}", "name_font": {"size": 10}})
            g4.set_size({"width": 470, "height": 280})
            ws.insert_chart(T0 + 39, PC + 5, g4)
        g5 = wb.add_chart({"type": "bar"})
        nt = min(15, k)
        g5.add_series({"name": f"Pts Mes {M.MES_CORTE - 1}", "categories": [nm, T0 + 1, 2, T0 + nt, 2], "values": [nm, T0 + 1, 10, T0 + nt, 10],
                       "fill": {"color": "#CBD5E1"}, "border": {"color": "#64748B"}})
        g5.add_series({"name": f"Pts {M.ETQ['MES']}", "categories": [nm, T0 + 1, 2, T0 + nt, 2], "values": [nm, T0 + 1, 11, T0 + nt, 11],
                       "fill": {"color": "#60A5FA"}, "border": {"color": "#1D4ED8"}, "data_labels": {"value": True, "num_format": "0.0", "font": {"size": 7}}})
        if plan:
            g5.add_series({"name": f"{M.MES_ABR[M.MESES_AÑO]}: plan base", "categories": [nm, T0 + 1, 2, T0 + nt, 2], "values": [nm, T0 + 1, 19, T0 + nt, 19],
                           "fill": {"color": "#34D399"}, "border": {"color": "#059669"}})
        g5.set_title({"name": f"Top {nt} IPRESS por puntaje — {disp}", "name_font": {"size": 10}})
        g5.set_y_axis({"reverse": True, "num_font": {"size": 7}})
        g5.set_x_axis({"min": 0, "max": 100})
        g5.set_legend({"position": "bottom"})
        g5.set_size({"width": 960, "height": 420})
        ws.insert_chart(T0 + 54, PC, g5)

    # ---------------------------------------------------------------- hoja Evolución
    if serie:
        import ficha32_evolucion as EVO
        EVO.escribir_hoja(wb, F, serie, M.MES_CORTE, ws=ws_evo, desde=evo_desde)

    wb.close()
    return OUT


if __name__ == "__main__":
    print("OK", build())
