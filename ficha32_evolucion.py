# -*- coding: utf-8 -*-
"""Evolución mes a mes — Ficha N.° 32 (Telemedicina).

La API devuelve el acumulado ene→mes para cada mes pedido, por lo que la producción
de cada mes individual se obtiene por diferencia:

    producción(mes m) = acumulado(m) − acumulado(m−1)

El puntaje de cada mes se calcula con las metas de ese mismo mes. Conviene recordar que
el HIS re-digitaliza retroactivamente, así que un mes pedido hoy puede dar cifras ligeramente
mayores que las que se informaron en su momento.

Salida: la tabla que alimenta la hoja "Evolución" del libro Excel.
"""
import numpy as np
import pandas as pd

import ficha32_modelo as M

COL = {"TI": "#1F3E8F", "TC": "#F59E0B", "TM": "#006B5F"}


def _num(v, defecto=0):
    """Convierte a número apto para xlsxwriter; NaN/inf -> `defecto`."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return defecto
    return f if np.isfinite(f) else defecto


def serie(diresa_token, hasta, anio=2026, refrescar=False, avance=None, desde=1):
    """{(mes): (df_A, df_B, ctx)} de `desde` hasta `hasta`. Aprovecha la caché local.

    Para que la producción del primer mes se pueda diferenciar, conviene pedir
    `desde - 1` también (ver construir()); el generador lo hace por defecto.
    """
    import ficha32_api as api
    sal = {}
    for m in range(max(1, int(desde) - 1), int(hasta) + 1):
        if avance:
            avance(m, int(hasta))
        sal[m] = api.obtener(diresa_token, m, anio=anio, refrescar=refrescar)
    return sal


def _puntaje(df_realizadas, df_metas):
    """Puntaje de logro (0-100) con pesos 60 TI / 10 TC / 30 TM, tope en la meta.

    Los dos DataFrames deben estar indexados por código IPRESS y tener columnas
    "TI", "TC", "TM" (realizadas y metas respectivamente).
    """
    p = 0.0
    for s in M.SERV:
        r = pd.to_numeric(df_realizadas[s], errors="coerce").fillna(0)
        m = pd.to_numeric(df_metas[s], errors="coerce")
        p = p + M.PESO[s] * np.minimum(r, m) / m
    return p


def construir(serie_meses, mes_corte=None, desde=None):
    """Construye las tablas de evolución a partir de la serie descargada.

    `desde` (opcional) fija el primer mes mostrado; no hace falta que sea el mes 1.
    Si la serie incluye también el mes anterior a `desde`, ese mes se usa solo como
    base para diferenciar la producción (acumulado(m) − acumulado(m−1)) y no se
    muestra en la hoja.

    Devuelve (detalle, resumen_red, meses):
      * detalle    : una fila por IPRESS con la producción de cada mes y su puntaje.
      * resumen_red: promedio de puntaje y totales de producción por red y mes.
      * meses      : lista de meses incluidos.
    """
    serie_meses = {int(k): v for k, v in serie_meses.items()}
    meses = [m for m in sorted(serie_meses) if desde is None or m >= int(desde)]
    if not meses:
        raise ValueError("La serie está vacía.")
    # Mes anterior al rango: sirve de base para la diferencia del primer mes.
    pre = meses[0] - 1 if meses[0] > 1 and (meses[0] - 1) in serie_meses else None
    todos = [pre] + meses if pre else meses
    mes_ref = int(mes_corte or meses[-1])
    A_ref, B_ref, _ = serie_meses[meses[-1]] if mes_ref not in serie_meses else serie_meses[mes_ref]

    # Padrón y metas vigentes (las metas del primer nivel son anuales e invariantes)
    base = pd.concat([A_ref, B_ref], ignore_index=True)
    base["_COD"] = (base["ESTABLECIMIENTO / IPRESS"].str.split(" - ").str[1].str.strip())

    filas = {s: {} for s in M.SERV}              # producción del mes
    acums = {s: {} for s in M.SERV}              # acumulado
    punt = {}                                    # puntaje del mes

    for m in todos:
        A_m, B_m, _ = serie_meses[m]
        d_m = pd.concat([A_m, B_m], ignore_index=True)
        d_m["_COD"] = d_m["ESTABLECIMIENTO / IPRESS"].str.split(" - ").str[1].str.strip()
        d_m = d_m.set_index("_COD")
        # metas y realizadas del propio mes (en el Reporte B la meta crece con el denominador)
        real = d_m[[f"{s} REALIZADAS" for s in M.SERV]].copy()
        real.columns = list(M.SERV)
        meta = d_m[[f"{s} META" for s in M.SERV]].copy()
        meta.columns = list(M.SERV)
        for s in M.SERV:
            acums[s][m] = pd.to_numeric(real[s], errors="coerce").fillna(0)
        if m not in meses:
            continue                            # mes base: solo aporta el acumulado
        punt[m] = _puntaje(real, meta).round(1)
        # Logro B (hospitales consultores): solo Teleinterconsulta con peso 100;
        # el resto de servicios no aplica y sus metas son "--".
        if "LOGRO" in d_m.columns:
            es_b = d_m["LOGRO"].astype(str).str.upper().str.startswith("B")
            if bool(es_b.any()):
                r_ti = pd.to_numeric(real.loc[es_b, "TI"], errors="coerce").fillna(0)
                m_ti = pd.to_numeric(meta.loc[es_b, "TI"], errors="coerce").mask(lambda s: s <= 0)
                punt[m].loc[es_b] = (100 * np.minimum(r_ti, m_ti) / m_ti).round(1)
        for s in M.SERV:
            if m - 1 in acums[s]:
                filas[s][m] = (acums[s][m] - acums[s][m - 1]).clip(lower=0)
            else:                               # mes 1: el acumulado de enero es su producción
                filas[s][m] = acums[s][m]

    info = base.set_index("_COD")[["DIRESA", "RED", "PROVINCIA", "ESTABLECIMIENTO / IPRESS",
                                   "LOGRO", "DENOMINADOR"]].copy()
    info["RED"] = M.resolver_red_salud(info, diresa=M.DIRESA)
    info["NIVEL"] = info["ESTABLECIMIENTO / IPRESS"].str.split().str[0]
    info["NOMBRE"] = info["ESTABLECIMIENTO / IPRESS"].map(M.short_name)
    info["CODIGO"] = info["ESTABLECIMIENTO / IPRESS"].map(M.codigo)
    nombres = M.NOMBRES_HOSPITAL
    info["NOMBRE"] = [nombres.get(c, n) for c, n in zip(info["CODIGO"], info["NOMBRE"])]

    for m in meses:
        for s in M.SERV:
            info[f"{s} PROD M{m}"] = info.index.map(filas[s][m])
            info[f"{s} ACUM M{m}"] = info.index.map(acums[s][m])
        info[f"PUNTAJE M{m}"] = info.index.map(punt[m])

    info["PUNTAJE INI"] = info[f"PUNTAJE M{meses[0]}"]
    info["PUNTAJE FINAL"] = info[f"PUNTAJE M{meses[-1]}"]
    info["DELTA"] = (info["PUNTAJE FINAL"] - info["PUNTAJE INI"]).round(1)
    info["SEMAFORO FINAL"] = info["PUNTAJE FINAL"].apply(M.clasificar)
    info["SEMAFORO INI"] = info["PUNTAJE INI"].apply(M.clasificar)
    info = info.reset_index(drop=True)

    # Respetar el alcance activo (M.FILTRO): la evolución lista los mismos eess
    # que el resto del informe (toda la DIRESA, una red o una lista de códigos).
    red_f, cod_f = M.FILTRO.get("red"), M.FILTRO.get("codigos") or []
    if red_f:
        pat = str(red_f).upper()
        info = info[info["RED"].astype(str).str.upper().str.contains(pat, regex=False, na=False)]
    if cod_f:
        cods = {str(c).strip().zfill(8) for c in cod_f}
        info = info[info["CODIGO"].astype(str).str.zfill(8).isin(cods)]
    info = info.reset_index(drop=True)

    # Resumen por red de salud
    resumen = []
    for red, sub in info.groupby("RED"):
        fila = {"RED": red, "IPRESS": len(sub)}
        for m in meses:
            fila[f"PUNT M{m}"] = round(float(sub[f"PUNTAJE M{m}"].mean()), 1)
            for s in M.SERV:
                fila[f"{s} M{m}"] = int(sub[f"{s} PROD M{m}"].sum())
        resumen.append(fila)
    resumen = pd.DataFrame(resumen).sort_values(f"PUNT M{meses[-1]}", ascending=False).reset_index(drop=True)
    return info, resumen, meses


def escribir_hoja(wb, F, serie_meses, mes_corte=None, ws=None, desde=None):
    """Escribe la hoja "Evolución" (xlsxwriter). Devuelve la hoja.

    `ws` permite indicar una hoja ya creada (para posicionarla donde se quiera dentro
    de un libro mayor); si se omite, se crea al final del libro.
    `desde` fija el primer mes de la serie mostrada (ver construir()).
    """
    info, resumen, meses = construir(serie_meses, mes_corte, desde=desde)
    if ws is None:
        ws = wb.add_worksheet("Evolución")
    ws.set_tab_color("#0EA5E9")

    c = 0
    fijas = [("Red de Salud", 20), ("Cat.", 6), ("Código", 10), ("Establecimiento / IPRESS", 42),
             ("Población", 10), ("Puntaje inicial", 11), ("Puntaje final", 11), ("Δ Puntaje", 9)]
    for h, w in fijas:
        ws.write(0, c, h, F["h"]); ws.set_column(c, c, w); c += 1
    cols_mes = []
    for m in meses:
        for s in M.SERV:
            cols_mes.append((m, s, f"{s} PROD M{m}"))
    for m, s, _ in cols_mes:
        ws.write(0, c, f"{s} M{m}", F["h"]); ws.set_column(c, c, 7); c += 1
    for m in meses:
        ws.write(0, c, f"Puntaje M{m}", F["h"]); ws.set_column(c, c, 10); c += 1
    ws.set_row(0, 30)
    nfilas = len(info)

    for i, r in info.iterrows():
        x, e = i + 1, i + 2
        ws.write(x, 0, r["RED"], F["txt"]); ws.write(x, 1, r["NIVEL"], F["c"])
        ws.write(x, 2, r["CODIGO"], F["c"]); ws.write(x, 3, r["NOMBRE"], F["txt"])
        ws.write_number(x, 4, _num(r["DENOMINADOR"]), F["int"])
        ws.write_number(x, 5, _num(r["PUNTAJE INI"]), F["d1"])
        ws.write_number(x, 6, _num(r["PUNTAJE FINAL"]), F["d1"])
        ws.write_formula(x, 7, f"=G{e}-F{e}", F["dlt"], _num(r["DELTA"]))
        j = 8
        for m, s, col in cols_mes:
            ws.write_number(x, j, _num(r[col]), F["int"]); j += 1
        for m in meses:
            ws.write_number(x, j, _num(r[f"PUNTAJE M{m}"]), F["d1"]); j += 1

    ult = nfilas + 1
    ws.freeze_panes(1, 4)
    ws.autofilter(0, 0, nfilas, c - 1)
    verde = wb.add_format({"font_color": "#166534", "bold": True})
    ws.conditional_format(f"H2:H{ult + 1}", {"type": "cell", "criteria": ">", "value": 0,
                                             "format": verde})
    # Puntaje final (col. G) coloreado como el semáforo del tablero
    ws.conditional_format(f"G2:G{ult + 1}", {"type": "cell", "criteria": ">=", "value": 60,
                                             "format": wb.add_format({"bg_color": "#DCFCE7"})})
    ws.conditional_format(f"G2:G{ult + 1}", {"type": "cell", "criteria": "between",
                                             "minimum": 30, "maximum": 59.999,
                                             "format": wb.add_format({"bg_color": "#FEF3C7"})})
    ws.conditional_format(f"G2:G{ult + 1}", {"type": "cell", "criteria": "<", "value": 30,
                                             "format": wb.add_format({"bg_color": "#FEE2E2"})})

    # Bloque de resumen por red, debajo de la tabla
    fila0 = ult + 2
    ws.write(fila0, 0, f"Resumen por Red de Salud — evolución del puntaje medio ({M.ETQ['RANGO']})", F["t"])
    fila0 += 1
    hdr = ["Red de Salud", "N.° IPRESS"] + [f"Puntaje M{m}" for m in meses] + \
          [f"Total {s}" for s in M.SERV]
    for j, h in enumerate(hdr):
        ws.write(fila0, j, h, F["h2"])
    for i, r in resumen.iterrows():
        x = fila0 + 1 + i
        ws.write(x, 0, r["RED"], F["txt"]); ws.write_number(x, 1, r["IPRESS"], F["int"])
        for j, m in enumerate(meses):
            v = r[f"PUNT M{m}"]
            if pd.isna(v):                      # red sin datos en ese mes
                ws.write_blank(x, 2 + j, None, F["d1"])
            else:
                ws.write_number(x, 2 + j, _num(v), F["d1"])
        for j, s in enumerate(M.SERV):
            ws.write_number(x, 2 + len(meses) + j, int(r[f"{s} M{meses[-1]}"]), F["int"])
    return ws


def formatos(wb):
    """Conjunto de formatos que usa la hoja «Evolución» (xlsxwriter)."""
    base = {"font_name": "Arial", "font_size": 9}
    borde, color_borde = 1, "#94A3B8"
    return {
        "h": wb.add_format({**base, "bold": True, "bg_color": "#1E3A5F", "font_color": "#FFFFFF",
                            "border": borde, "border_color": color_borde, "align": "center",
                            "valign": "vcenter", "text_wrap": True}),
        "h2": wb.add_format({**base, "bold": True, "bg_color": "#E2E8F0", "border": borde,
                             "border_color": color_borde, "align": "center", "valign": "vcenter",
                             "text_wrap": True}),
        "txt": wb.add_format({**base, "border": borde, "border_color": "#CBD5E1"}),
        "c": wb.add_format({**base, "border": borde, "border_color": "#CBD5E1", "align": "center"}),
        "int": wb.add_format({**base, "border": borde, "border_color": "#CBD5E1",
                              "num_format": "#,##0", "align": "center"}),
        "d1": wb.add_format({**base, "border": borde, "border_color": "#CBD5E1",
                             "num_format": "0.0", "align": "center"}),
        "dlt": wb.add_format({**base, "border": borde, "border_color": "#CBD5E1",
                              "num_format": "+0.0;-0.0;0.0", "align": "center"}),
        "t": wb.add_format({**base, "bold": True, "font_size": 13, "font_color": "#0F172A"}),
    }


def crear_libro(ruta, serie_meses, mes_corte=None, wb=None):
    """Crea un libro Excel nuevo con la hoja «Evolución» (xlsxwriter). Devuelve la ruta.

    Se escribe nativamente con xlsxwriter en lugar de reabrir el libro principal con
    openpyxl: openpyxl degrada las barras de datos (78 → 39) y borra las extensiones
    de formato (26 → 0) al volver a guardar. Para que la evolución viva dentro del
    libro principal, llame generar_excel.py a escribir_hoja() con su propio workbook.
    """
    import xlsxwriter

    propios = wb is None
    if propios:
        wb = xlsxwriter.Workbook(ruta)
        wb.set_calc_mode("auto")
        F = formatos(wb)
    else:
        F = None
    escribir_hoja(wb, F, serie_meses, mes_corte)
    if propios:
        wb.close()
    return ruta
