# -*- coding: utf-8 -*-
"""Modelo de datos y de proyección — Ficha Técnica N.° 32 (Telemedicina).

Fuente única de verdad para los Word por red y para el libro Excel.

Fuentes de datos (elegible con `configurar(fuente=...)`):
  * "api"    → descarga automática desde el dashboard (ficha32_api), sin exportar nada a mano.
  * "xls"    → los .xls exportados a mano del tablero (modo original / sin internet).
  * "cache"  → la API solo si ya hay descargas locales en cache_api/.

Hallazgos metodológicos verificados con los reportes de los meses 7 y 8:
  * Las METAS son anuales y no cambian entre meses (Reporte A: idénticas en 371/371 IPRESS).
  * Las atenciones REALIZADAS son ACUMULADAS ene→mes (nunca disminuyen entre mes 7 y mes 8).
  * Producción del mes N = REALIZADAS(mes N) − REALIZADAS(mes N−1).
  * Reporte B (hospitales): la meta TI = máx(100, redondeo(40/10000 × atenciones HIS acumuladas)),
    por lo que crece con el avance del año.
  * El HIS re-digitaliza retroactivamente: la API devuelve el acumulado ya corregido, por lo que
    los meses anteriores pueden ser MAYORES que en el .xls exportado en su momento (verificado:
    mes 7 con +32 TI, +8 TC y +117 TM en 42 IPRESS de DIRESA Puno).
"""
import os
import re
import numpy as np
import pandas as pd

import config_manager as CM

BASE = CM.get_app_dir()
# Copias locales de los .xls exportados a mano del tablero (modo "xls", sin internet).
XLS_BASE = CM.resolver_recurso("xls_base")

F_A8 = os.path.join(XLS_BASE, "Reporte_DITEL_2026_agosto.xls")
F_B8 = os.path.join(XLS_BASE, "Reporte_DITEL_agosto_B.xls")
F_A7 = os.path.join(XLS_BASE, "Reporte_DITEL_2026_mes7_A.xls")
F_B7 = os.path.join(XLS_BASE, "Reporte_DITEL_2026_mes7_B.xls")

# Configuración dinámica y extensible de servicios desde config_manager (preparado para 2026, 2027 y nuevos servicios)
SERV = CM.get_servicio_codigos()
PESO = CM.get_pesos()
NOMBRE_SERV = CM.get_servicio_nombres()
NUM = ["DENOMINADOR"] + [f"{s} REALIZADAS" for s in SERV] + [f"{s} META" for s in SERV]

_CFG_FORMULAS = CM.get_metas_config()
_CFG_PLAN = CM.get_plan_config()

ANIO = 2026               # año del convenio de gestión
DIRESA = "PUNO"           # dirección regional activa (por defecto, la del proyecto)
FUENTE = "xls"            # "api" | "cache" | "xls"
MES_CORTE = 8             # mes de corte del informe (acumulado ene→mes)
MES_INI_PLAN = 11         # inicio de la ventana del plan (por defecto Noviembre)
MES_FIN_PLAN = 12         # fin de la ventana del plan (por defecto Diciembre)
MES_PROY = 10             # horizonte de la proyección previa (ej. cierre de octubre)
MESES_AÑO = 12
MESES_PLAN = max(1, MES_FIN_PLAN - MES_INI_PLAN + 1)     # ventana del plan
MESES_PROY_PREVIOS = max(0, MES_PROY - MES_CORTE)

MES_TXT = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
           7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
MES_ABR = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun", 7: "Jul", 8: "Ago",
           9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}


def _rangos():
    """Etiquetas y cantidades derivadas del mes de corte y ventana del plan."""
    mes_corte = min(12, max(1, MES_CORTE))
    ini_plan = min(12, max(1, MES_INI_PLAN))
    fin_plan = min(12, max(ini_plan, MES_FIN_PLAN))
    meses_plan = max(1, fin_plan - ini_plan + 1)

    mes_proy = min(12, max(mes_corte, MES_PROY))
    meses_prev = max(0, mes_proy - mes_corte)

    if ini_plan == fin_plan:
        plan_abr = MES_ABR[ini_plan]
        plan_largo = MES_TXT[ini_plan]
    else:
        plan_abr = f"{MES_ABR[ini_plan]}–{MES_ABR[fin_plan]}"
        plan_largo = f"{MES_TXT[ini_plan]}–{MES_TXT[fin_plan]}"

    return {
        "MES_CORTE": mes_corte, "MES_PROY": mes_proy, "MESES_PLAN": meses_plan,
        "MES_INI_PLAN": ini_plan, "MES_FIN_PLAN": fin_plan,
        "MESES_PROY_PREVIOS": meses_prev, "MESES_AÑO": MESES_AÑO, "ANIO": ANIO,
        "MES": f"Mes {mes_corte}",
        "MES_LARGO": MES_TXT[mes_corte],
        "MES_ABR": MES_ABR[mes_corte],
        "RANGO": f"ene–{MES_ABR[mes_corte]}",
        "PLAN": plan_abr,
        "PLAN_LARGO": plan_largo,
        "PLAN_INI_TXT": MES_TXT[ini_plan],
        "PLAN_FIN_TXT": MES_TXT[fin_plan],
        "PROY": MES_ABR[mes_proy],
        "PREVIO": MES_ABR[mes_proy - 1] if meses_prev and mes_proy > 1 else "",
        "HASTA": MES_TXT[mes_proy],
        "ANIO_TXT": str(ANIO),
    }


ETQ = _rangos()          # etiquetas de presentación, se refresca con configurar()
SERIE = None             # serie ya descargada {mes: (A, B, ctx)}; la pone generar_todo.py


def configurar(mes=None, anio=None, diresa=None, fuente=None, mes_proy=None,
               mes_ini_plan=None, mes_fin_plan=None, **_):
    """Ajusta el periodo, ventana de planificación y fuente de datos de todo el modelo.

    mes          : mes de corte (1..12). Por defecto 8.
    anio         : año del convenio. Por defecto 2026.
    diresa       : p. ej. "PUNO". Para la API se traduce al token DIRESA_PUNO.
    fuente       : "api" (descarga automática), "xls" (archivos exportados), "cache" (sin red) o "renipress" (Cartera SUSALUD).
    mes_proy     : horizonte de proyección previa (por defecto mes_ini_plan - 1 o mes corte).
    mes_ini_plan : mes de inicio del plan de cumplimiento (1..12). Por defecto 11 (Nov).
    mes_fin_plan : mes de fin del plan de cumplimiento (1..12). Por defecto 12 (Dic).
    """
    global MES_CORTE, MESES_PROY_PREVIOS, MESES_PLAN, MESES_AÑO, ANIO, DIRESA, FUENTE, MES_PROY, ETQ
    global MES_INI_PLAN, MES_FIN_PLAN
    global SERV, PESO, NOMBRE_SERV, NUM, PISO_CAP_MES, PCT_CAP_META, RATIO_CERCA, FRAC_ARRANQUE

    # Recargar parámetros de config_manager por si config.json cambió
    SERV = CM.get_servicio_codigos()
    PESO = CM.get_pesos()
    NOMBRE_SERV = CM.get_servicio_nombres()
    NUM = ["DENOMINADOR"] + [f"{s} REALIZADAS" for s in SERV] + [f"{s} META" for s in SERV]

    cfg_form = CM.get_metas_config()
    cfg_plan = CM.get_plan_config()
    PISO_CAP_MES = cfg_plan.get("piso_capacidad_mes", 12)
    PCT_CAP_META = cfg_plan.get("pct_capacidad_meta", 0.05)
    RATIO_CERCA = cfg_plan.get("ratio_esfuerzo_cerca", 1 / 3)
    FRAC_ARRANQUE = cfg_plan.get("fraccion_arranque", 1 / 3)

    if mes is not None:
        MES_CORTE = int(mes)
    if anio is not None:
        ANIO = int(anio)
    if diresa is not None:
        DIRESA = diresa
    if fuente is not None:
        FUENTE = fuente

    # Ventana de Planificación configurable
    default_ini = int(cfg_plan.get("mes_inicio_plan", 11))
    default_fin = int(cfg_plan.get("mes_fin_plan", 12))

    if mes_ini_plan is not None:
        MES_INI_PLAN = int(mes_ini_plan)
    elif MES_CORTE >= default_ini:
        MES_INI_PLAN = min(12, MES_CORTE + 1)
    else:
        MES_INI_PLAN = default_ini

    if mes_fin_plan is not None:
        MES_FIN_PLAN = int(mes_fin_plan)
    else:
        MES_FIN_PLAN = default_fin

    if MES_INI_PLAN > MES_FIN_PLAN:
        MES_INI_PLAN, MES_FIN_PLAN = MES_FIN_PLAN, MES_INI_PLAN

    MESES_AÑO = 12
    MESES_PLAN = max(1, MES_FIN_PLAN - MES_INI_PLAN + 1)

    # Horizonte de proyección previa antes del plan
    if mes_proy is not None:
        MES_PROY = int(mes_proy)
    elif MES_INI_PLAN > MES_CORTE + 1:
        MES_PROY = MES_INI_PLAN - 1
    else:
        MES_PROY = MES_CORTE

    MESES_PROY_PREVIOS = max(0, MES_PROY - MES_CORTE)
    ETQ = _rangos()
    # Una configuración nueva invalida cualquier serie descargada anterior.
    globals()["SERIE"] = None
    FILTRO["red"] = None
    FILTRO["codigos"] = []
    return ETQ


def _token_diresa():
    """Token del tablero para la DIRESA activa ("PUNO" -> "DIRESA_PUNO")."""
    try:
        import ficha32_api as api
        return api.TOKEN_POR_NOMBRE.get(DIRESA, f"DIRESA_{DIRESA}")
    except Exception:  # noqa: BLE001
        return f"DIRESA_{DIRESA}"


def cargar():
    """(a_corte, b_corte, a_prev, b_prev) según la fuente configurada. Deja el contexto en CTX.

    Si ya hay una serie descargada (M.SERIE, la pone el orquestador generar_todo.py),
    se reutiliza y no se vuelve a llamar a la API.
    """
    global CTX
    if FUENTE in ("api", "cache"):
        import ficha32_api as api
        serie = SERIE or {}
        if MES_CORTE in serie:
            a_corte, b_corte, ctx = serie[MES_CORTE]
            a_prev, b_prev, _ = serie.get(MES_CORTE - 1, (a_corte.iloc[0:0].copy(),
                                                          b_corte.iloc[0:0].copy(), {}))
        else:
            a_corte, b_corte, ctx = api.obtener(_token_diresa(), MES_CORTE, anio=ANIO,
                                                refrescar=(FUENTE == "api"))
            if MES_CORTE > 1:
                a_prev, b_prev, _ = api.obtener(_token_diresa(), MES_CORTE - 1, anio=ANIO,
                                                refrescar=(FUENTE == "api"))
            else:
                # Mes 1: no hay mes anterior; se devuelve vacío para que el modelo trabaje con cero.
                a_prev, b_prev = a_corte.iloc[0:0].copy(), b_corte.iloc[0:0].copy()
        for d in (a_corte, a_prev, b_corte, b_prev):
            d["ESTABLECIMIENTO / IPRESS"] = d["ESTABLECIMIENTO / IPRESS"].astype(str).str.strip()
            # El indicador B trae "--" como meta de TC/TM, lo que deja la columna como
            # dtype object; se fuerza a numérico (NaN donde no hay meta).
            for col in NUM:
                d[col] = pd.to_numeric(d[col], errors="coerce")
        CTX = {**ctx, "fuente": FUENTE}
        return a_corte, b_corte, a_prev, b_prev, ctx

    if FUENTE == "renipress":
        cartera, resumen = obtener_cartera_renipress()
        a_corte = pd.DataFrame()
        clasi = cartera["clasificacion"].replace("", "I-1").fillna("I-1")
        a_corte["ESTABLECIMIENTO / IPRESS"] = clasi + " - " + cartera["CODIGO_UNICO"] + " - " + cartera["nombre_ipress"]
        a_corte["NOMBRE"] = cartera["nombre_ipress"]
        a_corte["DIRESA"] = cartera["diresa"]
        a_corte["PROVINCIA"] = cartera["provincia"]
        a_corte["RED_SALUD"] = cartera["RED_SALUD"]
        a_corte["TELESALUD_ROL"] = cartera["telemedicina"]
        a_corte["ESTADO"] = cartera["estado"]
        for col in NUM:
            a_corte[col] = 1.0 if "META" in col or col == "DENOMINADOR" else 0.0
        b_corte = a_corte.iloc[0:0].copy()
        a_prev = a_corte.iloc[0:0].copy()
        b_prev = b_corte.iloc[0:0].copy()
        CTX = {
            "fuente": "renipress",
            "scope_vista": "REGIONAL",
            "nombre_diresa": DIRESA,
            "total_establecimientos": len(cartera),
            "resumen_resolucion": resumen
        }
        return a_corte, b_corte, a_prev, b_prev, CTX

    a8, b8 = leer(F_A8), leer(F_B8)
    a7, b7 = leer(F_A7), leer(F_B7)
    for d in (a8, b8, a7, b7):
        for col in NUM:
            d[col] = pd.to_numeric(d[col], errors="coerce")
    CTX = {"scope_vista": "XLS", "nombre_diresa": DIRESA, "fuente": "xls"}
    return a8, b8, a7, b7, CTX


def obtener_cartera_renipress(token=None, rol=None, refrescar=False):
    """Obtiene el DataFrame oficial de SUSALUD RENIPRESS enriquecido con Categorización y Cartera de Servicios.

    Resuelve la Red de Salud para cada establecimiento y estandariza nombres de columnas.
    """
    import ficha32_api as api
    tok = token or _token_diresa()
    raw = api.obtener_datos_renipress(tok, mes=MES_CORTE, anio=ANIO, rol=rol, refrescar=refrescar)
    df = api.datos_a_dataframe(raw.get("todos_establecimientos") or raw.get("establecimientos", []))
    resumen = dict(raw.get("resumen_resolucion", {}))
    if not df.empty:
        df["RED_SALUD"] = resolver_red_salud(df, diresa=DIRESA)
        df["CODIGO_UNICO"] = df["id_ipress"].astype(str).str.zfill(8)

        # Cruce con resoluciones de Categorización y Cartera
        cods = df["CODIGO_UNICO"].tolist()
        mapa_res = api.obtener_resoluciones_renipress(cods, token=tok)

        df["categoria_susalud"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("categoria_susalud") or ""))
        df["nro_categorizacion"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("nro_categorizacion") or "—"))
        df["fecha_categorizacion"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("fecha_categorizacion") or "—"))
        df["subio_categorizacion"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("subio_categorizacion") or "NO SUBIÓ"))
        df["archivo_categorizacion"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("archivo_categorizacion") or ""))

        df["nro_cartera"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("nro_cartera") or "—"))
        df["fecha_cartera"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("fecha_cartera") or "—"))
        df["subio_cartera"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("subio_cartera") or "NO SUBIÓ"))
        df["archivo_cartera"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("archivo_cartera") or ""))

        df["nro_resolucion"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("nro_resolucion") or "—"))
        df["fecha_resolucion"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("fecha_resolucion") or "—"))
        df["variacion_resolucion"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("variacion_resolucion") or "NO SUBIÓ"))
        df["subio_resolucion"] = df["CODIGO_UNICO"].apply(lambda c: (mapa_res.get(c, {}).get("subio_resolucion") or "NO SUBIÓ"))

        # Actualizar resumen de resoluciones
        resumen["cat_subio"] = sum(1 for v in df["subio_categorizacion"] if v == "SUBIÓ")
        resumen["car_subio"] = sum(1 for v in df["subio_cartera"] if v == "SUBIÓ")
        resumen["car_con_nro"] = sum(1 for v in df["nro_cartera"] if v and v != "—")

    return df, resumen


def enriquecer_con_renipress(df):
    """Enriquece un DataFrame de Ficha 32 (A o B) con la condición oficial RENIPRESS de Categorización, Cartera de Servicios y Telesalud.

    Agrega:
      * 'TELESALUD_ROL': 'CONSULTOR', 'CONSULTANTE', 'AMBOS' o 'NO REGISTRADO'
      * 'NRO_CATEGORIZACION': N° Doc. de Categorización (ej. RDR N° 1543-2024/DRS-PUNO-DESP-DSS)
      * 'FECHA_CATEGORIZACION': Fecha de Categorización (ej. 25/07/2024)
      * 'SUBIO_CATEGORIZACION': 'SUBIÓ' o 'NO SUBIÓ'
      * 'NRO_CARTERA': N° de Resolución de Cartera de Servicios
      * 'FECHA_CARTERA': Fecha de Cartera de Servicios
      * 'SUBIO_CARTERA': 'SUBIÓ' o 'NO SUBIÓ'
      * 'NRO_RESOLUCION': Nro. resolución principal
      * 'FECHA_RESOLUCION': Fecha resolución principal
      * 'VARIACION_RESOLUCION': 'SUBIÓ' o 'NO SUBIÓ'
    """
    if df is None or df.empty or "ESTABLECIMIENTO / IPRESS" not in df.columns:
        return df
    try:
        import ficha32_api as api
        tok = _token_diresa()
        cods = [str(codigo(v)).zfill(8) for v in df["ESTABLECIMIENTO / IPRESS"]]
        mapa_res = api.obtener_resoluciones_renipress(cods, token=tok)

        cartera, _ = obtener_cartera_renipress()
        mapa_roles = dict(zip(cartera["CODIGO_UNICO"], cartera["telemedicina"])) if not cartera.empty and "CODIGO_UNICO" in cartera.columns else {}

        cod_series = pd.Series(cods, index=df.index)
        df["TELESALUD_ROL"] = cod_series.map(mapa_roles).fillna("NO REGISTRADO")

        # Categoría oficial SUSALUD y comparativa con Ficha 32
        df["CATEGORIA_SUSALUD"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("categoria_susalud") or ""))

        def _comparar_cat(row):
            cat_f32 = str(row.get("NIVEL") or row.get("CATEGORIA") or "").strip()
            if not cat_f32 and "ESTABLECIMIENTO / IPRESS" in row:
                cat_f32 = str(row["ESTABLECIMIENTO / IPRESS"]).split(" - ")[0].strip()
            cat_su = str(row.get("CATEGORIA_SUSALUD") or "").strip()
            if not cat_su or cat_su in ("—", "Sin Categoría"):
                return "—"
            if cat_f32 and cat_su:
                if cat_f32 == cat_su:
                    return "✓ IGUAL"
                else:
                    return f"⚠ VARIÓ ({cat_f32} → {cat_su})"
            return "—"

        df["ESTADO_CATEGORIA"] = df.apply(_comparar_cat, axis=1)

        # Asignar campos detallados de Categorización y Cartera
        df["NRO_CATEGORIZACION"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("nro_categorizacion") or "—"))
        df["FECHA_CATEGORIZACION"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("fecha_categorizacion") or "—"))
        df["SUBIO_CATEGORIZACION"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("subio_categorizacion") or "NO SUBIÓ"))

        df["NRO_CARTERA"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("nro_cartera") or "—"))
        df["FECHA_CARTERA"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("fecha_cartera") or "—"))
        df["SUBIO_CARTERA"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("subio_cartera") or "NO SUBIÓ"))

        df["NRO_RESOLUCION"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("nro_resolucion") or "—"))
        df["FECHA_RESOLUCION"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("fecha_resolucion") or "—"))
        df["VARIACION_RESOLUCION"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("variacion_resolucion") or "NO SUBIÓ"))
        df["SUBIO_RESOLUCION"] = cod_series.apply(lambda c: (mapa_res.get(c, {}).get("subio_resolucion") or "NO SUBIÓ"))

    except Exception:
        for c in ["TELESALUD_ROL", "NRO_CATEGORIZACION", "FECHA_CATEGORIZACION", "SUBIO_CATEGORIZACION",
                  "NRO_CARTERA", "FECHA_CARTERA", "SUBIO_CARTERA", "NRO_RESOLUCION", "FECHA_RESOLUCION",
                  "VARIACION_RESOLUCION", "SUBIO_RESOLUCION"]:
            if c not in df.columns:
                df[c] = "—"
    return df

# (RED_SALUD en datos, slug archivo, nombre para mostrar) — 12 Redes de Salud oficiales de DIRESA Puno
REDES = [
    ("PUNO", "Puno", "PUNO"),
    ("AZANGARO", "Azangaro", "AZÁNGARO"),
    ("EL COLLAO", "El_Collao", "EL COLLAO"),
    ("MELGAR", "Melgar", "MELGAR"),
    ("HUANCANE", "Huancane", "HUANCANÉ"),
    ("CARABAYA", "Carabaya", "CARABAYA"),
    ("SAN ROMAN", "San_Roman", "SAN ROMÁN"),
    ("LAMPA", "Lampa", "LAMPA"),
    ("YUNGUYO", "Yunguyo", "YUNGUYO"),
    ("SANDIA", "Sandia", "SANDIA"),
    ("SAN ANTONIO DE PUTINA", "San_Antonio_de_Putina", "SAN ANTONIO DE PUTINA"),
    ("CHUCUITO", "Chucuito", "CHUCUITO"),
]

# Hospital consultor (Reporte B) por red y nombres propios de hospitales:
# viven en nombres_hospitales.py, que el usuario puede ampliar sin tocar el código.
try:
    import nombres_hospitales as _NH
    _, HOSPITAL_REF = _NH.cargar()          # HOSPITAL_POR_RED, indexada por red de salud
    NOMBRES_HOSPITAL = dict(_NH.NOMBRES_HOSPITAL)
except Exception:  # noqa: BLE001 - si falta el archivo, se trabaja con las tablas mínimas
    HOSPITAL_REF = {
        "PUNO": "Hospital Regional Manuel Núñez Butrón",
        "SAN ROMAN": "Hospital Carlos Monje Medrano de Juliaca",
        "CARABAYA": "Hospital San Martín de Porres de Macusani",
        "EL COLLAO": "Hospital de Ilave",
        "MELGAR": "Hospital San Juan de Dios de Ayaviri",
    }
    NOMBRES_HOSPITAL = {"00003251": "HOSPITAL REGIONAL MANUEL NÚÑEZ BUTRÓN"}

# Parámetros del plan Nov-Dic
PISO_CAP_MES = 12       # capacidad mínima realista (atenciones/mes, 3 servicios)
PCT_CAP_META = 0.05     # o 5 % de la meta anual total, el mayor
RATIO_CERCA = 1 / 3     # requerido/capacidad <= 1/3 -> "CERCA"; <= 1 -> "MODERADO"; > 1 -> "BRECHA ALTA"
FRAC_ARRANQUE = 1 / 3   # piso de arranque del plan base para BRECHA ALTA (fracción de la capacidad mensual)

CATEGORIAS = ["CUMPLE", "CERCA", "MODERADO", "BRECHA ALTA"]

# Hospitales de referencia identificados en el Reporte A (II-1) — solo para redactar estrategias
HOSPITAL_ESTRATEGIA = {
    "AZANGARO": "Hospital Carlos Cornejo Rosello",
    "LAMPA": "Hospital Antonio Barrionuevo de Lampa",
    "HUANCANE": "Hospital Lucio Aldazábal Pauca",
}


def leer(path):
    df = pd.read_html(path, encoding="utf-8")[0]
    df.columns = [str(c).strip() for c in df.columns]
    for c in NUM:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["PROVINCIA"] = df["PROVINCIA"].astype(str).str.strip()
    df["ESTABLECIMIENTO / IPRESS"] = df["ESTABLECIMIENTO / IPRESS"].astype(str).str.strip()
    return df


def short_name(raw):
    parts = str(raw).split(" - ")
    return parts[-1].strip() if len(parts) > 1 else str(raw).strip()


def codigo(raw):
    parts = str(raw).split(" - ")
    return parts[1].strip() if len(parts) > 2 else ""


def clasificar(s):
    sem = CM.get_metas_config().get("semaforo", {})
    v_min = float(sem.get("verde_min", 60.0))
    a_min = float(sem.get("amarillo_min", 30.0))
    if s >= v_min:
        return "VERDE"
    if s >= a_min:
        return "AMARILLO"
    return "ROJO"


def puntaje_a(df, suf=""):
    """Puntaje de Logro oficial (0-100) para el primer nivel. suf permite columnas con sufijo."""
    p = 0.0
    for s in SERV:
        r = df[f"{s} REALIZADAS{suf}"]
        m = df[f"{s} META{suf}"]
        p = p + PESO[s] * np.minimum(r, m) / m
    return p


def pts_hasta(df, realizadas):
    """Puntaje a partir de un dict servicio->serie de realizadas (con tope por meta)."""
    p = 0.0
    for s in SERV:
        p = p + PESO[s] * np.minimum(realizadas[s], df[f"{s} META"]) / df[f"{s} META"]
    return p


def atenciones_min_para_puntos(row_vals, objetivo):
    """Mínimo de atenciones adicionales para alcanzar `objetivo` puntos (greedy por puntos/atención).

    row_vals: lista de (peso, meta, realizado_proyectado). Devuelve (n_atenciones, dict servicio->atenciones).
    """
    pts_actual = sum(w * min(r, m) / m for _, w, m, r in row_vals)
    falta = objetivo - pts_actual
    if falta <= 1e-9:
        return 0, {s: 0 for s, *_ in row_vals}
    # ordenar por puntos por atención (peso/meta) descendente
    orden = sorted(row_vals, key=lambda t: t[1] / t[2], reverse=True)
    total = 0
    detalle = {s: 0 for s, *_ in row_vals}
    for s, w, m, r in orden:
        if falta <= 1e-9:
            break
        cupo = max(0, m - r)
        if cupo == 0:
            continue
        pts_por_at = w / m
        necesarias = int(np.ceil(falta / pts_por_at - 1e-9))
        usar = min(cupo, necesarias)
        detalle[s] = usar
        total += usar
        falta -= usar * pts_por_at
    if falta > 1e-6:
        return None, detalle   # inalcanzable
    return total, detalle


MAPEO_REDES_PUNO = {
    "PUNO": "PUNO",
    "AZANGARO": "AZANGARO",
    "EL COLLAO": "EL COLLAO",
    "COLLAO": "EL COLLAO",
    "MELGAR": "MELGAR",
    "HUANCANE": "HUANCANE",
    "MOHO": "HUANCANE",
    "CARABAYA": "CARABAYA",
    "MACUSANI": "CARABAYA",
    "SAN ROMAN": "SAN ROMAN",
    "LAMPA": "LAMPA",
    "YUNGUYO": "YUNGUYO",
    "SANDIA": "SANDIA",
    "SAN ANTONIO DE PUTINA": "SAN ANTONIO DE PUTINA",
    "CHUCUITO": "CHUCUITO",
}


def resolver_red_salud(df, diresa=""):
    """Resuelve la Red de Salud oficial para cada fila de un DataFrame de IPRESS.

    En DIRESA Puno, las 12 Redes de Salud oficiales corresponden a las provincias
    (Moho pertenece administrativamente a la Red Huancané, Carabaya a la Red Carabaya/Macusani,
    y El Collao a la Red El Collao/Collao). La API del tablero entrega el campo 'red'
    con inconsistencias en Puno (p. ej. asigna erróneamente establecimientos de Sandia
    a Melgar o Putina). Esta función garantiza la asignación correcta tanto para Puno
    como para las demás DIRESAS del país.
    """
    c_prov = "PROVINCIA" if "PROVINCIA" in df.columns else ("provincia" if "provincia" in df.columns else None)
    c_red = "RED" if "RED" in df.columns else ("red" if "red" in df.columns else None)
    c_dir = "DIRESA" if "DIRESA" in df.columns else ("diresa" if "diresa" in df.columns else None)

    es_puno = "PUNO" in str(diresa).upper()
    if not es_puno and c_dir and (df[c_dir].astype(str).str.upper().str.contains("PUNO")).any():
        es_puno = True
    if not es_puno and c_prov:
        provs = set(df[c_prov].dropna().astype(str).str.upper().str.strip())
        if provs and provs.issubset(set(MAPEO_REDES_PUNO.keys())):
            es_puno = True

    if es_puno and c_prov:
        prov = df[c_prov].astype(str).str.strip().str.upper()
        return prov.map(MAPEO_REDES_PUNO).fillna(prov)

    if c_red and df[c_red].astype(str).str.strip().ne("").any():
        ra = df[c_red].astype(str).str.strip().str.upper()
        fb = df[c_prov].astype(str).str.strip().str.upper() if c_prov else ra
        return ra.mask(ra.isin(["NO PERTENECE A NINGUNA RED", "", "NONE", "NAN"])).fillna(fb)

    if c_prov:
        return df[c_prov].astype(str).str.strip().str.upper()
    return pd.Series("SIN_RED", index=df.index)


ALIAS_REDES = {
    "CARABAYA": "MACUSANI", "MACUSANI": "CARABAYA",
    "EL COLLAO": "COLLAO", "COLLAO": "EL COLLAO",
    "MOHO": "HUANCANE",
}


def hospital_consultor(prov, A=None, B=None):
    """Hospital consultor de referencia de la red.

    1) Tabla fija (DIRESA Puno u otras configuradas).
    2) Si hay B (Reporte B), el hospital asociado a esa red con mayor denominador.
    3) Si no está en B, el de mayor categoría (II o III) de esa red en A.
    4) Texto de referencia seguro.
    """
    p_norm = str(prov).strip().upper()
    if p_norm in HOSPITAL_REF:
        return HOSPITAL_REF[p_norm]
    alias = ALIAS_REDES.get(p_norm)
    if alias and alias in HOSPITAL_REF:
        return HOSPITAL_REF[alias]
    if p_norm in HOSPITAL_ESTRATEGIA:
        return HOSPITAL_ESTRATEGIA[p_norm]

    # Si se pasa B o hay hospitales del Reporte B para esta red
    if B is not None and len(B):
        sub_b = B[(B["RED_SALUD"].astype(str).str.upper() == p_norm) |
                  (B["RED_SALUD"].astype(str).str.upper().apply(lambda x: bool(x and (x in p_norm or p_norm in x))))]
        if len(sub_b):
            return str(sub_b.sort_values("DENOMINADOR", ascending=False)["NOMBRE"].iloc[0]).title()

    if A is not None and len(A):
        sub = A[A["RED_SALUD"] == prov]
        hosp = sub[sub["NIVEL"].astype(str).str.startswith(("II", "III"))]
        if len(hosp):
            return str(hosp.sort_values("DENOMINADOR", ascending=False)["NOMBRE"].iloc[0]).title()
    return "el hospital de referencia de la red"


def limpiar_slug_hoja(nombre):
    """Limpia el nombre para que sea un slug válido y un nombre de hoja seguro para Excel.

    Excel prohíbe los caracteres: [ ] : * ? / \\
    """
    s = re.sub(r"[\[\]:*?/\\]", "_", str(nombre))
    slug = (s.lower().replace(" ", "_").replace("á", "a").replace("é", "e")
            .replace("í", "i").replace("ó", "o").replace("ú", "u").replace("ñ", "n"))
    return re.sub(r"_+", "_", slug).strip("_")


def redes(A):
    """(RED_SALUD, slug de archivo, nombre para mostrar) derivados de los datos.

    Permite usar el modelo con cualquier DIRESA sin editar la lista fija de Puno.
    Si todos los datos pertenecen a la lista fija se respeta su orden (que es el de
    las hojas del libro); si no, se deriva de los datos presentes.
    """
    pres = set(A["RED_SALUD"].dropna().astype(str)) if A is not None and len(A) else set()
    if REDES and pres and pres <= {r[0] for r in REDES}:
        return [r for r in REDES if r[0] in pres]
    out = []
    for r in sorted(pres):
        slug = limpiar_slug_hoja(r)
        out.append((r, slug[:1].upper() + slug[1:], r))
    return out
    return out


CTX = {"scope_vista": "", "nombre_diresa": DIRESA}


def texto_accion(r, hosp):
    """Meta de producción sugerida para Nov-Dic (por mes) y acción concreta por servicio."""
    cat = r["CATEGORIA"]
    if cat == "CUMPLE":
        return "Meta anual asegurada. Mantener el ritmo, verificar digitación y servir como referente."
    partes = []

    def tramo(nd, etiqueta, extra=""):
        mes = int(np.ceil(nd / MESES_PLAN))
        cuando = f"≈{mes}/mes" if nd > 1 else "en el periodo"
        return f"{nd} {etiqueta}{extra} en Nov–Dic ({cuando})"
    if r["TC OBJ_ND"]:
        partes.append(tramo(r["TC OBJ_ND"], "teleconsulta(s)", f" [+{PESO['TC'] / r['TC META']:.1f} pts c/u]"))
    if r["TI OBJ_ND"]:
        partes.append(tramo(r["TI OBJ_ND"], "teleinterconsulta(s)", f" con {hosp} [+{PESO['TI'] / r['TI META']:.1f} pts c/u]"))
    if r["TM OBJ_ND"]:
        partes.append(tramo(r["TM OBJ_ND"], "telemonitoreo(s)", f" del Anexo 1 [+{PESO['TM'] / r['TM META']:.1f} pts c/u]"))
    return "; ".join(partes) if partes else "Sostener la producción actual."


def agregar_acciones(A, B=None):
    for s in SERV:
        A[f"{s} OBJ_ND"] = A[f"{s} BASE_ND"]
        A[f"{s} OBJ_MES"] = np.ceil(A[f"{s} OBJ_ND"] / MESES_PLAN).astype(int)
    A["OBJ_MES_TOTAL"] = sum(A[f"{s} OBJ_MES"] for s in SERV)
    A["ACCION"] = [texto_accion(r, "hospitales de mayor complejidad" if r["NIVEL"] == "II-1"
                                else hospital_consultor(r["RED_SALUD"], A, B)) for _, r in A.iterrows()]
    return A


def construir():
    A, B = _construir()
    A = agregar_acciones(A, B)
    return aplicar_filtro(A, B)


# Alcance del informe: se puede limitar a una o varias redes o a una lista de
# establecimientos desde la GUI sin tocar el resto del pipeline. Ver aplicar_filtro().
FILTRO = {"red": None, "codigos": []}


def aplicar_filtro(A, B):
    """Restringe A y B a la red o a los códigos activos (FILTRO). Vacío = toda la DIRESA.

    * FILTRO["red"]: la red se acepta completa o parcial ("AREQUIPA" encuentra
      "AREQUIPA CAYLLOMA").
    * FILTRO["codigos"]: lista de códigos IPRESS (8 dígitos); vacía = sin filtro.
    """
    red, codigos = FILTRO.get("red"), FILTRO.get("codigos") or []
    if red:
        pat = str(red).upper()
        A = A[A["RED_SALUD"].astype(str).str.upper().str.contains(pat, regex=False, na=False)]
        B = B[B["RED_SALUD"].astype(str).str.upper().str.contains(pat, regex=False, na=False)]
    if codigos:
        cods = {str(c).strip().zfill(8) for c in codigos}
        A = A[A["CODIGO"].astype(str).str.zfill(8).isin(cods)]
        B = B[B["CODIGO"].astype(str).str.zfill(8).isin(cods)]
    return A.reset_index(drop=True), B.reset_index(drop=True)


def _construir():
    a8, b8, a7, b7, CTX = cargar()
    for d in (a8, a7):
        d["NIVEL"] = d["ESTABLECIMIENTO / IPRESS"].str.split().str[0]
    k = "ESTABLECIMIENTO / IPRESS"
    # Integridad: el padrón puede variar entre meses (altas/bajas en el HIS); no se aborta, se avisa.
    if len(a7):
        faltan = set(a8[k]) - set(a7[k])
        if faltan:
            print(f"[aviso] {len(faltan)} IPRESS sin dato en el mes anterior; se asume 0.")

    A = a8.copy()
    # La API entrega la red de salud oficial; los .xls solo traen provincia
    # (en Puno, la provincia de Moho pertenece administrativamente a la Red Huancané).
    A["RED_SALUD"] = resolver_red_salud(A, diresa=DIRESA)
    prev = a7.set_index(k) if len(a7) else pd.DataFrame(columns=a7.columns)
    for s in SERV:
        A[f"{s} ANT"] = A[k].map(prev[f"{s} REALIZADAS"] if len(prev) else pd.Series(dtype=float)) \
            .astype(float).fillna(0)
        A[f"{s} MES{MES_CORTE}"] = (A[f"{s} REALIZADAS"] - A[f"{s} ANT"]).clip(lower=0)
        if ((A[f"{s} REALIZADAS"] - A[f"{s} ANT"]) < 0).any():
            print(f"[aviso] {s}: el acumulado bajó respecto del mes anterior "
                  f"(re-digitalización del HIS). Se ajusta la producción mensual a 0.")
    A["PUNTAJE_ANT"] = pts_hasta(A.assign(**{f"{s} META": A[f"{s} META"] for s in SERV}),
                                 {s: A[f"{s} ANT"] for s in SERV}).round(1)
    A["PUNTAJE"] = puntaje_a(A).round(1)
    A["SEMAFORO"] = A["PUNTAJE"].apply(clasificar)
    A["SEMAFORO_ANT"] = A["PUNTAJE_ANT"].apply(clasificar)
    A["DELTA_PTS"] = (A["PUNTAJE"] - A["PUNTAJE_ANT"]).round(1)

    co_cfg = CM.get_metas_config().get("cumplimiento_oficial", {})
    b_restar = float(co_cfg.get("base_restar", 10.0))
    divisor = float(co_cfg.get("divisor", 90.0))
    A["CUMPL_OFICIAL"] = np.where(A["PUNTAJE"] > b_restar, (A["PUNTAJE"] - b_restar) / divisor * 100, 0.0).round(1)
    A["NOMBRE"] = A[k].map(short_name)
    A["CODIGO"] = A[k].map(codigo)

    # --- Proyección por servicio ---
    # Ritmo conservador: promedio mensual ene→ago (R8/8). Ritmo agosto: producción del mes 8.
    pts_cons, pts_ago = 0.0, 0.0
    proj_cons = {}
    for s in SERV:
        M, R = A[f"{s} META"], A[f"{s} REALIZADAS"]
        A[f"{s} RITMO_PROM"] = R / MES_CORTE
        A[f"{s} RITMO_AGO"] = A[f"{s} MES{MES_CORTE}"]
        # fin de octubre (2 meses proyectados) – escenario conservador
        A[f"{s} PROY_OCT"] = (R + MESES_PROY_PREVIOS * A[f"{s} RITMO_PROM"]).round(0)
        A[f"{s} PROY_OCT_AGO"] = (R + MESES_PROY_PREVIOS * A[f"{s} RITMO_AGO"]).round(0)
        # brecha que quedaría para nov-dic
        A[f"{s} BRECHA_ND"] = np.maximum(0, M - A[f"{s} PROY_OCT"]).astype(int)
        A[f"{s} REQ_MES"] = np.ceil(A[f"{s} BRECHA_ND"] / MESES_PLAN).astype(int)
        # ritmo vigente redondeado (atenciones/mes) para comparar
        A[f"{s} PROY_DIC_TEND"] = np.minimum(M, R + (MESES_AÑO - MES_CORTE) * A[f"{s} RITMO_PROM"]).round(0)
        A[f"{s} PROY_DIC_AGO"] = np.minimum(M, R + (MESES_AÑO - MES_CORTE) * A[f"{s} RITMO_AGO"]).round(0)
        proj_cons[s] = np.minimum(M, A[f"{s} PROY_OCT"])

    A["PUNTAJE_PROY_OCT"] = pts_hasta(A, proj_cons).round(1)
    A["PUNTAJE_DIC_TEND"] = pts_hasta(A, {s: A[f"{s} PROY_DIC_TEND"] for s in SERV}).round(1)
    A["PUNTAJE_DIC_AGO"] = pts_hasta(A, {s: A[f"{s} PROY_DIC_AGO"] for s in SERV}).round(1)
    A["BRECHA_TOTAL_ND"] = sum(A[f"{s} BRECHA_ND"] for s in SERV)
    A["REQ_MES_TOTAL"] = np.ceil(A["BRECHA_TOTAL_ND"] / MESES_PLAN).astype(int)
    A["EXTRA_TOTAL_RESTANTE"] = sum(np.maximum(0, A[f"{s} META"] - A[f"{s} REALIZADAS"]) for s in SERV).astype(int)

    # Capacidad realista de producción mensual (atenciones/mes, 3 servicios) durante Nov-Dic:
    # el mayor entre un piso de 12, el 5 % de la meta anual total y el rendimiento mensual ya demostrado.
    meta_total = sum(A[f"{s} META"] for s in SERV)
    ritmo_total = sum(A[f"{s} RITMO_PROM"] for s in SERV)
    A["CAP_MES"] = np.maximum.reduce([np.full(len(A), PISO_CAP_MES),
                                      np.ceil(PCT_CAP_META * meta_total),
                                      np.ceil(ritmo_total)]).astype(int)
    A["RATIO_ESFUERZO"] = (A["REQ_MES_TOTAL"] / A["CAP_MES"]).round(2)

    def cat(r):
        if r["BRECHA_TOTAL_ND"] == 0:
            return "CUMPLE"
        if r["RATIO_ESFUERZO"] <= RATIO_CERCA:
            return "CERCA"
        if r["RATIO_ESFUERZO"] <= 1.0:
            return "MODERADO"
        return "BRECHA ALTA"
    A["CATEGORIA"] = A.apply(cat, axis=1)

    # Plan Nov-Dic: cada IPRESS destina hasta CAP_MES x 2 atenciones, priorizando el servicio con más puntos/atención
    plan = {s: [] for s in SERV}
    for _, r in A.iterrows():
        presupuesto = int(r["CAP_MES"]) * MESES_PLAN
        orden = sorted(SERV, key=lambda s: PESO[s] / r[f"{s} META"], reverse=True)
        for s in orden:
            usar = int(min(r[f"{s} BRECHA_ND"], presupuesto))
            plan[s].append(usar)
            presupuesto -= usar
    for s in SERV:
        A[f"{s} PLAN_ND"] = plan[s]
        A[f"{s} PLAN_MES"] = np.ceil(A[f"{s} PLAN_ND"] / MESES_PLAN).astype(int)
    A["PLAN_ND_TOTAL"] = sum(A[f"{s} PLAN_ND"] for s in SERV)
    A["PUNTAJE_PLAN"] = pts_hasta(A, {s: np.minimum(A[f"{s} META"], A[f"{s} PROY_OCT"] + A[f"{s} PLAN_ND"])
                                      for s in SERV}).round(1)
    A["SEMAFORO_PLAN"] = A["PUNTAJE_PLAN"].apply(clasificar)
    A["SEMAFORO_OCT"] = A["PUNTAJE_PROY_OCT"].apply(clasificar)

    # Plan base (realista): CERCA/MODERADO/CUMPLE completan su meta; BRECHA ALTA duplica su ritmo en Nov-Dic,
    # con un piso de "arranque" equivalente a 1/3 de su capacidad mensual (prioriza el servicio de mayor puntos/atención)
    extra_base = {s: [] for s in SERV}
    for _, r in A.iterrows():
        if r["CATEGORIA"] != "BRECHA ALTA":
            for s in SERV:
                extra_base[s].append(int(r[f"{s} BRECHA_ND"]))
            continue
        ext = {s: int(min(r[f"{s} BRECHA_ND"], round(MESES_PLAN * 2 * r[f"{s} RITMO_PROM"]))) for s in SERV}
        faltante = int(np.ceil(r["CAP_MES"] * FRAC_ARRANQUE)) * MESES_PLAN - sum(ext.values())
        for s in sorted(SERV, key=lambda s: PESO[s] / r[f"{s} META"], reverse=True):
            if faltante <= 0:
                break
            cupo = int(r[f"{s} BRECHA_ND"]) - ext[s]
            usar = max(0, min(cupo, faltante))
            ext[s] += usar
            faltante -= usar
        for s in SERV:
            extra_base[s].append(ext[s])
    for s in SERV:
        A[f"{s} BASE_ND"] = extra_base[s]
        A[f"{s} BASE_DIC"] = np.round(np.minimum(A[f"{s} META"], A[f"{s} PROY_OCT"] + A[f"{s} BASE_ND"])).astype(int)
    A["PUNTAJE_PLAN_BASE"] = pts_hasta(A, {s: A[f"{s} BASE_DIC"] for s in SERV}).round(1)
    A["SEMAFORO_PLAN_BASE"] = A["PUNTAJE_PLAN_BASE"].apply(clasificar)

    def ruta(r):
        if r["CATEGORIA"] != "BRECHA ALTA":
            return ""
        if r["PUNTAJE_PLAN"] >= 60:
            return "Ruta a Verde"
        if r["PUNTAJE_PLAN"] >= 30:
            return "Ruta a Amarillo"
        return "Arranque"
    A["RUTA"] = A.apply(ruta, axis=1)

    # Atenciones mínimas (sobre proyección fin de octubre) para escalar a Amarillo(30) / Verde(60)
    for objetivo, col in ((30, "AT_AMARILLO"), (60, "AT_VERDE")):
        vals = []
        servs = []
        for _, r in A.iterrows():
            lst = [(s, PESO[s], r[f"{s} META"], r[f"{s} PROY_OCT"]) for s in SERV]
            n, det = atenciones_min_para_puntos(lst, objetivo)
            vals.append(n if n is not None else np.nan)
            servs.append(det)
        A[col] = vals
        A[col + "_DET"] = servs

    # Servicio de mayor rendimiento (pts por atención) con cupo – para la "jugada rápida"
    def mejor_servicio(r):
        cand = []
        for s in SERV:
            if r[f"{s} META"] - r[f"{s} PROY_OCT"] > 0:
                cand.append((PESO[s] / r[f"{s} META"], s))
        return max(cand)[1] if cand else ""
    A["SERV_PALANCA"] = A.apply(mejor_servicio, axis=1)
    A["PTS_POR_AT_PALANCA"] = [
        round(PESO[s] / r[f"{s} META"], 2) if s else 0.0 for s, (_, r) in zip(A["SERV_PALANCA"], A.iterrows())]

    # --- Hospitales (Reporte B) ---
    B = b8.copy()
    B["RED_SALUD"] = resolver_red_salud(B, diresa=DIRESA)

    # Mapear cada hospital de B a la Red de Salud correspondiente de A
    if len(B) and len(A):
        redes_a = list(A["RED_SALUD"].dropna().unique())

        def asociar_red(r):
            rs = str(r.get("RED_SALUD", "")).strip().upper()
            if rs in redes_a:
                return rs
            alias = ALIAS_REDES.get(rs)
            if alias and alias in redes_a:
                return alias
            for ra in redes_a:
                if rs and (rs in ra or ra in rs):
                    return ra
            prov = str(r.get("PROVINCIA", "")).strip().upper()
            if prov in redes_a:
                return prov
            for ra in redes_a:
                if prov and (prov in ra or ra in prov):
                    return ra
            return rs
        B["RED_SALUD"] = B.apply(asociar_red, axis=1)

    prevb = b7.set_index(k) if len(b7) else pd.DataFrame(columns=b7.columns)
    B["TI ANT"] = B[k].map(prevb["TI REALIZADAS"] if len(prevb) else pd.Series(dtype=float)) \
        .astype(float).fillna(0)
    B["DEN ANT"] = B[k].map(prevb["DENOMINADOR"] if len(prevb) else pd.Series(dtype=float)) \
        .astype(float).fillna(0)
    B[f"TI MES{MES_CORTE}"] = (B["TI REALIZADAS"] - B["TI ANT"]).clip(lower=0)
    B["PUNTAJE"] = (100 * np.minimum(B["TI REALIZADAS"], B["TI META"]) / B["TI META"]).round(1)
    meta_ant = B[k].map(prevb["TI META"] if len(prevb) else pd.Series(dtype=float)) \
        .astype(float).fillna(B["TI META"]).values.astype(float)
    B["PUNTAJE_ANT"] = (100 * np.minimum(B["TI ANT"], meta_ant) / meta_ant).round(1)
    B["SEMAFORO"] = B["PUNTAJE"].apply(clasificar)
    B["SEMAFORO_ANT"] = B["PUNTAJE_ANT"].apply(clasificar)
    B["NIVEL"] = B[k].str.split().str[0]
    # Nombres propios de hospitales conocidos; se pueden ampliar con configurar(nombres_hospital=...)
    nombres_b = dict(NOMBRES_HOSPITAL)
    B["CODIGO"] = B[k].map(codigo)
    B["NOMBRE"] = [nombres_b.get(c, short_name(n)) for c, n in zip(B["CODIGO"], B[k])]
    B["DEN_ACT_MES"] = B["DENOMINADOR"] - B["DEN ANT"]            # atenciones HIS del mes de corte
    B["DEN_PROY_DIC"] = (B["DENOMINADOR"] / MES_CORTE * MESES_AÑO).round(0)   # lineal sobre acumulado
    B["META_PROY_DIC"] = np.maximum(100, np.round(40 / 10000 * B["DEN_PROY_DIC"])).astype(int)
    B["RITMO_PROM"] = B["TI REALIZADAS"] / MES_CORTE
    B["RITMO_AGO"] = B[f"TI MES{MES_CORTE}"]
    B["TI_PROY_OCT"] = (B["TI REALIZADAS"] + MESES_PROY_PREVIOS * B["RITMO_PROM"]).round(0)
    B["META_PROY_OCT"] = np.maximum(100, np.round(40 / 10000 * B["DENOMINADOR"] / MES_CORTE * MES_PROY)).astype(int)
    B["TI_PROY_DIC_TEND"] = (B["TI REALIZADAS"] + (MESES_AÑO - MES_CORTE) * B["RITMO_PROM"]).round(0)
    B["BRECHA_DIC"] = np.maximum(0, B["META_PROY_DIC"] - B["TI_PROY_OCT"]).astype(int)
    B["REQ_MES"] = np.ceil(B["BRECHA_DIC"] / MESES_PLAN).astype(int)
    B["PUNTAJE_DIC_TEND"] = (100 * np.minimum(B["TI_PROY_DIC_TEND"], B["META_PROY_DIC"]) / B["META_PROY_DIC"]).round(1)
    # Esfuerzo exigente sostenido en Nov-Dic: el mayor entre el ritmo de agosto y 2 veces el promedio ene-ago
    B["CAP_MES"] = np.ceil(np.maximum(B["RITMO_AGO"], 2 * B["RITMO_PROM"])).astype(int)
    B["TI_PLAN_DIC"] = np.minimum(B["META_PROY_DIC"], B["TI_PROY_OCT"] + B["CAP_MES"] * MESES_PLAN).round(0)
    B["PUNTAJE_PLAN"] = (100 * B["TI_PLAN_DIC"] / B["META_PROY_DIC"]).round(1)
    A = enriquecer_con_renipress(A)
    B = enriquecer_con_renipress(B)
    return A, B


if __name__ == "__main__":
    pd.set_option("display.width", 250)
    pd.set_option("display.max_columns", 40)
    A, B = construir()
    print(A["SEMAFORO"].value_counts(), A["SEMAFORO_ANT"].value_counts())
    print("Puntaje medio ant/act:", A["PUNTAJE_ANT"].mean().round(2), A["PUNTAJE"].mean().round(2),
          "| proy oct:", A["PUNTAJE_PROY_OCT"].mean().round(2), "| dic tend:", A["PUNTAJE_DIC_TEND"].mean().round(2),
          "| dic ago:", A["PUNTAJE_DIC_AGO"].mean().round(2))
    print(A["CATEGORIA"].value_counts())
    print(A.groupby("CATEGORIA")[["BRECHA_TOTAL_ND", "OBJ_MES_TOTAL", "PUNTAJE"]].agg(["mean", "sum"]))
    print(A[["AT_AMARILLO", "AT_VERDE"]].describe())
    print(B[["NOMBRE", "TI REALIZADAS", "TI META", "PUNTAJE", "PUNTAJE_ANT", "META_PROY_DIC", "TI_PROY_OCT",
             "BRECHA_DIC", "REQ_MES", "PUNTAJE_DIC_TEND"]])
