# -*- coding: utf-8 -*-
"""Conector a la API del Dashboard Gerencial MINSA | Ficha N.° 32 (Telemedicina).

Reemplaza el paso manual de "Exportar a Excel" del tablero
(https://ficha32-ditel.web.app/) por una descarga programática.

ENDPOINT
    GET https://api-ficha32-ditel-303392539829.europe-west1.run.app?mes={1..12}
    Header: Authorization: Bearer <código de usuario o IPRESS>

ALCANCE (control de acceso por RLS, sin contraseña):
    * "DIRESA_<REGIÓN>"      -> REGIONAL  (toda la red de la dirección regional)
    * "<id_ipress>" numérico -> LOCAL     (un solo establecimiento)
    El token es sensible a mayúsculas. No hay vista nacional por este camino.

RESPUESTA (JSON):
    {
      "success": true,
      "user_context": {"rol": ..., "scope_vista": "REGIONAL", "region_asignada": "PUNO"},
      "kpis_consolidados": {"atenciones_validas": ..., "cump_total": ..., "ind_total": ...},
      "data_establecimientos": [
        {"id_ipress": 2935, "tipo_indicador": "A", "mes": 8, "diresa": "PUNO", "red": "AZANGARO",
         "provincia": "AZANGARO", "nombre_ipress": "I-4   - 00002935 - ARAPA", "denominador": 5426,
         "numerador": .., "indicador": .., "ti": 4, "tc": 37, "tm": 29, "estado": 1}, ...]
    }

DATOS VERIFICADOS frente al .xls exportado a mano (DIRESA Puno, mes 8):
    * 375/375 establecimientos presentes y con denominador, TI, TC y TM idénticos.
    * Las METAS NO vienen en la API; se reproducen con las fórmulas del propio dashboard:
        tipo A:  TI = máx(24, 120/10000 x den)   TC = máx(1, 10/10000 x den)   TM = máx(24, 72/10000 x den)
        tipo B:  TI = máx(100, 40/10000 x den)   TC y TM sin meta ("--")
      Se replica el redondeo "hacia arriba en mitad" de JavaScript (Math.round), no el de numpy.
    * Las atenciones son ACUMULADAS ene->mes (nunca disminuyen entre meses).
    * Solo hay datos de los meses en curso: los meses futuros responden 404.

CACHÉ: las respuestas se guardan en ./cache_api/ para no repetir llamadas y
trabajar sin internet una vez descargadas.

NUEVA API RENIPRESS (SUSALUD)
-------------------------------
Endpoint: GET https://app20.susalud.gob.pe:8086/registro-renipress-webapp/listadoEstablecimientosRegistrados.htm?action=mostrarBuscar
Sin autenticación (acceso público). Devuelve HTML con listado de establecimientos.
Los datos se cachean en ./cache_api/renipress_ para trabajo sin internet.
"""

import json
import math
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request

try:
    import requests
except ImportError:
    requests = None

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

import pandas as pd
import config_manager as CM

API = CM.get_api_url()
CACHE = os.path.join(CM.get_app_dir(), "cache_api")
MESES = list(range(1, 13))
MESES_TXT = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
             7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
SERV = ["TI", "TC", "TM"]
COLS_REPORTE = ["DIRESA", "RED", "PROVINCIA", "ESTABLECIMIENTO / IPRESS", "MES", "LOGRO", "DENOMINADOR",
                "TI REALIZADAS", "TI META", "TC REALIZADAS", "TC META", "TM REALIZADAS", "TM META"]

# Códigos de usuario del tablero (datalist del login), en el orden del MINSA.
DIRESAS = [
    ("DIRESA_AMAZONAS", "AMAZONAS"), ("DIRESA_ANCASH", "ÁNCASH"), ("DIRESA_APURIMAC", "APURÍMAC"),
    ("DIRESA_AREQUIPA", "AREQUIPA"), ("DIRESA_AYACUCHO", "AYACUCHO"), ("DIRESA_CAJAMARCA", "CAJAMARCA"),
    ("DIRESA_CALLAO", "CALLAO"), ("DIRESA_CUSCO", "CUSCO"), ("DIRESA_HUANCAVELICA", "HUANCAVELICA"),
    ("DIRESA_HUANUCO", "HUÁNUCO"), ("DIRESA_ICA", "ICA"), ("DIRESA_JUNIN", "JUNÍN"),
    ("DIRESA_LA_LIBERTAD", "LA LIBERTAD"), ("DIRESA_LAMBAYEQUE", "LAMBAYEQUE"),
    ("DIRESA_LIMA_DIRIS_CENTRO", "LIMA - DIRIS CENTRO"), ("DIRESA_LIMA_DIRIS_ESTE", "LIMA - DIRIS ESTE"),
    ("DIRESA_LIMA_DIRIS_NORTE", "LIMA - DIRIS NORTE"), ("DIRESA_LIMA_DIRIS_SUR", "LIMA - DIRIS SUR"),
    ("DIRESA_LIMA_REGION", "LIMA REGIÓN"), ("DIRESA_LORETO", "LORETO"), ("DIRESA_MADRE_DE_DIOS", "MADRE DE DIOS"),
    ("DIRESA_MOQUEGUA", "MOQUEGUA"), ("DIRESA_PASCO", "PASCO"), ("DIRESA_PIURA", "PIURA"),
    ("DIRESA_PUNO", "PUNO"), ("DIRESA_SAN_MARTIN", "SAN MARTÍN"), ("DIRESA_TACNA", "TACNA"),
    ("DIRESA_TUMBES", "TUMBES"), ("DIRESA_UCAYALI", "UCAYALI"),
]
NOMBRE_POR_TOKEN = dict(DIRESAS)
TOKEN_POR_NOMBRE = {v: k for k, v in DIRESAS}


class ErrorFicha32(Exception):
    """Error controlado del conector (credencial inválida, sin conexión, etc.)."""


def r_media(x):
    """Redondeo "hacia arriba en la mitad" como Math.round de JavaScript.

    numpy usa redondeo bancario (46.5 -> 46), el tablero usa 46.5 -> 47. Las metas
    del .xls exportado a mano coinciden con este comportamiento.
    """
    return int(math.floor(x + 0.5))


def metas(tipo, denominador):
    """Metas anuales por servicio según la Ficha Técnica N.° 32 (reglas configurables)."""
    den = float(denominador)
    cfg = CM.get_metas_config()
    if tipo == "B":
        mb = cfg.get("metas_tipo_b", {}).get("TI", {"piso": 100, "factor_denominador": 0.0040})
        piso = mb.get("piso", 100)
        factor = mb.get("factor_denominador", 0.0040)
        return {"TI": max(piso, r_media(factor * den)), "TC": "--", "TM": "--"}

    ma = cfg.get("metas_tipo_a", {})
    ti_cfg = ma.get("TI", {"piso": 24, "factor_denominador": 0.0120})
    tc_cfg = ma.get("TC", {"piso": 1, "factor_denominador": 0.0010})
    tm_cfg = ma.get("TM", {"piso": 24, "factor_denominador": 0.0072})
    return {
        "TI": max(ti_cfg.get("piso", 24), r_media(ti_cfg.get("factor_denominador", 0.0120) * den)),
        "TC": max(tc_cfg.get("piso", 1), r_media(tc_cfg.get("factor_denominador", 0.0010) * den)),
        "TM": max(tm_cfg.get("piso", 24), r_media(tm_cfg.get("factor_denominador", 0.0072) * den)),
    }


def _archivo_cache(token, mes, anio):
    nombre = f"{token}_{anio}_mes{mes}.json"
    p_res = CM.resolver_recurso(os.path.join("cache_api", nombre))
    if os.path.exists(p_res):
        return p_res
    return os.path.join(CACHE, nombre)


def _limpiar_nombre(raw):
    """'I-4   - 00002935 - ARAPA' -> 'I-4 - 00002935 - ARAPA' (como el .xls del tablero)."""
    return re.sub(r"\s+", " ", str(raw)).strip()


def pedir(token, mes, anio=2026, refrescar=False, reintentos=None, timeout=None):
    """Descarga (o lee de caché) el JSON crudo de la API para una credencial y un mes."""
    if not token:
        raise ErrorFicha32("Falta el código de usuario (token).")
    os.makedirs(CACHE, exist_ok=True)
    ruta = _archivo_cache(token, mes, anio)

    if not refrescar and os.path.exists(ruta):
        with open(ruta, encoding="utf-8") as fh:
            return json.load(fh)

    url_base = CM.get_api_url()
    timeout_sec = timeout if timeout is not None else CM.get_api_timeout()
    max_reintentos = reintentos if reintentos is not None else CM.get_api_reintentos()

    url = f"{url_base}?mes={mes}"
    ultimo = None
    for intento in range(max_reintentos):
        try:
            req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
            with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
                datos = json.load(resp)
            break
        except urllib.error.HTTPError as e:
            cuerpo = e.read().decode("utf-8", "replace")
            if e.code == 401:
                raise ErrorFicha32("La API no aceptó la credencial (401: falta o mal formada).")
            if e.code == 404:
                # No hay datos para esa combinación credencial/mes: no es un fallo de red.
                raise ErrorFicha32(f"No hay datos publicados para el mes {mes} con esa credencial.")
            ultimo = f"HTTP {e.code}: {cuerpo[:160]}"
        except urllib.error.URLError as e:
            ultimo = f"sin conexión ({e.reason})"
        except Exception as e:  # noqa: BLE001 - cualquier fallo de lectura se reintenta
            ultimo = str(e)
        time.sleep(1.2 * (intento + 1))
    else:
        if os.path.exists(ruta):
            try:
                with open(ruta, encoding="utf-8") as fh:
                    return json.load(fh)
            except Exception:
                pass
        raise ErrorFicha32(f"No se pudo consultar la API: {ultimo}")

    if datos.get("error"):
        raise ErrorFicha32(str(datos["error"]))
    with open(ruta, "w", encoding="utf-8") as fh:
        json.dump({**datos, "_descargado_en": time.strftime("%Y-%m-%d %H:%M:%S")}, fh,
                  ensure_ascii=False, indent=1)
    return datos


def filas_a_reporte(datos, mes):
    """Convierte el JSON de la API a las mismas columnas del .xls del tablero.

    Devuelve un DataFrame con la fila por establecimiento y las 13 columnas de
    COLS_REPORTE, separadas por tipo de indicador:
        LOGRO 'A' -> primer nivel, CSMC e hospitales (Reporte A)
        LOGRO 'B' -> hospitales e institutos (Reporte B)
    """
    filas = []
    for r in datos.get("data_establecimientos", []):
        m = metas(r.get("tipo_indicador"), r.get("denominador", 0))
        filas.append({
            "DIRESA": (r.get("diresa") or "").strip(),
            "RED": (r.get("red") or "").strip().upper(),
            "PROVINCIA": (r.get("provincia") or "").strip(),
            "ESTABLECIMIENTO / IPRESS": _limpiar_nombre(r.get("nombre_ipress")),
            "MES": int(r.get("mes") or mes),
            "LOGRO": r.get("tipo_indicador"),
            "DENOMINADOR": int(r.get("denominador") or 0),
            "TI REALIZADAS": int(r.get("ti") or 0), "TI META": m["TI"],
            "TC REALIZADAS": int(r.get("tc") or 0), "TC META": m["TC"],
            "TM REALIZADAS": int(r.get("tm") or 0), "TM META": m["TM"],
            # Columnas extra de la API (no están en el .xls; útiles para trazabilidad)
            "ID IPRESS": r.get("id_ipress"),
            "INDICADOR OFICIAL": r.get("indicador"),
            "ESTADO": r.get("estado"),
        })
    df = pd.DataFrame(filas, columns=COLS_REPORTE + ["ID IPRESS", "INDICADOR OFICIAL", "ESTADO"])
    return df.sort_values(["RED", "PROVINCIA", "ESTABLECIMIENTO / IPRESS"]).reset_index(drop=True)


def obtener(diresa_token, mes, anio=2026, refrescar=False):
    """(df_A, df_B, contexto) para una DIRESA y un mes, con la forma de los .xls del tablero."""
    datos = pedir(diresa_token, mes, anio=anio, refrescar=refrescar)
    df = filas_a_reporte(datos, mes)
    A = df[df["LOGRO"] == "A"].reset_index(drop=True)
    B = df[df["LOGRO"] == "B"].reset_index(drop=True)
    ctx = datos.get("user_context", {})
    ctx["nombre_diresa"] = NOMBRE_POR_TOKEN.get(diresa_token, diresa_token)
    ctx["kpis"] = datos.get("kpis_consolidados", {})
    ctx["filas"] = len(df)
    return A, B, ctx


def obtener_todos(diresa_token, mes, anio=2026, refrescar=False, avance=None):
    """{(mes): (df_A, df_B)} de enero hasta `mes`. Base para cualquier comparativa."""
    sal = {}
    for m in range(1, int(mes) + 1):
        if avance:
            avance(m, int(mes))
        sal[m] = obtener(diresa_token, m, anio=anio, refrescar=refrescar)
    return sal


def meses_disponibles(diresa_token, anio=2026, avance=None):
    """Lista de meses con datos publicados (los meses futuros responden 404)."""
    ok = []
    for m in MESES:
        if avance:
            avance(m, 12)
        try:
            pedir(diresa_token, m, anio=anio)
        except ErrorFicha32:
            continue
        ok.append(m)
    return ok


# ---------- NUEVA API RENIPRESS (SUSALUD) ----------
RENIPRESS_CACHE = os.path.join(CM.get_app_dir(), "cache_api", "renipress")

RENIPRESS_DEPTO = {
    "AMAZONAS": "01", "ÁNCASH": "02", "ANCASH": "02", "APURÍMAC": "03", "APURIMAC": "03",
    "AREQUIPA": "04", "AYACUCHO": "05", "CAJAMARCA": "06", "CALLAO": "07", "CUSCO": "08",
    "HUANCAVELICA": "09", "HUÁNUCO": "10", "HUANUCO": "10", "ICA": "11", "JUNÍN": "12", "JUNIN": "12",
    "LA LIBERTAD": "13", "LAMBAYEQUE": "14", "LIMA": "15", "LIMA REGIÓN": "15",
    "LIMA - DIRIS CENTRO": "15", "LIMA - DIRIS ESTE": "15", "LIMA - DIRIS NORTE": "15", "LIMA - DIRIS SUR": "15",
    "LORETO": "16", "MADRE DE DIOS": "17", "MOQUEGUA": "18", "PASCO": "19", "PIURA": "20",
    "PUNO": "21", "SAN MARTÍN": "22", "SAN MARTIN": "22", "TACNA": "23", "TUMBES": "24", "UCAYALI": "25"
}

RENIPRESS_AUTORIDAD = {
    "DIRESA_AMAZONAS": "1", "DIRESA_ANCASH": "2", "DIRESA_APURIMAC": "3", "DIRESA_AREQUIPA": "4",
    "DIRESA_AYACUCHO": "5", "DIRESA_CAJAMARCA": "7", "DIRESA_CALLAO": "8", "DIRESA_CUSCO": "11",
    "DIRESA_HUANCAVELICA": "13", "DIRESA_HUANUCO": "14", "DIRESA_ICA": "15", "DIRESA_JUNIN": "17",
    "DIRESA_LA_LIBERTAD": "18", "DIRESA_LAMBAYEQUE": "19", "DIRESA_LIMA_REGION": "22",
    "DIRESA_LIMA_DIRIS_CENTRO": "36", "DIRESA_LIMA_DIRIS_ESTE": "21", "DIRESA_LIMA_DIRIS_NORTE": "35",
    "DIRESA_LIMA_DIRIS_SUR": "23", "DIRESA_LORETO": "24", "DIRESA_MADRE_DE_DIOS": "25",
    "DIRESA_MOQUEGUA": "26", "DIRESA_PASCO": "27", "DIRESA_PIURA": "28", "DIRESA_PUNO": "29",
    "DIRESA_SAN_MARTIN": "30", "DIRESA_TACNA": "32", "DIRESA_TUMBES": "33", "DIRESA_UCAYALI": "34"
}


def _archivo_cache_renipress(token, mes=None, anio=2026):
    """Ruta del archivo de caché para la API RENIPRESS."""
    nombre = f"renipress_{token}.json"
    p_res = CM.resolver_recurso(os.path.join("cache_api", "renipress", nombre))
    if os.path.exists(p_res):
        return p_res
    return os.path.join(RENIPRESS_CACHE, nombre)

def _limpiar_texto(raw):
    """Normalizar texto: remover espacios múltiples y strip."""
    return re.sub(r"\s+", " ", str(raw or "")).strip()


def obtener_resoluciones_renipress(codigos, token="DIRESA_PUNO", session=None, refrescar=False):
    """Descarga y cachea los datos oficiales de Categorización y Cartera de Servicios desde SUSALUD RENIPRESS.

    Devuelve dict por código de 8 dígitos con:
        {
            codigo: {
                "nro_categorizacion": "RDR N° 1543-2024/...",
                "fecha_categorizacion": "25/07/2024",
                "subio_categorizacion": "SUBIÓ" | "NO SUBIÓ",
                "archivo_categorizacion": "ATUNCOLLA.pdf",
                "nro_cartera": "1364-2025/...",
                "fecha_cartera": "30/06/2025",
                "subio_cartera": "SUBIÓ" | "NO SUBIÓ",
                "archivo_cartera": "IPRESS I-3 ATUNCOLLA.pdf",
                "nro_resolucion": ...,
                "fecha_resolucion": ...,
                "variacion_resolucion": ...,
                "subio_resolucion": ...
            }
        }
    """
    if not codigos:
        return {}
    os.makedirs(RENIPRESS_CACHE, exist_ok=True)
    nombre_res = f"renipress_resoluciones_{token}.json"
    cache_file_read = CM.resolver_recurso(os.path.join("cache_api", "renipress", nombre_res))
    cache_file = cache_file_read if os.path.exists(cache_file_read) else os.path.join(RENIPRESS_CACHE, nombre_res)
    cached = {}
    if not refrescar and os.path.exists(cache_file):
        try:
            with open(cache_file, encoding="utf-8") as f:
                cached = json.load(f)
        except Exception:
            cached = {}

    missing = [str(c).strip().zfill(8) for c in codigos if str(c).strip().zfill(8) not in cached]
    if missing:
        from concurrent.futures import ThreadPoolExecutor

        def _crear_ses():
            s = requests.Session()
            s.headers.update({
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                "X-Requested-With": "XMLHttpRequest"
            })
            try:
                s.get("https://app20.susalud.gob.pe:8086/registro-renipress-webapp/listadoEstablecimientosRegistrados.htm?action=mostrarBuscar", timeout=12)
            except Exception:
                pass
            return s

        def _fetch_uno(c):
            url = "https://app20.susalud.gob.pe:8086/registro-renipress-webapp/ipress.htm?action=cargarIpress"
            for _ in range(2):
                try:
                    s = _crear_ses()
                    r = s.post(url, data={"idipress": c}, timeout=12)
                    time.sleep(0.04)  # Control de tasa (rate limiting)
                    d = r.json().get("datos", {}).get("p_crCURSOR_DATOS", [{}])[0]
                    cat_su = str(d.get("ADICIONAL_CATEGORIA", "") or "").strip()
                    # Categorizacion
                    nro_cat = str(d.get("ADICIONAL_CATEGORIA_NUMERO", "") or "").strip()
                    fec_cat = str(d.get("ADICIONAL_DOC_FECHA", "") or "").strip()
                    doc_cat = str(d.get("ADICIONAL_DOCUMENTO_DIGITAL", "") or "").strip()
                    nom_cat = str(d.get("ADICIONAL_DOC_NOMORIGINAL", "") or "").strip()
                    subio_cat = "SUBIÓ" if (doc_cat or nom_cat) else "NO SUBIÓ"

                    # Cartera de Servicios
                    nro_car = str(d.get("ADICIONAL_CATEGORIA_NUMEROCAR", "") or "").strip()
                    fec_car = str(d.get("ADICIONAL_DOC_FECHACAR", "") or "").strip()
                    doc_car = str(d.get("ADICIONAL_DOCUMENTO_DIGITALCAR", "") or "").strip()
                    nom_car = str(d.get("ADICIONAL_DOC_NOMORIGINALCAR", "") or "").strip()
                    subio_car = "SUBIÓ" if (doc_car or nom_car) else "NO SUBIÓ"

                    nro_res = nro_car if (nro_car and nro_car != "--") else nro_cat
                    fec_res = fec_car if fec_car else fec_cat
                    subio_res = "SUBIÓ" if (subio_car == "SUBIÓ" or subio_cat == "SUBIÓ") else "NO SUBIÓ"

                    return c, {
                        "categoria_susalud": cat_su,
                        "nro_categorizacion": nro_cat,
                        "fecha_categorizacion": fec_cat,
                        "subio_categorizacion": subio_cat,
                        "archivo_categorizacion": nom_cat,
                        "nro_cartera": nro_car,
                        "fecha_cartera": fec_car,
                        "subio_cartera": subio_car,
                        "archivo_cartera": nom_car,
                        "nro_resolucion": nro_res,
                        "fecha_resolucion": fec_res,
                        "variacion_resolucion": subio_car,
                        "subio_resolucion": subio_res
                    }
                except Exception:
                    time.sleep(0.4)
            return c, {
                "categoria_susalud": "",
                "nro_categorizacion": "", "fecha_categorizacion": "", "subio_categorizacion": "NO SUBIÓ",
                "archivo_categorizacion": "", "nro_cartera": "", "fecha_cartera": "",
                "subio_cartera": "NO SUBIÓ", "archivo_cartera": "", "nro_resolucion": "",
                "fecha_resolucion": "", "variacion_resolucion": "NO SUBIÓ", "subio_resolucion": "NO SUBIÓ"
            }

        with ThreadPoolExecutor(max_workers=5) as ex:
            for c, data in ex.map(_fetch_uno, missing):
                cached[c] = data

        try:
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump(cached, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    return cached


def pedir_renipress(token, mes=None, anio=2026, refrescar=False):
    """Descarga (o lee de caché) los establecimientos registrados en SUSALUD / RENIPRESS.

    Obtiene la relación de establecimientos de la DIRESA especificada junto con su
    condición oficial de Telesalud (CONSULTOR, CONSULTANTE, AMBOS o NO REGISTRADO).

    Retorna dict con estructura:
        {
            "token": token,
            "diresa": region_nombre,
            "mes": mes,
            "anio": anio,
            "datos": [filas parseadas],
            "fuente": "renipress",
            "_descargado_en": timestamp
        }
    """
    if not token:
        raise ErrorFicha32("Falta el código de usuario o token de región.")
    os.makedirs(RENIPRESS_CACHE, exist_ok=True)
    ruta = _archivo_cache_renipress(token, mes, anio)

    if not refrescar and os.path.exists(ruta):
        try:
            with open(ruta, encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            pass  # Si el archivo está corrupto, re-descargar

    # Resolver departamento y autoridad sanitaria para la consulta
    nombre_region = NOMBRE_POR_TOKEN.get(token, token)
    diresa_token = TOKEN_POR_NOMBRE.get(token, token)
    if not diresa_token.startswith("DIRESA_") and token in NOMBRE_POR_TOKEN:
        diresa_token = token

    cod_depto = RENIPRESS_DEPTO.get(nombre_region.upper(), "21")
    cod_autoridad = RENIPRESS_AUTORIDAD.get(diresa_token, "29")

    url_base = getattr(CM, "get_renipress_url", lambda: "https://app20.susalud.gob.pe:8086/registro-renipress-webapp/listadoEstablecimientosRegistrados.htm")()
    if "?" in url_base:
        url_base = url_base.split("?")[0]

    timeout_sec = max(20, CM.get_api_timeout())
    max_reintentos = CM.get_api_reintentos()

    if requests is None:
        raise ErrorFicha32("Se requiere la librería 'requests' para consultar la API de RENIPRESS.")

    ultimo_err = None
    for intento in range(max_reintentos):
        try:
            session = requests.Session()
            session.headers.update({
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                "X-Requested-With": "XMLHttpRequest"
            })

            # 1. Obtener cookie de sesión inicial
            r_init = session.get(f"{url_base}?action=mostrarBuscar", timeout=timeout_sec)
            if r_init.status_code != 200:
                raise ErrorFicha32(f"SUSALUD respondió con código {r_init.status_code} al inicializar.")

            # Parámetros base del formulario
            base_params = {
                "txt_filtrar": "",
                "cmb_estado": "1",  # Solo activos
                "cmb_departamento": cod_depto,
                "cmb_provincia": "",
                "cmb_distrito": "",
                "cmb_institucion": "0",  # Todas las instituciones
                "cmb_tipo_establecimiento": "0",
                "cmb_clasificacion": "",
                "cmb_categoria": "0",
                "cmb_unidadEjecutora": "0",
                "cmb_servicio": "0",
                "cmb_autoridadSanitaria": cod_autoridad,
                "cmb_red": "",
                "cmb_microRed": "",
                "cmb_clas": "0",
                "cmb_colegio": "0",
                "cmb_especialidad": "",
                "cmb_quintil": "0",
                "dat_fd_quintil": "",
                "ra_reg": "on",
                "dat_fd_desde": "",
                "dat_fd_hasta": "",
            }

            # Consultar conjuntos de Telesalud
            # 1: CONSULTOR, 2: CONSULTANTE
            consultores_set = set()
            consultantes_set = set()

            for r_code, r_nombre in [("1", "CONSULTOR"), ("2", "CONSULTANTE")]:
                p = dict(base_params)
                p["cmb_telesalud"] = r_code
                session.post(f"{url_base}?action=buscarEstablecimientos", data=p, timeout=timeout_sec)
                qs = urllib.parse.urlencode(p)
                url_cargar = f"{url_base}?action=cargarEstablecimientos&{qs}"
                resp = session.post(url_cargar, data={"draw": "1", "start": "0", "length": "1000"}, timeout=timeout_sec)
                if resp.status_code == 200:
                    jdata = resp.json().get("data", [])
                    for it in jdata:
                        c = str(it.get("codigounico", "")).strip().zfill(8)
                        if r_code == "1":
                            consultores_set.add(c)
                        else:
                            consultantes_set.add(c)

            # Consultar todos los establecimientos de la región (cmb_telesalud="0")
            p_todos = dict(base_params)
            p_todos["cmb_telesalud"] = "0"
            session.post(f"{url_base}?action=buscarEstablecimientos", data=p_todos, timeout=timeout_sec)
            qs_todos = urllib.parse.urlencode(p_todos)
            url_cargar_todos = f"{url_base}?action=cargarEstablecimientos&{qs_todos}"
            resp_todos = session.post(url_cargar_todos, data={"draw": "1", "start": "0", "length": "2000"}, timeout=timeout_sec)
            
            raw_items = []
            if resp_todos.status_code == 200:
                raw_items = resp_todos.json().get("data", [])

            # Si no devolvió con cmb_telesalud="0", usar la consulta cmb_telesalud="3" (Telesalud)
            if not raw_items:
                p_tele = dict(base_params)
                p_tele["cmb_telesalud"] = "3"
                session.post(f"{url_base}?action=buscarEstablecimientos", data=p_tele, timeout=timeout_sec)
                qs_tele = urllib.parse.urlencode(p_tele)
                url_cargar_tele = f"{url_base}?action=cargarEstablecimientos&{qs_tele}"
                resp_tele = session.post(url_cargar_tele, data={"draw": "1", "start": "0", "length": "1000"}, timeout=timeout_sec)
                if resp_tele.status_code == 200:
                    raw_items = resp_tele.json().get("data", [])

            # Obtener resoluciones de cartera de servicios (desde caché o API)
            all_raw_codes = [str(it.get("codigounico", "")).strip().zfill(8) for it in raw_items if it.get("codigounico")]
            res_map = obtener_resoluciones_renipress(all_raw_codes, token=token, session=session)

            # Construir registros unificados
            datos = []
            seen_codes = set()
            for it in raw_items:
                cod = str(it.get("codigounico", "")).strip().zfill(8)
                if not cod or cod in seen_codes:
                    continue
                seen_codes.add(cod)

                es_consultor = cod in consultores_set
                es_consultante = cod in consultantes_set

                if es_consultor and es_consultante:
                    rol = "AMBOS"
                elif es_consultor:
                    rol = "CONSULTOR"
                elif es_consultante:
                    rol = "CONSULTANTE"
                else:
                    rol = "NO REGISTRADO"

                r_info = res_map.get(cod, {})
                nro_res = r_info.get("nro_resolucion") or _limpiar_texto(it.get("numeroCategoria", ""))
                fec_res = r_info.get("fecha_resolucion") or _limpiar_texto(it.get("fecha_inspecion", ""))
                var_res = r_info.get("variacion_resolucion") or ("SUBIÓ" if nro_res else "NO SUBIÓ")

                datos.append({
                    "id_ipress": cod,
                    "codigo_unico": cod,
                    "nombre_ipress": _limpiar_texto(it.get("nombrerazonsocial", "")),
                    "diresa": nombre_region,
                    "red": _limpiar_texto(it.get("red", "")),
                    "provincia": _limpiar_texto(it.get("provincia", "")),
                    "distrito": _limpiar_texto(it.get("distrito", "")),
                    "telemedicina": rol,
                    # Categoría oficial SUSALUD
                    "categoria_susalud": r_info.get("categoria_susalud", ""),
                    # Categorización oficial
                    "nro_categorizacion": r_info.get("nro_categorizacion") or _limpiar_texto(it.get("numeroCategoria", "")),
                    "fecha_categorizacion": r_info.get("fecha_categorizacion") or _limpiar_texto(it.get("fecha_inspecion", "")),
                    "subio_categorizacion": r_info.get("subio_categorizacion", "NO SUBIÓ"),
                    "archivo_categorizacion": r_info.get("archivo_categorizacion", ""),
                    # Cartera de Servicios oficial
                    "nro_cartera": r_info.get("nro_cartera", ""),
                    "fecha_cartera": r_info.get("fecha_cartera", ""),
                    "subio_cartera": r_info.get("subio_cartera", "NO SUBIÓ"),
                    "archivo_cartera": r_info.get("archivo_cartera", ""),
                    # Consolidado
                    "nro_resolucion": nro_res,
                    "fecha_resolucion": fec_res,
                    "variacion_resolucion": var_res,
                    "subio_resolucion": r_info.get("subio_resolucion", "NO SUBIÓ"),
                    "estado": _limpiar_texto(it.get("estado", "ACTIVO")),
                    "clasificacion": r_info.get("categoria_susalud", "") or _limpiar_texto(it.get("categoria", "") or it.get("clasificacion", "")),
                    "direccion": _limpiar_texto(it.get("direccion", "")),
                    "ubigeo": _limpiar_texto(it.get("ubigeo", "")),
                })

            resultado = {
                "token": token,
                "diresa": nombre_region,
                "mes": mes,
                "anio": anio,
                "datos": datos,
                "fuente": "renipress",
                "_descargado_en": time.strftime("%Y-%m-%d %H:%M:%S"),
            }
            with open(ruta, "w", encoding="utf-8") as fh:
                json.dump(resultado, fh, ensure_ascii=False, indent=2)
            return resultado

        except Exception as e:
            ultimo_err = str(e)
            time.sleep(1.2 * (intento + 1))

    # Si falló la red, intentar usar caché antiguo si existe
    if os.path.exists(ruta):
        try:
            with open(ruta, encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            pass

    raise ErrorFicha32(f"No se pudo consultar la API RENIPRESS / SUSALUD: {ultimo_err}")


def parsear_renipress(html):
    """Compatibilidad: parsear HTML si se provee directamente."""
    if not html:
        return []
    datos = []
    if BeautifulSoup:
        soup = BeautifulSoup(html, "html.parser")
        for tr in soup.find_all("tr")[1:]:
            cols = [c.get_text(strip=True) for c in tr.find_all("td")]
            if len(cols) >= 5:
                datos.append({
                    "id_ipress": cols[0].zfill(8),
                    "codigo_unico": cols[0].zfill(8),
                    "nombre_ipress": cols[1] if len(cols) > 1 else "",
                    "diresa": cols[2] if len(cols) > 2 else "",
                    "provincia": cols[3] if len(cols) > 3 else "",
                    "distrito": cols[4] if len(cols) > 4 else "",
                    "telemedicina": "AMBOS",
                    "estado": "ACTIVO",
                    "clasificacion": ""
                })
    return datos


def filtrar_telemedicina(datos, rol=None):
    """Filtrar datos RENIPRESS por rol de telemedicina.

    Args:
        datos: lista de dicts de establecimientos
        rol: "CONSULTOR", "CONSULTANTE", "AMBOS" o None/'TODOS'
    """
    if not datos:
        return []
    if not rol or str(rol).upper() in ("TODOS", "ALL", "0", ""):
        return [d for d in datos if d.get("telemedicina") in ("CONSULTOR", "CONSULTANTE", "AMBOS")]
    rol_u = str(rol).upper().strip()
    if rol_u == "AMBOS":
        return [d for d in datos if d.get("telemedicina") == "AMBOS"]
    elif rol_u in ("CONSULTOR", "CONSULTANTE"):
        # Un establecimiento con rol AMBOS también califica como Consultor y como Consultante
        return [d for d in datos if d.get("telemedicina") in (rol_u, "AMBOS")]
    return [d for d in datos if d.get("telemedicina") == rol_u]


def filtrar_resolucion_sucremec(datos):
    """Resumen estadístico de acreditación Telesalud y resoluciones RENIPRESS."""
    if not datos:
        return {
            "total_ipress": 0,
            "con_telesalud": 0,
            "consultores": 0,
            "consultantes": 0,
            "ambos": 0,
            "sin_telesalud": 0,
            "con_resolucion": 0,
            "pct_con_telesalud": 0.0,
        }

    total = len(datos)
    con_tele = [d for d in datos if d.get("telemedicina") in ("CONSULTOR", "CONSULTANTE", "AMBOS")]
    consultores = [d for d in datos if d.get("telemedicina") == "CONSULTOR"]
    consultantes = [d for d in datos if d.get("telemedicina") == "CONSULTANTE"]
    ambos = [d for d in datos if d.get("telemedicina") == "AMBOS"]
    con_res = [d for d in datos if d.get("nro_resolucion")]

    return {
        "total_ipress": total,
        "con_telesalud": len(con_tele),
        "consultores": len(consultores),
        "consultantes": len(consultantes),
        "ambos": len(ambos),
        "sin_telesalud": total - len(con_tele),
        "con_resolucion": len(con_res),
        "pct_con_telesalud": round(100.0 * len(con_tele) / total, 1) if total > 0 else 0.0,
    }


def obtener_datos_renipress(token, mes=None, anio=2026, rol=None, refrescar=False):
    """Función de alto nivel: descarga, parsea y filtra los datos de SUSALUD RENIPRESS."""
    raw = pedir_renipress(token, mes=mes, anio=anio, refrescar=refrescar)
    filtrados = filtrar_telemedicina(raw["datos"], rol=rol)
    resumen = filtrar_resolucion_sucremec(raw["datos"])
    return {
        "token": raw["token"],
        "diresa": raw.get("diresa", token),
        "mes": raw.get("mes", mes),
        "anio": raw.get("anio", anio),
        "fuente": "renipress",
        "establecimientos": filtrados,
        "todos_establecimientos": raw["datos"],
        "resumen_resolucion": resumen,
        "total_establecimientos": len(filtrados),
        "total_region": len(raw["datos"]),
    }


def datos_a_dataframe(datos):
    """Convertir lista de dicts RENIPRESS a DataFrame pandas con columnas estandarizadas."""
    if not datos:
        return pd.DataFrame(columns=[
            "id_ipress", "codigo_unico", "nombre_ipress", "diresa", "red", "provincia",
            "distrito", "telemedicina", "nro_resolucion", "fecha_resolucion",
            "variacion_resolucion", "estado", "clasificacion", "direccion", "ubigeo"
        ])
    df = pd.DataFrame(datos)
    columnas_esperadas = [
        "id_ipress", "codigo_unico", "nombre_ipress", "diresa", "red", "provincia",
        "distrito", "telemedicina", "nro_resolucion", "fecha_resolucion",
        "variacion_resolucion", "estado", "clasificacion", "direccion", "ubigeo"
    ]
    for c in columnas_esperadas:
        if c not in df.columns:
            df[c] = ""
    return df[columnas_esperadas]


if __name__ == "__main__":
    token = TOKEN_POR_NOMBRE.get("PUNO", "DIRESA_PUNO")
    print(f"Probando conexión a Ficha 32 y RENIPRESS para {token}...")
    try:
        ren = obtener_datos_renipress(token, refrescar=False)
        print("RENIPRESS descargado con éxito:")
        print(f"  Total establecimientos región: {ren['total_region']}")
        print(f"  Establecimientos con Telesalud: {ren['total_establecimientos']}")
        print(f"  Resumen: {ren['resumen_resolucion']}")
    except Exception as e:
        print(f"Error RENIPRESS: {e}")
