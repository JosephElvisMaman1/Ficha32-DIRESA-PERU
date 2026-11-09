# -*- coding: utf-8 -*-
"""Gestor de configuración centralizada — Ficha N.° 32 (Telemedicina).

Permite modificar la URL de la API (para 2026, 2027 o nuevos endpoints),
los pesos de servicios (TI, TC, TM), las metas mínimas y los parámetros
de cálculo sin necesidad de alterar el código fuente.
"""
import json
import os
import sys


def get_bundle_dir():
    """Directorio de recursos internos empaquetados (PyInstaller _MEIPASS o carpeta del código)."""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
    return os.path.dirname(os.path.abspath(__file__))


def get_app_dir():
    """Directorio persistente de trabajo (junto al .exe o raíz del proyecto)."""
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        parent_dir = os.path.dirname(exe_dir)
        # Si se ejecuta desde la subcarpeta dist/ dentro del proyecto, usar la raíz del proyecto
        if os.path.basename(exe_dir).lower() == "dist" and os.path.exists(os.path.join(parent_dir, "config.json")):
            return parent_dir
        return exe_dir
    return os.path.dirname(os.path.abspath(__file__))


def resolver_recurso(rel_path):
    """Devuelve la ruta existente priorizando APP_DIR (persistente) y luego BUNDLE_DIR (empaquetado)."""
    p_app = os.path.join(get_app_dir(), rel_path)
    if os.path.exists(p_app):
        return p_app
    p_bun = os.path.join(get_bundle_dir(), rel_path)
    if os.path.exists(p_bun):
        return p_bun
    return p_app


BASE_DIR = get_app_dir()
BUNDLE_DIR = get_bundle_dir()
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

DEFAULT_SERVICIOS = [
    {
        "codigo": "TI",
        "nombre": "Teleinterconsulta",
        "peso": 60,
        "activo": True,
        "metas_tipo_a": {"piso": 24, "factor_denominador": 0.0120},
        "metas_tipo_b": {"piso": 100, "factor_denominador": 0.0040}
    },
    {
        "codigo": "TC",
        "nombre": "Teleconsulta",
        "peso": 10,
        "activo": True,
        "metas_tipo_a": {"piso": 1, "factor_denominador": 0.0010},
        "metas_tipo_b": None
    },
    {
        "codigo": "TM",
        "nombre": "Telemonitoreo",
        "peso": 30,
        "activo": True,
        "metas_tipo_a": {"piso": 24, "factor_denominador": 0.0072},
        "metas_tipo_b": None
    }
]

DEFAULT_CONFIG = {
    "api": {
        "url": "https://api-ficha32-ditel-303392539829.europe-west1.run.app",
        "renipress_url": "https://app20.susalud.gob.pe:8086/registro-renipress-webapp/listadoEstablecimientosRegistrados.htm?action=mostrarBuscar",
        "timeout_segundos": 15,
        "reintentos": 2
    },
    "servicios": DEFAULT_SERVICIOS,
    "formulas": {
        "pesos": {
            "TI": 60,
            "TC": 10,
            "TM": 30
        },
        "metas_tipo_a": {
            "TI": {"piso": 24, "factor_denominador": 0.0120},
            "TC": {"piso": 1, "factor_denominador": 0.0010},
            "TM": {"piso": 24, "factor_denominador": 0.0072}
        },
        "metas_tipo_b": {
            "TI": {"piso": 100, "factor_denominador": 0.0040}
        },
        "semaforo": {
            "verde_min": 60.0,
            "amarillo_min": 30.0
        },
        "cumplimiento_oficial": {
            "base_restar": 10.0,
            "divisor": 90.0
        }
    },
    "plan": {
        "mes_inicio_plan": 11,
        "mes_fin_plan": 12,
        "piso_capacidad_mes": 12,
        "pct_capacidad_meta": 0.05,
        "ratio_esfuerzo_cerca": 0.3333,
        "fraccion_arranque": 0.3333
    }
}


def _deep_update(d, u):
    for k, v in u.items():
        if isinstance(v, dict) and isinstance(d.get(k), dict):
            _deep_update(d[k], v)
        else:
            d[k] = v
    return d


def cargar_config():
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    ruta_cfg = resolver_recurso("config.json")
    if os.path.exists(ruta_cfg):
        try:
            with open(ruta_cfg, "r", encoding="utf-8") as f:
                user_cfg = json.load(f)
            _deep_update(cfg, user_cfg)
        except Exception as e:
            if getattr(sys, "stdout", None) is not None:
                print(f"[aviso] No se pudo leer config.json: {e}. Se usan valores por defecto.")
    # Variable de entorno opcional tiene prioridad sobre el JSON
    env_api = os.environ.get("FICHA32_API_URL")
    if env_api:
        cfg["api"]["url"] = env_api.strip()
    return cfg


def guardar_config(nueva_cfg):
    cfg = cargar_config()
    _deep_update(cfg, nueva_cfg)
    os.makedirs(os.path.dirname(os.path.abspath(CONFIG_FILE)), exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return cfg


def get_api_url():
    return cargar_config()["api"]["url"]


def get_renipress_url():
    return cargar_config()["api"].get(
        "renipress_url",
        "https://app20.susalud.gob.pe:8086/registro-renipress-webapp/listadoEstablecimientosRegistrados.htm?action=mostrarBuscar",
    )


def get_api_timeout():
    return int(cargar_config()["api"].get("timeout_segundos", 15))


def get_api_reintentos():
    return int(cargar_config()["api"].get("reintentos", 2))


def get_servicios():
    """Retorna la lista dinámica de servicios configurados (ej. TI, TC, TM y nuevos para 2027)."""
    cfg = cargar_config()
    if "servicios" in cfg and isinstance(cfg["servicios"], list) and len(cfg["servicios"]) > 0:
        return cfg["servicios"]
    pesos = cfg.get("formulas", {}).get("pesos", {})
    servicios = []
    for s in DEFAULT_SERVICIOS:
        item = dict(s)
        if s["codigo"] in pesos:
            item["peso"] = pesos[s["codigo"]]
        servicios.append(item)
    return servicios


def get_servicio_codigos():
    """Lista ordenada de códigos de servicios activos, ej. ['TI', 'TC', 'TM']."""
    return [s["codigo"] for s in get_servicios() if s.get("activo", True)]


def get_servicio_nombres():
    """Diccionario de código -> nombre descriptivo del servicio."""
    return {s["codigo"]: s.get("nombre", s["codigo"]) for s in get_servicios()}


def guardar_servicios_pesos(pesos_dict):
    """Actualiza los pesos de los servicios tanto en la lista dinámica como en formulas.pesos."""
    cfg = cargar_config()
    if "servicios" not in cfg or not isinstance(cfg["servicios"], list):
        cfg["servicios"] = get_servicios()
    for s in cfg["servicios"]:
        cod = s.get("codigo")
        if cod in pesos_dict:
            s["peso"] = int(pesos_dict[cod])
    if "formulas" not in cfg:
        cfg["formulas"] = {}
    cfg["formulas"]["pesos"] = {k: int(v) for k, v in pesos_dict.items()}
    guardar_config(cfg)
    return cfg


def get_pesos():
    """Retorna el diccionario de pesos {codigo: peso} derivado de los servicios activos."""
    cfg = cargar_config()
    if "servicios" in cfg and isinstance(cfg["servicios"], list) and len(cfg["servicios"]) > 0:
        return {s["codigo"]: s.get("peso", 0) for s in cfg["servicios"] if s.get("activo", True)}
    return cfg.get("formulas", {}).get("pesos", {"TI": 60, "TC": 10, "TM": 30})


def get_metas_config():
    return cargar_config()["formulas"]


def get_plan_config():
    return cargar_config()["plan"]
