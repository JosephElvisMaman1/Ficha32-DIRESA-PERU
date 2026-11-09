# -*- coding: utf-8 -*-
"""Extrae los nombres propios de los hospitales (IPRESS de categoría II) de una DIRESA.

Sirve para generar la tabla NOMBRES_HOSPITAL de ficha32_modelo.py, que corrige los
nombres sin acentos que llegan desde la API (p. ej. 'MANUEL NUÑEZ BUTRÓN').
"""
import sys
import unicodedata

import ficha32_api as api


def partes(s):
    p = str(s).split(" - ")
    return p[1].strip(), p[-1].strip()


def a_nombre_propio(txt):
    """Pasa el nombre a formato título y repara los acentos de las mayúsculas."""
    s = " ".join(str(txt).split()).title()
    correcciones = {
        "Núñez": "Núñez", "Nuñez": "Núñez", "Nunez": "Núñez",
        "Manuel Nuñez Butron": "Manuel Núñez Butrón",
        "Butron": "Butrón", "Jose": "José", "Jesus": "Jesús",
        "Monje Medrano": "Monje Medrano", "Asencion": "Asención",
        "Salud Mental Comunitario": "Salud Mental Comunitario",
        "De Dios": "de Dios", "Del": "del", "De": "de", "La": "la",
        "El": "el", "Los": "los", "Las": "las", "San": "San", "Santa": "Santa",
    }
    for mal, bien in correcciones.items():
        s = s.replace(mal, bien)
    return s


if __name__ == "__main__":
    nombre = sys.argv[1] if len(sys.argv) > 1 else "PUNO"
    mes = int(sys.argv[2]) if len(sys.argv) > 2 else 8
    token = api.TOKEN_POR_NOMBRE.get(nombre, f"DIRESA_{nombre}")
    A, B, ctx = api.obtener(token, mes)
    print(f"# Nombres propios de hospitales — DIRESA {ctx['nombre_diresa']} (mes {mes})\n")
    print("NOMBRES_HOSPITAL = {")
    for _, r in B.iterrows():
        cod, nom = partes(r["ESTABLECIMIENTO / IPRESS"])
        print(f'    "{cod}": "{a_nombre_propio(nom)}",')
    ii = A[A["ESTABLECIMIENTO / IPRESS"].str.startswith(("II-1", "II-2"))]
    print("}\n")
    print("# --- categorías II que también figuran en el Reporte A ---")
    for _, r in ii.iterrows():
        cod, nom = partes(r["ESTABLECIMIENTO / IPRESS"])
        print(f'    "{cod}": "{a_nombre_propio(nom)}",   # red={r["RED"]}')
    # Detenerse si el normalizador rompió algo con acentos
    print("\n# Revisión: nombres con posibles acentos faltantes")
    for df_, col in ((B, "B"), (ii, "A")):
        for _, r in df_.iterrows():
            cod, nom = partes(r["ESTABLECIMIENTO / IPRESS"])
            if any(c in nom.upper() for c in "ÑÁÉÍÓÚ") or "NUNEZ" in nom.upper() or "NUÑEZ" in nom.upper():
                print(f"    [{col}] {cod} crudo='{nom}' -> propuesto='{a_nombre_propio(nom)}'")
