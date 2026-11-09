# -*- coding: utf-8 -*-
"""Nombres propios de hospitales e institutos — Ficha N.° 32.

La API del tablero entrega los nombres en mayúsculas y sin acentos
("HOSPITAL REGIONAL MANUEL NUÑEZ BUTRON"). Aquí se corrigen para los informes.

ARCHIVO EDITABLE: agréguese una línea por cada IPRESS de categoría II (hospitales e
institutos) que se desee mostrar con su nombre oficial. La clave es el código IPRESS
de 8 dígitos que aparece en el nombre (p. ej. "II-2 - 00003251 - ..." -> "00003251").

Si un código no está en esta tabla, el informe usa el nombre que viene de la API tal cual.
Para ver los códigos de tu DIRESA:  python extraer_hospitales.py TU_DIRESA 8
"""

# ---------------------------------------------------------------- DIRESA PUNO
NOMBRES_HOSPITAL = {
    # Reporte B (hospitales e institutos - Puno)
    "00003251": "Hospital Regional Manuel Núñez Butrón",
    "00003299": "Hospital Carlos Monje Medrano de Juliaca",
    "00031179": "Hospital San Martín de Porres de Macusani",
    "00035288": "Hospital San Juan de Dios de Ayaviri",
    "00003032": "Hospital de Ilave",
    # Categorías II que también figuran en el Reporte A (primer nivel - Puno)
    "00002941": "Hospital Carlos Cornejo Rosello",
    "00003068": "Hospital Lucio Aldazábal Pauca",
    "00003128": "Hospital Antonio Barrionuevo de Lampa",
    "00003349": "Hospital de Yunguyo",
    # Reporte B (Arequipa)
    "00001231": "Hospital Regional Honorio Delgado Espinoza",
    "00001232": "Hospital Goyeneche",
    "00007397": "Instituto Regional de Enfermedades Neoplásicas del Sur (IREN Sur)",
    "00016721": "Hospital Central de Majes Ing. Ángel Gabriel Chura Gallegos",
}

# Hospital consultor de referencia por red de salud (se usa para redactar las estrategias)
HOSPITAL_POR_RED = {
    "PUNO": "Hospital Regional Manuel Núñez Butrón",
    "SAN ROMAN": "Hospital Carlos Monje Medrano de Juliaca",
    "CARABAYA": "Hospital San Martín de Porres de Macusani",
    "MACUSANI": "Hospital San Martín de Porres de Macusani",
    "EL COLLAO": "Hospital de Ilave",
    "COLLAO": "Hospital de Ilave",
    "MELGAR": "Hospital San Juan de Dios de Ayaviri",
    "AZANGARO": "Hospital Carlos Cornejo Rosello",
    "LAMPA": "Hospital Antonio Barrionuevo de Lampa",
    "HUANCANE": "Hospital Lucio Aldazábal Pauca",
    "YUNGUYO": "Hospital de Yunguyo",
    "AREQUIPA CAYLLOMA": "Hospital Regional Honorio Delgado Espinoza",
}


def cargar():
    """Devuelve las dos tablas; si el archivo fue editado por el usuario, respeta lo nuevo."""
    return dict(NOMBRES_HOSPITAL), dict(HOSPITAL_POR_RED)


if __name__ == "__main__":
    print(f"{len(NOMBRES_HOSPITAL)} hospitales con nombre propio, "
          f"{len(HOSPITAL_POR_RED)} redes con hospital consultor.")
    for cod, nom in sorted(NOMBRES_HOSPITAL.items()):
        print(f"  {cod}  {nom}")
