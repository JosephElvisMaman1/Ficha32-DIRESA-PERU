# -*- coding: utf-8 -*-
"""Generador maestro de informes — Ficha N.° 32 (Telemedicina).

Descarga los datos del tablero MINSA, los modela y produce el libro Excel, los informes
por red de salud (Word/PDF) y los gráficos. Lo usa tanto la aplicación gráfica
(app_ficha32.py) como la línea de comandos.

Uso de consola:
    python generar_todo.py --diresa PUNO --mes 8 --anio 2026 --excel --evolucion
    python generar_todo.py --diresa PUNO --mes 8 --word --graficos --alcance red:AZANGARO
    python generar_todo.py --diresa PUNO --mes 8 --alcance ipress:00002936 --excel
    python generar_todo.py --diresa PUNO --mes 8 --excel --sin-plan --sin-redes   # solo situación actual
    python generar_todo.py --diresa PUNO --mes 8 --excel --evolucion --evo-desde 4 --evo-hasta 6

Salida: líneas en stdout con el prefijo
    [progreso] n  →  para la barra de progreso de la interfaz
    [info] texto  →  mensaje informativo
    [ok] ruta     →  archivo generado
    [error] texto →  fallo; termina con código distinto de cero
"""
import argparse
import os
import re
import sys
import time

if getattr(sys, "stdout", None) is not None and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if getattr(sys, "stderr", None) is not None and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import ficha32_modelo as M  # noqa: E402

OUT_CALLBACK = None


def out(txt):
    if OUT_CALLBACK is not None:
        try:
            OUT_CALLBACK(str(txt))
        except Exception:
            pass
    if getattr(sys, "stdout", None) is not None:
        try:
            print(txt, flush=True)
        except Exception:
            pass


def progreso(n, total, etiqueta=""):
    out(f"[progreso] {n}/{total} {etiqueta}".rstrip())


def resolver_slug_alcance(diresa, alcance=None):
    """Devuelve un identificador limpio y descriptivo del alcance solicitado."""
    if not alcance:
        return f"Toda_la_DIRESA_{diresa}"
    if isinstance(alcance, str):
        alcance = [alcance]
    for a in alcance:
        if ":" in a:
            tipo, valor = a.split(":", 1)
            tipo, valor = tipo.strip().lower(), valor.strip()
            if tipo == "red":
                slug_val = re.sub(r'[^a-zA-Z0-9_]+', '_', valor.upper().replace(" ", "_")).strip('_')
                return f"Red_{slug_val}"
            elif tipo in ("ipress", "establecimiento"):
                cods = [c.strip() for c in valor.split(",") if c.strip()]
                if len(cods) == 1:
                    return f"IPRESS_{cods[0]}"
                return f"Seleccion_{len(cods)}_IPRESS"
    return f"DIRESA_{diresa}"


def carpeta_salida(diresa, mes, anio, personalizada=None, alcance_slug=None, timestamp=None):
    """Crea una carpeta de salida única para la ejecución, identificando red/alcance y fecha."""
    raiz = personalizada or os.path.join(M.BASE, "informes")
    base_mes = os.path.join(raiz, f"DIRESA_{diresa}", f"{anio}", f"Mes {mes:02d}")
    ts = timestamp or time.strftime("%Y-%m-%d_%H-%M-%S")
    slug = alcance_slug or f"Toda_la_DIRESA_{diresa}"
    slug_limpio = re.sub(r'[<>:"/\\|?*]+', '_', str(slug)).strip('_')
    nombre_carpeta = f"{slug_limpio}_{ts}"
    destino = os.path.join(base_mes, nombre_carpeta)
    os.makedirs(destino, exist_ok=True)
    return destino


def generar(diresa, mes, anio=2026, excel=True, word=False, graficos=False, evolucion=True,
            alcance=None, carpeta=None, refrescar=False, avance=None,
            plan=True, hospitales=True, redes=True, evo_desde=1, evo_hasta=None,
            plan_desde=None, plan_hasta=None, cartera=True):
    """Ejecuta el pipeline completo. Devuelve la lista de archivos generados.

    `alcance`: None, una cadena ("red:AZANGARO" o "ipress:C1,C2") o una lista de ellas.
    `plan`, `hospitales` y `redes`: qué partes incluye el libro Excel (ver generar_excel.build).
    `evo_desde` / `evo_hasta`: rango de meses de la hoja Evolución (por defecto 1..mes de corte).
    `plan_desde` / `plan_hasta`: ventana de meses del Plan de Cumplimiento (ej. 11 a 12 o 7 a 12).
    """
    if isinstance(alcance, str):
        alcance = [alcance]
    inicio = time.time()
    salidas = []
    slug_alcance = resolver_slug_alcance(diresa, alcance)
    ts_ejecucion = time.strftime("%Y-%m-%d_%H-%M-%S")
    destino = carpeta_salida(diresa, mes, anio, carpeta, alcance_slug=slug_alcance, timestamp=ts_ejecucion)
    out(f"[carpeta] {destino}")

    # 1) Configurar el modelo ANTES de importar los generadores:
    #    ellos congelan rutas y etiquetas en el momento del import.
    etq = M.configurar(mes=mes, anio=anio, diresa=diresa, fuente="api",
                       mes_ini_plan=plan_desde, mes_fin_plan=plan_hasta)
    out(f"[info] DIRESA {diresa} · {etq['MES']} {anio} · proyección a {etq['HASTA']} · plan {etq['PLAN_LARGO']}")

    # 2) Descarga de datos (corte + mes anterior; la serie del rango de evolución si se pide)
    import ficha32_api as api
    token = api.TOKEN_POR_NOMBRE.get(diresa, f"DIRESA_{diresa}")
    if evolucion:
        # El rango pedido se amplía un mes hacia atrás: su acumulado es la base para
        # diferenciar la producción del primer mes del rango.
        desde = max(1, int(evo_desde or 1))
        hasta = min(int(mes), int(evo_hasta or mes))
        if desde > hasta:
            desde, hasta = hasta, desde
        meses_evo = list(range(max(1, desde - 1), hasta + 1))
        # El modelo del plan siempre necesita el mes de corte y el anterior
        meses_pedir = sorted(set(meses_evo) | {int(mes), max(1, int(mes) - 1)})
        out(f"[info] Descargando la serie M{desde}→M{hasta} ({len(meses_pedir)} meses)…")
        serie = {}
        for m in meses_pedir:
            if avance:
                avance(m, len(meses_pedir))
            serie[m] = api.obtener(token, m, anio=anio, refrescar=refrescar)
        serie_evo = {m: serie[m] for m in meses_evo if m in serie}
    else:
        meses_pedir = sorted({int(mes), max(1, int(mes) - 1)})
        serie = {m: api.obtener(token, m, anio=anio, refrescar=refrescar) for m in meses_pedir}
        serie_evo = None
    out(f"[progreso] {len(meses_pedir)}/{len(meses_pedir)} datos listos")
    # Solo los dos meses que usa el modelo de plan
    M.SERIE = {m: serie[m] for m in (int(mes), max(1, int(mes) - 1)) if m in serie}

    # 3) Alcance (toda la DIRESA, una o varias redess, o una lista de establecimientos)
    M.FILTRO["red"] = None
    M.FILTRO["codigos"] = []
    for a in (alcance or []):
        if ":" not in a:
            raise ValueError(f"Alcance mal formado: {a} (use red:NOMBRE o ipress:COD1,COD2)")
        tipo, valor = a.split(":", 1)
        tipo, valor = tipo.strip().lower(), valor.strip()
        if tipo == "red":
            M.FILTRO["red"] = valor
        elif tipo in ("ipress", "establecimiento"):
            M.FILTRO["codigos"] = [c.strip() for c in valor.split(",") if c.strip()]
        else:
            raise ValueError(f"Alcance desconocido: {tipo}")
        out(f"[info] Alcance: {tipo} {valor[:80]}{'…' if len(valor) > 80 else ''}")

    A, B = M.construir()
    out(f"[info] Padrón modelado: {len(A)} IPRESS de primer nivel y {len(B)} hospitales/institutos")
    if len(A) == 0 and len(B) == 0:
        raise RuntimeError("El alcance seleccionado no contiene establecimientos.")
    if len(A) == 0:
        raise RuntimeError("El alcance no contiene IPRESS de primer nivel; pruebe con otro nombre de red o código.")

    # 4) Libro Excel (con la hoja «Evolución» integrada cuando se pide)
    if excel:
        import generar_excel
        importlib_reload(generar_excel)
        generar_excel.ESPERADO = M.MES_CORTE / M.MESES_AÑO
        nombre_excel = f"Ficha32_Telemedicina_{diresa}_Mes{mes}_{slug_alcance}.xlsx"
        ruta = os.path.join(destino, nombre_excel)
        generar_excel.build(serie_evo if evolucion else None, plan=plan, hospitales=hospitales,
                            redes=redes, evo_desde=desde if evolucion else evo_desde, out_path=ruta,
                            cartera=cartera)
        if evolucion:
            out(f"[info] Hoja 'Evolución' integrada en el libro (M{desde}→M{hasta})")
        if not (plan and hospitales and redes):
            omitidas = [n for n, ok in (("planificación", plan), ("hospitales", hospitales),
                                        ("parte de redes", redes)) if not ok]
            out(f"[info] Libro sin: {', '.join(omitidas)}")
        salidas.append(ruta)
        out(f"[ok] {ruta}")

    # 5) Gráficos PNG por red (también necesarios para incrustar en los Word)
    if graficos or word:
        import ficha32_graficos as G
        importlib_reload(G)
        G.OUT = os.path.join(destino, "graficos")
        os.makedirs(G.OUT, exist_ok=True)

    if graficos:
        G.graficos_red(A, f"DIRESA {diresa} ({len(A)} IPRESS)", "REGIONAL")
        for prov, slug, disp in M.redes(A):
            sub = A[A["RED_SALUD"] == prov]
            if len(sub):
                G.graficos_red(sub, f"Red de Salud {disp}", slug)
        out(f"[ok] {G.OUT}")

    # 6) Informes Word (Informe Regional + informes por red de salud)
    if word:
        import generar_reportes_redes as GR
        importlib_reload(GR)
        GR.ESPERADO = M.MES_CORTE / M.MESES_AÑO * 100
        GR.OUT = os.path.join(destino, "words")
        os.makedirs(GR.OUT, exist_ok=True)
        solo_red = [M.FILTRO["red"]] if M.FILTRO.get("red") else None
        informes = GR.main(solo=solo_red, incluir_regional=(M.FILTRO.get("red") is None))
        salidas.extend(informes)
        out(f"[ok] {GR.OUT}")

    # Guardar manifiesto de resumen de entrega
    try:
        manifiesto_path = os.path.join(destino, "resumen_entrega.txt")
        with open(manifiesto_path, "w", encoding="utf-8") as f_man:
            f_man.write("=" * 72 + "\n")
            f_man.write("FICHA TÉCNICA N.° 32 - TELEMEDICINA (MINSA / DITEL)\n")
            f_man.write("MANIFIESTO DE INFORMES GENERADOS\n")
            f_man.write("=" * 72 + "\n\n")
            f_man.write(f"Fecha y hora de generación:  {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f_man.write(f"Dirección Regional (DIRESA): {diresa}\n")
            f_man.write(f"Año del Convenio:            {anio}\n")
            f_man.write(f"Mes de corte de evaluación:  Mes {mes} ({M.MES_TXT.get(int(mes), '')})\n")
            f_man.write(f"Alcance seleccionado:        {slug_alcance.replace('_', ' ')}\n")
            f_man.write(f"Establecimientos evaluados:  {len(A)} primer nivel, {len(B)} hospitales\n\n")
            f_man.write("Archivos entregables en esta carpeta:\n")
            for s in salidas:
                rel = os.path.relpath(s, destino)
                f_man.write(f"  • {rel}\n")
            f_man.write("\n" + "=" * 72 + "\n")
    except Exception:
        pass

    out(f"[info] Listo en {time.time() - inicio:.1f} s → {destino}")
    return salidas


def importlib_reload(mod):
    """Recarga un módulo para que sus constantes (rutas, etiquetas) se actualicen."""
    try:
        import importlib
        importlib.reload(mod)
    except Exception:
        pass


def main(argv=None):
    p = argparse.ArgumentParser(description="Genera los informes de la Ficha N.° 32.")
    p.add_argument("--diresa", default="PUNO", help="DIRESA (p. ej. PUNO, AREQUIPA)")
    p.add_argument("--mes", type=int, default=M.MES_CORTE, help="Mes de corte (1-12)")
    p.add_argument("--anio", type=int, default=M.ANIO, help="Año del convenio")
    p.add_argument("--excel", action="store_true", help="Generar el libro Excel")
    p.add_argument("--word", action="store_true", help="Generar los informes Word por red")
    p.add_argument("--graficos", action="store_true", help="Generar los gráficos PNG")
    p.add_argument("--evolucion", action="store_true", help="Incluir la hoja Evolución")
    p.add_argument("--evo-desde", type=int, default=1,
                   help="Primer mes de la hoja Evolución (no hace falta que sea el mes 1)")
    p.add_argument("--evo-hasta", type=int, default=None,
                   help="Último mes de la hoja Evolución (por defecto: el mes de corte)")
    p.add_argument("--sin-plan", action="store_true",
                   help="Omitir la parte de planificación (hojas Plan, Palanca y Supuestos)")
    p.add_argument("--plan-desde", type=int, default=None,
                   help="Mes de inicio del Plan (ej. 11 para Noviembre o 7 para Julio)")
    p.add_argument("--plan-hasta", type=int, default=None,
                   help="Mes de fin del Plan (ej. 12 para Diciembre)")
    p.add_argument("--sin-hospitales", action="store_true", help="Omitir la hoja Hospitales")
    p.add_argument("--sin-redes", action="store_true",
                   help="Omitir la hoja Redes y las hojas por red (Red_*)")
    p.add_argument("--sin-cartera", action="store_true",
                   help="Omitir la hoja Cartera Servicios (SUSALUD RENIPRESS)")
    p.add_argument("--alcance", action="append", default=None,
                   help="red:NOMBRE o ipress:COD1,COD2 (se puede repetir)")
    p.add_argument("--carpeta", default=None, help="Carpeta de salida")
    p.add_argument("--refrescar", action="store_true", help="Ignorar la caché y volver a descargar")
    a = p.parse_args(argv)
    if not (a.excel or a.word or a.graficos):
        a.excel = True
    generar(a.diresa, a.mes, a.anio, excel=a.excel, word=a.word, graficos=a.graficos,
            evolucion=a.evolucion, alcance=a.alcance, carpeta=a.carpeta, refrescar=a.refrescar,
            plan=not a.sin_plan, hospitales=not a.sin_hospitales, redes=not a.sin_redes,
            evo_desde=a.evo_desde, evo_hasta=a.evo_hasta,
            plan_desde=a.plan_desde, plan_hasta=a.plan_hasta,
            cartera=not a.sin_cartera)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        out(f"[error] {e}")
        sys.exit(1)
