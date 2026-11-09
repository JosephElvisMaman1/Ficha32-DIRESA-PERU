# -*- coding: utf-8 -*-
"""Gráficos de la Ficha N.° 32 (periodo según ficha32_modelo: M.ETQ / M.DIRESA)."""
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import ficha32_modelo as M

OUT = os.path.join(M.BASE, "graficos")
os.makedirs(OUT, exist_ok=True)

C_V, C_A, C_R = "#86efac", "#fde047", "#fca5a5"
E_V, E_A, E_R = "#16a34a", "#ca8a04", "#dc2626"
COL_SEM = {"VERDE": C_V, "AMARILLO": C_A, "ROJO": C_R}
EDG_SEM = {"VERDE": E_V, "AMARILLO": E_A, "ROJO": E_R}
AZUL, GRIS, OSC = "#3b82f6", "#94a3b8", "#1f2937"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "axes.spines.top": False, "axes.spines.right": False,
    "axes.titleweight": "bold", "axes.titlesize": 11, "axes.labelsize": 9.5,
})


def _fmt(n):
    return f"{int(round(n)):,}".replace(",", ".")


def _guardar(fig, slug, nombre):
    path = os.path.join(OUT, f"{slug.lower()}_{nombre}.png")
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)
    return path


def conteo(df, col):
    return [int((df[col] == k).sum()) for k in ("VERDE", "AMARILLO", "ROJO")]


def fig_semaforo_comparativo(df, titulo, slug):
    fig, ax = plt.subplots(figsize=(7, 3.4))
    x = np.arange(3)
    w = 0.36
    c7, c8 = conteo(df, "SEMAFORO_ANT"), conteo(df, "SEMAFORO")
    etq_ant = f"Mes {M.MES_CORTE - 1}" if M.MES_CORTE > 1 else "Línea base"
    b1 = ax.bar(x - w / 2, c7, w, color=["#d1fae5", "#fef9c3", "#fee2e2"], edgecolor=[E_V, E_A, E_R], hatch="//",
                label=etq_ant)
    b2 = ax.bar(x + w / 2, c8, w, color=[C_V, C_A, C_R], edgecolor=[E_V, E_A, E_R], label=M.ETQ['MES'])
    for bars in (b1, b2):
        for b in bars:
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() + max(c7 + c8 + [1]) * 0.02, str(int(b.get_height())),
                    ha="center", fontsize=9, fontweight="bold", color=OSC)
    ax.set_xticks(x)
    ax.set_xticklabels(["Verde\n(≥60 pts)", "Amarillo\n(30–59.9 pts)", "Rojo\n(<30 pts)"])
    ax.set_ylabel(f"N.° de IPRESS (total = {len(df)})")
    ax.set_ylim(0, max(c7 + c8 + [1]) * 1.18)
    ax.grid(axis="y", ls="--", alpha=0.4)
    ax.set_title(f"Semáforo de gestión: {etq_ant} vs {M.ETQ['MES']}\n{titulo}", pad=8, fontsize=10.5)
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(facecolor="#e5e7eb", edgecolor="#6b7280", hatch="//", label=etq_ant),
                       Patch(facecolor="#9ca3af", edgecolor="#6b7280", label=M.ETQ['MES'])], frameon=False, fontsize=8.5)
    return _guardar(fig, slug, "semaforo")


def fig_avance_servicios(df, titulo, slug):
    fig, ax = plt.subplots(figsize=(7, 3.4))
    esperado = M.MES_CORTE / M.MESES_AÑO * 100
    vals, rot = [], []
    for s in M.SERV:
        r, m = df[f"{s} REALIZADAS"].sum(), df[f"{s} META"].sum()
        vals.append(r / m * 100 if m else 0)
        rot.append(f"{_fmt(r)} de {_fmt(m)}")
    colores = ["#60a5fa", "#a78bfa", "#34d399"]
    tope = 150
    shown = [min(v, tope) for v in vals]
    bars = ax.barh([f"{M.NOMBRE_SERV[s]} ({s})\npeso {M.PESO[s]}" for s in M.SERV], shown, color=colores,
                   edgecolor="#374151", height=0.55)
    l1 = ax.axvline(esperado, color=E_R, ls="--", lw=1.6, label=f"Avance lineal esperado a {M.ETQ['MES_LARGO'].lower()} ({esperado:.1f} %)")
    l2 = ax.axvline(100, color=OSC, ls=":", lw=1.3, label="Meta anual (100 %)")
    for b, v, t in zip(bars, vals, rot):
        marca = " ▶" if v > tope else ""
        ax.text(min(v, tope) + 3, b.get_y() + b.get_height() / 2, f"{v:.1f} %{marca}\n({t})", va="center",
                fontsize=8.3, fontweight="bold", color=OSC)
    ax.set_xlim(0, 250)
    ax.invert_yaxis()
    ax.set_xlabel(f"% de la meta anual acumulada al {M.ETQ['MES']} (barra recortada en 150 %)")
    ax.legend(handles=[l1, l2], loc="upper center", bbox_to_anchor=(0.5, -0.27), ncol=2, frameon=False, fontsize=8)
    ax.grid(axis="x", ls="--", alpha=0.4)
    ax.set_title(f"Avance acumulado vs meta anual por servicio\n{titulo}", pad=8, fontsize=10.5)
    return _guardar(fig, slug, "avance")


def fig_ranking(df, titulo, slug, n=15):
    top = df.sort_values(["PUNTAJE", "PUNTAJE_ANT"], ascending=False).head(n).iloc[::-1]
    fig, ax = plt.subplots(figsize=(7, max(3.0, 0.36 * len(top) + 1.3)))
    y = np.arange(len(top))
    ax.barh(y, top["PUNTAJE"], color=[COL_SEM[s] for s in top["SEMAFORO"]],
            edgecolor=[EDG_SEM[s] for s in top["SEMAFORO"]], height=0.65)
    ax.scatter(top["PUNTAJE_ANT"], y, marker="D", s=26, color=OSC, zorder=3, label=f"Puntaje Mes {M.MES_CORTE - 1}")
    for yi, v in zip(y, top["PUNTAJE"]):
        ax.text(min(v, 100) + 1.2, yi, f"{v:.1f}", va="center", fontsize=8, fontweight="bold", color=OSC)
    ax.set_yticks(y)
    ax.set_yticklabels([n_[:30] for n_ in top["NOMBRE"]], fontsize=8)
    ax.axvline(60, color=E_V, ls="--", lw=1)
    ax.axvline(30, color=E_A, ls="--", lw=1)
    ax.set_xlim(0, 112)
    ax.set_xlabel("Puntaje de Logro (0–100)")
    ax.grid(axis="x", ls="--", alpha=0.4)
    ax.legend(loc="lower right", frameon=False, fontsize=8)
    ax.set_title(f"Ranking de IPRESS por Puntaje de Logro ({M.ETQ['MES']})\n{titulo}", pad=8, fontsize=10.5)
    return _guardar(fig, slug, "ranking")


def get_escenarios():
    etq_ant = f"Mes {M.MES_CORTE - 1}\n(oficial)" if M.MES_CORTE > 1 else "Línea base\n(oficial)"
    return [
        (etq_ant, "PUNTAJE_ANT", GRIS),
        (f"{M.ETQ['MES']}\n(oficial)", "PUNTAJE", "#475569"),
        (f"Proy. {M.ETQ['PROY']}\n(ritmo {M.ETQ['RANGO'].lower()})", "PUNTAJE_PROY_OCT", "#93c5fd"),
        (f"{M.MES_ABR[M.MESES_AÑO]}: sigue\nla tendencia", "PUNTAJE_DIC_TEND", "#fbbf24"),
        (f"{M.MES_ABR[M.MESES_AÑO]}: Plan\nbase", "PUNTAJE_PLAN_BASE", "#34d399"),
        (f"{M.MES_ABR[M.MESES_AÑO]}: Plan\nampliado", "PUNTAJE_PLAN", "#059669"),
    ]

ESCENARIOS = get_escenarios()


def fig_proyeccion(df, titulo, slug):
    fig, ax = plt.subplots(figsize=(7.4, 3.6))
    escenarios = get_escenarios()
    vals = [df[c].mean() for _, c, _ in escenarios]
    bars = ax.bar(range(len(vals)), vals, color=[c for *_, c in escenarios], edgecolor="#374151", width=0.62)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + 1.5, f"{v:.1f}", ha="center", fontsize=9.5, fontweight="bold",
                color=OSC)
    ax.axhline(60, color=E_V, ls="--", lw=1.2)
    ax.text(-0.45, 61.5, "Umbral Verde (60)", ha="left", fontsize=8, color=E_V)
    ax.axhline(100, color=OSC, ls=":", lw=1.2)
    ax.text(-0.45, 101.5, "Logro esperado (100)", ha="left", fontsize=8, color=OSC)
    ax.set_xticks(range(len(vals)))
    ax.set_xticklabels([e[0] for e in escenarios], fontsize=8)
    ax.set_ylim(0, 112)
    ax.set_ylabel("Puntaje de Logro promedio")
    ax.grid(axis="y", ls="--", alpha=0.4)
    ax.set_title(f"Proyección del puntaje promedio a {M.MES_TXT[M.MESES_AÑO].lower()}\n{titulo}", pad=8, fontsize=10.5)
    return _guardar(fig, slug, "proyeccion")


def fig_semaforo_proyectado(df, titulo, slug):
    cols = [(M.ETQ['MES'], "SEMAFORO"), (f"Proy. {M.ETQ['PROY']}", "SEMAFORO_OCT"),
            (f"{M.MES_ABR[M.MESES_AÑO]}: tendencia", None),
            (f"{M.MES_ABR[M.MESES_AÑO]}: Plan base", "SEMAFORO_PLAN_BASE"),
            (f"{M.MES_ABR[M.MESES_AÑO]}: Plan ampliado", "SEMAFORO_PLAN")]
    df = df.copy()
    df["SEM_DIC_TEND"] = df["PUNTAJE_DIC_TEND"].apply(M.clasificar)
    cols[2] = (f"{M.MES_ABR[M.MESES_AÑO]}: tendencia", "SEM_DIC_TEND")
    data = np.array([conteo(df, c) for _, c in cols]) / max(len(df), 1) * 100
    fig, ax = plt.subplots(figsize=(7.4, 3.3))
    left = np.zeros(len(cols))
    for j, (lab, col, ed) in enumerate((("Verde", C_V, E_V), ("Amarillo", C_A, E_A), ("Rojo", C_R, E_R))):
        ax.barh(range(len(cols)), data[:, j], left=left, color=col, edgecolor=ed, label=lab, height=0.6)
        for i, v in enumerate(data[:, j]):
            if v >= 6:
                ax.text(left[i] + v / 2, i, f"{v:.0f}%", ha="center", va="center", fontsize=8.5, fontweight="bold")
        left += data[:, j]
    ax.set_yticks(range(len(cols)))
    ax.set_yticklabels([c[0] for c in cols], fontsize=8.5)
    ax.invert_yaxis()
    ax.set_xlim(0, 100)
    ax.set_xlabel("% de IPRESS")
    ax.legend(ncol=3, frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.2), fontsize=8.5)
    ax.set_title(f"Composición del semáforo proyectado\n{titulo}", pad=8, fontsize=10.5)
    return _guardar(fig, slug, "semaforo_proy")


def fig_categorias(df, titulo, slug):
    cats = M.CATEGORIAS
    colores = ["#86efac", "#bef264", "#fde047", "#fca5a5"]
    cnt = [int((df["CATEGORIA"] == c).sum()) for c in cats]
    req = [int(df.loc[df["CATEGORIA"] == c, "OBJ_MES_TOTAL"].sum()) for c in cats]
    fig, ax = plt.subplots(figsize=(7, 3.2))
    bars = ax.barh(cats, cnt, color=colores, edgecolor="#374151", height=0.6)
    top = max(cnt + [1])
    for b, c, r in zip(bars, cnt, req):
        txt = f"{c} IPRESS" + (f" · {_fmt(r)} " + ("atención" if r == 1 else "atenciones") + f"/mes en {M.ETQ['PLAN']}" if c else "")
        ax.text(b.get_width() + top * 0.02, b.get_y() + b.get_height() / 2, txt, va="center", fontsize=8.5,
                fontweight="bold", color=OSC)
    ax.invert_yaxis()
    ax.set_xlim(0, top * 2.0)
    ax.set_xlabel("N.° de IPRESS")
    ax.grid(axis="x", ls="--", alpha=0.4)
    ax.set_title(f"IPRESS según esfuerzo requerido para cumplir la meta\n{titulo}", pad=8, fontsize=10.5)
    return _guardar(fig, slug, "categorias")


def fig_palanca(df, titulo, slug, n=12):
    """IPRESS que más puntaje ganan con el plan base (proy. del mes proyectado -> plan base de cierre)."""
    d = df[df["CATEGORIA"] != "CUMPLE"].copy()
    d["GANA"] = d["PUNTAJE_PLAN_BASE"] - d["PUNTAJE_PROY_OCT"]
    d = d.sort_values("GANA", ascending=False).head(n).iloc[::-1]
    if d.empty:
        return None
    fig, ax = plt.subplots(figsize=(7, max(3.0, 0.38 * len(d) + 1.3)))
    y = np.arange(len(d))
    ax.barh(y, d["PUNTAJE_PROY_OCT"], color="#cbd5e1", edgecolor="#64748b", height=0.62, label=f"Proy. fin de {M.ETQ['HASTA'].lower()}")
    ax.barh(y, d["GANA"], left=d["PUNTAJE_PROY_OCT"], color="#34d399", edgecolor="#059669", height=0.62,
            label=f"Aporte del plan {M.ETQ['PLAN']}")
    for yi, a, g in zip(y, d["PUNTAJE_PROY_OCT"], d["GANA"]):
        ax.text(a + g + 1, yi, f"{a + g:.1f}  (+{g:.1f})", va="center", fontsize=8, fontweight="bold", color=OSC)
    ax.set_yticks(y)
    ax.set_yticklabels([n_[:30] for n_ in d["NOMBRE"]], fontsize=8)
    ax.axvline(60, color=E_V, ls="--", lw=1)
    ax.set_xlim(0, 120)
    ax.set_xlabel("Puntaje de Logro")
    ax.grid(axis="x", ls="--", alpha=0.4)
    ax.legend(loc="lower right", frameon=False, fontsize=8)
    ax.set_title(f"IPRESS con mayor ganancia de puntaje con el plan\n{titulo}", pad=8, fontsize=10.5)
    return _guardar(fig, slug, "palanca")


def fig_redes(A, slug="REGIONAL"):
    """Comparativo regional por red: puntaje promedio del mes anterior, mes de corte y plan base."""
    redes = M.redes(A)
    g = A.groupby("RED_SALUD")[["PUNTAJE_ANT", "PUNTAJE", "PUNTAJE_PLAN_BASE"]].mean().sort_values("PUNTAJE")
    fig, ax = plt.subplots(figsize=(8, 5.2))
    y = np.arange(len(g))
    h = 0.27
    ax.barh(y - h, g["PUNTAJE_ANT"], h, color="#cbd5e1", edgecolor="#64748b", label=f"Mes {M.MES_CORTE - 1}")
    ax.barh(y, g["PUNTAJE"], h, color="#60a5fa", edgecolor="#1d4ed8", label=M.ETQ['MES'])
    ax.barh(y + h, g["PUNTAJE_PLAN_BASE"], h, color="#34d399", edgecolor="#059669", label=f"{M.MES_ABR[M.MESES_AÑO]}: Plan base")
    for yi, v in zip(y, g["PUNTAJE"]):
        ax.text(v + 1, yi, f"{v:.1f}", va="center", fontsize=8, fontweight="bold")
    ax.set_yticks(y)
    ax.set_yticklabels([dict([(p, d) for p, _, d in redes]).get(k, k) for k in g.index], fontsize=8.5)
    ax.axvline(60, color=E_V, ls="--", lw=1)
    ax.set_xlabel("Puntaje de Logro promedio")
    ax.grid(axis="x", ls="--", alpha=0.4)
    ax.legend(frameon=False, fontsize=8.5, loc="lower right")
    ax.set_title(f"Puntaje promedio por Red de Salud — DIRESA {M.DIRESA.title()} ({len(redes)} Redes)", pad=10)
    return _guardar(fig, slug, "redes")


def graficos_red(df, titulo, slug):
    p = {
        "semaforo": fig_semaforo_comparativo(df, titulo, slug),
        "avance": fig_avance_servicios(df, titulo, slug),
        "ranking": fig_ranking(df, titulo, slug),
        "proyeccion": fig_proyeccion(df, titulo, slug),
        "semaforo_proy": fig_semaforo_proyectado(df, titulo, slug),
        "categorias": fig_categorias(df, titulo, slug),
        "palanca": fig_palanca(df, titulo, slug),
    }
    return p


if __name__ == "__main__":
    A, B = M.construir()
    redes = M.redes(A)
    for prov, slug, disp in redes:
        graficos_red(A[A["RED_SALUD"] == prov], f"Red de Salud {disp}", slug)
    graficos_red(A, f"DIRESA {M.DIRESA.title()} ({len(A)} IPRESS)", "REGIONAL")
    fig_redes(A)
    print("OK", len(os.listdir(OUT)), "gráficos en", OUT)


