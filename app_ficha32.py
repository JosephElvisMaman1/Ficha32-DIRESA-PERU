# -*- coding: utf-8 -*-
"""Aplicación gráfica — Generador de informes de la Ficha N.° 32 (Telemedicina).

Interfaz para uso administrativo: elige la DIRESA, el mes y qué se quiere generar; el
botón "Generar informes" hace el resto (descarga, cálculo y archivos). No requiere
conocimientos técnicos.

Arranque:  python app_ficha32.py     (o doble clic en Iniciar_Ficha32.bat)
"""
import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Con pythonw (doble clic) no hay consola y sys.stdout es None; se protege y se
# dejan los errores de arranque en un archivo para poder diagnosticarlos.
def _seguro(fn):
    try:
        fn()
    except Exception:
        pass


_seguro(lambda: sys.stdout.reconfigure(encoding="utf-8", errors="replace"))

import config_manager as CM        # noqa: E402
import ficha32_api as api          # noqa: E402
import ficha32_modelo as M         # noqa: E402
import generar_todo                # noqa: E402

PYTHON = sys.executable
CARPETA = CM.get_app_dir()
LOG_ERRORES = os.path.join(CARPETA, "errores_app.log")
MESES_TXT = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
             7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
MARCAS = ["#573375", "#006b5f", "#1f3e8f"]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ficha N.° 32 · Telemedicina — Generador de informes")
        self.geometry("1000x750")
        self.minsize(860, 520)
        self.configure(bg="#F8F9FA")
        self._proceso = None
        self._generando = False
        self._redes = []
        self._ipress = []       # [(codigo, nombre, red)] de la DIRESA consultada
        self._marcados = set()  # códigos con la casilla marcada
        self._ultima_carpeta_salida = None
        self.protocol("WM_DELETE_WINDOW", self._al_cerrar)

        self._estilo()
        self._cabecera()
        self._crear_contenedor_desplazable()
        self._panel_datos(self.cuerpo)
        self._panel_alcance(self.cuerpo)
        self._panel_salidas(self.cuerpo)
        self._panel_carpeta(self.cuerpo)
        self._panel_accion(self.cuerpo)
        self._panel_log(self.cuerpo)
        self._pie(self.cuerpo)

        self._id_prono = self.after(200, self._cargar_pronostico)

    # ------------------------------------------------------------------ estilo
    def _estilo(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure(".", font=("Segoe UI", 10))
        s.configure("TFrame", background="#F8F9FA")
        s.configure("TLabel", background="#F8F9FA", foreground="#191C1D")
        s.configure("TLabelframe", background="#F8F9FA", foreground="#573375",
                    font=("Segoe UI", 10, "bold"))
        s.configure("TLabelframe.Label", background="#F8F9FA", foreground="#573375",
                    font=("Segoe UI", 10, "bold"))
        s.configure("TButton", font=("Segoe UI", 10))
        s.configure("Big.TButton", font=("Segoe UI", 13, "bold"), padding=10)
        s.configure("TCheckbutton", background="#F8F9FA", font=("Segoe UI", 10))
        s.configure("TRadiobutton", background="#F8F9FA", font=("Segoe UI", 10))

    # ------------------------------------------------------------------ bloques
    def _cabecera(self):
        f = tk.Frame(self, bg="#573375", pady=10, padx=14)
        f.pack(fill="x")
        top_bar = tk.Frame(f, bg="#573375")
        top_bar.pack(fill="x")
        tk.Label(top_bar, text="MINSA · DITEL", font=("Segoe UI", 11, "bold"),
                 bg="#573375", fg="#FFFFFF").pack(side="left")
        btn_cfg = tk.Button(top_bar, text="⚙ Ajustes API / Fórmulas", font=("Segoe UI", 9, "bold"),
                            bg="#7C3AED", fg="#FFFFFF", activebackground="#6D28D9", activeforeground="#FFFFFF",
                            relief="flat", padx=8, pady=2, cursor="hand2", command=self._abrir_dialogo_config)
        btn_cfg.pack(side="right")
        tk.Label(f, text="Indicador de Desempeño en Telemedicina — Ficha N.° 32",
                 font=("Segoe UI", 15, "bold"), bg="#573375", fg="#FFFFFF").pack(anchor="w")
        tk.Label(f, text="Convenios de Gestión · Tasas de uso de los servicios de telemedicina",
                 font=("Segoe UI", 9), bg="#573375", fg="#E8C7FF").pack(anchor="w")

    def _crear_contenedor_desplazable(self):
        """Crea un área deslizable (Canvas + Scrollbar) para que quepan todos los paneles en cualquier monitor."""
        contenedor = ttk.Frame(self)
        contenedor.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(contenedor, bg="#F8F9FA", highlightthickness=0)
        self.sb_v = ttk.Scrollbar(contenedor, orient="vertical", command=self.canvas.yview)
        self.cuerpo = ttk.Frame(self.canvas)

        self._win_cuerpo = self.canvas.create_window((0, 0), window=self.cuerpo, anchor="nw")

        def _on_cuerpo_config(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def _on_canvas_config(event):
            # Mantener el ancho del contenido alineado al canvas para que expanda correctamente
            self.canvas.itemconfig(self._win_cuerpo, width=event.width)

        self.cuerpo.bind("<Configure>", _on_cuerpo_config)
        self.canvas.bind("<Configure>", _on_canvas_config)
        self.canvas.configure(yscrollcommand=self.sb_v.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.sb_v.pack(side="right", fill="y")

        # Scroll suave con la rueda del ratón
        self.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        """Maneja el desplazamiento de rueda del ratón en Windows de forma inteligente."""
        try:
            w = event.widget
            # Si el cursor está sobre el árbol de IPRESS o sobre el cuadro de Log, desplazar ese widget
            if hasattr(self, "tree") and (w == self.tree or str(w).startswith(str(self.tree))):
                self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
                return
            if hasattr(self, "log") and (w == self.log or str(w).startswith(str(self.log))):
                self.log.yview_scroll(int(-1 * (event.delta / 120)), "units")
                return
        except Exception:
            pass
        if hasattr(self, "canvas"):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _panel_datos(self, parent=None):
        p = parent or self
        f = ttk.LabelFrame(p, text=" 1 · Datos a reportar ")
        f.pack(fill="x", padx=14, pady=(12, 6))
        g = ttk.Frame(f)
        g.pack(fill="x", padx=10, pady=8)
        g.columnconfigure((1, 3, 5), weight=1)

        ttk.Label(g, text="Dirección Regional:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.v_diresa = tk.StringVar(value="PUNO")
        self.cb_diresa = ttk.Combobox(g, textvariable=self.v_diresa, state="readonly", width=26,
                                      values=[d for _, d in api.DIRESAS])
        self.cb_diresa.grid(row=0, column=1, sticky="ew", padx=(0, 16))
        self.cb_diresa.bind("<<ComboboxSelected>>", lambda e: self._cargar_pronostico())

        ttk.Label(g, text="Año del convenio:").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.v_anio = tk.IntVar(value=M.ANIO)
        sp = ttk.Spinbox(g, from_=2024, to=2035, width=8, textvariable=self.v_anio,
                         command=self._marcar_cambio)
        sp.grid(row=0, column=3, sticky="w", padx=(0, 16))
        sp.bind("<KeyRelease>", lambda e: self._marcar_cambio())

        ttk.Label(g, text="Mes de corte:").grid(row=0, column=4, sticky="w", padx=(0, 6))
        self.v_mes = tk.StringVar(value="Mes 8 (Agosto)")
        self.cb_mes = ttk.Combobox(g, textvariable=self.v_mes, state="readonly", width=24)
        self.cb_mes.grid(row=0, column=5, sticky="ew")
        self.cb_mes.bind("<<ComboboxSelected>>", lambda e: (self._marcar_cambio(), self._mes_cambiado()))

        op = ttk.Frame(f)
        op.pack(fill="x", padx=10, pady=(0, 8))
        self.v_refrescar = tk.BooleanVar(value=False)
        ttk.Checkbutton(op, text="Volver a descargar los datos (ignorar la copia guardada)",
                        variable=self.v_refrescar).pack(anchor="w")
        self.v_estado = ttk.Label(op, text="Consultando el tablero…", foreground="#64748B")
        self.v_estado.pack(anchor="w", pady=(2, 0))

    def _panel_alcance(self, parent=None):
        p = parent or self
        f = ttk.LabelFrame(p, text=" 2 · ¿Qué establecimientos (eess) incluir? ")
        f.pack(fill="x", padx=14, pady=6)

        g = ttk.Frame(f)
        g.pack(fill="x", padx=10, pady=(8, 4))
        self.v_alcance = tk.StringVar(value="todo")
        ttk.Radiobutton(g, text="Toda la DIRESA", value="todo", variable=self.v_alcance,
                        command=self._cambiar_alcance).pack(side="left")
        ttk.Radiobutton(g, text="Elegir eess con casillas", value="eess", variable=self.v_alcance,
                        command=self._cambiar_alcance).pack(side="left", padx=(14, 0))
        self.lb_alcance = ttk.Label(g, text="", foreground="#64748B")
        self.lb_alcance.pack(side="right")

        # Barra de filtros y acciones sobre las casillas
        g2 = ttk.Frame(f)
        g2.pack(fill="x", padx=10, pady=2)
        ttk.Label(g2, text="Red de salud:").pack(side="left")
        self.v_red = tk.StringVar(value="(todas)")
        self.cb_red = ttk.Combobox(g2, textvariable=self.v_red, state="disabled", width=24,
                                   values=["(todas)"])
        self.cb_red.pack(side="left", padx=(4, 10))
        self.cb_red.bind("<<ComboboxSelected>>", lambda e: self._refrescar_arbol())
        ttk.Label(g2, text="Categoría:").pack(side="left")
        self.v_nivel = tk.StringVar(value="(todas)")
        self.cb_nivel = ttk.Combobox(g2, textvariable=self.v_nivel, state="disabled", width=9,
                                     values=["(todas)"])
        self.cb_nivel.pack(side="left", padx=(4, 10))
        self.cb_nivel.bind("<<ComboboxSelected>>", lambda e: self._refrescar_arbol())
        ttk.Label(g2, text="Buscar:").pack(side="left")
        self.v_buscar = tk.StringVar()
        ent = ttk.Entry(g2, textvariable=self.v_buscar, width=26)
        ent.pack(side="left", padx=(4, 12))
        ent.bind("<KeyRelease>", lambda e: self._refrescar_arbol())
        self.bt_marcar = ttk.Button(g2, text="Marcar visibles", state="disabled",
                                    command=lambda: self._marcar_visibles(True))
        self.bt_marcar.pack(side="left", padx=(0, 4))
        self.bt_desmarcar = ttk.Button(g2, text="Desmarcar visibles", state="disabled",
                                       command=lambda: self._marcar_visibles(False))
        self.bt_desmarcar.pack(side="left")

        # Lista de eess con casillas (☑ / ☐)
        ft = ttk.Frame(f)
        ft.pack(fill="x", padx=10, pady=(4, 2))
        self.tree = ttk.Treeview(ft, columns=("sel", "red", "cat_f32", "cat_su", "estado_cat", "cod", "nom",
                                              "res_cat", "fec_cat", "sub_cat",
                                              "fec_car", "sub_car"),
                                 show="headings", height=7, selectmode="none")
        self.tree.heading("sel", text="✓")
        self.tree.heading("red", text="Red de salud")
        self.tree.heading("cat_f32", text="Cat. F32")
        self.tree.heading("cat_su", text="Cat. SUSALUD")
        self.tree.heading("estado_cat", text="¿Cambió Cat.?")
        self.tree.heading("cod", text="Código")
        self.tree.heading("nom", text="Establecimiento / IPRESS")
        self.tree.heading("res_cat", text="N° Res. Categorización")
        self.tree.heading("fec_cat", text="Fecha Categorización")
        self.tree.heading("sub_cat", text="¿Subió Categorización?")
        self.tree.heading("fec_car", text="Fecha Res. Cartera")
        self.tree.heading("sub_car", text="¿Subió Cartera?")

        self.tree.column("sel", width=28, anchor="center", minwidth=28)
        self.tree.column("red", width=105, anchor="w", minwidth=85)
        self.tree.column("cat_f32", width=55, anchor="center", minwidth=45)
        self.tree.column("cat_su", width=75, anchor="center", minwidth=55)
        self.tree.column("estado_cat", width=95, anchor="center", minwidth=80)
        self.tree.column("cod", width=68, anchor="center", minwidth=60)
        self.tree.column("nom", width=190, anchor="w", minwidth=140)
        self.tree.column("res_cat", width=180, anchor="w", minwidth=130)
        self.tree.column("fec_cat", width=110, anchor="center", minwidth=90)
        self.tree.column("sub_cat", width=125, anchor="center", minwidth=95)
        self.tree.column("fec_car", width=110, anchor="center", minwidth=90)
        self.tree.column("sub_car", width=90, anchor="center", minwidth=75)

        self.tree.tag_configure("vario", foreground="#B45309", background="#FEF3C7")
        self.tree.tag_configure("igual", foreground="#15803D")

        sb_y = ttk.Scrollbar(ft, orient="vertical", command=self.tree.yview)
        sb_x = ttk.Scrollbar(ft, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb_y.grid(row=0, column=1, sticky="ns")
        sb_x.grid(row=1, column=0, sticky="ew")
        ft.grid_rowconfigure(0, weight=1)
        ft.grid_columnconfigure(0, weight=1)
        self.tree.bind("<Button-1>", self._click_arbol)

        pie = ttk.Frame(f)
        pie.pack(fill="x", padx=10, pady=(0, 8))
        self.lb_cuenta = ttk.Label(pie, text="Consultando el tablero para cargar los eess…",
                                   foreground="#64748B")
        self.lb_cuenta.pack(side="left")
        ttk.Label(pie, text="Consejo: elija la red para ver solo sus eess y marcar en bloque.",
                  foreground="#94A3B8", font=("Segoe UI", 8)).pack(side="right")

    def _panel_salidas(self, parent=None):
        p = parent or self
        f = ttk.LabelFrame(p, text=" 3 · Qué generar ")
        f.pack(fill="x", padx=14, pady=6)
        g = ttk.Frame(f)
        g.pack(fill="x", padx=10, pady=(8, 4))
        g.columnconfigure((0, 1), weight=1)
        self.v_excel = tk.BooleanVar(value=True)
        self.v_plan = tk.BooleanVar(value=True)
        self.v_hosp = tk.BooleanVar(value=True)
        self.v_redes = tk.BooleanVar(value=True)
        self.v_word = tk.BooleanVar(value=False)
        self.v_graf = tk.BooleanVar(value=False)
        self.v_cartera = tk.BooleanVar(value=True)
        ttk.Checkbutton(g, text="Libro Excel (Resumen, IPRESS y datos actuales)",
                        variable=self.v_excel, command=self._alternar_evo).grid(row=0, column=0, sticky="w")
        self.cb_plan_bt = ttk.Checkbutton(g, text="Plan de Cumplimiento (hojas Plan, Palanca y Supuestos)",
                                           variable=self.v_plan, command=self._alternar_plan)
        self.cb_plan_bt.grid(row=0, column=1, sticky="w", padx=(18, 0))
        ttk.Checkbutton(g, text="Hospitales (hoja Hospitales)",
                        variable=self.v_hosp).grid(row=1, column=0, sticky="w")
        self.cb_redes_bt = ttk.Checkbutton(g, text="Parte de redes (hoja Redes + una hoja por red)",
                                           variable=self.v_redes)
        self.cb_redes_bt.grid(row=1, column=1, sticky="w", padx=(18, 0))
        ttk.Checkbutton(g, text="Informes Word (Informe Ejecutivo Regional + por Red)",
                        variable=self.v_word).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(g, text="Gráficos PNG por red de salud",
                        variable=self.v_graf).grid(row=2, column=1, sticky="w", padx=(18, 0))
        ttk.Checkbutton(g, text="Hoja «Cartera Servicios» (SUSALUD / RENIPRESS)",
                        variable=self.v_cartera).grid(row=3, column=0, sticky="w")

        # Rango del Plan de Cumplimiento (configurable)
        import config_manager as CM
        _cplan = CM.get_plan_config()
        gp = ttk.Frame(f)
        gp.pack(fill="x", padx=10, pady=(2, 4))
        ttk.Label(gp, text="Rango del Plan:", font=("Segoe UI", 9, "bold")).pack(side="left")
        ttk.Label(gp, text="de Mes").pack(side="left", padx=(8, 4))
        self.v_plan_desde = tk.IntVar(value=_cplan.get("mes_inicio_plan", 11))
        self.v_plan_hasta = tk.IntVar(value=_cplan.get("mes_fin_plan", 12))
        self.sp_plan_desde = ttk.Spinbox(gp, from_=1, to=12, width=4, textvariable=self.v_plan_desde,
                                         command=self._plan_cambiado)
        self.sp_plan_desde.pack(side="left", padx=(0, 6))
        self.sp_plan_desde.bind("<KeyRelease>", lambda e: self._plan_cambiado())
        ttk.Label(gp, text="a Mes").pack(side="left", padx=(4, 4))
        self.sp_plan_hasta = ttk.Spinbox(gp, from_=1, to=12, width=4, textvariable=self.v_plan_hasta,
                                         command=self._plan_cambiado)
        self.sp_plan_hasta.pack(side="left", padx=(0, 8))
        self.sp_plan_hasta.bind("<KeyRelease>", lambda e: self._plan_cambiado())
        self.lb_plan_txt = ttk.Label(gp, text="", foreground="#573375", font=("Segoe UI", 9, "bold"))
        self.lb_plan_txt.pack(side="left")

        # Hoja Evolución: rango de meses seleccionable (mes X → mes Z, no solo desde el mes 1)
        ge = ttk.Frame(f)
        ge.pack(fill="x", padx=10, pady=(0, 8))
        self.v_evo = tk.BooleanVar(value=True)
        ttk.Checkbutton(ge, text="Hoja «Evolución»: de Mes", variable=self.v_evo,
                        command=self._alternar_evo).pack(side="left")
        self.v_evo_desde = tk.IntVar(value=1)
        self.v_evo_hasta = tk.IntVar(value=8)
        self.sp_desde = ttk.Spinbox(ge, from_=1, to=12, width=4, textvariable=self.v_evo_desde,
                                    state="disabled")
        self.sp_desde.pack(side="left", padx=(4, 6))
        ttk.Label(ge, text="a Mes").pack(side="left")
        self.sp_hasta = ttk.Spinbox(ge, from_=1, to=12, width=4, textvariable=self.v_evo_hasta,
                                    state="disabled")
        self.sp_hasta.pack(side="left", padx=(6, 0))
        self.lb_evo = ttk.Label(ge, text="", foreground="#64748B")
        self.lb_evo.pack(side="left", padx=12)
        self._alternar_evo()
        self._alternar_plan()
        self._plan_cambiado()

    def _panel_carpeta(self, parent=None):
        p = parent or self
        f = ttk.LabelFrame(p, text=" 4 · Guardar en ")
        f.pack(fill="x", padx=14, pady=6)
        g = ttk.Frame(f)
        g.pack(fill="x", padx=10, pady=8)
        g.columnconfigure(0, weight=1)
        self.v_carpeta = tk.StringVar(value=os.path.join(CARPETA, "informes"))
        ttk.Entry(g, textvariable=self.v_carpeta).grid(row=0, column=0, sticky="ew")
        ttk.Button(g, text="Examinar…", command=self._elegir_carpeta).grid(row=0, column=1,
                                                                           padx=(8, 0))

    def _panel_accion(self, parent=None):
        p = parent or self
        f = ttk.Frame(p)
        f.pack(fill="x", padx=14, pady=(10, 4))
        self.bt_gen = ttk.Button(f, text="▶  GENERAR INFORMES", style="Big.TButton",
                                 command=self._generar)
        self.bt_gen.pack(side="left", ipadx=18)
        self.pb = ttk.Progressbar(f, mode="determinate", length=260)
        self.pb.pack(side="left", fill="x", expand=True, padx=14)
        self.bt_abrir = ttk.Button(f, text="Abrir carpeta", state="disabled",
                                   command=self._abrir_carpeta)
        self.bt_abrir.pack(side="left")

    def _panel_log(self, parent=None):
        p = parent or self
        f = ttk.Frame(p)
        f.pack(fill="both", expand=True, padx=14, pady=(4, 6))
        self.log = tk.Text(f, height=8, wrap="word", font=("Consolas", 9), bg="#111827",
                           fg="#D1D5DB", insertbackground="#FFFFFF", relief="flat", padx=8, pady=6)
        sb = ttk.Scrollbar(f, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.log.tag_config("ok", foreground="#4ADE80")
        self.log.tag_config("err", foreground="#F87171")
        self.log.tag_config("info", foreground="#93C5FD")
        self.log.tag_config("h2", foreground="#FBBF24")

    def _pie(self, parent=None):
        p = parent or self
        tk.Label(p, text="Fuente de datos: Dashboard Gerencial MINSA | Ficha 32 · "
                            "Los datos se descargan con las credenciales de su DIRESA.",
                 font=("Segoe UI", 8), bg="#F8F9FA", fg="#64748B").pack(side="bottom", pady=(0, 6))

    # ------------------------------------------------------------------ lógica UI
    def _marcar_cambio(self, *_):
        self.v_estado.configure(text="Cambios pendientes: vuelva a consultar el tablero "
                                     "para actualizar la lista de redes e IPRESS.",
                                foreground="#B45309")

    def _token(self):
        return api.TOKEN_POR_NOMBRE.get(self.v_diresa.get().strip().upper(),
                                        f"DIRESA_{self.v_diresa.get().strip().upper()}")

    def _mes_num(self):
        txt = self.v_mes.get()
        for m, nom in MESES_TXT.items():
            if txt.startswith(f"Mes {m}") or f"({nom}" in txt:
                return m
        return M.MES_CORTE

    def _cargar_pronostico(self):
        """Descubre los meses con datos y carga el padrón de eess de la DIRESA."""
        self.cb_diresa.configure(state="disabled")
        self._marcados = set()          # DIRESA nueva: casillas al principio
        self.v_estado.configure(text="Consultando el tablero…", foreground="#64748B")
        threading.Thread(target=self._hilo_pronostico, daemon=True).start()

    def _hilo_pronostico(self):
        token, anio = self._token(), int(self.v_anio.get())
        # 1) Descubrir qué meses tiene publicados esa DIRESA
        try:
            meses = []
            for m in api.MESES:
                try:
                    api.pedir(token, m, anio=anio)
                    meses.append(m)
                except api.ErrorFicha32:
                    continue
        except Exception as e:  # noqa: BLE001
            self.after(0, self._fin_pronostico, [], f"No se pudo consultar: {e}", [], [])
            return
        if not meses:
            self.after(0, self._fin_pronostico, [],
                       "Sin datos publicados para esa credencial.", [], [])
            return

        # 2) Padrón del mes más reciente disponible
        try:
            A, B, ctx = api.obtener(token, meses[-1], anio=anio)
        except Exception as e:  # noqa: BLE001
            self.after(0, self._fin_pronostico, meses, f"No se pudo leer el padrón: {e}", [], [])
            return

        import pandas as pd
        todos = pd.concat([A, B], ignore_index=True)
        todos["_RED"] = M.resolver_red_salud(todos, diresa=token)
        # Categoría del eess: el primer token del nombre ("I-1 HOSPITAL …" → I-1)
        todos["_CAT"] = todos["ESTABLECIMIENTO / IPRESS"].astype(str).str.split().str[0]
        redes = sorted(set(todos["_RED"].dropna()))
        # Enriquecer con resoluciones y categorías oficiales desde SUSALUD RENIPRESS
        self.after(0, lambda: self.v_estado.configure(text="⏳ Sincronizando establecimientos con SUSALUD...", foreground="#D97706"))
        codigos_todos = [str(r["ID IPRESS"]).strip().zfill(8) for _, r in todos.iterrows()]
        try:
            mapa_res = api.obtener_resoluciones_renipress(codigos_todos, token=token)
        except Exception:
            mapa_res = {}

        # (código, nombre_limpio, red, cat_f32, cat_su, estado_cat, res_cat, fec_cat, sub_cat, fec_car, sub_car)
        ipress = []
        for _, r in todos.sort_values(["_RED", "ESTABLECIMIENTO / IPRESS"]).iterrows():
            c = str(r["ID IPRESS"]).strip().zfill(8)
            clean_n = M.short_name(r["ESTABLECIMIENTO / IPRESS"])
            cat_f32 = r["_CAT"]
            r_info = mapa_res.get(c, {})

            cat_su = str(r_info.get("categoria_susalud") or "").strip()
            if not cat_su or cat_su in ("—", "Sin Categoría"):
                cat_su = "—"
                estado_cat = "—"
            elif cat_f32 == cat_su:
                estado_cat = "✓ Igual"
            else:
                estado_cat = f"⚠ Varió ({cat_f32}→{cat_su})"

            res_cat = r_info.get("nro_categorizacion") or "—"
            fec_cat = r_info.get("fecha_categorizacion") or "—"
            sub_cat = r_info.get("subio_categorizacion") or "NO SUBIÓ"
            fec_car = r_info.get("fecha_cartera") or "—"
            sub_car = r_info.get("subio_cartera") or "NO SUBIÓ"
            ipress.append((c, clean_n, r["_RED"], cat_f32, cat_su, estado_cat, res_cat, fec_cat, sub_cat, fec_car, sub_car))
        self.after(0, self._fin_pronostico, meses, f"{len(todos)} establecimientos · "
                                                   f"{ctx.get('rol', '')}", redes, ipress)

    def _fin_pronostico(self, meses, estado, redes, ipress):
        if meses:
            self.cb_mes.configure(values=[f"Mes {m} ({MESES_TXT[m]})" for m in meses])
            actual = self.v_mes.get()
            if not any(actual.startswith(f"Mes {m}") for m in meses):
                self.v_mes.set(f"Mes {meses[-1]} ({MESES_TXT[meses[-1]]})")
        self._redes, self._ipress = redes, ipress
        self.cb_red.configure(values=["(todas)"] + redes)
        if self.v_red.get() not in ["(todas)"] + redes:
            self.v_red.set("(todas)")
        cats = sorted({item[3] for item in ipress if len(item) > 3})
        self.cb_nivel.configure(values=["(todas)"] + cats)
        if self.v_nivel.get() not in ["(todas)"] + cats:
            self.v_nivel.set("(todas)")
        if meses:
            # Rango de la hoja Evolución: hasta el mes de corte disponible
            self.v_evo_hasta.set(min(self._mes_num(), meses[-1]))
            self.v_evo_desde.set(1)
            self.sp_hasta.configure(to=meses[-1])
            self.sp_desde.configure(to=max(1, meses[-1] - 1))
            self.lb_evo.configure(text=f"(datos disponibles: M1–M{meses[-1]})")
        self.cb_diresa.configure(state="readonly")
        self.v_estado.configure(text=f"Listo · {estado}", foreground="#166534")
        self._cambiar_alcance()

    def _mes_cambiado(self):
        """El rango de la hoja Evolución no puede pasar del mes de corte."""
        m = self._mes_num()
        self.v_evo_hasta.set(m)
        self.sp_hasta.configure(to=m)
        self.sp_desde.configure(to=max(1, m - 1))

    def _cambiar_alcance(self):
        modo = self.v_alcance.get()
        activo = modo == "eess" and bool(self._ipress)
        st_red = "readonly" if activo else "disabled"
        st_bt = "normal" if activo else "disabled"
        self.cb_red.configure(state=st_red)
        self.cb_nivel.configure(state=st_red)
        self.bt_marcar.configure(state=st_bt)
        self.bt_desmarcar.configure(state=st_bt)
        # Al informar sobre eess seleccionados no se genera la parte de redes
        self.cb_redes_bt.configure(state="disabled" if activo else "normal")
        if activo:
            self.v_redes.set(False)
        self._refrescar_arbol()
        if activo and len(self._marcados) == 0:
            # Al pasar a modo casillas con la red "(todas)", se marcan todos:
            # el caso habitual es reportar toda la DIRESA y luego desmarcar alguna.
            self._marcar_visibles(True)

    # ------------------------------------------------------------------ casillas
    def _filtrados(self):
        """Eess que deben mostrarse según la red, la categoría y el texto de búsqueda."""
        red, nivel = self.v_red.get(), self.v_nivel.get()
        txt = self.v_buscar.get().strip().upper()
        sal = []
        for item in self._ipress:
            if len(item) >= 11:
                cod, nom, prov, cat_f32, cat_su, estado_cat, res_cat, fec_cat, sub_cat, fec_car, sub_car = item[:11]
            else:
                cod, nom, prov, cat_f32 = item[0], item[1], item[2], item[3]
                cat_su, estado_cat = "—", "—"
                res_cat = item[4] if len(item) > 4 else "—"
                fec_cat = item[5] if len(item) > 5 else "—"
                sub_cat = item[6] if len(item) > 6 else "NO SUBIÓ"
                fec_car = item[7] if len(item) > 7 else "—"
                sub_car = item[8] if len(item) > 8 else "NO SUBIÓ"

            if red and red != "(todas)" and prov != red:
                continue
            if nivel and nivel != "(todas)" and cat_f32 != nivel and cat_su != nivel:
                continue
            if txt:
                buscables = f"{nom} {cod} {cat_f32} {cat_su} {estado_cat} {res_cat} {fec_cat} {sub_cat} {fec_car} {sub_car}".upper()
                if txt not in buscables:
                    continue
            sal.append((cod, nom, prov, cat_f32, cat_su, estado_cat, res_cat, fec_cat, sub_cat, fec_car, sub_car))
        return sal

    def _refrescar_arbol(self):
        self.tree.delete(*self.tree.get_children())
        for cod, nom, prov, cat_f32, cat_su, estado_cat, res_cat, fec_cat, sub_cat, fec_car, sub_car in self._filtrados():
            marcado = cod in self._marcados
            tags = ()
            if str(estado_cat).startswith("⚠"):
                tags = ("vario",)
            elif estado_cat == "✓ Igual":
                tags = ("igual",)
            self.tree.insert("", "end", iid=cod,
                             values=("☑" if marcado else "☐", prov, cat_f32, cat_su, estado_cat, cod, nom,
                                     res_cat, fec_cat, sub_cat, fec_car, sub_car),
                             tags=tags)
        self._actualizar_cuenta()

    def _click_arbol(self, ev):
        """Un clic sobre una fila alterna su casilla."""
        if self.v_alcance.get() != "eess":
            return
        item = self.tree.identify_row(ev.y)
        if not item:
            return
        if item in self._marcados:
            self._marcados.discard(item)
            self.tree.set(item, "sel", "☐")
        else:
            self._marcados.add(item)
            self.tree.set(item, "sel", "☑")
        self._actualizar_cuenta()

    def _marcar_visibles(self, valor):
        for cod, *_ in self._filtrados():
            if valor:
                self._marcados.add(cod)
            else:
                self._marcados.discard(cod)
            if self.tree.exists(cod):
                self.tree.set(cod, "sel", "☑" if valor else "☐")
        self._actualizar_cuenta()

    def _actualizar_cuenta(self):
        n, t = len(self._marcados), len(self._ipress)
        if self.v_alcance.get() != "eess":
            self.lb_cuenta.configure(text=f"Se incluirán los {t} establecimientos de la DIRESA.",
                                     foreground="#64748B")
            self.lb_alcance.configure(text="")
        else:
            color = "#166534" if n else "#B45309"
            self.lb_cuenta.configure(text=f"{n} de {t} eess marcados.",
                                     foreground=color)
            self.lb_alcance.configure(text="")

    def _alternar_plan(self):
        st = "normal" if self.v_plan.get() else "disabled"
        self.sp_plan_desde.configure(state=st)
        self.sp_plan_hasta.configure(state=st)

    def _plan_cambiado(self):
        try:
            d = max(1, min(12, int(self.v_plan_desde.get())))
            h = max(1, min(12, int(self.v_plan_hasta.get())))
            if d > h:
                txt = "⚠ (Inicio mayor que fin)"
            else:
                meses = h - d + 1
                txt = f"({M.MES_TXT[d]}–{M.MES_TXT[h]} · {meses} mes{'es' if meses > 1 else ''})"
            self.lb_plan_txt.configure(text=txt)
        except Exception:
            pass

    def _alternar_evo(self):
        if not self.v_excel.get():
            self.v_evo.set(False)
        st = "normal" if (self.v_excel.get() and self.v_evo.get()) else "disabled"
        self.sp_desde.configure(state=st)
        self.sp_hasta.configure(state=st)

    def _elegir_carpeta(self):
        d = filedialog.askdirectory(initialdir=self.v_carpeta.get() or CARPETA,
                                    title="Carpeta donde guardar los informes")
        if d:
            self.v_carpeta.set(d)

    # ------------------------------------------------------------------ generación
    def _generar(self):
        if self._generando or (self._proceso and self._proceso.poll() is None):
            return
        if not (self.v_excel.get() or self.v_word.get() or self.v_graf.get()):
            messagebox.showwarning("Nada que generar",
                                   "Marque al menos el libro Excel, el informe Word o los gráficos.")
            return
        alcance = []
        if self.v_alcance.get() == "eess":
            if not self._marcados:
                messagebox.showwarning("Sin eess marcados",
                                       "Marque al menos un establecimiento con la casilla ☑ "
                                       "o elija «Toda la DIRESA».")
                return
            red_sel = self.v_red.get()
            cods_red = {it[0] for it in self._ipress if it[2] == red_sel} if (red_sel and red_sel != "(todas)") else set()
            if cods_red and self._marcados == cods_red:
                alcance.append(f"red:{red_sel}")
            elif len(self._marcados) < len(self._ipress):
                alcance.append("ipress:" + ",".join(sorted(self._marcados)))

        diresa_val = self.v_diresa.get().strip().upper()
        mes_val = int(self._mes_num())
        anio_val = int(self.v_anio.get())
        carpeta_val = self.v_carpeta.get().strip() or os.path.join(CARPETA, "informes")

        p_d = max(1, min(12, int(self.v_plan_desde.get())))
        p_h = max(1, min(12, int(self.v_plan_hasta.get())))
        if p_d > p_h:
            p_d, p_h = p_h, p_d

        evo_activo = bool(self.v_evo.get() and self.v_excel.get())
        e_d = max(1, int(self.v_evo_desde.get()))
        e_h = min(mes_val, int(self.v_evo_hasta.get()))
        if e_d > e_h:
            e_d, e_h = e_h, e_d

        params = {
            "diresa": diresa_val,
            "mes": mes_val,
            "anio": anio_val,
            "excel": bool(self.v_excel.get()),
            "word": bool(self.v_word.get()),
            "graficos": bool(self.v_graf.get()),
            "evolucion": evo_activo,
            "alcance": alcance or None,
            "carpeta": carpeta_val,
            "refrescar": bool(self.v_refrescar.get()),
            "plan": bool(self.v_plan.get()),
            "hospitales": bool(self.v_hosp.get()),
            "redes": bool(self.v_redes.get()),
            "evo_desde": e_d,
            "evo_hasta": e_h,
            "plan_desde": p_d if self.v_plan.get() else None,
            "plan_hasta": p_h if self.v_plan.get() else None,
            "cartera": bool(self.v_cartera.get()),
        }

        self.log.delete("1.0", "end")
        resumen_cmd = f"Generando: DIRESA={diresa_val} · Mes={mes_val} · Año={anio_val}"
        if alcance:
            txt_alc = alcance[0]
            if len(txt_alc) > 60:
                txt_alc = f"ipress ({len(self._marcados)} seleccionados)"
            resumen_cmd += f" · Alcance={txt_alc}"
        self._escribir(resumen_cmd + "\n", "h2")

        self._generando = True
        self.bt_gen.configure(state="disabled")
        self.bt_abrir.configure(state="disabled")
        self.pb.configure(mode="indeterminate")
        self.pb.start(12)
        threading.Thread(target=self._hilo_generar_directo, args=(params,), daemon=True).start()

    def _procesar_linea_salida(self, linea):
        linea = str(linea).rstrip("\n")
        if not linea:
            return
        if linea.startswith("[progreso]"):
            try:
                n = int(linea.split()[1].split("/")[0])
                t = int(linea.split()[1].split("/")[1])
                self._tot_progreso = t
                self.after(0, self._poner_progreso, n, t)
            except (IndexError, ValueError):
                pass
            return
        if linea.startswith("[carpeta]"):
            self._ultima_carpeta_salida = linea.split("[carpeta]", 1)[1].strip()
        elif "→" in linea:
            partes = linea.split("→", 1)
            candidato = partes[1].strip()
            if os.path.isdir(candidato):
                self._ultima_carpeta_salida = candidato

        etiqueta = {"[ok]": "ok", "[error]": "err", "[info]": "info"}.get(linea[:5], None)
        self.after(0, self._escribir, linea + "\n", etiqueta)

    def _hilo_generar_directo(self, params):
        self._tot_progreso = 0
        codigo = 0
        generar_todo.OUT_CALLBACK = self._procesar_linea_salida
        try:
            generar_todo.generar(**params)
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            try:
                with open(LOG_ERRORES, "a", encoding="utf-8") as fh:
                    fh.write(f"\n=== ERROR GENERACIÓN {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n{tb}")
            except Exception:
                pass
            self._procesar_linea_salida(f"[error] {e}")
            codigo = 1
        finally:
            generar_todo.OUT_CALLBACK = None
            self._generando = False
        self.after(0, self._terminar, codigo, getattr(self, "_tot_progreso", 0))

    def _poner_progreso(self, n, t):
        self.pb.stop()
        self.pb.configure(mode="determinate", maximum=max(t, 1), value=n)

    def _escribir(self, txt, tag=None):
        self.log.insert("end", txt, tag or "")
        self.log.see("end")

    def _terminar(self, codigo, _tot):
        self.pb.stop()
        self.bt_gen.configure(state="normal")
        if codigo == 0:
            self.pb.configure(value=100)
            carp = getattr(self, "_ultima_carpeta_salida", None) or self.v_carpeta.get().strip()
            self._escribir(f"\n✓ Informes generados correctamente.\n📁 Guardado en:\n  {carp}\n", "ok")
            self.bt_abrir.configure(state="normal")
        else:
            self.pb.configure(value=0)
            self._escribir("\n✗ Hubo un problema. Revise el mensaje en rojo.\n", "err")

    def _abrir_carpeta(self):
        d = getattr(self, "_ultima_carpeta_salida", None)
        if not d or not os.path.isdir(d):
            d = self.v_carpeta.get().strip()
        if os.path.isdir(d):
            os.startfile(d)

    def _al_cerrar(self):
        if self._generando or (self._proceso and self._proceso.poll() is None):
            if messagebox.askyesno("Proceso en ejecución",
                                   "Hay un proceso de generación en marcha.\n¿Desea cancelarlo y cerrar la aplicación?"):
                try:
                    self._proceso.terminate()
                except Exception:
                    pass
                self.destroy()
        else:
            self.destroy()


    def _abrir_dialogo_config(self):
        """Diálogo para configurar URL de la API y parámetros de fórmulas (2026 / 2027)."""
        import config_manager as CM
        cfg = CM.cargar_config()

        top = tk.Toplevel(self)
        top.title("Configuración de API y Fórmulas — Ficha 32")
        top.geometry("740x630")
        top.minsize(700, 560)
        top.resizable(True, True)
        top.transient(self)
        top.grab_set()

        # Cabecera
        f_head = tk.Frame(top, bg="#573375", padx=16, pady=10)
        f_head.pack(fill="x")
        tk.Label(f_head, text="Ajustes de Conexión y Fórmulas de Cálculo", font=("Segoe UI", 12, "bold"),
                 bg="#573375", fg="#FFFFFF").pack(anchor="w")
        tk.Label(f_head, text="Personalice el endpoint de la API (para 2026/2027), las ponderaciones de servicios "
                              "y los factores de cálculo sin tocar el código fuente.",
                 font=("Segoe UI", 9), bg="#573375", fg="#E8C7FF", wraplength=640).pack(anchor="w", pady=(2, 0))

        # Pestañas
        nb = ttk.Notebook(top)
        nb.pack(fill="both", expand=True, padx=14, pady=(10, 6))

        tab1 = ttk.Frame(nb, padding=12)
        tab2 = ttk.Frame(nb, padding=12)
        nb.add(tab1, text="  ⚙ Conexión y Ponderaciones  ")
        nb.add(tab2, text="  📐 Metas y Fórmulas 2026/2027  ")

        # -------------------------------------------------------------
        # PESTAÑA 1: CONEXIÓN, PESOS Y SEMÁFORO
        # -------------------------------------------------------------
        # Marco API
        f_api = ttk.LabelFrame(tab1, text=" Conexión a la API del Tablero ", padding=10)
        f_api.pack(fill="x", pady=(0, 10))

        ttk.Label(f_api, text="URL del Endpoint:").grid(row=0, column=0, sticky="w", pady=4)
        v_url = tk.StringVar(value=cfg["api"]["url"])
        ent_url = ttk.Entry(f_api, textvariable=v_url, width=54)
        ent_url.grid(row=0, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=4)
        f_api.columnconfigure(1, weight=1)

        ttk.Label(f_api, text="Timeout (segundos):").grid(row=1, column=0, sticky="w", pady=4)
        v_timeout = tk.IntVar(value=cfg["api"].get("timeout_segundos", 15))
        ttk.Spinbox(f_api, from_=5, to=120, textvariable=v_timeout, width=8).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=4)

        ttk.Label(f_api, text="Reintentos:").grid(row=1, column=2, sticky="e", padx=(16, 8), pady=4)
        v_reintentos = tk.IntVar(value=cfg["api"].get("reintentos", 2))
        ttk.Spinbox(f_api, from_=1, to=5, textvariable=v_reintentos, width=6).grid(row=1, column=3, sticky="w", pady=4)

        # Marco Ponderaciones
        f_pesos = ttk.LabelFrame(tab1, text=" Ponderaciones de Servicios de Telemedicina (Pesos del Indicador) ", padding=10)
        f_pesos.pack(fill="x", pady=(0, 10))

        g_p = ttk.Frame(f_pesos)
        g_p.pack(fill="x")
        g_p.columnconfigure(2, weight=1)

        servicios_cfg = CM.get_servicios()
        vars_pesos = {}
        for s_idx, s in enumerate(servicios_cfg):
            cod = s.get("codigo")
            nom = s.get("nombre", cod)
            peso_val = s.get("peso", 0)
            v = tk.IntVar(value=peso_val)
            vars_pesos[cod] = v
            desc = s.get("descripcion") or ("Servicio principal (consultas de apoyo entre IPRESS y Hospitales)" if cod == "TI"
                                           else ("Atención directa médico-paciente a distancia" if cod == "TC"
                                           else "Seguimiento a distancia de pacientes crónicos / Anexo 1"))
            ttk.Label(g_p, text=f"{nom} ({cod} %):", font=("Segoe UI", 9, "bold")).grid(row=s_idx, column=0, sticky="w", pady=4)
            ttk.Spinbox(g_p, from_=0, to=100, textvariable=v, width=7).grid(row=s_idx, column=1, sticky="w", padx=(10, 14), pady=4)
            ttk.Label(g_p, text=desc, font=("Segoe UI", 8)).grid(row=s_idx, column=2, sticky="w", pady=4)

        # Separador y Totalizador en tiempo real
        sep_row = len(servicios_cfg)
        ttk.Separator(g_p, orient="horizontal").grid(row=sep_row, column=0, columnspan=3, sticky="ew", pady=(8, 8))

        ttk.Label(g_p, text="Suma total de pesos:", font=("Segoe UI", 9, "bold")).grid(row=sep_row + 1, column=0, sticky="w")
        lbl_total = tk.Label(g_p, font=("Segoe UI", 9, "bold"), padx=8, pady=2)
        lbl_total.grid(row=sep_row + 1, column=1, columnspan=2, sticky="w", padx=(6, 0))

        def actualizar_total(*_):
            try:
                tot = sum(int(v.get()) for v in vars_pesos.values())
            except Exception:
                tot = 0
            if tot == 100:
                lbl_total.config(text=f"{tot}%  ✓  Correcto (suma 100%)", bg="#DCFCE7", fg="#15803D")
            else:
                lbl_total.config(text=f"{tot}%  ⚠  Atención: La suma debe ser exactamente 100%", bg="#FEE2E2", fg="#B91C1C")

        for v in vars_pesos.values():
            v.trace_add("write", actualizar_total)
        actualizar_total()

        # Marco Semáforo
        f_sem = ttk.LabelFrame(tab1, text=" Umbrales de Calificación (Semáforo) ", padding=10)
        f_sem.pack(fill="x")

        g_s = ttk.Frame(f_sem)
        g_s.pack(fill="x")
        ttk.Label(g_s, text="Mínimo Verde (Cumple):").grid(row=0, column=0, sticky="w")
        v_sverde = tk.DoubleVar(value=cfg["formulas"]["semaforo"].get("verde_min", 60.0))
        ttk.Spinbox(g_s, from_=10, to=100, textvariable=v_sverde, width=7).grid(row=0, column=1, padx=(6, 24))

        ttk.Label(g_s, text="Mínimo Amarillo (Riesgo):").grid(row=0, column=2, sticky="w")
        v_samarillo = tk.DoubleVar(value=cfg["formulas"]["semaforo"].get("amarillo_min", 30.0))
        ttk.Spinbox(g_s, from_=0, to=90, textvariable=v_samarillo, width=7).grid(row=0, column=3, padx=(6, 24))

        ttk.Label(g_s, text="Rojo (Crítico): < Amarillo", font=("Segoe UI", 8, "italic")).grid(row=0, column=4, sticky="w")

        # -------------------------------------------------------------
        # PESTAÑA 2: METAS Y FÓRMULAS 2026/2027
        # -------------------------------------------------------------
        f_tipo_a = ttk.LabelFrame(tab2, text=" Metas Anuales — Primer Nivel (Reporte A) ", padding=10)
        f_tipo_a.pack(fill="x", pady=(0, 10))

        ma = cfg["formulas"].get("metas_tipo_a", {})
        g_a = ttk.Frame(f_tipo_a)
        g_a.pack(fill="x")

        # TI
        ttk.Label(g_a, text="TI:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=2)
        ttk.Label(g_a, text="Piso mínimo:").grid(row=0, column=1, sticky="w", padx=(10, 4))
        v_piso_ti = tk.IntVar(value=ma.get("TI", {}).get("piso", 24))
        ttk.Spinbox(g_a, from_=1, to=1000, textvariable=v_piso_ti, width=6).grid(row=0, column=2, sticky="w")
        ttk.Label(g_a, text="Factor Denominador:").grid(row=0, column=3, sticky="w", padx=(16, 4))
        v_fac_ti = tk.DoubleVar(value=ma.get("TI", {}).get("factor_denominador", 0.0120))
        ttk.Spinbox(g_a, from_=0.0001, to=0.5, increment=0.001, textvariable=v_fac_ti, width=8).grid(row=0, column=4, sticky="w")
        ttk.Label(g_a, text="(ej. 0.012 = 1.2% pob)", font=("Segoe UI", 8)).grid(row=0, column=5, sticky="w", padx=(6, 0))

        # TC
        ttk.Label(g_a, text="TC:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=2)
        ttk.Label(g_a, text="Piso mínimo:").grid(row=1, column=1, sticky="w", padx=(10, 4))
        v_piso_tc = tk.IntVar(value=ma.get("TC", {}).get("piso", 1))
        ttk.Spinbox(g_a, from_=1, to=1000, textvariable=v_piso_tc, width=6).grid(row=1, column=2, sticky="w")
        ttk.Label(g_a, text="Factor Denominador:").grid(row=1, column=3, sticky="w", padx=(16, 4))
        v_fac_tc = tk.DoubleVar(value=ma.get("TC", {}).get("factor_denominador", 0.0010))
        ttk.Spinbox(g_a, from_=0.0001, to=0.5, increment=0.001, textvariable=v_fac_tc, width=8).grid(row=1, column=4, sticky="w")
        ttk.Label(g_a, text="(ej. 0.001 = 0.1% pob)", font=("Segoe UI", 8)).grid(row=1, column=5, sticky="w", padx=(6, 0))

        # TM
        ttk.Label(g_a, text="TM:", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", pady=2)
        ttk.Label(g_a, text="Piso mínimo:").grid(row=2, column=1, sticky="w", padx=(10, 4))
        v_piso_tm = tk.IntVar(value=ma.get("TM", {}).get("piso", 24))
        ttk.Spinbox(g_a, from_=1, to=1000, textvariable=v_piso_tm, width=6).grid(row=2, column=2, sticky="w")
        ttk.Label(g_a, text="Factor Denominador:").grid(row=2, column=3, sticky="w", padx=(16, 4))
        v_fac_tm = tk.DoubleVar(value=ma.get("TM", {}).get("factor_denominador", 0.0072))
        ttk.Spinbox(g_a, from_=0.0001, to=0.5, increment=0.001, textvariable=v_fac_tm, width=8).grid(row=2, column=4, sticky="w")
        ttk.Label(g_a, text="(ej. 0.0072 = 0.72% pob)", font=("Segoe UI", 8)).grid(row=2, column=5, sticky="w", padx=(6, 0))

        # Hospitales (Tipo B)
        f_tipo_b = ttk.LabelFrame(tab2, text=" Metas Anuales — Hospitales e Institutos (Reporte B) ", padding=10)
        f_tipo_b.pack(fill="x", pady=(0, 10))

        mb = cfg["formulas"].get("metas_tipo_b", {})
        g_b = ttk.Frame(f_tipo_b)
        g_b.pack(fill="x")
        ttk.Label(g_b, text="TI Hospitales:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(g_b, text="Piso mínimo:").grid(row=0, column=1, sticky="w", padx=(10, 4))
        v_piso_b = tk.IntVar(value=mb.get("TI", {}).get("piso", 100))
        ttk.Spinbox(g_b, from_=10, to=5000, textvariable=v_piso_b, width=6).grid(row=0, column=2, sticky="w")
        ttk.Label(g_b, text="Factor Denominador:").grid(row=0, column=3, sticky="w", padx=(16, 4))
        v_fac_b = tk.DoubleVar(value=mb.get("TI", {}).get("factor_denominador", 0.0040))
        ttk.Spinbox(g_b, from_=0.0001, to=0.5, increment=0.001, textvariable=v_fac_b, width=8).grid(row=0, column=4, sticky="w")
        ttk.Label(g_b, text="(ej. 0.004 = 0.4% pob)", font=("Segoe UI", 8)).grid(row=0, column=5, sticky="w", padx=(6, 0))

        # Ventana de Planificación por Defecto
        f_plan_cfg = ttk.LabelFrame(tab2, text=" Ventana y Capacidad del Plan de Cumplimiento ", padding=10)
        f_plan_cfg.pack(fill="x", pady=(0, 10))

        cplan = cfg.get("plan", {})
        g_pc = ttk.Frame(f_plan_cfg)
        g_pc.pack(fill="x")

        ttk.Label(g_pc, text="Mes Inicio Plan:").grid(row=0, column=0, sticky="w")
        v_plan_ini_cfg = tk.IntVar(value=cplan.get("mes_inicio_plan", 11))
        ttk.Spinbox(g_pc, from_=1, to=12, textvariable=v_plan_ini_cfg, width=6).grid(row=0, column=1, padx=(6, 20), sticky="w")

        ttk.Label(g_pc, text="Mes Fin Plan:").grid(row=0, column=2, sticky="w")
        v_plan_fin_cfg = tk.IntVar(value=cplan.get("mes_fin_plan", 12))
        ttk.Spinbox(g_pc, from_=1, to=12, textvariable=v_plan_fin_cfg, width=6).grid(row=0, column=3, padx=(6, 20), sticky="w")

        ttk.Label(g_pc, text="Piso Capacidad Mensual:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        v_piso_cap = tk.IntVar(value=cplan.get("piso_capacidad_mes", 12))
        ttk.Spinbox(g_pc, from_=1, to=100, textvariable=v_piso_cap, width=6).grid(row=1, column=1, padx=(6, 20), pady=(8, 0), sticky="w")

        ttk.Label(g_pc, text="% Capacidad s/Meta:").grid(row=1, column=2, sticky="w", pady=(8, 0))
        v_pct_cap = tk.DoubleVar(value=cplan.get("pct_capacidad_meta", 0.05))
        ttk.Spinbox(g_pc, from_=0.01, to=0.5, increment=0.01, textvariable=v_pct_cap, width=6).grid(row=1, column=3, padx=(6, 20), pady=(8, 0), sticky="w")

        # Guía sobre otros indicadores en 2027
        f_info = ttk.LabelFrame(tab2, text=" ¿Cómo se adaptan otros indicadores o cambios para el año 2027? ", padding=10)
        f_info.pack(fill="x")
        info_txt = (
            "\u2022 Nuevas ponderaciones o servicios: Si MINSA modifica los pesos en 2027 (ej. 50% TI, 25% TC, 25% TM) "
            "o incorpora nuevos servicios (Teleorientaci\u00f3n / TO, Teleapoyo / TAD), se actualizan los pesos en la Pesta\u00f1a 1 o en config.json.\n"
            "\u2022 Ajuste de metas: Si los convenios 2027 elevan los pisos m\u00ednimos o los porcentajes poblacionales, "
            "se ajustan en los controles superiores y se recalculan al instante todos los informes.\n"
            "\u2022 Cambio de API / Servidor 2027: Bastar\u00e1 con colocar el nuevo enlace en 'URL del Endpoint'."
        )
        tk.Label(f_info, text=info_txt, font=("Segoe UI", 8), fg="#334155", justify="left", wraplength=600).pack(anchor="w")

        # -------------------------------------------------------------
        # BOTONES
        # -------------------------------------------------------------
        f_btns = ttk.Frame(top, padding=12)
        f_btns.pack(fill="x", side="bottom")

        def guardar():
            try:
                tot = sum(int(v.get()) for v in vars_pesos.values())
            except Exception:
                tot = 0
            if tot != 100:
                msg_advertencia = f"La suma actual de pesos es {tot}% (la ficha oficial requiere 100%).\n\n¿Desea guardar de todas formas?"
                if not messagebox.askyesno("Advertencia de Ponderación", msg_advertencia, parent=top):
                    return

            # Guardar pesos dinámicos en configuración
            pesos_map = {cod: int(v.get()) for cod, v in vars_pesos.items()}
            CM.guardar_servicios_pesos(pesos_map)

            nueva_cfg = {
                "api": {
                    "url": v_url.get().strip(),
                    "timeout_segundos": int(v_timeout.get()),
                    "reintentos": int(v_reintentos.get()),
                },
                "formulas": {
                    "pesos": pesos_map,
                    "metas_tipo_a": {
                        "TI": {"piso": int(v_piso_ti.get()), "factor_denominador": float(v_fac_ti.get())},
                        "TC": {"piso": int(v_piso_tc.get()), "factor_denominador": float(v_fac_tc.get())},
                        "TM": {"piso": int(v_piso_tm.get()), "factor_denominador": float(v_fac_tm.get())},
                    },
                    "metas_tipo_b": {
                        "TI": {"piso": int(v_piso_b.get()), "factor_denominador": float(v_fac_b.get())},
                    },
                    "semaforo": {
                        "verde_min": float(v_sverde.get()),
                        "amarillo_min": float(v_samarillo.get()),
                    }
                },
                "plan": {
                    "mes_inicio_plan": int(v_plan_ini_cfg.get()),
                    "mes_fin_plan": int(v_plan_fin_cfg.get()),
                    "piso_capacidad_mes": int(v_piso_cap.get()),
                    "pct_capacidad_meta": float(v_pct_cap.get()),
                }
            }
            CM.guardar_config(nueva_cfg)
            # Actualizar controles en la ventana principal
            self.v_plan_desde.set(int(v_plan_ini_cfg.get()))
            self.v_plan_hasta.set(int(v_plan_fin_cfg.get()))
            self._plan_cambiado()
            # Actualizar en caliente módulos
            api.API = v_url.get().strip()
            M.configurar(diresa=self.v_diresa.get().strip().upper(), mes=self._mes_num(),
                         mes_ini_plan=int(v_plan_ini_cfg.get()), mes_fin_plan=int(v_plan_fin_cfg.get()))
            messagebox.showinfo("Configuración guardada",
                                "Los cambios han sido guardados en config.json y aplicados correctamente.",
                                parent=top)
            top.destroy()
            self._cargar_pronostico()

        def restaurar():
            v_url.set("https://api-ficha32-ditel-303392539829.europe-west1.run.app")
            v_timeout.set(15)
            v_reintentos.set(2)
            pesos_def = {"TI": 60, "TC": 10, "TM": 30}
            for cod, var_p in vars_pesos.items():
                var_p.set(pesos_def.get(cod, 0))
            v_sverde.set(60.0)
            v_samarillo.set(30.0)
            v_piso_ti.set(24)
            v_fac_ti.set(0.0120)
            v_piso_tc.set(1)
            v_fac_tc.set(0.0010)
            v_piso_tm.set(24)
            v_fac_tm.set(0.0072)
            v_piso_b.set(100)
            v_plan_ini_cfg.set(11)
            v_plan_fin_cfg.set(12)
            v_piso_cap.set(12)
            v_pct_cap.set(0.05)
            v_fac_b.set(0.0040)

        ttk.Button(f_btns, text="Restaurar valores oficiales", command=restaurar).pack(side="left")
        ttk.Button(f_btns, text="Guardar y aplicar", style="Big.TButton", command=guardar).pack(side="right")
        ttk.Button(f_btns, text="Cancelar", command=top.destroy).pack(side="right", padx=(0, 8))

if __name__ == "__main__":
    if "--probar-licencia" in sys.argv:
        os.environ["FICHA32_ENFORCE_LICENSE"] = "1"
    if "--cli-generar" in sys.argv:
        args_cli = [a for a in sys.argv[1:] if a != "--cli-generar"]
        generar_todo.main(args_cli)
        sys.exit(0)
    import licencia_manager as LM
    LM.verificar_o_solicitar_licencia()
    try:
        App().mainloop()
    except Exception as _e:  # noqa: BLE001
        import traceback
        with open(LOG_ERRORES, "a", encoding="utf-8") as _fh:
            _fh.write(f"\n=== {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")
            _fh.write(traceback.format_exc())
        try:
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("Error al abrir la aplicación",
                                 f"Ocurrió un error:\n\n{_e}\n\nDetalle guardado en:\n{LOG_ERRORES}")
        except Exception:
            pass
        raise
