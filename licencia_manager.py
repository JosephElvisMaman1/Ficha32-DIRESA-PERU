# -*- coding: utf-8 -*-
"""Módulo de Control de Licencia y Activación por Hardware (Machine ID) — Ficha N.° 32.

Permite vincular la ejecución del ejecutable (.exe) de forma exclusiva a equipos
autorizados mediante un identificador criptográfico derivado del hardware de la máquina.

Algoritmo:
  1. Identificador de Máquina (Machine ID): Hash SHA-256 de MachineGuid (Windows) + UUID.
     Formato: DIR-XXXX-XXXX-XXXX (12 caracteres en 3 bloques).
  2. Clave de Activación: HMAC-SHA256(SECRET_KEY, Machine ID).
     Formato: ACT-XXXX-XXXX-XXXX.
  3. Almacenamiento: Guarda la firma en APPDATA local y/o en la carpeta de la app.
"""

import hashlib
import hmac
import json
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk

# Clave maestra para la generación de firmas criptográficas de activación
# (Solo el generador privado puede generar claves válidas que coincidan con este secreto)
_SECRET_KEY = b"Ficha32-DIRESA-Telemedicina-MINSA-2026-SecuritySalt#77!"


def obtener_machine_id():
    """Obtiene un identificador único y persistente de la máquina (Windows)."""
    partes = []

    # 1. Windows MachineGuid (clave criptográfica única instalada por Windows)
    try:
        import winreg
        clave = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography")
        guid, _ = winreg.QueryValueEx(clave, "MachineGuid")
        winreg.CloseKey(clave)
        if guid:
            partes.append(str(guid).strip().lower())
    except Exception:
        pass

    # 2. Nombre del equipo y fallback de MAC address
    try:
        comp_name = os.environ.get("COMPUTERNAME", "")
        if comp_name:
            partes.append(comp_name.strip().upper())
    except Exception:
        pass

    try:
        import uuid
        node = uuid.getnode()
        if node:
            partes.append(hex(node))
    except Exception:
        pass

    cadena_base = "|".join(partes) if partes else "FICHA32_FALLBACK_DEVICE_DEFAULT"
    h = hashlib.sha256(cadena_base.encode("utf-8")).hexdigest().upper()
    # Tomar 12 caracteres en 3 bloques de 4: DIR-XXXX-XXXX-XXXX
    return f"DIR-{h[0:4]}-{h[4:8]}-{h[8:12]}"


def calcular_codigo_activacion(machine_id):
    """Calcula el código de activación válido correspondiente a un Machine ID."""
    mid_limpio = str(machine_id).strip().upper()
    firma = hmac.new(_SECRET_KEY, mid_limpio.encode("utf-8"), hashlib.sha256).hexdigest().upper()
    return f"ACT-{firma[0:4]}-{firma[4:8]}-{firma[8:12]}"


def verificar_codigo(machine_id, codigo_ingresado):
    """Verifica si un código de activación coincide con el Machine ID."""
    if not codigo_ingresado:
        return False
    esperado = calcular_codigo_activacion(machine_id)
    c_limpio = str(codigo_ingresado).strip().upper()
    return hmac.compare_digest(esperado, c_limpio)


def _ruta_archivo_licencia():
    """Devuelve la ruta persistente del archivo de licencia."""
    # 1. Preferir APPDATA local para que persista aun si el usuario mueve el .exe
    appdata = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    if appdata and os.path.exists(appdata):
        carpeta = os.path.join(appdata, "Ficha32_Telemedicina")
        os.makedirs(carpeta, exist_ok=True)
        return os.path.join(carpeta, "ficha32.lic")

    # 2. Fallback: junto a la aplicación
    base = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
    return os.path.join(base, "ficha32.lic")


def esta_activado():
    """Comprueba si el equipo actual tiene una licencia válida y registrada."""
    # En desarrollo directo (código fuente Python), se permite bypass automático
    # para no entorpecer el desarrollo, salvo que se active explícitamente la prueba.
    if not getattr(sys, "frozen", False) and os.environ.get("FICHA32_ENFORCE_LICENSE", "0") != "1":
        return True

    ruta = _ruta_archivo_licencia()
    if not os.path.exists(ruta):
        return False

    try:
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
        mid_actual = obtener_machine_id()
        if data.get("machine_id") != mid_actual:
            return False
        return verificar_codigo(mid_actual, data.get("license_key", ""))
    except Exception:
        return False


def guardar_licencia(machine_id, codigo):
    """Guarda la activación en el archivo de licencia local."""
    try:
        ruta = _ruta_archivo_licencia()
        datos = {
            "machine_id": str(machine_id).strip().upper(),
            "license_key": str(codigo).strip().upper(),
            "status": "ACTIVE",
            "system": "MINSA / DIRESA - Ficha 32"
        }
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=2)
        return True
    except Exception:
        return False


class DialogoActivacion(tk.Toplevel):
    """Ventana modal de activación que bloquea el acceso si no hay licencia."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.title("Activación Requerida — Ficha N.° 32 Telemedicina")
        self.geometry("540x440")
        self.resizable(False, False)
        self.configure(bg="#F8F9FA")
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

        self.exito = False
        self.machine_id = obtener_machine_id()

        self._construir_ui()

        # Forzar foco modal
        self.transient(parent)
        self.grab_set()
        self.focus_force()

    def _construir_ui(self):
        # Cabecera
        cab = tk.Frame(self, bg="#573375", pady=14, padx=16)
        cab.pack(fill="x")
        tk.Label(cab, text="🔒 CONTROL DE ACCESO INSTITUCIONAL", font=("Segoe UI", 10, "bold"),
                 bg="#573375", fg="#E8C7FF").pack(anchor="w")
        tk.Label(cab, text="Activación de Software — Ficha N.° 32", font=("Segoe UI", 14, "bold"),
                 bg="#573375", fg="#FFFFFF").pack(anchor="w")

        cuerpo = tk.Frame(self, bg="#F8F9FA", padx=24, pady=16)
        cuerpo.pack(fill="both", expand=True)

        msg = ("Este ejecutable requiere autorización para su uso en este equipo.\n"
               "Proporcione el siguiente Identificador de Máquina al administrador "
               "para recibir su clave de activación:")
        tk.Label(cuerpo, text=msg, font=("Segoe UI", 9), bg="#F8F9FA", fg="#49454F",
                 justify="left", wraplength=480).pack(anchor="w", pady=(0, 10))

        # Cuadro de Identificador de Máquina
        f_id = tk.LabelFrame(cuerpo, text=" Identificador Único de este Equipo ", font=("Segoe UI", 9, "bold"),
                             bg="#F8F9FA", fg="#573375", padx=10, pady=8)
        f_id.pack(fill="x", pady=6)

        sub_id = tk.Frame(f_id, bg="#F8F9FA")
        sub_id.pack(fill="x")

        self.txt_mid = tk.Entry(sub_id, font=("Consolas", 12, "bold"), bg="#EDE7F6", fg="#311B92",
                                relief="solid", bd=1, justify="center")
        self.txt_mid.insert(0, self.machine_id)
        self.txt_mid.configure(state="readonly")
        self.txt_mid.pack(side="left", fill="x", expand=True, padx=(0, 8))

        btn_copiar = tk.Button(sub_id, text="📋 Copiar ID", font=("Segoe UI", 9),
                               bg="#7C3AED", fg="#FFFFFF", relief="flat", padx=10, pady=4,
                               cursor="hand2", command=self._copiar_id)
        btn_copiar.pack(side="right")

        # Cuadro para ingresar Clave
        f_act = tk.LabelFrame(cuerpo, text=" Ingrese su Clave de Activación ", font=("Segoe UI", 9, "bold"),
                              bg="#F8F9FA", fg="#573375", padx=10, pady=10)
        f_act.pack(fill="x", pady=(12, 6))

        self.txt_key = tk.Entry(f_act, font=("Consolas", 12, "bold"), bg="#FFFFFF", fg="#191C1D",
                                relief="solid", bd=1, justify="center")
        self.txt_key.pack(fill="x", pady=4)
        self.txt_key.bind("<Return>", lambda e: self._activar())

        self.lbl_estado = tk.Label(cuerpo, text="", font=("Segoe UI", 9, "bold"), bg="#F8F9FA", fg="#B3261E")
        self.lbl_estado.pack(pady=4)

        # Botones de acción
        f_botones = tk.Frame(cuerpo, bg="#F8F9FA")
        f_botones.pack(fill="x", pady=(10, 0))

        btn_salir = tk.Button(f_botones, text="Salir", font=("Segoe UI", 10),
                              bg="#E0E0E0", fg="#333333", relief="flat", padx=14, pady=6,
                              command=self._cancelar)
        btn_salir.pack(side="left")

        btn_activar = tk.Button(f_botones, text="🔓 Activar Sistema", font=("Segoe UI", 10, "bold"),
                                bg="#573375", fg="#FFFFFF", relief="flat", padx=18, pady=6,
                                cursor="hand2", command=self._activar)
        btn_activar.pack(side="right")

    def _copiar_id(self):
        self.clipboard_clear()
        self.clipboard_append(self.machine_id)
        messagebox.showinfo("Copiado", "Identificador de equipo copiado al portapapeles.\n"
                                       "Envíelo al administrador para solicitar su código.")

    def _activar(self):
        codigo = self.txt_key.get().strip().upper()
        if not codigo:
            self.lbl_estado.configure(text="Por favor ingrese una clave de activación.", fg="#B3261E")
            return

        if verificar_codigo(self.machine_id, codigo):
            if guardar_licencia(self.machine_id, codigo):
                self.exito = True
                messagebox.showinfo("Activación Exitosa",
                                    "¡El sistema ha sido activado exitosamente en este equipo!\n"
                                    "Bienvenido a Ficha N.° 32 Telemedicina.")
                self.grab_release()
                self.destroy()
            else:
                self.lbl_estado.configure(text="Error al guardar el archivo de licencia local.", fg="#B3261E")
        else:
            self.lbl_estado.configure(text="Clave de activación inválida para este equipo.", fg="#B3261E")

    def _cancelar(self):
        self.exito = False
        self.grab_release()
        self.destroy()


def verificar_o_solicitar_licencia():
    """Punto de entrada principal. Devuelve True si está activado o logra activarse.

    Si no se activa, sale de la aplicación de inmediato.
    """
    if esta_activado():
        return True

    # Inicializar ventana temporal de Tk si no hay una activa
    root_creado = False
    root = tk._default_root
    if root is None:
        root = tk.Tk()
        root.withdraw()
        root_creado = True

    dialogo = DialogoActivacion(root)
    root.wait_window(dialogo)

    exito = dialogo.exito
    if root_creado:
        root.destroy()

    if not exito:
        sys.exit(0)

    return True


if __name__ == "__main__":
    mid = obtener_machine_id()
    print("Machine ID actual:", mid)
    print("Código activación correspondiente:", calcular_codigo_activacion(mid))
    print("¿Está activado?:", esta_activado())
