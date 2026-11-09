# Ficha N.° 32 — Indicador de Telemedicina (MINSA / DITEL)
## Guía de Instalación, Uso y Arquitectura Escalable

Esta aplicación procesa las evaluaciones oficiales de la **Ficha Técnica N.° 32** de Convenios de Gestión (Tasa de uso de los servicios de telemedicina), genera libros Excel con fórmulas nativas, informes ejecutivos en Word y cruza automáticamente la información con el **Registro Nacional de IPRESS (SUSALUD / RENIPRESS)**.

---

### 🚀 Formas de Iniciar la Aplicación

#### Método 1: Acceso Directo de Escritorio Silencioso (Recomendado)
* **Archivo:** `Iniciar_Ficha32.vbs`
* Haga **doble clic** sobre `Iniciar_Ficha32.vbs`.
* La aplicación se abrirá de inmediato como una **aplicación nativa de Windows**, sin mostrar ninguna ventana de consola negra (CMD).
* *Consejo:* Puede hacer clic derecho en `Iniciar_Ficha32.vbs` -> *Enviar a* -> *Escritorio (crear acceso directo)*.

#### Método 2: Lanzador con Diagnóstico y Auto-Instalación
* **Archivo:** `Iniciar_Ficha32.bat`
* Si alguna librería requerida (`pandas`, `openpyxl`, `docx`, `requests`, `matplotlib`, `xlsxwriter`) falta en su computadora, este script la detectará e instalará automáticamente con `pip` antes de iniciar la ventana.

#### Método 3: Compilar un Ejecutable Portable `.exe` (Sin requerir Python)
* **Archivo:** `compilar_ejecutable.bat`
* Si necesita distribuir la herramienta a directores, coordinadores de red o computadoras de centros de salud que **no tienen Python instalado**:
  1. Ejecute `compilar_ejecutable.bat`.
  2. El compilador empaquetará todas las dependencias, librerías y la base de datos offline.
  3. En la carpeta `dist/` se generará el archivo único:
     ```
     dist/Ficha32_Telemedicina.exe
     ```
  4. Ese archivo `.exe` puede copiarse en una memoria USB y ejecutarse en cualquier computadora con Windows (10 u 11).

---

### 📦 Nuevas Capacidades Implementadas

#### 1. Comparativa de Categorías: Ficha 32 vs. SUSALUD (RENIPRESS)
* **Problema:** Con frecuencia los establecimientos de salud son categorizados o recategorizados legalmente por la autoridad sanitaria (e.g. ascienden de I-1 a I-2, o de I-3 a I-4), pero los sistemas de registro del HIS o tableros de DITEL mantienen su nivel desactualizado.
* **Solución:**
  * En la **interfaz gráfica**, la tabla muestra ahora: `Cat. F32`, `Cat. SUSALUD` y la columna `¿Cambió Cat.?` resaltando en color ámbar las discrepancias (ej. `⚠ Varió (I-3→I-4)`).
  * En el **Libro Excel** (hoja **Cartera Servicios** y hoja **IPRESS**), se registran ambas columnas y el diagnóstico automático del estado de categorización.

#### 2. Control de Tasa (Rate Limiting) y Prioridad Absoluta de Caché
* Se implementó un control de concurrencia y retardo regulado (`pacing delay` de 40ms por petición y máximo 5 trabajadores simultáneos) hacia el servidor de SUSALUD.
* La aplicación **siempre consulta primero el archivo local** `cache_api/renipress/renipress_resoluciones_*.json`. Si la IPRESS ya está registrada en la copia local, **no realiza ninguna llamada de red innecesaria**, acelerando el inicio a menos de un segundo.

#### 3. Estructura de Servicios Dinámica (Preparada para 2027)
* En versiones anteriores, los servicios estaban codificados de forma fija (`TI`, `TC`, `TM`).
* Ahora se gestionan como una lista dinámica en `config.json` y `config_manager.py`:
  ```json
  "servicios": [
    {"codigo": "TI", "nombre": "Teleinterconsulta", "peso": 60, "activo": true, ...},
    {"codigo": "TC", "nombre": "Teleconsulta", "peso": 10, "activo": true, ...},
    {"codigo": "TM", "nombre": "Telemonitoreo", "peso": 30, "activo": true, ...}
  ]
  ```
* Si en el año 2027 el MINSA añade un cuarto servicio (ej. `TO` - Teleorientación) o modifica las ponderaciones:
  * Basta con añadirlo en `config.json` o desde la ventana emergente **Ajustes de API / Fórmulas**.
  * El modelo matemático, el cálculo de puntajes de logro y las hojas del Excel se adaptarán automáticamente sin alterar el código fuente.

#### 4. Indicador Visual de Sincronización
* Mientras la aplicación cruza los datos con la API de SUSALUD en segundo plano, se muestra el estado en color ámbar:
  `⏳ Sincronizando establecimientos con SUSALUD...`
* Al finalizar, se actualiza a verde: `Listo · 375 establecimientos`.

#### 5. Resiliencia 100% Desconectada (Offline)
* Todos los datos regionales de DIRESA Puno (meses 1 a 8 de Ficha 32 y las 736 IPRESS con sus resoluciones de categorización y cartera de SUSALUD) quedan almacenados en la carpeta `cache_api/`.
* Si la computadora no dispone de conexión a internet o el servidor de SUSALUD se encuentra en mantenimiento, la aplicación funciona al 100% sin emitir errores de red.

#### 6. Guardado Ordenado por Fecha y Alcance
* Cada vez que se genera un informe, se crea una carpeta exclusiva fechada y con sello de tiempo para evitar sobreescrituras accidentales o bloqueos de archivos abiertos en Excel:
  ```
  informes/DIRESA_PUNO/2026/Mes 08/Red_EL_COLLAO_2026-10-30_11-45-00/
  ```
