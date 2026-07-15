================================================================================
  README — Carpeta Valeria
  Archivos documentados: Rep_Infra.py | infra_logic.py | Rep_Infra_App.py
                          Rep_Infra.spec | build.bat
================================================================================

DESCRIPCIÓN GENERAL
-------------------
Esta carpeta contiene una aplicación para generar el Reporte de Infraestructura
de Red Noroeste. El sistema toma dos archivos Excel semanales (Semáforo AB y
Semáforo Interconexiones), los procesa y combina en un único CSV normalizado
que describe el estado de cada equipo de red: utilización, semáforo, clientes,
interconexiones, velocidades y más.

La carpeta tiene dos formas de ejecutar el mismo proceso:
  - Rep_Infra.py:     script standalone (directo, sin GUI).
  - Rep_Infra_App.py: aplicación de escritorio con interfaz gráfica (Tkinter),
                      que usa infra_logic.py como motor de procesamiento.

================================================================================

  ARCHIVO: Rep_Infra.py  (script standalone)
  -------------------------------------------

  PROPÓSITO
  ---------
  Versión de script directo. Lee los archivos de entrada con rutas fijas,
  ejecuta el pipeline completo y guarda el CSV de salida. Es la versión
  original/prototipo antes de la integración con la GUI.

  ARCHIVOS DE ENTRADA (rutas fijas)
  ----------------------------------
  - Valeria/SEMAFORO AB_SEM 24-2026.xlsx
      Hoja "2.-Reporte Infra": datos principales de infraestructura por equipo.
      Hoja "3.-Detalle":       datos de Ancho de Banda (utilización, semáforo,
                               tipo de dato, KB x cliente).
  - Valeria/Semaforo Interconexiones SEM24-2026.xlsx
      Hoja "4.-Detalle": datos de interconexiones (promedio utilización, semáforo).

  ARCHIVO DE SALIDA
  -----------------
  - Valeria/Reporte_infraestructura_Red Noroeste_SEM24_2026.csv

  FLUJO DE PROCESAMIENTO
  ----------------------

  Paso 1 — Carga de fuentes
    - df_INFRA:  hoja "2.-Reporte Infra" (primeras 29 columnas).
    - df_AB:     hoja "3.-Detalle" con encabezados en fila 1 (se elimina
                 fila de encabezado duplicada).
    - df_INTXS:  hoja "4.-Detalle" de interconexiones.

  Paso 2 — Construcción de df_SEM1 (tabla intermedia)
    Tabla con una fila por equipo que cruza las tres fuentes:
    - De df_INFRA: AREA, POBLACION, CENTRAL, TECNOLOGIA, NombrePisa,
                   Nombre configurado en equipo, AB REI (interface),
                   TOTAL y OCUPADOS (puertos), NOMBRE BAS, SUBINT BAS.
    - De df_AB:    Avg Util (%), Tipo de Dato, SEMAFORO ACTUAL.
    - De df_INTXS: Prom. de Utilizacion (%) de la interconexión correspondiente
                   (cruzado por BAS + Subinterfaz concatenados como "BAS_INTERFAZ"),
                   SEMAFORO de interconexión.
    - Se renombra POBLACION → "Catálogo TECNOLOGÍA" y NombrePisa → "EQUIPO".

  Paso 3 — Construcción de df_SEM2 (tabla final del reporte)
    Tabla con 24 columnas en el formato requerido por los sistemas de Red Noroeste:

    | Columna                              | Descripción                          |
    |--------------------------------------|--------------------------------------|
    | Division                             | Siempre "RED NOROESTE"               |
    | Area                                 | Área geográfica del equipo           |
    | Siglas_central                       | Siglas de la central                 |
    | NombrePisa                           | Nombre del equipo en PISA            |
    | ip_nom_pisa                          | IP del equipo                        |
    | Tecnologia                           | Tecnología (GPON, ADSL, etc.)        |
    | equipo_410                           | Nombre concatenado TBA               |
    | medio_tx                             | Interfaz del equipo normalizada      |
    |                                      | (GE→GB, CE→"", 1GB→"1 GB 1000",     |
    |                                      |  10GB→"1 GB 10000")                  |
    |                                      | Para tecnología GPON con 1 GB 1000:  |
    |                                      | se cambia a "1 GB 2000"              |
    | clientes                             | Puertos ocupados                     |
    | nombre_configurado                   | equipo_410 si existe, NombrePisa si  |
    |                                      | equipo_410 es "-"                    |
    | nombre_router                        | Nombre del BAS                       |
    | ip_nombre_router                     | IP del BAS                           |
    | interface                            | Subinterfaz BAS                      |
    | subinterface                         | Fijo: "S.D"                          |
    | velocidad                            | KB x cliente (del Semáforo AB)       |
    | promedio_util_salida_interfaz INTX   | % utilización de la interconexión    |
    | semaforo                             | Semáforo del enlace (V/A/R/S)        |
    | color                                | Traducción del semáforo:             |
    |                                      | V→VERDE, A→AMARILLO, R→ROJO,        |
    |                                      | S→SATURADO                           |
    | nombre_subinterace                   | BAS_INTERFAZ (BAS + Subinterfaz)     |
    | VLANHSI                              | VPI/VLAN del equipo                  |
    | velocidad_subint                     | Fijo: 1,000,000,000                  |
    | promedio_util_salida_subinterfaz     | Avg Util (%) de Ancho de Banda       |
    |   ACCESO                             | (redondeado a 2 decimales)           |
    | Prom_util_salida_subinterfaz_kbps    | Avg Util (%) × 10,000,000            |
    | porcentaje_ocupacion INTX            | Igual a promedio INTX                |

================================================================================

  ARCHIVO: infra_logic.py  (motor de procesamiento para la GUI)
  -------------------------------------------------------------

  PROPÓSITO
  ---------
  Encapsula el mismo pipeline de Rep_Infra.py en una función reutilizable
  llamada run(), diseñada para ser invocada desde Rep_Infra_App.py (la GUI)
  en un hilo secundario. Acepta las rutas de los archivos como parámetros
  y un callback on_log para emitir mensajes de progreso en tiempo real.

  USO
  ---
  resultado = infra_logic.run(path_ab, path_intx, on_log=print)
  df_reporte = resultado["df_reporte"]

  PARÁMETROS
  ----------
  - path_ab:   ruta al Excel del Semáforo AB (hojas "2.-Reporte Infra" y "3.-Detalle").
  - path_intx: ruta al Excel del Semáforo Interconexiones (hoja "4.-Detalle").
  - on_log:    función callback para enviar mensajes de progreso (opcional).

  RETORNA
  -------
  Un diccionario: {"df_reporte": pd.DataFrame}
  El DataFrame es la tabla df_SEM2 con las 24 columnas del reporte final.

  DIFERENCIAS CON Rep_Infra.py
  ----------------------------
  - Las rutas de los archivos NO están fijas; se reciben como parámetros.
  - No guarda el CSV automáticamente; devuelve el DataFrame para que la GUI
    decida dónde y cómo guardarlo.
  - Emite mensajes de progreso por cada paso del pipeline (via on_log).
  - La lógica de procesamiento es idéntica a Rep_Infra.py.

================================================================================

  ARCHIVO: Rep_Infra_App.py  (interfaz gráfica — punto de entrada)
  ----------------------------------------------------------------

  PROPÓSITO
  ---------
  Aplicación Tkinter con diseño oscuro y moderno para generar el reporte de
  infraestructura de forma interactiva. El usuario selecciona los dos archivos
  Excel, la app procesa los datos en segundo plano y muestra el resultado en
  una tabla, con opción de descarga como CSV o Excel.

  PANTALLAS
  ---------
  1. UploadFrame (pantalla de carga de archivos)
     - Dos áreas de clic para seleccionar archivos (drag & drop visual):
         · Semáforo AB (.xlsx)
         · Semáforo Interconexión (.xlsx)
     - El botón "GENERAR REPORTE" se habilita solo cuando ambos archivos
       han sido seleccionados.
     - Canvas decorativo de fondo con óvalos.

  2. LoadingFrame
     - Se muestra mientras infra_logic.run() ejecuta en un hilo secundario.
     - Barra de progreso indeterminada (animada).
     - Log en tiempo real con los mensajes del pipeline.

  3. ResultsFrame (pantalla de resultados)
     - Muestra el DataFrame resultante en un Treeview con filas alternadas.
     - Información: número de registros y columnas.
     - Botón "Descargar CSV" (o xlsx) para guardar el reporte.
     - Botón "Nuevo Reporte" para volver a la pantalla de carga.

  FLUJO
  -----
  Upload → (hilo: infra_logic.run) → Loading → Results → (nuevo reporte) → Upload
  Error → regresa a Upload con mensaje de error.

  DISEÑO VISUAL
  -------------
  - Paleta oscura: fondo #0f0f1a, tarjetas #1a1a2e, acento azul #4361ee.
  - Fuente: Segoe UI. Botones con efecto hover (HoverButton).
  - Áreas de carga con efecto hover y confirmación visual (verde) al seleccionar.
  - Treeview con filas alternadas y encabezados azules.

================================================================================

  ARCHIVO: Rep_Infra.spec  (configuración de PyInstaller)
  --------------------------------------------------------

  Define cómo empaquetar la aplicación GUI en un ejecutable Windows (.exe).

  Detalles:
  - Script principal:    Rep_Infra_App.py
  - Archivos incluidos:  infra_logic.py (en el directorio raíz del exe)
  - Imports ocultos:     pandas, numpy, openpyxl
  - Modo: onedir, sin consola (windowed), compresión UPX
  - Nombre del exe:      Rep_Infra.exe

================================================================================

  ARCHIVO: build.bat  (generación automática del ejecutable)
  -----------------------------------------------------------

  Script de Windows utilizado para generar automáticamente el archivo ejecutable
  Rep_Infra.exe a partir del código fuente Python mediante PyInstaller.
  Basta con ejecutarlo (doble clic o desde la terminal) para compilar toda
  la aplicación en un exe listo para distribuir.

  Al ejecutarse:
  1. Verifica/instala PyInstaller.
  2. Llama a PyInstaller con todos los parámetros (infra_logic.py como dato
     adicional, imports ocultos: pandas, numpy, openpyxl, modo sin consola).
  3. Si tiene éxito, informa la ruta: dist\Rep_Infra\Rep_Infra.exe
  4. Indica que para distribuir se debe compartir la carpeta completa:
     dist\Rep_Infra\

================================================================================
  Dependencias: pandas, numpy, openpyxl, tkinter
  Herramienta de compilación: PyInstaller
  Python: 3.10+
================================================================================
