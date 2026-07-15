================================================================================
  README — Carpeta Ubaldo
  Archivos documentados: Reportes.py | app.py | Reportes.spec | build.bat
================================================================================

DESCRIPCIÓN GENERAL
-------------------
Esta carpeta contiene una aplicación de escritorio para el monitoreo y reporte
de Escalaciones B de la subdirección TELNOR. El sistema realiza web scraping
automatizado sobre portales internos para obtener información de quejas, visitas
y casas cerradas del día, construye una tabla pivote con métricas de atención y
genera un reporte resumen por COPE y Área. Todo esto se presenta en una interfaz
gráfica oscura y moderna construida con Tkinter.

La aplicación se compone de dos módulos:
  - app.py (scraper): lógica de extracción y procesamiento de datos.
  - Reportes.py:      interfaz gráfica y orquestador del flujo completo.

================================================================================

  ARCHIVO: app.py  (módulo de scraping — importado como "scraper")
  -----------------------------------------------------------------

  PROPÓSITO
  ---------
  Extrae datos de tres portales internos vía Selenium y construye:
    - df_Pivote: tabla detallada por queja (estado, técnico, cita, área).
    - df_TEB:    tabla resumen Escalada B con conteos por COPE y Área.

  PORTALES CONSULTADOS
  --------------------
  1. Escaladas B: descarga CSV con quejas tipo "B" de TELNOR.
     Columnas usadas: QUEJA, citaPeriodo, QuejaRNUM, entre otras.

  2. Quejas Hoy: requiere login. Descarga detalle de PLEX del día.
     Columnas clave: QUEJA_AZUL, ESTATUS_PLEX, CALIF_PLEX, FECHA_CITA_PLEX,
                     DISTRITO, TAREA, EMPLEADO_PLEX.

  3. Casas Cerradas: presiona "Consultar" y descarga CSV.
     Columnas clave: FOLIO, CALIFICACION.

  ARCHIVO AUXILIAR
  ----------------
  - ESCALADAS 2026_03_20.xlsx (hoja "SIGLAS COPE"):
    Catálogo siglas de distrito → COPE → ÁREA. Clasifica cada queja
    dentro de su área geográfica (Ensenada, Mexicali, Tijuana).

  LÓGICA PRINCIPAL — df_Pivote
  ----------------------------
  Se construye fila a fila combinando los tres datasets. Los campos clave:
  - ESTADO: determinado por una cadena de condiciones que evalúa CALIF_PLEX,
    FUTURA, ESTATUS_PLEX, TAREA_PLEX y si la casa fue cerrada hoy.
    Valores posibles: COMPLETADA, RE-AGENDA FUTURA, CASA_CERRADA,
                      MANTENIMIENTO, PENDIENTE, ASIGNADA.
  - ESC B FECHA: clasifica la escalada según citaPeriodo
    ("Hoy" / "Manana" / otro).
  - COMPLETADA_CITA FUTURA: "ATENDIDA/FUTURA" o "NO ATENDIDA" según ESTADO.
  - AREA_ACT / COPE_ACT: cruce por las 3 primeras letras del DISTRITO.

  LÓGICA PRINCIPAL — df_TEB (Tabla Escalada B)
  --------------------------------------------
  Tabla resumen con una fila por COPE_ACT y filas de subtotales por área.
  Cada queja incrementa el contador de su estado en la fila del COPE
  y en el subtotal de su ÁREA. Al final se ordena y exporta a EscalaB.xlsx.

  SALIDAS
  -------
  - descargas/          → CSVs descargados (se limpian al inicio).
  - Pivote_EscalaB.csv  → Tabla pivote completa.
  - EscalaB.xlsx        → Tabla resumen Escalada B.

================================================================================

  ARCHIVO: Reportes.py  (interfaz gráfica — punto de entrada)
  ------------------------------------------------------------

  PROPÓSITO
  ---------
  Aplicación Tkinter con diseño oscuro y moderno. Orquesta: login → scraping
  en hilo secundario → visualización de resultados. Importa app.py (como scraper).

  PANTALLAS
  ---------
  1. LoginFrame   — Solicita usuario y contraseña. Valida campos no vacíos.
  2. LoadingFrame — Barra de progreso animada + log en tiempo real del scraping.
  3. HomeFrame    — Panel principal con tabla df_Pivote, botón de descarga Excel
                    y acceso al módulo Reportes. Botón "Actualizar datos".
  4. ReportesFrame — Muestra df_TEB, botón de descarga Excel/CSV, botón "Inicio".

  FLUJO
  -----
  Login → (hilo de scraping) → Loading → Home ↔ Reportes
  Error → regresa a Login con mensaje.

  DISEÑO VISUAL
  -------------
  - Paleta oscura: fondo #0f0f1a, tarjetas #1a1a2e, acento azul #4361ee.
  - Fuente: Segoe UI. Botones con efecto hover (HoverButton).
  - Treeview con filas alternadas y encabezados azules.

================================================================================

  ARCHIVO: Reportes.spec  (configuración de PyInstaller)
  -------------------------------------------------------

  Define cómo empaquetar la aplicación en un ejecutable Windows (.exe).

  Detalles:
  - Script principal:    Reportes.py
  - Archivos incluidos:  scraper.py, ESCALADAS 2026_03_20.xlsx
  - Imports ocultos:     pandas, openpyxl
  - Dependencias auto:   selenium, webdriver_manager (collect_all)
  - Modo: onedir, sin consola (windowed), compresión UPX
  - Nombre del exe:      Reportes.exe

================================================================================

  ARCHIVO: build.bat  (generación automática del ejecutable)
  -----------------------------------------------------------

  Script de Windows utilizado para generar automáticamente el archivo ejecutable
  Reportes.exe a partir del código fuente Python mediante PyInstaller.
  Basta con ejecutarlo (doble clic o desde la terminal) para compilar toda
  la aplicación en un exe listo para distribuir.

  Al ejecutarse:
  1. Verifica/instala PyInstaller.
  2. Llama a PyInstaller con todos los parámetros (archivos adicionales,
     imports ocultos, modo ventana sin consola).
  3. Si tiene éxito, informa la ruta: dist\Reportes\Reportes.exe
  4. Indica que para distribuir se debe compartir la carpeta completa:
     dist\Reportes\

================================================================================
  Dependencias: pandas, selenium, webdriver-manager, openpyxl, tkinter
  Herramienta de compilación: PyInstaller
  Python: 3.10+
================================================================================
