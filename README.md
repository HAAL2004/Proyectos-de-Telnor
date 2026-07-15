================================================================================

  PORTAFOLIO DE TRABAJO — ESTADÍA EN TELNOR / RED NOROESTE
  Periodo de desarrollo: 2026
  Autor: Heber Alejandro Ahumada López

================================================================================

DESCRIPCIÓN GENERAL
-------------------
Este repositorio reúne todos los proyectos, scripts, herramientas y aplicaciones
desarrollados durante mi estadía dentro de la empresa Telnor (Red Noroeste).

El trabajo estuvo enfocado en tres grandes líneas:

  1. Automatización de procesos de datos mediante scripts Python (pandas).
  2. Desarrollo de aplicaciones de escritorio (Tkinter + PyInstaller).
  3. Desarrollo de un sistema web interno (FastAPI + HTML/CSS) para la gestión
     y visualización de información operativa.

Cada subcarpeta corresponde al trabajo realizado en coordinación con un
colaborador o área específica del equipo.

================================================================================

  ESTRUCTURA DEL REPOSITORIO
  --------------------------

  Anti/
  ├── Christian/          → Automatización del seguimiento de programas FTTH/TBA
  ├── Ubaldo/             → Sistema de reportes de Escalaciones B (app de escritorio)
  └── Valeria/            → Generador de Reporte de Infraestructura (app de escritorio)

================================================================================

  CARPETA: Christian
  ------------------
  Trabajo realizado en conjunto con Christian para automatizar la actualización
  del archivo maestro de seguimiento de programas FTTH y TBA.

  PROBLEMA QUE RESUELVE
  ----------------------
  Cada semana se recibe un listado de nuevos registros del día ("Distritos HOY")
  que debe compararse contra el archivo maestro de seguimiento para identificar:
    - Registros que no existen aún en la base de datos.
    - Registros donde solo falta la clave SGS.
    - Cambios de etapa en registros ya existentes.
  Este proceso era manual y propenso a errores.

  SOLUCIÓN DESARROLLADA
  ---------------------
  - Fase1.py:  script con la lógica de comparación y clasificación
                 (Fase 1) y consulta de estados SGS-PON (Fase 2).
  - Fase2.py:    versión final que además escribe automáticamente los nuevos
                 registros en la hoja "BD" del archivo XLSM (preservando macros VBA)
                 y actualiza las pestañas SGS-PON con datos frescos del día.

  Archivos de entrada: Excel de distritos del día, archivo maestro XLSM,
                       archivos TSV de SGS-PON y catálogo de distritos prioritarios.
  Archivo de salida:   Seguimiento_ACTUALIZADO.xlsm con los nuevos registros
                       ya insertados y fórmulas Excel incluidas.

  → Ver: Christian/README_Christian.txt para documentación detallada.

================================================================================

  CARPETA: Ubaldo
  ---------------
  Trabajo realizado en conjunto con Ubaldo para automatizar el monitoreo diario
  de Escalaciones B dentro de la subdirección TELNOR.

  PROBLEMA QUE RESUELVE
  ----------------------
  El seguimiento de quejas escaladas (tipo B) requería revisar manualmente varios
  portales internos, cruzar la información y construir un reporte resumen por COPE
  y Área. Este proceso se realizaba a diario y consumía tiempo considerable.

  SOLUCIÓN DESARROLLADA
  ---------------------
  Aplicación de escritorio completa con interfaz gráfica oscura y moderna:

  - app.py (scraper): realiza web scraping automatizado con Selenium sobre tres
    portales internos (Escalaciones B, Quejas Hoy, Casas Cerradas), cruza la
    información y construye la tabla pivote (df_Pivote) y el reporte resumen
    por COPE (df_TEB — Tabla Escalada B).
  - Reportes.py: interfaz Tkinter con login, pantalla de carga con log en
    tiempo real, panel principal con la tabla pivote y módulo de reportes.
    Permite descargar los resultados como Excel o CSV.
  - build.bat / Reportes.spec: permiten compilar todo en un ejecutable
    Reportes.exe distribuible (sin necesidad de Python instalado).

  → Ver: Ubaldo/README_Ubaldo.txt para documentación detallada.

================================================================================

  CARPETA: Valeria
  ----------------
  Trabajo realizado en conjunto con Valeria para automatizar la generación del
  Reporte de Infraestructura de Red Noroeste.

  PROBLEMA QUE RESUELVE
  ----------------------
  Semanalmente se reciben dos archivos Excel (Semáforo de Ancho de Banda y
  Semáforo de Interconexiones) que deben procesarse y combinarse para generar
  un CSV con el estado de cada equipo de red, incluyendo: utilización, semáforo,
  clientes, interconexiones, velocidades e IPs. Este proceso requería varias
  operaciones manuales de cruce de datos.

  SOLUCIÓN DESARROLLADA
  ---------------------
  - Rep_Infra.py:    script standalone original (rutas fijas, salida directa a CSV).
  - infra_logic.py:  motor de procesamiento encapsulado en una función reutilizable
                     (run), con soporte para callbacks de progreso en tiempo real.
  - Rep_Infra_App.py: aplicación Tkinter que permite seleccionar los dos archivos
                     Excel, procesa los datos en segundo plano y muestra el resultado
                     en una tabla interactiva con opción de descarga.
  - build.bat / Rep_Infra.spec: compilación en Rep_Infra.exe distribuible.

  Reporte de salida: CSV con 24 columnas normalizadas (División, Área, Central,
  Tecnología, IPs, utilización, semáforo en texto, medio de transmisión, etc.)

  → Ver: Valeria/README_Valeria.txt para documentación detallada.

================================================================================

  RESUMEN DE TECNOLOGÍAS UTILIZADAS
  -----------------------------------

  Lenguajes:    Python 3.10+, HTML5, CSS3
  Librerías:    pandas, numpy, openpyxl, selenium, webdriver-manager,
                tkinter
  Herramientas: PyInstaller (generación de .exe), Git

  Tipos de solución entregados:
    ✔ Scripts de automatización de datos (Python + pandas)
    ✔ Aplicaciones de escritorio compiladas (.exe con Tkinter + PyInstaller)

================================================================================

  NOTAS FINALES
  -------------
  Cada carpeta contiene su propio README con la documentación técnica detallada
  de cada archivo. Este README general sirve como índice y punto de entrada
  para entender el alcance completo del trabajo realizado durante la estadía.

  Para cualquier duda sobre el código, revisar primero el README de la carpeta
  correspondiente antes de explorar los scripts directamente.

