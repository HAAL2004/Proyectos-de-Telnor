================================================================================
  README — Carpeta Christian
  Archivos documentados: Fase1.py | Fase2.py
================================================================================

DESCRIPCIÓN GENERAL
-------------------
Los scripts de esta carpeta forman parte de un proceso de actualización automática
del archivo de seguimiento de programas FTTH/TBA. Su objetivo es comparar los datos
del día ("Distritos HOY") contra el archivo maestro de seguimiento, clasificar los
registros nuevos o discrepantes, y finalmente escribir esos registros directamente
en el archivo Excel/XLSM de seguimiento.

El proceso se divide conceptualmente en dos fases:
  - Fase 1: Comparación y clasificación de registros.
  - Fase 2: Escritura en el archivo de seguimiento y actualización de pestañas SGS.

================================================================================

  ARCHIVO: Fase1.py
  --------------------

  PROPÓSITO
  ---------
  Script de prueba/prototipo. Contiene las Fases 1 y 2 del proceso, pero trabaja
  directamente con los archivos en la carpeta raíz (sin subcarpeta "PRUEBAS/").
  Se usa para validar la lógica antes de integrarla en Fase2.py.

  ARCHIVOS DE ENTRADA REQUERIDOS
  --------------------------------
  - Distritos HOY.xlsx              → Registros nuevos del día (columnas: PROG, SGS,
                                       DIST, PTOS, Pon, Terminal).
  - Seguimiento PROGRAMAS FTTH-TBA_27ABR26.xlsm  → Archivo maestro (hoja "BD").
  - SGS-PON NUEVO.xls               → Tabla TSV con estado de SGS para PON Nuevo.
  - SGS-PON EXISTENTE.xls           → Tabla TSV con estado de SGS para PON Existente.
  - CATALOGO DISTRITOS PRIORITARIOS SEM15 2026.xlsx → Catálogo de distritos prioritarios.

  FLUJO DE EJECUCIÓN
  ------------------

  [ FASE 1 — Comparación y clasificación ]

  1. Carga "Distritos HOY" (df_ND): registros nuevos del día.
     - Normaliza el campo "Terminal": reemplaza puntos por comas.
     - Rellena valores vacíos en "SGS" con 'N/A'.

  2. Carga la hoja "BD" del archivo maestro (df_TP):
     - Filtra por la semana actual (isocalendar().week).
     - Excluye registros con etapa "En Servicio".
     - Trabaja solo con columnas: SEM PROG, DISTRITO, SGS, Terminales,
       ETAPA SGS ACTUALIZADO, PUENTES.

  3. Construye tabla de comparación (df_C) para cada registro nuevo:
     - BUSCA X SGS:   busca el registro en df_TP por clave SGS.
     - B X DISTRITO:  busca el primer SGS que comparte el mismo distrito.
     - Ter x Dis:     busca terminales que coincidan con el distrito.

  4. Determina la columna "¿Igual?" para cada registro:
     - "SI"  → el registro ya existe en el maestro con los mismos datos.
     - "NO"  → existe pero con datos distintos.
     - "N/A" → no se encontró en el maestro.

  5. Filtra para quedarse solo con "NO" y "N/A" (los que necesitan acción).

  6. Aplica depuración adicional: elimina registros que ya están cubiertos
     por otro SGS del mismo distrito con la misma terminal.

  7. Clasifica los registros filtrados en dos grupos:
     - df_SGS: registros donde solo falta registrar la clave SGS
               (el distrito y terminal ya existen en la BD).
     - df_NR:  registros completamente nuevos que deben insertarse.

  8. Marca con "Prioridad" los registros cuyos distritos aparecen en el
     catálogo de distritos prioritarios.

  [ FASE 2 — Consulta de estado SGS ]

  9. Carga TablaNuevos y TablaExistente (archivos .xls tipo TSV).
     - Cruza con el catálogo de etapas (hoja "cat_eFTTH") para traducir
       la etapa a texto y marcar como "En Servicio" si orden == 99.

  10. Para cada registro en df_NR y df_SGS, busca su estado actual en
      TablaNuevos o TablaExistente según el tipo de PON (Nuevo/Exist)
      y lo escribe en la columna "ETAPA SGS ACTUALIZADO".

  11. Detecta cambios de etapa en registros ya existentes en df_TP (df_TPCE):
      registros cuya etapa en la BD maestro difiere de la etapa actual
      en los archivos SGS-PON.

  SALIDAS (impresas en consola)
  -----------------------------
  - Lista de terminales disponibles por distrito (LD).
  - Tabla completa antes y después de la clasificación.
  - df_SGS: "Registrar solamente la clave SGS".
  - df_NR:  "Nuevos Registros" con etapa actualizada.
  - df_TPCE: "Cambios de Etapa en BD".

================================================================================

  ARCHIVO: Fase2.py
  ------------------

  PROPÓSITO
  ---------
  Versión evolucionada y final del proceso. Extiende la lógica de pruebas.py con
  tres mejoras clave:
    1. Usa una variable centralizada para el nombre del archivo de seguimiento.
    2. Escribe automáticamente los nuevos registros (df_NR) en la hoja "BD"
       del archivo XLSM usando openpyxl (preservando macros VBA).
    3. Actualiza las pestañas "SGS-PON NUEVO" y "SGS-PON EXISTENTE" con datos
       frescos del día.
  El resultado se guarda como un archivo de salida separado para no sobreescribir
  el original.

  ARCHIVOS DE ENTRADA REQUERIDOS
  --------------------------------
  - PRUEBAS/Distritos HOY.XLSX
  - PRUEBAS/Seguimiento PROGRAMAS FTTH-TBA_27ABR26.xlsm   (también es salida)
  - PRUEBAS/SGS-PON NUEVO.xls
  - PRUEBAS/SGS-PON EXISTENTE.xls
  - PRUEBAS/CATALOGO DISTRITOS PRIORITARIOS SEM15 2026.xlsx

  ARCHIVO DE SALIDA
  -----------------
  - PRUEBAS/Seguimiento_ACTUALIZADO.xlsm  → Copia del archivo maestro con los
                                            cambios aplicados.

  FLUJO DE EJECUCIÓN
  ------------------

  [ FASE 1 — Comparación y clasificación ]
  Idéntica a pruebas.py (misma lógica descrita arriba).

  [ FASE 2 — Inserción en la hoja "BD" del archivo XLSM ]

  1. Abre el archivo XLSM con openpyxl (keep_vba=True para preservar macros).
  2. Lee la fila 14 de la hoja "BD" para mapear nombres de columna → índice.
  3. Inserta filas nuevas al INICIO de los datos (fila 15), empujando los
     registros existentes hacia abajo.
  4. Actualiza el rango de todas las tablas de Excel en la hoja para que
     incluyan las filas recién insertadas (evita que queden fuera del filtro).
  5. Escribe los datos de df_NR en las filas insertadas con:
     - Valores directos: PROGRAMA, SEM PROG, DISTRITO, SGS, Terminales,
       Puertos Construidos, PUENTES, PRIORIDAD.
     - Fórmulas Excel: VLOOKUP para ZONA COMERCIAL, COPE, LOCALIDAD,
       GENERICO, Prioridad etapa, y IFERROR+VLOOKUP para ETAPA SGS ACTUALIZADO.

  [ FASE 3 — Actualización de pestañas SGS-PON ]

  6. Lee los archivos .xls (formato TSV) de SGS-PON NUEVO y SGS-PON EXISTENTE.
  7. Para cada pestaña del XLSM ("SGS-PON NUEVO" y "SGS-PON EXISTENTE"):
     - Mapea los encabezados de la fila 1 a índices de columna.
     - Identifica las columnas comunes entre el archivo externo y la pestaña.
     - Sobreescribe celda a celda los datos actualizados.
     - Actualiza el rango de la tabla (Table2 / Table3) para reflejar el
       número real de filas.

  8. Guarda el resultado como PRUEBAS/Seguimiento_ACTUALIZADO.xlsm.

  DIFERENCIAS CLAVE RESPECTO A pruebas.py
  -----------------------------------------
  | Característica               | pruebas.py      | Fase2.py                  |
  |------------------------------|-----------------|---------------------------|
  | Archivos de entrada          | Raíz del proy.  | Subcarpeta PRUEBAS/       |
  | Escritura en XLSM            | No              | Sí (openpyxl + keep_vba)  |
  | Actualización pestañas SGS   | No              | Sí (Fase 3)               |
  | Archivo de salida            | Solo consola    | Seguimiento_ACTUALIZADO   |

================================================================================
  Dependencias: pandas, numpy, openpyxl
  Python: 3.10+
================================================================================
