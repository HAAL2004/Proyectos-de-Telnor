# -*- coding: utf-8 -*-
"""
infra_logic.py
==============
Lógica de procesamiento para el Reporte de Infraestructura.
Encapsula todo el pipeline de datos de Rep_Infra.py en una función reutilizable.

Uso:
    resultado = run(path_ab, path_intx, on_log=print)
    df_reporte = resultado["df_reporte"]
"""

import pandas as pd
import numpy as np


def _log(on_log, msg: str):
    """Helper para enviar mensajes de progreso."""
    if on_log:
        on_log(msg)


def run(path_ab: str, path_intx: str, on_log=None) -> dict:
    """
    Ejecuta el pipeline completo de generación del reporte de infraestructura.

    Parameters
    ----------
    path_ab : str
        Ruta al archivo Excel "Semáforo AB" (contiene hojas "2.-Reporte Infra" y "3.-Detalle").
    path_intx : str
        Ruta al archivo Excel "Semáforo Interconexión" (contiene hoja "4.-Detalle").
    on_log : callable, optional
        Callback para enviar mensajes de progreso.  Firma: on_log(msg: str)

    Returns
    -------
    dict
        {"df_reporte": pd.DataFrame}  — El DataFrame final del reporte.
    """

    # ─── Paso 1: Leer hoja "2.-Reporte Infra" ────────────────────────────────
    _log(on_log, "→ Leyendo hoja '2.-Reporte Infra'...")
    df_INFRA = pd.read_excel(path_ab, sheet_name="2.-Reporte Infra")
    df_INFRA = df_INFRA.iloc[:, :29]
    _log(on_log, f"  ✓ {len(df_INFRA)} registros cargados")

    # ─── Paso 2: Leer hoja "3.-Detalle" (Ancho de Banda) ─────────────────────
    _log(on_log, "→ Leyendo hoja '3.-Detalle' (Ancho de Banda)...")
    df_AB = pd.read_excel(path_ab, sheet_name="3.-Detalle")
    df_AB.columns = df_AB.iloc[0]
    df_AB = df_AB.iloc[1:].reset_index(drop=True)
    df_AB = df_AB[["AREA", "POBLACION", "CENTRAL", "TECNOLOGIA", "EQUIPO", "Nombre configurado en equipo", "Enlace", "INTERFACE EQUIPO",
                   "KB x cliente", "Avg Util (%)", "Tipo de Dato", "Puertos Totales", "Puertos Ocupados", "SEMAFORO ACTUAL"]]
    _log(on_log, f"  ✓ {len(df_AB)} registros cargados")

    # ─── Paso 3: Leer hoja "4.-Detalle" (Interconexiones) ────────────────────
    _log(on_log, "→ Leyendo hoja '4.-Detalle' (Interconexiones)...")
    df_INTXS = pd.read_excel(path_intx, sheet_name="4.-Detalle")
    df_INTXS.columns = df_INTXS.iloc[0]
    df_INTXS = df_INTXS.iloc[1:].reset_index(drop=True)
    df_INTXS = df_INTXS[["BAS_INTERFAZ", "Nombre de ITX", "Speed (Kbps)", "Prom. de Utilizacion (%)", "No. de Equipos (en interfaz)",
                         "Clientes en Interconexión (Ocupados)", "Total Puertos", "SEMAFORO"]]
    _log(on_log, f"  ✓ {len(df_INTXS)} registros cargados")

    # ─── Paso 4: Construir df_SEM1 (intermedio) ──────────────────────────────
    _log(on_log, "→ Construyendo tabla intermedia (Semáforo Infra)...")

    df_SEM1 = pd.DataFrame(columns=["AREA", "POBLACION", "CENTRAL", "TECNOLOGIA", "NombrePisa", "Nombre configurado en equipo", "AB REI", "Avg Util (%)", "Peak Util (%)", "TIPO de DATO", "TOTAL", "OCUPADOS", "SEMAFORO ENLACE", "BAS", "Subinterfaz", "Interconexión", "Avg Util (%) INTERCONEXIÓN", "SEMAFORO INTERCONEXIÓN"])

    df_SEM1["AREA"] = df_INFRA["AREA"]
    df_SEM1["POBLACION"] = df_INFRA["POBLACION"]
    df_SEM1["CENTRAL"] = df_INFRA["CENTRAL"]
    df_SEM1["TECNOLOGIA"] = df_INFRA["TECNOLOGIA"]
    df_SEM1["NombrePisa"] = df_INFRA["EQUIPO"]
    df_SEM1["Nombre configurado en equipo"] = df_INFRA["Nombre configurado en equipo"]

    df_infra_lookup = df_INFRA[["EQUIPO", "INTERFACE EQUIPO", "TOTAL", "OCUPADO (I)", "NOMBRE BAS", "SUBINT BAS"]].drop_duplicates(subset="EQUIPO", keep="first")
    df_SEM1 = df_SEM1.merge(df_infra_lookup, left_on="NombrePisa", right_on="EQUIPO", how="left")
    df_SEM1["TOTAL_x"] = df_SEM1["TOTAL_y"].fillna("0").astype(int)
    df_SEM1["OCUPADOS"] = df_SEM1["OCUPADO (I)"].fillna("0").astype(int)
    df_SEM1.rename(columns={'TOTAL_x': 'TOTAL'}, inplace=True)
    df_SEM1["AB REI"] = df_SEM1["INTERFACE EQUIPO"].fillna("")
    df_SEM1["BAS"] = df_SEM1["NOMBRE BAS"].fillna("")
    df_SEM1["Subinterfaz"] = df_SEM1["SUBINT BAS"].fillna("")
    df_SEM1.drop(columns=["INTERFACE EQUIPO", "EQUIPO", "TOTAL_y", "OCUPADO (I)", "NOMBRE BAS", "SUBINT BAS"], inplace=True)

    df_ab_lookup = df_AB[["EQUIPO", "Avg Util (%)", "Tipo de Dato", "SEMAFORO ACTUAL"]].drop_duplicates(subset="EQUIPO", keep="first")
    df_SEM1 = df_SEM1.merge(df_ab_lookup, left_on="NombrePisa", right_on="EQUIPO", how="left")
    df_SEM1["Avg Util (%)_x"] = df_SEM1["Avg Util (%)_y"]
    df_SEM1.rename(columns={'Avg Util (%)_x': 'Avg Util (%)'}, inplace=True)
    df_SEM1["TIPO de DATO"] = df_SEM1["Tipo de Dato"].fillna("")
    df_SEM1["SEMAFORO ENLACE"] = df_SEM1["SEMAFORO ACTUAL"].fillna("")
    df_SEM1.drop(columns=["Avg Util (%)_y", "EQUIPO", "Tipo de Dato", "SEMAFORO ACTUAL"], inplace=True)

    df_SEM1["Peak Util (%)"] = 0

    df_SEM1["Interconexión"] = df_SEM1["BAS"].str.cat(df_SEM1["Subinterfaz"], sep="_", na_rep="")

    df_intx_lookup = df_INTXS[["BAS_INTERFAZ", "Prom. de Utilizacion (%)", "SEMAFORO"]].drop_duplicates(subset="BAS_INTERFAZ", keep="first")
    df_SEM1 = df_SEM1.merge(df_intx_lookup, left_on="Interconexión", right_on="BAS_INTERFAZ", how="left")
    df_SEM1["Avg Util (%) INTERCONEXIÓN"] = pd.to_numeric(df_SEM1["Prom. de Utilizacion (%)"], errors='coerce').round(2).fillna(0)
    df_SEM1["SEMAFORO INTERCONEXIÓN"] = df_SEM1["SEMAFORO"].fillna("")
    df_SEM1.drop(columns=["BAS_INTERFAZ", "Prom. de Utilizacion (%)", "SEMAFORO"], inplace=True)

    df_SEM1.rename(columns={'POBLACION': 'Catálogo TECNOLOGÍA', 'NombrePisa': 'EQUIPO'}, inplace=True)
    _log(on_log, f"  ✓ Tabla intermedia: {len(df_SEM1)} registros")

    # ─── Paso 5: Construir df_SEM2 (reporte final) ──────────────────────────
    _log(on_log, "→ Construyendo reporte final...")

    df_SEM2 = pd.DataFrame(columns=["Division", "Area", "Siglas_central", "NombrePisa", "ip_nom_pisa", "Tecnologia",
                                "equipo_410", "medio_tx", "clientes", "nombre_configurado", "nombre_router", "ip_nombre_router",
                                "interface", "subinterface", "velocidad", "promedio_util_salida_interfaz INTX", "semaforo", "color",
                                "nombre_subinterace", "VLANHSI", "velocidad_subint", "promedio_util_salida_subinterfaz ACCESO",
                                "Prom_util_salida_subinterfaz_kbps", "porcentaje_ocupacion INTX"])

    df_SEM2["Area"] = df_INFRA["AREA"]
    df_SEM2["Siglas_central"] = df_INFRA["CENTRAL"]
    df_SEM2["NombrePisa"] = df_INFRA["EQUIPO"]
    df_SEM2["Tecnologia"] = df_INFRA["TECNOLOGIA"]
    df_SEM2["Division"] = "RED NOROESTE"
    df_SEM2["subinterface"] = "S.D"
    df_SEM2["velocidad_subint"] = 1000000000

    df_infra_lookup = df_INFRA[["EQUIPO", "IP EQUIPO", "NOM CONC TBA", "INTERFACE EQUIPO", "VPI/VLAN EQUIPO"]].drop_duplicates(subset="EQUIPO", keep="first")
    df_SEM2 = df_SEM2.merge(df_infra_lookup, left_on="NombrePisa", right_on="EQUIPO", how="left")
    df_SEM2["ip_nom_pisa"] = df_SEM2["IP EQUIPO"]
    df_SEM2["equipo_410"] = df_SEM2["NOM CONC TBA"]
    df_SEM2["medio_tx"] = df_SEM2["INTERFACE EQUIPO"].astype(str)
    df_SEM2["medio_tx"] = df_SEM2["medio_tx"].str.replace("GE", "GB", regex=False).str.replace("CE", "", regex=False)
    df_SEM2["VLANHSI"] = df_SEM2["VPI/VLAN EQUIPO"]
    df_SEM2.drop(columns=["EQUIPO", "IP EQUIPO", "NOM CONC TBA", "INTERFACE EQUIPO", "VPI/VLAN EQUIPO"], inplace=True)

    _log(on_log, "  → Aplicando merges con tabla intermedia...")

    df_int_lookup = df_SEM1[["EQUIPO", "OCUPADOS", "BAS", "Subinterfaz", "Avg Util (%) INTERCONEXIÓN", "SEMAFORO ENLACE", "Interconexión",
                             "Avg Util (%)"]].drop_duplicates(subset="EQUIPO", keep="first")
    df_SEM2 = df_SEM2.merge(df_int_lookup, left_on="NombrePisa", right_on="EQUIPO", how="left")
    df_SEM2["clientes"] = df_SEM2["OCUPADOS"]
    df_SEM2["nombre_router"] = df_SEM2["BAS"]
    df_SEM2["interface"] = df_SEM2["Subinterfaz"]
    df_SEM2["promedio_util_salida_interfaz INTX"] = df_SEM2["Avg Util (%) INTERCONEXIÓN"]
    df_SEM2["semaforo"] = df_SEM2["SEMAFORO ENLACE"]
    df_SEM2["nombre_subinterace"] = df_SEM2["Interconexión"]
    df_SEM2["promedio_util_salida_subinterfaz ACCESO"] = pd.to_numeric(df_SEM2["Avg Util (%)"], errors='coerce').fillna(0).round(2)
    df_SEM2["Prom_util_salida_subinterfaz_kbps"] = df_SEM2["Avg Util (%)"] * 10000000
    df_SEM2.drop(columns=["EQUIPO", "OCUPADOS", "BAS", "Subinterfaz", "Avg Util (%) INTERCONEXIÓN", "SEMAFORO ENLACE", "Interconexión",
                          "Avg Util (%)"], inplace=True)

    df_SEM2["porcentaje_ocupacion INTX"] = df_SEM2["promedio_util_salida_interfaz INTX"]

    _log(on_log, "  → Aplicando merge IP BAS...")

    df_infra_lookup = df_INFRA[["NOMBRE BAS", "IP BAS"]].drop_duplicates(subset="NOMBRE BAS", keep="first")
    df_SEM2 = df_SEM2.merge(df_infra_lookup, left_on="nombre_router", right_on="NOMBRE BAS", how="left")
    df_SEM2["ip_nombre_router"] = df_SEM2["IP BAS"]
    df_SEM2.drop(columns=["NOMBRE BAS", "IP BAS"], inplace=True)

    df_ab_lookup = df_AB[["EQUIPO", "KB x cliente"]].drop_duplicates(subset="EQUIPO", keep="first")
    df_SEM2 = df_SEM2.merge(df_ab_lookup, left_on="NombrePisa", right_on="EQUIPO", how="left")
    df_SEM2["velocidad"] = df_SEM2["KB x cliente"].fillna("")
    df_SEM2.drop(columns=["EQUIPO", "KB x cliente"], inplace=True)

    _log(on_log, "  → Aplicando transformaciones finales...")

    df_SEM2["nombre_configurado"] = np.where(
        df_SEM2["equipo_410"] == "-",
        df_SEM2["NombrePisa"],
        df_SEM2["equipo_410"]
    )

    semaforo_map = {"V": "VERDE", "A": "AMARILLO", "R": "ROJO", "S": "SATURADO"}
    df_SEM2["color"] = df_SEM2["semaforo"].map(semaforo_map)

    df_SEM2["medio_tx"] = df_SEM2["medio_tx"].replace({"1 GB": "1 GB 1000", "10 GB": "1 GB 10000", "1 GB ": "1 GB 1000", "10 GB ": "1 GB 10000"})

    mask_gpon = df_SEM2["Tecnologia"].str.contains("GPON", na=False) & (df_SEM2["medio_tx"] == "1 GB 1000")
    df_SEM2.loc[mask_gpon, "medio_tx"] = "1 GB 2000"

    mask_gpon = df_SEM2["Tecnologia"].str.contains("GPON", na=False) & (df_SEM2["medio_tx"] == "1 GB 1000 ")
    df_SEM2.loc[mask_gpon, "medio_tx"] = "1 GB 2000"

    _log(on_log, f"  ✓ Reporte final generado: {len(df_SEM2)} registros, {len(df_SEM2.columns)} columnas")

    return {"df_reporte": df_SEM2}
