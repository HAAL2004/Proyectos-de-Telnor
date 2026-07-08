import pandas as pd
import re
import os
import time
import glob
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from datetime import date




print(f"\n{'='*75}")
print(f"  Web Scraper - Escalado B")
print(f"{'='*75}\n")

# ============================================================
#  Links a las paginas

URL_EscalaB = "https://reportesplanta.intranet.telmex.com/testing/Operacion/New/SeguimientoAEscalaciones/Dashboard.asp"
URL_Qhoy = "https://mioficina.intranet.rednoroeste.com/group/monitoreo-de-os-y-gestion/detalle-plex-abiertas-cerradas-hoy?p_auth=Z5g1ifbi&p_p_id=menufecha_WAR_Utilerias2019&p_p_lifecycle=1&p_p_state=normal&p_p_mode=view&p_p_col_id=column-1&p_p_col_count=3&_menufecha_WAR_Utilerias2019_javax.portlet.action=ejecutaAccion"
URL_CCerradas = "https://mioficina.intranet.rednoroeste.com/group/monitoreo-de-os-y-gestion/casas-cerradas"

# Carpeta donde Selenium guardará las descargas
Descargas = os.path.join(os.path.dirname(os.path.abspath(__file__)), "descargas")
os.makedirs(Descargas, exist_ok=True)


# Limpiar archivos previos en la carpeta de descargas
for f in glob.glob(os.path.join(Descargas, "*.*")):
    os.remove(f)

# ============================================================
#  ESCALA B 

# Configuración de Selenium
print("→ Abriendo navegador...")
s = Service(ChromeDriverManager().install())
opc = Options()
opc.add_argument("--window-size=1020,1200")

prefs = {
    "download.default_directory": Descargas,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
}
opc.add_experimental_option("prefs", prefs)

navegador = webdriver.Chrome(service=s, options=opc)

df_EscalaB = pd.DataFrame()  

# Extracion de datos

try:
    print(f"→ Navegando a: {URL_EscalaB}")
    navegador.get(URL_EscalaB)
    time.sleep(5)

    # Buscar y hacer clic en el enlace del CSV de TELNOR
    print("→ Buscando enlace de exportación CSV (TELNOR)...")
    enlace = navegador.find_element(
        By.CSS_SELECTOR,
        'a[href*="ExportCSVEscalacionResumenEjecutivo.asp?caso=1"]'
    )
    print(f"  Enlace encontrado: {enlace.text.strip()}")
    enlace.click()

    print("→ Descargando archivo...")
    print("  Esperando descarga...", end="")
    descarga_ok = False

    for _ in range(60):
        time.sleep(1)
        archivos_temp = glob.glob(os.path.join(Descargas, "*.crdownload"))
        archivos = glob.glob(os.path.join(Descargas, "*.csv")) + glob.glob(os.path.join(Descargas, "*.xls*"))
        if not archivos_temp and archivos:
            print(" ✓")
            descarga_ok = True
            break

    if not descarga_ok:
        print("\n⚠  No se descargó ningún archivo (timeout).")

    else:
        archivo = max(archivos, key=os.path.getmtime)
        print(f"\n→ Procesando: {os.path.basename(archivo)}")

        ext = os.path.splitext(archivo)[1].lower()
        print(f"  Archivo detectado: {os.path.basename(archivo)} ({ext})")

        df_EscalaB = pd.read_csv(archivo, encoding="latin-1", usecols=["QUEJA", "ETIQUETA_QUEJA", "Distrito", "TMXSubDireccion", "tieneCitaAtendida", "citaPeriodo", "QuejaRNUM"])

        df_EscalaB['TMXSubDireccion'] = df_EscalaB['TMXSubDireccion'].astype(str).str.strip()
        df_EscalaB['ETIQUETA_QUEJA'] = df_EscalaB['ETIQUETA_QUEJA'].astype(str).str.strip()
        df_EscalaB['QuejaRNUM'] = pd.to_numeric(df_EscalaB['QuejaRNUM'], errors='coerce')
        df_EscalaB['QuejaRNUM'] = df_EscalaB['QuejaRNUM'].astype('Int64')

        df_EscalaB = df_EscalaB[df_EscalaB['TMXSubDireccion'].isin(["TELNOR"])].reset_index(drop=True)
        df_EscalaB = df_EscalaB[df_EscalaB['ETIQUETA_QUEJA'].isin(["B"])].reset_index(drop=True)



        if df_EscalaB.empty:
            print("\n⚠  El archivo está vacío.")
        else:
            print("\n→ Datos de EscalaB:")
            print(f"  Se obtuvieron {len(df_EscalaB)} registros y {len(df_EscalaB.columns)} columnas:\n")
            print(df_EscalaB.head())

except Exception as e:
    print(f"\n✖  Error en EscalaB: {e}")

finally:
    pass  # No cerrar navegador, se reutiliza para descargar Qhoy

# =========================================================================================

# Credenciales (modifica con tus datos)
USUARIO = "ubfl@rednoroeste.com"
CONTRASENA = "MiN94729$%&%$#jhgjh"

# =========================================================================================

# Quejas abiertas y cerradas hoy (descarga via Selenium)

df_Qhoy = pd.DataFrame()

try:
    print(f"\n{'='*75}")
    print(f"→ Navegando a: {URL_Qhoy}")
    navegador.get(URL_Qhoy)
    time.sleep(5)

    # Login
    print("→ Ingresando credenciales...")
    campo_usuario = navegador.find_element(By.ID, "_58_login")
    campo_contrasena = navegador.find_element(By.ID, "_58_password")
    campo_usuario.clear()
    campo_usuario.send_keys(USUARIO)
    campo_contrasena.clear()
    campo_contrasena.send_keys(CONTRASENA)

    boton_login = navegador.find_element(By.CSS_SELECTOR, 'input[type="submit"], button[type="submit"]')
    boton_login.click()
    print("  Login enviado, esperando carga...")
    time.sleep(5)

    # Navegar de nuevo a la URL de descarga (por si el login redirige)
    navegador.get(URL_Qhoy)
    time.sleep(5)

    # Guardar lista de archivos existentes ANTES del click
    archivos_previos = set(glob.glob(os.path.join(Descargas, "*.csv")) + glob.glob(os.path.join(Descargas, "*.xls*")))

    # Buscar y hacer clic en "Descarga CSV"
    print("→ Buscando botón 'Descarga CSV'...")
    boton_csv = navegador.find_element(By.CSS_SELECTOR, 'input[type="submit"][value="Descarga CSV"]')
    boton_csv.click()

    print("→ Descargando archivo...")
    print("  Esperando descarga...", end="")
    descarga_ok = False

    for _ in range(60):
        time.sleep(1)
        archivos_temp = glob.glob(os.path.join(Descargas, "*.crdownload"))
        archivos_actuales = set(glob.glob(os.path.join(Descargas, "*.csv")) + glob.glob(os.path.join(Descargas, "*.xls*")))
        archivos_nuevos = archivos_actuales - archivos_previos  # Solo archivos NUEVOS
        if not archivos_temp and archivos_nuevos:
            print(" ✓")
            descarga_ok = True
            break

    if not descarga_ok:
        print("\n⚠  No se descargó ningún archivo (timeout).")
    else:
        archivo_qhoy = max(archivos_nuevos, key=os.path.getmtime)
        print(f"\n→ Procesando: {os.path.basename(archivo_qhoy)}")

        df_Qhoy = pd.read_csv(archivo_qhoy, encoding='latin-1', engine='python')
        df_Qhoy["FECHA_CITA_PLEX"] = pd.to_datetime(df_Qhoy["FECHA_CITA_PLEX"])


        print("→ Extrayendo datos de Quejas abiertas y cerradas hoy...")
        print(f"  Se obtuvieron {len(df_Qhoy)} registros y {len(df_Qhoy.columns)} columnas:\n")
        print(df_Qhoy.head())

except Exception as e:
    print(f"\n✖  Error en Quejas Hoy: {e}")

finally:
    pass

fecha = date.today()

print(f"\n{'='*75}")

# =========================================================================================

# Casas Cerradas
df_CCerradas = pd.DataFrame()

try:
    print(f"\n{'='*75}")
    print(f"→ Navegando a: {URL_CCerradas}")
    navegador.get(URL_CCerradas)
    time.sleep(5)

    # Click en "Consultar" para cargar los datos del día
    print("→ Buscando botón 'Consultar'...")
    boton_consultar = navegador.find_element(By.CSS_SELECTOR, 'button[type="submit"]')
    boton_consultar.click()
    print("  Consulta enviada, esperando carga...")
    time.sleep(5)

    # Guardar lista de archivos existentes ANTES del click
    archivos_previos_cc = set(glob.glob(os.path.join(Descargas, "*.csv")) + glob.glob(os.path.join(Descargas, "*.xls*")))

    # Buscar y hacer clic en "Descarga CSV"
    print("→ Buscando botón 'Descarga CSV'...")
    boton_csv_cc = navegador.find_element(By.CSS_SELECTOR, 'input[type="submit"][value="Descarga CSV"]')
    boton_csv_cc.click()

    print("→ Descargando archivo...")
    print("  Esperando descarga...", end="")
    descarga_ok = False

    for _ in range(60):
        time.sleep(1)
        archivos_temp = glob.glob(os.path.join(Descargas, "*.crdownload"))
        archivos_actuales = set(glob.glob(os.path.join(Descargas, "*.csv")) + glob.glob(os.path.join(Descargas, "*.xls*")))
        archivos_nuevos = archivos_actuales - archivos_previos_cc
        if not archivos_temp and archivos_nuevos:
            print(" ✓")
            descarga_ok = True
            break

    if not descarga_ok:
        print("\n⚠  No se descargó ningún archivo (timeout).")
    else:
        archivo_cc = max(archivos_nuevos, key=os.path.getmtime)
        print(f"\n→ Procesando: {os.path.basename(archivo_cc)}")

        df_CCerradas = pd.read_csv(archivo_cc, encoding='latin-1', engine='python')

        print("→ Extrayendo datos de Casas Cerradas...")
        print(f"  Se obtuvieron {len(df_CCerradas)} registros y {len(df_CCerradas.columns)} columnas:\n")
        print(df_CCerradas.head())

except Exception as e:
    print(f"\n✖  Error en Casas Cerradas: {e}")

finally:
    navegador.quit()

fecha = date.today()

print(f"\n{'='*75}")

# =========================================================================================

# Siglas COPE

df_SC = pd.read_excel("ESCALADAS 2026_03_20.xlsx", sheet_name="SIGLAS COPE", usecols=["siglas", "cope", "AREA"])

print("→ Extrayendo datos de SIGLAS COPE...")
print(f"  Se obtuvieron {len(df_SC)} registros y {len(df_SC.columns)} columnas:\n")

# =========================================================================================

print("  Recoleccion de informacion finalizada")
print(f"{'='*75}\n")

# =========================================================================================

# Manejo de datos

print("Base Pivote\n")

df_Pivote = pd.DataFrame(columns= ["QUEJA TNO", "QUEJA RUMN", "DISTRITO", "ESTATUS_PLEX", "CALIF_PLEX", "FUTURA", "ESTADO", "TAREA_PLEX", 
                                    "CITA PLEX", "citaPeriodo", "ESC B FECHA", "AREA_ACT", "COPE_ACT", "TECNICO", 
                                    "COMPLETADA_CITA FUTURA", "CASA CERRADA HOY"])


# A) QUEJA TNO
"""='BASE QUEJAS'!B2"""
""" BASE QUEJAS - B2 = QUEJA_AZUL"""
df_Pivote["QUEJA TNO"] = df_Qhoy["QUEJA_AZUL"]

# --------------------------------------------

# B) QUEJA RUMN
"""='BASE QUEJAS'!T2"""
""" BASE QUEJAS - T2 = QJ_PLEX"""
df_Pivote["QUEJA RUMN"] = pd.to_numeric(df_Qhoy["QJ_PLEX"], errors='coerce').astype('Int64')

# -----------------------------------------

# F) ESTATUS_PLEX
"""='BASE QUEJAS'!X2"""
""" BASE QUEJAS - X2 = ESTATUS_PLEX"""
df_Pivote["ESTATUS_PLEX"] = df_Qhoy["ESTATUS_PLEX"]

# ----------------------------------------

# K) TAREA_PLEX
"""='BASE QUEJAS'!Z2"""
""" BASE QUEJAS - Z2 = TAREA"""
df_Pivote["TAREA_PLEX"] = df_Qhoy["TAREA "]

# ----------------------------------------

# P) citaPeriodo
"""=IFNA(VLOOKUP(A2,'BASE PORTAL'!H:BE,50,0),"")
                 A2 = QUEJA TNO   
                    BASE PORTAL - H = QUEJA
                      BASE PORTAL - BE = citaPeriodo"""
df_escala_lookup = df_EscalaB[["QUEJA", "citaPeriodo"]].drop_duplicates(subset="QUEJA", keep="first")
df_Pivote = df_Pivote.merge(df_escala_lookup, left_on="QUEJA TNO", right_on="QUEJA", how="left")
df_Pivote["citaPeriodo"] = df_Pivote["citaPeriodo_y"].fillna("")
df_Pivote.drop(columns=["QUEJA", "citaPeriodo_x", "citaPeriodo_y"], inplace=True)

# ----------------------------------------------------------------------------------------------------

# AF) TECNICO
"""=IFNA(VLOOKUP(A8,'BASE QUEJAS'!B:AB,27,0),"")
                 A8 = QUEJA TNO   
                    BASE QUEJAS - B = QUEJA_AZUL
                      BASE QUEJAS - AB = EMPLEADO_PLEX"""
df_escala_lookup = df_Qhoy[["QUEJA_AZUL", "EMPLEADO_PLEX"]].drop_duplicates(subset="QUEJA_AZUL", keep="first")
df_Pivote = df_Pivote.merge(df_escala_lookup, left_on="QUEJA TNO", right_on="QUEJA_AZUL", how="left")
df_Pivote["TECNICO"] = df_Pivote["EMPLEADO_PLEX"].fillna("")
df_Pivote.drop(columns=["QUEJA_AZUL", "EMPLEADO_PLEX"], inplace=True)

# ----------------------------------------------------------------------------------------------------

# AH) CASA CERRADA HOY
"""=IFNA(VLOOKUP(B8,'CASA CERRADA HOY'!A:C,3,0),"")
                 B8 = QUEJA RUMN
                    CASA CERRADA HOY - A = FOLIO
                      CASA CERRADA HOY - C = CALIFICACION
"""
df_CCH_lookup = df_CCerradas[["FOLIO", "CALIFICACION"]].drop_duplicates(subset="FOLIO", keep="first")
df_Pivote = df_Pivote.merge(df_CCH_lookup, left_on="QUEJA RUMN", right_on="FOLIO", how="left")
df_Pivote.drop(columns=["FOLIO", "CALIFICACION"], inplace=True)


# ---------------------------------------------------------------------------------------------------------------------------------------------------------


for i in range(len(df_Qhoy)):

    # D) DISTRITO
    """=IF(AI2="",'BASE QUEJAS'!AC2,AI2)"""
    """    AI2 = DIST MANUAL   
                  BASE QUEJAS - AC2 = DISTRITO"""
    df_Pivote.loc[i, "DISTRITO"] = df_Qhoy.loc[i, "DISTRITO"]

    # ----------------------------------------------------------------

    # G) CALIF_PLEX
    """=IF(OR('BASE QUEJAS'!Y2=0,'BASE QUEJAS'!Y2=""),"",'BASE QUEJAS'!Y2)"""
    """       BASE QUEJAS - Y2 = CALIF_PLEX"""
    CPlex = df_Qhoy.loc[i, "CALIF_PLEX"]
    if CPlex == "" or CPlex == 0:
        df_Pivote.loc[i, "CALIF_PLEX"] = ""
    else:
        df_Pivote.loc[i, "CALIF_PLEX"] = CPlex

    # -------------------------------------------------------------------------

    # M) CITA PLEX
    """=IFERROR(IF(OR(YEAR('BASE QUEJAS'!AD2)=1940,YEAR('BASE QUEJAS'!AD2)=1900),"",'BASE QUEJAS'!AD2),"")"""
    """                    BASE QUEJAS - AD2 = FECHA_CITA_PLEX"""
    if pd.isna(df_Qhoy.loc[i, "FECHA_CITA_PLEX"]):
        df_Pivote.loc[i, "CITA PLEX"] = ""
    elif df_Qhoy.loc[i, "FECHA_CITA_PLEX"].year == 1940 or df_Qhoy.loc[i, "FECHA_CITA_PLEX"].year == 1900:
        df_Pivote.loc[i, "CITA PLEX"] = ""
    else:
        df_Pivote.loc[i, "CITA PLEX"] = df_Qhoy.loc[i, "FECHA_CITA_PLEX"].date()

    # ---------------------------------------------------------------------------------------------------------

    # Q-R-S) ESC B FECHA
    """=IF(P2="Hoy","1 Escalada B cita Hoy","")"""
    """=IF(P2="Manana","Escalada B MAÑANA","")"""
    """=IF(P2="Hoy","",IF(P2="","","5 Escalada B no para hoy"))"""
    """          Pivote - P2 = citaPeriodo"""
    cp = df_Pivote.loc[i, "citaPeriodo"]
    if cp == "Hoy":
        df_Pivote.loc[i, "ESC B FECHA"] = "1 Escalada B cita Hoy"
    elif cp == "Manana":
        df_Pivote.loc[i, "ESC B FECHA"] = "Escalada B MAÑANA"
    elif cp != "":
        df_Pivote.loc[i, "ESC B FECHA"] = "5 Escalada B no para hoy"
    else:
        df_Pivote.loc[i, "ESC B FECHA"] = ""

    # --------------------------------------------------------------

    # H) FUTURA
    """=IF(M2="","",IF(M2>TODAY(),"RE-AGENDA FUTURA",""))"""
    """       Pivote - M2 = CITA PLEX"""
    CP = df_Pivote.loc[i, "CITA PLEX"]
    if CP == "":
        df_Pivote.loc[i, "FUTURA"] = ""
    elif CP > fecha:
        df_Pivote.loc[i, "FUTURA"] = "RE-AGENDA FUTURA"
    else:
        df_Pivote.loc[i, "FUTURA"] = ""

    # --------------------------------------------------------------


    # I-J) ESTADO
    """ =IF(G8="COMPLETADA",G8, IF
            G8 = CALIF_PLEX
            (H8="RE-AGENDA FUTURA",H8, IF
             H8 = FUTURA
                (TRIM(G8)="", IF
                    (AND(F8="PENDIENTE",AH8="CASA_CERRADA"), AH8, IF
                        F8 = ESTATUS_PLEX
                                        AH8 = CASA CERRADA HOY
                        (AND(F8="PENDIENTE",OR(K8="F841",K8="F597")),"MANTENIMIENTO",F8)),TRIM(G8)))) 
                                               K8 = TAREA_PLEX
                                                                                                
    """
    if df_Pivote.loc[i, "FUTURA"] == "RE-AGENDA FUTURA":
        df_Pivote.loc[i, "ESTADO"] = "RE-AGENDA FUTURA"
    elif df_Pivote.loc[i, "CALIF_PLEX"] == "COMPLETADA":
        df_Pivote.loc[i, "ESTADO"] = "COMPLETADA"
    elif pd.isna(df_Pivote.loc[i, "CALIF_PLEX"]):
        if df_Pivote.loc[i, "ESTATUS_PLEX"] == "PENDIENTE" and df_Pivote.loc[i, "CASA CERRADA HOY"] == "CASA_CERRADA":
            df_Pivote.loc[i, "ESTADO"] = "CASA_CERRADA"
        elif df_Pivote.loc[i, "ESTATUS_PLEX"] == "PENDIENTE" and (df_Pivote.loc[i, "TAREA_PLEX"] == "F841" or df_Pivote.loc[i, "TAREA_PLEX"] == "F597"):
            df_Pivote.loc[i, "ESTADO"] = "MANTENIMIENTO"
        else:
            df_Pivote.loc[i, "ESTADO"] = df_Pivote.loc[i, "ESTATUS_PLEX"]
    else:
        df_Pivote.loc[i, "ESTADO"] = df_Pivote.loc[i, "CALIF_PLEX"]
    
    # --------------------------------------------------------------

    # COMPLETADA_CITA FUTURA
    """=IF(OR(J8="COMPLETADA",J8="RE-AGENDA FUTURA"),"ATENDIDA/FUTURA","NO ATENDIDA")
                    J8 = ESTADO"""
    estado = df_Pivote.loc[i, "ESTADO"]
    if estado == "COMPLETADA" or estado == "RE-AGENDA FUTURA":
        df_Pivote.loc[i, "COMPLETADA_CITA FUTURA"] = "ATENDIDA/FUTURA"
    else:
        df_Pivote.loc[i, "COMPLETADA_CITA FUTURA"] = "NO ATENDIDA"

    # --------------------------------------------------------------

# Z - AA)  AREA_ACT y COPE_ACT
""" =IFNA(VLOOKUP(LEFT(D2,3),'SIGLAS COPE'!A:C,3,0),"-")
                       D2 = DISTRITO
                             SIGLAS COPE - A = siglas
                               SIGLAS COPE - C = COPE"""
df_SC_lookup = df_SC.drop_duplicates(subset="siglas", keep="first")
df_Pivote = df_Pivote.merge(df_SC_lookup, left_on=[df_Pivote["DISTRITO"].str[:3]], right_on="siglas", how="left")
df_Pivote["COPE_ACT"] = df_Pivote["cope"].fillna("")
df_Pivote["AREA_ACT"] = df_Pivote["AREA"].fillna("")
df_Pivote.drop(columns=["siglas","cope", "AREA"], inplace=True)

# ---------------------------------------------------------------------------------------------------

print(df_Pivote.head())
df_Pivote.to_csv("Pivote_EscalaB.csv", index=False)
print(f"\n→ Pivote guardado en Pivote_EscalaB.csv ({len(df_Pivote)} registros)")

print("\n" + "="*75)
print("→ Tabla de EscalaB")

df_TEB = pd.DataFrame(columns=["AREA_ACT", "COPE_ACT", "COMPLETADA", "ASIGNADA", "PENDIENTE", "CASA_CERRADA", "MANTENIMIENTO", "RE-AGENDA FUTURA", "Grand Total"])

df_TEB["COPE_ACT"] = df_Pivote["COPE_ACT"].unique()

df_TEB.loc[len(df_TEB), "AREA_ACT"] = "Ensenada Total"
df_TEB.loc[len(df_TEB), "AREA_ACT"] = "Mexicali Total"
df_TEB.loc[len(df_TEB), "AREA_ACT"] = "Tijuana Total"
df_TEB.loc[len(df_TEB), "AREA_ACT"] = "Total General"

df_TEB.fillna("0", inplace=True)
df_TEB = df_TEB.astype({"COMPLETADA": int, "ASIGNADA": int, "PENDIENTE": int, "CASA_CERRADA": int, "MANTENIMIENTO": int, "RE-AGENDA FUTURA": int, "Grand Total": int})


for i in range(len(df_TEB)):
    for n in range(len(df_Pivote)):
        if df_TEB.loc[i, "COPE_ACT"] == df_Pivote.loc[n, "COPE_ACT"]:
            df_TEB.loc[i, "AREA_ACT"] = df_Pivote.loc[n, "AREA_ACT"]

            estado = df_Pivote.loc[n, "ESTADO"]
            area = df_Pivote.loc[n, "AREA_ACT"]

            if estado == "COMPLETADA":
                df_TEB.loc[i, "COMPLETADA"] += 1

                if area == "ENSENADA":
                    df_TEB.loc[12, "COMPLETADA"] += 1
                elif area == "MEXICALI":
                    df_TEB.loc[13, "COMPLETADA"] += 1
                elif area == "TIJUANA":
                    df_TEB.loc[14, "COMPLETADA"] += 1

                df_TEB.loc[15, "COMPLETADA"] += 1

            elif estado == "ASIGNADA":
                df_TEB.loc[i, "ASIGNADA"] += 1

                if area == "ENSENADA":
                    df_TEB.loc[12, "ASIGNADA"] += 1
                elif area == "MEXICALI":
                    df_TEB.loc[13, "ASIGNADA"] += 1
                elif area == "TIJUANA":
                    df_TEB.loc[14, "ASIGNADA"] += 1

                df_TEB.loc[15, "ASIGNADA"] += 1
                    
            elif estado == "PENDIENTE":
                df_TEB.loc[i, "PENDIENTE"] += 1

                if area == "ENSENADA":
                    df_TEB.loc[12, "PENDIENTE"] += 1
                elif area == "MEXICALI":
                    df_TEB.loc[13, "PENDIENTE"] += 1
                elif area == "TIJUANA":
                    df_TEB.loc[14, "PENDIENTE"] += 1

                df_TEB.loc[15, "PENDIENTE"] += 1

            elif estado == "CASA_CERRADA":
                df_TEB.loc[i, "CASA_CERRADA"] += 1

                if area == "ENSENADA":
                    df_TEB.loc[12, "CASA_CERRADA"] += 1
                elif area == "MEXICALI":
                    df_TEB.loc[13, "CASA_CERRADA"] += 1
                elif area == "TIJUANA":
                    df_TEB.loc[14, "CASA_CERRADA"] += 1

                df_TEB.loc[15, "CASA_CERRADA"] += 1

            elif estado == "MANTENIMIENTO":
                df_TEB.loc[i, "MANTENIMIENTO"] += 1

                if area == "ENSENADA":
                    df_TEB.loc[12, "MANTENIMIENTO"] += 1
                elif area == "MEXICALI":
                    df_TEB.loc[13, "MANTENIMIENTO"] += 1
                elif area == "TIJUANA":
                    df_TEB.loc[14, "MANTENIMIENTO"] += 1

                df_TEB.loc[15, "MANTENIMIENTO"] += 1

            elif estado == "RE-AGENDA FUTURA":
                df_TEB.loc[i, "RE-AGENDA FUTURA"] += 1

                if area == "ENSENADA":
                    df_TEB.loc[12, "RE-AGENDA FUTURA"] += 1
                elif area == "MEXICALI":
                    df_TEB.loc[13, "RE-AGENDA FUTURA"] += 1
                elif area == "TIJUANA":
                    df_TEB.loc[14, "RE-AGENDA FUTURA"] += 1

                df_TEB.loc[15, "RE-AGENDA FUTURA"] += 1

for i in range(len(df_TEB)):
    df_TEB.loc[i, "Grand Total"] = df_TEB.loc[i, "COMPLETADA"] + df_TEB.loc[i, "PENDIENTE"] + df_TEB.loc[i, "CASA_CERRADA"] + df_TEB.loc[i, "MANTENIMIENTO"] + df_TEB.loc[i, "RE-AGENDA FUTURA"]


df_TEB.sort_values(by=['AREA_ACT'], inplace=True)
df_TEB.reset_index(drop=True, inplace=True)

df_TEB.replace({"0": ""}, inplace=True)
df_TEB.replace({0: ""}, inplace=True)

print(df_TEB)

print("="*75 + "\n")

df_TEB.to_excel("EscalaB.xlsx", index=False)
