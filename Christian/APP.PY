import os
import sys
import json
import uuid
import tempfile
import re

import pandas as pd
from datetime import date
from fastapi import FastAPI, Request, Form, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware

if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    # Cuando esté empaquetado como .exe (PyInstaller)
    BASE_DIR = sys._MEIPASS
else:
    # Cuando lo ejecutes normal con Python
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

WEB_DIR = os.path.join(BASE_DIR, "web")
USERS_PATH = os.path.join(BASE_DIR, "users.json")


app = FastAPI(title="TELNOR Local App")

# Sesión por cookie (para proteger /valeria y /christian)
app.add_middleware(SessionMiddleware, secret_key="CAMBIA_ESTA_CLAVE_POR_ALGO_LARGO")

# Archivos estáticos (css/img/js) desde /web como /static
app.mount("/static", StaticFiles(directory=WEB_DIR), name="static")

# Descargas temporales compartidas (Christian + Valeria)
# {id: filepath}
DOWNLOADS: dict[str, str] = {}


def read_html(filename: str) -> str:
    path = os.path.join(WEB_DIR, filename)
    if not os.path.exists(path):
        return HTMLResponse(f"No encontré {filename} en {WEB_DIR}", status_code=404)
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def load_users() -> dict:
    try:
        with open(USERS_PATH, "r", encoding="utf-8") as f:
            # 1️⃣ Si el archivo existe y se puede leer, usamos eso
            return json.load(f)
    except Exception as e:
        print("ERROR al leer users.json:", e)
        # 2️⃣ Si algo falla (como en el .exe), usamos estos usuarios por defecto
        return {
            "valeria": {
                "password": "129",
                "role": "valeria",
                "redirect": "/valeria"
            },
            "christian": {
                "password": "123",
                "role": "christian",
                "redirect": "/christian"
            },
            "joel": {
                "password": "123",
                "role": "joel",
                "redirect": "/joel"
            }
        }


def require_role(request: Request, role: str):
    user = request.session.get("user")
    if not user:
        raise HTTPException(status_code=401, detail="No autenticado")
    if user.get("role") != role:
        raise HTTPException(status_code=403, detail="Sin permisos")


# =========================
# LOGIN / LOGOUT
# =========================

@app.get("/", response_class=HTMLResponse)
def login_page():
    return read_html("login.html")


@app.post("/api/login")
def api_login(request: Request, usuario: str = Form(...), password: str = Form(...)):
    users = load_users()
    u = users.get(usuario.lower())

    if not u or u.get("password") != password:
        return JSONResponse({"ok": False, "msg": "Usuario o contraseña incorrectos"}, status_code=401)

    request.session["user"] = {"username": usuario.lower(), "role": u["role"]}
    return {"ok": True, "redirect": u["redirect"]}


@app.post("/api/logout")
def api_logout(request: Request):
    request.session.clear()
    return {"ok": True}


# =========================
# PÁGINAS PROTEGIDAS
# =========================

@app.get("/valeria", response_class=HTMLResponse)
def valeria_page(request: Request):
    require_role(request, "valeria")
    return read_html("valeria.html")


@app.get("/christian", response_class=HTMLResponse)
def christian_page(request: Request):
    require_role(request, "christian")
    return read_html("christian.html")

@app.get("/joel", response_class=HTMLResponse)
def joel_page(request: Request):
    require_role(request, "joel")
    return read_html("joel.html")

# =========================
# CHRISTIAN - BACKEND REAL
# =========================

@app.post("/api/christian/procesar")
async def christian_procesar(
    request: Request,
    documento: UploadFile = File(...),
    base: UploadFile = File(...),
    semana: int = Form(None),
):
    require_role(request, "christian")

    try:
        with tempfile.TemporaryDirectory(prefix="telnor_christian_") as tmp:
            nd_path = os.path.join(tmp, documento.filename or "Distritos.xlsx")
            tp_path = os.path.join(tmp, base.filename or "Seguimiento.xlsm")

            with open(nd_path, "wb") as f:
                f.write(await documento.read())
            with open(tp_path, "wb") as f:
                f.write(await base.read())

            # Tabla con Nuevos Datos a agregar
            df_ND = pd.read_excel(nd_path)
            df_ND["SGS"] = df_ND["SGS"].fillna('N/A', inplace=False)
            for i in range(len(df_ND)):
                if "." in df_ND.loc[i, "Terminal"]:
                    df_ND.loc[i, "Terminal"] = df_ND.loc[i, "Terminal"].replace(".", ",")
                    
            if semana is None:
                Sem = date.today().isocalendar().week
            else:
                Sem = int(semana)

            # Tabla Principal
            df_TP = pd.read_excel(tp_path, sheet_name="BD")
            df_TP = df_TP.iloc[12:].reset_index(drop=True)
            df_TP = df_TP.drop(df_TP.columns[:1], axis=1)
            df_TP.columns = df_TP.iloc[0]
            df_TP = df_TP.iloc[1:].reset_index(drop=True)
            df_TP = df_TP.iloc[:, :19]
            df_TP = df_TP[["SEM PROG", "DISTRITO", "SGS", "Terminales"]]
            df_TP = df_TP[df_TP['SEM PROG'].isin([Sem])].reset_index(drop=True)
            for i in range(len(df_TP)):
                if "." in df_TP.loc[i, "Terminales"]:
                    df_TP.loc[i, "Terminales"] = df_TP.loc[i, "Terminales"].replace(".", ",")

            if df_TP.empty:
                return JSONResponse(
                    {"error": f"No hay registros en BD para SEM PROG = {Sem}."},
                    status_code=400,
                )

            y = 0
            df_C = pd.DataFrame(columns=["BUSCA X SGS","B X DISTRITO", "Ter x Dis", "¿Igual?"])

            # Busqueda de SGS
            for i in range(len(df_ND)):
                for n in range(len(df_TP)):
                    if df_ND.loc[i, "SGS"] == "N/A":
                        df_C.loc[y, "BUSCA X SGS"] = "N/A"
                        y += 1
                        break
                    elif df_ND.loc[i, "SGS"] == df_TP.loc[n, "SGS"]:
                        df_C.loc[y, "BUSCA X SGS"] = df_TP.loc[n, "Terminales"]
                        y += 1
                        break
                else:
                    y += 1
            else:
                y += 1
        
            y = 0

            # Busqueda de Distrito
            for i in range(len(df_ND)):
                P = None
                for n in range(len(df_TP)):
                    
                    if df_ND.loc[i, "DIST"] == df_TP.loc[n, "DISTRITO"]:
                        if P == None:
                            P = df_TP.loc[n, "SGS"]
                            if P == df_ND.loc[i, "SGS"]:
                                break 


                        else:
                            P = df_TP.loc[n, "SGS"]
                            if P == df_ND.loc[i, "SGS"]:
                                break 

                if P == None:
                    df_C.loc[y, "B X DISTRITO"] = "N/A"
                    y += 1
                else:
                    df_C.loc[y, "B X DISTRITO"] = P
                    y += 1

            y = 0
            LD = pd.DataFrame(columns=["SGS", "Distrito", "Terminales"])


            # Busqueda de Terminales por Distrito
            for i in range(len(df_ND)):
                P = None
                for n in range(len(df_TP)):
                    if df_ND.loc[i, "DIST"] == df_TP.loc[n, "DISTRITO"]:
                        if P == None:
                            P = df_TP.loc[n, "Terminales"]
                            if P != df_ND.loc[i, "Terminal"]:
                                LD.loc[len(LD)] = [df_TP.loc[n, "SGS"], df_TP.loc[n, "DISTRITO"], df_TP.loc[n, "Terminales"]]
                        elif P == df_ND.loc[i, "Terminal"]:

                            break
                        else:
                            LD.loc[len(LD)] = [df_TP.loc[n, "SGS"], df_TP.loc[n, "DISTRITO"], df_TP.loc[n, "Terminales"]]
                            P = df_TP.loc[n, "Terminales"]



                if P == None:
                    df_C.loc[y, "Ter x Dis"] = "N/A"
                    y += 1

                else:        
                    df_C.loc[y, "Ter x Dis"] = P
                    y += 1



            LD.drop_duplicates(inplace=True)
            LD.reset_index(drop=True, inplace=True)
            print(LD)

            df_C["BUSCA X SGS"] = df_C["BUSCA X SGS"].fillna('N/A', inplace=False)
            df_C["B X DISTRITO"] = df_C["B X DISTRITO"].fillna('N/A', inplace=False)

            # Comparacion
            for i in range(len(df_C)):
                if df_C.loc[i, "BUSCA X SGS"] == "N/A" and df_C.loc[i, "B X DISTRITO"] == "N/A" and df_C.loc[i, "Ter x Dis"] == "N/A":
                    df_C.loc[i, "¿Igual?"] = "N/A"
                elif df_C.loc[i, "B X DISTRITO"] == df_ND.loc[i, "SGS"] and df_C.loc[i, "Ter x Dis"] == df_ND.loc[i, "Terminal"]:
                    df_C.loc[i, "¿Igual?"] = "SI"
                else:
                    df_C.loc[i, "¿Igual?"] = "NO"

            df = pd.concat([df_ND, df_C], axis=1)
            df = df[df['¿Igual?'].isin(["NO", "N/A"])].reset_index(drop=True)

            i = 0
            while i < len(df):
                if df.loc[i, "Terminal"] == df.loc[i, "Ter x Dis"] and df.loc[i, "SGS"] != df.loc[i, "B X DISTRITO"] and df.loc[i, "SGS"] == "N/A":
                    df.drop(i, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                elif df.loc[i, "SGS"] != df.loc[i, "B X DISTRITO"] and df.loc[i, "SGS"] == "N/A" and df.loc[i, "Terminal"]  in df.loc[i, "Ter x Dis"]: 
                    df.drop(i, inplace=True)
                    df.reset_index(drop=True, inplace=True)
                else:
                    i += 1

            df["B X DISTRITO"] = df["B X DISTRITO"].fillna('N/A', inplace=False)

            df_SGS = pd.DataFrame(columns=["SGS", "DIST", "Terminal", "Comentarios"])

            # Solo falta SGS

            for i in range(len(df)):
                if df.loc[i, "B X DISTRITO"] == "N/A" and df.loc[i, "SGS"] != "N/A":
                    if df.loc[i, "Terminal"] == df.loc[i, "Ter x Dis"]:
                        df_SGS.loc[len(df_SGS)] = [df.loc[i, "SGS"], df.loc[i, "DIST"], df.loc[i, "Terminal"], ""]
        
                    else:
                        J = False
                        for n in range(len(LD)):

                            if df.loc[i, "DIST"] == LD.loc[n, "Distrito"] and LD.loc[n, "Terminales"][1:] in df.loc[i, "Terminal"]:
                                J = True
                                break
            
                        if J:
                            df_SGS.loc[len(df_SGS)] = [df.loc[i, "SGS"], df.loc[i, "DIST"], df.loc[i, "Terminal"], "Se encontraron terminales"]
           

            #Eliminar registros ya clasificados

            for i in (df_SGS.index):
                for n in (df).index:
                    if df_SGS.loc[i, "SGS"] == df.loc[n, "SGS"] and df_SGS.loc[i, "DIST"] == df.loc[n, "DIST"] and df_SGS.loc[i, "Terminal"] == df.loc[n, "Terminal"]:
                        df.drop(n, inplace=True)
                        df.reset_index(drop=True, inplace=True)
                        break

            # Nuevos Registros

            df_NR = pd.DataFrame(columns=["PROG", "SGS", "DIST", "PTOS", "Pon", "Terminal"])


            for i in range(len(df)):
                if df.loc[i, "¿Igual?"] == "N/A":
                    df_NR.loc[len(df_NR)] = [df.loc[i, "PROG"], df.loc[i, "SGS"], df.loc[i, "DIST"], df.loc[i, "PTOS"], df.loc[i, "Pon"], df.loc[i, "Terminal"]]

                elif df.loc[i, "Terminal"] != df.loc[i, "Ter x Dis"]:
                    df_NR.loc[len(df_NR)] = [df.loc[i, "PROG"], df.loc[i, "SGS"], df.loc[i, "DIST"], df.loc[i, "PTOS"], df.loc[i, "Pon"], df.loc[i, "Terminal"]]
                elif df.loc[i, "B X DISTRITO"] != "N/A" and df.loc[i, "B X DISTRITO"] != df.loc[i, "SGS"] and df.loc[i, "SGS"] != "N/A":
                    df_NR.loc[len(df_NR)] = [df.loc[i, "PROG"], df.loc[i, "SGS"], df.loc[i, "DIST"], df.loc[i, "PTOS"], df.loc[i, "Pon"], df.loc[i, "Terminal"]]

            # Eliminar registros ya clasificados

            for i in range(len(df_NR)):
                for n in range(len(df)):
                    if df_NR.loc[i, "PROG"] == df.loc[n, "PROG"] and df_NR.loc[i, "SGS"] == df.loc[n, "SGS"] and df_NR.loc[i, "DIST"] == df.loc[n, "DIST"] and df_NR.loc[i, "Terminal"] == df.loc[n, "Terminal"]:
                        df.drop(n, inplace=True)
                        df.reset_index(drop=True, inplace=True)
                        break

            # Guardar Excel
            out_path = os.path.join(tmp, "diferencias.xlsx")
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                df_NR.to_excel(writer, index=False, sheet_name="Nuevos_Registros")
                df_SGS.to_excel(writer, index=False, sheet_name="Registrar_SGS")

            file_id = str(uuid.uuid4())
            saved_path = os.path.join(tempfile.gettempdir(), f"Distritos_Reg.xlsx")
            with open(out_path, "rb") as src, open(saved_path, "wb") as dst:
                dst.write(src.read())
            DOWNLOADS[file_id] = saved_path

            return {"ok": True, "download_id": file_id}

    except Exception as e:
        return JSONResponse(
            {"error": f"Error inesperado al procesar archivos: {e}"},
            status_code=500,
        )


@app.get("/api/christian/descargar/{file_id}")
def christian_descargar(request: Request, file_id: str):
    require_role(request, "christian")

    if file_id not in DOWNLOADS:
        return JSONResponse({"error": "Archivo no encontrado"}, status_code=404)

    path = DOWNLOADS[file_id]
    return FileResponse(path, filename=os.path.basename(path))
# =========================
# JOEL - BACKEND REAL
# =========================

@app.post("/api/joel/procesar")
async def joel_procesar(
    request: Request,
    archivo_nuevo: UploadFile = File(...),
    archivo_viejo: UploadFile = File(...),
):
    require_role(request, "joel")

    try:
        with tempfile.TemporaryDirectory(prefix="telnor_joel_") as tmpdir:
            path_nuevo = os.path.join(tmpdir, archivo_nuevo.filename or "nuevo.txt")
            path_viejo = os.path.join(tmpdir, archivo_viejo.filename or "viejo.txt")

            with open(path_nuevo, "wb") as f:
                f.write(await archivo_nuevo.read())
            with open(path_viejo, "wb") as f:
                f.write(await archivo_viejo.read())



            try:
                df_N = pd.read_csv(path_nuevo, sep=r"\s*\|\s*", encoding="latin-1",
                        on_bad_lines="skip", engine="python",
                        usecols=['idReg',"folio", 'fase', "NombreUsuarioInserta"])
                df_V = pd.read_csv(path_viejo, sep=r"\s*\|\s*", encoding="latin-1",
                        on_bad_lines="skip", engine="python",
                        usecols=["idReg", "fase"])

                
            except Exception as e:
                return JSONResponse(
                    {"ok": False, "error": f"Error leyendo archivos: {e}"},
                    status_code=400,
                )

            df_N['idReg'] = df_N['idReg'].astype(str).str.strip()
            df_V['idReg'] = df_V['idReg'].astype(str).str.strip()

            df = pd.merge(df_N, df_V, on='idReg', how='left',suffixes=('_actual', '_anterior'))

            for i in range(len(df)):
                if df.loc[i, "fase_actual"] == df.loc[i, "fase_anterior"]:
                    df.loc[i,"Estado_fase"] = False
                else:
                    df.loc[i, "Estado_fase"] = True


            df = df[['idReg', 'folio', 'fase_actual', 'fase_anterior', 'Estado_fase', 'usuarioInserta']]

            # Exportar como CSV (mucho más rápido que Excel para archivos grandes)
            out_path = os.path.join(tmpdir, "Estado_Fase.csv")
            df.to_csv(out_path, index=False, encoding="utf-8-sig")

            file_id = str(uuid.uuid4())
            saved_path = os.path.join(tempfile.gettempdir(), f"Estado_Fase.csv")
            with open(out_path, "rb") as src, open(saved_path, "wb") as dst:
                dst.write(src.read())
            DOWNLOADS[file_id] = saved_path

            return {"ok": True, "download_id": file_id}

    except Exception as e:
        return JSONResponse(
            {"ok": False, "error": f"Error inesperado al procesar: {e}"},
            status_code=500,
        )


@app.get("/api/joel/descargar/{file_id}")
def joel_descargar(request: Request, file_id: str):
    require_role(request, "joel")
    if file_id not in DOWNLOADS:
        return JSONResponse({"error": "Archivo no encontrado"}, status_code=404)
    path = DOWNLOADS[file_id]
    return FileResponse(path, filename="resultado_joel.csv")

# =========================
# VALERIA - BACKEND REAL (NODOS + ITX)
# =========================

# =========================
# VALERIA - BACKEND REAL (NODOS + ITX)
# =========================

# --- CONFIGURACIÓN ---
SHEET_NODOS = "Nodos"
SHEET_ITX = "ITX"
OUTPUT_FILE_NAME_VALERIA = "semaforo_limpio.xlsx"

# ----------------- HELPERS -----------------
def encontrar_fila_header(ruta_archivo, sheet_name, keyword="Sensor", nrows=80):
    df_temp = pd.read_excel(ruta_archivo, sheet_name=sheet_name, header=None, nrows=nrows)
    for idx, row in df_temp.iterrows():
        if keyword in row.astype(str).str.cat(sep=" "):
            return idx
    return 0


def normalizar_columnas(df):
    df.columns = df.columns.astype(str).str.replace("\u00a0", " ").str.strip()
    df = df.loc[:, ~df.columns.str.contains(r"^Unnamed", na=False)]
    return df


def limpiar_probe_group_device(valor):
    """Local Probe » TIJUANA » BASTIJ26.TELNOR.NET -> BASTIJ26"""
    if pd.isna(valor):
        return ""
    s = str(valor).replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    partes = re.split(r"\s»\s", s)
    s = partes[-1].strip()
    s = s.split(".", 1)[0].strip()
    return s


def limpiar_kbits_a_numero(valor):
    """
    "5,082,895 kbit/s" -> 5082895
    ""/NaN -> NaN
    """
    if pd.isna(valor):
        return pd.NA
    if isinstance(valor, (int, float)) and not pd.isna(valor):
        return valor

    s = str(valor).replace("\u00a0", " ").strip()
    if s == "":
        return pd.NA

    s = re.sub(r"\s*kbits?/s\s*$", "", s, flags=re.IGNORECASE).strip()
    s = s.replace(",", "").strip()
    return pd.to_numeric(s, errors="coerce")


def obtener_col_percentil(df):
    return next((c for c in df.columns if c.strip().lower() in ["percentile", "percentil"]), None)


# ---------- NODOS ----------
def extraer_vlan(sensor_val):
    if pd.isna(sensor_val):
        return pd.NA
    s = str(sensor_val).replace("\u00a0", " ").strip()

    m = re.search(r":\s*\d+\.(\d+)", s)
    if m:
        return pd.to_numeric(m.group(1), errors="coerce")

    m = re.search(r"/\s*\d+\.(\d+)", s)
    if m:
        return pd.to_numeric(m.group(1), errors="coerce")

    return pd.NA


def extraer_nom_equipo(sensor_val):
    if pd.isna(sensor_val):
        return ""
    s = str(sensor_val).replace("\u00a0", " ").strip()
    m = re.search(r"DSL\[[^\]]+\]\[([^\]]+)\]", s, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""


# ---------- ITX ----------
def extraer_interfaz_del_sensor(sensor_val):
    if pd.isna(sensor_val):
        return ""
    s = str(sensor_val).replace("\u00a0", " ").strip()
    s = re.sub(r"^\(\s*\d+\s*\)\s*", "", s)

    m = re.search(
        r"(xe-\S+|ge-\S+|TenGigE?\S+|TenGig\S+|GigabitEthernet\S+)",
        s,
        flags=re.IGNORECASE
    )
    return m.group(1).strip() if m else ""


def normalizar_interfaz(interfaz):
    if not interfaz:
        return ""
    x = interfaz.strip()

    if re.match(r"^xe-", x, flags=re.IGNORECASE):
        return "XE" + x[3:]
    if re.match(r"^ge-", x, flags=re.IGNORECASE):
        return "GE" + x[3:]
    if re.match(r"^GigabitEthernet", x, flags=re.IGNORECASE):
        return "GE" + x[len("GigabitEthernet"):]
    if re.match(r"^TenGigE", x, flags=re.IGNORECASE):
        return "TE" + x[len("TenGigE"):]
    if re.match(r"^TenGig", x, flags=re.IGNORECASE):
        return "TE" + x[len("TenGig"):]
    return x


def construir_bas_interfaz(probe_group_device, sensor):
    p = "" if pd.isna(probe_group_device) else str(probe_group_device).strip()
    inter = normalizar_interfaz(extraer_interfaz_del_sensor(sensor))
    if p and inter:
        return f"{p}_{inter}"
    if p:
        return p
    return inter


def extraer_nombre_itx(sensor_val):
    if pd.isna(sensor_val):
        return ""
    s = str(sensor_val).replace("\u00a0", " ").strip()
    m = re.search(r"\[(ITX[^\]]+)\]", s, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""


# ----------------- LIMPIEZA POR SHEET -----------------
def limpiar_sheet_nodos(ruta_archivo):
    fila_header = encontrar_fila_header(ruta_archivo, SHEET_NODOS, keyword="Sensor")
    df = pd.read_excel(ruta_archivo, sheet_name=SHEET_NODOS, header=fila_header)
    df = normalizar_columnas(df)

    if "Probe Group Device" in df.columns:
        df["Probe Group Device"] = df["Probe Group Device"].apply(limpiar_probe_group_device)

    speed_targets = ["Average", "Minimum", "Maximum", "Percentile", "Percentil"]
    speed_cols = [c for c in df.columns if c.strip().lower() in [x.lower() for x in speed_targets]]
    for col in speed_cols:
        df[col] = df[col].apply(limpiar_kbits_a_numero)

    # SOLO Minimum: vacíos -> 0
    col_min = next((c for c in df.columns if c.strip().lower() == "minimum"), None)
    if col_min:
        df[col_min] = pd.to_numeric(df[col_min], errors="coerce").fillna(0)

    if "Sensor" in df.columns:
        df["VLAN"] = df["Sensor"].apply(extraer_vlan)
        df["NomEquipo"] = df["Sensor"].apply(extraer_nom_equipo)

        df = df.dropna(subset=["Sensor"])
        df = df[df["Sensor"].astype(str).str.strip() != ""]

    # ✅ Puertos ocupados = conteo de repeticiones de NomEquipo
    if "NomEquipo" in df.columns:
        df["Puertos ocupados"] = df.groupby("NomEquipo")["NomEquipo"].transform("size").astype("Int64")

    # Ordenar por percentil si existe
    col_p = obtener_col_percentil(df)
    if col_p:
        df[col_p] = pd.to_numeric(df[col_p], errors="coerce")
        df = df.sort_values(by=col_p, ascending=False)

    # Rellenar SOLO texto con ""
    cols_texto = df.select_dtypes(include=["object"]).columns
    df[cols_texto] = df[cols_texto].fillna("")

    # Export compacto (A-H) para que ITX quepa desde J
    col_p = obtener_col_percentil(df)
    core = []
    for c in ["NomEquipo", "Sensor", "Probe Group Device", "Average", "Minimum", "Maximum"]:
        if c in df.columns:
            core.append(c)
    if col_p and col_p in df.columns:
        core.append(col_p)
    if "VLAN" in df.columns:
        core.append("VLAN")
    # 🔥 La nueva columna debe ir al lado de VLAN
    if "Puertos ocupados" in df.columns:
        core.append("Puertos ocupados")

    return df[core] if core else df


def limpiar_sheet_itx(ruta_archivo):
    fila_header = encontrar_fila_header(ruta_archivo, SHEET_ITX, keyword="Sensor")
    df = pd.read_excel(ruta_archivo, sheet_name=SHEET_ITX, header=fila_header)
    df = normalizar_columnas(df)

    if "Probe Group Device" in df.columns:
        df["Probe Group Device"] = df["Probe Group Device"].apply(limpiar_probe_group_device)

    speed_targets = ["Average", "Minimum", "Maximum", "Percentile", "Percentil"]
    speed_cols = [c for c in df.columns if c.strip().lower() in [x.lower() for x in speed_targets]]
    for col in speed_cols:
        df[col] = df[col].apply(limpiar_kbits_a_numero)

    if "Sensor" in df.columns:
        df["BAS_INTERFAZ"] = df.apply(
            lambda r: construir_bas_interfaz(r.get("Probe Group Device", ""), r.get("Sensor", "")),
            axis=1
        )
        df["Nombre de ITX"] = df["Sensor"].apply(extraer_nombre_itx)

        df = df.dropna(subset=["Sensor"])
        df = df[df["Sensor"].astype(str).str.strip() != ""]

    col_p = obtener_col_percentil(df)
    if col_p:
        df[col_p] = pd.to_numeric(df[col_p], errors="coerce")
        df = df.sort_values(by=col_p, ascending=False)

    cols_texto = df.select_dtypes(include=["object"]).columns
    df[cols_texto] = df[cols_texto].fillna("")

    if "Probe Group Device" in df.columns:
        df = df.drop(columns=["Probe Group Device"])

    col_p = obtener_col_percentil(df)
    core = []
    if "BAS_INTERFAZ" in df.columns:
        core.append("BAS_INTERFAZ")
    if "Nombre de ITX" in df.columns:
        core.append("Nombre de ITX")
    for c in ["Sensor", "Average", "Minimum", "Maximum"]:
        if c in df.columns:
            core.append(c)
    if col_p and col_p in df.columns:
        core.append(col_p)

    return df[core] if core else df


# ----------------- ENDPOINT FASTAPI -----------------
@app.post("/api/valeria/limpiar")
async def valeria_limpiar(request: Request, archivo: UploadFile = File(...)):
    """
    Recibe un Excel, limpia Nodos + ITX y genera:
      - Hoja 'Semaforo_Limpio'
      - Nodos desde la columna A
      - ITX desde la columna J
    """
    require_role(request, "valeria")

    nombre = archivo.filename or "archivo.xlsx"
    ext = os.path.splitext(nombre)[1].lower()
    if ext not in {".xlsx", ".xls"}:
        return JSONResponse(
            {
                "ok": False,
                "error": f"Extensión no válida: {ext}. Solo se permiten .xlsx o .xls",
            },
            status_code=400,
        )

    try:
        # Carpeta temporal propia de esta petición
        tmpdir = tempfile.mkdtemp(prefix="telnor_valeria_")
        input_path = os.path.join(tmpdir, nombre)

        # Guardar archivo subido
        with open(input_path, "wb") as f:
            f.write(await archivo.read())

        # Revisar hojas disponibles
        xls = pd.ExcelFile(input_path)
        faltantes = []
        if SHEET_NODOS not in xls.sheet_names:
            faltantes.append(SHEET_NODOS)
        if SHEET_ITX not in xls.sheet_names:
            faltantes.append(SHEET_ITX)

        if faltantes:
            return JSONResponse(
                {
                    "ok": False,
                    "error": f"No encontré las hojas requeridas: {faltantes}. "
                             f"Hojas disponibles: {xls.sheet_names}",
                },
                status_code=400,
            )

        # LIMPIEZA
        df_nodos = limpiar_sheet_nodos(input_path)
        df_itx = limpiar_sheet_itx(input_path)

        # Crear archivo de salida
        output_path = os.path.join(tmpdir, OUTPUT_FILE_NAME_VALERIA)
        sheet_out = "Semaforo_Limpio"

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df_nodos.to_excel(writer, sheet_name=sheet_out, index=False, startrow=0, startcol=0)   # A
            df_itx.to_excel(writer,   sheet_name=sheet_out, index=False, startrow=0, startcol=9)   # J

        # Copiar a carpeta de descargas compartida
        file_id = str(uuid.uuid4())
        saved_path = os.path.join(tempfile.gettempdir(), f"valeria_{file_id}.xlsx")

        with open(output_path, "rb") as src, open(saved_path, "wb") as dst:
            dst.write(src.read())

        DOWNLOADS[file_id] = saved_path

        return {
            "ok": True,
            "message": "Limpieza completada (Nodos + ITX)",
            "download_id": file_id,
            "nodos_rows": len(df_nodos),
            "itx_rows": len(df_itx),
        }

    except Exception as e:
        return JSONResponse(
            {
                "ok": False,
                "error": f"Error inesperado al procesar el archivo: {e}",
            },
            status_code=500,
        )


@app.get("/api/valeria/descargar/{file_id}")
def valeria_descargar(request: Request, file_id: str):
    require_role(request, "valeria")

    path = DOWNLOADS.get(file_id)
    if not path or not os.path.exists(path):
        return JSONResponse({"error": "Archivo no encontrado"}, status_code=404)

    return FileResponse(path, filename=OUTPUT_FILE_NAME_VALERIA)