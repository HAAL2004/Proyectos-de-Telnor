import pandas as pd
from datetime import date, timedelta
import numpy as np

pd.set_option('display.max_rows', None)



#Tabla con Nuevos Datos a agregar 
df_ND = pd.read_excel("Distritos HOY.xlsx")
df_ND["SGS"] = df_ND["SGS"].fillna('N/A', inplace=False)
df_ND = df_ND.iloc[:, :6]
for i in range(len(df_ND)):

    if "." in str(df_ND.loc[i, "Terminal"]):
        df_ND.loc[i, "Terminal"] = df_ND.loc[i, "Terminal"].replace(".", ",")




Sem = date.today().isocalendar().week
M =  14
print("Semana Actual:")
print(Sem)

#Tabla Principal
df_TP = pd.read_excel("Seguimiento PROGRAMAS FTTH-TBA_27ABR26.xlsm", sheet_name="BD")
df_TP = df_TP.iloc[12:].reset_index(drop=True)
df_TP = df_TP.drop(df_TP.columns[:1], axis=1)
df_TP.columns = df_TP.iloc[0]
df_TP = df_TP.iloc[1:].reset_index(drop=True)
df_TP = df_TP.iloc[:, :19]
df_TP = df_TP[df_TP["ETAPA SGS ACTUALIZADO"] != "En Servicio"]
df_TP = df_TP[["SEM PROG", "DISTRITO", "SGS", "Terminales", "ETAPA SGS ACTUALIZADO", "PUENTES"]]
df_TP = df_TP[df_TP['SEM PROG'].isin([Sem])].reset_index(drop=True)
df_TP["PUENTES"] = df_TP["PUENTES"].fillna('Nuevo', inplace=False)
for i in range(len(df_TP)):

    if "." in df_TP.loc[i, "Terminales"]:
        df_TP.loc[i, "Terminales"] = df_TP.loc[i, "Terminales"].replace(".", ",")



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
                if P[:2] in df_ND.loc[i, "Terminal"]:
                    break

                else:
                    LD.loc[len(LD)] = [df_TP.loc[n, "SGS"], df_TP.loc[n, "DISTRITO"], df_TP.loc[n, "Terminales"]]

            else:
                P = df_TP.loc[n, "Terminales"]
                if P[:2] in df_ND.loc[i, "Terminal"]:
                    break
                else:
                    LD.loc[len(LD)] = [df_TP.loc[n, "SGS"], df_TP.loc[n, "DISTRITO"], df_TP.loc[n, "Terminales"]]



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
    elif df_C.loc[i, "B X DISTRITO"] == df_ND.loc[i, "SGS"] and (df_C.loc[i, "Ter x Dis"] == df_ND.loc[i, "Terminal"] or df_C.loc[i, "Ter x Dis"] == df_C.loc[i, "BUSCA X SGS"]):
        df_C.loc[i, "¿Igual?"] = "SI"
    else:
        df_C.loc[i, "¿Igual?"] = "NO"

df = pd.concat([df_ND, df_C], axis=1)

print("Tabla Completa Antes de la Clasificacion:\n")
print(df,"\n")
print("--"*50,"\n")

df = df[df['¿Igual?'].isin(["NO", "N/A"])].reset_index(drop=True)

print("Tabla Completa Despues de la Clasificacion:\n")
print(df,"\n")
print("--"*50,"\n")

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

df_SGS = pd.DataFrame(columns=["SGS", "DIST", "Pon", "Terminal", "ETAPA SGS ACTUALIZADO", "Comentarios"])

# Solo falta SGS

for i in range(len(df)):
    if df.loc[i, "B X DISTRITO"] == "N/A" and df.loc[i, "SGS"] != "N/A":
        if df.loc[i, "Terminal"] == df.loc[i, "Ter x Dis"]:
            df_SGS.loc[len(df_SGS)] = [df.loc[i, "SGS"], df.loc[i, "DIST"], df.loc[i, "Pon"], df.loc[i, "Terminal"], "", ""]
    
        else:
            J = False
            for n in range(len(LD)):
                if df.loc[i, "DIST"] == LD.loc[n, "Distrito"]:

                    if LD.loc[n, "Terminales"][:2] in df.loc[i, "Terminal"]:
                        J = True
                        break


            
            if J:
                df_SGS.loc[len(df_SGS)] = [df.loc[i, "SGS"], df.loc[i, "DIST"], df.loc[i, "Pon"], df.loc[i, "Terminal"], "", "Validar terminales"]

        
           

#Eliminar registros ya clasificados

for i in (df_SGS.index):
    for n in (df).index:
        if df_SGS.loc[i, "SGS"] == df.loc[n, "SGS"] and df_SGS.loc[i, "DIST"] == df.loc[n, "DIST"] and df_SGS.loc[i, "Terminal"] == df.loc[n, "Terminal"]:
            df.drop(n, inplace=True)
            df.reset_index(drop=True, inplace=True)
            break

# Nuevos Registros

df_NR = pd.DataFrame(columns=["PROG", "SGS", "DIST", "PTOS", "Pon", "Terminal", "ETAPA SGS ACTUALIZADO", "Comentarios"])


for i in range(len(df)):
    if df.loc[i, "¿Igual?"] == "N/A":
        df_NR.loc[len(df_NR)] = [df.loc[i, "PROG"], df.loc[i, "SGS"], df.loc[i, "DIST"], df.loc[i, "PTOS"], df.loc[i, "Pon"], df.loc[i, "Terminal"], "", ""]


    elif df.loc[i, "Terminal"] != df.loc[i, "Ter x Dis"]:
        df_NR.loc[len(df_NR)] = [df.loc[i, "PROG"], df.loc[i, "SGS"], df.loc[i, "DIST"], df.loc[i, "PTOS"], df.loc[i, "Pon"], df.loc[i, "Terminal"], "", ""]
    elif df.loc[i, "B X DISTRITO"] != "N/A" and df.loc[i, "B X DISTRITO"] != df.loc[i, "SGS"] and df.loc[i, "SGS"] != "N/A":
        df_NR.loc[len(df_NR)] = [df.loc[i, "PROG"], df.loc[i, "SGS"], df.loc[i, "DIST"], df.loc[i, "PTOS"], df.loc[i, "Pon"], df.loc[i, "Terminal"], "", ""]

# Eliminar registros ya clasificados

for i in range(len(df_NR)):
    for n in range(len(df)):
        if df_NR.loc[i, "PROG"] == df.loc[n, "PROG"] and df_NR.loc[i, "SGS"] == df.loc[n, "SGS"] and df_NR.loc[i, "DIST"] == df.loc[n, "DIST"] and df_NR.loc[i, "Terminal"] == df.loc[n, "Terminal"]:
            df.drop(n, inplace=True)
            df.reset_index(drop=True, inplace=True)
            break

df_Prioridad = pd.read_excel("CATALOGO DISTRITOS PRIORITARIOS SEM15 2026.xlsx")

for i in range(len(df_NR)):
    for n in range(len(df_Prioridad)):
        if df_NR.loc[i, "DIST"] == df_Prioridad.loc[n, "Distrito"]:
            df_NR.loc[i, "Comentarios"] = "Prioridad"
            break

for i in range(len(df_SGS)):
    for n in range(len(df_Prioridad)):
        if df_SGS.loc[i, "DIST"] == df_Prioridad.loc[n, "Distrito"]:
            if df_SGS.loc[i, "Comentarios"] == "":
                df_SGS.loc[i, "Comentarios"] = "Prioridad"

            else:
                df_SGS.loc[i, "Comentarios"] = "Prioridad, " + df_SGS.loc[i, "Comentarios"]

            break




print("Faltantes por clasificar:\n")
print(df,"\n")
print("--"*50,"\n")            

print("Registrar solamente la clave SGS:\n")
print(df_SGS,"\n")
print("--"*50,"\n")

print("Nuevos Registros:\n")
print(df_NR,"\n")
print("--"*50,"\n")

# ================================================== FASE 2 ==================================================

print("================================================== FASE 2 ==================================================")
print("\n")

# Leer cada archivo UNA sola vez con ambas columnas
TablaNuevos = pd.read_csv("SGS-PON NUEVO.xls", sep='\t', encoding='latin-1', usecols=["SOLMASTER", "servicio_actual"])
TablaExistente = pd.read_csv("SGS-PON EXISTENTE.xls", sep='\t', encoding='latin-1', usecols=["SOLMASTER", "servicio_actual"])

ValorEstado = pd.read_excel("Seguimiento PROGRAMAS FTTH-TBA_27ABR26.xlsm", sheet_name="cat_eFTTH", usecols=["etapa", "orden"])

lookup = ValorEstado[["etapa", "orden"]].drop_duplicates(subset="etapa", keep="first")

TablaNuevos = TablaNuevos.merge(lookup, left_on="servicio_actual", right_on="etapa", how="left")
TablaNuevos["Etapa Reporte"] = np.where(TablaNuevos["orden"] == 99, "En Servicio", TablaNuevos["servicio_actual"])
TablaNuevos.drop(columns=["etapa", "orden", "servicio_actual"], inplace=True)

TablaExistente = TablaExistente.merge(lookup, left_on="servicio_actual", right_on="etapa", how="left")
TablaExistente["Etapa Reporte"] = np.where(TablaExistente["orden"] == 99, "En Servicio", TablaExistente["servicio_actual"])
TablaExistente.drop(columns=["etapa", "orden", "servicio_actual"], inplace=True)

for i in range(len(df_NR)):
    if df_NR.loc[i, "SGS"] == "N/A":
        df_NR.loc[i, "ETAPA SGS ACTUALIZADO"] = "PENDIENTE SGS"


    elif df_NR.loc[i, "Pon"] == "Nuevo":

        for n in range(len(TablaNuevos)):
            if df_NR.loc[i, "SGS"] == TablaNuevos.loc[n, "SOLMASTER"]:
                df_NR.loc[i, "ETAPA SGS ACTUALIZADO"] = TablaNuevos.loc[n, "Etapa Reporte"]
                break
    
    elif df_NR.loc[i, "Pon"] == "Exist":
        for n in range(len(TablaExistente)):
            if df_NR.loc[i, "SGS"] == TablaExistente.loc[n, "SOLMASTER"]:
                df_NR.loc[i, "ETAPA SGS ACTUALIZADO"] = TablaExistente.loc[n, "Etapa Reporte"]
                break

for i in range(len(df_SGS)):
    if df_SGS.loc[i, "Pon"] == "Nuevo":

        for n in range(len(TablaNuevos)):
            if df_SGS.loc[i, "SGS"] == TablaNuevos.loc[n, "SOLMASTER"]:
                df_SGS.loc[i, "ETAPA SGS ACTUALIZADO"] = TablaNuevos.loc[n, "Etapa Reporte"]
                break
    
    elif df_SGS.loc[i, "Pon"] == "Exist":
        for n in range(len(TablaExistente)):
            if df_SGS.loc[i, "SGS"] == TablaExistente.loc[n, "SOLMASTER"]:
                df_SGS.loc[i, "ETAPA SGS ACTUALIZADO"] = TablaExistente.loc[n, "Etapa Reporte"]
                break

print("Nuevos Registros:")
print(df_NR,"\n")
print("--"*50,"\n")

print("SGS:")
print(df_SGS,"\n")
print("--"*50,"\n")


df_TPCE = pd.DataFrame(columns=["SGS", "DIST", "ETAPA SGS ACTUALIZADO"])

for i in range(len(df_TP)):
    if df_TP.loc[i, "SGS"] == "N/A":
        pass

    elif df_TP.loc[i, "PUENTES"] == "Nuevo":
        for n in range(len(TablaNuevos)):
            if df_TP.loc[i, "SGS"] == TablaNuevos.loc[n, "SOLMASTER"] and df_TP.loc[i, "ETAPA SGS ACTUALIZADO"] != TablaNuevos.loc[n, "Etapa Reporte"]:
                df_TPCE.loc[len(df_TPCE)] = [df_TP.loc[i, "SGS"], df_TP.loc[i, "DISTRITO"], TablaNuevos.loc[n, "Etapa Reporte"]]
                break

    elif df_TP.loc[i, "PUENTES"] == "Exist":
        for n in range(len(TablaExistente)):
            if df_TP.loc[i, "SGS"] == TablaExistente.loc[n, "SOLMASTER"] and df_TP.loc[i, "ETAPA SGS ACTUALIZADO"] != TablaExistente.loc[n, "Etapa Reporte"]:
                df_TPCE.loc[len(df_TPCE)] = [df_TP.loc[i, "SGS"], df_TP.loc[i, "DISTRITO"], TablaExistente.loc[n, "Etapa Reporte"]]
                break

print("Cambios de Etapa en BD:")
print(df_TPCE,"\n")
print("--"*50,"\n")