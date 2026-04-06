# ============================================================================
# Limpieza del Anexo 1 - Inversiones en infraestructura educativa
# Base declarada por los GR y GL
# ============================================================================

import pandas as pd
import re
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# rutas (mi estructura de carpetas)
ruta = Path(r"C:\Users\iraid\Documents\DISCO TERA\MINEDU\TRABAJO-MINEDU\01_MINEDU")
entrada = ruta / "01_input"
salida  = ruta / "03_output"
temporal = ruta / "04_temporal"

# buscar archivos de entrada (a veces cambian de nombre)
archivo_anexo1 = entrada / "Anexo_1_avance_GR_GL.xlsx"

archivo_vinc = entrada / "Vinculaciones_compartido.xlsx"
archivo_inv = entrada / "Base_inversiones.xlsx"

print("=" * 50)
print("LIMPIEZA DEL ANEXO 1")
print("=" * 50)

# ============================================================================
# PASO 1: Cargar el Anexo 1
# ============================================================================
print("\n>> Cargando Anexo 1...")
df = pd.read_excel(archivo_anexo1, sheet_name="Hoja1", header=0, dtype=str)
print(f"   {len(df)} filas, {len(df.columns)} columnas")

# las primeras 2 filas son basura del encabezado
df = df.drop(index=[0, 1]).reset_index(drop=True)

# renombro por posicion porque los nombres del excel son un desastre
nombres = ["nro","cod_local","region","provincia","distrito","nombre_ie","cui",
           "tipo","monto","avance","fecha","f9","comp","unid","cod_mod",
           "demol","nueva","reforz","cerco","sust","ampl","mobil","agua","elec","comentarios"]
for i, nom in enumerate(nombres):
    if i < len(df.columns):
        df = df.rename(columns={df.columns[i]: nom})

# si hay columnas de mas las saco
extras = [c for c in df.columns if c not in nombres]
df = df.drop(columns=extras)

# me quedo solo con los que tienen CUI (los demas estan pendientes de registro)
df = df[df["cui"].notna() & (df["cui"].str.strip() != "")].copy()
total_con_cui = len(df)
print(f"   Filas con CUI registrado: {total_con_cui}")

# ============================================================================
# PASO 2: Cargar las bases maestras
# ============================================================================

# --- Vinculaciones ---
vinc = None
if archivo_vinc and archivo_vinc.exists():
    print(f"\n>> Cargando Vinculaciones: {archivo_vinc.name}")
    vinc = pd.read_excel(archivo_vinc, sheet_name="Vinculaciones", dtype=str)
    # limpio el CUI para que sea string sin .0
    vinc["cui_limpio"] = vinc["CUI"].apply(
        lambda x: str(int(float(x))) if pd.notna(x) and re.match(r'^\d+\.?\d*$', str(x).strip()) else str(x).strip() if pd.notna(x) else ""
    )
    print(f"   {len(vinc)} registros")
else:
    print("\n>> Vinculaciones no encontrada, se omite")

# --- Base de inversiones ---
inv = None
if archivo_inv and archivo_inv.exists():
    print(f"\n>> Cargando Base Inversiones: {archivo_inv.name}")
    xls_inv = pd.ExcelFile(archivo_inv)
    inv = pd.read_excel(xls_inv, sheet_name=xls_inv.sheet_names[0], dtype=str)
    # detectar columna de CUI
    col_cui_inv = None
    for c in ["CODIGO_UNICO", "CUI", "CODIGO_INVERSION"]:
        if c in inv.columns:
            col_cui_inv = c
            break
    if col_cui_inv:
        inv["cui_limpio"] = inv[col_cui_inv].apply(
            lambda x: str(int(float(x))) if pd.notna(x) and re.match(r'^\d+\.?\d*$', str(x).strip()) else str(x).strip() if pd.notna(x) else ""
        )
    print(f"   {len(inv)} registros, columna CUI: {col_cui_inv}")
else:
    print("\n>> Base Inversiones no encontrada, se omite")

# ============================================================================
# PASO 3: Funciones para limpiar cada campo
# ============================================================================

# para marcar errores uso este flag
ERR = "__ERROR__"

def limpiar_cui(val):
    if pd.isna(val): return val
    s = str(val).strip()
    if re.match(r'^\d+$', s): return s
    # intento rescatar una secuencia de 5+ digitos
    m = re.search(r'\d{5,}', s)
    if m: return m.group()
    return ERR

def limpiar_tipo(val):
    if pd.isna(val) or str(val).strip() == "": return val
    s = str(val).strip().upper()
    s = re.sub(r'\s+', ' ', s)
    # si pusieron un numero (como un monto) eso no es tipo
    if re.match(r'^\d+\.?\d*$', s): return ERR
    if s in ("S/N", "SN", ""): return ERR
    # normalizar
    if "IOAR" in s or "FUR" in s: return "IOARR"
    if "IRI" in s: return "IRI"
    if "PI" in s or "PROYECTO" in s: return "PI"
    return ERR

def limpiar_monto(val):
    if pd.isna(val) or str(val).strip() == "": return val
    s = str(val).strip().replace(",", ".").replace(" ", "")
    s = s.replace("S/", "").replace("s/", "")
    try:
        n = float(s)
        if n >= 0: return str(n)
        return ERR
    except:
        return ERR

def limpiar_avance(val):
    if pd.isna(val) or str(val).strip() == "": return val
    s = str(val).strip().replace("%", "").replace(",", ".")
    try:
        n = float(s)
        # si esta entre 0 y 1 es proporcion, lo paso a porcentaje
        if 0 <= n <= 1: return str(round(n * 100, 2))
        if 0 <= n <= 100: return str(round(n, 2))
        return ERR
    except:
        return ERR

def limpiar_f9(val):
    if pd.isna(val) or str(val).strip() == "": return val
    s = str(val).strip().upper().replace("Í", "I")
    # a veces se corren las columnas y aca aparece el componente
    if any(x in s for x in ["INFRAESTRUCTURA","EQUIPAMIENTO","MOBILIARIO","INTEGRAL"]):
        return ERR
    if s in ("SI","S","1","SÍ") or s.startswith("SI "): return "SI"
    if s in ("NO","N","0") or s.startswith("NO "): return "NO"
    return ERR

def limpiar_comp(val):
    if pd.isna(val) or str(val).strip() == "": return val
    s = re.sub(r'\s+', ' ', str(val).strip().upper())
    # basura
    if s in ("-","_","-1","--","---","----","-----","------","-------","-----------------------"):
        return ERR
    # integral (4) - combinaciones de 2 o mas
    if any(x in s for x in ["INTEGRAL","1,2,3","1, 2 Y 3","1 Y 2 Y 3","(1,2","(1-2"]):
        return "4. Integral (1, 2 y 3)"
    if any(x in s for x in ["1,2","1, 2","1 Y 2","1,3","1 Y 3","2,3","2 Y 3","1.2","1.3","2.3"]):
        return "4. Integral (1, 2 y 3)"
    # individuales
    if s in ("1","1.") or "INFRAESTRUCTURA" in s:
        return "1. Infraestructura."
    if s in ("2","2.") or "EQUIPAMIENTO" in s or "EQUIPAMIENT" in s:
        return "2. Equipamiento."
    if s in ("3","3.") or "MOBILIARIO" in s:
        return "3. Mobiliario."
    if s in ("4","4."):
        return "4. Integral (1, 2 y 3)"
    # si tiene numeros mezclados intento deducir
    tiene1 = bool(re.search(r'\b1\b', s))
    tiene2 = bool(re.search(r'\b2\b', s))
    tiene3 = bool(re.search(r'\b3\b', s))
    if (tiene1 and tiene2) or (tiene1 and tiene3) or (tiene2 and tiene3):
        return "4. Integral (1, 2 y 3)"
    if tiene1: return "1. Infraestructura."
    if tiene2: return "2. Equipamiento."
    if tiene3: return "3. Mobiliario."
    return ERR

def limpiar_sino(val):
    """para los campos de SI/NO (demol, nueva, etc)"""
    if pd.isna(val) or str(val).strip() == "": return val
    s = re.sub(r'\s+', ' ', str(val).strip().upper().replace("Í","I"))
    if s in ("-","_","P","MO",""): return ERR
    # textos largos que no son SI/NO son basura
    if len(s) > 30 and not s.startswith("SI") and not s.startswith("NO"): return ERR
    # numeros raros
    if re.match(r'^\d+$', s) and s not in ("0","1"): return ERR
    if s.startswith("SI") or s in ("1","S"): return "SI"
    if s.startswith("NO") or s in ("0","N"): return "NO"
    return ERR

def limpiar_cod_mod(val):
    if pd.isna(val) or str(val).strip() == "": return val
    s = str(val).strip()
    # guiones y underscores solos no son error, simplemente no aplica
    if re.match(r'^[\s\-_\.]+$', s): return pd.NA
    # SI/NO/NINGUNA tampoco es error
    if s.upper().replace("Í","I") in ("SI","SÍ","NO","NINGUNA","0"): return pd.NA
    # extraer codigos de 5+ digitos
    codigos = re.findall(r'\d{5,}', s)
    if codigos: return ", ".join(codigos)
    return ERR

# ============================================================================
# PASO 4: Aplicar limpieza
# ============================================================================
print("\n>> Limpiando campos...")

# primero muevo textos largos de elec a comentarios (ahi no deberian estar)
for i in df.index:
    val = df.at[i, "elec"]
    if pd.notna(val) and len(str(val).strip()) > 30:
        if not str(val).strip().upper().startswith(("SI","NO")):
            com = df.at[i, "comentarios"]
            if pd.isna(com) or str(com).strip() == "":
                df.at[i, "comentarios"] = str(val).strip()
            else:
                df.at[i, "comentarios"] = str(com).strip() + " | " + str(val).strip()
            df.at[i, "elec"] = pd.NA

# limpio f9 antes (detectar corrimiento de columnas)
for i in df.index:
    val = df.at[i, "f9"]
    if pd.notna(val):
        s = str(val).strip().upper()
        if any(x in s for x in ["INFRAESTRUCTURA","EQUIPAMIENTO","MOBILIARIO","INTEGRAL"]):
            df.at[i, "f9"] = pd.NA

# ahora si aplico las funciones
df["cui"]   = df["cui"].apply(limpiar_cui)
df["tipo"]  = df["tipo"].apply(limpiar_tipo)
df["monto"] = df["monto"].apply(limpiar_monto)
df["avance"]= df["avance"].apply(limpiar_avance)
df["f9"]    = df["f9"].apply(limpiar_f9)
df["comp"]  = df["comp"].apply(limpiar_comp)

for campo in ["demol","nueva","reforz","cerco","sust","ampl","mobil","agua","elec","unid"]:
    df[campo] = df[campo].apply(limpiar_sino)

df["cod_mod"] = df["cod_mod"].apply(lambda x: limpiar_cod_mod(x) if pd.notna(x) and str(x).strip() != "" else x)

print("   Limpieza de campos completada")

# ============================================================================
# PASO 5: Validar contra bases maestras y corregir
# ============================================================================
print("\n>> Validando contra bases maestras...")

# lista para guardar las correcciones que hago
correcciones = []

# limpio el CUI del anexo para poder cruzar
df["cui_limpio"] = df["cui"].apply(
    lambda x: str(int(float(x))) if pd.notna(x) and re.match(r'^\d+\.?\d*$', str(x).strip()) and str(x) != ERR else str(x).strip() if pd.notna(x) else ""
)

# --- Cruzar con Vinculaciones ---
if vinc is not None:
    # armo un diccionario de CUI -> datos de vinculaciones (tomo el primer registro por CUI)
    vinc_por_cui = vinc.drop_duplicates(subset="cui_limpio", keep="first").set_index("cui_limpio")

    for i in df.index:
        cui = df.at[i, "cui_limpio"]
        if cui == "" or cui == ERR:
            continue
        if cui not in vinc_por_cui.index:
            continue

        fila_vinc = vinc_por_cui.loc[cui]

        # corregir nombre_ie si es diferente
        nombre_vinc = str(fila_vinc.get("Nombre IIEE", "")).strip()
        nombre_anexo = str(df.at[i, "nombre_ie"]).strip() if pd.notna(df.at[i, "nombre_ie"]) else ""
        if nombre_vinc and nombre_anexo != nombre_vinc:
            correcciones.append({
                "cui": cui, "campo": "nombre_ie",
                "valor_original": nombre_anexo,
                "valor_corregido": nombre_vinc,
                "fuente": "Vinculaciones"
            })
            df.at[i, "nombre_ie"] = nombre_vinc

        # completar cod_local si esta vacio
        cod_local_vinc = str(fila_vinc.get("Código Local", "")).strip()
        # limpiar .0 si viene de float
        if cod_local_vinc.endswith(".0"):
            cod_local_vinc = cod_local_vinc[:-2]
        cod_local_anexo = str(df.at[i, "cod_local"]).strip() if pd.notna(df.at[i, "cod_local"]) else ""
        if cod_local_vinc and (cod_local_anexo == "" or pd.isna(df.at[i, "cod_local"])):
            correcciones.append({
                "cui": cui, "campo": "cod_local",
                "valor_original": cod_local_anexo,
                "valor_corregido": cod_local_vinc,
                "fuente": "Vinculaciones"
            })
            df.at[i, "cod_local"] = cod_local_vinc

    print(f"   Vinculaciones: {len([c for c in correcciones if c['fuente']=='Vinculaciones'])} correcciones")

# --- Cruzar con Base de Inversiones ---
if inv is not None and col_cui_inv:
    # detectar columnas disponibles
    def buscar_col(df_inv, opciones):
        for c in opciones:
            if c in df_inv.columns: return c
        return None

    col_tipo = buscar_col(inv, ["DES_TIPO_FORMATO","TIPO_FORMATO","TIPO_INVERSION"])
    col_monto_inv = buscar_col(inv, ["COSTO_ACTUALIZADO_BI","COSTO_INV_TOTAL_BI"])
    col_f9_inv = buscar_col(inv, ["TIENE_F9","F9"])
    col_nombre_inv = buscar_col(inv, ["NOMBRE_INVERSION","NOMBRE"])
    col_estado = buscar_col(inv, ["ESTADO"])
    col_avance_f12b = buscar_col(inv, ["AVANCE_FISICO_F12B"])

    inv_por_cui = inv.drop_duplicates(subset="cui_limpio", keep="first").set_index("cui_limpio")

    # agrego columna de nombre de inversion
    df["nombre_inversion"] = ""

    for i in df.index:
        cui = df.at[i, "cui_limpio"]
        if cui == "" or cui == ERR:
            continue
        if cui not in inv_por_cui.index:
            continue

        fila_inv = inv_por_cui.loc[cui]

        # nombre de la inversion (siempre lo agrego)
        if col_nombre_inv:
            nombre = str(fila_inv.get(col_nombre_inv, "")).strip()
            if nombre: df.at[i, "nombre_inversion"] = nombre

        # si tipo esta vacio o con error, intento completar desde inversiones
        tipo_actual = str(df.at[i, "tipo"]).strip()
        if tipo_actual in ("", ERR, "nan") and col_tipo:
            tipo_inv = str(fila_inv.get(col_tipo, "")).strip().upper()
            tipo_nuevo = ""
            if "IOARR" in tipo_inv: tipo_nuevo = "IOARR"
            elif "PROYECTO" in tipo_inv: tipo_nuevo = "PI"
            elif "IRI" in tipo_inv: tipo_nuevo = "IRI"
            if tipo_nuevo:
                correcciones.append({
                    "cui": cui, "campo": "tipo",
                    "valor_original": tipo_actual,
                    "valor_corregido": tipo_nuevo,
                    "fuente": "Base Inversiones"
                })
                df.at[i, "tipo"] = tipo_nuevo

        # si f9 esta vacio intento completar
        f9_actual = str(df.at[i, "f9"]).strip()
        if f9_actual in ("", ERR, "nan") and col_f9_inv:
            f9_inv = str(fila_inv.get(col_f9_inv, "")).strip().upper()
            if f9_inv in ("SI","NO"):
                correcciones.append({
                    "cui": cui, "campo": "f9",
                    "valor_original": f9_actual,
                    "valor_corregido": f9_inv,
                    "fuente": "Base Inversiones"
                })
                df.at[i, "f9"] = f9_inv

    print(f"   Base Inversiones: {len([c for c in correcciones if c['fuente']=='Base Inversiones'])} correcciones")

print(f"   Total correcciones: {len(correcciones)}")

# ============================================================================
# PASO 6: Clasificar comentarios
# ============================================================================
print("\n>> Clasificando comentarios...")

def clasificar_comentario(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    s = str(val).strip().upper()

    # las categorias van de mas especifica a mas general
    if re.search(r'CULMINA|FINALIZ|TERMINAD|100\s*%', s):
        return "OBRA CULMINADA"
    if re.search(r'NO EJECUT|SIN EJECUCI', s):
        return "NO EJECUTADO"
    if re.search(r'PARALIZ|SUSPENDID', s):
        return "PARALIZADA/SUSPENDIDA"
    if re.search(r'ARBITRAJE', s):
        return "EN ARBITRAJE"
    if re.search(r'LIQUIDACI', s):
        return "EN LIQUIDACION"
    if re.search(r'CIERRE|F9|FORMATO 9|FORMATO9', s):
        return "EN PROCESO DE CIERRE (F9)"
    if re.search(r'EN EJECUCI|EJECUTANDO|EJECUCION FISICA', s):
        return "EN EJECUCION"
    if re.search(r'EXPEDIENTE|PERFIL|E\.?T\.?|ESTUDIO', s):
        return "EXPEDIENTE TECNICO/PERFIL"
    if re.search(r'NO INICIAD|SIN INICIO|NO HA INIC', s):
        return "SIN INICIAR"
    if re.search(r'FINANCIAMIENTO|SIN PRESUPUESTO|RECURSO', s):
        return "SIN FINANCIAMIENTO"
    if re.search(r'INTERVINO \d+ IE|INDIVIDUALIZAR|COBERTURA.*IE|COMPRENDE.*IE', s):
        return "MULTIPLES IE (CUI AGRUPADO)"
    if re.search(r'DEPORT|LOSA|TRIBUNA|CANCHA|ESTADIO', s):
        return "INFRAESTRUCTURA DEPORTIVA"
    if re.search(r'SOLAR|TENSIONAR', s):
        return "PROTECCION SOLAR"
    if re.search(r'MODULO|CAPACITACI|TRANSFERENCIA TEC', s):
        return "MODULOS EDUCATIVOS"
    if re.search(r'EQUIP|MOBILIAR|COMPUT|LAPTOP', s):
        return "EQUIPAMIENTO/MOBILIARIO"
    if re.search(r'CERCO|PERIMÉTRIC|PERIMETRIC', s):
        return "CERCO PERIMETRICO"
    if re.search(r'AGUA|DESAGUE|SANEAMIENT|ALCANTARILL', s):
        return "SERVICIOS BASICOS (AGUA)"
    if re.search(r'ELECTRI|ENERGIA|LUZ', s):
        return "SERVICIOS BASICOS (ELECTRICIDAD)"
    if re.search(r'DEMOL|REFORZAMIENT', s):
        return "DEMOLICION/REFORZAMIENTO"
    if re.search(r'AMPLIA', s):
        return "AMPLIACION"
    if re.search(r'CONSTRUCCI', s):
        return "CONSTRUCCION NUEVA"
    return "OTRO"

df["tipo_comentario"] = df["comentarios"].apply(clasificar_comentario)

# ============================================================================
# PASO 7: Marcar duplicados, renumerar, detectar errores y excluidos
# ============================================================================
print("\n>> Clasificando registros...")

# duplicados de cod_local+cui (un local puede tener varios CUI, eso es valido)
dup = df.duplicated(subset=["cod_local","cui"], keep=False) & df["cod_local"].notna()
df["cod_local_dup"] = ""
df.loc[dup, "cod_local_dup"] = "SI"

# renumerar
df["nro"] = range(1, len(df) + 1)

# campos que deben estar llenos si o si (menos comentarios, cod_mod, fecha, nro)
campos_check = ["cui","tipo","monto","avance","f9","comp","unid",
                "demol","nueva","reforz","cerco","sust","ampl","mobil","agua","elec"]
campos_requeridos = ["cod_local","region","provincia","distrito","nombre_ie"] + campos_check

# marcar errores por campo
for c in campos_check:
    df["err_" + c] = (df[c] == ERR).astype(int)

# tiene algun error?
cols_err = [f"err_{c}" for c in campos_check]
df["tiene_error"] = df[cols_err].max(axis=1)

# campos vacios (requeridos que quedaron sin llenar despues de correcciones)
def campos_incompletos(fila):
    vacios = []
    for c in campos_requeridos:
        val = fila.get(c)
        if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "nan":
            vacios.append(c)
    return vacios

df["campos_faltantes"] = df.apply(lambda f: ", ".join(campos_incompletos(f)), axis=1)
df["esta_completo"] = df["campos_faltantes"].apply(lambda x: x == "")

# clasificar en 3 grupos
# limpio = sin errores Y completo
# con errores = tiene algun campo con error pero tiene datos
# excluido = le faltan campos requeridos
df["status"] = "LIMPIO"
df.loc[df["tiene_error"] == 1, "status"] = "CON ERRORES"
df.loc[~df["esta_completo"], "status"] = "EXCLUIDO"

# reemplazar el flag por texto legible en los que tienen error
for c in campos_check:
    df.loc[df[c] == ERR, c] = "<<ERROR>>"

n_limpio = (df["status"] == "LIMPIO").sum()
n_error = (df["status"] == "CON ERRORES").sum()
n_excluido = (df["status"] == "EXCLUIDO").sum()
print(f"   Limpios: {n_limpio}")
print(f"   Con errores: {n_error}")
print(f"   Excluidos (incompletos): {n_excluido}")

# ============================================================================
# PASO 8: Armar reportes
# ============================================================================
print("\n>> Generando reportes...")

# --- Reporte de correcciones ---
df_correcciones = pd.DataFrame(correcciones)

# --- Status por region ---
# incluyo tambien los pendientes (sin CUI) para el reporte completo
df_total = pd.read_excel(archivo_anexo1, sheet_name="Hoja1", header=0, dtype=str)
df_total = df_total.drop(index=[0,1]).reset_index(drop=True)
col_region = df_total.columns[2]
col_cui_total = df_total.columns[6]
df_total["tiene_cui"] = df_total[col_cui_total].notna() & (df_total[col_cui_total].str.strip() != "")

# cuento por region en el total
status_region = df_total.groupby(col_region).agg(
    total_locales=("tiene_cui", "count"),
    con_cui=("tiene_cui", "sum")
).reset_index()
status_region.columns = ["region", "total_locales", "con_info"]
status_region["pendientes"] = status_region["total_locales"] - status_region["con_info"]
status_region["pct_avance"] = round(status_region["con_info"] / status_region["total_locales"] * 100, 1)

# agrego el desglose de los que tienen info (limpio/error/excluido)
desglose = df.groupby("region")["status"].value_counts().unstack(fill_value=0).reset_index()
for col in ["LIMPIO","CON ERRORES","EXCLUIDO"]:
    if col not in desglose.columns:
        desglose[col] = 0
status_region = status_region.merge(desglose[["region","LIMPIO","CON ERRORES","EXCLUIDO"]], on="region", how="left")
status_region = status_region.fillna(0)

# --- Errores por campo ---
resumen_errores = []
for c in campos_check:
    n = (df["err_" + c] == 1).sum()
    if n > 0:
        resumen_errores.append({"campo": c, "cantidad_errores": n})
df_errores_campo = pd.DataFrame(resumen_errores)

# ============================================================================
# PASO 9: Graficos
# ============================================================================
print("\n>> Generando graficos...")

# colores que uso siempre
verde = "#2ecc71"
amarillo = "#f39c12"
rojo = "#e74c3c"
gris = "#95a5a6"
azul = "#3498db"

try:
    # grafico 1: status por region (las 10 con mas registros)
    top10 = status_region.nlargest(10, "total_locales").sort_values("total_locales", ascending=True)
    fig, ax = plt.subplots(figsize=(12, 7))
    y = range(len(top10))
    ax.barh(y, top10["LIMPIO"], color=verde, label="Limpios")
    ax.barh(y, top10["CON ERRORES"], left=top10["LIMPIO"], color=amarillo, label="Con errores")
    ax.barh(y, top10["EXCLUIDO"], left=top10["LIMPIO"]+top10["CON ERRORES"], color=rojo, label="Excluidos")
    ax.barh(y, top10["pendientes"], left=top10["LIMPIO"]+top10["CON ERRORES"]+top10["EXCLUIDO"], color=gris, label="Pendientes")
    ax.set_yticks(y)
    ax.set_yticklabels(top10["region"], fontsize=9)
    ax.set_xlabel("Cantidad de locales educativos")
    ax.set_title("Status de envío por región (Top 10)", fontsize=13, fontweight="bold")
    ax.legend(loc="lower right")
    plt.tight_layout()
    plt.savefig(salida / "Anexo1_grafico_status.png", dpi=150)
    plt.close()
    print("   Grafico status guardado")
except Exception as e:
    print(f"   Error en grafico status: {e}")

try:
    # grafico 2: categorias de comentarios
    cats = df[df["tipo_comentario"] != ""]["tipo_comentario"].value_counts()
    if len(cats) > 0:
        fig, ax = plt.subplots(figsize=(10, 7))
        colores_cat = plt.cm.Set3(range(len(cats)))
        cats.plot.barh(ax=ax, color=colores_cat)
        ax.set_xlabel("Cantidad")
        ax.set_title("Distribución de categorías de comentarios", fontsize=13, fontweight="bold")
        ax.invert_yaxis()
        plt.tight_layout()
        plt.savefig(salida / "Anexo1_grafico_comentarios.png", dpi=150)
        plt.close()
        print("   Grafico comentarios guardado")
except Exception as e:
    print(f"   Error en grafico comentarios: {e}")

try:
    # grafico 3: errores por campo
    if len(df_errores_campo) > 0:
        fig, ax = plt.subplots(figsize=(10, 6))
        df_errores_campo_sorted = df_errores_campo.sort_values("cantidad_errores", ascending=True)
        ax.barh(df_errores_campo_sorted["campo"], df_errores_campo_sorted["cantidad_errores"], color=rojo)
        ax.set_xlabel("Cantidad de errores")
        ax.set_title("Errores por campo", fontsize=13, fontweight="bold")
        plt.tight_layout()
        plt.savefig(salida / "Anexo1_grafico_errores.png", dpi=150)
        plt.close()
        print("   Grafico errores guardado")
except Exception as e:
    print(f"   Error en grafico errores: {e}")

# ============================================================================
# PASO 10: Exportar todo
# ============================================================================
print("\n>> Exportando archivos...")

# diccionario de datos
dicc = pd.DataFrame([
    ("nro", "Numero correlativo"),
    ("cod_local", "Codigo de local educativo"),
    ("region", "Region"),
    ("provincia", "Provincia"),
    ("distrito", "Distrito"),
    ("nombre_ie", "Nombre de la IE (corregido desde Vinculaciones si aplica)"),
    ("cui", "Codigo Unico de Inversion - solo numeros"),
    ("tipo", "Tipo: PI / IOARR / IRI"),
    ("monto", "Monto de inversion en soles"),
    ("avance", "Avance fisico (0-100%)"),
    ("fecha", "Fecha de recepcion de obra"),
    ("f9", "Tiene Formato 9 (SI/NO)"),
    ("comp", "Componente: 1.Infraestructura / 2.Equipamiento / 3.Mobiliario / 4.Integral"),
    ("unid", "Intervino todas las unidades productoras (SI/NO)"),
    ("cod_mod", "Codigos modulares intervenidos"),
    ("demol", "Demolicion (SI/NO)"),
    ("nueva", "Construccion nueva (SI/NO)"),
    ("reforz", "Reforzamiento (SI/NO)"),
    ("cerco", "Cerco perimetrico (SI/NO)"),
    ("sust", "Sustitucion (SI/NO)"),
    ("ampl", "Ampliacion (SI/NO)"),
    ("mobil", "Mobiliario (SI/NO)"),
    ("agua", "Agua y desague (SI/NO)"),
    ("elec", "Electricidad (SI/NO)"),
    ("comentarios", "Comentarios (texto libre)"),
    ("nombre_inversion", "Nombre del proyecto (desde Base Inversiones)"),
    ("tipo_comentario", "Categoria del comentario"),
    ("cod_local_dup", "SI si cod_local+CUI esta duplicado"),
], columns=["Campo", "Descripcion"])

# columnas para exportar (sin las auxiliares)
cols_export = [c for c in df.columns if not c.startswith("err_") and c not in
               ("tiene_error","campos_faltantes","esta_completo","status","cui_limpio")]

# base limpia
df_limpia = df[df["status"] == "LIMPIO"][cols_export]
with pd.ExcelWriter(salida / "Anexo1_base_limpia.xlsx", engine="openpyxl") as w:
    df_limpia.to_excel(w, index=False, sheet_name="Base Limpia")
    dicc.to_excel(w, index=False, sheet_name="Diccionario de Datos")
print(f"   Base limpia: {len(df_limpia)} registros")

# base con errores
df_errores = df[df["status"] == "CON ERRORES"].copy()
with pd.ExcelWriter(salida / "Anexo1_base_errores.xlsx", engine="openpyxl") as w:
    df_errores[cols_export + [f"err_{c}" for c in campos_check]].to_excel(w, index=False, sheet_name="Registros con Errores")
    dicc.to_excel(w, index=False, sheet_name="Diccionario de Datos")
print(f"   Base errores: {len(df_errores)} registros")

# base excluidos
df_excluidos = df[df["status"] == "EXCLUIDO"].copy()
with pd.ExcelWriter(salida / "Anexo1_base_excluidos.xlsx", engine="openpyxl") as w:
    df_excluidos[cols_export + ["campos_faltantes"]].to_excel(w, index=False, sheet_name="Excluidos")
    dicc.to_excel(w, index=False, sheet_name="Diccionario de Datos")
print(f"   Base excluidos: {len(df_excluidos)} registros")

# reporte general
with pd.ExcelWriter(salida / "Anexo1_reporte.xlsx", engine="openpyxl") as w:
    if len(df_correcciones) > 0:
        df_correcciones.to_excel(w, index=False, sheet_name="Correcciones")
    status_region.to_excel(w, index=False, sheet_name="Status por Region")
    if len(df_errores_campo) > 0:
        df_errores_campo.to_excel(w, index=False, sheet_name="Errores por Campo")
    # resumen general
    resumen = pd.DataFrame([
        ("Total locales en Anexo 1", len(df_total)),
        ("Con CUI registrado", total_con_cui),
        ("Pendientes (sin CUI)", len(df_total) - total_con_cui),
        ("", ""),
        ("Registros limpios", n_limpio),
        ("Registros con errores", n_error),
        ("Registros excluidos", n_excluido),
        ("", ""),
        ("Correcciones realizadas", len(correcciones)),
    ], columns=["Concepto", "Cantidad"])
    resumen.to_excel(w, index=False, sheet_name="Resumen")
print("   Reporte generado")

# ============================================================================
# RESUMEN FINAL
# ============================================================================
print("\n" + "=" * 50)
print("RESUMEN")
print("=" * 50)
print(f"Total locales: {len(df_total):,}")
print(f"Con CUI: {total_con_cui:,} ({round(total_con_cui/len(df_total)*100,1)}%)")
print(f"  - Limpios: {n_limpio}")
print(f"  - Con errores: {n_error}")
print(f"  - Excluidos: {n_excluido}")
print(f"Correcciones desde maestras: {len(correcciones)}")
print(f"\nArchivos en: {salida}")
