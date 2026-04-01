"""
LIMPIEZA - Anexo 01: Relación de locales educativos intervenidos
================================================================
Replica exactamente el flujo del do-file Stata:

  Stata:
    import excel "...Anexo_1_avance_GR_GL.xlsx", sheet("Hoja1") firstrow clear
    → Lee fila 1 (título del Anexo) como nombres de variables
    → Observations: 55,714  |  Variables: 29
    drop in 1/2  → elimina fila vacía y fila de headers reales
    rename ...   → asigna nombres cortos operativos

  Python replica:
    pd.read_excel(..., header=0, sheet_name="Hoja1")
    → misma lógica: fila 1 como header, filas 2-55715 como datos (55,714 obs)
    drop primeras 2 filas (vacía + headers reales)
    rename con mismo diccionario de nombres cortos
"""

import pandas as pd
import re
from pathlib import Path

# ── 0. RUTAS (equivalente a los globals de Stata) ─────────────────────────────
reporte    = Path(r"C:\Users\iraid\Documents\DISCO TERA\MINEDU\TRABAJO-MINEDU\01_MINEDU")
rep_cons_i = reporte / "01_input"
rep_cons_o = reporte / "03_output"
rep_cons_t = reporte / "04_temporal"

INPUT      = rep_cons_i / "Anexo_1_avance_GR_GL.xlsx"
OUT_CLEAN  = rep_cons_o / "Anexo1_base_limpia.xlsx"
OUT_ERRORS = rep_cons_o / "Anexo1_base_errores.xlsx"

# ── 1. IMPORTAR (equivalente a: import excel ..., firstrow clear) ─────────────
# header=0 → fila 1 del Excel como nombres de columna (igual que Stata firstrow)
# dtype=str → todo como texto para preservar los valores tal como Stata string vars
df = pd.read_excel(INPUT, sheet_name="Hoja1", header=0, dtype=str)
print(f"Importado: {df.shape[0]:,} obs, {df.shape[1]} vars")  # → 55,714 obs, 29 vars

# ── 2. DROP IN 1/2 (eliminar fila vacía y fila con headers reales) ─────────────
df = df.drop(index=[0, 1]).reset_index(drop=True)
print(f"Tras drop in 1/2: {df.shape[0]:,} obs")

# ── 3. RENAME (equivalente al rename masivo del do-file) ──────────────────────
# Col 0  = título largo del Anexo (merged en Excel)
# Col 1-24 = Unnamed: 1 .. Unnamed: 24
# Col 25-28 = columnas extra vacías → se eliminan (drop extra1-extra4)
rename_map = {
    df.columns[0] : "nro",
    df.columns[1] : "cod_local",
    df.columns[2] : "region",
    df.columns[3] : "provincia",
    df.columns[4] : "distrito",
    df.columns[5] : "nombre_ie",
    df.columns[6] : "cui",
    df.columns[7] : "tipo",
    df.columns[8] : "monto",
    df.columns[9] : "avance",
    df.columns[10]: "fecha",
    df.columns[11]: "f9",
    df.columns[12]: "comp",
    df.columns[13]: "unid",
    df.columns[14]: "cod_mod",
    df.columns[15]: "demol",
    df.columns[16]: "nueva",
    df.columns[17]: "reforz",
    df.columns[18]: "cerco",
    df.columns[19]: "sust",
    df.columns[20]: "ampl",
    df.columns[21]: "mobil",
    df.columns[22]: "agua",
    df.columns[23]: "elec",
    df.columns[24]: "comentarios",
}
df = df.rename(columns=rename_map)
# Eliminar columnas extra (cols 25-28) igual que: drop extra1 extra2 extra3 extra4
extra_cols = [c for c in df.columns if c not in rename_map.values()]
df = df.drop(columns=extra_cols)

# ── 4. ELIMINAR FILAS SIN CUI ──────────────────────────────────────────────────
df = df[df["cui"].notna() & (df["cui"].str.strip() != "")].copy()
print(f"Filas con CUI: {len(df):,}")

# Columna de error por campo (0=ok, 1=error) — equivalente a gen byte err_* = 0
for field in ["cui","tipo","monto","avance","f9","comp","unid",
              "demol","nueva","reforz","cerco","sust","ampl","mobil","agua","elec"]:
    df[f"err_{field}"] = 0

ERROR_FLAG = "__ERROR__"


# ── 5. FUNCIONES DE NORMALIZACIÓN ─────────────────────────────────────────────

def norm_cui(v):
    """Solo dígitos. Intenta rescatar extrayendo secuencia de 5+ dígitos."""
    if pd.isna(v): return v
    s = str(v).strip()
    if re.match(r"^\d+$", s): return s
    m = re.search(r"\d{5,}", s)
    return m.group() if m else ERROR_FLAG


def norm_tipo(v):
    """Normaliza a PI / IOARR / IRI."""
    if pd.isna(v): return v
    s = re.sub(r"\s+", " ", str(v).strip().upper())
    if re.search(r"\bIOARR?\b|\bFUR\b", s): return "IOARR"
    if re.search(r"\bIRI\b", s):            return "IRI"
    if re.search(r"\bPI\b|PROYECTO", s):    return "PI"
    return ERROR_FLAG


def norm_monto(v):
    """Número positivo. Acepta coma como separador decimal."""
    if pd.isna(v): return v
    s = str(v).strip().replace(",", ".")
    try:
        f = float(s)
        return str(f) if f >= 0 else ERROR_FLAG
    except ValueError:
        return ERROR_FLAG


def norm_avance(v):
    """Porcentaje 0–100. Acepta '96.10%', proporción 0.976 → 97.6."""
    if pd.isna(v): return v
    s = str(v).strip().replace("%", "").replace(",", ".")
    try:
        f = float(s)
        if 0 <= f <= 1:    return str(round(f * 100, 2))
        if 0 <= f <= 100:  return str(round(f, 2))
        return ERROR_FLAG
    except ValueError:
        return ERROR_FLAG


def norm_f9(v):
    """SI o NO."""
    if pd.isna(v): return v
    s = str(v).strip().upper().replace("Í","I")
    if s in {"SI","S","1","SÍ"}: return "SI"
    if s in {"NO","N","0","NO "}: return "NO"
    return ERROR_FLAG


COMP_MAP = {
    # solo infraestructura
    "1":"1. Infraestructura.", "1.":"1. Infraestructura.",
    "1. INFRAESTRUCTURA.":"1. Infraestructura.",
    "1.INFRAESTRUCTURA":"1. Infraestructura.", "INFRAESTRUCTURA":"1. Infraestructura.",
    # solo equipamiento
    "2":"2. Equipamiento.", "2.":"2. Equipamiento.",
    "2. EQUIPAMIENTO.":"2. Equipamiento.", "2.EQUIPAMIENTO":"2. Equipamiento.",
    "2. EQUIPAMIENT":"2. Equipamiento.",
    # solo mobiliario
    "3":"3. Mobiliario.", "3.":"3. Mobiliario.", "3. MOBILIARIO.":"3. Mobiliario.",
    # integral (y combos)
    "4":"4. Integral (1, 2 y 3)", "4. INTEGRAL (1, 2 Y 3)":"4. Integral (1, 2 y 3)",
    "INTEGRAL":"4. Integral (1, 2 y 3)", "(1,2 Y 3)":"4. Integral (1, 2 y 3)",
    "1,2,3":"4. Integral (1, 2 y 3)", "1, 2 Y 3":"4. Integral (1, 2 y 3)",
    "1 Y 2 Y 3":"4. Integral (1, 2 y 3)",
    "1,2":"4. Integral (1, 2 y 3)", "1, 2":"4. Integral (1, 2 y 3)",
    "1 Y 2":"4. Integral (1, 2 y 3)", "1,3":"4. Integral (1, 2 y 3)",
    "1 Y 3":"4. Integral (1, 2 y 3)", "2,3":"4. Integral (1, 2 y 3)",
    "2 Y 3":"4. Integral (1, 2 y 3)", "2. EQUIPAMIENTO 3.MOBILIARIO":"4. Integral (1, 2 y 3)",
    "1.2":"4. Integral (1, 2 y 3)", "1.3":"4. Integral (1, 2 y 3)",
    "2.3":"4. Integral (1, 2 y 3)",
}

def norm_comp(v):
    if pd.isna(v): return v
    s = re.sub(r"\s+", " ", str(v).strip().upper())
    if s in COMP_MAP: return COMP_MAP[s]
    has1 = bool(re.search(r"\b1\b|INFRA", s))
    has2 = bool(re.search(r"\b2\b|EQUIP", s))
    has3 = bool(re.search(r"\b3\b|MOBIL", s))
    has4 = bool(re.search(r"\b4\b|INTEGR", s))
    if has4 or (has1 and has2 and has3): return "4. Integral (1, 2 y 3)"
    if (has1 and has2) or (has1 and has3) or (has2 and has3): return "4. Integral (1, 2 y 3)"
    if has1: return "1. Infraestructura."
    if has2: return "2. Equipamiento."
    if has3: return "3. Mobiliario."
    return ERROR_FLAG


def norm_sino(v):
    """SI o NO — acepta SI (PARCIAL), SI-TOTAL, etc. → SI."""
    if pd.isna(v): return v
    s = re.sub(r"\s+", " ", str(v).strip().upper().replace("Í","I"))
    if re.match(r"^SI\b", s) or s in {"1","S","SÍ"}: return "SI"
    if re.match(r"^NO\b", s) or s in {"0","N"}:       return "NO"
    return ERROR_FLAG


# ── 6. APLICAR NORMALIZACIÓN ──────────────────────────────────────────────────
SINO_VARS = ["demol","nueva","reforz","cerco","sust","ampl","mobil","agua","elec","unid"]

df["cui"]   = df["cui"].apply(norm_cui)
df["tipo"]  = df["tipo"].apply(norm_tipo)
df["monto"] = df["monto"].apply(norm_monto)
df["avance"]= df["avance"].apply(norm_avance)
df["f9"]    = df["f9"].apply(norm_f9)
df["comp"]  = df["comp"].apply(norm_comp)
for v in SINO_VARS:
    df[v] = df[v].apply(norm_sino)


# ── 7. MARCAR ERRORES POR CAMPO ───────────────────────────────────────────────
CHECK_VARS = ["cui","tipo","monto","avance","f9","comp"] + SINO_VARS

for v in CHECK_VARS:
    df[f"err_{v}"] = df[v].apply(lambda x: 1 if str(x) == ERROR_FLAG else 0)

df["tiene_error"] = df[[f"err_{v}" for v in CHECK_VARS]].max(axis=1)


# ── 8. SEPARAR BASES ──────────────────────────────────────────────────────────
df_clean  = df[df["tiene_error"] == 0].drop(columns=[f"err_{v}" for v in CHECK_VARS] + ["tiene_error"])
df_errors = df[df["tiene_error"] == 1].copy()

# Reemplazar flag por etiqueta legible
for v in CHECK_VARS:
    df_errors[v] = df_errors[v].replace(ERROR_FLAG, "ERROR")

print(f"\nRegistros limpios : {len(df_clean):,}")
print(f"Registros con error: {len(df_errors):,}")

# Resumen de errores por campo
print("\n=== RESUMEN DE ERRORES POR CAMPO ===")
for v in CHECK_VARS:
    n = df[f"err_{v}"].sum()
    if n > 0:
        print(f"  err_{v}: {n} registros")


# ── 9. EXPORTAR ───────────────────────────────────────────────────────────────
with pd.ExcelWriter(OUT_CLEAN, engine="openpyxl") as w:
    df_clean.to_excel(w, index=False, sheet_name="Base Limpia")

with pd.ExcelWriter(OUT_ERRORS, engine="openpyxl") as w:
    df_errors.to_excel(w, index=False, sheet_name="Registros con Errores")

print(f"\nArchivos generados:")
print(f"  {OUT_CLEAN}")
print(f"  {OUT_ERRORS}")
