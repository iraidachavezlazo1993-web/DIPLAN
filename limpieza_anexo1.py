"""
LIMPIEZA - Anexo 01: Relación de locales educativos intervenidos
================================================================
Limpia los campos declarados por GR y GL en el Anexo 1.
Genera: base limpia, base con errores, y reporte de completitud.
"""

import pandas as pd
import re
from pathlib import Path

# ── 0. RUTAS ──────────────────────────────────────────────────────────────────
import platform
if platform.system() == "Windows":
    ruta = Path(r"C:\Users\diplan11\Documents\00_MINEDU\TRABAJO-MINEDU\SOL_GR_GL")
    entrada = ruta / "01_input"
    salida  = ruta / "03_output"
    temporal = ruta / "04_temporal"
else:
    entrada = Path(__file__).resolve().parent
    salida  = entrada
    temporal = entrada

INPUT = entrada / "Anexo_1_avance_GR_GL.xlsx"
INPUT_INV = entrada / "2026.03.23 Base de Inversiones_.xlsx"
OUT_CLEAN = salida / "Anexo1_base_limpia.xlsx"
OUT_ERRORS = salida / "Anexo1_base_errores.xlsx"
OUT_PENDIENTES = salida / "Anexo1_pendientes_registro.xlsx"
OUT_REPORTE = salida / "Anexo1_reporte_completitud.xlsx"

print("=" * 60)
print("LIMPIEZA DEL ANEXO 1 - Inversiones GR/GL")
print("=" * 60)

# ── 1. IMPORTAR ──────────────────────────────────────────────────────────────
df = pd.read_excel(INPUT, sheet_name="Hoja1", header=0, dtype=str)
print(f"\n>> Importado: {df.shape[0]:,} filas, {df.shape[1]} columnas")

df = df.drop(index=[0, 1]).reset_index(drop=True)
print(f"   Tras eliminar filas de encabezado: {df.shape[0]:,} filas")

# ── 2. RENOMBRAR COLUMNAS ────────────────────────────────────────────────────
NOMBRES = [
    "nro", "cod_local", "region", "provincia", "distrito", "nombre_ie",
    "cui", "tipo", "monto", "avance", "fecha", "f9", "comp", "unid",
    "cod_mod", "demol", "nueva", "reforz", "cerco", "sust", "ampl",
    "mobil", "agua", "elec", "comentarios",
]

rename_map = {}
for i, name in enumerate(NOMBRES):
    if i < len(df.columns):
        rename_map[df.columns[i]] = name
df = df.rename(columns=rename_map)

extra_cols = [c for c in df.columns if c not in NOMBRES]
if extra_cols:
    df = df.drop(columns=extra_cols)

# ── 3. SEPARAR: con CUI (registrados) vs sin CUI (pendientes) ───────────────
tiene_cui = df["cui"].notna() & (df["cui"].str.strip() != "")
df_pendientes = df[~tiene_cui].copy()
df = df[tiene_cui].copy()

print(f"\n>> Filas con CUI (registrados por GR/GL): {len(df):,}")
print(f"   Filas sin CUI (pendientes de registro): {len(df_pendientes):,}")

# ── 4. CAMPOS A LIMPIAR ─────────────────────────────────────────────────────
CAMPOS_GR_GL = [
    "cui", "tipo", "monto", "avance", "fecha", "f9", "comp", "unid",
    "cod_mod", "demol", "nueva", "reforz", "cerco", "sust", "ampl",
    "mobil", "agua", "elec", "comentarios",
]

ERROR_FLAG = "__ERROR__"

df["region"] = df["region"].str.strip().str.upper()
df_pendientes["region"] = df_pendientes["region"].str.strip().str.upper()

for field in CAMPOS_GR_GL:
    if field != "comentarios" and field != "cod_mod" and field != "fecha":
        df[f"err_{field}"] = 0


# ── 5. FUNCIONES DE NORMALIZACIÓN ────────────────────────────────────────────

def norm_cod_local(v):
    """Código local: 6 dígitos, pad con ceros a la izquierda. Letras = missing."""
    if pd.isna(v) or str(v).strip() == "":
        return v
    s = re.sub(r"\.0+$", "", str(v).strip())
    s = re.sub(r"\s+", "", s)
    if not re.match(r"^\d+$", s):
        return None
    return s.zfill(6)


def norm_cui(v):
    """CUI: solo números. Intenta rescatar secuencia de 5+ dígitos."""
    if pd.isna(v):
        return v
    s = str(v).strip()
    s = re.sub(r"\.0+$", "", s)
    if re.match(r"^\d+$", s):
        return s
    m = re.search(r"\d{5,}", s)
    return m.group() if m else ERROR_FLAG


def norm_cod_mod(v):
    """Código modular: 7 dígitos cada uno, separados por /.
    Múltiples códigos se unifican con /. Letras = missing."""
    if pd.isna(v) or str(v).strip() == "":
        return v
    s = str(v).strip()
    if s in {"-", "_", "--", "---", "-----------------------"}:
        return None
    s_upper = s.upper()
    if s_upper in {"SI", "SÍ", "NO", "NINGUNA", "0"}:
        return None
    if re.search(r"[a-zA-Z]", s):
        return None
    s = s.replace("/", ",").replace(" - ", ",")
    s = s.replace("\n", ",").replace("\r", ",")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r",+", ",", s)
    s = s.strip(",")
    if not re.match(r"^[\d,]+$", s):
        return None
    codigos = []
    for cod in s.split(","):
        cod = cod.strip()
        if cod:
            cod = re.sub(r"\.0+$", "", cod)
            cod = cod.zfill(7)
            codigos.append(cod)
    return "/".join(codigos) if codigos else None


def norm_tipo(v):
    if pd.isna(v) or str(v).strip() == "":
        return v
    s = re.sub(r"\s+", " ", str(v).strip().upper())
    if re.search(r"\bIOARR?\b|\bFUR\b", s):
        return "IOARR"
    if re.search(r"\bIRI\b", s):
        return "IRI"
    if re.search(r"\bPI\b|PROYECTO", s):
        return "PI"
    return ERROR_FLAG


def norm_monto(v):
    if pd.isna(v) or str(v).strip() == "":
        return v
    s = str(v).strip().replace(",", ".").replace(" ", "")
    s = re.sub(r"[sS]/\.?\s*", "", s)
    try:
        f = float(s)
        return str(round(f, 2)) if f >= 0 else ERROR_FLAG
    except ValueError:
        return ERROR_FLAG


def norm_avance(v):
    if pd.isna(v) or str(v).strip() == "":
        return v
    s = str(v).strip().replace("%", "").replace(",", ".")
    try:
        f = float(s)
        if 0 <= f <= 1:
            return str(round(f * 100, 2))
        if 0 <= f <= 100:
            return str(round(f, 2))
        return ERROR_FLAG
    except ValueError:
        return ERROR_FLAG


def norm_f9(v):
    if pd.isna(v) or str(v).strip() == "":
        return v
    s = str(v).strip().upper()
    s = s.replace("Í", "I").replace("í", "i")
    s = s.strip()
    if s in {"SI", "S", "1", "SÍ", "SI.", "YES"}:
        return "SI"
    if s in {"NO", "N", "0", "NO.", "NO "}:
        return "NO"
    if re.match(r"^SI\b", s):
        return "SI"
    if re.match(r"^NO\b", s):
        return "NO"
    return ERROR_FLAG


COMP_MAP = {
    "1": "1. Infraestructura.",
    "1.": "1. Infraestructura.",
    "1. INFRAESTRUCTURA.": "1. Infraestructura.",
    "1.INFRAESTRUCTURA": "1. Infraestructura.",
    "INFRAESTRUCTURA": "1. Infraestructura.",
    "2": "2. Equipamiento.",
    "2.": "2. Equipamiento.",
    "2. EQUIPAMIENTO.": "2. Equipamiento.",
    "2.EQUIPAMIENTO": "2. Equipamiento.",
    "2. EQUIPAMIENT": "2. Equipamiento.",
    "EQUIPAMIENTO": "2. Equipamiento.",
    "3": "3. Mobiliario.",
    "3.": "3. Mobiliario.",
    "3. MOBILIARIO.": "3. Mobiliario.",
    "MOBILIARIO": "3. Mobiliario.",
    "4": "4. Integral (1, 2 y 3)",
    "4.": "4. Integral (1, 2 y 3)",
    "4. INTEGRAL (1, 2 Y 3)": "4. Integral (1, 2 y 3)",
    "INTEGRAL": "4. Integral (1, 2 y 3)",
    "(1,2 Y 3)": "4. Integral (1, 2 y 3)",
    "1,2,3": "4. Integral (1, 2 y 3)",
    "1, 2, 3": "4. Integral (1, 2 y 3)",
    "1, 2 Y 3": "4. Integral (1, 2 y 3)",
    "1 Y 2 Y 3": "4. Integral (1, 2 y 3)",
    "1,2": "4. Integral (1, 2 y 3)",
    "1, 2": "4. Integral (1, 2 y 3)",
    "1 Y 2": "4. Integral (1, 2 y 3)",
    "1,3": "4. Integral (1, 2 y 3)",
    "1 Y 3": "4. Integral (1, 2 y 3)",
    "2,3": "4. Integral (1, 2 y 3)",
    "2 Y 3": "4. Integral (1, 2 y 3)",
    "2. EQUIPAMIENTO 3.MOBILIARIO": "4. Integral (1, 2 y 3)",
    "1.2": "4. Integral (1, 2 y 3)",
    "1.3": "4. Integral (1, 2 y 3)",
    "2.3": "4. Integral (1, 2 y 3)",
}


def norm_comp(v):
    if pd.isna(v) or str(v).strip() == "":
        return v
    s = re.sub(r"\s+", " ", str(v).strip().upper())
    if s in COMP_MAP:
        return COMP_MAP[s]
    has1 = bool(re.search(r"\b1\b|INFRA", s))
    has2 = bool(re.search(r"\b2\b|EQUIP", s))
    has3 = bool(re.search(r"\b3\b|MOBIL", s))
    has4 = bool(re.search(r"\b4\b|INTEGR", s))
    if has4 or (has1 and has2 and has3):
        return "4. Integral (1, 2 y 3)"
    if (has1 and has2) or (has1 and has3) or (has2 and has3):
        return "4. Integral (1, 2 y 3)"
    if has1:
        return "1. Infraestructura."
    if has2:
        return "2. Equipamiento."
    if has3:
        return "3. Mobiliario."
    return ERROR_FLAG


def norm_sino(v):
    if pd.isna(v) or str(v).strip() == "":
        return v
    s = re.sub(r"\s+", " ", str(v).strip().upper())
    s = s.replace("Í", "I")
    if re.match(r"^SI\b", s) or s in {"1", "S", "SI"}:
        return "SI"
    if re.match(r"^NO\b", s) or s in {"0", "N"}:
        return "NO"
    if s in {"_", "-", "-----------------------", "P"}:
        return ERROR_FLAG
    return ERROR_FLAG


def norm_fecha(v):
    if pd.isna(v) or str(v).strip() == "":
        return v
    s = str(v).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        try:
            dt = pd.to_datetime(s)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            return s
    if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", s):
        return s
    if re.match(r"^\d{1,2}-\d{1,2}-\d{4}$", s):
        return s.replace("-", "/")
    return s


def limpiar_comentarios(v):
    if pd.isna(v):
        return v
    s = str(v).strip()
    s = re.sub(r"\s+", " ", s)
    return s if s else None


# ── 6. APLICAR NORMALIZACIÓN ─────────────────────────────────────────────────
print("\n>> Limpiando campos...")

SINO_VARS = ["demol", "nueva", "reforz", "cerco", "sust", "ampl", "mobil", "agua", "elec", "unid"]

df["cod_local"] = df["cod_local"].apply(norm_cod_local)
df["cui"] = df["cui"].apply(norm_cui)
df["cod_mod"] = df["cod_mod"].apply(norm_cod_mod)
df["tipo"] = df["tipo"].apply(norm_tipo)
df["monto"] = df["monto"].apply(norm_monto)
df["avance"] = df["avance"].apply(norm_avance)
df["f9"] = df["f9"].apply(norm_f9)
df["comp"] = df["comp"].apply(norm_comp)
df["fecha"] = df["fecha"].apply(norm_fecha)
df["comentarios"] = df["comentarios"].apply(limpiar_comentarios)

for v in SINO_VARS:
    df[v] = df[v].apply(norm_sino)

print("   Limpieza de campos completada")

# ── 7. MARCAR ERRORES POR CAMPO ─────────────────────────────────────────────
CHECK_VARS = ["cui", "tipo", "monto", "avance", "f9", "comp"] + SINO_VARS

for v in CHECK_VARS:
    df[f"err_{v}"] = df[v].apply(lambda x: 1 if str(x) == ERROR_FLAG else 0)

df["tiene_error"] = df[[f"err_{v}" for v in CHECK_VARS]].max(axis=1)

# ── 8. CRUZAR CON BASE DE INVERSIONES ───────────────────────────────────────
if INPUT_INV.exists():
    print(f"\n>> Cruzando con Base de Inversiones: {INPUT_INV.name}")
    try:
        inv = pd.read_excel(INPUT_INV, dtype=str)
        print(f"   {len(inv):,} registros en Base de Inversiones")

        inv["cui_limpio"] = inv["CODIGO_UNICO"].apply(
            lambda x: re.sub(r"\.0+$", "", str(x).strip()) if pd.notna(x) else ""
        )

        cui_validos = set(inv["cui_limpio"].unique())

        df["cui_en_banco"] = df["cui"].apply(
            lambda x: "SI" if str(x) in cui_validos else "NO"
        )
        n_encontrados = (df["cui_en_banco"] == "SI").sum()
        n_no_encontrados = (df["cui_en_banco"] == "NO").sum()
        print(f"   CUI encontrados en Banco: {n_encontrados}")
        print(f"   CUI NO encontrados en Banco: {n_no_encontrados}")

    except Exception as e:
        print(f"   Error al cargar Base de Inversiones: {e}")
        df["cui_en_banco"] = "NO VERIFICADO"
else:
    print("\n>> Base de Inversiones no encontrada, se omite cruce")
    df["cui_en_banco"] = "NO VERIFICADO"

# ── 9. SEPARAR BASES ────────────────────────────────────────────────────────
err_cols = [f"err_{v}" for v in CHECK_VARS] + ["tiene_error"]
df_clean = df[df["tiene_error"] == 0].drop(columns=err_cols)
df_errors = df[df["tiene_error"] == 1].copy()

for v in CHECK_VARS:
    df_errors[v] = df_errors[v].replace(ERROR_FLAG, "<<ERROR>>")

print(f"\n>> Registros limpios:    {len(df_clean):,}")
print(f"   Registros con error:  {len(df_errors):,}")
print(f"   Pendientes registro:  {len(df_pendientes):,}")

# ── 10. RESUMEN DE ERRORES ──────────────────────────────────────────────────
print("\n" + "=" * 60)
print("RESUMEN DE ERRORES POR CAMPO")
print("=" * 60)
resumen = []
for v in CHECK_VARS:
    n = df[f"err_{v}"].sum()
    if n > 0:
        print(f"  {v:12s}: {n:,} registros con error")
        resumen.append({"campo": v, "errores": int(n)})

# ── 11. REPORTE DE COMPLETITUD POR REGIÓN ────────────────────────────────────
print("\n" + "=" * 60)
print("COMPLETITUD POR REGIÓN")
print("=" * 60)

df_all = pd.concat([df, df_pendientes], ignore_index=True)
completitud = []
for region in sorted(df_all["region"].dropna().unique()):
    total = len(df_all[df_all["region"] == region])
    registrados = len(df[df["region"] == region])
    limpios = len(df_clean[df_clean["region"] == region])
    con_error = len(df_errors[df_errors["region"] == region])
    pct = round(registrados / total * 100, 1) if total > 0 else 0

    completitud.append({
        "region": region,
        "total_locales": total,
        "registrados_cui": registrados,
        "limpios": limpios,
        "con_error": con_error,
        "pct_completitud": pct,
    })
    print(f"  {region:25s}: {registrados:5,} / {total:5,} ({pct:5.1f}%)")

df_completitud = pd.DataFrame(completitud)

# ── 12. EXPORTAR ─────────────────────────────────────────────────────────────
print("\n>> Exportando archivos...")

with pd.ExcelWriter(OUT_CLEAN, engine="openpyxl") as w:
    df_clean.to_excel(w, index=False, sheet_name="Base Limpia")
print(f"   {OUT_CLEAN.name}")

with pd.ExcelWriter(OUT_ERRORS, engine="openpyxl") as w:
    df_errors.to_excel(w, index=False, sheet_name="Registros con Errores")
    if resumen:
        pd.DataFrame(resumen).to_excel(w, index=False, sheet_name="Resumen Errores")
print(f"   {OUT_ERRORS.name}")

with pd.ExcelWriter(OUT_PENDIENTES, engine="openpyxl") as w:
    df_pendientes.to_excel(w, index=False, sheet_name="Pendientes de Registro")
print(f"   {OUT_PENDIENTES.name}")

with pd.ExcelWriter(OUT_REPORTE, engine="openpyxl") as w:
    df_completitud.to_excel(w, index=False, sheet_name="Completitud por Región")
print(f"   {OUT_REPORTE.name}")

print("\n" + "=" * 60)
print("LIMPIEZA COMPLETADA")
print("=" * 60)
