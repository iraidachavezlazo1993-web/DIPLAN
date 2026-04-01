"""
LIMPIEZA - Anexo 01: Relación de locales educativos intervenidos
================================================================
Limpia los campos ingresados por GR/GL en el Anexo 1 y valida contra
la Base de Inversiones MINEDU.

Campos limpiados:
  - CUI (solo numérico, validado contra Base MINEDU)
  - Tipo de inversión (PI/IOARR/IRI, cruzado con MINEDU)
  - Monto de inversión (numérico positivo)
  - Avance físico (0-100%)
  - Fecha de recepción de obra
  - Tiene F9 (SI/NO, cruzado con MINEDU)
  - Componentes (1. Infraestructura / 2. Equipamiento / 3. Mobiliario / 4. Integral)
  - Campos SI/NO de intervención (unid, demol, nueva, reforz, cerco, sust, ampl, mobil, agua, elec)
  - Comentarios

Genera archivos:
  - Anexo1_base_limpia.xlsx     (registros sin errores)
  - Anexo1_base_errores.xlsx    (registros con al menos un error)
  - Anexo1_validacion_cruce.xlsx (reporte de cruce con Base MINEDU)
"""

import pandas as pd
import re
from pathlib import Path

# ── 0. RUTAS (misma estructura que el do-file de Stata) ──────────────────────
reporte      = Path(r"C:\Users\iraid\Documents\DISCO TERA\MINEDU\TRABAJO-MINEDU\01_MINEDU")
rep_cons_i   = reporte / "01_input"
rep_cons_o   = reporte / "03_output"
rep_cons_t   = reporte / "04_temporal"

INPUT        = rep_cons_i / "Anexo_1_avance_GR_GL.xlsx"
INPUT_MINEDU = rep_cons_i / "2026.03.23 Base de Inversiones_.xlsx"
OUT_CLEAN    = rep_cons_o / "Anexo1_base_limpia.xlsx"
OUT_ERRORS   = rep_cons_o / "Anexo1_base_errores.xlsx"
OUT_CRUCE    = rep_cons_o / "Anexo1_validacion_cruce.xlsx"

# ── 1. IMPORTAR ──────────────────────────────────────────────────────────────
# header=0 → fila 1 del Excel como nombres de columna (igual que Stata firstrow)
# dtype=str → todo como texto para preservar los valores tal cual
df = pd.read_excel(INPUT, sheet_name="Hoja1", header=0, dtype=str)
print(f"Importado: {df.shape[0]:,} obs, {df.shape[1]} vars")

# ── 2. DROP filas de encabezado (fila vacía + fila con headers reales) ───────
df = df.drop(index=[0, 1]).reset_index(drop=True)
print(f"Tras drop encabezados: {df.shape[0]:,} obs")

# ── 3. RENAME a nombres cortos operativos ────────────────────────────────────
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
# Eliminar columnas extra (cols 25+)
extra_cols = [c for c in df.columns if c not in rename_map.values()]
df = df.drop(columns=extra_cols)

# ── 4. ELIMINAR FILAS SIN CUI (pendientes de registro) ──────────────────────
df = df[df["cui"].notna() & (df["cui"].str.strip() != "")].copy()
print(f"Filas con CUI: {len(df):,}")

# Columnas de error por campo (0=ok, 1=error)
CHECK_VARS = ["cui", "tipo", "monto", "avance", "f9", "comp",
              "unid", "demol", "nueva", "reforz", "cerco", "sust",
              "ampl", "mobil", "agua", "elec"]
for field in CHECK_VARS:
    df[f"err_{field}"] = 0

ERROR_FLAG = "__ERROR__"


# ── 5. FUNCIONES DE NORMALIZACIÓN ────────────────────────────────────────────

def norm_cui(v):
    """CUI debe ser solo dígitos. Intenta rescatar extrayendo secuencia de 5+ dígitos."""
    if pd.isna(v):
        return v
    s = str(v).strip()
    if re.match(r"^\d+$", s):
        return s
    # Intentar rescatar secuencia de 5+ dígitos
    m = re.search(r"\d{5,}", s)
    return m.group() if m else ERROR_FLAG


def norm_tipo(v):
    """Normaliza a PI / IOARR / IRI."""
    if pd.isna(v):
        return v
    s = re.sub(r"\s+", " ", str(v).strip().upper())
    # Detectar valores que claramente no son tipo (números, montos)
    if re.match(r"^\d+\.?\d*$", s):
        return ERROR_FLAG
    if s in {"S/N", ""}:
        return ERROR_FLAG
    if re.search(r"\bIOARR?\b|\bFUR\b", s):
        return "IOARR"
    if re.search(r"\bIRI\b", s):
        return "IRI"
    if re.search(r"\bPI\b|PROYECTO", s):
        return "PI"
    return ERROR_FLAG


def norm_monto(v):
    """Número positivo. Acepta coma como separador decimal."""
    if pd.isna(v):
        return v
    s = str(v).strip().replace(",", ".")
    # Eliminar espacios internos (separadores de miles)
    s = s.replace(" ", "")
    try:
        f = float(s)
        return str(f) if f >= 0 else ERROR_FLAG
    except ValueError:
        return ERROR_FLAG


def norm_avance(v):
    """Porcentaje 0–100. Acepta '96.10%', proporción 0.976 → 97.6."""
    if pd.isna(v):
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
    """SI o NO. Detecta valores que pertenecen a comp (corrimiento de columnas)."""
    if pd.isna(v):
        return v
    s = str(v).strip().upper().replace("Í", "I")
    # Detectar valores que claramente son de comp (corrimiento de columnas)
    if re.search(r"INFRAESTRUCTURA|EQUIPAMIENTO|MOBILIARIO|INTEGRAL|\d\.\s", s):
        return ERROR_FLAG
    if s in {"SI", "S", "1", "SÍ", "SI SECCIÓN B", "SI SECCION B"}:
        return "SI"
    if re.match(r"^SI\b", s):
        return "SI"
    if s in {"NO", "N", "0"}:
        return "NO"
    if re.match(r"^NO\b", s):
        return "NO"
    return ERROR_FLAG


# Mapeo de componentes
COMP_MAP = {
    # solo infraestructura
    "1": "1. Infraestructura.",
    "1.": "1. Infraestructura.",
    "1. INFRAESTRUCTURA.": "1. Infraestructura.",
    "1. INFRAESTRUCTURA": "1. Infraestructura.",
    "1.INFRAESTRUCTURA": "1. Infraestructura.",
    "INFRAESTRUCTURA": "1. Infraestructura.",
    # solo equipamiento
    "2": "2. Equipamiento.",
    "2.": "2. Equipamiento.",
    "2. EQUIPAMIENTO.": "2. Equipamiento.",
    "2.EQUIPAMIENTO": "2. Equipamiento.",
    "2. EQUIPAMIENT": "2. Equipamiento.",
    "2. EQUIPAMIENTO": "2. Equipamiento.",
    # solo mobiliario
    "3": "3. Mobiliario.",
    "3.": "3. Mobiliario.",
    "3. MOBILIARIO.": "3. Mobiliario.",
    "3. MOBILIARIO": "3. Mobiliario.",
    # integral
    "4": "4. Integral (1, 2 y 3)",
    "4. INTEGRAL (1, 2 Y 3)": "4. Integral (1, 2 y 3)",
    "INTEGRAL": "4. Integral (1, 2 y 3)",
    "(1,2 Y 3)": "4. Integral (1, 2 y 3)",
    "(1,2,3)": "4. Integral (1, 2 y 3)",
    "(1-2-3)": "4. Integral (1, 2 y 3)",
    "1,2,3": "4. Integral (1, 2 y 3)",
    "1, 2 Y 3": "4. Integral (1, 2 y 3)",
    "1 Y 2 Y 3": "4. Integral (1, 2 y 3)",
    # combos de dos → integral
    "1,2": "4. Integral (1, 2 y 3)",
    "1, 2": "4. Integral (1, 2 y 3)",
    "1 Y 2": "4. Integral (1, 2 y 3)",
    "1,3": "4. Integral (1, 2 y 3)",
    "1 Y 3": "4. Integral (1, 2 y 3)",
    "2,3": "4. Integral (1, 2 y 3)",
    "2 Y 3": "4. Integral (1, 2 y 3)",
    "1.2": "4. Integral (1, 2 y 3)",
    "1.3": "4. Integral (1, 2 y 3)",
    "2.3": "4. Integral (1, 2 y 3)",
    "2. EQUIPAMIENTO 3.MOBILIARIO": "4. Integral (1, 2 y 3)",
    "1,2 Y 3": "4. Integral (1, 2 y 3)",
    "1,2 Y 3": "4. Integral (1, 2 y 3)",
    "1 Y 2": "4. Integral (1, 2 y 3)",
}


def norm_comp(v):
    """Normaliza componentes a las 4 categorías estándar."""
    if pd.isna(v):
        return v
    s = re.sub(r"\s+", " ", str(v).strip().upper())
    # Valores claramente inválidos
    if s in {"-", "_", "-1", "-----------------------", ""}:
        return ERROR_FLAG
    if s in COMP_MAP:
        return COMP_MAP[s]
    # Detección por contenido
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
    """Normaliza campos SI/NO. Acepta variantes como 'SI (PARCIAL)', 'SI-TOTAL', etc."""
    if pd.isna(v):
        return v
    s = re.sub(r"\s+", " ", str(v).strip().upper().replace("Í", "I"))
    # Valores claramente inválidos (guiones, underscores, textos largos)
    if s in {"-", "_", "P", "MO", ""}:
        return ERROR_FLAG
    # Detectar textos largos que son comentarios (>30 chars y no empiezan con SI/NO)
    if len(s) > 30 and not re.match(r"^(SI|NO)\b", s):
        return ERROR_FLAG
    # Números que no son 0/1 → error
    if re.match(r"^\d+$", s) and s not in {"0", "1"}:
        return ERROR_FLAG
    # Normalizar SI
    if re.match(r"^SI\b", s) or s in {"1", "S", "SI"}:
        return "SI"
    # Normalizar NO
    if re.match(r"^NO\b", s) or s in {"0", "N"}:
        return "NO"
    return ERROR_FLAG


# ── 6. APLICAR NORMALIZACIÓN ────────────────────────────────────────────────
SINO_VARS = ["demol", "nueva", "reforz", "cerco", "sust", "ampl", "mobil", "agua", "elec", "unid"]

# Antes de normalizar elec, mover textos largos a comentarios si están vacíos
for idx in df.index:
    elec_val = df.at[idx, "elec"]
    if pd.notna(elec_val):
        s = str(elec_val).strip()
        # Si es un texto largo (>30 chars) que no es SI/NO → mover a comentarios
        if len(s) > 30 and not re.match(r"^(SI|NO)\b", s.upper()):
            comment = df.at[idx, "comentarios"]
            if pd.isna(comment) or str(comment).strip() == "":
                df.at[idx, "comentarios"] = s
            else:
                df.at[idx, "comentarios"] = str(comment).strip() + " | " + s
            df.at[idx, "elec"] = pd.NA

# Detectar corrimiento de columnas en f9 (cuando f9 tiene valor de comp)
for idx in df.index:
    f9_val = df.at[idx, "f9"]
    if pd.notna(f9_val):
        s = str(f9_val).strip().upper()
        if re.search(r"INFRAESTRUCTURA|EQUIPAMIENTO|MOBILIARIO|INTEGRAL|\d\.\s", s):
            # Este valor pertenece a comp, no a f9 → marcar f9 como vacío
            df.at[idx, "f9"] = pd.NA

print("Aplicando normalización...")
df["cui"]    = df["cui"].apply(norm_cui)
df["tipo"]   = df["tipo"].apply(norm_tipo)
df["monto"]  = df["monto"].apply(norm_monto)
df["avance"] = df["avance"].apply(norm_avance)
df["f9"]     = df["f9"].apply(norm_f9)
df["comp"]   = df["comp"].apply(norm_comp)
for v in SINO_VARS:
    df[v] = df[v].apply(norm_sino)


# ── 7. MARCAR ERRORES POR CAMPO ─────────────────────────────────────────────
for v in CHECK_VARS:
    df[f"err_{v}"] = df[v].apply(lambda x: 1 if str(x) == ERROR_FLAG else 0)

df["tiene_error"] = df[[f"err_{v}" for v in CHECK_VARS]].max(axis=1)


# ── 8. SEPARAR BASES ────────────────────────────────────────────────────────
df_clean = df[df["tiene_error"] == 0].drop(
    columns=[f"err_{v}" for v in CHECK_VARS] + ["tiene_error"]
)
df_errors = df[df["tiene_error"] == 1].copy()

# Reemplazar flag por etiqueta legible en base de errores
for v in CHECK_VARS:
    df_errors[v] = df_errors[v].replace(ERROR_FLAG, "<<ERROR>>")

print(f"\nRegistros limpios : {len(df_clean):,}")
print(f"Registros con error: {len(df_errors):,}")

# Resumen de errores por campo
print("\n=== RESUMEN DE ERRORES POR CAMPO ===")
for v in CHECK_VARS:
    n = df[f"err_{v}"].sum()
    if n > 0:
        print(f"  err_{v}: {n} registros")


# ── 9. EXPORTAR ─────────────────────────────────────────────────────────────
with pd.ExcelWriter(OUT_CLEAN, engine="openpyxl") as w:
    df_clean.to_excel(w, index=False, sheet_name="Base Limpia")

with pd.ExcelWriter(OUT_ERRORS, engine="openpyxl") as w:
    df_errors.to_excel(w, index=False, sheet_name="Registros con Errores")


# ══════════════════════════════════════════════════════════════════════════════
# 10. VALIDACIÓN CRUZADA CON BASE DE INVERSIONES MINEDU
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("VALIDACIÓN CRUZADA CON BASE DE INVERSIONES MINEDU")
print("=" * 60)

if INPUT_MINEDU.exists():
    minedu = pd.read_excel(INPUT_MINEDU, sheet_name="Data", dtype=str)
    print(f"Base MINEDU cargada: {len(minedu):,} registros")

    # Normalizar tipo MINEDU para comparación
    def tipo_minedu(v):
        if pd.isna(v):
            return v
        s = str(v).strip().upper()
        if "IOARR" in s:
            return "IOARR"
        if "PROYECTO" in s:
            return "PI"
        if "IRI" in s:
            return "IRI"
        return s

    minedu["tipo_norm"] = minedu["DES_TIPO_FORMATO"].apply(tipo_minedu)
    minedu["cui_norm"] = minedu["CODIGO_UNICO"].str.strip()

    # Conjunto de CUIs válidos MINEDU
    minedu_cuis = set(minedu["cui_norm"].dropna())

    # Usar la base completa (limpia + errores) para el cruce
    df_all = df.copy()
    # Solo filas con CUI numérico válido (no ERROR_FLAG)
    df_cruce = df_all[df_all["cui"].apply(lambda x: str(x) != ERROR_FLAG and pd.notna(x))].copy()
    anexo_cuis = set(df_cruce["cui"].str.strip())

    match_cuis = minedu_cuis & anexo_cuis
    solo_minedu = minedu_cuis - anexo_cuis
    solo_anexo = anexo_cuis - minedu_cuis

    print(f"\nCUIs en Base MINEDU: {len(minedu_cuis)}")
    print(f"CUIs en Anexo1 (con CUI válido): {len(anexo_cuis)}")
    print(f"CUIs que coinciden: {len(match_cuis)}")
    print(f"CUIs solo en MINEDU (no declarados en Anexo1): {len(solo_minedu)}")
    print(f"CUIs solo en Anexo1 (no en Base MINEDU): {len(solo_anexo)}")

    # ── Construir reporte de cruce ──
    rows_cruce = []

    # 1) CUIs que coinciden: comparar tipo, monto, f9, avance
    for cui in sorted(match_cuis):
        m = minedu[minedu["cui_norm"] == cui].iloc[0]
        a_rows = df_cruce[df_cruce["cui"].str.strip() == cui]
        for _, a in a_rows.iterrows():
            row = {"cui": cui, "status_cruce": "COINCIDE"}
            # Tipo
            row["tipo_anexo1"] = a.get("tipo", "")
            row["tipo_minedu"] = m.get("tipo_norm", "")
            row["tipo_ok"] = "OK" if row["tipo_anexo1"] == row["tipo_minedu"] else "DIFERENTE"
            # Monto
            row["monto_anexo1"] = a.get("monto", "")
            row["monto_minedu"] = m.get("COSTO_ACTUALIZADO_BI", "")
            # F9
            row["f9_anexo1"] = a.get("f9", "")
            row["f9_minedu"] = m.get("TIENE_F9", "")
            row["f9_ok"] = "OK" if str(row["f9_anexo1"]).upper() == str(row["f9_minedu"]).upper() else "DIFERENTE"
            # Avance
            row["avance_anexo1"] = a.get("avance", "")
            row["avance_minedu_f9"] = m.get("AVANCE_FISICO_F9", "")
            row["avance_minedu_f12b"] = m.get("AVANCE_FISICO_F12B", "")
            # Estado MINEDU
            row["estado_minedu"] = m.get("ESTADO", "")
            row["situacion_minedu"] = m.get("SITUACION", "")
            row["nombre_ie"] = a.get("nombre_ie", "")
            row["nombre_inv_minedu"] = m.get("NOMBRE_INVERSION", "")
            rows_cruce.append(row)

    # 2) CUIs solo en MINEDU
    for cui in sorted(solo_minedu):
        m = minedu[minedu["cui_norm"] == cui].iloc[0]
        rows_cruce.append({
            "cui": cui,
            "status_cruce": "SOLO EN MINEDU (no declarado por GR/GL)",
            "tipo_minedu": m.get("tipo_norm", ""),
            "monto_minedu": m.get("COSTO_ACTUALIZADO_BI", ""),
            "f9_minedu": m.get("TIENE_F9", ""),
            "estado_minedu": m.get("ESTADO", ""),
            "situacion_minedu": m.get("SITUACION", ""),
            "nombre_inv_minedu": m.get("NOMBRE_INVERSION", ""),
        })

    # 3) CUIs solo en Anexo1 (muestra: no están validados por MINEDU)
    for cui in sorted(solo_anexo):
        a_rows = df_cruce[df_cruce["cui"].str.strip() == cui]
        for _, a in a_rows.iterrows():
            rows_cruce.append({
                "cui": cui,
                "status_cruce": "SOLO EN ANEXO1 (no validado por MINEDU)",
                "tipo_anexo1": a.get("tipo", ""),
                "monto_anexo1": a.get("monto", ""),
                "f9_anexo1": a.get("f9", ""),
                "avance_anexo1": a.get("avance", ""),
                "nombre_ie": a.get("nombre_ie", ""),
            })

    df_cruce_out = pd.DataFrame(rows_cruce)

    # Ordenar columnas
    col_order = ["cui", "status_cruce", "nombre_ie", "nombre_inv_minedu",
                 "tipo_anexo1", "tipo_minedu", "tipo_ok",
                 "monto_anexo1", "monto_minedu",
                 "f9_anexo1", "f9_minedu", "f9_ok",
                 "avance_anexo1", "avance_minedu_f9", "avance_minedu_f12b",
                 "estado_minedu", "situacion_minedu"]
    col_order = [c for c in col_order if c in df_cruce_out.columns]
    df_cruce_out = df_cruce_out[col_order]

    # Exportar
    with pd.ExcelWriter(OUT_CRUCE, engine="openpyxl") as w:
        df_cruce_out.to_excel(w, index=False, sheet_name="Validacion Cruce")

    print(f"\nReporte de cruce generado: {OUT_CRUCE}")

    # Resumen de discrepancias
    if len(match_cuis) > 0:
        coinciden = df_cruce_out[df_cruce_out["status_cruce"] == "COINCIDE"]
        if "tipo_ok" in coinciden.columns:
            tipo_diff = (coinciden["tipo_ok"] == "DIFERENTE").sum()
            print(f"  Tipo discrepante: {tipo_diff}")
        if "f9_ok" in coinciden.columns:
            f9_diff = (coinciden["f9_ok"] == "DIFERENTE").sum()
            print(f"  F9 discrepante: {f9_diff}")
else:
    print(f"AVISO: No se encontró {INPUT_MINEDU}")
    print("  Se omite la validación cruzada.")


# ── RESUMEN FINAL ────────────────────────────────────────────────────────────
print(f"\nArchivos generados:")
print(f"  {OUT_CLEAN}")
print(f"  {OUT_ERRORS}")
if INPUT_MINEDU.exists():
    print(f"  {OUT_CRUCE}")
