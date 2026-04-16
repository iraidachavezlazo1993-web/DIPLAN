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
INPUT_INV = entrada / "2026.04.13 Base de Inversiones.xlsx"
INPUT_VINC = entrada / "2026.04.14 Vinculaciones.xlsx"
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

def _parse_codigos(v, n_digitos, letras_missing=True):
    """Lógica común: separa múltiples códigos por /, pad cada uno a n_digitos.
    Letras = missing si letras_missing=True."""
    if pd.isna(v) or str(v).strip() == "":
        return v
    s = str(v).strip()
    if s in {"-", "_", "--", "---", "-----------------------"}:
        return None
    s_upper = s.upper()
    if s_upper in {"SI", "SÍ", "NO", "NINGUNA", "0"}:
        return None
    if letras_missing and re.search(r"[a-zA-Z]", s):
        return None
    s = re.sub(r"\.0+$", "", s)
    s = s.replace("/", ",").replace("-", ",")
    s = s.replace("\n", ",").replace("\r", ",")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r",+", ",", s).strip(",")
    if not re.match(r"^[\d,]+$", s):
        return None
    codigos = []
    for cod in s.split(","):
        cod = cod.strip()
        if cod:
            cod = re.sub(r"\.0+$", "", cod)
            cod = cod.zfill(n_digitos)
            codigos.append(cod)
    return "/".join(codigos) if codigos else None


def norm_cod_local(v):
    """Código local: 6 dígitos cada uno, separados por /. Letras = missing."""
    return _parse_codigos(v, 6)


def norm_cui(v):
    """CUI: 7 dígitos cada uno, separados por /. Intenta rescatar si tiene letras."""
    if pd.isna(v):
        return v
    s = str(v).strip()
    s = re.sub(r"\.0+$", "", s)
    # si ya es numérico puro (o con separadores), usar lógica estándar
    clean = s.replace("/", ",").replace("-", ",").replace(" ", "")
    clean = re.sub(r",+", ",", clean).strip(",")
    if re.match(r"^[\d,]+$", clean):
        codigos = [c.zfill(7) for c in clean.split(",") if c]
        return "/".join(codigos) if codigos else ERROR_FLAG
    # intentar rescatar secuencia de 5+ dígitos
    m = re.search(r"\d{5,}", s)
    return m.group().zfill(7) if m else ERROR_FLAG


def norm_cod_mod(v):
    """Código modular: 7 dígitos cada uno, separados por /. Letras = missing."""
    return _parse_codigos(v, 7)


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
# Mapeo: campo Anexo1 → campo Banco de Inversiones
MAPEO_BANCO = {
    "tipo":   "DES_TIPO_FORMATO",
    "monto":  "COSTO_INV_TOTAL_BI",
    "f9":     "TIENE_F9",
}
# Campos adicionales del Banco para enriquecer
CAMPOS_EXTRA_BANCO = {
    "banco_estado":     "ESTADO",
    "banco_situacion":  "SITUACION",
    "banco_avance_f9":  "AVANCE_FISICO_F9",
    "banco_avance_f12b":"AVANCE_FISICO_F12B",
    "banco_monto_actualizado": "COSTO_ACTUALIZADO_BI",
    "banco_fecha_registro":    "FECHA_REGISTRO",
    "banco_fecha_viabilidad":  "FECHA_VIABILIDAD",
    "banco_tiene_f8":   "TIENE_F8",
    "banco_nombre_inv": "NOMBRE_INVERSION",
    "banco_departamento": "DEPARTAMENTO_CUI",
    "banco_provincia":    "PROVINCIA_CUI",
    "banco_distrito":     "DISTRITO",
}

def tipo_banco_a_anexo(v):
    if pd.isna(v) or str(v).strip() == "":
        return None
    s = str(v).strip().upper()
    if "IOARR" in s or "FUR" in s:
        return "IOARR"
    if "IRI" in s:
        return "IRI"
    if "PROYECTO" in s or "PI" == s:
        return "PI"
    return None

# ── 8.1 Cargar Base de Inversiones ──────────────────────────────────────────
inv_dict = None
if INPUT_INV.exists():
    print(f"\n>> Cargando Base de Inversiones: {INPUT_INV.name}")
    try:
        # Intentar leer con header=4 (formato nuevo) o header=0 (antiguo)
        try:
            inv = pd.read_excel(INPUT_INV, sheet_name="Data", header=4, dtype=str)
        except Exception:
            inv = pd.read_excel(INPUT_INV, dtype=str)

        if "CODIGO_UNICO" not in inv.columns:
            raise ValueError("No se encontró la columna CODIGO_UNICO")

        inv["cui_limpio"] = inv["CODIGO_UNICO"].apply(
            lambda x: re.sub(r"\.0+$", "", str(x).strip()).zfill(7) if pd.notna(x) else ""
        )
        inv_dedup = inv.drop_duplicates(subset=["cui_limpio"], keep="first")
        inv_dict = inv_dedup.set_index("cui_limpio")
        print(f"   {len(inv_dict):,} CUIs únicos en Banco")
    except Exception as e:
        print(f"   Error: {e}")
        inv_dict = None
else:
    print("\n>> Base de Inversiones no encontrada")

# ── 8.2 Cargar Vinculaciones (cod_local y cod_mod por CUI) ──────────────────
vinc_por_cui = None
vinc_por_local = None
vinc_por_mod = None
if INPUT_VINC.exists():
    print(f"\n>> Cargando Vinculaciones: {INPUT_VINC.name}")
    try:
        vinc = pd.read_excel(INPUT_VINC, sheet_name="Vinculaciones", dtype=str)
        print(f"   {len(vinc):,} registros en Vinculaciones")

        # Solo registros Vinculados (confirmados por MEF)
        vinc = vinc[vinc["ESTADO_VINCULACION"].str.strip().str.lower().str.startswith("vinculado", na=False)]
        print(f"   {len(vinc):,} registros con estado Vinculado")

        # Normalizar códigos: CUI a 7, cod_local a 6, cod_mod a 7
        vinc["cui_7"] = vinc["CUI"].apply(
            lambda x: re.sub(r"\.0+$", "", str(x).strip()).zfill(7) if pd.notna(x) and str(x).strip() else ""
        )
        vinc["cod_local_6"] = vinc["CODIGO_LOCAL"].apply(
            lambda x: re.sub(r"\.0+$", "", str(x).strip()).zfill(6) if pd.notna(x) and str(x).strip() else ""
        )
        vinc["cod_mod_7"] = vinc["CODIGO_MODULAR"].apply(
            lambda x: re.sub(r"\.0+$", "", str(x).strip()).zfill(7) if pd.notna(x) and str(x).strip() else ""
        )

        # Agrupar por CUI: obtener listas de cod_local y cod_mod únicos
        vinc_por_cui = vinc.groupby("cui_7").agg({
            "cod_local_6": lambda x: sorted(set(v for v in x if v)),
            "cod_mod_7":   lambda x: sorted(set(v for v in x if v)),
            "GRUPO":       lambda x: list(set(x.dropna())),
        }).to_dict("index")

        # Lookup inverso: cod_local → CUI(s), cod_mod → CUI(s)
        vinc_por_local = vinc.groupby("cod_local_6")["cui_7"].apply(
            lambda x: sorted(set(v for v in x if v))
        ).to_dict()
        vinc_por_mod = vinc.groupby("cod_mod_7")["cui_7"].apply(
            lambda x: sorted(set(v for v in x if v))
        ).to_dict()

        print(f"   CUIs únicos Vinculados: {len(vinc_por_cui):,}")

    except Exception as e:
        print(f"   Error: {e}")
        import traceback
        traceback.print_exc()
else:
    print("\n>> Vinculaciones no encontrado")

# ── 8.3 CRUZAR CON BANCO DE INVERSIONES (por CUI) ────────────────────────────
if inv_dict is not None:
    print(f"\n>> Cruzando Anexo 1 con Banco de Inversiones...")
    cui_validos = set(inv_dict.index)
    df["cui_en_banco"] = df["cui"].apply(
        lambda x: "SI" if str(x).split("/")[0] in cui_validos else "NO"
    )
    n_encontrados = (df["cui_en_banco"] == "SI").sum()
    print(f"   CUI encontrados en Banco: {n_encontrados}")
    print(f"   CUI NO encontrados en Banco: {len(df) - n_encontrados}")

    n_reemplazados = 0
    n_completados = 0
    for idx, row in df.iterrows():
        cui_val = str(row["cui"]).split("/")[0]
        if cui_val not in cui_validos:
            continue
        banco_row = inv_dict.loc[cui_val]

        # tipo: prevalece Banco
        banco_tipo = tipo_banco_a_anexo(banco_row.get("DES_TIPO_FORMATO"))
        if banco_tipo:
            if pd.isna(row["tipo"]) or row["tipo"] == "" or row["tipo"] == ERROR_FLAG:
                df.at[idx, "tipo"] = banco_tipo
                df.at[idx, "err_tipo"] = 0
                n_completados += 1
            elif row["tipo"] != banco_tipo:
                df.at[idx, "tipo"] = banco_tipo
                df.at[idx, "err_tipo"] = 0
                n_reemplazados += 1

        # monto: prevalece Banco
        banco_monto = banco_row.get("COSTO_INV_TOTAL_BI")
        if pd.notna(banco_monto) and str(banco_monto).strip() != "":
            try:
                monto_banco = str(round(float(str(banco_monto).strip()), 2))
                if pd.isna(row["monto"]) or row["monto"] == "" or row["monto"] == ERROR_FLAG:
                    df.at[idx, "monto"] = monto_banco
                    df.at[idx, "err_monto"] = 0
                    n_completados += 1
                else:
                    df.at[idx, "monto"] = monto_banco
                    df.at[idx, "err_monto"] = 0
                    n_reemplazados += 1
            except ValueError:
                pass

        # f9: prevalece Banco
        banco_f9 = banco_row.get("TIENE_F9")
        if pd.notna(banco_f9) and str(banco_f9).strip().upper() in {"SI", "NO"}:
            banco_f9_clean = str(banco_f9).strip().upper()
            if pd.isna(row["f9"]) or row["f9"] == "" or row["f9"] == ERROR_FLAG:
                df.at[idx, "f9"] = banco_f9_clean
                df.at[idx, "err_f9"] = 0
                n_completados += 1
            elif row["f9"] != banco_f9_clean:
                df.at[idx, "f9"] = banco_f9_clean
                df.at[idx, "err_f9"] = 0
                n_reemplazados += 1

        # avance: completar si falta
        if pd.isna(row["avance"]) or row["avance"] == "" or row["avance"] == ERROR_FLAG:
            banco_av = banco_row.get("AVANCE_FISICO_F9")
            if pd.isna(banco_av) or str(banco_av).strip() in {"", "0"}:
                banco_av = banco_row.get("AVANCE_FISICO_F12B")
            if pd.notna(banco_av) and str(banco_av).strip() != "":
                try:
                    av = float(str(banco_av).strip())
                    if 0 <= av <= 1:
                        av = round(av * 100, 2)
                    df.at[idx, "avance"] = str(round(av, 2))
                    df.at[idx, "err_avance"] = 0
                    n_completados += 1
                except ValueError:
                    pass

        # campos extra del Banco (enriquecer)
        for col_nueva, col_banco in CAMPOS_EXTRA_BANCO.items():
            val = banco_row.get(col_banco)
            if pd.notna(val) and str(val).strip() != "":
                df.at[idx, col_nueva] = str(val).strip()

    # recalcular errores
    for v in CHECK_VARS:
        df[f"err_{v}"] = df[v].apply(lambda x: 1 if str(x) == ERROR_FLAG else 0)
    df["tiene_error"] = df[[f"err_{v}" for v in CHECK_VARS]].max(axis=1)

    print(f"   Campos reemplazados (prevalece Banco): {n_reemplazados}")
    print(f"   Campos completados (faltaban en Anexo): {n_completados}")
else:
    df["cui_en_banco"] = "NO VERIFICADO"

# ── 8.4 CRUZAR CON VINCULACIONES (cod_local y cod_mod oficiales) ────────────
if vinc_por_cui is not None:
    print(f"\n>> Cruzando Anexo 1 con Vinculaciones...")

    df["cui_en_vinc"] = ""
    df["cod_local_oficial"] = ""
    df["cod_mod_oficial"] = ""
    df["cod_local_match"] = ""
    df["cod_mod_match"] = ""
    df["vinc_grupo"] = ""

    n_cui_vinc = 0
    n_completados_local = 0
    n_completados_mod = 0
    n_mismatch_local = 0
    n_mismatch_mod = 0

    for idx, row in df.iterrows():
        cui_val = str(row["cui"]).split("/")[0]
        if cui_val not in vinc_por_cui:
            df.at[idx, "cui_en_vinc"] = "NO"
            continue
        df.at[idx, "cui_en_vinc"] = "SI"
        n_cui_vinc += 1
        v = vinc_por_cui[cui_val]

        # Listas oficiales desde Vinculaciones
        locales_oficiales = v.get("cod_local_6", [])
        mods_oficiales = v.get("cod_mod_7", [])
        grupos = v.get("GRUPO", [])

        df.at[idx, "cod_local_oficial"] = "/".join(locales_oficiales)
        df.at[idx, "cod_mod_oficial"] = "/".join(mods_oficiales)
        df.at[idx, "vinc_grupo"] = "/".join(grupos) if grupos else ""

        # --- cod_local: validar y completar ---
        local_anexo = str(row["cod_local"]) if pd.notna(row["cod_local"]) else ""
        if local_anexo and local_anexo != "nan":
            set_anexo = set(local_anexo.split("/"))
            set_oficial = set(locales_oficiales)
            if set_anexo.issubset(set_oficial) or set_oficial.issubset(set_anexo):
                df.at[idx, "cod_local_match"] = "SI"
            else:
                df.at[idx, "cod_local_match"] = "MISMATCH"
                n_mismatch_local += 1
        elif locales_oficiales:
            df.at[idx, "cod_local"] = "/".join(locales_oficiales)
            df.at[idx, "cod_local_match"] = "COMPLETADO"
            n_completados_local += 1

        # --- cod_mod: validar y completar ---
        mod_anexo = str(row["cod_mod"]) if pd.notna(row["cod_mod"]) else ""
        if mod_anexo and mod_anexo != "nan":
            set_anexo_m = set(mod_anexo.split("/"))
            set_oficial_m = set(mods_oficiales)
            if set_anexo_m.issubset(set_oficial_m) or set_oficial_m.issubset(set_anexo_m):
                df.at[idx, "cod_mod_match"] = "SI"
            else:
                df.at[idx, "cod_mod_match"] = "MISMATCH"
                n_mismatch_mod += 1
        elif mods_oficiales:
            df.at[idx, "cod_mod"] = "/".join(mods_oficiales)
            df.at[idx, "cod_mod_match"] = "COMPLETADO"
            n_completados_mod += 1

    print(f"   CUIs encontrados en Vinculaciones: {n_cui_vinc}")
    print(f"   cod_local completados desde Vinculaciones: {n_completados_local}")
    print(f"   cod_mod completados desde Vinculaciones: {n_completados_mod}")
    print(f"   cod_local con mismatch (no coinciden): {n_mismatch_local}")
    print(f"   cod_mod con mismatch (no coinciden): {n_mismatch_mod}")
else:
    df["cui_en_vinc"] = "NO VERIFICADO"
    df["cod_local_oficial"] = ""
    df["cod_mod_oficial"] = ""
    df["cod_local_match"] = ""
    df["cod_mod_match"] = ""
    df["vinc_grupo"] = ""

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

# ── 12. DICCIONARIO DE VARIABLES ────────────────────────────────────────────
DICCIONARIO = [
    {"variable": "nro", "descripcion": "Número correlativo", "tipo": "numérico", "regla_limpieza": "Sin cambios"},
    {"variable": "cod_local", "descripcion": "Código de local educativo", "tipo": "texto (6 dígitos)", "regla_limpieza": "Pad con ceros a la izquierda hasta 6 dígitos. Letras = missing"},
    {"variable": "region", "descripcion": "Región", "tipo": "texto", "regla_limpieza": "Normalizado a mayúsculas"},
    {"variable": "provincia", "descripcion": "Provincia", "tipo": "texto", "regla_limpieza": "Sin cambios"},
    {"variable": "distrito", "descripcion": "Distrito", "tipo": "texto", "regla_limpieza": "Sin cambios"},
    {"variable": "nombre_ie", "descripcion": "Nombre de la institución educativa", "tipo": "texto", "regla_limpieza": "Sin cambios"},
    {"variable": "cui", "descripcion": "Código Único de Inversión (CUI)", "tipo": "texto (7 dígitos)", "regla_limpieza": "Solo números. Pad con ceros hasta 7 dígitos. Se rescatan secuencias de 5+ dígitos"},
    {"variable": "tipo", "descripcion": "Tipo de inversión (PI/IOARR/IRI)", "tipo": "texto", "regla_limpieza": "Estandarizado a PI, IOARR o IRI. Prevalece info del Banco de Inversiones"},
    {"variable": "monto", "descripcion": "Monto de inversión (S/)", "tipo": "numérico", "regla_limpieza": "Convertido a numérico. Prevalece info del Banco de Inversiones"},
    {"variable": "avance", "descripcion": "Avance físico (%)", "tipo": "numérico (0-100)", "regla_limpieza": "Proporción 0-1 convertida a %. Completado con Banco si falta"},
    {"variable": "fecha", "descripcion": "Fecha de recepción de obra (o estimada)", "tipo": "fecha (dd/mm/yyyy)", "regla_limpieza": "Formato estandarizado a dd/mm/yyyy"},
    {"variable": "f9", "descripcion": "Tiene Formato 9 (SI/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO. Prevalece info del Banco de Inversiones"},
    {"variable": "comp", "descripcion": "Componentes de la inversión", "tipo": "texto", "regla_limpieza": "Mapeado a: 1.Infraestructura, 2.Equipamiento, 3.Mobiliario, 4.Integral"},
    {"variable": "unid", "descripcion": "¿Intervino todas las unidades productoras? (SÍ/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO"},
    {"variable": "cod_mod", "descripcion": "Códigos modulares intervenidos", "tipo": "texto (7 dígitos c/u)", "regla_limpieza": "Cada código con 7 dígitos, pad con ceros. Múltiples separados por /. Letras = missing"},
    {"variable": "demol", "descripcion": "¿Demolición total o parcial? (SÍ/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO"},
    {"variable": "nueva", "descripcion": "¿Construcción de nueva infraestructura? (SÍ/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO"},
    {"variable": "reforz", "descripcion": "¿Reforzamiento estructural? (SÍ/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO"},
    {"variable": "cerco", "descripcion": "¿Cerco perimétrico? (SÍ/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO"},
    {"variable": "sust", "descripcion": "¿Sustitución parcial de edificaciones? (SÍ/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO"},
    {"variable": "ampl", "descripcion": "¿Ampliación del área? (SÍ/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO"},
    {"variable": "mobil", "descripcion": "¿Reposición/dotación de mobiliario y equipamiento? (SÍ/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO"},
    {"variable": "agua", "descripcion": "¿Acceso a agua y desagüe? (SÍ/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO"},
    {"variable": "elec", "descripcion": "¿Acceso a energía eléctrica? (SÍ/NO)", "tipo": "texto", "regla_limpieza": "Estandarizado a SI o NO"},
    {"variable": "comentarios", "descripcion": "Comentarios adicionales", "tipo": "texto", "regla_limpieza": "Limpieza de espacios extra"},
    {"variable": "cui_en_banco", "descripcion": "¿El CUI existe en el Banco de Inversiones?", "tipo": "texto", "regla_limpieza": "SI / NO / NO VERIFICADO"},
    {"variable": "banco_estado", "descripcion": "Estado de la inversión (del Banco)", "tipo": "texto", "regla_limpieza": "Campo adicional del Banco de Inversiones"},
    {"variable": "banco_situacion", "descripcion": "Situación de la inversión (del Banco)", "tipo": "texto", "regla_limpieza": "Campo adicional del Banco de Inversiones"},
    {"variable": "banco_avance_f9", "descripcion": "Avance físico F9 (del Banco)", "tipo": "numérico", "regla_limpieza": "Campo adicional del Banco de Inversiones"},
    {"variable": "banco_avance_f12b", "descripcion": "Avance físico F12B (del Banco)", "tipo": "numérico", "regla_limpieza": "Campo adicional del Banco de Inversiones"},
    {"variable": "banco_monto_actualizado", "descripcion": "Costo actualizado (del Banco)", "tipo": "numérico", "regla_limpieza": "Campo adicional del Banco de Inversiones"},
    {"variable": "banco_nombre_inv", "descripcion": "Nombre de la inversión (del Banco)", "tipo": "texto", "regla_limpieza": "Campo adicional del Banco de Inversiones"},
    {"variable": "cui_en_vinc", "descripcion": "¿El CUI existe en Vinculaciones MEF?", "tipo": "texto", "regla_limpieza": "SI / NO"},
    {"variable": "cod_local_oficial", "descripcion": "Códigos locales oficiales desde Vinculaciones", "tipo": "texto (sep /)", "regla_limpieza": "Lista desde Vinculaciones MEF para ese CUI"},
    {"variable": "cod_mod_oficial", "descripcion": "Códigos modulares oficiales desde Vinculaciones", "tipo": "texto (sep /)", "regla_limpieza": "Lista desde Vinculaciones MEF para ese CUI"},
    {"variable": "cod_local_match", "descripcion": "¿cod_local del Anexo coincide con el oficial?", "tipo": "texto", "regla_limpieza": "SI / MISMATCH / COMPLETADO"},
    {"variable": "cod_mod_match", "descripcion": "¿cod_mod del Anexo coincide con el oficial?", "tipo": "texto", "regla_limpieza": "SI / MISMATCH / COMPLETADO"},
    {"variable": "vinc_grupo", "descripcion": "Grupo de la inversión (Básica, Superior, etc.)", "tipo": "texto", "regla_limpieza": "Desde Vinculaciones MEF"},
]
df_diccionario = pd.DataFrame(DICCIONARIO)

# ── 13. EXPORTAR ─────────────────────────────────────────────────────────────
print("\n>> Exportando archivos...")

with pd.ExcelWriter(OUT_CLEAN, engine="openpyxl") as w:
    df_clean.to_excel(w, index=False, sheet_name="Base Limpia")
    df_diccionario.to_excel(w, index=False, sheet_name="Diccionario de Variables")
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
