import re
import pandas as pd
import numpy as np
from dateutil import parser

# ---------------- Utilidades de normalización ----------------

_NULL_TOKENS = {"none", "nan", "null", "n/a", "na", "-"}

def _collapse_spaces(x):
    """Convierte \u00a0 a espacio, colapsa espacios múltiples."""
    if pd.isna(x):
        return x
    s = str(x).replace("\u00a0", " ")
    return " ".join(s.split())

def _normalize_nulls(x):
    """
    Convierte cadenas como 'None', 'NONE', 'nan', 'NaN', 'NULL', vacío, etc. a NaN.
    """
    if x is None:
        return np.nan
    s = str(x).strip()
    if s == "":
        return np.nan
    if s.lower() in _NULL_TOKENS:
        return np.nan
    return x

def apply_non_destructive_normalization(df: pd.DataFrame) -> pd.DataFrame:
    """
    - Convierte tokens nulos comunes a NaN
    - Colapsa espacios
    """
    df = df.copy()
    for c in df.columns:
        df[c] = df[c].map(_normalize_nulls).map(_collapse_spaces)
    return df

def apply_auto_fixes(df: pd.DataFrame, rules_pkg: dict) -> pd.DataFrame:
    """
    Aplica correcciones ligeras ANTES de validar:
      - default_if_empty: "TEXTO" -> completa si está vacío
      - uppercase: true           -> valor.upper()
      - siempre colapsa espacios
    """
    df2 = df.copy()
    for r in rules_pkg.get("rules", []):
        col = r.get("campo")
        if not col or col not in df2.columns:
            continue

        # Normalización básica previa
        df2[col] = df2[col].map(_normalize_nulls).map(_collapse_spaces)

        # Completar si vacío
        if r.get("default_if_empty") is not None:
            def_val = r["default_if_empty"]
            mask = df2[col].isna() | (df2[col].astype(str).str.strip() == "")
            if mask.any():
                df2.loc[mask, col] = def_val

        # Mayúsculas
        if r.get("uppercase") is True:
            df2[col] = df2[col].map(lambda v: v if pd.isna(v) else str(v).upper())

    return df2

# ---------------- Validaciones ----------------

def parse_date_ddmmyyyy(x):
    """Parsea con dayfirst=True. Acepta 'dd/mm/yyyy', 'dd-mm-yyyy' y variantes."""
    if pd.isna(x) or str(x).strip() == "":
        return pd.NaT
    try:
        return pd.to_datetime(parser.parse(str(x), dayfirst=True, fuzzy=True))
    except Exception:
        return pd.NaT

def check_rule_series(series: pd.Series, rule: dict) -> pd.DataFrame:
    col = rule["campo"]
    errs = []

    # Requerido
    if rule.get("obligatorio"):
        missing_idx = series[series.isna() | (series.astype(str).str.strip()=="")].index.tolist()
        for i in missing_idx:
            errs.append((i, col, "obligatorio", "Valor requerido faltante"))

    # Tipo
    tipo = rule.get("tipo")
    if tipo == "int":
        bad = series.dropna().astype(str).map(lambda s: re.fullmatch(r"[+-]?\d+", s) is None)
        for i in series.dropna().index[bad]:
            errs.append((i, col, "tipo:int", f"No entero: {series.loc[i]}"))
        if rule.get("min") is not None:
            try:
                minv = float(rule["min"])
                bad = series.dropna().astype(float) < minv
                for i in series.dropna().index[bad]:
                    errs.append((i, col, "min", f"< {minv}"))
            except Exception:
                pass
        if rule.get("max") is not None:
            try:
                maxv = float(rule["max"])
                bad = series.dropna().astype(float) > maxv
                for i in series.dropna().index[bad]:
                    errs.append((i, col, "max", f"> {maxv}"))
            except Exception:
                pass

    elif tipo == "float":
        def can_float(s):
            try:
                float(str(s).replace(",", ""))
                return True
            except:
                return False
        bad = series.dropna().map(lambda s: not can_float(s))
        for i in series.dropna().index[bad]:
            errs.append((i, col, "tipo:float", f"No número: {series.loc[i]}"))

    elif tipo == "date":
        parsed = series.map(parse_date_ddmmyyyy)
        # formato inválido
        bad2 = series.notna() & parsed.isna() & (series.astype(str).str.strip()!="")
        for i in series.index[bad2]:
            errs.append((i, col, "tipo:date", f"No válido (use dd/mm/aaaa): {series.loc[i]}"))
        # rango
        bad = parsed.notna() & ((parsed > pd.Timestamp.today()) | (parsed < pd.Timestamp("1900-01-01")))
        for i in parsed.index[bad]:
            errs.append((i, col, "rango_fecha", "Fuera de rango (1900..hoy)"))

    # Regex
    if rule.get("regex"):
        pat = re.compile(str(rule["regex"]))
        bad = series.dropna().astype(str).map(lambda s: re.fullmatch(pat, s) is None)
        for i in series.dropna().index[bad]:
            errs.append((i, col, "regex", "No cumple patrón"))

    # Catálogo
    if rule.get("catalogo"):
        allowed = set([str(v) for v in rule["catalogo"]])
        bad = series.dropna().astype(str).map(lambda s: s not in allowed)
        for i in series.dropna().index[bad]:
            errs.append((i, col, "catalogo", f"Valor no permitido: {series.loc[i]}"))

    # Longitudes
    if rule.get("len_min") is not None:
        bad = series.dropna().astype(str).map(lambda s: len(s) < int(rule["len_min"]))
        for i in series.dropna().index[bad]:
            errs.append((i, col, "len_min", f"Longitud < {int(rule['len_min'])}"))
    if rule.get("len_max") is not None:
        bad = series.dropna().astype(str).map(lambda s: len(s) > int(rule["len_max"]))
        for i in series.dropna().index[bad]:
            errs.append((i, col, "len_max", f"Longitud > {int(rule['len_max'])}"))

    return pd.DataFrame(errs, columns=["fila", "columna", "regla", "detalle"])

def check_uniques(df: pd.DataFrame, rules: list) -> pd.DataFrame:
    errs = []
    for r in rules:
        if r.get("unico"):
            col = r["campo"]
            if col in df.columns:
                dup = df[col].astype(str).duplicated(keep=False) & df[col].notna()
                for i in df.index[dup]:
                    errs.append((i, col, "unique", "Valor duplicado"))
    if not errs:
        return pd.DataFrame(columns=["fila", "columna", "regla", "detalle"])
    return pd.DataFrame(errs, columns=["fila", "columna", "regla", "detalle"])

def check_conditionals(df: pd.DataFrame, rules: list) -> pd.DataFrame:
    errs = []
    for r in rules:
        cond = r.get("condicion")
        target = r["campo"]
        if not cond:
            continue
        m_eq = re.fullmatch(r"\s*([A-Za-z0-9_ /]+)\s*=\s*([^=]+)\s*", str(cond))
        m_in = re.fullmatch(r"\s*([A-Za-z0-9_ /]+)\s*in\s*\[([^\]]+)\]\s*", str(cond), re.IGNORECASE)
        if m_eq:
            a = m_eq.group(1).strip()
            v = m_eq.group(2).strip().strip("'\"")
            if a in df.columns and target in df.columns:
                mask = df[a].astype(str) == v
                miss = df[target].isna() | (df[target].astype(str).str.strip()=="")
                bad_idx = df.index[mask & miss].tolist()
                for i in bad_idx:
                    errs.append((i, target, "condicion", f"Obligatorio si {a}={v}"))
        elif m_in:
            a = m_in.group(1).strip()
            vals = [x.strip().strip("'\"") for x in m_in.group(2).split("|")]
            if a in df.columns and target in df.columns:
                mask = df[a].astype(str).isin(vals)
                miss = df[target].isna() | (df[target].astype(str).str.strip()=="")
                bad_idx = df.index[mask & miss].tolist()
                for i in bad_idx:
                    errs.append((i, target, "condicion", f"Obligatorio si {a} in {vals}"))
    if not errs:
        return pd.DataFrame(columns=["fila","columna","regla","detalle"])
    return pd.DataFrame(errs, columns=["fila","columna","regla","detalle"])

def validate_dataframe(df: pd.DataFrame, rules_pkg: dict) -> pd.DataFrame:
    rules = rules_pkg.get("rules", [])
    errs_list = []
    for r in rules:
        col = r["campo"]
        if col not in df.columns:
            continue
        errs_list.append(check_rule_series(df[col], r))
    errs_list.append(check_uniques(df, rules))
    errs_list.append(check_conditionals(df, rules))
    errs = pd.concat(errs_list, ignore_index=True) if errs_list else pd.DataFrame(columns=["fila","columna","regla","detalle"])
    if not errs.empty:
        errs = errs.sort_values(["fila","columna"]).reset_index(drop=True)
    return errs
