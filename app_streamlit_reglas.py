# app_streamlit_reglas.py
import io, os, re, yaml, difflib, datetime
import pandas as pd
import streamlit as st
from validator import apply_non_destructive_normalization, apply_auto_fixes, validate_dataframe

from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

MAESTRO_SHEET = "MAESTRO"

# ----- Posiciones (0-based) -----
UPPERCASE_N = 4            # columnas 0..3 -> MAYÚSCULAS
TIPO_DOC_POS = 5           # col 6 (1-based) -> *TIPO DOCUMENTO*
DOC_NUM_POS  = 6           # col 7 (1-based) -> *N. DOCUMENTO*
EMAIL_POS    = 7           # col 8 (1-based) -> *E-MAIL*
# FECHAS por POSICIÓN 0-based (obligatorias, dd/mm/yyyy)
DATE_POSITIONS = [12, 13, 14, 15]

OPTIONAL_COLUMNS = {
    "ZONA",
    "TIPO CALLE",
    "TIPO DE CONJUNTO",
    "TIPO VIVIENDA",
    "RUTA",
    "ORDEN RUTA",
    "CALLE PRINCIPAL",
    "NUMERO DE CASA",
    "CALLE SECUNDARIA",
    "LATITUD",
    "LONGITUD",
}

# ----- Dominios de e-mail permitidos -----
ALLOWED_DOMAINS = [
    "gmail.com",
    "hotmail.com", "hotmail.es",
    "outlook.com", "outlook.es",
    "yahoo.com",
]
ALLOWED_DOMAINS_STRIPPED = {d.replace(".", ""): d for d in ALLOWED_DOMAINS}
COMMON_DOMAIN_TYPO_MAP = {
    "gmai.com": "gmail.com",
    "gmaii.com": "gmail.com",
    "gmal.com": "gmail.com",
    "gmail.con": "gmail.com",
    "gmail,com": "gmail.com",
    "hotmail.con": "hotmail.com",
    "hotmai.com": "hotmail.com",
    "hotmal.com": "hotmail.com",
    "hotmail,es": "hotmail.es",
    "outlok.com":  "outlook.com",
    "outlok.es":   "outlook.es",
    "outlook,es":  "outlook.es",
    "yahho.com":   "yahoo.com",
    "yaho.com":    "yahoo.com",
}

st.set_page_config(page_title="Validador MAESTRO ", layout="wide")
st.title("Validador MAESTRO")

# =================== LECTURA INTELIGENTE DE ENCABEZADOS ===================
HEADER_ROW_INDEX = 1  # segunda fila (0-based)

def smart_read_maestro(file) -> pd.DataFrame:
    """Lee la hoja MAESTRO y detecta si la primera fila es un título/merge.
    Si lo es, usa la fila 2 como encabezado y devuelve solo las filas de datos.
    """
    df0 = pd.read_excel(file, sheet_name=MAESTRO_SHEET, header=None, engine="openpyxl")

    first_row = df0.iloc[0].astype(str)
    many_unnamed = (first_row.str.startswith("Unnamed")).mean() > 0.5
    has_title = first_row.str.contains("INFORMACION", case=False, na=False).any()

    header_row = HEADER_ROW_INDEX if (many_unnamed or has_title) else 0

    # Heurística extra: si la fila 2 contiene "* CODIGO CONTRATO", usa fila 2
    try:
        row1 = df0.iloc[1].astype(str).str.strip()
        if "* CODIGO CONTRATO" in row1.values:
            header_row = HEADER_ROW_INDEX
    except Exception:
        pass

    header = df0.iloc[header_row].astype(str).str.strip().tolist()
    df = df0.iloc[header_row + 1 : ].copy()
    df.columns = header
    df.reset_index(drop=True, inplace=True)

    # Limpia posibles "Unnamed"
    df.columns = [c if not str(c).lower().startswith("unnamed") else "" for c in df.columns]
    return df

# =================== CARGA DE REGLAS (estructura/fechas) ===================
@st.cache_data
def load_rules():
    with open("rules/schema_rules.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)
rules_pkg = load_rules()
COLUMN_ORDER = rules_pkg.get("column_order", None)

# ---------- Helpers encabezados (solo visual) ----------
def is_unnamed(col_name) -> bool:
    return str(col_name).lower().startswith("unnamed")

def build_column_config(df: pd.DataFrame):
    cfg = {}
    for c in df.columns:
        label = "" if is_unnamed(c) else str(c)
        cfg[c] = st.column_config.Column(label=label)
    return cfg

# ---------- Columnas duplicadas: helpers ----------
def make_unique_labels(labels):
    seen = {}
    out = []
    for name in labels:
        name = "" if name is None else str(name)
        if name not in seen:
            seen[name] = 1
            out.append(name)
        else:
            seen[name] += 1
            out.append(f"{name} ({seen[name]})")
    return out

def ui_df_with_unique_columns(df: pd.DataFrame):
    """Devuelve una copia con nombres únicos SOLO para UI (misma forma y orden)."""
    dfd = df.copy()
    dfd.columns = make_unique_labels(df.columns)
    return dfd

# ---------- Helpers fechas ----------
def date_columns_from_rules(rules_pkg):
    return [r["campo"] for r in rules_pkg.get("rules", []) if str(r.get("tipo")) == "date"]

def ensure_datetime_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Convierte columnas de fecha a datetime (para escribir sin hora)."""
    out = df.copy()
    for col in date_columns_from_rules(rules_pkg):
        if col in out.columns:
            out[col] = pd.to_datetime(out[col], errors="coerce")
    for pos in DATE_POSITIONS:
        if pos < len(out.columns):
            c = out.columns[pos]
            out[c] = pd.to_datetime(out[c], errors="coerce")
    return out

def format_dates_for_display(df: pd.DataFrame, rules_pkg):
    """Solo para mostrar en la UI: dd/mm/yyyy."""
    out = df.copy()
    for col in date_columns_from_rules(rules_pkg):
        if col in out.columns:
            s = pd.to_datetime(out[col], errors="coerce")
            out[col] = s.dt.strftime("%d/%m/%Y")
    for pos in DATE_POSITIONS:
        if pos < len(out.columns):
            c = out.columns[pos]
            s = pd.to_datetime(out[c], errors="coerce")
            out[c] = s.dt.strftime("%d/%m/%Y")
    return out

# ---------- Transformaciones ----------
def force_uppercase_first_cols(df: pd.DataFrame, n: int) -> pd.DataFrame:
    df2 = df.copy()
    for c in list(df2.columns)[:max(0, n)]:
        df2[c] = df2[c].map(lambda v: v if pd.isna(v) else str(v).upper())
    return df2

def digits_only_column(df: pd.DataFrame, pos: int) -> pd.DataFrame:
    if pos >= len(df.columns):
        return df
    c = df.columns[pos]
    df2 = df.copy()
    df2[c] = df2[c].map(lambda v: v if pd.isna(v) else re.sub(r"[^0-9]", "", str(v)))
    return df2

def clean_mac(value: str) -> str:
    """Limpia la MAC quitando separadores y dejando solo hexadecimales."""
    if pd.isna(value):
        return ""
    s = str(value).strip()
    # eliminar todos los separadores y símbolos comunes
    s = re.sub(r'[^0-9A-Fa-f]', '', s)
    # devolver limpio pero sin validar aún
    return s.upper()

def normalize_vlan(df: pd.DataFrame) -> pd.DataFrame:
    if "VLAN" not in df.columns:
        return df
    df2 = df.copy()
    c = "VLAN"
    for i, v in df2[c].items():
        if pd.isna(v) or str(v).strip() == "":
            continue
        s = str(v).strip()
        # convertir 1003.0 -> 1003
        if re.fullmatch(r"\d+\.0+", s):
            df2.at[i, c] = s.split(".")[0]
        # eliminar caracteres NO numéricos
        df2.at[i, c] = re.sub(r"[^\d]", "", str(df2.at[i, c]))
    return df2

# ---------- Normalización de texto: acentos/ñ -> con apóstrofe pegado ----------
def normalize_text_apostrophe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Reemplaza acentos y ñ por su versión con apóstrofe pegado en TODAS las columnas de texto,
    excepto en columnas sensibles (E-MAIL/EMAIL, IP, MAC).
    Trabaja por POSICIÓN para evitar problemas con nombres duplicados.
    """
    df2 = df.copy()

    # columnas a excluir por nombre (en mayúsculas para comparar)
    excluded_cols = {"* E-MAIL", "E-MAIL", "EMAIL", "IP", "MAC"}

    # mapeo de reemplazos
    repl = {
        "á": "a'", "é": "e'", "í": "i'", "ó": "o'", "ú": "u'",
        "Á": "A'", "É": "E'", "Í": "I'", "Ó": "O'", "Ú": "U'",
        "ñ": "n", "Ñ": "N",
    }

    def _fix_text(val):
        if pd.isna(val):
            return val
        s = str(val)
        for k, v in repl.items():
            s = s.replace(k, v)
        return s

    # por posición (para no chocar con nombres duplicados)
    for j in range(df2.shape[1]):
        col_name_upper = str(df2.columns[j]).strip().upper()
        if col_name_upper in excluded_cols:
            continue
        col = df2.iloc[:, j]
        if pd.api.types.is_object_dtype(col) or pd.api.types.is_string_dtype(col):
            df2.iloc[:, j] = col.map(_fix_text)

    return df2

DOC_PREFIX_MAP = {"J": "JURIDICO", "V": "CEDULA", "G": "GOBIERNO", "E": "EXTRANJERO"}

def infer_tipo_documento_from_docnum(df: pd.DataFrame, tipo_pos: int, doc_pos: int):
    df2 = df.copy()
    errs = []
    if tipo_pos >= len(df2.columns) or doc_pos >= len(df2.columns):
        return df2, pd.DataFrame(columns=["fila","columna","col_idx","regla","detalle","severity"])
    tipo_col = df2.columns[tipo_pos]
    doc_col  = df2.columns[doc_pos]
    for i in df2.index:
        s = "" if pd.isna(df2.at[i, doc_col]) else str(df2.at[i, doc_col]).strip()
        if s == "":
            errs.append((i, tipo_col, tipo_pos, "tipo_documento_desconocido", "N. DOCUMENTO vacío", "error")); continue
        m = re.search(r"[A-Za-z]", s)
        if not m:
            df2.at[i, tipo_col] = "CEDULA"; continue
        pref = m.group(0).upper()
        mapped = DOC_PREFIX_MAP.get(pref)
        if mapped is None:
            errs.append((i, tipo_col, tipo_pos, "tipo_documento_desconocido", f"Prefijo '{pref}' no reconocido", "error"))
        else:
            df2.at[i, tipo_col] = mapped
    if not errs:
        return df2, pd.DataFrame(columns=["fila","columna","col_idx","regla","detalle","severity"])
    return df2, pd.DataFrame(errs, columns=["fila","columna","col_idx","regla","detalle","severity"])

# ---------- E-mail: normalización + autocorrección ----------
EMAIL_REGEX = re.compile(r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$')

def _clean_email_text(s: str) -> str:
    s = s.strip()
    s = s.replace(" ", "").replace(";", ",").lower()
    s = s.replace(",",".")
    s = s.replace("..", ".")
    s = s.replace("@@", "@").replace("@.", "@")
    return s

def _insert_at_using_suffix_match(s: str) -> str | None:
    if "@" in s:  # ya trae @
        return s
    best = None
    best_ratio = 0.0
    best_pos = None
    best_domain = None
    for d in ALLOWED_DOMAINS:
        L = len(d)
        for k in range(max(1, len(s)-L-2), len(s)):
            suf = s[k:]
            ratio = difflib.SequenceMatcher(None, suf, d).ratio()
            if ratio > best_ratio:
                best_ratio, best, best_pos, best_domain = ratio, suf, k, d
    if best is not None and best_ratio >= 0.72:
        local = s[:best_pos]
        if local:
            return f"{local}@{best_domain}"
    return None

def autocorrect_email(value: str):
    """Devuelve (new_value, corrected:bool, detail:str, valid:bool)."""
    original = "" if pd.isna(value) else str(value)
    s = _clean_email_text(original)

    if s == "" or s.lower() in {"none", "nan", "null"}:
        return original, False, "", False

    # Sin '@': intento insertar por sufijo aproximado o por dominio pegado.
    if "@" not in s:
        candidate = _insert_at_using_suffix_match(s)
        if candidate:
            s = candidate
        else:
            for d in ALLOWED_DOMAINS:
                idx = s.rfind(d)
                if idx > 0:
                    local = s[:idx]
                    if local:
                        s = f"{local}@{d}"
                        break

    # Si trae más de un '@', dejo solo el primero.
    if s.count("@") > 1:
        parts = s.split("@")
        s = parts[0] + "@" + "".join(parts[1:])

    if "@" not in s:
        return s, False, "", False

    local, domain = s.split("@", 1)
    if local == "":
        return s, False, "", False

    corrected = False
    detail = ""

    # Correcciones comunes directas
    if domain in COMMON_DOMAIN_TYPO_MAP:
        new_domain = COMMON_DOMAIN_TYPO_MAP[domain]
        detail = f"{domain} → {new_domain}"
        domain = new_domain
        corrected = True

    if domain in ALLOWED_DOMAINS:
        final_email = f"{local}@{domain}"
        return final_email, corrected, detail, EMAIL_REGEX.fullmatch(final_email) is not None

    # Fuzzy contra dominios permitidos
    dom_stripped = domain.replace(".", "")
    candidates = list(ALLOWED_DOMAINS_STRIPPED.keys())
    close = difflib.get_close_matches(dom_stripped, candidates, n=1, cutoff=0.8)
    if close:
        best = ALLOWED_DOMAINS_STRIPPED[close[0]]
        detail = f"{domain} → {best}" if not detail else f"{detail}; {domain} → {best}"
        domain = best
        corrected = True
    else:
        if dom_stripped in ALLOWED_DOMAINS_STRIPPED:
            best = ALLOWED_DOMAINS_STRIPPED[dom_stripped]
            detail = f"{domain} → {best}" if not detail else f"{detail}; {domain} → {best}"
            domain = best
            corrected = True

    final_email = f"{local}@{domain}"
    valid = EMAIL_REGEX.fullmatch(final_email) is not None and (domain in ALLOWED_DOMAINS)
    return final_email, corrected, detail, valid

def apply_email_autocorrect(df: pd.DataFrame, email_pos: int, enable: bool):
    corr_rows = []
    info_errors = []
    if not enable or email_pos >= len(df.columns):
        return df, pd.DataFrame(columns=["fila","original","nuevo","detalle","timestamp"]), pd.DataFrame(columns=["fila","columna","col_idx","regla","detalle","severity"])

    c = df.columns[email_pos]
    for i in df.index:
        orig = df.at[i, c]
        new_val, corrected, detail, valid = autocorrect_email(orig)
        if corrected:
            df.at[i, c] = new_val
            corr_rows.append({
                "fila": i,
                "original": orig,
                "nuevo": new_val,
                "detalle": detail if detail else "Autocorrección",
                "timestamp": datetime.datetime.utcnow().isoformat(timespec="seconds") + "Z",
            })
            info_errors.append((i, c, email_pos, "email_autocorregido", detail or "Corrección automática de dominio", "info"))

    corr_df = pd.DataFrame(corr_rows, columns=["fila","original","nuevo","detalle","timestamp"])
    info_df = pd.DataFrame(info_errors, columns=["fila","columna","col_idx","regla","detalle","severity"]) if info_errors else pd.DataFrame(columns=["fila","columna","col_idx","regla","detalle","severity"])

    # Guardar historial
    if not corr_df.empty:
        try:
            os.makedirs("history", exist_ok=True)
            path = os.path.join("history", "email_corrections.csv")
            append_header = not os.path.exists(path)
            corr_to_save = corr_df.copy()
            corr_to_save["fila"] = corr_to_save["fila"].astype(int)
            corr_to_save.to_csv(path, index=False, mode="a", header=append_header, encoding="utf-8")
        except Exception:
            pass

    return df, corr_df, info_df

# ---------- Validaciones ----------
IP_COL_NAME = "IP"
IP_REGEX = re.compile(r'^(?:(?:25[0-5]|2[0-4][0-9]|1?[0-9]?[0-9])\.){3}(?:25[0-5]|2[0-4][0-9]|1?[0-9]?[0-9])$')
PHONE_COL_NAME = "TELEFONOS"  # columna de teléfonos
MAC_COL_NAME = "MAC"
MAC_REGEX = re.compile(r'^[0-9A-Fa-f]{12}$')  # 12 caracteres hexadecimales corridos

def validate_position_rules(df: pd.DataFrame) -> pd.DataFrame:
    errs = []
    # Email obligatorio + dominio permitido
    if EMAIL_POS < len(df.columns):
        c = df.columns[EMAIL_POS]
        series = df[c].astype(str).str.strip()
        for i, v in series.items():
            ok_regex = EMAIL_REGEX.fullmatch(v or "")
            ok_domain = False
            if ok_regex:
                try:
                    domain = v.split("@", 1)[1].lower()
                    ok_domain = domain in ALLOWED_DOMAINS
                except Exception:
                    ok_domain = False
            if (v == "" or v.lower() in {"none","nan","null"}):
                errs.append((i, c, EMAIL_POS, "email_vacio", "Correo vacío", "error"))
            elif "@" not in v:
                errs.append((i, c, EMAIL_POS, "email_sin_arroba", "Correo sin '@' o dominio", "error"))
            elif (not ok_regex) or (not ok_domain):
                errs.append((i, c, EMAIL_POS, "email_invalido", "Correo inválido o dominio no permitido", "warn"))

    # Fechas obligatorias
    for pos in DATE_POSITIONS:
        if pos < len(df.columns):
            c = df.columns[pos]
            s = pd.to_datetime(df[c], errors="coerce")
            for i in df.index:
                if pd.isna(s.loc[i]):
                    errs.append((i, c, pos, "fecha_obligatoria", "Fecha vacía o inválida (dd/mm/yyyy)", "error"))

    # IP válida (4 octetos 0–255) — si existe la columna "IP"
    if IP_COL_NAME in df.columns:
        c = IP_COL_NAME
        col_idx = df.columns.get_loc(c)
        for i, v in df[c].items():
            if pd.isna(v) or str(v).strip() == "":
                continue  # si quieres que sea obligatoria, cambia a error por vacío aquí
            val = str(v).strip()
            if not IP_REGEX.fullmatch(val):
                errs.append((i, df.columns[col_idx], col_idx, "ip_invalida", "IP inválida: debe tener 4 segmentos 0–255 (ej. 172.18.9.10)", "error"))

    # TELEFONO válido — si existe la columna "TELEFONOS"
    if PHONE_COL_NAME in df.columns:
        c = PHONE_COL_NAME
        col_idx = df.columns.get_loc(c)
        for i, v in df[c].items():
            s = "" if pd.isna(v) else str(v).strip()

            if s == "":
                errs.append(
                    (
                        i,
                        c,
                        col_idx,
                        "telefono_vacio",
                        "Teléfono vacío o incompleto (se esperan 11 dígitos).",
                        "error",
                    )
                )
                continue

            # Puede venir "04121234567" o "04121234567 / 04141234567"
            parts = [p.strip() for p in s.split("/") if p.strip() != ""]

            # Si no hay partes válidas -> error
            if not parts:
                errs.append(
                    (
                        i,
                        c,
                        col_idx,
                        "telefono_invalido",
                        "Teléfono inválido (use 11 dígitos, opcionalmente separados por '/').",
                        "error",
                    )
                )
                continue

            ok = True
            for p in parts:
                digits = re.sub(r"\D", "", p)  # solo números
                if len(digits) != 11:
                    ok = False
                    break

            if not ok:
                errs.append(
                    (
                        i,
                        c,
                        col_idx,
                        "telefono_invalido",
                        "Cada número debe tener exactamente 11 dígitos (ej. 04121234567 / 04141234567).",
                        "error",
                    )
                )

    # MAC: 12 caracteres hexadecimales corridos, sin separadores
    if MAC_COL_NAME in df.columns:
        c = MAC_COL_NAME
        col_idx = df.columns.get_loc(c)

        for i, v in df[c].items():
            s = "" if pd.isna(v) else str(v).strip()

            if s == "":
                continue  # No obligatorio

            # después de limpieza debe ser 12 hexadecimales
            if not MAC_REGEX.fullmatch(s):
                errs.append(
                    (
                        i,
                        c,
                        col_idx,
                        "mac_invalida",
                        "MAC inválida: debe tener exactamente 12 caracteres hexadecimales sin símbolos. Ej: 001A2B3C4D5E",
                        "error",
                    )
                )

    # VLAN opcional pero si viene debe ser entera
    if "VLAN" in df.columns:
        c = "VLAN"
        col_idx = df.columns.get_loc(c)
        for i, v in df[c].items():
            if pd.isna(v) or str(v).strip() == "":
                continue  # es opcional

            val = str(v).strip()

            # eliminar decimales tipo "1003.0"
            if re.fullmatch(r"\d+\.0+", val):
                continue  # lo damos como válido porque igual se autocorregirá

            # validar solo enteros puros
            if not re.fullmatch(r"\d+", val):
                errs.append((i, c, col_idx,
                            "vlan_invalida",
                            "VLAN debe ser número entero (ej: 1003), sin decimales ni letras.",
                            "error"))

    # --- NUMERO DE SERIE ONT: limpiar y validar ---
    ONT_COL = "NUMERO DE SERIE ONT"
    if ONT_COL in df.columns:
        col_idx = df.columns.get_loc(ONT_COL)

        for i, v in df[ONT_COL].items():
            raw = "" if pd.isna(v) else str(v)

            # 1. Limpieza → solo letras y números
            cleaned = re.sub(r"[^A-Za-z0-9]", "", raw).upper()

            # Guardamos limpieza en el DF para mostrarla corregida
            df.at[i, ONT_COL] = cleaned

            # 2. Validación
            if cleaned == "":
                errs.append((
                    i, ONT_COL, col_idx,
                    "ont_vacio",
                    "NUMERO DE SERIE ONT es obligatorio.",
                    "error"
                ))
                continue

            if len(cleaned) not in (14, 16):
                errs.append((
                    i, ONT_COL, col_idx,
                    "ont_longitud_invalida",
                    f"Debe tener exactamente 14 o 16 caracteres (actual: {len(cleaned)}).",
                    "error"
                ))
                continue

            # 3. Verificación final: solo alfanumérico
            if not re.fullmatch(r"[A-Za-z0-9]+", cleaned):
                errs.append((
                    i, ONT_COL, col_idx,
                    "ont_formato_invalido",
                    "Debe contener solo letras y números.",
                    "error"
                ))

    # Campos obligatorios adicionales
    mandatory_cols = [
        "* CODIGO CONTRATO",
        "* NOMBRES",
        "* APELLIDOS",
        "* NOMBRES / RAZON SOCIAL",
        "* ESTADO",
        "* TIPO DOCUMENTO",
        "* N. DOCUMENTO",
        "* E-MAIL",
        "TELEFONOS",
        "* DURACION (MESES)",
        "* FECHA NACIMIENTO",
        "* FECHA CONTRATO",
        "* FECHA FIRMA",
        "FECHA ULTIMO CORTE",
        "DIA COBRO",
        "DIA CORTE",
        "PLAN DE INTERNET",
        "PRECIO DE INTERNET",
        "NUMERO DE SERIE ONT",
        "CALLE / DIRECCION",
        "IP",
        "VLAN",
        "MAC",
        "NAP",
        "OLT",
        "MIKROTIK",
        "DIRECCION",
    ]

    for col_name in mandatory_cols:
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name)
            series = df[col_name]
            for i, v in series.items():
                s = "" if pd.isna(v) else str(v).strip()
                if s == "" or s.lower() in {"none", "nan", "null"}:
                    errs.append(
                        (
                            i,
                            col_name,
                            col_idx,
                            "campo_obligatorio",
                            f"{col_name} es obligatorio y no puede estar vacío.",
                            "error",
                        )
                    )

    if not errs:
        return pd.DataFrame(columns=["fila","columna","col_idx","regla","detalle","severity"])
    return pd.DataFrame(errs, columns=["fila","columna","col_idx","regla","detalle","severity"])

# ---------- Estilos de errores ----------
def build_error_mask_by_position(df: pd.DataFrame, errores: pd.DataFrame):
    mask = pd.DataFrame("", index=df.index, columns=range(len(df.columns)))
    if errores is None or errores.empty:
        return mask
    for _, r in errores.iterrows():
        fila = int(r["fila"])
        # usa col_idx si existe; si no, marca todas las coincidencias por nombre
        if "col_idx" in r and not pd.isna(r["col_idx"]):
            j = int(r["col_idx"])
            if 0 <= j < len(df.columns) and fila in df.index:
                mask.iat[fila, j] = str(r.get("severity", "error"))
        else:
            # fallback por nombre
            col_name = str(r["columna"])
            for j, c in enumerate(df.columns):
                if str(c) == col_name and fila in df.index:
                    mask.iat[fila, j] = str(r.get("severity", "error"))
    return mask

def style_error_cells_ui(df: pd.DataFrame, errores: pd.DataFrame):
    # Creamos una copia con columnas únicas SOLO para la UI
    disp = ui_df_with_unique_columns(df)
    pos_mask = build_error_mask_by_position(df, errores)

    def apply_colors(row):
        i = row.name
        styles = []
        for j in range(len(row)):
            sev = pos_mask.iat[i, j] if i < pos_mask.shape[0] else ""
            if sev == "info":
                styles.append("background-color: #e7f7e7")
            elif sev == "warn":
                styles.append("background-color: #fff3cd")
            elif sev == "error":
                styles.append("background-color: #ffd6d6")
            else:
                styles.append("")
        return styles

    return disp.style.apply(apply_colors, axis=1)

# ---------- Exportar Excel ----------
def _blank_unnamed_headers(ws, header_row_index: int, col_names):
    """Deja en blanco los encabezados que empiezan por 'Unnamed' en el archivo."""
    for i, name in enumerate(col_names, start=1):
        if str(name).lower().startswith("unnamed"):
            ws.cell(row=header_row_index, column=i).value = ""

def write_excel_with_errors(df: pd.DataFrame, errores: pd.DataFrame, rules_pkg, title_for_unnamed="INFORMACION DEL ABONADO"):
    from openpyxl import Workbook
    df_to_write = ensure_datetime_columns(df)

    wb = Workbook(); ws = wb.active; ws.title = MAESTRO_SHEET
    for r in dataframe_to_rows(df_to_write, index=False, header=True):
        ws.append(r)

    col_names = list(df_to_write.columns)
    # Inserta título para el bloque de columnas sin nombre y deja headers en blanco
    unnamed_idxs = [i for i, c in enumerate(col_names, start=1) if str(c).lower().startswith("unnamed")]
    if unnamed_idxs:
        ws.insert_rows(1)
        start = min(unnamed_idxs); end = max(unnamed_idxs)
        ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
        cell = ws.cell(row=1, column=start)
        cell.value = title_for_unnamed
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(bold=True)
        ws.row_dimensions[1].height = 22

    header_row_index = 2 if unnamed_idxs else 1
    _blank_unnamed_headers(ws, header_row_index, col_names)

    # Bordes rojos + RELLENO rojo claro
    red = Side(style="thin", color="FF0000")
    red_border = Border(left=red, right=red, top=red, bottom=red)
    light_red_fill = PatternFill(fill_type="solid", fgColor="FFFFC7CE")  # rojo claro

    if errores is not None and not errores.empty:
        # Usamos col_idx si está disponible
        for _, r in errores.iterrows():
            fila1 = int(r["fila"])
            if "col_idx" in r and not pd.isna(r["col_idx"]):
                col_idx0 = int(r["col_idx"])
                if 0 <= col_idx0 < len(col_names):
                    excel_row = fila1 + header_row_index + 1
                    excel_col = col_idx0 + 1
                    cell = ws.cell(row=excel_row, column=excel_col)
                    cell.border = red_border
                    cell.fill = light_red_fill
            else:
                col_name = str(r["columna"])
                matches = [j for j, name in enumerate(col_names) if str(name) == col_name]
                for col_idx0 in matches:
                    excel_row = fila1 + header_row_index + 1
                    excel_col = col_idx0 + 1
                    cell = ws.cell(row=excel_row, column=excel_col)
                    cell.border = red_border
                    cell.fill = light_red_fill

    # Ancho columnas
    for i, c in enumerate(col_names, start=1):
        ws.column_dimensions[ws.cell(row=header_row_index, column=i).column_letter].width = max(12, min(35, len(str(c)) + 4))

    # Formato fechas DD/MM/YYYY
    date_cols_by_name = date_columns_from_rules(rules_pkg)
    for pos in DATE_POSITIONS:
        if pos < len(col_names):
            date_cols_by_name.append(col_names[pos])
    if date_cols_by_name:
        name_to_idx = {name: idx for idx, name in enumerate(col_names, start=1)}
        for dc in date_cols_by_name:
            if dc in name_to_idx:
                cidx = name_to_idx[dc]
                for r in range(header_row_index + 1, ws.max_row + 1):
                    ws.cell(row=r, column=cidx).number_format = "DD/MM/YYYY"

    out = io.BytesIO(); wb.save(out); return out.getvalue()

def write_excel_validated(df: pd.DataFrame, rules_pkg, title_for_unnamed="INFORMACION DEL ABONADO"):
    return write_excel_with_errors(df, errores=pd.DataFrame(), rules_pkg=rules_pkg, title_for_unnamed=title_for_unnamed)

# ---------- Helper para elegir DF vivo ----------
def pick_live_df():
    """Devuelve el DF que debe usarse: el editado si existe, sino el actual."""
    ed = st.session_state.get("edited_df", None)
    cur = st.session_state.get("current_df", None)
    return ed if ed is not None else cur

# ---------- Interfaz ----------
st.sidebar.subheader("Opciones")
enable_autocorrect = st.sidebar.checkbox("Autocorregir dominios de correo", value=True)

archivo = st.file_uploader("Sube el Excel (hoja 'MAESTRO')", type=["xlsx", "xls"])

if "current_df" not in st.session_state:
    st.session_state["current_df"] = None
if "edited_df" not in st.session_state:
    st.session_state["edited_df"] = None

if archivo:
    try:
        # Solo procesamos desde cero si aún no hay DF en sesión
        if st.session_state["current_df"] is None:
            # LECTURA INTELIGENTE (usa fila 2 como encabezado si corresponde)
            df_raw = smart_read_maestro(archivo)

            # Normalización y fixes base
            df_norm  = apply_non_destructive_normalization(df_raw)
            df_fixed = apply_auto_fixes(df_norm, rules_pkg)
            df_fixed = force_uppercase_first_cols(df_fixed, n=UPPERCASE_N)
            df_fixed, tipo_errs_all = infer_tipo_documento_from_docnum(df_fixed, TIPO_DOC_POS, DOC_NUM_POS)
            df_fixed = normalize_text_apostrophe(df_fixed)

            df_fixed = digits_only_column(df_fixed, DOC_NUM_POS)
            if "MAC" in df_fixed.columns:
                df_fixed["MAC"] = df_fixed["MAC"].map(clean_mac)

            df_fixed = normalize_vlan(df_fixed)
            df_fixed, corrections_df, info_email_errors = apply_email_autocorrect(df_fixed, EMAIL_POS, enable_autocorrect)

            st.session_state["current_df"] = df_fixed.copy()
            st.session_state["edited_df"]  = None  # se poblará desde el editor
        else:
            # Ya existe DF en sesión: usamos ese (con posibles ediciones previas)
            df_fixed = st.session_state["current_df"].copy()
            tipo_errs_all = pd.DataFrame(columns=["fila","columna","col_idx","regla","detalle","severity"])
            corrections_df = pd.DataFrame(columns=["fila","original","nuevo","detalle","timestamp"])
            info_email_errors = pd.DataFrame(columns=["fila","columna","col_idx","regla","detalle","severity"])

        # Validación desde fila 0 sobre el DF actual
        df_valid = df_fixed.copy()
        df_valid.index = range(len(df_valid))
        errores_yaml = validate_dataframe(df_valid, rules_pkg)
        errores_pos  = validate_position_rules(df_valid)

        tipo_errs = tipo_errs_all.copy()
        if not tipo_errs.empty:
            tipo_errs["fila"] = tipo_errs["fila"].astype(int)

        frames = []
        for e in (errores_yaml, errores_pos, tipo_errs, info_email_errors):
            if e is not None and not e.empty:
                if "severity" not in e.columns:
                    e = e.copy(); e["severity"] = "error"
                frames.append(e)
        errores = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=["fila","columna","col_idx","regla","detalle","severity"])

        # --- VISTA PREVIA con sombreado ---
        st.subheader("Vista previa (errores=rojo, avisos=amarillo, info=verde)")
        df_preview = format_dates_for_display(st.session_state["current_df"].head(20), rules_pkg)
        st.dataframe(style_error_cells_ui(df_preview, errores), use_container_width=True)

        # Estructura (compara con nombres REALES, no los únicos de UI)
        if COLUMN_ORDER is not None and list(st.session_state["current_df"].columns) != list(COLUMN_ORDER):
            st.error("❌ Las columnas NO coinciden exactamente con el formato esperado.")
            with st.expander("Ver detalle de columnas"):
                st.write("Esperado:", COLUMN_ORDER); st.write("Recibido:", list(st.session_state["current_df"].columns))
            st.stop()
        else:
            st.success("✅ Estructura OK (columnas y orden idénticos).")

        # Correcciones de correo aplicadas (info)
        if enable_autocorrect and corrections_df is not None and not corrections_df.empty:
            corr_show = corrections_df.copy()
            corr_show["fila"] = corr_show["fila"].astype(int)
            st.info("Se aplicaron correcciones automáticas de e-mail:")
            st.dataframe(corr_show, use_container_width=True)

        # ---- Errores ----
        if not errores.empty:
            st.warning(f"⚠️ Se encontraron {len(errores)} incidencias.")
            st.dataframe(errores, use_container_width=True)

            # ERRORES.xlsx (con borde y RELLENO rojo claro)
            st.download_button(
                "⬇️ Descargar ERRORES.xlsx",
                write_excel_with_errors(st.session_state["current_df"], errores, rules_pkg, title_for_unnamed="INFORMACION DEL ABONADO"),
                "ERRORES.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ============== EDITOR DE INCIDENCIAS (CELDA POR CELDA) ==============
            st.markdown("## Editor de incidencias (por fila/columna)")

            # Solo filas con col_idx válido (podemos ubicar la celda exacta)
            if "col_idx" in errores.columns:
                errores_editables = errores.dropna(subset=["col_idx"]).copy()
            else:
                errores_editables = pd.DataFrame(columns=errores.columns)

            if errores_editables.empty:
                st.info("No hay celdas editables con posición de columna definida.")
            else:
                errores_editables["fila"] = errores_editables["fila"].astype(int)
                errores_editables["col_idx"] = errores_editables["col_idx"].astype(int)

                def severity_icon(sev: str) -> str:
                    sev = str(sev).lower()
                    if sev == "error":
                        return "🔴 ERROR"
                    if sev == "warn":
                        return "🟡 AVISO"
                    if sev == "info":
                        return "🟢 INFO"
                    return "⚪"

                errores_editables["estado"] = errores_editables["severity"].map(severity_icon)


                # valor_actual desde el DF
                valores_actuales = []
                for _, r in errores_editables.iterrows():
                    fila_i = int(r["fila"])
                    col_i = int(r["col_idx"])
                    val = None
                    if 0 <= fila_i < st.session_state["current_df"].shape[0] and 0 <= col_i < st.session_state["current_df"].shape[1]:
                        val = st.session_state["current_df"].iat[fila_i, col_i]
                    valores_actuales.append(val)

                errores_editables["valor_actual"] = valores_actuales
                errores_editables["nuevo_valor"] = errores_editables["valor_actual"]

                editor_cols = ["estado", "fila", "columna", "col_idx", "valor_actual", "detalle", "severity", "nuevo_valor"]
                editor_view = errores_editables[editor_cols]

                edited_issues = st.data_editor(
                    editor_view,
                    use_container_width=True,
                    num_rows="fixed",
                    key="editor_issues",
                    column_config={
                        "estado": st.column_config.TextColumn("Estado", disabled=True),
                        "fila": st.column_config.NumberColumn("Fila DF", disabled=True),
                        "columna": st.column_config.TextColumn("Columna", disabled=True),
                        "col_idx": st.column_config.NumberColumn("Idx Columna", disabled=True),
                        "valor_actual": st.column_config.TextColumn("Valor actual", disabled=True),
                        "detalle": st.column_config.TextColumn("Detalle", disabled=True),
                        "severity": st.column_config.TextColumn("Severidad", disabled=True),
                        "nuevo_valor": st.column_config.TextColumn("Nuevo valor"),
                    }
                )

                # Aplicar cambios al DF vivo
                df_editado = st.session_state["current_df"].copy()
                for _, row in edited_issues.iterrows():
                    fila_i = int(row["fila"])
                    col_i = int(row["col_idx"])
                    new_val = row["nuevo_valor"]
                    if 0 <= fila_i < df_editado.shape[0] and 0 <= col_i < df_editado.shape[1]:
                        df_editado.iat[fila_i, col_i] = new_val

                st.session_state["current_df"] = df_editado.copy()
                st.session_state["edited_df"] = df_editado.copy()

            # Descarga "de todos modos" ANTES de revalidar
            st.info("¿Deseas descargar el MAESTRO corregido aunque aún haya incidencias?")
            confirm_anytime = st.checkbox("Sí, descargar MAESTRO_corregido.xlsx con posibles errores.", key="dl_anytime")
            if confirm_anytime:
                df_export_any = pick_live_df().copy()
                if enable_autocorrect:
                    df_export_any, _, _ = apply_email_autocorrect(df_export_any, EMAIL_POS, True)
                st.download_button(
                    "⬇️ Descargar MAESTRO_corregido.xlsx (con posibles errores)",
                    write_excel_validated(df_export_any, rules_pkg, title_for_unnamed="INFORMACION DEL ABONADO"),
                    "MAESTRO_corregido.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_anytime_btn"
                )

            # Revalidar (respeta edición)
            st.markdown("---")
            if st.button("Revalidar"):
                df2 = pick_live_df().copy()
                df2 = force_uppercase_first_cols(df2, n=UPPERCASE_N)
                df2, _ = infer_tipo_documento_from_docnum(df2, TIPO_DOC_POS, DOC_NUM_POS)

                # Forzar PAIS = VENEZUELA (si existe la columna)
                for country_col_name in ["PAIS", "* PAIS"]:
                    if country_col_name in df2.columns:
                        df2[country_col_name] = "VENEZUELA"

                df2 = digits_only_column(df2, DOC_NUM_POS)
                if enable_autocorrect:
                    df2, _, _ = apply_email_autocorrect(df2, EMAIL_POS, True)
                st.session_state["current_df"] = df2
                st.session_state["edited_df"] = None  # editor se recalculará

                # Recalcular errores con el DF actualizado
                df2_valid = df2.copy()
                df2_valid.index = range(len(df2_valid))
                e_yaml2 = validate_dataframe(df2_valid, rules_pkg)
                e_pos2 = validate_position_rules(df2_valid)
                frames2 = []
                for e in (e_yaml2, e_pos2):
                    if e is not None and not e.empty:
                        if "severity" not in e.columns:
                            e = e.copy(); e["severity"] = "error"
                        frames2.append(e)
                errs2 = pd.concat(frames2, ignore_index=True) if frames2 else pd.DataFrame(columns=["fila","columna","col_idx","regla","detalle","severity"])

                if errs2.empty:
                    st.success("✅ Sin errores. Puedes descargar el MAESTRO corregido:")
                    df_export = pick_live_df().copy()
                    if enable_autocorrect:
                        df_export, _, _ = apply_email_autocorrect(df_export, EMAIL_POS, True)
                    st.download_button(
                        "⬇️ Descargar Excel validado",
                        write_excel_validated(df_export, rules_pkg, title_for_unnamed="INFORMACION DEL ABONADO"),
                        "MAESTRO_corregido.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning(f"Aún hay {len(errs2)} errores. Corrige y vuelve a revalidar.")

        else:
            st.success("✅ Sin errores desde la fila 1.")
            # Editor opcional antes de descargar — UI con columnas únicas
            disp_ok = ui_df_with_unique_columns(format_dates_for_display(st.session_state["current_df"].copy(), rules_pkg))
            edited_ok_disp = st.data_editor(
                disp_ok,
                use_container_width=True, num_rows="dynamic",
                key="editor_ok"
            )
            edited_ok = edited_ok_disp.copy()
            edited_ok.columns = st.session_state["current_df"].columns
            st.session_state["edited_df"] = edited_ok.copy()

            df_export = pick_live_df().copy()
            if enable_autocorrect:
                df_export, _, _ = apply_email_autocorrect(df_export, EMAIL_POS, True)

            st.download_button(
                "⬇️ Descargar Excel validado",
                write_excel_validated(df_export, rules_pkg, title_for_unnamed="INFORMACION DEL ABONADO"),
                "MAESTRO_corregido.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.exception(e)
else:
    st.info("Sube un archivo para validar.")
