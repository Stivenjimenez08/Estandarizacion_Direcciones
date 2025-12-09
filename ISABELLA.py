# --------------------------------------------------------------------------------------------------
import os
import re
import pandas as pd

# ================== Configuración fija ==================
RUTA_ENTRADA = "CICLO 53_PDIRECCION.csv"   # Cambia si tu archivo se llama distinto
RUTA_SALIDA  = "CICLOS_PROCESADOS.xlsx"
HOJA_SALIDA  = "C53_CR_CL"
COL_DIR      = "DIRECCION"

# ================== Utilidades de carga ==================
def leer_entrada_flexible(path, col_dir=COL_DIR) -> pd.DataFrame:
    """
    Lee CSV (coma o punto y coma) o Excel y garantiza que exista la columna de dirección.
    """
    if not os.path.exists(path):
        raise ValueError(f"No se encuentra el archivo: {path}")

    ext = os.path.splitext(path)[1].lower()

    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(path, dtype=str)
    elif ext in [".csv", ".txt"]:
        # Primero intentamos con ';', si queda una sola columna probamos con ','
        df = pd.read_csv(path, dtype=str, sep=";")
        if df.shape[1] == 1:
            df = pd.read_csv(path, dtype=str, sep=",")
    else:
        raise ValueError(f"Extensión de archivo no soportada: {ext}")

    # Normalizar nombres de columnas a mayúsculas para búsqueda
    cols_map = {c: c.upper().strip() for c in df.columns}

    # Buscar columna de dirección
    if col_dir not in df.columns:
        # Intentar variantes por mayúsculas
        for c, up in cols_map.items():
            if up.replace(" ", "") in ("DIRECCION", "DIRECCIÓN", "DIR"):
                df.rename(columns={c: col_dir}, inplace=True)
                break

    if col_dir not in df.columns:
        raise ValueError("No se encontró la columna de dirección en el archivo de entrada.")

    return df

# ================== Patrones para intersecciones ==================

# 1) CLL <cll> CR <cr> - <guion> + cola
P_CLL_CR = re.compile(
    r"""^\s*
        CLL\s*(?P<cll>\d+[A-Z]?)\s*
        (?:CR|CRA|KR|KRA|K|CL)\s*(?P<cr>\d+[A-Z]?)\s*
        (?:-|–|—|\#)\s*(?P<guion>\d+)\s*
        (?P<tail>.*)$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# 2) CRA <cra> CL <cl> - <guion> + cola
P_CRA_CL = re.compile(
    r"""^\s*
        (?:CR|CRA|KR|KRA|K)\s*(?P<cra>\d+[A-Z]?)\s*
        CLL?\s*(?P<cl>\d+[A-Z]?)\s*
        (?:-|–|—|\#)\s*(?P<guion>\d+)\s*
        (?P<tail>.*)$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# 3.a) CRA <cra> CL <cl> + cola con CS explícito (ej. "CRA 25 CL 50B CS 2")
P_CRA_CL_CS = re.compile(
    r"""^\s*
        (?:CR|CRA|KR|KRA|K)\s*(?P<cra>\d+[A-Z]?)\s*
        CLL?\s*(?P<cl>\d+[A-Z]?)\s*
        (?P<tail>.*\bCS\b.*)$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# 3.b) CRA <cra> CL <cl> <num> (sin signo explícito, ej. "CRA 18 CL 49 34")
P_CRA_CL_SIMPLE = re.compile(
    r"""^\s*
        (?:CR|CRA|KR|KRA|K)\s*(?P<cra>\d+[A-Z]?)\s*
        CLL?\s*(?P<cl>\d+[A-Z]?)\s+
        (?P<guion>\d+)\b
        (?P<tail>.*)$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# 4.a) CLL <cll> <cr> - <guion> (sin "CR" explícito, ej. "CLL 26 15 -57")
P_CLL_NUMNUM = re.compile(
    r"""^\s*
        CLL\s*(?P<cll>\d+[A-Z]?)\s+
        (?P<cr>\d+[A-Z]?)\s*
        (?:-|–|—|\#)\s*(?P<guion>\d+)\s*
        (?P<tail>.*)$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# 4.b) CRA <cra> <cl> <guion> (sin "CL" explícito, ej. "CRA 17 27 63")
P_CRA_NUMNUM = re.compile(
    r"""^\s*
        (?:CR|CRA|KR|KRA|K)\s*(?P<cra>\d+[A-Z]?)\s+
        (?P<cl>\d+[A-Z]?)\s+
        (?P<guion>\d+)\s*
        (?P<tail>.*)$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# ============= Limpieza de cola (complementos) =============

def norm_tail(tail: str) -> str:
    """
    Limpia la 'cola' de la dirección (complementos) sin perder información útil.
    Se eliminan cosas tipo NIU, 8000 sueltos, espacios extra, etc.
    """
    if tail is None:
        return ""

    t = tail.upper()

    # Quitar NIU y números asociados
    t = re.sub(r"-?\s*NIU\s*\#?\s*\d+\b", "", t)
    # Quitar 8000 si aparece como token aislado
    t = re.sub(r"\s+8000\b", "", t)
    # Quitar 'ARMENIA' final
    t = re.sub(r"\s*-\s*ARMENIA\b", "", t)

    # Unificar espacios
    t = re.sub(r"\s+", " ", t).strip()

    if not t:
        return ""
    return " " + t

# ============= Normalizador de intersecciones =============

def normalizar_direccion_interseccion(direccion: str):
    """
    Normaliza solo direcciones de INTERSECCIÓN (CLL/CR, CRA/CL y variantes tipo
    CLL 26 15 - 57, CRA 17 27 63, etc.).

    Retorna:
        (direccion_normalizada, "1") si se normaliza
        (direccion_original_limpia, "0") si no se reconoce como intersección
    """
    if pd.isna(direccion):
        return direccion, "0"

    u = str(direccion).upper().strip()

    # Normalización básica de guiones y espacios
    u = re.sub(r"[\u00A0\u2000-\u200B\u202F\u205F\u3000]", " ", u)  # espacios raros
    u = re.sub(r"[‐‒–—―-]", "-", u)                                 # distintos guiones -> '-'
    u = re.sub(r"\s+", " ", u).strip()

    # 1) CLL <cll> CR <cr> - <guion> tail
    m = P_CLL_CR.match(u)
    if m:
        out = f"CLL {m.group('cll').replace(' ', '')} CR {m.group('cr').replace(' ', '')} - {m.group('guion')}"
        tail = norm_tail(m.group("tail"))
        return (out + tail).strip(), "1"

    # 2) CRA <cra> CL <cl> - <guion> tail
    m = P_CRA_CL.match(u)
    if m and m.group("guion"):
        out = f"CRA {m.group('cra').replace(' ', '')} CL {m.group('cl').replace(' ', '')} - {m.group('guion')}"
        tail = norm_tail(m.group("tail"))
        return (out + tail).strip(), "1"

    # 3.a) CRA <cra> CL <cl> con CS en la cola
    m = P_CRA_CL_CS.match(u)
    if m:
        cra = m.group("cra").replace(" ", "")
        cl  = m.group("cl").replace(" ", "")
        tail = norm_tail(m.group("tail"))
        # Si no encontramos número, dejamos sin '-'
        out = f"CRA {cra} CL {cl}"
        return (out + tail).strip(), "1"

    # 3.b) CRA <cra> CL <cl> <num> (sin signo)
    m = P_CRA_CL_SIMPLE.match(u)
    if m:
        out = f"CRA {m.group('cra').replace(' ', '')} CL {m.group('cl').replace(' ', '')} - {m.group('guion')}"
        tail = norm_tail(m.group("tail"))
        return (out + tail).strip(), "1"

    # 4.a) CLL <cll> <cr> - <guion>
    m = P_CLL_NUMNUM.match(u)
    if m:
        out = f"CLL {m.group('cll').replace(' ', '')} CR {m.group('cr').replace(' ', '')} - {m.group('guion')}"
        tail = norm_tail(m.group("tail"))
        return (out + tail).strip(), "1"

    # 4.b) CRA <cra> <cl> <guion>
    m = P_CRA_NUMNUM.match(u)
    if m:
        out = f"CRA {m.group('cra').replace(' ', '')} CL {m.group('cl').replace(' ', '')} - {m.group('guion')}"
        tail = norm_tail(m.group("tail"))
        return (out + tail).strip(), "1"

    # No cumple intersección -> no normaliza
    return u, "0"

# ================== Función estándar para el pipeline ==================

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe un DataFrame con al menos:
        - NIU o CLIENTE_ID
        - DIRECCION

    Devuelve un DataFrame con:
        NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION

    Solo incluye direcciones que aparentan ser intersecciones CLL/CR o CRA/CL.
    """
    # Copia para no mutar el original
    df = df_in.copy()

    # --- Asegurar NIU ---
    cols_map = {c: c.upper().strip() for c in df.columns}
    if "NIU" not in df.columns:
        for c, up in cols_map.items():
            if up == "CLIENTE_ID":
                df.rename(columns={c: "NIU"}, inplace=True)
                break
    if "NIU" not in df.columns:
        raise ValueError("El DataFrame de entrada debe contener columna 'NIU' o 'CLIENTE_ID'.")

    # --- Asegurar DIRECCION ---
    if COL_DIR not in df.columns:
        for c, up in cols_map.items():
            if up.replace(" ", "") in ("DIRECCION", "DIRECCIÓN", "DIR"):
                df.rename(columns={c: COL_DIR}, inplace=True)
                break
    if COL_DIR not in df.columns:
        raise ValueError("El DataFrame de entrada debe contener columna de dirección.")

    # Nos quedamos solo con NIU + DIRECCION
    df = df[["NIU", COL_DIR]].copy()
    df.rename(columns={COL_DIR: "DIRECCION"}, inplace=True)

    # Aplicar normalizador de intersección
    df[["DIRECCION_NORMALIZADA", "VALIDACION"]] = df["DIRECCION"].apply(
        lambda x: pd.Series(normalizar_direccion_interseccion(x))
    )

    # Filtrar solo las que parecen intersección (por la salida)
    patron_inter = re.compile(r"^\s*(CLL|CRA)\b", re.IGNORECASE)
    mask = df["DIRECCION_NORMALIZADA"].astype(str).str.match(patron_inter, na=False)
    df_filtrado = df[mask].copy()

    return df_filtrado[["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"]]

# ================== Flujo principal standalone ==================

def main():
    # Leer entrada flexible (CSV / Excel)
    df = leer_entrada_flexible(RUTA_ENTRADA, col_dir=COL_DIR)

    # Asegurar NIU (igual que procesar)
    cols_map = {c: c.upper().strip() for c in df.columns}
    if "NIU" not in df.columns:
        for c, up in cols_map.items():
            if up == "CLIENTE_ID":
                df.rename(columns={c: "NIU"}, inplace=True)
                break
    if "NIU" not in df.columns:
        raise ValueError("No se encontró la columna NIU (ni CLIENTE_ID para renombrar).")

    df_out = procesar(df)

    # Estadísticas
    total = len(df_out)
    normalizadas = (df_out["VALIDACION"] == "1").sum()
    efectividad = (normalizadas / total * 100) if total > 0 else 0

    # Guardar en Excel (adjuntando si existe)
    modo = "a" if os.path.exists(RUTA_SALIDA) else "w"
    with pd.ExcelWriter(RUTA_SALIDA, engine="openpyxl", mode=modo) as writer:
        df_out.to_excel(writer, sheet_name=HOJA_SALIDA, index=False)

    print(f"✅ Archivo procesado y guardado como {RUTA_SALIDA}")
    print(f"📄 Hoja: {HOJA_SALIDA}")
    print(f"🔍 Total direcciones filtradas (intersección): {total}")
    print(f"📊 Direcciones normalizadas: {normalizadas}")
    print(f"📈 Efectividad: {round(efectividad, 2)}%")

if __name__ == "__main__":
    main()
