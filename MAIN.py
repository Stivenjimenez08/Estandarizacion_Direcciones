import pandas as pd
import re
import os
import unidecode

# ============= CONFIGURACIÓN MODO SCRIPT =============
RUTA_ENTRADA = "CICLO 49_PDIRECCION.xlsx"
RUTA_SALIDA = "CICLOS_PROCESADOS.xlsx"
HOJA_SALIDA = "Direcciones_Normalizadas1"

COL_DIR = "DIRECCION"

# ============= BLOQUE MERCAR =============

palabras_clave_bodega = [
    "PABELLON", "PABELL", "PTS", "SECTOR", "PU",
    "BODEGA", "BOD", "BOD.",
    "B.VERDE", "AZUL", "AMARILLAS",
    "BQ", "BG", "B.", "BODE", "BODEG", "BLOQUE"
]

def categorizar_bodega(direccion: str) -> str:
    if pd.isna(direccion):
        return ""
    direccion = str(direccion).upper()
    return "BODEGA" if any(
        re.search(rf'\b{p}\b', direccion, re.IGNORECASE)
        for p in palabras_clave_bodega
    ) else ""

def extraer_bodega(direccion: str) -> str:
    if pd.isna(direccion):
        return ""
    direccion = str(direccion).upper()

    colores = {
        "BODEGA VERDE": "VERDE", "B.VERDE": "VERDE",
        "BODEGA AMARILLA": "AMARILLA", "AMARILLAS": "AMARILLA",
        "BODEGA AZUL": "AZUL", "3 AZUL": "AZUL",
        "PLATANEROS": "PLATANOS", "PLATANERO": "PLATANOS",
        "PLATANERA": "PLATANOS", "PLATANOS": "PLATANOS",
        "PLATANO": "PLATANOS",
        "BODEGA VERDURAS": "DE VERDURAS",
        "DE VERDURAS": "DE VERDURAS",
        "VERDURA": "DE VERDURAS",
    }
    for key, value in colores.items():
        if key in direccion:
            return value

    # BODEGA 3, BG 5, BQ 2-3, etc.
    m = re.search(
        r'\b(BODEGA|BOD|BG|BQ|PABELLON|BLOQUE|B\.VERDE|SECCION)\s+'
        r'(\d+[A-Z]?(?:[-,]\d+[A-Z]?)*)\b',
        direccion
    )
    if m:
        return m.group(2)
    return ""

def extraer_puesto(direccion: str) -> str:
    if pd.isna(direccion):
        return ""
    direccion = str(direccion).upper()
    matches = re.findall(
        r'\b(?:LOCAL|LOC|LC|LCAL|PT|PU|PTO|PTOS|L|MERCAR)\s+'
        r'(\d+[A-Z]?(?:[-,]\d+[A-Z]?)*)\b',
        direccion
    )
    puestos = set()
    for match in matches:
        numeros = re.split(r'[-,\s]+', match.strip())
        for n in numeros:
            if n:
                puestos.add(n)

    # ordenar por parte numérica
    def clave(x: str):
        m = re.search(r'\d+', x)
        return int(m.group()) if m else float('inf')

    puestos_ordenados = sorted(puestos, key=clave)
    return "-".join(puestos_ordenados).strip("-")


regex_merc = re.compile(
    r'^(LOC MERCAR (BODEGA(?: DE VERDURAS| PLATANOS| VERDE| AZUL| AMARILLA)?'
    r'(?: \d+[A-Z]?)?|BG \d+[A-Z]?)'
    r'(?: PTO \d+[A-Z]?(?:[-,]\d+[A-Z]?)*)?'
    r'|KMT 2 ARMENIA PLAZA BG \d+(?:-\d+)?)$',
    re.IGNORECASE
)

# ============= BLOQUE NUEVO ARMENIA =============

ROMAN_RE = re.compile(r"^(I{1,3}|IV|V|VI|VII|VIII|IX|X)$")
NUM_ALFA_RE = re.compile(r"^\d+[A-Z]{0,2}$")
TOKEN_ET = {"ET", "ETAPA", "ETP", "TP", "P"}
TOKEN_MZ = {"MNZ", "MZ", "MNZA", "MNA"}
TOKEN_CS = {"CS", "CASA", "C"}
TOKEN_PI = {"PI", "PISO"}
TOKEN_AP = {"AP"}

def romano_a_arabigo(romano: str) -> str | None:
    valores = {"I": 1, "V": 5, "X": 10}
    total = 0
    prev = 0
    for letra in reversed(romano):
        actual = valores.get(letra, 0)
        if actual < prev:
            total -= actual
        else:
            total += actual
            prev = actual
    return str(total) if total > 0 else None

def normalizar_armenia(direccion: str) -> tuple[str, int]:
    if pd.isna(direccion):
        return direccion, 0

    txt = unidecode.unidecode(str(direccion).upper())
    txt = re.sub(r"[-.,]", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    txt = re.sub(r"\bBAJOS\b", " PI 1", txt)
    txt = re.sub(r"\bALTOS\b", " PI 2", txt)
    tokens = txt.split()

    # Debe contener NUEVO ARMENIA
    if not ("NUEVO" in tokens and "ARMENIA" in tokens):
        return direccion, 0

    etapa = manzana = casa = piso = ap = None

    # 1) Buscar ETAPA explícita
    for i, tk in enumerate(tokens):
        if tk in TOKEN_ET and i + 1 < len(tokens):
            nxt = tokens[i + 1]
            etapa = romano_a_arabigo(nxt) if ROMAN_RE.fullmatch(nxt) else nxt
            break

    # 2) Si no hay ETAPA, intentar tomar el token antes de MZ como etapa
    if etapa is None:
        for i in range(1, len(tokens)):
            if tokens[i] in TOKEN_MZ:
                cand = tokens[i - 1]
                if ROMAN_RE.fullmatch(cand):
                    etapa = romano_a_arabigo(cand)
                elif cand.isdigit():
                    etapa = cand
                break

    # 3) Barrido para MZ / CS / PI / AP
    i = 0
    while i < len(tokens):
        tk = tokens[i]
        if tk in TOKEN_MZ and i + 1 < len(tokens):
            manzana = tokens[i + 1]
            i += 2
            continue
        if tk in TOKEN_CS and i + 1 < len(tokens):
            casa = tokens[i + 1]
            i += 2
            continue
        if tk in TOKEN_PI and i + 1 < len(tokens):
            piso = tokens[i + 1]
            i += 2
            continue
        if tk in TOKEN_AP and i + 1 < len(tokens):
            ap = tokens[i + 1]
            i += 2
            continue
        if manzana and casa is None and NUM_ALFA_RE.fullmatch(tk):
            casa = tk
        i += 1

    if manzana and casa:
        out = ["URB NUEVO ARMENIA", "MZ", manzana, "CS", casa]
        if piso:
            out += ["PI", piso]
        if ap:
            out += ["AP", ap]
        if etapa:
            out += ["ET", etapa]
        return " ".join(out), 1

    return direccion, 0

# ============= FUNCIÓN PÚBLICA =============

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Espera un DataFrame con al menos:
      - 'DIRECCION'
      - 'CLIENTE_ID' o 'NIU'

    Devuelve:
      NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION

    Incluye:
      - MERCAR / ARMENIA PLAZA
      - URB NUEVO ARMENIA
    """
    if COL_DIR not in df_in.columns:
        raise ValueError("El DataFrame de entrada debe tener columna 'DIRECCION'.")

    df = df_in.copy()

    # Identificador estándar
    if "NIU" in df.columns:
        id_col = "NIU"
    elif "CLIENTE_ID" in df.columns:
        id_col = "CLIENTE_ID"
    else:
        raise ValueError("No se encontró columna 'NIU' ni 'CLIENTE_ID'.")

    df = df[[id_col, COL_DIR]].copy()
    df.rename(columns={id_col: "NIU", COL_DIR: "DIRECCION"}, inplace=True)

    # --- BLOQUE MERCAR ---
    df_merc = df.copy()
    df_merc["INDICADOR"] = df_merc["DIRECCION"].str.split(expand=True)[0]
    df_merc["PREFIJO_LOC"] = "LOC"
    df_merc["LOCALIDAD"] = "MERCAR"

    df_merc["CATEGORIA"] = df_merc["DIRECCION"].apply(categorizar_bodega)
    df_merc["# BODEGA"] = df_merc["DIRECCION"].apply(extraer_bodega)
    df_merc["PUESTO #"] = df_merc["DIRECCION"].apply(extraer_puesto)

    indicadores_a_reemplazar = {"KMT", "CLL", "SAS", "URB", "VIA", "GAL"}
    mask_reemplazo = (
        df_merc["# BODEGA"].str.strip().ne("") &
        df_merc["PUESTO #"].str.strip().ne("") &
        df_merc["INDICADOR"].isin(indicadores_a_reemplazar)
    )
    df_merc.loc[mask_reemplazo, "INDICADOR"] = "PTO"

    df_merc["INDICADOR"] = df_merc["INDICADOR"].replace("LOC", "PTO")
    df_merc.loc[
        df_merc["DIRECCION"].str.contains("ARMENIA PLAZA", case=False, na=False),
        "INDICADOR"
    ] = "ARMENIA PLAZA"
    df_merc.loc[df_merc["# BODEGA"].str.strip().ne(""), "CATEGORIA"] = "BODEGA"

    df_merc["DIRECCION_NORMALIZADA"] = df_merc.apply(
        lambda row: row["DIRECCION"]
        if any(pd.isna(row[col]) or str(row[col]).strip() == ""
               for col in ["CATEGORIA", "# BODEGA", "INDICADOR", "PUESTO #"])
        else f'{row["PREFIJO_LOC"]} {row["LOCALIDAD"]} {row["CATEGORIA"]} '
             f'{row["# BODEGA"]} {row["INDICADOR"]} {row["PUESTO #"]}',
        axis=1
    )

    df_merc["CUMPLE_ESTANDAR"] = df_merc["DIRECCION_NORMALIZADA"].apply(
        lambda d: 1 if isinstance(d, str) and regex_merc.match(d.strip()) else 0
    )

    # Filtro de líneas MERCAR
    filtro_merc = df_merc["DIRECCION"].str.contains(
        r"armenia plaza|mercar|pabellon|mayorista|bodega|verduras|verdur|azul|verde|amar|plastico",
        case=False, na=False
    )

    df_filtrado_merc = df_merc[filtro_merc].copy()

    # --- BLOQUE NUEVO ARMENIA ---
    df_arm = df.copy()
    df_arm[["DIRECCION_NORMALIZADA", "CUMPLE_ESTANDAR"]] = df_arm["DIRECCION"].apply(
        lambda x: pd.Series(normalizar_armenia(x))
    )
    filtro_arm = df_arm["DIRECCION"].str.contains("nuevo armenia", case=False, na=False)
    df_filtrado_arm = df_arm[filtro_arm].copy()

    # --- CONSOLIDADO ---
    cols_base = ["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "CUMPLE_ESTANDAR"]
    partes = []
    if not df_filtrado_merc.empty:
        partes.append(df_filtrado_merc[cols_base])
    if not df_filtrado_arm.empty:
        partes.append(df_filtrado_arm[cols_base])

    if not partes:
        return pd.DataFrame(columns=["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"])

    df_final = pd.concat(partes, ignore_index=True)

    df_final["VALIDACION"] = df_final["CUMPLE_ESTANDAR"].astype(int).astype(str)
    df_final = df_final[["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"]]

    return df_final

# ============= MODO SCRIPT STANDALONE =============

if __name__ == "__main__":
    df_in = pd.read_excel(RUTA_ENTRADA, dtype=str)
    df_out = procesar(df_in)

    total = df_out.shape[0]
    normalizadas = (df_out["VALIDACION"] == "1").sum()
    efectividad = (normalizadas * 100 / total) if total > 0 else 0

    modo = "a" if os.path.exists(RUTA_SALIDA) else "w"
    with pd.ExcelWriter(RUTA_SALIDA, engine="openpyxl", mode=modo) as writer:
        df_out.to_excel(writer, sheet_name=HOJA_SALIDA, index=False)

    print(f"✅ Archivo procesado: {RUTA_SALIDA}")
    print(f"📄 Hoja: {HOJA_SALIDA}")
    print(f"📦 Total direcciones procesadas: {total}")
    print(f"✔️ Direcciones normalizadas correctamente: {normalizadas}")
    print(f"📈 Efectividad global: {round(efectividad, 2)}%")
