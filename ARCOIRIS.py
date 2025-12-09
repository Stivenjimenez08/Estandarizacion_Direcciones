import pandas as pd
import re
import os

# ============================================================
#  Utilidades comunes
# ============================================================

romanos = {
    "I": 1, "II": 2, "III": 3, "IV": 4, "V": 5,
    "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10
}


def etapa_a_numero(etapa: str | None) -> str | None:
    if etapa is None:
        return None
    etapa = (
        etapa.strip()
        .upper()
        .replace("ETAPA", "")
        .replace("ETP", "")
        .replace("ET", "")
        .replace("ETA", "")
        .strip()
    )
    if not etapa:
        return None
    return str(romanos.get(etapa, etapa))


def limpiar_espacios(txt: str) -> str:
    return re.sub(r"\s+", " ", txt or "").strip()


# ============================================================
#  ARCO IRIS
# ============================================================

regex_arco_iris = re.compile(
    r"""
    (?:URB|URBANIZACION|BRR)?\s*
    (?:ARCO\s*IRIS|ARCOIRIS)
    (?:
        .*?(?:ETAPA|ETP|ET)\s*(?P<etapa>[IVXLCDM]+|\d+)
    )?
    .*?(?:MNZ|MZN|MZ|MZA|MANZANA)\s*(?P<manzana>[A-Z0-9]{1,3})
    .*?(?:CS|CASA|C)?\s*(?P<casa>\d{1,4}[A-Z]?)
    (?:.*?(?:PI|PISO|PIS|P)\s*(?P<piso>\d{1,2}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_arco_iris(direccion: str):
    if pd.isna(direccion):
        return direccion, "0"

    direccion_up = limpiar_espacios(str(direccion).upper())
    match = regex_arco_iris.search(direccion_up)
    if not match:
        return direccion_up, "0"

    mz = match.group("manzana")
    cs = match.group("casa")
    piso = match.group("piso")
    etapa = etapa_a_numero(match.group("etapa"))

    if not (mz and cs):
        return direccion_up, "0"

    nueva = f"URB ARCO IRIS MZ {mz} CS {cs}"
    if piso:
        nueva += f" PI {int(piso)}"
    if etapa:
        nueva += f" ET {etapa}"

    return nueva.strip(), "1"


# ============================================================
#  GIBRALTAR / BOSQUES DE GIBRALTAR
# ============================================================

# URB GIBRALTAR
regex_gibraltar = re.compile(
    r"""
    (?:URB|URBANIZACION|BRR)?\s*
    GIBRALTAR
    (?:
        .*?(?:ETAPA|ETP|ET)\s*(?P<etapa>[IVXLCDM]+|\d+)
    )?
    .*?(?:MNZ|MZN|MZ|MZA|MANZANA)?\s*(?P<manzana>[A-Z0-9]{1,3})
    .*?(?:CS|CASA|C|LT|LOTE)?\s*(?P<casa>\d{1,4}[A-Z]?)
    (?:.*?(?:PI|PISO|PIS|P)\s*(?P<piso>\d{1,2}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

# URB BOSQUES DE GIBRALTAR
regex_bosques_gibraltar = re.compile(
    r"""
    (?:URB|URBANIZACION|BRR)?\s*
    (?:BOSQUES|BOSQ|BQ|BQS)\s+(?:DE\s+)?GIBRALTAR
    (?:
        .*?(?:ETAPA|ETP|ET)\s*(?P<etapa>[IVXLCDM]+|\d+)
    )?
    .*?(?:MNZ|MZN|MZ|MZA|MANZANA)?\s*(?P<manzana>[A-Z0-9]{1,3})
    .*?(?:CS|CASA|C|LT|LOTE)?\s*(?P<casa>\d{1,4}[A-Z]?)
    (?:.*?(?:PI|PISO|PIS|P)\s*(?P<piso>\d{1,2}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_gibraltar(direccion: str):
    if pd.isna(direccion):
        return direccion, "0"

    direccion_up = limpiar_espacios(str(direccion).upper())
    match = regex_gibraltar.search(direccion_up)
    if not match:
        return direccion_up, "0"

    mz = match.group("manzana")
    cs = match.group("casa")
    piso = match.group("piso")
    etapa = etapa_a_numero(match.group("etapa"))

    if not (mz and cs):
        return direccion_up, "0"

    nueva = f"URB GIBRALTAR MZ {mz} CS {cs}"
    if piso:
        nueva += f" PI {int(piso)}"
    if etapa:
        nueva += f" ET {etapa}"

    return nueva.strip(), "1"


def normalizar_bosques_gibraltar(direccion: str):
    if pd.isna(direccion):
        return direccion, "0"

    direccion_up = limpiar_espacios(str(direccion).upper())
    match = regex_bosques_gibraltar.search(direccion_up)
    if not match:
        return direccion_up, "0"

    mz = match.group("manzana")
    cs = match.group("casa")
    piso = match.group("piso")
    etapa = etapa_a_numero(match.group("etapa"))

    if not (mz and cs):
        return direccion_up, "0"

    nueva = f"URB BOSQUES DE GIBRALTAR MZ {mz} CS {cs}"
    if piso:
        nueva += f" PI {int(piso)}"
    if etapa:
        nueva += f" ET {etapa}"

    return nueva.strip(), "1"


# ============================================================
#  Aplicación general del módulo
# ============================================================

def aplicar_normalizacion(direccion: str):
    if not isinstance(direccion, str):
        return "", "0"

    d = direccion.upper()

    if "ARCO IRIS" in d or "ARCOIRIS" in d:
        return normalizar_arco_iris(d)

    if "GIBRALTAR" in d:
        # Si parece un BOSQUES DE GIBRALTAR
        if re.search(r"\b(BOSQUES|BOSQ|BQ|BQS)\b.*GIBRALTAR", d):
            return normalizar_bosques_gibraltar(d)
        return normalizar_gibraltar(d)

    return d.strip(), "0"


# ============================================================
#  Función estándar para el script maestro
# ============================================================

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe un DataFrame con al menos:
      - NIU o CLIENTE_ID
      - DIRECCION

    Devuelve SOLO filas de:
      - URB ARCO IRIS
      - URB GIBRALTAR
      - URB BOSQUES DE GIBRALTAR

    con columnas:
      NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION
    """
    df = df_in.copy()
    cols_upper = {c: c.upper().strip() for c in df.columns}

    # Asegurar NIU
    if "NIU" not in df.columns:
        for c, up in cols_upper.items():
            if up == "CLIENTE_ID":
                df.rename(columns={c: "NIU"}, inplace=True)
                break
        if "NIU" not in df.columns:
            raise ValueError("El DataFrame no tiene columna NIU ni CLIENTE_ID.")

    # Asegurar DIRECCION
    if "DIRECCION" not in df.columns:
        for c, up in cols_upper.items():
            if up.replace(" ", "") == "DIRECCION":
                df.rename(columns={c: "DIRECCION"}, inplace=True)
                break
        if "DIRECCION" not in df.columns:
            raise ValueError("El DataFrame no tiene columna DIRECCION.")

    # Aplicar normalización
    df[["DIRECCION_NORMALIZADA", "VALIDACION"]] = df["DIRECCION"].apply(
        lambda x: pd.Series(aplicar_normalizacion(x))
    )

    # Filtrar solo ARCO IRIS / GIBRALTAR
    filtro = df["DIRECCION"].str.contains(
        r"ARCO IRIS|ARCOIRIS|GIBRALTAR",
        case=False,
        na=False,
    )
    df_filtrado = df[filtro].copy()

    columnas_salida = ["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"]
    for col in columnas_salida:
        if col not in df_filtrado.columns:
            df_filtrado[col] = None

    return df_filtrado[columnas_salida]


# ============================================================
#  Ejecución individual opcional
# ============================================================

if __name__ == "__main__":
    ruta_entrada = "CICLO 49_PDIRECCION.xlsx"
    ruta_salida = "CICLOS_PROCESADOS.xlsx"

    # CLIENTE_ID → NIU como en los otros scripts
    df_local = pd.read_excel(ruta_entrada, usecols=["CLIENTE_ID", "DIRECCION"])
    df_resultado = procesar(df_local)

    modo = "a" if os.path.exists(ruta_salida) else "w"
    with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode=modo) as writer:
        df_resultado.to_excel(writer, sheet_name="ARCO_GIBRALTAR", index=False)

    total = df_resultado.shape[0]
    normalizadas = df_resultado[df_resultado["VALIDACION"] == "1"].shape[0]
    efec = (normalizadas / total * 100) if total > 0 else 0

    print(f"Total direcciones procesadas: {total}")
    print(f"Direcciones normalizadas correctamente: {normalizadas}")
    print(f"Efectividad: {efec:.2f}%")
