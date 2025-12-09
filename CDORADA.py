import pandas as pd
import re
import os

# ============================================================
#  Utilidades comunes
# ============================================================

def limpiar_espacios(txt: str) -> str:
    return re.sub(r"\s+", " ", txt or "").strip()


# ============================================================
#  CIUDAD DORADA
# ============================================================

regex_ciudad_dorada = re.compile(
    r"""
    (?:BRR|URB|URBANIZACION|CND|CONJ|CONJUNTO)?\s*
    (?:CIUDAD\s+)?DORADA
    .*?(?:MNZ|MZN|MZ|MZA|MANZANA)?\s*(?P<manzana>[A-Z0-9]{1,3})
    .*?(?:CS|CASA|C|LT|LOTE)?\s*(?P<casa>\d{1,4}[A-Z]?)
    (?:.*?(?:PI|PISO|PIS|P)\s*(?P<piso>\d{1,2}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_ciudad_dorada(direccion: str):
    """
    Intenta normalizar direcciones de CIUDAD DORADA a formato:
      URB CIUDAD DORADA MZ X CS Y [PI Z]
    """
    if pd.isna(direccion):
        return direccion, "0"

    d = limpiar_espacios(str(direccion).upper())

    match = regex_ciudad_dorada.search(d)
    if not match:
        return d, "0"

    mz = match.group("manzana")
    cs = match.group("casa")
    piso = match.group("piso")

    if not (mz and cs):
        return d, "0"

    nueva = f"URB CIUDAD DORADA MZ {mz} CS {cs}"
    if piso:
        nueva += f" PI {int(piso)}"

    return nueva.strip(), "1"


# ============================================================
#  COOP / COOPERATIVO (ej. URB COOPERATIVO CIUDAD DORADA, etc.)
# ============================================================

regex_coop = re.compile(
    r"""
    (?:URB|URBANIZACION|BRR|CJT|CND)?\s*
    (?:COOP(?:ERATIVO)?|COOPERATIVO|COOPERATIVA)
    (?:\s+CIUDAD\s+DORADA|\s+DORADA)?
    .*?(?:MNZ|MZN|MZ|MZA|MANZANA)?\s*(?P<manzana>[A-Z0-9]{1,3})
    .*?(?:CS|CASA|C|LT|LOTE)?\s*(?P<casa>\d{1,4}[A-Z]?)
    (?:.*?(?:PI|PISO|PIS|P)\s*(?P<piso>\d{1,2}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_coop(direccion: str):
    """
    Intenta normalizar direcciones de tipo COOPERATIVO / COOP DORADA a formato:
      URB COOPERATIVO CIUDAD DORADA MZ X CS Y [PI Z]
    """
    if pd.isna(direccion):
        return direccion, "0"

    d = limpiar_espacios(str(direccion).upper())

    match = regex_coop.search(d)
    if not match:
        return d, "0"

    mz = match.group("manzana")
    cs = match.group("casa")
    piso = match.group("piso")

    if not (mz and cs):
        return d, "0"

    nueva = f"URB COOPERATIVO CIUDAD DORADA MZ {mz} CS {cs}"
    if piso:
        nueva += f" PI {int(piso)}"

    return nueva.strip(), "1"


# ============================================================
#  Aplicación general del módulo
# ============================================================

def aplicar_normalizacion(direccion: str):
    """
    Decide si la dirección es de CIUDAD DORADA o de COOPERATIVO
    y llama al normalizador correspondiente.
    """
    if not isinstance(direccion, str):
        return "", "0"

    d = direccion.upper()

    # CIUDAD DORADA
    if "DORADA" in d and "COOP" not in d and "COOPERAT" not in d:
        return normalizar_ciudad_dorada(d)

    # COOPERATIVO / COOP CIUDAD DORADA
    if "COOP" in d or "COOPERAT" in d:
        return normalizar_coop(d)

    # Si solo dice "DORADA" (muy suelto), igual intentamos con ciudad dorada
    if "DORADA" in d:
        return normalizar_ciudad_dorada(d)

    return limpiar_espacios(d), "0"


# ============================================================
#  Función estándar para el script maestro
# ============================================================

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe un DataFrame con al menos:
      - NIU o CLIENTE_ID
      - DIRECCION

    Devuelve SOLO filas de:
      - URB CIUDAD DORADA
      - URB COOPERATIVO CIUDAD DORADA (COOP / COOPERATIVO / COOPERATIVA)

    con columnas:
      NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION
    """
    df = df_in.copy()

    # Normalizar nombres de columnas
    cols_upper = {c: c.upper().strip() for c in df.columns}

    # Asegurar columna NIU
    if "NIU" not in df.columns:
        for c, up in cols_upper.items():
            if up == "CLIENTE_ID":
                df.rename(columns={c: "NIU"}, inplace=True)
                break
        if "NIU" not in df.columns:
            raise ValueError("El DataFrame no tiene columna NIU ni CLIENTE_ID.")

    # Asegurar columna DIRECCION
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

    # Filtro de interés: DORADA / COOP / COOPERATIVO
    filtro = df["DIRECCION"].str.contains(
        r"DORADA|COOP|COOPERAT",
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
    # Igual que en tu script original: trabaja sobre CICLO 49_PDIRECCION.xlsx
    ruta_entrada = "CICLO 49_PDIRECCION.xlsx"
    ruta_salida = "CICLOS_PROCESADOS.xlsx"

    # Leer archivo de entrada
    df_local = pd.read_excel(ruta_entrada, usecols=["CLIENTE_ID", "DIRECCION"])

    # Procesar
    df_resultado = procesar(df_local)

    # Guardar en una sola hoja de salida (puedes cambiarle el nombre si prefieres)
    modo = "a" if os.path.exists(ruta_salida) else "w"
    with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode=modo) as writer:
        df_resultado.to_excel(writer, sheet_name="CDORADA_COOP", index=False)

    # Estadísticas globales
    total = df_resultado.shape[0]
    normalizadas = df_resultado[df_resultado["VALIDACION"] == "1"].shape[0]
    efectividad_global = (normalizadas * 100) / total if total else 0

    print(f"✅ Archivo procesado y guardado como {ruta_salida}")
    print(f"📦 Total direcciones intentadas: {total}")
    print(f"✔️ Direcciones normalizadas correctamente: {normalizadas}")
    print(f"📈 Efectividad global: {round(efectividad_global, 2)}%")
