import pandas as pd
import re
import os

# ============================================================
#    Utilidades comunes
# ============================================================

romanos = {
    "I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6,
    "VII": 7, "VIII": 8, "IX": 9, "X": 10
}


def etapa_a_numero(etapa):
    if etapa is None:
        return None
    etapa = (
        etapa.strip()
        .replace("ETAPA", "")
        .replace("ETAP", "")
        .replace("ETP", "")
        .replace("ET", "")
        .strip()
        .upper()
    )
    return str(romanos.get(etapa, etapa))


# ============================================================
#    LAS COLINAS
# ============================================================

regex_colinas = re.compile(
    r"""
    (?:URB|BRR)?\s*(?:LAS\s+)?COLINAS.*?
    (?P<sector>(?:SC|SEC|SECT)\s*\d{1,2})
    .*?(?:MNZ|MZN|MZ|MANZANA)\s+(?P<manzana>[A-Z0-9]{1,3})
    .*?(?:CS|C|CASA)?\s*(?P<casa>\d{1,4}[A-Z]?)
    (?:\s*(?:PI|PIS|PISO|P)?[\s\-]*(?P<piso>\d{1,2}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_las_colinas(direccion):
    if pd.isna(direccion):
        return direccion, "0"
    direccion = str(direccion).upper().strip()
    match = regex_colinas.search(direccion)
    if match:
        sector_raw = match.group("sector")
        sector = re.sub(r"[^0-9]", "", sector_raw or "")
        manzana = match.group("manzana")
        casa = match.group("casa")
        piso = match.group("piso")
        if sector and manzana and casa:
            nueva = f"URB LAS COLINAS SC {sector} MZ {manzana} CS {casa}"
            if piso:
                nueva += f" PI {piso}"
            return nueva.strip(), "1"
    return direccion, "0"


# ============================================================
#    LA ADIELA
# ============================================================

regex_adiela = re.compile(
    r"""
    (?:URB|BRR)?\s*(?:LA\s+)?ADIELA(?:\s+II)?
    (?:.*?(ETA|ETAPA|ETP|ET)\s*(?P<etapa>[IVXLCDM]+|\d+))?
    .*?(?:MNZ|MZN|MZ)\s*(?P<manzana>\d{1,3})
    .*?(?:CS|C|CASA)?\s*(?P<casa>\d{1,4}[A-Z]?)
    (?:.*?(?:PI|PIS|PISO|P)?[\s\-]*(?P<piso>\d{1,2}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_adiela(direccion):
    if pd.isna(direccion):
        return direccion, "0"
    direccion = str(direccion).upper().strip()
    match = regex_adiela.search(direccion)
    if match:
        manzana = match.group("manzana")
        casa = match.group("casa")
        piso = match.group("piso")
        etapa = etapa_a_numero(match.group("etapa"))
        if manzana and casa:
            nueva = f"BRR LA ADIELA MZ {manzana} CS {casa}"
            if piso:
                nueva += f" PI {piso}"
            if etapa:
                nueva += f" ET {etapa}"
            return nueva.strip(), "1"
    return direccion, "0"


# ============================================================
#    LA ESMERALDA
# ============================================================

regex_esmeralda = re.compile(
    r"""
    (?:URB|BRR)?\s*(?:LA\s+)?ESMERALDA
    .*?(?:MNZ|MZN|MZ)?\s*(?P<manzana>\d{1,3})
    .*?(?:CS|C|CASA)?\s*(?P<casa>\d{1,3}[A-Z]?)
    (?:[\s\-]*(?:PI|PIS|PISO|P)?[\s\-]*(?P<piso>\d{1,2}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_esmeralda(direccion):
    if pd.isna(direccion):
        return direccion, "0"
    direccion = str(direccion).upper().strip()
    # Filtrar cosas que no son viviendas
    if any(x in direccion for x in ["LOC", "ZONA", "NIU", "CAMBUCHE", "LC"]):
        return direccion, "0"
    match = regex_esmeralda.search(direccion)
    if match:
        manzana = match.group("manzana")
        casa = match.group("casa")
        piso = match.group("piso")
        if manzana and casa:
            nueva = f"BRR LA ESMERALDA MZ {manzana} CS {casa}"
            if piso:
                nueva += f" PI {piso}"
            return nueva.strip(), "1"
    return direccion, "0"


# ============================================================
#    7 DE AGOSTO
# ============================================================

regex_agosto = re.compile(
    r"""
    (?:URB|BRR)?\s*7\s+DE\s+AGOSTO
    .*?(?:MNZ|MZN|MZ|MN|MNZS|MMZ)\s*(?P<manzana>\d{1,3})
    .*?(?:CS|C|CASA)?\s*(?P<casa>\d{1,4}[A-Z]?)
    (?:[\s\-]*(?:PI|PIS|PISO|P)?[\s\-]*(?P<piso>\d{1,2}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_agosto(direccion):
    if pd.isna(direccion):
        return direccion, "0"
    direccion = str(direccion).upper().strip()
    match = regex_agosto.search(direccion)
    if match:
        manzana = match.group("manzana")
        casa = match.group("casa")
        piso = match.group("piso")
        if manzana and casa:
            nueva = f"BRR 7 DE AGOSTO MZ {manzana} CS {casa}"
            if piso:
                nueva += f" PI {piso}"
            return nueva.strip(), "1"
    return direccion, "0"


# ============================================================
#    LA UNION
# ============================================================

regex_union = re.compile(
    r"""
    (?:BRR|URB)?\s*(?:LA\s+)?UNION.*?
    (?:MNZ|MZN|MZ|MZA)\s+(?P<manzana>[A-Z0-9]{1,3})
    (?:.*?(?:CS|CASA|C|LT))?\s*(?P<casa>\d{1,4}[A-Z]?)
    (?:.*?(?:PI|PISO|PIS|P)\s*-?(?P<piso>\d{1,2}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_union(direccion):
    if pd.isna(direccion):
        return direccion, "0"
    direccion = str(direccion).upper().strip()

    # Excluir cosas no relevantes
    if any(
        x in direccion
        for x in ["NIU", "LOCAL", "FTE", "ENT", "ENTR", "JUNT", "ALT", "BAJ", "ZONA", "ARA"]
    ):
        return direccion, "0"

    match = regex_union.search(direccion)
    if match:
        manzana = match.group("manzana")
        casa = match.group("casa")
        piso = match.group("piso")
        if manzana and casa:
            nueva = f"BRR LA UNION MZ {manzana} CS {casa}"
            if piso:
                nueva += f" PI {piso}"
            return nueva.strip(), "1"
    return direccion, "0"


# ============================================================
#    LA GRECIA
# ============================================================

regex_lagrecia = re.compile(
    r"""
    (?:URB|BRR)?\s*LA\s+GRECIA
    .*?(?:MNZZ|MNZ|MZN|MZ|MZA|MLZ)\s*(\d{1,3})        # Manzana
    .*?(?:CS|CASA)?\s*(\d{1,4})                       # Casa
    (?:.*?(APT|AP|APTO|APARTAMENTO)\s*(\d{1,4}))?     # Apto
    (?:.*?(PI|PISO|PIS|P)\s*(\d{1,2}))?               # Piso
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_lagrecia(direccion):
    if pd.isna(direccion):
        return direccion, "0"
    direccion_upper = str(direccion).upper().strip()

    if "GRECIA" not in direccion_upper:
        return direccion, "0"

    # Descartar ambigüedades (ej. "Y 12", "12-34")
    if re.search(r"\bY\s+\d{1,4}", direccion_upper) or re.search(r"\d+\s*-\s*\d+", direccion_upper):
        return direccion, "0"

    match = regex_lagrecia.search(direccion_upper)
    if match:
        manzana = match.group(1)
        casa = match.group(2)
        apt_key = match.group(3)
        apto = match.group(4)
        piso_key = match.group(5)
        piso = match.group(6)

        if not (manzana and casa):
            return direccion, "0"

        nueva = f"URB LA GRECIA MZ {manzana} CS {casa}"
        if apt_key and apto:
            nueva += f" AP {apto}"
        if piso_key and piso and piso != "8000":
            nueva += f" PI {piso}"

        return nueva.strip(), "1"

    return direccion, "0"


# ============================================================
#    APLICACIÓN GENERAL DEL MÓDULO
# ============================================================

def aplicar_normalizacion(direccion):
    direccion_upper = direccion.upper() if isinstance(direccion, str) else ""
    if "COLINAS" in direccion_upper:
        return normalizar_las_colinas(direccion)
    elif "ADIELA" in direccion_upper:
        return normalizar_adiela(direccion)
    elif "ESMERALDA" in direccion_upper:
        return normalizar_esmeralda(direccion)
    elif "7 DE AGOSTO" in direccion_upper:
        return normalizar_agosto(direccion)
    elif "UNION" in direccion_upper:
        return normalizar_union(direccion)
    elif "GRECIA" in direccion_upper:
        return normalizar_lagrecia(direccion)
    return direccion, "0"


# ============================================================
#    FUNCIÓN ESTÁNDAR PARA EL SCRIPT MAESTRO
# ============================================================

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe un DataFrame con al menos:
      - NIU o CLIENTE_ID
      - DIRECCION

    Devuelve SOLO filas de:
      - LAS COLINAS
      - LA ADIELA
      - LA ESMERALDA
      - 7 DE AGOSTO
      - LA UNION
      - LA GRECIA

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

    # Filtrar solo estos barrios
    filtro = df["DIRECCION"].str.contains(
        "COLINAS|ADIELA|ESMERALDA|7 DE AGOSTO|UNION|GRECIA",
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
#    EJECUCIÓN INDIVIDUAL OPCIONAL
# ============================================================

if __name__ == "__main__":
    ruta_entrada = "CICLO 47_PDIRECCION.xlsx"
    ruta_salida = "CICLOS_PROCESADOS.xlsx"

    df_local = pd.read_excel(ruta_entrada, usecols=["CLIENTE_ID", "DIRECCION"], engine="openpyxl")
    df_resultado = procesar(df_local)

    modo = "a" if os.path.exists(ruta_salida) else "w"
    with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode=modo) as writer:
        df_resultado.to_excel(writer, sheet_name="NORMALIZADAS", index=False)

    total_direcciones = df_resultado.shape[0]
    normalizadas = df_resultado[df_resultado["VALIDACION"] == "1"].shape[0]
    efec = (normalizadas / total_direcciones) * 100 if total_direcciones > 0 else 0

    print(f"Total direcciones procesadas: {total_direcciones}")
    print(f"Direcciones normalizadas correctamente: {normalizadas}")
    print(f"Efectividad: {efec:.2f}%")
