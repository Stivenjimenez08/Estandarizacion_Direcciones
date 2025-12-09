import pandas as pd
import re
import os

# ============================================================
#  Regex y funciones de normalización para LA MIRANDA / ACACIAS
# ============================================================

# Detectar locales tipo "LC 12", "LOCAL 5A", etc.
regex_local = re.compile(
    r"\b(LC|LOC|LOCAL|LO)\s*(\d{1,4}[A-Z]?)\b", re.IGNORECASE
)

# Barrios (con muchas variantes de escritura)
PATRON_MIRANDA = r"(LA\s+MIRANDA|MIRANDA)"
PATRON_ACACIAS = r"((LAS\s+)?ACACIAS|ACACIA)"


def limpiar_espacios(txt: str) -> str:
    txt = re.sub(r"\s+", " ", txt or "").strip()
    return txt


def normalizar_barrio_y_manzana(
    direccion: str,
    patron_barrio: str,
    nombre_barrio_normalizado: str
):
    """
    Busca estructuras tipo:

      BRR LA MIRANDA MZ C CS 12
      URB ACACIAS MANZANA B CASA 4
      CND LAS ACACIAS MZN A  CS 3
      BRR LA MIRANDA MZ C 12   (sin 'CS' explícito)
      etc.

    y devuelve algo como:
      "BARRIO LA MIRANDA MZ C CS 12"
    """

    if not isinstance(direccion, str):
        return "", "0"

    d = direccion.upper()
    d = limpiar_espacios(d)

    # 1) Encontrar barrio
    m_barrio = re.search(patron_barrio, d, flags=re.IGNORECASE)
    if not m_barrio:
        return d, "0"

    # 2) Extraer posibles manzana/casa/apto/piso
    #    Permitimos varios prefijos antes del barrio: BRR, URB, CND, CONJ, etc.
    patron_completo = re.compile(
        rf"""
        (?:
            BRR|BARRIO|URB(?:ANIZACION)?|CND|CONJ|CONJUNTO
        )?
        \s*
        {patron_barrio}
        (?P<resto>.*)
        """,
        re.IGNORECASE | re.VERBOSE
    )

    m = patron_completo.search(d)
    if not m:
        # Si por alguna razón no cuadra, devolvemos solo el barrio normalizado
        return f"BARRIO {nombre_barrio_normalizado}", "0"

    resto = m.group("resto") or ""
    resto = limpiar_espacios(resto)

    # Manzana
    m_mz = re.search(
        r"\b(MNZ|MNZ\.|MZN|MZNA|MZ\.?|MANZANA|MZA|MZNA)\s*([A-ZÑ]{1,2}|\d{1,3})\b",
        resto,
        flags=re.IGNORECASE
    )
    # Casa
    m_cs = re.search(
        r"\b(CS|CASA|CAS|C)\s*([A-Z]?\d{1,4}[A-Z]?)\b",
        resto,
        flags=re.IGNORECASE
    )
    # Apto
    m_ap = re.search(
        r"\b(APTO?|APARTAMENTO|AP)\s*(\d{1,4}[A-Z]?)\b",
        resto,
        flags=re.IGNORECASE
    )
    # Piso
    m_pi = re.search(
        r"\b(PI|PISO|PIS)\s*(\d{1,2})\b",
        resto,
        flags=re.IGNORECASE
    )
    # Local
    m_lc = regex_local.search(resto)

    partes = [f"BARRIO {nombre_barrio_normalizado}"]

    if m_mz:
        mz_letra = m_mz.group(2).upper()
        partes.append(f"MZ {mz_letra}")

    if m_cs:
        cs_num = m_cs.group(2).upper()
        partes.append(f"CS {cs_num}")
    else:
        # Si no hay 'CASA' explícito pero sí un número suelto,
        # lo asumimos como casa.
        m_num_suelto = re.search(r"\b(\d{1,4}[A-Z]?)\b", resto)
        if m_num_suelto and not m_lc:
            partes.append(f"CS {m_num_suelto.group(1).upper()}")

    if m_ap:
        ap = m_ap.group(2).upper()
        partes.append(f"AP {ap}")

    if m_pi:
        pi = m_pi.group(2)
        partes.append(f"PI {int(pi)}")

    if m_lc:
        lc = m_lc.group(2).upper()
        partes.append(f"LC {lc}")

    direccion_normalizada = limpiar_espacios(" ".join(partes))

    # Consideramos "1" si al menos barrio + algún dato numérico
    tiene_numero = re.search(r"\d", direccion_normalizada) is not None
    validacion = "1" if tiene_numero else "0"

    return direccion_normalizada, validacion


def normalizar_direccion(direccion: str):
    """
    Lógica de normalización para:
      - LA MIRANDA
      - ACACIAS

    Si no matchea ningún patrón de estos barrios,
    devuelve la dirección en mayúsculas con VALIDACION = "0".
    """

    if not isinstance(direccion, str):
        return "", "0"

    d = limpiar_espacios(direccion.upper())

    # Primero intentamos LA MIRANDA
    if re.search(PATRON_MIRANDA, d, flags=re.IGNORECASE):
        out, val = normalizar_barrio_y_manzana(
            d,
            PATRON_MIRANDA,
            "LA MIRANDA"
        )
        return out, val

    # Luego intentamos ACACIAS
    if re.search(PATRON_ACACIAS, d, flags=re.IGNORECASE):
        out, val = normalizar_barrio_y_manzana(
            d,
            PATRON_ACACIAS,
            "ACACIAS"
        )
        return out, val

    # Si no es ninguno de los dos barrios, no normalizamos
    return d, "0"


# ============================================================
#      FUNCIÓN ESTÁNDAR PARA USAR DESDE EL SCRIPT MAESTRO
# ============================================================

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe un DataFrame con al menos:
      - NIU o CLIENTE_ID
      - DIRECCION

    Devuelve SOLO filas de los barrios LA MIRANDA y ACACIAS,
    con columnas:
      NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION
    """

    df = df_in.copy()

    # Normalizar nombres de columnas a un mapa mayúscula
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
        lambda x: pd.Series(normalizar_direccion(x))
    )

    # Filtros por barrio (igual idea que el script original)
    filtro_miranda = df["DIRECCION"].str.contains("MIRANDA", case=False, na=False)
    filtro_acacias = df["DIRECCION"].str.contains("ACACIAS", case=False, na=False)

    df_filtrado = df[filtro_miranda | filtro_acacias].copy()

    columnas_salida = ["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"]
    for col in columnas_salida:
        if col not in df_filtrado.columns:
            df_filtrado[col] = None

    return df_filtrado[columnas_salida]


# ============================================================
#      BLOQUE OPCIONAL PARA EJECUTAR ESTE SCRIPT SOLO
# ============================================================

if __name__ == "__main__":
    # Rutas de prueba (como en tu script actual)
    ruta_entrada = "CICLO 49_PDIRECCION.xlsx"
    ruta_salida = "CICLOS_PROCESADOS.xlsx"

    df_local = pd.read_excel(ruta_entrada, usecols=["CLIENTE_ID", "DIRECCION"])
    df_resultado = procesar(df_local)

    # Separar en hojas LA_MIRANDA y ACACIAS como antes
    filtro_miranda = df_resultado["DIRECCION"].str.contains("MIRANDA", case=False, na=False)
    filtro_acacias = df_resultado["DIRECCION"].str.contains("ACACIAS", case=False, na=False)

    modo = "a" if os.path.exists(ruta_salida) else "w"
    with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode=modo) as writer:
        df_resultado[filtro_miranda].to_excel(writer, sheet_name="LA_MIRANDA", index=False)
        df_resultado[filtro_acacias].to_excel(writer, sheet_name="ACACIAS", index=False)

    # Estadísticas simples
    total_acacias = df_resultado[filtro_acacias].shape[0]
    normalizadas_acacias = df_resultado[filtro_acacias][
        df_resultado["DIRECCION_NORMALIZADA"].notna()
        & (df_resultado["VALIDACION"] == "1")
    ].shape[0]
    efectividad_acacias = (normalizadas_acacias / total_acacias * 100) if total_acacias > 0 else 0

    print(f"Total ACACIAS procesadas: {total_acacias}")
    print(f"Normalizadas: {normalizadas_acacias}")
    print(f"Efectividad: {efectividad_acacias:.2f}%")
