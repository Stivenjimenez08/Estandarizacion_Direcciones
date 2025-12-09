import pandas as pd
import re
import os

# ============================================================
# Conversión de romanos → número para la etapa
# ============================================================

romanos = {
    "I": "1",
    "II": "2",
    "III": "3",
    "IV": "4",
    "V": "5",
    "VI": "6",
    "VII": "7",
    "VIII": "8",
    "IX": "9",
    "X": "10",
}


def etapa_a_numero(etapa):
    if not etapa:
        return ""
    etapa = (
        etapa.strip()
        .upper()
        .replace("ETAPA", "")
        .replace("ET", "")
        .strip()
    )
    return romanos.get(etapa, etapa)


# ============================================================
# REGEX para cada barrio
# ============================================================

regex_juliana = re.compile(
    r"""(?:URB|BRR)?\s*V(?:D|\.|ILLA)?\s*JULIANA
        (?P<etapa>(?:\s+(?:ET|ETAPA)?\s*\d+|\s+I{1,3}|\s+IV|\s+V|\s+VI|\s+VII|\s+VIII|\s+IX|\s+X))?
        \s+(?:MNZ|MZN|MZ|MNA|M)\s+(?P<manzana>[A-Z0-9]{1,3})
        \s+(?:CASA|CS|C)?\s*(?P<casa>\d{1,3}[A-Z]?)
        (?:\s*(?:PISO|PI)\s*(?P<piso>\d{1,2}))?
        (?:\s*(?:APT|AP)\s*(?P<apt>\d{1,4}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_italia = re.compile(
    r"""(?:URB|BRR)?\s*VILLA\s+ITALIA
        .*?(?:MNZ|MZN|MZ)\s*(?P<manzana>\d{1,3})
        .*?(?:CASA|CS|C)?\s*(?P<casa>\d{1,3}[A-Z]?)
        (?:.*?(?P<marcador>PISO|PI|AP)\s*(?P<piso>\d{1,3}))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_cafe = re.compile(
    r"""(?:URB|BRR)?\s*VILLA\s+DEL\s+CAFE
        .*?(?:MNZ|MZN|MZ|MZA|MHNZ)?\s*(?P<manzana>[A-Z])
        \s*(?:CS|CASA)?\s*(?P<casa>\d+[A-Z]?)
        (?:\s*(?:PISO|PI|P)?\s*(?P<piso>\d+))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_marina = re.compile(
    r'\b'
    r'(?:URB|BRR|CJT|UR)?\s*'
    r'(?:QUINTAS|QUITAS|QUINT|QUNT|QTAS|Q|QUINTA|QUIN|Q\-TAS|QNTS|QTS)?'
    r'(?:\s+(?:DE|D))?(?:\s+(?:LA|ELA))?\s+(?:SAN\s)?M\s*ARIN[A]?'
    r'\s*'
    r'(?:\b(?:ET|ETAPA|ETP)?\s*(\d+)|(\d+|I{1,3}|IV|V|VI{0,3}|IX|X))?\s*'
    r'M?(?:ZN?|NZ|ZA|N)?\s*([A-Z\d]+)\s*'
    r'(?:CS|#|CASA|CAS|C|CA|APT|AP|INT|APU)?\s*(\d+)?\s*'
    r'(?:PISO|PI|P|PIS)?\s*(\d+)?\s*'
    r'(?:ETAPA|ET|TP|E.|ETP)?\s*(\d+)?',
    re.IGNORECASE,
)


# ============================================================
# Función de normalización por dirección
# ============================================================

def normalizar_direccion(direccion):
    if pd.isna(direccion):
        return direccion, "0"

    direccion = str(direccion).upper().strip()

    # ---------- VILLA JULIANA ----------
    if "JULIANA" in direccion:
        # Excluir cosas como CASETA, ZONA, LOCAL, etc.
        if any(x in direccion for x in ["CASETA", "NIU", "ZONA", "LOCAL"]):
            return direccion, "0"

        match = regex_juliana.search(direccion)
        if match:
            manzana = match.group("manzana")
            casa = match.group("casa")
            piso = match.group("piso")
            apt = match.group("apt")
            etapa = etapa_a_numero(match.group("etapa"))

            nueva = f"URB VILLA JULIANA MZ {manzana} CS {casa}"
            if apt:
                nueva += f" AP {apt}"
            if piso:
                nueva += f" PI {piso}"
            if etapa:
                nueva += f" ET {etapa}"
            return nueva.strip(), "1"

    # ---------- VILLA ITALIA ----------
    if "VILLA ITALIA" in direccion:
        match = regex_italia.search(direccion)
        if match:
            manzana = match.group("manzana")
            casa = match.group("casa")
            piso = match.group("piso")
            marcador = match.group("marcador")

            nueva = f"URB VILLA ITALIA MZ {manzana} CS {casa}"
            if piso and marcador:
                nueva += f" {marcador} {piso}"
            return nueva.strip(), "1"

    # ---------- VILLA DEL CAFE ----------
    if "VILLA DEL CAFE" in direccion:
        match = regex_cafe.search(direccion)
        if match:
            manzana = match.group("manzana")
            casa = match.group("casa").lstrip("0")
            piso = match.group("piso")

            nueva = f"URB VILLA DEL CAFE MZ {manzana} CS {casa}"
            if piso:
                nueva += f" PI {piso}"
            return nueva.strip(), "1"

    # ---------- QUINTAS DE LA MARINA ----------
    if any(x in direccion for x in ["MARINA", "MARIN", "M ARINA"]):
        match = regex_marina.search(direccion)
        if match:
            etapa = match.group(1) or match.group(2) or ""
            manzana = match.group(3)
            casa = match.group(4) or ""
            piso = match.group(5) or ""
            etapa = etapa_a_numero(etapa)

            nueva = f"URB QUINTAS DE LA MARINA MZ {manzana} CS {casa}"
            if piso:
                nueva += f" PI {piso}"
            if etapa:
                nueva += f" ET {etapa}"
            return nueva.strip(), "1"

    # Si no se pudo normalizar
    return direccion, "0"


# ============================================================
#   FUNCIÓN ESTÁNDAR PARA USAR DESDE EL SCRIPT MAESTRO
# ============================================================

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe un DataFrame con al menos:
      - NIU o CLIENTE_ID
      - DIRECCION

    Devuelve SOLO filas de:
      - VILLA JULIANA
      - VILLA ITALIA
      - VILLA DEL CAFE
      - QUINTAS DE LA MARINA

    con columnas:
      NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION
    """
    df = df_in.copy()

    # Normalizar nombres de columnas a un mapa en mayúsculas
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

    # Filtrar direcciones de interés (igual que el script original)
    filtro = df["DIRECCION"].str.contains(
        "JULIANA|VILLA ITALIA|VILLA DEL CAFE|MARINA|MARIN|M ARINA",
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
#   BLOQUE OPCIONAL PARA EJECUTAR ESTE SCRIPT SOLO
# ============================================================

if __name__ == "__main__":
    # Configuración de rutas para ejecución individual
    ruta_entrada = "CICLO 49_PDIRECCION.xlsx"
    ruta_salida = "CICLOS_PROCESADOS.xlsx"

    # Leer datos como en el script original
    df_local = pd.read_excel(ruta_entrada, usecols=["CLIENTE_ID", "DIRECCION"])

    # Procesar usando la función estándar
    df_resultado = procesar(df_local)

    # Guardar resultado
    modo = "a" if os.path.exists(ruta_salida) else "w"
    with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode=modo) as writer:
        df_resultado.to_excel(writer, sheet_name="UNIFICADO_VILLAS_MARINA", index=False)

    # Estadísticas
    total = df_resultado.shape[0]
    normalizadas = df_resultado[df_resultado["VALIDACION"] == "1"].shape[0]
    efectividad = (normalizadas / total) * 100 if total > 0 else 0

    print(f"✅ Archivo procesado y guardado como {ruta_salida}")
    print(f"🔍 Total direcciones filtradas: {total}")
    print(f"📊 Direcciones normalizadas: {normalizadas}")
    print(f"📈 Efectividad: {round(efectividad, 2)}%")
