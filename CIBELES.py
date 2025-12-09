import pandas as pd
import re
import os

# ---------------------- CIBELES ----------------------
regex_cibeles = re.compile(
    r"""
    (?:CIBELES.*)?
    (?:TO(?:RRE)?|T)\s*(\d{1,2})
    .*?(?:APT|AP|APU)\s*([A-Z])\s*(\d{1,3})
    |
    (?:CIBELES.*)?
    (?:TO(?:RRE)?|T)\s*(\d{1,2})
    .*?(?:APT|AP|APU)\s*(\d{1,3})([A-Z])
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_cibeles(direccion):
    if pd.isna(direccion):
        return direccion, "0"

    original = str(direccion).upper().strip()
    direccion = original

    # Si viene ya con la estructura de CRA 40 CL 51 -41 al inicio
    if original.startswith("CRA 40 CL 51 -41"):
        direccion = re.sub(r"^CRA\s*40\s*CL\s*51\s*-?41\s*", "", original)
        direccion += " CIBELES"

    match = regex_cibeles.search(direccion)

    if match:
        torre = match.group(1) or match.group(4)
        letra = match.group(2) or match.group(6)
        numero = match.group(3) or match.group(5)
        if torre and letra and numero:
            return f"CRA 40 CL 51 -41 TO {torre} AP {numero}{letra}", "1"

    return original, "0"


# ---------------------- VILLA LILIANA ----------------------

regex_liliana_patron1 = re.compile(
    r"""
    CRA\s+(\d+[A-Z]?)\s+(?:CL\s*)?(\d+)[\s\-]+(\d+)
    .*?(?:PI|PISO|PIS)?\s*(\d{1,2})?
    .*?LILIANA
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_quintas_liliana = re.compile(
    r"""
    CRA\s*(\d+[A-Z]?)\s+(\d+)\s+(\d+)
    .*?(?:Q(?:UINTAS|TAS|TAS\.?|TA))
    .*?LILIANA
    .*?(\d{1,4}[A-Z]?)
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_quintas_directa = re.compile(
    r"^URB QUINTAS DE VILLA LILIANA CS \d+[A-Z]?$",
    re.IGNORECASE,
)

regex_liliana_manzana = re.compile(
    r"""
    (?:
        URB|BRR|URBANIZACION|\s*
    )?
    \s*(?:V(?:\.|D\.?|D)?\s?)?LILIANA
    .*?
    (?:MNZ|MZN|MZ|MZ\.?)\s*([A-ZÑ]{1,3})           
    .*?(?:CS|CASA|C)?\s*(\d+[A-Z]?)                 
    (?:.*?(?:PI|PISO|PIS|P)?\s*(\d{1,2}))?          
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_bosques_liliana = re.compile(
    r"""
    URB(?:\.|\s+)?                     
    BOSQUES\s+                         
    (?:DE\s+)?                         
    V(?:\.|ILLA|ILA)?\s*LILIANA        
    (?:\s+CS)?                         
    \s+(\d{1,4}[A-Z]?)\b              
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_liliana(direccion):
    if pd.isna(direccion):
        return direccion, "0"

    direccion = str(direccion).upper().strip()

    # URB QUINTAS DE VILLA LILIANA con CRA/CL
    match1 = regex_quintas_liliana.search(direccion)
    if match1:
        cra = match1.group(1)
        calle = match1.group(2)
        numero = match1.group(3)
        casa = match1.group(4)
        nueva = f"CRA {cra} CL {calle}-{numero} URB QUINTAS DE VILLA LILIANA CS {casa}"
        return nueva.strip(), "1"

    # CRA XX CL YY-ZZ ... LILIANA
    match2 = regex_liliana_patron1.search(direccion)
    if match2:
        cra = match2.group(1)
        calle = match2.group(2)
        numero = match2.group(3)
        piso = match2.group(4)
        nueva = f"CRA {cra} CL {calle}-{numero}"
        if piso:
            nueva += f" PI {piso}"
        nueva += " URB VILLA LILIANA"
        return nueva.strip(), "1"

    # Ya normalizada como QUINTAS
    if regex_quintas_directa.match(direccion):
        return direccion, "1"

    # URB VILLA LILIANA MZ X CS Y (opcional PI)
    match3 = regex_liliana_manzana.search(direccion)
    if match3:
        manzana = match3.group(1)
        casa = match3.group(2)
        piso = match3.group(3)
        nueva = f"URB VILLA LILIANA MZ {manzana} CS {casa}"
        if piso:
            nueva += f" PI {piso}"
        return nueva.strip(), "1"

    # URB BOSQUES DE VILLA LILIANA CS X
    match_bosques = regex_bosques_liliana.search(direccion)
    if match_bosques:
        cs = match_bosques.group(1)
        nueva = f"URB BOSQUES DE VILLA LILIANA CS {cs}"
        return nueva.strip(), "1"

    return direccion, "0"


# ---------------------- Aplicación general ----------------------

def aplicar_normalizacion(direccion):
    direccion = str(direccion).upper()
    if "CIBELES" in direccion or "CRA 40 CL 51 -41" in direccion:
        return normalizar_cibeles(direccion)
    elif "LILIANA" in direccion:
        return normalizar_liliana(direccion)
    return direccion, "0"


# ---------------------- Función estándar para el maestro ----------------------

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe un DataFrame con al menos:
      - NIU o CLIENTE_ID
      - DIRECCION

    Devuelve SOLO filas de:
      - CIBELES (CRA 40 CL 51 -41, torres / aptos)
      - URB VILLA LILIANA / QUINTAS DE VILLA LILIANA
      - URB BOSQUES DE VILLA LILIANA

    con columnas:
      NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION
    """
    df = df_in.copy()
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

    # Mismo filtro que el script original
    filtro = df["DIRECCION"].str.contains(
        "CIBELES|CRA 40 CL 51 -41|LILIANA|BOSQUES",
        case=False,
        na=False,
    )
    df_filtrado = df[filtro].copy()

    columnas_salida = ["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"]
    for col in columnas_salida:
        if col not in df_filtrado.columns:
            df_filtrado[col] = None

    return df_filtrado[columnas_salida]


# ---------------------- Ejecución individual opcional ----------------------

if __name__ == "__main__":
    ruta_entrada = "CICLO 47_PDIRECCION.xlsx"
    ruta_salida = "CICLOS_PROCESADOS.xlsx"

    # Igual que antes: CLIENTE_ID → NIU
    df_local = pd.read_excel(ruta_entrada, usecols=["CLIENTE_ID", "DIRECCION"])
    df_resultado = procesar(df_local)

    modo = "a" if os.path.exists(ruta_salida) else "w"
    with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode=modo) as writer:
        df_resultado.to_excel(writer, sheet_name="CIBELES_LILIANA", index=False)

    total = df_resultado.shape[0]
    normalizadas = df_resultado[df_resultado["VALIDACION"] == "1"].shape[0]
    efectividad = (normalizadas / total * 100) if total > 0 else 0

    print(f"Total direcciones procesadas: {total}")
    print(f"Direcciones normalizadas correctamente: {normalizadas}")
    print(f"Efectividad: {efectividad:.2f}%")
