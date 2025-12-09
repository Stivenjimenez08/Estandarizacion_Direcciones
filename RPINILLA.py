import pandas as pd
import re
import os

# ================== REGEX / NORMALIZACIÓN ==================
ORIENT_MAP = {
    "NORTE": "N", "SUR": "S", "ESTE": "E", "OESTE": "O", "ORIENTE": "E", "OCCIDENTE": "O",
    "NTE": "N", "STE": "S", "OTE": "E", "OCC": "O", "N": "N", "S": "S", "E": "E", "O": "O"
}

ROTULOS_A_REMOVER = [
    "CLUB HOUSE", "CLUBHOUSE", "CENTENARIO MALL", "MALL CENTENARIO", "FLORIDA BAJA",
    "BALEARES", "AV BOLIVAR", "ED EL PILAR", "AMANECER", "MALL ZN ORO", "LUXOR"
]


def limpiar_rotulos_finales(t: str) -> str:
    s = t
    for r in ROTULOS_A_REMOVER:
        s = re.sub(rf"{re.escape(r)}\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalizar_orientacion(txt: str) -> str:
    if not txt:
        return ""
    t = txt.strip().upper()
    return ORIENT_MAP.get(t, t)


def extraer_segmento_lc(d: str) -> str:
    m_lc = re.search(
        r"\bLC\s+(\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+)",
        d,
        flags=re.IGNORECASE,
    )
    if m_lc:
        return f"LC {m_lc.group(1).strip()}"
    return ""


def normalizar_cola_lc(seg: str) -> str:
    seg = seg.upper().replace("LOCAL", "LC").replace("LOCALES", "LC")
    seg = re.sub(r"\s+", " ", seg)
    seg = seg.replace("LC LC", "LC").strip()
    return seg


def normalize_tipo_num(tipo: str, num: str) -> str:
    """
    Normaliza combinaciones como BLOQUE 5, BQ 5, BLQ 5, etc.
    """
    if not tipo or not num:
        return ""
    t = tipo.strip().upper()
    n = num.strip().upper()

    if t in ["BL", "BQ", "BLQ", "BLOQUE", "TORRE", "T"]:
        if t in ["BL", "BQ", "BLQ", "BLOQUE"]:
            pref = "BQ"
        else:
            pref = "T"
        return f"{pref} {n}"
    if t in ["TO", "TORRE"]:
        return f"T {n}"
    return f"{t} {n}"


# --- Regex para CLL/CR específicos (CR/CL, etc.) ---
regex_cll_cr = re.compile(
    r"""
    ^\s*CLL\s*(?P<cll>\d+(?:[A-Z]{1,2})?)\s*
    (?:NORTE|SUR|ESTE|OESTE|ORIENTE|OCCIDENTE|N|S|E|O|NTE|STE|OTE|OCC)?
    \s*(?:CR|CRA|KR|KRA|K)\s*(?P<cr>\d+(?:[A-Z]{1,2})?)\s*
    (?:
        (?:-|–|—)\s*(?P<guion>\d+)?
    )?
    (?:\s*(?:APTO?|APARTAMENTO|AP)\s*(?P<ap>\d+))?
    (?:\s*(?:PI|PISO|PIS)\s*(?P<piso>\d+))?
    (?:\s*LC\s+(?P<lc_raw>(?:\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+)))?
    (?:\b.*)?$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# --- Regex CLL/CR con texto intermedio (ej. "CLL 26 CR 15 -57 LC 1") ---
regex_cll_cr_intermedio = re.compile(
    r"""
    ^\s*CLL\s*(?P<cll>\d+(?:[A-Z]{1,2})?)\s*
    (?:NORTE|SUR|ESTE|OESTE|ORIENTE|OCCIDENTE|N|S|E|O|NTE|STE|OTE|OCC)?
    .*?
    (?:CR|CRA|KR|KRA|K)\s*(?P<cr>\d+(?:[A-Z]{1,2})?)\s*
    (?:
        (?:-|–|—)\s*(?P<guion>\d+)?
    )?
    (?:\s*(?:APTO?|APARTAMENTO|AP)\s*(?P<ap>\d+))?
    (?:\s*(?:PI|PISO|PIS)\s*(?P<piso>\d+))?
    (?:\s*LC\s+(?P<lc_raw>(?:\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+)))?
    (?:\b.*)?$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# --- Regex CRA/CL ---
regex_cra_cl = re.compile(
    r"""
    ^\s*CRA\s*(?P<cra>\d+(?:[A-Z]{1,2})?)\s*
    (?:NORTE|SUR|ESTE|OESTE|ORIENTE|OCCIDENTE|N|S|E|O|NTE|STE|OTE|OCC)?
    \s*CL\s*(?P<cl>\d+(?:[A-Z]{1,2})?)\s*
    (?:
        (?:-|–|—)\s*(?P<guion>\d+)?
    )?
    (?:\s*(?:APTO?|APARTAMENTO|AP)\s*(?P<ap>\d+))?
    (?:\s*(?:PI|PISO|PIS)\s*(?P<piso>\d+))?
    (?:\s*LC\s+(?P<lc_raw>(?:\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+)))?
    (?:\b.*)?$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# --- CRA/CL con texto intermedio ---
regex_cra_cl_intermedio = re.compile(
    r"""
    ^\s*CRA\s*(?P<cra>\d+(?:[A-Z]{1,2})?)\s*
    (?:NORTE|SUR|ESTE|OESTE|ORIENTE|OCCIDENTE|N|S|E|O|NTE|STE|OTE|OCC)?
    .*?
    CL\s*(?P<cl>\d+(?:[A-Z]{1,2})?)\s*
    (?:
        (?:-|–|—)\s*(?P<guion>\d+)?
    )?
    (?:\s*(?:APTO?|APARTAMENTO|AP)\s*(?P<ap>\d+))?
    (?:\s*(?:PI|PISO|PIS)\s*(?P<piso>\d+))?
    (?:\s*LC\s+(?P<lc_raw>(?:\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+)))?
    (?:\b.*)?$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# --- CRA/CL general (permite letras en CRA y cola) ---
regex_cra_cl_general = re.compile(
    r"""CRA\s*(?P<cra>\d+(?:[A-Z]{1,2})?)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        (?:(?:-|–|—)\s*|\s+)?(?P<guion>0*\d+)? 
        (?:\s*(?:APTO?|APARTAMENTO|AP|APU|APT)\s*(?P<ap>\d+))?
        (?:\s*(?:PI|PISO|PIS)\s*(?P<piso>\d+))?
        (?:\s*(?:OFI(?:CINA)?|OF(?:ICINA)?)\s*(?P<of>\d+(?:\s*-\s*\d+)?))?
        (?:\s*LC\s+(?P<lc>(?:\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+)))?
        (?:
            .*?(?P<tipo>(?:\bTO\b|\bTORRE\b|\bT\b|\bBL\b|\bBQ\b|\bBLQ\b|\bBLOQUE\b))\s*
            (?P<numtipo>(?!AP\b)[A-Z0-9]+)\s*
            (?:\s*(?:APTO?|APARTAMENTO|AP)\s*(?P<ap2>\d+))?
            (?:\s*(?:PI|PISO|PIS)\s*(?P<piso2>\d+))?
        )?
        (?:\b.*)?$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# --- CLL/CR general (permite cola) ---
regex_cll_cr_general = re.compile(
    r"""CLL\s*(?P<cll>\d+(?:[A-Z]{1,2})?(?:\s*BIS(?:\s*[A-Z])?)?)
        (?:\s*(?:NORTE|SUR|ESTE|OESTE|ORIENTE|OCCIDENTE|N|S|E|O|NTE|STE|OTE|OCC))?
        \s*(?:CR|CRA|KR|KRA|K|CL)\s*(?P<cr>\d+(?:[A-Z]{1,2})?(?:\s*BIS(?:\s*[A-Z])?)?)
        (?:\s*(?:NORTE|SUR|ESTE|OESTE|ORIENTE|OCCIDENTE|N|S|E|O|NTE|STE|OTE|OCC))?
        (?:\s*(?P<orient1>NORTE|SUR|ESTE|OESTE|ORIENTE|OCCIDENTE|N|S|E|O|NTE|STE|OTE|OCC))?
        \s*(?:
            (?:-|–|—)\s*(?P<guion>0*\d+)?
        )?
        (?:\s*(?:APTO?|APARTAMENTO|AP|APT)\s*(?P<ap>\d+))?
        (?:\s*(?:PI|PISO|PIS)\s*(?P<piso>\d+))?
        (?:\s*(?:OFI(?:CINA)?|OF(?:ICINA)?)\s*(?P<of>\d+(?:\s*-\s*\d+)?))?
        (?:\s*LC\s+(?P<lc_raw>(?:\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+)))?
        (?:\b.*)?$
    """,
    re.IGNORECASE | re.VERBOSE,
)


def normalizar_direccion(direccion: str):
    if not isinstance(direccion, str):
        return "", "0"

    d = direccion.upper().strip()
    d = re.sub(r"\s+", " ", d)

    # --- Rojas Pinilla ---
    if "ROJAS PINILLA" in d:
        d = d.replace("KR ", "CRA ")
        d = d.replace("KR", "CRA ")
        d = re.sub(r"\s+", " ", d)
        if "CRA" not in d:
            d = "CRA " + d
        # No se fuerza CL porque en Rojas Pinilla la estructura es CRA XX (sin CL)
        d = limpiar_rotulos_finales(d)
        return d, "1"

    # --- CLL/CR directos ---
    m = regex_cll_cr.search(d)
    if not m:
        m = regex_cll_cr_intermedio.search(d)
    if m:
        cll = m.group("cll")
        cr = m.group("cr")
        guion = m.group("guion")
        ap = m.group("ap")
        piso = m.group("piso")
        lc_raw = m.group("lc_raw")

        out = f"CLL {cll} CR {cr}"
        if guion:
            out += f" - {int(guion)}"
        if ap:
            out += f" AP {ap}"
        if piso:
            out += f" PI {int(piso)}"
        if lc_raw:
            out += f" {normalizar_cola_lc('LC ' + lc_raw.strip())}"

        out = limpiar_rotulos_finales(out)
        return out, "1"

    # --- CRA/CL directos ---
    m = regex_cra_cl.search(d)
    if not m:
        m = regex_cra_cl_intermedio.search(d)
    if m:
        cra = m.group("cra")
        cl = m.group("cl")
        guion = m.group("guion")
        ap = m.group("ap")
        piso = m.group("piso")
        lc_raw = m.group("lc_raw")

        out = f"CRA {cra} CL {cl}"
        if guion:
            out += f" - {int(guion)}"
        if ap:
            out += f" AP {ap}"
        if piso:
            out += f" PI {int(piso)}"
        if lc_raw:
            out += f" {normalizar_cola_lc('LC ' + lc_raw.strip())}"

        out = limpiar_rotulos_finales(out)
        return out, "1"

    # --- CLL/CR general ---
    m = regex_cll_cr_general.search(d)
    if m:
        cll = m.group("cll")
        cr = m.group("cr")
        guion = m.group("guion")
        ap = m.group("ap")
        piso = m.group("piso")
        of = m.group("of")
        lc_raw = m.group("lc_raw")

        out = f"CLL {cll} CR {cr}"
        if guion:
            out += f" - {int(guion)}"
        if ap:
            out += f" AP {ap}"
        if piso:
            out += f" PI {int(piso)}"
        if of:
            out += f" OF {of}"
        if lc_raw:
            out += f" {normalizar_cola_lc('LC ' + lc_raw.strip())}"

        out = limpiar_rotulos_finales(out)
        return out, "1"

    # --- CRA/CL general ---
    m = regex_cra_cl_general.search(d)
    if m:
        cra = m.group("cra")
        cl = m.group("cl")
        guion = m.group("guion")
        ap = m.group("ap") or m.group("ap2")
        piso = m.group("piso") or m.group("piso2")
        of = m.group("of")
        lc = m.group("lc")
        tipo = m.group("tipo")
        numtipo = m.group("numtipo")

        out = f"CRA {cra} CL {cl}"
        if guion:
            out += f" - {int(guion)}"
        if ap:
            out += f" AP {ap}"
        if piso:
            out += f" PI {int(piso)}"
        if of:
            out += f" OF {of}"
        if tipo and numtipo:
            out += f" {normalize_tipo_num(tipo, numtipo)}"
        if lc:
            out += f" {normalizar_cola_lc('LC ' + lc.strip())}"

        out = limpiar_rotulos_finales(out)
        return out, "1"

    return d, "0"


def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """Procesa un DataFrame con columnas NIU/CLIENTE_ID y DIRECCION.

    Devuelve solo las filas filtradas (Rojas Pinilla / CLL/CR / CRA/CL)
    con las columnas: NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION.
    """
    df = df_in.copy()

    # Normalizar nombre de columnas a mayúsculas temporales
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

    # Aplicar la función de normalización original
    df[["DIRECCION_NORMALIZADA", "VALIDACION"]] = df["DIRECCION"].apply(
        lambda x: pd.Series(normalizar_direccion(x))
    )

    # Mismo filtro que en el script original
    filtro = df["DIRECCION"].str.contains(
        r"ROJAS\s+PINILLA|^\s*CLL\b|^\s*CRA\b|^\s*CR\b",
        case=False,
        na=False,
    )
    df_filtrado = df[filtro].copy()

    columnas_salida = ["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"]
    for col in columnas_salida:
        if col not in df_filtrado.columns:
            df_filtrado[col] = None

    return df_filtrado[columnas_salida]


if __name__ == "__main__":
    # Comportamiento opcional para probar este módulo de forma independiente.
    ruta_entrada = "CICLO 47_PDIRECCION.xlsx"
    df_local = pd.read_excel(ruta_entrada)

    df_resultado = procesar(df_local)

    ruta_salida = "CICLOS_PROCESADOS.xlsx"
    hoja_salida = "C47_ROJAS_PINILLA"

    modo = "a" if os.path.exists(ruta_salida) else "w"
    with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode=modo) as writer:
        df_resultado.to_excel(writer, sheet_name=hoja_salida, index=False)

    total = df_resultado.shape[0]
    normalizadas = (df_resultado["VALIDACION"] == "1").sum()
    efectividad = (normalizadas / total * 100) if total > 0 else 0

    print(f"✅ Archivo procesado y guardado como {ruta_salida}")
    print(f"📄 Hoja: {hoja_salida}")
    print(f"🔍 Total direcciones filtradas: {total}")
    print(f"📊 Direcciones normalizadas: {normalizadas}")
    print(f"📈 Efectividad: {round(efectividad, 2)}%")