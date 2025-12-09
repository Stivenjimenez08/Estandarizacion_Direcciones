import pandas as pd
import re
import os

# ================== Configuración por defecto (modo script) ==================
RUTA_ENTRADA = "CICLO 49_PDIRECCION.xlsx"
RUTA_SALIDA  = "CICLOS_PROCESADOS.xlsx"
HOJA_SALIDA  = "25MAYO"
COL_DIR      = "DIRECCION"

# ================== Patrones 25 DE MAYO / CIUDADELA EL SOL / EL PLACER ==================

# 25 DE MAYO: con manzana, casa, opcional apto y piso
REGEX_25_MAYO = re.compile(
    r"""
    25\s+DE\s+MAYO
    (?:\s+(?:URB|BRR))?
    .*?(?:MNZ|MZ|MZA|MANZANA|M)\s*(?P<mz>[A-Z0-9]+)
    .*?(?:CS|CASA|C)\s*(?P<cs>\d+[A-Z]?)
    (?:.*?(?:APTO?|APARTAMENTO|APT|AP)\s*(?P<ap>\d+))?
    (?:.*?(?:PI|PISO|P)\s*(?P<piso>\d+))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

# Variante que usaba APT después del número de casa, por ejemplo:
# "25 DE MAYO MZ A 12 APT 302"
REGEX_25_MAYO_APT = re.compile(
    r"""
    25\s+DE\s+MAYO
    (?:\s+(?:MNZ|MZ|MZA|MANZANA|M))?\s*(?P<mz>[A-Z0-9]+)
    \s+(?P<cs>\d+[A-Z]?)
    \s+APT\s+(?P<ap>\d+)
    """,
    re.IGNORECASE | re.VERBOSE,
)

# CIUDADELA EL SOL
REGEX_CIUD_SOL = re.compile(
    r"""
    CIUDADELA\s+EL\s+SOL
    .*?(?:MNZ|MZ|MZA|MANZANA|M)\s*(?P<mz>[A-Z0-9]+)
    .*?(?:CS|CASA|C)\s*(?P<cs>\d+[A-Z]?)
    (?:.*?(?:PI|PISO|P)\s*(?P<piso>\d+))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

# EL PLACER (a veces sólo "PLACER")
REGEX_EL_PLACER = re.compile(
    r"""
    (?:URB\s+)?EL\s+PLACER|BRR\s+EL\s+PLACER|\bPLACER\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

REGEX_EL_PLACER_MZ_CS = re.compile(
    r"""
    (?:URB\s+)?EL\s+PLACER
    .*?(?:MNZ|MZ|MZA|MANZANA|M)\s*(?P<mz>[A-Z0-9]+)
    .*?(?:CS|CASA|C)\s*(?P<cs>\d+[A-Z]?)
    (?:.*?(?:PI|PISO|P)\s*(?P<piso>\d+))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

# ================== Normalizadores por barrio ==================

def _limpiar_base(d: str) -> str:
    """Limpieza suave común."""
    s = d.upper().strip()
    s = re.sub(r"[\u00A0\u2000-\u200B\u202F\u205F\u3000]", " ", s)  # espacios raros
    s = re.sub(r"[‐‒–—―-]", "-", s)                                 # guiones raros
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"-?\s*NIU\s*\#?\s*\d+\b", "", s)                    # NIU suelto
    s = re.sub(r"\s+ARMENIA\b", "", s)                             # sufijo ciudad
    return s.strip()

def normalizar_25_de_mayo(direccion: str):
    if pd.isna(direccion):
        return direccion, "0"
    txt = _limpiar_base(str(direccion))

    # Variante APT explícito
    m2 = REGEX_25_MAYO_APT.search(txt)
    if m2:
        mz, cs, ap = m2.group("mz"), m2.group("cs"), m2.group("ap")
        out = f"URB 25 DE MAYO MZ {mz} CS {cs} AP {int(ap)}"
        return out, "1"

    m = REGEX_25_MAYO.search(txt)
    if not m:
        return direccion, "0"

    mz  = m.group("mz")
    cs  = m.group("cs")
    ap  = m.group("ap")
    piso = m.group("piso")

    out = f"URB 25 DE MAYO MZ {mz} CS {cs}"
    if ap:
        out += f" AP {int(ap)}"
    if piso:
        out += f" PI {int(piso)}"
    return out, "1"

def normalizar_ciudadela_el_sol(direccion: str):
    if pd.isna(direccion):
        return direccion, "0"
    txt = _limpiar_base(str(direccion))
    m = REGEX_CIUD_SOL.search(txt)
    if not m:
        return direccion, "0"
    mz   = m.group("mz")
    cs   = m.group("cs")
    piso = m.group("piso")
    out = f"URB CIUDADELA EL SOL MZ {mz} CS {cs}"
    if piso:
        out += f" PI {int(piso)}"
    return out, "1"

def normalizar_el_placer(direccion: str):
    if pd.isna(direccion):
        return direccion, "0"
    txt = _limpiar_base(str(direccion))

    # Si no aparece "PLACER" en absoluto, salir
    if not re.search(r"\bPLACER\b", txt, flags=re.IGNORECASE):
        return direccion, "0"

    m = REGEX_EL_PLACER_MZ_CS.search(txt)
    if m:
        mz   = m.group("mz")
        cs   = m.group("cs")
        piso = m.group("piso")
        out = f"URB EL PLACER MZ {mz} CS {cs}"
        if piso:
            out += f" PI {int(piso)}"
        return out, "1"

    # Si no tenemos MZ/CS claros, al menos devolver un formato de barrio marcado
    return "URB EL PLACER", "0"

# ================== Despachador principal ==================

BARRIOS_SOPORTADOS = [
    "25 DE MAYO",
    "CIUDADELA EL SOL",
    "EL PLACER",
    "PLACER",
]

def normalizar_direccion(direccion: str):
    """
    Intenta normalizar la dirección usando las reglas de:
    - URB 25 DE MAYO
    - URB CIUDADELA EL SOL
    - URB EL PLACER

    Retorna (direccion_normalizada, "1") ó (direccion_limpia, "0").
    """
    if pd.isna(direccion):
        return direccion, "0"

    texto = _limpiar_base(str(direccion))

    if "25 DE MAYO" in texto:
        return normalizar_25_de_mayo(texto)

    if "CIUDADELA EL SOL" in texto:
        return normalizar_ciudadela_el_sol(texto)

    if "PLACER" in texto:
        return normalizar_el_placer(texto)

    return texto, "0"

# ================== FUNCIÓN PÚBLICA PARA EL PIPELINE ==================

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Espera un DataFrame con al menos:
      - 'DIRECCION'
      - 'NIU' o 'CLIENTE_ID'

    Devuelve un DataFrame con columnas:
      - NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION

    Filtra SOLO las filas cuyos barrios pertenecen a BARRIOS_SOPORTADOS.
    """
    if "DIRECCION" not in df_in.columns:
        raise ValueError("El DataFrame de entrada debe contener la columna 'DIRECCION'.")

    df = df_in.copy()

    # Asegurar columna de identificador NIU
    if "NIU" in df.columns:
        id_col = "NIU"
    elif "CLIENTE_ID" in df.columns:
        id_col = "CLIENTE_ID"
    else:
        raise ValueError("No se encontró columna 'NIU' ni 'CLIENTE_ID' en el DataFrame de entrada.")

    df = df[[id_col, "DIRECCION"]].copy()
    df.rename(columns={id_col: "NIU"}, inplace=True)

    # Aplicar normalizador
    df[["DIRECCION_NORMALIZADA", "VALIDACION"]] = df["DIRECCION"].apply(
        lambda x: pd.Series(normalizar_direccion(x))
    )

    # Filtrar por barrios soportados
    patron = "|".join(BARRIOS_SOPORTADOS)
    mask = df["DIRECCION"].astype(str).str.contains(patron, case=False, na=False)
    df_filtrado = df[mask].copy()

    return df_filtrado[["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"]]

# ================== MODO SCRIPT (reproduce comportamiento original) ==================

if __name__ == "__main__":
    # Leer Excel de entrada
    df_in = pd.read_excel(RUTA_ENTRADA, dtype=str)

    df_out = procesar(df_in)

    # Estadísticas
    total = df_out.shape[0]
    normalizadas = (df_out["VALIDACION"] == "1").sum()
    efectividad = (normalizadas / total * 100) if total > 0 else 0

    # Guardar en el archivo consolidado
    modo = "a" if os.path.exists(RUTA_SALIDA) else "w"
    with pd.ExcelWriter(RUTA_SALIDA, engine="openpyxl", mode=modo) as writer:
        df_out.to_excel(writer, sheet_name=HOJA_SALIDA, index=False)

    print(f"✅ Archivo procesado y guardado como {RUTA_SALIDA}")
    print(f"📄 Hoja: {HOJA_SALIDA}")
    print(f"🔍 Total direcciones filtradas: {total}")
    print(f"📊 Direcciones normalizadas: {normalizadas}")
    print(f"📈 Efectividad: {round(efectividad, 2)}%")
