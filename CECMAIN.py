import pandas as pd
import re
import os

# ================== Configuración por defecto (modo script) ==================
RUTA_ENTRADA = "CICLO 49_PDIRECCION.xlsx"
RUTA_SALIDA  = "CICLOS_PROCESADOS.xlsx"
HOJA_SALIDA  = "Direcciones_Cecilia"
COL_DIR      = "DIRECCION"

# Palabras clave para filtrar las direcciones de interés
PALABRAS_CLAVE = ["cecilia", "ceciclia", "ecilia", "villa yolanda", "koa"]

# ================== Diccionario romanos ==================
romanos_a_numeros = {
    "I": "1", "II": "2", "III": "3", "IV": "4", "V": "5",
    "VI": "6", "VII": "7", "VIII": "8", "IX": "9", "X": "10"
}

# ================== Patrones principales ==================

# Captura sectores: LA CECILIA / BOSQUES DE LA CECILIA / VILLA YOLANDA
regex = re.compile(
    r'\b(?:URB|BRR|CJT|UR)\s+'
    r'((?:LA\s+)?(?:CECILIA|CECICLIA|ECILIA)|'
    r'(?:BOSQUES|BQ|BQUES|UES|BQDE)\s+(?:DE\s+)?LA\s+CECILIA|'
    r'(?:VILLA\s+)?YOLANDA'
    r')\s*'
    r'(?:E(?:T|TAPA|TP)?\s*(\d+)|(\d+|I{1,3}|IV|V|VI{0,3}|IX|X))?\s*'
    r'M(?:ZN?|NZ|ZA|N)?\s*([A-Z\d]+)\s*'
    r'(?:CS|#|CASA|CAS|C)?\s*(\d+)?\s*'
    r'(?:PISO|PI|P|PIS)?\s*(\d+)?\s*'
    r'(?:ETAPA|ET|TP)?\s*(\d+)?'
    r'(?:-?\s*ARMENIA)?',  # Permite que "- ARMENIA" sea opcional
    re.IGNORECASE
)

# Lotes que NO queremos normalizar
regex_lote = re.compile(r'\b(LOTE|LT)\b', re.IGNORECASE)

# Palabras adicionales (frente a cancha, ens, etc.) para limpiar
regex_info_adicional = re.compile(
    r'\b('
    r'FTE(?:\s+A)?|'
    r'FT|'
    r'FRENTE(?:\s+A\s+LA\s+CANCHA)?|'
    r'CANCHA|'
    r'ENS(?:\s+CANCHA)?|'
    r'DETRÁS(?:\s+DE\s+LA\s+CANCHA)?|'
    r'DETRAS|'
    r'DT|'
    r'ETPI|'
    r'ENSEG'
    r')\b',
    re.IGNORECASE
)

# Formatos "ideales" de salida
regex_estandar = re.compile(
    r'^(URB LA CECILIA MZ \d+[A-Z]? CS \d+(?: PI \d+)?(?: ET \d+)?)$|'
    r'^(BRR BOSQUES DE LA CECILIA MZ \d+[A-Z]? CS \d+(?: PI \d+)?(?: ET \d+)?)$|'
    r'^(URB VILLA YOLANDA MZ [A-Z\d]+ CS \d+(?: PI \d+)?)$|'
    r'^(CLL \d+ CR \d+\s*-\d+ TO \d+ (?:AC|AP \d+) VILLA CECILIA)$|'
    r'^(CRA \d+[A-Z]? CL \d+(?:\s*-\s*|\s+)\d+ AP \d+ ED [A-Z\d\s]+)$',
    re.IGNORECASE
)

# ================== Normalización ==================

def normalizar_direccion_cecilia(direccion: str) -> str:
    """
    Aplica la lógica original de normalización de Cecilia/Bosques/Villa Yolanda.
    Devuelve SOLO la dirección normalizada (no la info adicional).
    """
    if pd.isna(direccion):
        return direccion

    direccion = str(direccion).upper().strip()

    # Eliminar palabras adicionales primero
    direccion = regex_info_adicional.sub("", direccion).strip()

    # Si es lote, no normalizamos
    if regex_lote.search(direccion):
        return direccion

    # Quitar sufijos tipo "- AV ..." al final
    direccion = re.sub(
        r'\s*-\s*(AV|CALLE|CARRERA|CLL|CR|TO|AP|VILLA CECILIA)$',
        '',
        direccion
    )

    # Caso especial: CLL x CR y - z TO t AP a  -> agregar VILLA CECILIA
    match_calle_torre = re.match(r'^CLL \d+ CR \d+\s*-\d+ TO \d+ AP \d+', direccion)
    if match_calle_torre and "VILLA CECILIA" not in direccion:
        return f"{direccion} VILLA CECILIA"

    # Normalización por patrón URB/BRR
    match = regex.search(direccion)

    if match:
        sector = match.group(1)
        etapa  = match.group(2) or match.group(3) or ""
        manzana = match.group(4)
        casa    = match.group(5) or ""
        piso    = match.group(6) or ""

        # Convertir romanos
        etapa = romanos_a_numeros.get(etapa, etapa)

        # Determinar barrio
        if ("BOSQUES" in sector) or ("BQ" in sector) or ("BQUES" in sector) or ("UES" in sector):
            nueva = f"BRR BOSQUES DE LA CECILIA MZ {manzana} CS {casa}"
        elif "YOLANDA" in sector:
            nueva = f"URB VILLA YOLANDA MZ {manzana} CS {casa}"
        else:
            nueva = f"URB LA CECILIA MZ {manzana} CS {casa}"

        if piso:
            nueva += f" PI {piso}"
        if etapa:
            nueva += f" ET {etapa}"

        return nueva

    # Si nada matchea, devolvemos la dirección limpia tal cual
    return direccion

def cumple_estandar(direccion: str) -> int:
    """
    Usa el regex_estandar original para marcar si la dirección está
    en uno de los formatos esperados.
    """
    if not isinstance(direccion, str):
        return 0
    return 1 if regex_estandar.match(direccion.strip()) else 0

# ================== Función pública para el pipeline ==================

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Espera un DataFrame con al menos:
      - 'DIRECCION'
      - 'NIU' o 'CLIENTE_ID'

    Devuelve:
      NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION

    Solo se devuelven filas cuyas DIRECCION contienen
    alguna palabra en PALABRAS_CLAVE (cecilia, villa yolanda, koa, etc.)
    """
    if COL_DIR not in df_in.columns:
        raise ValueError("El DataFrame de entrada debe contener la columna 'DIRECCION'.")

    df = df_in.copy()

    # Asegurar columna de identificador NIU
    if "NIU" in df.columns:
        id_col = "NIU"
    elif "CLIENTE_ID" in df.columns:
        id_col = "CLIENTE_ID"
    else:
        raise ValueError("No se encontró columna 'NIU' ni 'CLIENTE_ID' en el DataFrame de entrada.")

    df = df[[id_col, COL_DIR]].copy()
    df.rename(columns={id_col: "NIU", COL_DIR: "DIRECCION"}, inplace=True)

    # Normalización
    df["DIRECCION_NORMALIZADA"] = df["DIRECCION"].apply(normalizar_direccion_cecilia)

    # Validación usando el regex_estandar
    df["VALIDACION"] = df["DIRECCION_NORMALIZADA"].apply(
        lambda x: "1" if cumple_estandar(x) == 1 else "0"
    )

    # Filtro por palabras clave (como en el script original)
    patron = "|".join(PALABRAS_CLAVE)
    mask = df["DIRECCION"].astype(str).str.contains(patron, case=False, na=False)
    df_filtrado = df[mask].copy()

    return df_filtrado[["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"]]

# ================== Modo script standalone ==================

if __name__ == "__main__":
    # Carga
    df_in = pd.read_excel(RUTA_ENTRADA, dtype=str)

    # Procesar
    df_out = procesar(df_in)

    # Estadísticas
    total = df_out.shape[0]
    normalizadas = (df_out["VALIDACION"] == "1").sum()
    efectividad = (normalizadas * 100 / total) if total > 0 else 0

    # Guardado (adjunta si ya existe)
    modo = "a" if os.path.exists(RUTA_SALIDA) else "w"
    with pd.ExcelWriter(RUTA_SALIDA, engine="openpyxl", mode=modo) as writer:
        df_out.to_excel(writer, sheet_name=HOJA_SALIDA, index=False)

    print(f"✅ Archivo procesado y guardado como {RUTA_SALIDA}")
    print(f"📄 Hoja: {HOJA_SALIDA}")
    print(f"🔍 Total de direcciones filtradas: {total}")
    print(f"📊 Direcciones que cumplen el estándar: {normalizadas}")
    print(f"📈 La efectividad es: {round(efectividad, 2)}%")