import pandas as pd
import re
import os

# ============================================================
#  RECREO / PALMARES DEL RECREO - NORMALIZACIÓN
# ============================================================

# Patrones de barrio
PATRON_RECREO = r"(?:EL\s+)?RECREO"
PATRON_PALMARES = (
    r"(PALMARES\s+DEL\s+RECREO|PALMARES\s+DE\s+RECREO|"
    r"PALMARES\s+RECREO|PALMA\s+DEL\s+RECREO|PALMA\s+DE\s+RECREO)"
)


def limpiar_espacios(txt: str) -> str:
    txt = re.sub(r"\s+", " ", txt or "").strip()
    return txt


def normalizar_barrio_cerrado(
    direccion: str,
    patron_barrio: str,
    nombre_barrio_normalizado: str,
) -> tuple[str, str]:
    """
    Busca estructuras tipo:

      BRR EL RECREO MZ A CS 12
      URB PALMARES DEL RECREO MZN B CASA 4
      CND PALMARES DE RECREO MZ C  CS 3
      BRR EL RECREO MZ D 12   (sin 'CS' explícito)

    y devuelve algo como:
      "BARRIO EL RECREO MZ A CS 12"
      "URB PALMARES DEL RECREO MZ B CS 4"
    """

    if not isinstance(direccion, str):
        return "", "0"

    d = limpiar_espacios(direccion.upper())

    # 1) Buscar el barrio, permitiendo prefijos como BRR, URB, CND, etc.
    patron_completo = re.compile(
        rf"""
        (?:
            BRR|BARRIO|URB(?:ANIZACION)?|CND|CONJ|CONJUNTO|CJT
        )?
        \s*
        {patron_barrio}
        (?P<resto>.*)
        """,
        re.IGNORECASE | re.VERBOSE,
    )

    m = patron_completo.search(d)
    if not m:
        return d, "0"

    resto = limpiar_espacios(m.group("resto") or "")

    # 2) Extraer manzana / casa / piso / apto
    m_mz = re.search(
        r"\b(MNZ|MNZ\.|MZN|MZNA|MZ\.?|MANZANA|MZA|MZNA|M)\s*([A-ZÑ0-9]{1,3})\b",
        resto,
        flags=re.IGNORECASE,
    )
    m_cs = re.search(
        r"\b(CS|CASA|CAS|C)\s*([A-Z]?\d{1,4}[A-Z]?)\b",
        resto,
        flags=re.IGNORECASE,
    )
    m_ap = re.search(
        r"\b(APTO?|APARTAMENTO|AP)\s*(\d{1,4}[A-Z]?)\b",
        resto,
        flags=re.IGNORECASE,
    )
    m_pi = re.search(
        r"\b(PI|PISO|PIS)\s*(\d{1,2})\b",
        resto,
        flags=re.IGNORECASE,
    )

    # Construcción de salida
    partes = []

    # Para PALMARES DEL RECREO lo catalogamos como URB,
    # para EL RECREO lo dejamos como BARRIO.
    if "PALMARES" in nombre_barrio_normalizado.upper():
        partes.append(f"URB {nombre_barrio_normalizado}")
    else:
        partes.append(f"BARRIO {nombre_barrio_normalizado}")

    if m_mz:
        mz = m_mz.group(2).upper()
        partes.append(f"MZ {mz}")

    if m_cs:
        cs = m_cs.group(2).upper()
        partes.append(f"CS {cs}")
    else:
        # Si no hay 'CASA/CS' explícito pero hay un número suelto, asumir casa
        m_num = re.search(r"\b(\d{1,4}[A-Z]?)\b", resto)
        if m_num:
            partes.append(f"CS {m_num.group(1).upper()}")

    if m_ap:
        ap = m_ap.group(2).upper()
        partes.append(f"AP {ap}")

    if m_pi:
        pi = m_pi.group(2)
        partes.append(f"PI {int(pi)}")

    direccion_normalizada = limpiar_espacios(" ".join(partes))

    # Consideramos validación = "1" si hay barrio + algún número
    tiene_numero = re.search(r"\d", direccion_normalizada) is not None
    validacion = "1" if tiene_numero else "0"

    return direccion_normalizada, validacion


def normalizar_direccion(direccion: str) -> tuple[str, str]:
    """
    Normalización de direcciones para:
      - BARRIO EL RECREO
      - URB PALMARES DEL RECREO

    Si no matchea ninguno de estos barrios,
    devuelve la dirección en mayúsculas con VALIDACION = "0".
    """

    if not isinstance(direccion, str):
        return "", "0"

    d = limpiar_espacios(direccion.upper())

    # --- PALMARES DEL RECREO primero (para no confundir con RECREO a secas) ---
    if re.search(PATRON_PALMARES, d, flags=re.IGNORECASE):
        out, val = normalizar_barrio_cerrado(
            d,
            PATRON_PALMARES,
            "PALMARES DEL RECREO",
        )
        return out, val

    # --- BARRIO EL RECREO ---
    if re.search(PATRON_RECREO, d, flags=re.IGNORECASE):
        # Excluir otros 'RECREO' raros si quisieras (ej: PARQUE EL RECREO),
        # pero por ahora cualquier RECREO lo tomamos como barrio.
        out, val = normalizar_barrio_cerrado(
            d,
            PATRON_RECREO,
            "EL RECREO",
        )
        return out, val

    # No pertenece a estos barrios
    return d, "0"


# ============================================================
#      FUNCIÓN ESTÁNDAR PARA USAR DESDE EL SCRIPT MAESTRO
# ============================================================

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe un DataFrame con al menos:
      - NIU o CLIENTE_ID
      - DIRECCION

    Devuelve SOLO filas de los barrios:
      - EL RECREO
      - PALMARES DEL RECREO

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

    # Filtrar SOLO las direcciones que contengan RECREO / PALMARES DEL RECREO
    filtro = df["DIRECCION"].str.contains(
        r"RECREO|PALMARES",
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
#      BLOQUE OPCIONAL PARA EJECUTAR ESTE SCRIPT SOLO
# ============================================================

if __name__ == "__main__":
    # Rutas como en tu script original
    ruta_entrada = "CICLO 47_PDIRECCION.xlsx"
    ruta_salida = "CICLOS_PROCESADOS.xlsx"

    # Leer datos
    df_local = pd.read_excel(ruta_entrada, usecols=["CLIENTE_ID", "DIRECCION"])

    # Procesar usando la función estándar
    df_resultado = procesar(df_local)

    # Guardar resultados en la hoja "RECREO"
    modo = "a" if os.path.exists(ruta_salida) else "w"
    with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode=modo) as writer:
        df_resultado.to_excel(writer, sheet_name="RECREO", index=False)

    # Estadísticas
    total = df_resultado.shape[0]
    normalizadas = df_resultado[df_resultado["VALIDACION"] == "1"].shape[0]
    efectividad = (normalizadas / total) * 100 if total > 0 else 0

    print(f"✅ Archivo procesado y guardado como {ruta_salida}")
    print(f"📄 Hoja: RECREO")
    print(f"🔍 Total direcciones procesadas (RECREO / PALMARES): {total}")
    print(f"📊 Direcciones normalizadas correctamente: {normalizadas}")
    print(f"📈 Efectividad: {efectividad:.2f}%")
