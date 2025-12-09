import os
import io
import importlib.util

import pandas as pd
import streamlit as st

# ================== CONFIGURACIÓN ==================

# Lista de scripts de normalización que quieres usar
# (nombre_amigable, archivo_py)
SCRIPTS = [
    ("MERCAR_ARMENIA", "main.py"),        # aquí está la lógica de MERCAR + NUEVO ARMENIA
    ("25MAYO",         "25mayo.py"),
    ("CECILIA",        "CecMain.py"),
    ("ARCOIRIS",       "Arcoiris.py"),
    ("CDORADA",        "CDorada.py"),
    ("CENTROCLL",      "centro29cll.py"),
    ("CHAMBRANA",      "chambranas.py"),
    ("CIBELES",        "cibeles.py"),
    ("COLINAS",        "colinas.py"),
    ("ISABELLA",       "Isabella.py"),
    ("MIRANDA",        "miranda.py"),
    ("QMARINA",        "Qmarina.py"),
    ("RECREO",         "recreo.py"),
    ("RPINILLA",       "Rpinilla.py"),
]

COLUMNAS_SALIDA = ["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"]


# ================== HELPERS PARA CARGAR MÓDULOS ==================

def cargar_modulo_desde_archivo(ruta_py: str, nombre_modulo: str):
    """
    Carga dinámicamente un archivo .py (aunque empiece con número, ej. 25mayo.py)
    y devuelve el módulo importado.
    """
    spec = importlib.util.spec_from_file_location(nombre_modulo, ruta_py)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)  # type: ignore[attr-defined]
    return modulo


@st.cache_resource
def cargar_scripts_normalizacion():
    """
    Carga todos los scripts definidos en SCRIPTS y devuelve:
    [(nombre_amigable, modulo), ...]
    solo para los que tienen procesar(df).
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    modulos = []

    for nombre, archivo in SCRIPTS:
        ruta_script = os.path.join(base_dir, archivo)
        if not os.path.exists(ruta_script):
            st.warning(f"Script no encontrado y se omitirá: **{archivo}**")
            continue

        try:
            modulo = cargar_modulo_desde_archivo(ruta_script, nombre)
        except Exception as e:
            st.warning(f"Error cargando `{archivo}`: {e}")
            continue

        if not hasattr(modulo, "procesar"):
            st.warning(f"El script `{archivo}` no define la función `procesar(df)`, se omitirá.")
            continue

        modulos.append((nombre, modulo))

    return modulos


def procesar_df_con_modulos(df_in: pd.DataFrame, modulos):
    """
    Aplica TODOS los módulos de normalización a un DataFrame de entrada.
    Devuelve un DataFrame con columnas estándar:
    NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION
    """
    resultados = []

    for nombre, modulo in modulos:
        try:
            df_out = modulo.procesar(df_in)
        except Exception as e:
            st.warning(f"Error en `procesar()` de **{nombre}**: {e}")
            continue

        if df_out is None or df_out.empty:
            # st.info(f"{nombre}: 0 filas devueltas.")
            continue

        # Verificar columnas
        if not set(COLUMNAS_SALIDA).issubset(df_out.columns):
            st.warning(
                f"El módulo **{nombre}** no devuelve todas las columnas estándar "
                f"{COLUMNAS_SALIDA}. Se omitirá su resultado."
            )
            continue

        resultados.append(df_out[COLUMNAS_SALIDA])

    if resultados:
        return pd.concat(resultados, ignore_index=True)
    else:
        return pd.DataFrame(columns=COLUMNAS_SALIDA)


def procesar_archivos_subidos(archivos, modulos):
    """
    Recorre la lista de archivos subidos en Streamlit, lee cada uno
    como Excel y aplica los módulos de normalización.
    Devuelve un DataFrame unificado.
    """
    todos = []

    for f in archivos:
        st.write(f"📂 Procesando archivo: **{f.name}**")
        try:
            df_in = pd.read_excel(f, dtype=str)
        except Exception as e:
            st.warning(f"No se pudo leer `{f.name}` como Excel: {e}")
            continue

        df_proc = procesar_df_con_modulos(df_in, modulos)
        if not df_proc.empty:
            # opcional: agregar de qué archivo viene
            df_proc = df_proc.copy()
            df_proc["ARCHIVO_ORIGEN"] = f.name
            todos.append(df_proc)

    if todos:
        return pd.concat(todos, ignore_index=True)
    else:
        return pd.DataFrame(columns=COLUMNAS_SALIDA + ["ARCHIVO_ORIGEN"])


# ================== INTERFAZ STREAMLIT ==================

def main():
    st.set_page_config(page_title="Normalizador de direcciones", layout="wide")
    st.title("🧠 Normalizador de direcciones (Streamlit)")

    st.markdown(
        """
        Carga **tus archivos de Excel** con las direcciones de los diferentes ciclos  
        (por ejemplo, los 5 archivos que usas actualmente), y el sistema aplicará
        todas las reglas de normalización que ya tienes programadas
        (MERCAR, Nuevo Armenia, Cecilio, Arcoiris, etc.).

        Al final podrás descargar **un solo archivo Excel** con las columnas:

        - `NIU`
        - `DIRECCION` (original)
        - `DIRECCION_NORMALIZADA`
        - `VALIDACION`
        """
    )

    archivos = st.file_uploader(
        "Sube aquí tus archivos de Excel (puedes subir hasta 5 a la vez)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )

    # Sólo informativo: no bloqueamos si son más o menos de 5
    if archivos and len(archivos) != 5:
        st.info(f"Has subido {len(archivos)} archivo(s). "
                f"Tu flujo habitual usa 5, pero el sistema funciona con cualquier número ≥ 1.")

    if st.button("🚀 Procesar archivos", type="primary"):
        if not archivos:
            st.warning("Primero debes subir al menos un archivo de Excel.")
            return

        with st.spinner("Cargando scripts y procesando direcciones..."):
            modulos = cargar_scripts_normalizacion()
            if not modulos:
                st.error("No se pudo cargar ningún script de normalización. Revisa los archivos .py en el repositorio.")
                return

            df_final = procesar_archivos_subidos(archivos, modulos)

        if df_final.empty:
            st.warning("No se generó ninguna dirección normalizada. "
                       "Verifica que los archivos tengan columnas esperadas (NIU/CLIENTE_ID y DIRECCION).")
            return

        st.success("¡Procesamiento completado!")

        # Métricas básicas
        total = df_final.shape[0]
        normalizadas = (df_final["VALIDACION"] == "1").sum()
        efectividad = (normalizadas * 100 / total) if total > 0 else 0

        col1, col2, col3 = st.columns(3)
        col1.metric("Total filas en salida", total)
        col2.metric("Normalizadas (VALIDACION = '1')", normalizadas)
        col3.metric("Efectividad global (%)", f"{efectividad:.2f}")

        st.subheader("Vista previa de resultados")
        st.dataframe(df_final.head(200))

        # ---- Generar Excel para descarga ----
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_final.to_excel(writer, sheet_name="NORMALIZADAS", index=False)
        buffer.seek(0)

        st.download_button(
            label="⬇️ Descargar Excel unificado",
            data=buffer,
            file_name="direcciones_normalizadas_unificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
