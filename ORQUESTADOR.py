import os
import importlib.util
import pandas as pd

# ================== CONFIGURACIÓN ==================

CARPETA_ENTRADA = "entradas"   # carpeta donde pondrás todos los Excel/CSV
ARCHIVO_SALIDA = "CICLOS_PROCESADOS_UNIFICADO.xlsx"
HOJA_SALIDA = "NORMALIZADAS"

EXT_PERMITIDAS = {".xlsx", ".xls", ".csv"}

# nombre_amigable, archivo_py
SCRIPTS = [
    ("MAIN_MERCAR_ARMENIA", "main.py"),      # <--- este
    ("25MAYO",              "25mayo.py"),
    ("CECILIA",             "CecMain.py"),
    ("ARCOIRIS",            "Arcoiris.py"),
    ("CDORADA",             "CDorada.py"),
    ("CENTROCLL",           "centro29cll.py"),
    ("CHAMBRANA",           "chambranas.py"),
    ("CIBELES",             "cibeles.py"),
    ("COLINAS",             "colinas.py"),
    ("ISABELLA",            "Isabella.py"),
    ("MIRANDA",             "miranda.py"),
    ("QMARINA",             "Qmarina.py"),
    ("RECREO",              "recreo.py"),
    ("RPINILLA",            "Rpinilla.py"),
]

# ================== HELPERS ==================

def leer_entrada_flexible(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return pd.read_excel(path, dtype=str)
    elif ext == ".csv":
        df = pd.read_csv(path, dtype=str, sep=";")
        if df.shape[1] == 1:
            df = pd.read_csv(path, dtype=str, sep=",")
        return df
    else:
        raise ValueError(f"Extensión no soportada: {ext}")

def cargar_modulo(ruta_py: str, nombre_modulo: str):
    spec = importlib.util.spec_from_file_location(nombre_modulo, ruta_py)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)  # type: ignore[attr-defined]
    return modulo

def cargar_scripts_normalizacion(base_dir: str):
    modulos = []
    for nombre, archivo in SCRIPTS:
        ruta_script = os.path.join(base_dir, archivo)
        if not os.path.exists(ruta_script):
            print(f"⚠️ Script {archivo} no encontrado, se omite.")
            continue
        try:
            modulo = cargar_modulo(ruta_script, nombre)
        except Exception as e:
            print(f"⚠️ Error cargando {archivo}: {e}")
            continue

        if not hasattr(modulo, "procesar"):
            print(f"⚠️ El script {archivo} no define procesar(df), se omite.")
            continue

        modulos.append((nombre, modulo))
        print(f"✅ Script cargado: {archivo} (alias: {nombre})")
    return modulos

def procesar_archivo_con_modulos(ruta_archivo: str, modulos):
    print(f"\n📂 Procesando archivo: {os.path.basename(ruta_archivo)}")
    try:
        df_in = leer_entrada_flexible(ruta_archivo)
    except Exception as e:
        print(f"❌ Error leyendo {ruta_archivo}: {e}")
        return pd.DataFrame(columns=["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"])

    resultados = []
    for nombre, modulo in modulos:
        try:
            df_out = modulo.procesar(df_in)
        except Exception as e:
            print(f"   ⚠️ Error en procesar() de {nombre}: {e}")
            continue

        if df_out is None or df_out.empty:
            print(f"   ➖ {nombre}: 0 filas")
            continue

        columnas_esperadas = {"NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"}
        if not columnas_esperadas.issubset(df_out.columns):
            print(f"   ⚠️ {nombre}: faltan columnas estándar, se omite.")
            continue

        resultados.append(df_out)
        print(f"   ✅ {nombre}: {len(df_out)} filas")

    if resultados:
        return pd.concat(resultados, ignore_index=True)
    else:
        return pd.DataFrame(columns=["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"])

# ================== MAIN ==================

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    carpeta_entrada = os.path.join(base_dir, CARPETA_ENTRADA)

    if not os.path.exists(carpeta_entrada):
        print(f"❌ La carpeta de entrada no existe: {carpeta_entrada}")
        print("   Crea la carpeta y copia allí los Excel/CSV.")
        return

    modulos = cargar_scripts_normalizacion(base_dir)
    if not modulos:
        print("❌ No se cargó ningún script de normalización.")
        return

    archivos = sorted(os.listdir(carpeta_entrada))
    if not archivos:
        print(f"❌ La carpeta {CARPETA_ENTRADA} está vacía.")
        return

    todos = []
    for nombre_archivo in archivos:
        ext = os.path.splitext(nombre_archivo)[1].lower()
        if ext not in EXT_PERMITIDAS:
            print(f"   (Se omite {nombre_archivo}, extensión no soportada)")
            continue
        ruta_archivo = os.path.join(carpeta_entrada, nombre_archivo)
        df_archivo = procesar_archivo_con_modulos(ruta_archivo, modulos)
        if not df_archivo.empty:
            todos.append(df_archivo)

    if not todos:
        print("❌ No se generó ninguna fila normalizada.")
        return

    df_final = pd.concat(todos, ignore_index=True)

    ruta_salida = os.path.join(base_dir, ARCHIVO_SALIDA)
    with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode="w") as writer:
        df_final.to_excel(writer, sheet_name=HOJA_SALIDA, index=False)

    total = df_final.shape[0]
    normalizadas = (df_final["VALIDACION"] == "1").sum()
    efectividad = (normalizadas * 100 / total) if total > 0 else 0

    print("\n✅ PROCESO COMPLETO")
    print(f"📁 Archivo de salida: {ruta_salida}")
    print(f"📄 Hoja: {HOJA_SALIDA}")
    print(f"🔢 Total filas: {total}")
    print(f"✔️ Normalizadas (VALIDACION='1'): {normalizadas}")
    print(f"📈 Efectividad global: {efectividad:.2f}%")

if __name__ == "__main__":
    main()
