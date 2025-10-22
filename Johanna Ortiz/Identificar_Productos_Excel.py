import pandas as pd
import csv
import os
import time

# --- Configuración ---
TXT_PATH = "Productos_Pendientes.txt"  # archivo con los códigos
INPUT_DIR = "Archivos"          # carpeta donde están tus archivos Excel o CSV
OUTPUT_CSV = "resultados_locales.csv"   # salida

# --- Leer los códigos del TXT ---
with open(TXT_PATH, "r", encoding="utf-8") as f:
    codigos = [line.strip() for line in f if line.strip()]
codigos_set = set(codigos)
print(f"📄 {len(codigos_set)} códigos cargados desde {TXT_PATH}")

# --- Función para leer archivos Excel o CSV ---
def leer_archivo(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx" or ext == ".xls":
        # Lee todas las hojas del Excel
        xls = pd.ExcelFile(path)
        return {hoja: pd.read_excel(xls, hoja) for hoja in xls.sheet_names}
    elif ext == ".csv":
        return {"Hoja1": pd.read_csv(path, dtype=str)}
    else:
        print(f"⚠️ Formato no compatible: {path}")
        return {}

# --- Procesar archivos ---
with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Archivo", "Hoja", "Fila", "Columna", "Código encontrado"])
    f.flush()

    for nombre_archivo in os.listdir(INPUT_DIR):
        ruta = os.path.join(INPUT_DIR, nombre_archivo)
        if not os.path.isfile(ruta):
            continue

        print(f"🔍 Buscando en: {nombre_archivo}")

        try:
            hojas = leer_archivo(ruta)
        except Exception as e:
            print(f"❌ Error al leer {nombre_archivo}: {e}")
            continue

        for hoja, df in hojas.items():
            print(f"   ↳ Revisando hoja: {hoja}")
            # Convertir todo a texto para evitar errores de tipo
            df = df.astype(str)
            filas, columnas = df.shape

            for i in range(filas):
                for j in range(columnas):
                    valor = str(df.iat[i, j]).strip()
                    if valor in codigos_set:
                        writer.writerow([nombre_archivo, hoja, i + 1, j + 1, valor])
                        f.flush()

print(f"✅ Búsqueda completada. Resultados guardados en {OUTPUT_CSV}")
