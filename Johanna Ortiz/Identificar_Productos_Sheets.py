import gspread
from google.oauth2.service_account import Credentials
import csv
import time
from gspread.exceptions import APIError, SpreadsheetNotFound

# --- Configuración ---
TXT_PATH = "Productos_Pendientes.txt"
SPREADSHEETS = [
    "1U7BuK-z9CYdYNb_6b2UY4AAjCYWp2QLVuOwaVgxiYb8",
    "1Abd2kbIc5qIxjHokYGwdKtHp0X137xgZ_5EeLPY5IPI", 
    "18ofBTsQeZCLMYIZh8Q3n3B9XM01hrz6-voo2gZunxAM",
    "1-iiYxxo39Nj9oGB4sKeXtdRofUeMGLjMpUD_kVjGSko",
    "1EDJplVvCbPONkKIEbFY_6zCS-1ZE88fe0TWohKfsQVs",
    "1mZju2KdQiq2xVHE21NCaI_0OEfphrVhrbPA598OcQYU",
    "1TsNVq7QLOdCH0zET8Z_4bPhjMe-3ApRYMUeiRAMNVvE",
    "1kBIcyVYiRFCH0d_EnvbZ8phu3gsbFt5FDgyjcsubwRg",
    "1N8fteVklo6GJq-5XjWIoOea-ahfbVtf7KCudqMfVzJY",
    "1zKuulN_Ug6t4K7HZbAtV0oroJ5Z5-dZxmHv0UnVuJew",
    "13fWRamivIUKyBpkC-pII48m5lcPB36f_04b3UAQIhJY",
    "1jWeu15ZsINBrW8fTR-SJR0Biuvv53hKhsKkddTWTpsc",
    "1C0T_0tJGI2zfuSrHfpDSJD2FRxt0VJPuQhJIDiuf_Ok",
    "1ONV4UBJh2W4LqAx21BNRzzTvDRfhVw_wZ21OholOMG0",
    "1Nln-vZARLkczE69QVG8vxvvs1DmOObdi83Ao_bPP_50",
    "1wy5lvDwYSAKXimlCXBkY_C56lkFf5pEoT1gti15bSks",
    "1mUxn05SPICcRqdy_o-QIhpulifNcVb9G-bV1Wnhd3pI",
    "1fIobXVtBj3ZP1jv9s6bA10CXemnO2avVODwU4FafXvM",
    "17I3clLM15KuJT6sztITHkO-JMI72-gxB-391g04HvrU",
    "15HZpd0Nei4d9JZCktKkJ__gFlR6tjkY0w8LEwsCfHl0",
    "1jodjob8RQJp9DQDx8A-MptGx5mPHgDbAYe42DkcLyFA",
    "1m8CLvRoglyquTP4CP7S9lqGP1-E-obvRK9JOelfLQXA",
    "1UNjtd11eSTbJo-uEYacH42xDT2Gl8bZUM2NEOTaI-qw",
    "1IrtQwEu3H0bIXVeBU6SJcrbsSFuJ2dJJgFnNzygwJkA",
    "1GY-itwxCDjnVjApT8ze4rXpusago8UrIW9vxXVh5PZ0",
    "14F1auvC559f-ls8T7m7Y--7hHWcGlicwpmFzAEmbrmc",
    "1G-FiSl4Zn1jY_FEgetxnW2GuLWFswtEJ7yqbuPoDgrA",
    "1y5s3UdS-D5k6ypwQ8R-Tnvqwzs9cKkaDFZGK9UNwJco",
    "1C24hD9LqQHksTQpzRfPjfMP8WlXwjBY9IRdSv6Ai7SM",
    "1AC0rOHHS4YlAXdlf-5Nckybwju1QqJ2UToqU8ONmx7c",
    "1nZVcvL_9WHS-NXg_yFzwdLK50okL6gTl2lFDgsXeQVw",
    "14wiOvdhhyTzhu9dxmUncClO1VanQr-pbEYJwI9ZNlrQ",
    "1D07kZYCJw1ztPaqhygASQca3r_ul44m7MzlddaGDBXY",
    "1AC0rOHHS4YlAXdlf-5Nckybwju1QqJ2UToqU8ONmx7c",
    "15oT_aREdalFmXLKlgLkX5NOel8XvJNSPS6LmZorp9Ss",
    "1xekPDOCKw4V_zeMrxrPJa06qSSDoYp65Z3b7poPuOTE"
    
]
OUTPUT_CSV = "resultados.csv"

# --- Autenticación ---
scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = Credentials.from_service_account_file("credenciales.json", scopes=scopes)
client = gspread.authorize(creds)

# --- Leer los códigos del TXT ---
with open(TXT_PATH, "r", encoding="utf-8") as f:
    codigos = [line.strip() for line in f if line.strip()]
codigos_set = set(codigos)  # más rápido para búsquedas

# --- Función segura para leer datos por bloques ---
def leer_datos_en_bloques(ws, bloque_filas=1000):
    """Lee una worksheet en bloques sin exceder sus límites reales"""
    fila_inicio = 1
    all_data = []

    # Detectar límites reales de la hoja
    max_filas = ws.row_count
    max_columnas = ws.col_count

    # Convertir número de columna a letra (por ejemplo, 29 -> AC)
    def numero_a_columna(n):
        resultado = ""
        while n > 0:
            n, resto = divmod(n - 1, 26)
            resultado = chr(65 + resto) + resultado
        return resultado

    ultima_columna = numero_a_columna(max_columnas)

    while fila_inicio <= max_filas:
        fila_fin = min(fila_inicio + bloque_filas - 1, max_filas)
        rango = f"A{fila_inicio}:{ultima_columna}{fila_fin}"
        try:
            bloque = ws.get(rango)
            if not bloque:
                break
            all_data.extend(bloque)
        except APIError as e:
            print(f"⚠️ Error al leer {ws.title} (filas {fila_inicio}-{fila_fin}): {e}")
            time.sleep(5)
            continue
        fila_inicio += bloque_filas
        time.sleep(1)

    return all_data

# --- Buscar en los Sheets ---
with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Archivo", "Hoja", "Fila", "Columna", "Código encontrado"])
    f.flush()

    for sid in SPREADSHEETS:
        if not sid.strip():
            continue

        try:
            sheet = client.open_by_key(sid)
        except SpreadsheetNotFound:
            print(f"❌ No se encontró una hoja con el ID {sid} o no hay permiso de acceso.")
            continue
        except APIError as e:
            print(f"❌ No se pudo abrir hoja {sid}: {e}")
            continue

        print(f"🔍 Buscando en: {sheet.title}")

        for ws in sheet.worksheets():
            print(f"   ↳ Leyendo hoja: {ws.title}")
            datos = leer_datos_en_bloques(ws)

            for i, fila in enumerate(datos):
                for j, celda in enumerate(fila):
                    valor = celda.strip()
                    if valor and valor in codigos_set:
                        writer.writerow([sheet.title, ws.title, i + 1, j + 1, celda])
                        f.flush()

print(f"✅ Búsqueda completada. Resultados guardados en {OUTPUT_CSV}")
