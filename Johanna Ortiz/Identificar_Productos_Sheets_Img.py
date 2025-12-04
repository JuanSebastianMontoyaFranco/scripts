import time
import re
import httplib2
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials

# ==============================
# CONFIGURACIÓN
# ==============================
SERVICE_ACCOUNT_FILE = "credenciales.json"  # tu archivo JSON con credenciales
SPREADSHEET_IDS = [  # 🔹 Lista de archivos a analizar
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
TIMEOUT_SECONDS = 300  # 5 minutos
MAX_RETRIES = 5
RANGE_LIMIT = "A1:Z300"

# ==============================
# CONEXIÓN A GOOGLE SHEETS API
# ==============================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Nueva forma de crear un cliente HTTP autorizado
authed_http = AuthorizedSession(creds)

# Creamos el cliente del servicio Sheets
service = build("sheets", "v4", credentials=creds)

# ==============================
# FUNCIÓN: OBTENER NOMBRES DE HOJAS
# ==============================
def obtener_hojas(spreadsheet_id):
    """Devuelve los nombres de todas las hojas de un archivo."""
    for intento in range(MAX_RETRIES):
        try:
            metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            return [s["properties"]["title"] for s in metadata.get("sheets", [])]
        except (TimeoutError, HttpError) as e:
            print(f"⚠️ Error al listar hojas ({spreadsheet_id}), intento {intento+1}: {e}")
            time.sleep(5)
    print(f"❌ No se pudieron obtener las hojas de {spreadsheet_id}.")
    return []

# ==============================
# FUNCIÓN: PROCESAR UNA HOJA
# ==============================
def procesar_hoja(spreadsheet_id, nombre_hoja):
    """Lee una hoja y detecta imágenes, URLs o fórmulas =IMAGE()."""
    for intento in range(MAX_RETRIES):
        try:
            sheet = service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                ranges=[f"{nombre_hoja}!{RANGE_LIMIT}"],
                includeGridData=True
            ).execute()
            break
        except (TimeoutError, HttpError) as e:
            print(f"⚠️ Intento {intento+1} falló en '{nombre_hoja}': {e}, reintentando...")
            time.sleep(5)
    else:
        print(f"❌ No se pudo obtener la hoja '{nombre_hoja}' del archivo {spreadsheet_id}.")
        return {}

    imagenes_por_fila = {}

    for s in sheet.get("sheets", []):
        for grid in s.get("data", []):
            for fila_idx, row in enumerate(grid.get("rowData", []), start=1):
                if "values" not in row:
                    continue

                for celda_idx, cell in enumerate(row["values"], start=1):
                    formula = cell.get("userEnteredValue", {}).get("formulaValue")
                    valor = cell.get("formattedValue", "")
                    url_img = None

                    # 1️⃣ Detectar fórmula =IMAGE("url")
                    if formula and formula.startswith("=IMAGE("):
                        match = re.search(r'"(https?://[^"]+)"', formula)
                        if match:
                            url_img = match.group(1)

                    # 2️⃣ Detectar URL de imagen en texto
                    elif valor and re.match(r"https?://.*\.(png|jpg|jpeg|gif|webp)", valor):
                        url_img = valor

                    if url_img:
                        imagenes_por_fila.setdefault(fila_idx, []).append(url_img)

                    # 3️⃣ Detectar posibles imágenes en notas o metadatos
                    note = cell.get("note", "")
                    if note and "https" in note:
                        matches = re.findall(r"https?://[^\s]+", note)
                        for m in matches:
                            imagenes_por_fila.setdefault(fila_idx, []).append(m)

    return imagenes_por_fila

# ==============================
# FUNCIÓN PRINCIPAL
# ==============================
def procesar_varios_archivos():
    """Procesa todos los archivos y hojas indicadas en SPREADSHEET_IDS."""
    resultados = {}

    for archivo_id in SPREADSHEET_IDS:
        print(f"\n🔍 Analizando archivo: {archivo_id}")
        hojas = obtener_hojas(archivo_id)
        if not hojas:
            print("⚠️ No se encontraron hojas.")
            continue

        resultados[archivo_id] = {}

        for hoja in hojas:
            print(f"   📄 Procesando hoja: {hoja}")
            imagenes = procesar_hoja(archivo_id, hoja)
            resultados[archivo_id][hoja] = imagenes

    return resultados

# ==============================
# EJECUCIÓN
# ==============================
if __name__ == "__main__":
    print("🚀 Iniciando análisis de múltiples archivos...")
    resultados = procesar_varios_archivos()

    print("\n📸 RESULTADOS GLOBALES:\n")
    for archivo, hojas in resultados.items():
        print(f"\n🗂️ Archivo: {archivo}")
        for hoja, filas in hojas.items():
            print(f"  📄 Hoja: {hoja}")
            if not filas:
                print("     (Sin imágenes detectadas)")
                continue
            for fila, urls in filas.items():
                print(f"     Fila {fila}:")
                for url in urls:
                    print(f"        🔗 {url}")
