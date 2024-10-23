from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By  # Importación correcta de By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
import zipfile
import sys

# Configura la variable PATH para ChromeDriver
os.environ['PATH'] += os.pathsep + r'C:\Users\crisrobles\Downloads\chromedriver'  # Cambia esta ruta a donde has descargado ChromeDriver

# Argumentos del script
archivo_control = sys.argv[1]
ruta_destino = sys.argv[2]

# Leer archivo de control
with open(archivo_control, "r") as f:
    contenido_archivo_entrada = f.read()

x = contenido_archivo_entrada.split("|")
Nombre_directorio_bajada = x[1]
Nombre_id_elemento = x[0].strip()
print(Nombre_id_elemento)

# Configura las opciones de Chrome
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": ruta_destino,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Eliminar cada archivo en el directorio de descargas
archivos_en_descargas = os.listdir(ruta_destino)
for archivo in archivos_en_descargas:
    try:
        os.remove(os.path.join(ruta_destino, archivo))
        print(f"Se eliminó correctamente el archivo: {archivo}")
    except Exception as e:
        print(f"No se pudo eliminar el archivo {archivo}: {e}")

# Inicializa el navegador usando WebDriver Manager
try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.maximize_window()

    # Abre la página web
    driver.get("https://guiadevalores.fasecolda.com/ConsultaExplorador/Default.aspx?url=C:\\inetpub\\wwwroot\\Fasecolda\\ConsultaExplorador\\Guias\\GuiaValores_NuevoFormato\\" + Nombre_directorio_bajada)
    
    # Descargar primer archivo con scrollIntoView para asegurar visibilidad
    nombre_archivo = "Guia_Excel"
    download_link = driver.find_element(By.XPATH, f"//a[contains(text(), '{nombre_archivo}')]")
    
    # Scroll hasta el enlace antes de hacer clic
    driver.execute_script("arguments[0].scrollIntoView(true);", download_link)
    time.sleep(1)  # Breve pausa
    download_link.click()

    # Espera a que la descarga comience
    timeout = 60  # Espera hasta 60 segundos
    start_time = time.time()
    while any(archivo.endswith('.crdownload') for archivo in os.listdir(ruta_destino)):
        if time.time() - start_time > timeout:
            print("Tiempo de espera excedido para Guia_Excel.")
            break
        time.sleep(1)
    else:
        print("El archivo Guia_Excel se ha descargado correctamente.")

    # Descargar segundo archivo, con espera explícita y manejo del elemento bloqueador
    nombre_archivo = "Tabla de Homologacion"

    # Ocultar temporalmente el elemento que está bloqueando el clic
    elemento_bloqueador = driver.find_element(By.CSS_SELECTOR, ".contacto")
    driver.execute_script("arguments[0].style.visibility='hidden'", elemento_bloqueador)

    # Usar una espera explícita hasta que el enlace sea clickeable
    download_link = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{nombre_archivo}')]"))
    )
    download_link.click()

    # Espera a que la descarga se complete
    start_time = time.time()
    while any(archivo.endswith('.crdownload') for archivo in os.listdir(ruta_destino)):
        if time.time() - start_time > timeout:
            print("Tiempo de espera excedido para Tabla de Homologacion.")
            break
        time.sleep(1)
    else:
        print("El archivo Tabla de Homologacion se ha descargado correctamente.")

finally:
    # Cierra el navegador
    driver.quit()
    print("El navegador se cerró correctamente.")

# Descomprimir el archivo descargado
for archivo in os.listdir(ruta_destino):
    if archivo.endswith('.zip'):
        ruta_archivo_zip = os.path.join(ruta_destino, archivo)
        with zipfile.ZipFile(ruta_archivo_zip, 'r') as zip_ref:
            zip_ref.extractall(ruta_destino)
        print("El archivo ZIP se ha descomprimido correctamente.")
        # Puedes eliminar el archivo ZIP si lo deseas
        os.remove(ruta_archivo_zip)
