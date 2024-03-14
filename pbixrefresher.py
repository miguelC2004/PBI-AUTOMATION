import sys
import time
import os
import psutil
from pywinauto.application import Application
from pywinauto.keyboard import send_keys

# Configuración del archivo de trabajo
WORKBOOK = r"ruta"  # Ruta al archivo de trabajo
INIT_WAIT = 180  # Tiempo de espera inicial
REFRESH_TIMEOUT = 300  # Tiempo de espera para la actualización

def main():
    # Terminar instancias de Power BI en ejecución
    procname = "PBIDesktop.exe"
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['name'] == procname:
            os.kill(proc.info['pid'], 9)
    time.sleep(3)

    # Iniciar Power BI y abrir el libro de trabajo
    print("Starting Power BI...")
    os.startfile(WORKBOOK)
    time.sleep(INIT_WAIT)

    # Conectar con Power BI usando pywinauto
    print("Identifying Power BI window...")
    try:
        app = Application(backend="uia").connect(path=procname, timeout=120)
        print("Application connected.")
    except Exception as e:
        print(f"Error connecting to Power BI: {e}")
        sys.exit(1)

    try:
        win = app.window(found_index=0)
        win.wait('ready', timeout=120)
        print("Power BI window is ready.")
    except Exception as e:
        print(f"Error finding Power BI window: {e}")
        sys.exit(1)

    # Refrescar
    print("Refreshing...")
    try:
        # Aquí asumimos que F5 es el atajo para refrescar; ajusta según sea necesario
        send_keys('{F5}')
        print("Refresh command sent, waiting...")
        time.sleep(REFRESH_TIMEOUT)  # Espera para el refresco, ajusta según necesidad
    except Exception as e:
        print(f"Error during refresh: {e}")
        sys.exit(1)

    # Guardar
    print("Saving...")
    try:
        send_keys('^s')  # Ctrl+S para guardar
        time.sleep(5)  # Espera para asegurar que el guardado se complete
    except Exception as e:
        print(f"Error during save: {e}")
        sys.exit(1)

    # Cerrar Power BI
    print("Closing Power BI...")
    try:
        win.close()
        time.sleep(5)  # Espera para asegurar que Power BI se haya cerrado
    except Exception as e:
        print(f"Error during close: {e}")
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
