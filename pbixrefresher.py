import time
import os
import sys
import argparse
import psutil
from pywinauto.application import Application


def type_keys(string, element):
    """Type a string char by char to Element window"""
    for char in string:
        element.type_keys(char)


def main():
    # Configuración del archivo de trabajo
    WORKBOOK = r"C:\Users\miguel\Desktop\sample.pbix"  # Ruta al archivo de trabajo

    # Otros parámetros de configuración
    INIT_WAIT = 15  # Tiempo de espera inicial
    REFRESH_TIMEOUT = 30000  # Tiempo de espera para la actualización

    # Kill running PBI
    PROCNAME = "PBIDesktop.exe"
    for proc in psutil.process_iter():
        # check whether the process name matches
        if proc.name() == PROCNAME:
            proc.kill()
    time.sleep(3)

    # Start PBI and open the workbook
    print("Starting Power BI")
    os.system('start "" "' + WORKBOOK + '"')
    print("Waiting ", INIT_WAIT, "sec")
    time.sleep(INIT_WAIT)

    # Connect pywinauto
    print("Identifying Power BI window")
    app = Application(backend='uia').connect(path=PROCNAME)
    win = app.window(title_re='.*Power BI Desktop')
    time.sleep(5)
    win.wait("enabled", timeout=300)
    win.Save.wait("enabled", timeout=300)
    win.set_focus()
    win.Home.click_input()
    win.Save.wait("enabled", timeout=300)
    win.wait("enabled", timeout=300)

    # Refresh
    print("Refreshing")
    win.Refresh.click_input()
    time.sleep(5)
    print("Waiting for refresh end (timeout in ", REFRESH_TIMEOUT, "sec)")

    try:
        win.wait("enabled", timeout=REFRESH_TIMEOUT)
    except TimeoutError:
        pass

    # Save
    print("Saving")
    type_keys("%1", win)
    time.sleep(5)
    win.wait("enabled", timeout=REFRESH_TIMEOUT)

    # Close
    print("Exiting")
    win.close()

    # Force close
    for proc in psutil.process_iter():
        if proc.name() == PROCNAME:
            proc.kill()


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(e)
        sys.exit(1)
