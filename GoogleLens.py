import pandas as pd
import logging
from datetime import datetime
import time
import os
import sys
import requests
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from utils import initialize_driver, setup_logging, enviar_texto_a_input, esperar_y_clicar, elemento_visible

# Declarar variables globales
fichero_entry = None
root = None

def seleccionar_fichero():
    global fichero_entry
    file_dir = filedialog.askdirectory()
    if file_dir:
        fichero_entry.delete(0, tk.END)
        fichero_entry.insert(0, file_dir)

def save_page_as_pdf(driver, url, save_dir, filename):
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)  # Crea el directorio si no existe

    save_path = os.path.join(save_dir, f"{filename}.pdf")

    # Navegar a la URL
    driver.get(url)
    driver.execute_script('window.print();')

    # Esperar a que el PDF se guarde
    WebDriverWait(driver, 30).until(lambda d: os.path.isfile(save_path))

    logging.info(f"El PDF ha sido guardado en: {save_path}")

def save_url_as_html(url, save_path):
    try:
        # Realizar la solicitud HTTP para obtener el contenido de la URL
        response = requests.get(url)
        response.raise_for_status()  # Asegúrate de que la solicitud fue exitosa

        # Guardar el contenido en un archivo HTML
        with open(save_path, 'w', encoding='utf-8') as file:
            file.write(response.text)

        logging.info(f"La página ha sido guardada en: {save_path}")

    except requests.exceptions.RequestException as e:
        logging.info(f"Error al intentar descargar la página: {e}")


def iniciar_programa():
    global fichero_entry
    file_dir = fichero_entry.get()

    # Generar Dataframe
    df = pd.DataFrame(columns=['File', 'Result'])
    index = 0

    # Logging
    directorio_base = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    current_time = datetime.now()
    setup_logging(current_time)

    # Inicializamos driver Chrome
    driver = initialize_driver(directorio_base)
    url = 'https://www.google.com/'
    logging.info("Driver inicializado")

    for filename in os.listdir(file_dir):
        if filename.lower().endswith((".jpg", ".jpeg", ".png")):
            # Entramos en la web de google
            driver.get(url)
            driver.maximize_window()

            # Verificar si se visualiza la seleccion de busqueda predeterminada
            es_visible = elemento_visible(driver, By.ID, "actionButton")
            # Selecciona el botón "Google" como predeterminado
            if es_visible:
                esperar_y_clicar(driver, 'name', "1", 1, 'Clicando en la opción Google')
                esperar_y_clicar(driver, 'id', "actionButton", 1, 'Clicando en la opción Establecer Predeterminado')
            else:
                logging.info("La seleccion de busqueda predeterminada no es visible.")

            # Verificar si se visualiza el inicio de sesión
            es_visible = elemento_visible(driver, By.ID, "W0wltc")
            # Selecciona el botón "Rechazar Todo"
            if es_visible:
                esperar_y_clicar(driver, 'id', "W0wltc", 1, 'Clicando en la opción Rechazar Todo')
            else:
                logging.info("El inicio de sesión de Google no es visible.")

            # Vamos a la sección de Google Lens
            # esperar_y_clicar(driver, 'id', "lensSearchButton", 1, 'Clicando en la opción Buscar por Imagenes')
            esperar_y_clicar(driver, 'css', "div[jsname='R5mgy']", 1, 'Clicando en la opción Buscar por Imagenes')

            # Pinchamos en Subir Archivo
            # esperar_y_clicar(driver, 'id', 'uploadText', 1, 'Clicando en el link Subir Archivo')
            # esperar_y_clicar(driver, 'css', "span[jsname='tAPGc']", 1, 'Clicando en el link Subir Archivo')

            # Busca el input de tipo 'file' que permite seleccionar un archivo para cargar
            file_input = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
            )
            # Guardamos el resultado en el dataFrame
            file_path = os.path.join(file_dir, filename)
            # Sube el archivo al input de tipo 'file'
            file_input.send_keys(file_path)

            # Esperar a cargar resultado
            time.sleep(5)

            # Añadir datos al Dataframe
            df.loc[index] = [filename, driver.current_url]
            index += 1

            # Guardar la página como PDF
            # save_page_as_pdf(driver, driver.current_url, directorio_base, os.path.splitext(filename)[0])

            # Guardar la página como HTML
            html_filename = os.path.join(file_dir, os.path.splitext(filename)[0] + ".html")
            save_url_as_html(driver.current_url, html_filename)

    # Guardar el DataFrame actualizado 
    timestamp = current_time.strftime('%Y%m%d__%H%M%S')
    output_file_final = os.path.join(directorio_base, f'Search_{timestamp}.xlsx')
    df.to_excel(output_file_final, index=False)

    logging.info("Proceso finalizado")
    messagebox.showinfo("FIN","Proceso completado")

def main():
    try:
        global fichero_entry, root

        # Crear la ventana de Tkinter
        root = tk.Tk()
        root.title("Selector de Fichero Excel")
        root.geometry("500x200")

        # Crear un cuadro de texto para mostrar el fichero seleccionado
        fichero_entry = tk.Entry(root, width=50)
        fichero_entry.pack(pady=20)

        # Botón para seleccionar el fichero
        boton_seleccionar = tk.Button(root, text="Seleccionar Directorio", command=seleccionar_fichero)
        boton_seleccionar.pack(pady=10)

        # Botón para iniciar el programa
        boton_iniciar = tk.Button(root, text="Iniciar Programa", command=iniciar_programa)
        boton_iniciar.pack(pady=10)

        # Ejecutar la ventana de Tkinter
        root.mainloop()

    except Exception as e:
        logging.exception("Error inesperado en el script principal.")
        error_message = f"Ha ocurrido un error: {str(e)}"
        messagebox.showerror("Error", error_message)
        exit(1)

if __name__ == "__main__":
    main()
