import os
import sys
import time
import zipfile
import shutil
import threading
from pathlib import Path
import subprocess
import configparser
from tkinter import filedialog, messagebox, scrolledtext
import tkinter as tk
from PIL import Image, ImageTk
import logging
from datetime import datetime

# Configuración del logging para guardar los errores en un archivo de log
logging.basicConfig(filename='impresion_etiquetas.log', level=logging.ERROR)

# Función para obtener la ruta correcta del archivo (para PyInstaller y en desarrollo)
def obtener_ruta_absoluta(relative_path):
    """Obtener la ruta absoluta para PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Archivo de configuración para guardar la ruta de Adobe Acrobat Reader
config_file = "config.ini"

# Función para cargar la ruta de Adobe Acrobat Reader desde el archivo de configuración
def cargar_ruta_adobe():
    config = configparser.ConfigParser()
    if os.path.exists(config_file):
        config.read(config_file)
        if 'Config' in config and 'adobe_path' in config['Config']:
            return config['Config']['adobe_path']
    return None

# Función para guardar la ruta de Adobe Acrobat Reader en el archivo de configuración
def guardar_ruta_adobe(ruta):
    config = configparser.ConfigParser()
    config['Config'] = {'adobe_path': ruta}
    with open(config_file, 'w') as configfile:
        config.write(configfile)

# Función para pedir la ruta de Adobe Acrobat Reader al usuario y guardarla
def obtener_ruta_adobe():
    ruta_adobe = cargar_ruta_adobe()
    if ruta_adobe and os.path.exists(ruta_adobe):
        return ruta_adobe
    else:
        messagebox.showwarning("Adobe Acrobat Reader no encontrado", "Seleccione la ubicación de AcroRd32.exe.")
        ruta_adobe = filedialog.askopenfilename(title="Seleccionar Adobe Acrobat Reader", filetypes=[("Ejecutable", "*.exe")])
        if ruta_adobe and os.path.exists(ruta_adobe):
            guardar_ruta_adobe(ruta_adobe)
            return ruta_adobe
        else:
            messagebox.showerror("Error", "No se encontró Adobe Acrobat Reader. La aplicación no puede continuar.")
            return None

# Clase principal para la aplicación
class Aplicacion:
    def __init__(self, root):
        self.root = root
        self.root.title("Atmósfera Sport - Monitoreo e Impresión")
        self.root.geometry("500x500")
        self.root.resizable(False, False)
        self.logo_tk = None
        self.crear_primera_ventana()

    def crear_primera_ventana(self):
        try:
            ruta_logo = obtener_ruta_absoluta('logo_atmosfera_sport.jpg')
            self.logo = Image.open(ruta_logo)
            self.logo = self.logo.resize((400, 100))
            self.logo_tk = ImageTk.PhotoImage(self.logo)
            self.logo_label = tk.Label(self.root, image=self.logo_tk)
            self.logo_label.pack(pady=10)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el logo. Detalles: {e}")

        self.label_carpeta = tk.Label(self.root, text="Seleccionar la carpeta a monitorear:", font=("Arial", 12))
        self.label_carpeta.pack(pady=5)
        self.boton_carpeta = tk.Button(self.root, text="Seleccionar carpeta", command=self.seleccionar_carpeta)
        self.boton_carpeta.pack(pady=5)

        self.label_impresora = tk.Label(self.root, text="Escriba el nombre de la impresora:", font=("Arial", 12))
        self.label_impresora.pack(pady=5)
        self.entry_impresora = tk.Entry(self.root, width=40)
        self.entry_impresora.pack(pady=5)

        self.boton_confirmar = tk.Button(self.root, text="Confirmar", command=self.abrir_segunda_ventana)
        self.boton_confirmar.pack(pady=10)

    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory()
        if carpeta:
            self.carpeta_seleccionada = carpeta
            messagebox.showinfo("Carpeta seleccionada", f"Has seleccionado: {carpeta}")
        else:
            self.carpeta_seleccionada = None

    def abrir_segunda_ventana(self):
        if not hasattr(self, 'carpeta_seleccionada') or not self.carpeta_seleccionada:
            messagebox.showwarning("Error", "Debe seleccionar una carpeta.")
            return
        impresora = self.entry_impresora.get().strip()
        if not impresora:
            messagebox.showwarning("Error", "Debe ingresar el nombre de la impresora.")
            return

        adobe_path = obtener_ruta_adobe()
        if not adobe_path:
            return

        self.root.withdraw()

        self.ventana_logs = tk.Toplevel()
        self.ventana_logs.title("Atmósfera Sport - Estado de la impresión")
        self.ventana_logs.geometry("500x500")
        self.ventana_logs.resizable(False, False)

        if self.logo_tk:
            self.logo_label_logs = tk.Label(self.ventana_logs, image=self.logo_tk)
            self.logo_label_logs.pack(pady=10)

        self.log_text = scrolledtext.ScrolledText(self.ventana_logs, width=60, height=15)
        self.log_text.pack(pady=10)
        self.ventana_logs.protocol("WM_DELETE_WINDOW", self.cerrar_aplicacion)

        self.mostrar_logs("Aplicación iniciada...\n")
        self.mostrar_logs(f"Usando Adobe Acrobat Reader en: {adobe_path}")

        threading.Thread(target=self.iniciar_monitoreo, args=(self.carpeta_seleccionada, impresora, adobe_path), daemon=True).start()

    def mostrar_logs(self, texto):
        self.root.after(0, lambda: self.log_text.insert(tk.END, texto + "\n"))
        self.root.after(0, lambda: self.log_text.see(tk.END))

    def iniciar_monitoreo(self, carpeta, impresora, adobe_path):
        self.mostrar_logs(f"Monitoreando la carpeta: {carpeta}")
        carpeta_descargas = Path(carpeta)

        archivos_anteriores = set(os.listdir(carpeta_descargas))

        try:
            while True:
                archivos_actuales = set(os.listdir(carpeta_descargas))
                nuevos_archivos = archivos_actuales - archivos_anteriores

                for archivo in nuevos_archivos:
                    if archivo.startswith("Etiquetas - ") and archivo.endswith(".zip"):
                        ruta_zip = carpeta_descargas / archivo
                        self.mostrar_logs(f"Nuevo archivo ZIP detectado: {ruta_zip}")
                        self.manejar_zip(ruta_zip, impresora, adobe_path)

                archivos_anteriores = archivos_actuales
                time.sleep(5)
        except Exception as e:
            self.mostrar_logs(f"Error en el monitoreo: {e}")

    def manejar_zip(self, ruta_zip, impresora, adobe_path):
        carpeta_temporal = Path('temp_etiquetas')
        carpeta_impresos = carpeta_temporal / 'impresos'

        carpeta_temporal.mkdir(exist_ok=True)
        carpeta_impresos.mkdir(exist_ok=True)

        with zipfile.ZipFile(ruta_zip, 'r') as zip_ref:
            zip_ref.extractall(carpeta_temporal)

        self.mostrar_logs(f"Archivos extraídos en: {carpeta_temporal}")

        list_files = os.listdir(carpeta_temporal)
        list_pdfs = [carpeta_temporal / file for file in list_files if file.endswith('.pdf')]

        for pdf_file in list_pdfs:
            self.mostrar_logs(f"Enviando a imprimir: {pdf_file}")
            comando_impresion = f'"{adobe_path}" /t "{pdf_file}" "{impresora}"'
            subprocess.run(comando_impresion, shell=True)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nuevo_nombre = f"{pdf_file.stem}_{timestamp}.pdf"
            destino = carpeta_impresos / nuevo_nombre
            pdf_file.rename(destino)
            self.mostrar_logs(f"Archivo movido a 'impresos' como: {destino}")

    def cerrar_aplicacion(self):
        self.ventana_logs.destroy()
        self.root.quit()

# Iniciar la aplicación
if __name__ == "__main__":
    root = tk.Tk()
    app = Aplicacion(root)
    root.mainloop()
