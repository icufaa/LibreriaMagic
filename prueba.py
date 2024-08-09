import fitz  # PyMuPDF
import cv2
import numpy as np
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, Toplevel, StringVar, IntVar, Label
from PIL import Image, ImageTk
import pandas as pd
import os
from tqdm import tqdm
from docx2pdf import convert as convert_docx
import tempfile
import shutil

# Variables globales para los precios de fotocopia según la tabla
PRECIOS_PATH = "precios.xlsx"
sensitivity = 50  # Valor predeterminado para la sensibilidad

# Cargar precios desde el archivo Excel si existe, de lo contrario usar precios por defecto
def cargar_precios():
    if os.path.exists(PRECIOS_PATH):
        df = pd.read_excel(PRECIOS_PATH, sheet_name=None)
        precios_publico = df['publico'].set_index('cantidad').T.fillna(0).to_dict('dict')
        precios_estudiante = df['estudiante'].set_index('cantidad').T.fillna(0).to_dict('dict')
        
        # Convertir los valores a enteros
        for tipo in precios_publico:
            precios_publico[tipo] = {int(key): int(value) for key, value in precios_publico[tipo].items()}
        for tipo in precios_estudiante:
            precios_estudiante[tipo] = {int(key): int(value) for key, value in precios_estudiante[tipo].items()}
            
        return precios_publico, precios_estudiante
    else:
        precios_publico = {
            "simple": {
                1: 100,
                2: 80,
                30: 60,
                80: 40
            },
            "doble": {
                1: 100,
                30: 80,
                80: 70
            }
        }

        precios_estudiante = {
            "simple": {
                1:100,
                10: 50,
                20: 50,
                80: 40,
                100: 40
            },
            "doble": {
                1:100,
                10: 80,
                20: 80,
                100: 70
            }
        }
        return precios_publico, precios_estudiante

def guardar_precios(precios_publico, precios_estudiante):
    df_publico = pd.DataFrame(precios_publico).T.reset_index().rename(columns={'index': 'cantidad'})
    df_estudiante = pd.DataFrame(precios_estudiante).T.reset_index().rename(columns={'index': 'cantidad'})
    try:
        with pd.ExcelWriter(PRECIOS_PATH) as writer:
            df_publico.to_excel(writer, sheet_name='publico', index=False)
            df_estudiante.to_excel(writer, sheet_name='estudiante', index=False)
        messagebox.showinfo("Éxito", "Precios actualizados con éxito")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudieron guardar los precios: {e}")

precios_publico, precios_estudiante = cargar_precios()

def obtener_precio(cantidad, tipo, usuario, porcentaje_color, color):
    precios = precios_estudiante if usuario == "estudiante" else precios_publico

    # Ajuste de precio según porcentaje de color
    precio_color = 100 + (porcentaje_color * 400)  # Precio base $100 y máximo $500
    if color:
        precio_color = precio_color   
        precio_color = min(precio_color, 500)  # Limitar a un máximo de $500
    else:
        precio_color = 0  # No hay ajuste por color si no es a color

    for limite, precio in sorted(precios[tipo].items(), reverse=True):
        if cantidad >= limite:
            return int(precio + precio_color)
    return 0

def obtener_porcentaje_color(pagina):
    # Convertir la página a una imagen de PIL
    pix = pagina.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    # Convertir la imagen de PIL a un arreglo de NumPy
    img_np = np.array(img)

    # Convertir la imagen a espacio de color HSV
    hsv = cv2.cvtColor(img_np, cv2.COLOR_RGB2HSV)

    # Definir umbral para detectar colores (excluyendo blanco y negro)
    lower_color = np.array([0, sensitivity, sensitivity])  # Usar sensibilidad global
    upper_color = np.array([179, 255, 255])  # Umbral máximo para colores

    # Crear una máscara para los colores
    mask = cv2.inRange(hsv, lower_color, upper_color)

    # Calcular el porcentaje de área coloreada
    color_percentage = np.sum(mask > 0) / mask.size

    return color_percentage

def calcular_precios(root, ruta_pdfs, doble_faz=False, usuario="publico", color=False):
    detalles_archivos = {}
    try:
        total_paginas = 0
        total_costo_archivos = 0

        # Crear la ventana de progreso
        progress_window = Toplevel(root)
        progress_window.title("Procesando PDFs")
        progress_window.geometry("400x100")

        # Crear la barra de progreso
        progress_bar = ttk.Progressbar(progress_window, length=300, mode='determinate')
        progress_bar.pack(pady=20)

        # Establecer el valor máximo de la barra de progreso
        progress_bar['maximum'] = len(ruta_pdfs)

        for i, ruta_pdf in enumerate(ruta_pdfs):
            doc = fitz.open(ruta_pdf)
            num_paginas = len(doc)
            total_paginas += num_paginas

            color_paginas = sum(obtener_porcentaje_color(pagina) for pagina in doc)
            porcentaje_color_promedio = color_paginas / num_paginas

            if doble_faz and not color:
                paginas_para_precio = (num_paginas + 1) // 2  # Redondear para arriba
                tipo = "doble"
            else:
                paginas_para_precio = num_paginas
                tipo = "simple"

            precio_fotocopia = obtener_precio(paginas_para_precio, tipo, usuario, porcentaje_color_promedio, color)
            costo_archivo = paginas_para_precio * precio_fotocopia
            total_costo_archivos += costo_archivo
            detalles_archivos[ruta_pdf] = (num_paginas, costo_archivo, precio_fotocopia)

            # Actualizar la barra de progreso
            progress_bar['value'] = i + 1
            progress_window.update()

        # Cerrar la ventana de progreso
        progress_window.destroy()

        return total_costo_archivos, detalles_archivos
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo PDF: {e}")
        return None, None

def convertir_a_pdf():
    archivo = filedialog.askopenfilename(filetypes=[("Todos los archivos", "*.*")])
    if archivo:
        pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Archivo PDF", "*.pdf")])
        if pdf_path:
            try:
                extension = os.path.splitext(archivo)[1].lower()
                temp_dir = tempfile.mkdtemp()
                if extension in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']:
                    imagen = Image.open(archivo)
                    imagen.convert("RGB").save(pdf_path)
                elif extension in ['.txt']:
                    with open(archivo, 'r', encoding='utf-8') as f:
                        contenido = f.read()
                    doc = fitz.open()
                    pagina = doc.new_page()
                    pagina.insert_text((72, 72), contenido, fontsize=12)
                    doc.save(pdf_path)
                elif extension in ['.pdf']:
                    shutil.copyfile(archivo, pdf_path)
                elif extension in ['.doc', '.docx']:
                    temp_pdf = os.path.join(temp_dir, "temp.pdf")
                    convert_docx(archivo, temp_pdf)
                    shutil.copyfile(temp_pdf, pdf_path)
                else:
                    messagebox.showwarning("Advertencia", "Tipo de archivo no soportado para conversión.")
                    shutil.rmtree(temp_dir)
                    return
                messagebox.showinfo("Éxito", f"El archivo {os.path.basename(archivo)} se ha convertido a PDF.")
                shutil.rmtree(temp_dir)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo convertir el archivo a PDF: {e}")

def menu_interactivo():
    def salir():
        root.destroy()

    def seleccionar_archivos():
        rutas_pdfs = filedialog.askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])
        if rutas_pdfs:
            rutas_var.set(", ".join(rutas_pdfs))

    def mostrar_ventana_detalles(total_copias, detalles_archivos, opciones_seleccionadas):
        detalles_window = Toplevel(root)
        detalles_window.title("Detalles de Archivos")
        detalles_window.geometry("600x400")

        canvas = ttk.Canvas(detalles_window)
        scrollbar = ttk.Scrollbar(detalles_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        canvas.configure(yscrollcommand=scrollbar.set)

        ttk.Label(scrollable_frame, text=f"Total de fotocopias: {total_copias}", font=("Helvetica", 12, "bold")).pack(pady=5)
        ttk.Label(scrollable_frame, text=f"Opción de usuario: {opciones_seleccionadas['usuario']}", font=("Helvetica", 12, "bold")).pack(pady=5)
        ttk.Label(scrollable_frame, text=f"Opción de color: {opciones_seleccionadas['color']}", font=("Helvetica", 12, "bold")).pack(pady=5)
        ttk.Label(scrollable_frame, text=f"Opción de doble faz: {opciones_seleccionadas['doble_faz']}", font=("Helvetica", 12, "bold")).pack(pady=5)

        for archivo, (paginas, costo, precio_fotocopia) in detalles_archivos.items():
            ttk.Label(scrollable_frame, text=f"{archivo} - Páginas: {paginas} - Precio por fotocopia: ${precio_fotocopia} - Costo Total: ${costo}", font=("Helvetica", 10)).pack(pady=2)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def calcular_y_mostrar_precios():
        rutas_pdfs = rutas_var.get().split(", ")
        usuario = usuario_var.get()
        color = color_var.get()
        doble_faz = doble_faz_var.get()
        if rutas_pdfs:
            total_costo, detalles_archivos = calcular_precios(root, rutas_pdfs, doble_faz, usuario, color)
            if total_costo is not None:
                opciones_seleccionadas = {
                    'usuario': usuario,
                    'color': 'Sí' if color else 'No',
                    'doble_faz': 'Sí' if doble_faz else 'No'
                }
                mostrar_ventana_detalles(total_costo, detalles_archivos, opciones_seleccionadas)

    root = ttk.Window(themename="superhero")
    root.title("Calculadora de Precios de Fotocopias")
    root.geometry("500x400")

    rutas_var = StringVar()
    usuario_var = StringVar(value="publico")
    color_var = IntVar(value=0)
    doble_faz_var = IntVar(value=0)
    sensitivity_var = IntVar(value=sensitivity)

    ttk.Label(root, text="Seleccionar archivos PDF:", font=("Helvetica", 12)).pack(pady=10)
    ttk.Button(root, text="Seleccionar archivos", command=seleccionar_archivos).pack()
    ttk.Label(root, textvariable=rutas_var, wraplength=400).pack(pady=10)

    ttk.Label(root, text="Seleccionar opciones:", font=("Helvetica", 12)).pack(pady=10)
    opciones_frame = ttk.Frame(root)
    opciones_frame.pack(pady=5)

    ttk.Radiobutton(opciones_frame, text="Público", variable=usuario_var, value="publico").grid(row=0, column=0, padx=5, pady=5)
    ttk.Radiobutton(opciones_frame, text="Estudiante", variable=usuario_var, value="estudiante").grid(row=0, column=1, padx=5, pady=5)
    ttk.Checkbutton(opciones_frame, text="Doble faz", variable=doble_faz_var).grid(row=1, column=0, padx=5, pady=5)
    ttk.Checkbutton(opciones_frame, text="Color", variable=color_var).grid(row=1, column=1, padx=5, pady=5)

    ttk.Label(opciones_frame, text="Sensibilidad de color:", font=("Helvetica", 10)).grid(row=2, column=0, padx=5, pady=5)
    sensitivity_scale = ttk.Scale(opciones_frame, from_=0, to_=255, variable=sensitivity_var)
    sensitivity_scale.set(sensitivity)
    sensitivity_scale.grid(row=2, column=1, padx=5, pady=5)

    ttk.Button(root, text="Calcular precios", command=calcular_y_mostrar_precios).pack(pady=10)
    ttk.Button(root, text="Salir", command=salir).pack(pady=10)
    ttk.Button(root, text="Convertir a PDF", command=convertir_a_pdf).pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    menu_interactivo()
