import fitz  # PyMuPDF
import cv2
import numpy as np
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, Toplevel, StringVar, IntVar
from PIL import Image
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
                1:80,
                6:70,
                20:60,
                50: 50,
                100: 40
            },
            "doble": {
                1:100,
                30: 80,
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
    lower_color = np.array([sensitivity, sensitivity, sensitivity])  # Usar sensibilidad global
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
        color_paginas = 0

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

            color_paginas += sum(obtener_porcentaje_color(pagina) for pagina in doc)

            # Actualizar la barra de progreso
            progress_bar['value'] = i + 1
            progress_window.update()

        porcentaje_color_promedio = color_paginas / total_paginas

        if doble_faz and not color:
            paginas_para_precio = (total_paginas + 1) // 2  # Redondear para arriba
            tipo = "doble"
        else:
            paginas_para_precio = total_paginas
            tipo = "simple"

        precio_fotocopia = obtener_precio(paginas_para_precio, tipo, usuario, porcentaje_color_promedio, color)
        total_costo_archivos = paginas_para_precio * precio_fotocopia

        for ruta_pdf in ruta_pdfs:
            detalles_archivos[ruta_pdf] = (num_paginas, total_costo_archivos)

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
                messagebox.showinfo("Éxito", f"El archivzo {os.path.basename(archivo)} se ha convertido a PDF.")
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

        ttk.Label(scrollable_frame, text="Opciones seleccionadas:", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5)
        ttk.Label(scrollable_frame, text=opciones_seleccionadas).grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        ttk.Label(scrollable_frame, text=f"Costo total de las fotocopias: ${total_copias}", font=("Arial", 12, "bold")).grid(row=2, column=0, columnspan=2, padx=5, pady=5)

        row = 3
        for ruta, (hojas, precio) in detalles_archivos.items():
            ttk.Label(scrollable_frame, text=f"Archivo: {os.path.basename(ruta)}", font=("Arial", 12, "bold")).grid(row=row, column=0, padx=5, pady=5, sticky="w")
            row += 1
            ttk.Label(scrollable_frame, text=f"Hojas: {hojas}").grid(row=row, column=0, padx=5, pady=5, sticky="w")
            ttk.Label(scrollable_frame, text=f"Precio: ${precio}").grid(row=row, column=1, padx=5, pady=5, sticky="w")
            row += 1

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    
    def calcular():
        rutas_pdfs = rutas_var.get().split(", ")
        if rutas_pdfs:
            doble_faz = opcion_doble_faz.get() == 1
            usuario = "estudiante" if opcion_usuario.get() == 1 else "publico"
            color = opcion_color.get() == 1
            total_copias, detalles_archivos = calcular_precios(root, rutas_pdfs, doble_faz, usuario, color)
            if total_copias is not None:
                opciones_seleccionadas = (
                    f"Tipo de usuario: {'Estudiante' if usuario == 'estudiante' else 'Público'}\n"
                    f"Tipo de fotocopia: {'Doble faz' if doble_faz else 'Simple'}\n"
                    f"Color: {'Sí' if color else 'No'}\n"
                )
                mostrar_ventana_detalles(total_copias, detalles_archivos, opciones_seleccionadas)
        else:
            messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un archivo PDF.")

    def mostrar_ventana_ajustes():
        def guardar_ajustes():
            global sensitivity  # Declarar sensibilidad como global
            try:
                nueva_sensibilidad = int(entry_sensitivity.get())
                if 0 <= nueva_sensibilidad <= 255:
                    sensitivity = nueva_sensibilidad
                    ajustes_window.destroy()
                else:
                    messagebox.showerror("Error", "El umbral de sensibilidad debe estar entre 0 y 255.")
            except ValueError:
                messagebox.showerror("Error", "El umbral de sensibilidad debe ser un número entero.")

        ajustes_window = Toplevel(root)
        ajustes_window.title("Ajustes Avanzados")
        ajustes_window.geometry("300x200")

        label_sensitivity = ttk.Label(ajustes_window, text="Umbral de Sensibilidad para Detección de Color:", font=("Arial", 12))
        label_sensitivity.pack(pady=10)

        entry_sensitivity = ttk.Entry(ajustes_window)
        entry_sensitivity.pack(pady=10)
        entry_sensitivity.insert(0, sensitivity)

        btn_guardar = ttk.Button(ajustes_window, text="Guardar", command=guardar_ajustes, bootstyle="success")
        btn_guardar.pack(pady=10)

    def mostrar_ventana_editar_precios():
        def guardar_precios_editar():
            try:
                for tipo in precios_publico:
                    for key in precios_publico[tipo]:
                        nuevo_precio = int(entry_publico[tipo][key].get())
                        precios_publico[tipo][key] = nuevo_precio

                for tipo in precios_estudiante:
                    for key in precios_estudiante[tipo]:
                        nuevo_precio = int(entry_estudiante[tipo][key].get())
                        precios_estudiante[tipo][key] = nuevo_precio

                guardar_precios(precios_publico, precios_estudiante)
                editar_precios_window.destroy()
            except ValueError:
                messagebox.showerror("Error", "Todos los campos deben ser números enteros.")

        editar_precios_window = Toplevel(root)
        editar_precios_window.title("Editar Precios")
        editar_precios_window.geometry("700x500")

        frame_publico = ttk.Frame(editar_precios_window)
        frame_publico.pack(side="left", padx=20, pady=20, fill="y", expand=True)

        frame_estudiante = ttk.Frame(editar_precios_window)
        frame_estudiante.pack(side="right", padx=20, pady=20, fill="y", expand=True)

        ttk.Label(frame_publico, text="Público - Simple Faz", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=2, pady=10)
        ttk.Label(frame_estudiante, text="Estudiante - Simple Faz", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=2, pady=10)

        entry_publico = {"simple": {}, "doble": {}}
        entry_estudiante = {"simple": {}, "doble": {}}

        row = 1
        for tipo in precios_publico:
            for key in sorted(precios_publico[tipo].keys()):
                if tipo == "doble":
                    ttk.Label(frame_publico, text="Público - Doble Faz", font=("Arial", 12, "bold")).grid(row=row, column=0, columnspan=2, pady=10)
                    row += 1
                ttk.Label(frame_publico, text=f"{key} hojas:").grid(row=row, column=0, padx=5, pady=5, sticky="e")
                entry = ttk.Entry(frame_publico, width=10)
                entry.grid(row=row, column=1, padx=5, pady=5, sticky="w")
                entry.insert(0, precios_publico[tipo][key])
                entry_publico[tipo][key] = entry
                row += 1

        row = 1
        for tipo in precios_estudiante:
            for key in sorted(precios_estudiante[tipo].keys()):
                if tipo == "doble":
                    ttk.Label(frame_estudiante, text="Estudiante - Doble Faz", font=("Arial", 12, "bold")).grid(row=row, column=0, columnspan=2, pady=10)
                    row += 1
                ttk.Label(frame_estudiante, text=f"{key} hojas:").grid(row=row, column=0, padx=5, pady=5, sticky="e")
                entry = ttk.Entry(frame_estudiante, width=10)
                entry.grid(row=row, column=1, padx=5, pady=5, sticky="w")
                entry.insert(0, precios_estudiante[tipo][key])
                entry_estudiante[tipo][key] = entry
                row += 1

        btn_guardar = ttk.Button(editar_precios_window, text="Guardar", command=guardar_precios_editar, bootstyle="success")
        btn_guardar.pack(pady=20)


    root = ttk.Window(themename="darkly")
    root.title("Calculadora de Fotocopias")
    root.geometry("800x600")


    marco_superior = ttk.Frame(root)
    marco_superior.pack(pady=20)

    etiqueta_titulo = ttk.Label(marco_superior, text="Calculadora de Fotocopias", font=("Arial", 24, "bold"))
    etiqueta_titulo.pack()

    marco_opciones = ttk.Frame(root)
    marco_opciones.pack(pady=10)

    ttk.Label(marco_opciones, text="Archivos seleccionados:").grid(row=0, column=0, sticky="w")
    rutas_var = StringVar()
    entry_rutas = ttk.Entry(marco_opciones, textvariable=rutas_var, width=60, state='readonly')
    entry_rutas.grid(row=1, column=0, padx=5, pady=5, sticky="we")

    btn_seleccionar = ttk.Button(marco_opciones, text="Seleccionar Archivos", command=seleccionar_archivos)
    btn_seleccionar.grid(row=1, column=1, padx=5, pady=5)

    marco_opciones_extra = ttk.Frame(root)
    marco_opciones_extra.pack(pady=10)

    opcion_doble_faz = IntVar()
    chk_doble_faz = ttk.Checkbutton(marco_opciones_extra, text="Doble Faz", variable=opcion_doble_faz)
    chk_doble_faz.grid(row=0, column=0, padx=10, pady=5)

    opcion_usuario = IntVar()
    rad_publico = ttk.Radiobutton(marco_opciones_extra, text="Público", variable=opcion_usuario, value=0)
    rad_publico.grid(row=0, column=1, padx=10, pady=5)
    rad_estudiante = ttk.Radiobutton(marco_opciones_extra, text="Estudiante", variable=opcion_usuario, value=1)
    rad_estudiante.grid(row=0, column=2, padx=10, pady=5)

    opcion_color = IntVar()
    chk_color = ttk.Checkbutton(marco_opciones_extra, text="Color", variable=opcion_color)
    chk_color.grid(row=0, column=3, padx=10, pady=5)


    btn_calcular = ttk.Button(root, text="Calcular", command=calcular, bootstyle="primary")
    btn_calcular.pack(pady=10)

    btn_convertir_pdf = ttk.Button(root, text="Convertir a PDF", command=convertir_a_pdf, bootstyle="secondary")
    btn_convertir_pdf.pack(pady=5)

    btn_ajustes = ttk.Button(root, text="Ajustes Avanzados", command=mostrar_ventana_ajustes, bootstyle="secondary")
    btn_ajustes.pack(pady=5)

    btn_editar_precios = ttk.Button(root, text="Editar Precios", command=mostrar_ventana_editar_precios, bootstyle="secondary")
    btn_editar_precios.pack(pady=5)

    btn_salir = ttk.Button(root, text="Salir", command=salir, bootstyle="danger")
    btn_salir.pack(pady=10)

    root.mainloop()

menu_interactivo()