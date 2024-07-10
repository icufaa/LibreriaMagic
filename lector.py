import fitz  # PyMuPDF
import cv2
import numpy as np
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, Toplevel, simpledialog
from tkinter import Canvas, Scrollbar, Frame
from PIL import Image
import pandas as pd
import os

# Variables globales para los precios de fotocopia según la tabla
PRECIOS_PATH = "precios.xlsx"

# Cargar precios desde el archivo Excel si existe, de lo contrario usar precios por defecto
def cargar_precios():
    if os.path.exists(PRECIOS_PATH):
        df = pd.read_excel(PRECIOS_PATH, sheet_name=None)
        precios_publico = df['publico'].set_index('cantidad').T.to_dict('list')
        precios_estudiante = df['estudiante'].set_index('cantidad').T.to_dict('list')
        
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
                10: 50,
                20: 50,
                80: 40,
                100: 40
            },
            "doble": {
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

    ajuste_color = 1.0 + (4.0 * porcentaje_color) if color else 1.0  # 1.0 para B/N, hasta 5.0 para 100% color

    for limite, precio in sorted(precios[tipo].items(), reverse=True):
        if cantidad >= limite:
            return int(precio * ajuste_color)
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
    sensitivity = 60
    lower_color = np.array([0, 0, 0])
    upper_color = np.array([179, 255, 255])

    # Crear una máscara para los colores
    mask = cv2.inRange(hsv, lower_color, upper_color)

    # Calcular el porcentaje de área coloreada
    color_percentage = np.sum(mask > sensitivity) / mask.size

    return color_percentage

def calcular_precios(ruta_pdfs, doble_faz=False, usuario="publico", color=False):
    detalles_archivos = {}
    try:
        total_paginas = 0
        total_color = 0
        total_costo_archivos = 0
        for ruta_pdf in ruta_pdfs:
            doc = fitz.open(ruta_pdf)
            num_paginas = len(doc)
            total_paginas += num_paginas

            color_paginas = sum(obtener_porcentaje_color(pagina) for pagina in doc)
            porcentaje_color_promedio = color_paginas / num_paginas

            if doble_faz:
                paginas_para_precio = (num_paginas + 1) // 2  # Redondear para arriba
                paginas_para_precio += paginas_para_precio % 2  # Asegurar que sea par
                tipo = "doble"
            else:
                paginas_para_precio = num_paginas
                tipo = "simple"

            precio_fotocopia = obtener_precio(paginas_para_precio, tipo, usuario, porcentaje_color_promedio, color)
            costo_archivo = paginas_para_precio * precio_fotocopia
            total_costo_archivos += costo_archivo
            detalles_archivos[ruta_pdf] = (num_paginas, costo_archivo)

        return total_costo_archivos, detalles_archivos
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo PDF: {e}")
        return None, None

def menu_interactivo():
    def salir():
        root.destroy()

    def seleccionar_archivos():
        rutas_pdfs = filedialog.askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])
        if rutas_pdfs:
            rutas_var.set(", ".join(rutas_pdfs))

    def calcular():
        rutas_pdfs = rutas_var.get().split(", ")
        if rutas_pdfs:
            doble_faz = opcion_doble_faz.get() == 1
            usuario = "estudiante" if opcion_usuario.get() == 1 else "publico"
            color = opcion_color.get() == 1
            total_copias, detalles_archivos = calcular_precios(rutas_pdfs, doble_faz, usuario, color)
            if total_copias is not None:
                opciones_seleccionadas = (
                    f"Tipo de usuario: {'Estudiante' if usuario == 'estudiante' else 'Público'}\n"
                    f"Tipo de fotocopia: {'Doble faz' if doble_faz else 'Simple'}\n"
                    f"Color: {'Sí' if color else 'No'}\n"
                )
                detalle_mensaje = "\n\n".join([f'Archivo: {ruta}\nHojas: {hojas}\nPrecio: ${precio}' 
                                               for ruta, (hojas, precio) in detalles_archivos.items()])
                messagebox.showinfo("Resultado", f'{opciones_seleccionadas}\nEl costo total de las fotocopias es: ${total_copias}\n\nDetalles:\n{detalle_mensaje}')
        else:
            messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un archivo PDF.")

    def mostrar_ventana_ajustes():
        def guardar_ajustes():
            nonlocal sensitivity
            sensitivity = int(entry_sensitivity.get())
            ajustes_window.destroy()
        
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
                # Actualizar precios
                for tipo in precios_publico:
                    for key in precios_publico[tipo]:
                        nuevo_precio = int(entry_publico[tipo][key].get())
                        precios_publico[tipo][key] = nuevo_precio
                for tipo in precios_estudiante:
                    for key in precios_estudiante[tipo]:
                        nuevo_precio = int(entry_estudiante[tipo][key].get())
                        precios_estudiante[tipo][key] = nuevo_precio
                # Guardar precios en archivo Excel
                guardar_precios(precios_publico, precios_estudiante)
                editar_precios_window.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron actualizar los precios: {e}")

        editar_precios_window = Toplevel(root)
        editar_precios_window.title("Editar Precios")
        editar_precios_window.geometry("500x400")

        canvas = Canvas(editar_precios_window)
        scrollbar = Scrollbar(editar_precios_window, orient="vertical", command=canvas.yview)
        scrollable_frame = Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        frame_publico = ttk.Frame(scrollable_frame, padding=10)
        frame_publico.pack(fill=BOTH, expand=True)
        label_publico = ttk.Label(frame_publico, text="Precios Público", font=("Arial", 12, "bold"))
        label_publico.pack(pady=5)

        entry_publico = {"simple": {}, "doble": {}}
        for tipo in precios_publico:
            label_tipo = ttk.Label(frame_publico, text=f"Tipo: {tipo.capitalize()}", font=("Arial", 10, "bold"))
            label_tipo.pack(pady=5)
            for key in precios_publico[tipo]:
                frame_precio = ttk.Frame(frame_publico)
                frame_precio.pack(fill=X, pady=2)
                label_key = ttk.Label(frame_precio, text=f"Cantidad mínima: {key}", font=("Arial", 10))
                label_key.pack(side=LEFT, padx=5)
                entry_precio = ttk.Entry(frame_precio)
                entry_precio.pack(side=LEFT, padx=5)
                entry_precio.insert(0, precios_publico[tipo][key])
                entry_publico[tipo][key] = entry_precio

        frame_estudiante = ttk.Frame(scrollable_frame, padding=10)
        frame_estudiante.pack(fill=BOTH, expand=True)
        label_estudiante = ttk.Label(frame_estudiante, text="Precios Estudiante", font=("Arial", 12, "bold"))
        label_estudiante.pack(pady=5)

        entry_estudiante = {"simple": {}, "doble": {}}
        for tipo in precios_estudiante:
            label_tipo = ttk.Label(frame_estudiante, text=f"Tipo: {tipo.capitalize()}", font=("Arial", 10, "bold"))
            label_tipo.pack(pady=5)
            for key in precios_estudiante[tipo]:
                frame_precio = ttk.Frame(frame_estudiante)
                frame_precio.pack(fill=X, pady=2)
                label_key = ttk.Label(frame_precio, text=f"Cantidad mínima: {key}", font=("Arial", 10))
                label_key.pack(side=LEFT, padx=5)
                entry_precio = ttk.Entry(frame_precio)
                entry_precio.pack(side=LEFT, padx=5)
                entry_precio.insert(0, precios_estudiante[tipo][key])
                entry_estudiante[tipo][key] = entry_precio

        btn_guardar = ttk.Button(scrollable_frame, text="Guardar Precios", command=guardar_precios_editar, bootstyle="success")
        btn_guardar.pack(pady=10)

    sensitivity = 60

    root = ttk.Window(themename="darkly")
    root.title("Calculadora de Precios de Fotocopias")
    root.geometry("600x400")

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill=BOTH, expand=True)

    label_titulo = ttk.Label(frame, text="Calculadora de Precios de Fotocopias", font=("Arial", 16, "bold"))
    label_titulo.pack(pady=10)

    rutas_var = ttk.StringVar()
    entry_rutas = ttk.Entry(frame, textvariable=rutas_var, state="readonly")
    entry_rutas.pack(fill=X, pady=5)

    btn_seleccionar = ttk.Button(frame, text="Seleccionar Archivos", command=seleccionar_archivos, bootstyle="primary")
    btn_seleccionar.pack(pady=5)

    opcion_doble_faz = ttk.IntVar()
    chk_doble_faz = ttk.Checkbutton(frame, text="Doble faz", variable=opcion_doble_faz, bootstyle="success")
    chk_doble_faz.pack(pady=5)

    opcion_usuario = ttk.IntVar()
    chk_estudiante = ttk.Checkbutton(frame, text="Estudiante", variable=opcion_usuario, bootstyle="success")
    chk_estudiante.pack(pady=5)

    opcion_color = ttk.IntVar()
    chk_color = ttk.Checkbutton(frame, text="Color", variable=opcion_color, bootstyle="success")
    chk_color.pack(pady=5)

    btn_calcular = ttk.Button(frame, text="Calcular Precios", command=calcular, bootstyle="primary")
    btn_calcular.pack(pady=10)

    btn_ajustes = ttk.Button(frame, text="Ajustes Avanzados", command=mostrar_ventana_ajustes, bootstyle="info")
    btn_ajustes.pack(pady=5)

    btn_editar_precios = ttk.Button(frame, text="Editar Precios", command=mostrar_ventana_editar_precios, bootstyle="info")
    btn_editar_precios.pack(pady=5)

    btn_salir = ttk.Button(frame, text="Salir", command=salir, bootstyle="danger")
    btn_salir.pack(pady=10)

    root.mainloop()

menu_interactivo()
