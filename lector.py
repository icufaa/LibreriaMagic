import fitz  # PyMuPDF
import cv2
import numpy as np
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, Toplevel, simpledialog
from PIL import Image
import pandas as pd
import os

# Variables globales para los precios de fotocopia según la tabla
PRECIOS_PATH = "precios.xlsx"

# Cargar precios desde el archivo Excel si existe, de lo contrario usar precios por defecto
def cargar_precios():
    if os.path.exists(PRECIOS_PATH):
        df = pd.read_excel(PRECIOS_PATH, sheet_name=None)
        return df['publico'].set_index('cantidad').T.to_dict('list'), df['estudiante'].set_index('cantidad').T.to_dict('list')
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
                30: 60,
                80: 70
            }
        }

        precios_estudiante = {
            "simple": {
                4: 50,
                20: 50,
                80: 40
            },
            "doble": {
                4: 80,
                20: 80,
                100: 70
            }
        }
        return precios_publico, precios_estudiante

def guardar_precios(precios_publico, precios_estudiante):
    df_publico = pd.DataFrame(precios_publico).T.reset_index().rename(columns={'index': 'cantidad'})
    df_estudiante = pd.DataFrame(precios_estudiante).T.reset_index().rename(columns={'index': 'cantidad'})
    with pd.ExcelWriter(PRECIOS_PATH) as writer:
        df_publico.to_excel(writer, sheet_name='publico', index=False)
        df_estudiante.to_excel(writer, sheet_name='estudiante', index=False)

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

        editar_precios_window = Toplevel(root)
        editar_precios_window.title("Editar Precios")
        editar_precios_window.geometry("500x400")

        frame_publico = ttk.Frame(editar_precios_window, padding=10)
        frame_publico.pack(fill=BOTH, expand=True)
        label_publico = ttk.Label(frame_publico, text="Precios Público", font=("Arial", 12, "bold"))
        label_publico.pack(pady=5)

        entry_publico = {"simple": {}, "doble": {}}
        for tipo in precios_publico:
            label_tipo = ttk.Label(frame_publico, text=f"Tipo: {tipo.capitalize()}", font=("Arial", 10, "italic"))
            label_tipo.pack(pady=5)
            for key, value in precios_publico[tipo].items():
                frame_entry = ttk.Frame(frame_publico)
                frame_entry.pack(pady=2)
                label_key = ttk.Label(frame_entry, text=f"Cantidad {key}: ")
                label_key.pack(side=LEFT)
                entry = ttk.Entry(frame_entry)
                entry.insert(0, str(value))
                entry.pack(side=LEFT)
                entry_publico[tipo][key] = entry

        frame_estudiante = ttk.Frame(editar_precios_window, padding=10)
        frame_estudiante.pack(fill=BOTH, expand=True)
        label_estudiante = ttk.Label(frame_estudiante, text="Precios Estudiante", font=("Arial", 12, "bold"))
        label_estudiante.pack(pady=5)

        entry_estudiante = {"simple": {}, "doble": {}}
        for tipo in precios_estudiante:
            label_tipo = ttk.Label(frame_estudiante, text=f"Tipo: {tipo.capitalize()}", font=("Arial", 10, "italic"))
            label_tipo.pack(pady=5)
            for key, value in precios_estudiante[tipo].items():
                frame_entry = ttk.Frame(frame_estudiante)
                frame_entry.pack(pady=2)
                label_key = ttk.Label(frame_entry, text=f"Cantidad {key}: ")
                label_key.pack(side=LEFT)
                entry = ttk.Entry(frame_entry)
                entry.insert(0, str(value))
                entry.pack(side=LEFT)
                entry_estudiante[tipo][key] = entry

        btn_guardar_precios = ttk.Button(editar_precios_window, text="Guardar Precios", command=guardar_precios_editar, bootstyle="success")
        btn_guardar_precios.pack(pady=10)

    sensitivity = 60

    root = ttk.Window(themename="darkly")
    root.title("Calculadora de Archivos")
    root.geometry("700x600")

    # Variables
    global rutas_var, opcion_doble_faz, opcion_usuario, opcion_color
    rutas_var = ttk.StringVar()
    opcion_doble_faz = ttk.IntVar(value=0)
    opcion_usuario = ttk.IntVar(value=0)
    opcion_color = ttk.IntVar(value=0)

    # Menú
    menu = ttk.Menu(root)
    root.config(menu=menu)
    archivo_menu = ttk.Menu(menu, tearoff=0)
    menu.add_cascade(label="Archivo", menu=archivo_menu)
    archivo_menu.add_command(label="Salir", command=salir)
    acciones_menu = ttk.Menu(menu, tearoff=0)
    menu.add_cascade(label='Acciones', menu=acciones_menu)
    acciones_menu.add_command(label='Calcular Precio', command=calcular)
    menu.add_command(label='Ajustes', command=mostrar_ventana_ajustes)
    menu.add_command(label='Editar Precios', command=mostrar_ventana_editar_precios)

    # Sección de Selección de Archivos
    frame_seleccion = ttk.Frame(root, padding=20)
    frame_seleccion.pack(fill=BOTH, expand=True)

    label_instruccion = ttk.Label(frame_seleccion, text="Selecciona uno o más archivos PDF y calcula el costo de las fotocopias.", font=("Arial", 14))
    label_instruccion.pack(pady=10)

    btn_seleccionar = ttk.Button(frame_seleccion, text="Seleccionar Archivos PDF", command=seleccionar_archivos, bootstyle="info-outline")
    btn_seleccionar.pack(pady=10)

    label_rutas = ttk.Label(frame_seleccion, textvariable=rutas_var, wraplength=600, font=("Arial", 12))
    label_rutas.pack(pady=10)

    # Sección de Opciones
    frame_opciones = ttk.Frame(frame_seleccion, padding=10)
    frame_opciones.pack(pady=10)

    check_doble_faz = ttk.Checkbutton(frame_opciones, text="Doble Faz", variable=opcion_doble_faz, bootstyle="success")
    check_doble_faz.pack(side=LEFT, padx=10)

    check_usuario = ttk.Checkbutton(frame_opciones, text="Estudiante", variable=opcion_usuario, bootstyle="success")
    check_usuario.pack(side=LEFT, padx=10)

    check_color = ttk.Checkbutton(frame_opciones, text="Color", variable=opcion_color, bootstyle="success")
    check_color.pack(side=LEFT, padx=10)

    # Botón de Calcular Precio
    frame_calcular = ttk.Frame(frame_seleccion, padding=20)
    frame_calcular.pack(pady=10)

    btn_calcular = ttk.Button(frame_calcular, text="Calcular Precio", command=calcular, bootstyle="primary")
    btn_calcular.pack()

    root.mainloop()

if __name__ == '__main__':
    menu_interactivo()
