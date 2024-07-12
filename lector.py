import fitz  # PyMuPDF
import cv2
import numpy as np
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, Toplevel
from PIL import Image
import pandas as pd
import os

# Variables globales para los precios de fotocopia según la tabla
PRECIOS_PATH = "precios.xlsx"
sensitivity = 50  # Valor predeterminado para la sensibilidad

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
        if 450 <= precio_color < 500:
            precio_color = round(precio_color / 50) * 50  # Redondear al más cercano entre 450 y 500
        elif 400 <= precio_color < 450:
            precio_color = round(precio_color / 50) * 50  # Redondear al más cercano entre 400 y 450
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

def calcular_precios(ruta_pdfs, doble_faz=False, usuario="publico", color=False):
    detalles_archivos = {}
    try:
        total_paginas = 0
        total_costo_archivos = 0
        for ruta_pdf in ruta_pdfs:
            doc = fitz.open(ruta_pdf)
            num_paginas = len(doc)
            total_paginas += num_paginas

            color_paginas = sum(obtener_porcentaje_color(pagina) for pagina in doc)
            porcentaje_color_promedio = color_paginas / num_paginas

            if doble_faz:
                paginas_para_precio = (num_paginas + 1) // 2  # Redondear para arriba
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
            global sensitivity  # Declarar sensibilidad como global
            try:
                sensitivity = int(entry_sensitivity.get())
                ajustes_window.destroy()
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
        editar_precios_window.geometry("500x400")

        frame_publico = ttk.Frame(editar_precios_window)
        frame_publico.pack(side="left", padx=20, pady=20)

        frame_estudiante = ttk.Frame(editar_precios_window)
        frame_estudiante.pack(side="right", padx=20, pady=20)

        ttk.Label(frame_publico, text="Público", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=2)
        ttk.Label(frame_estudiante, text="Estudiante", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=2)

        entry_publico = {"simple": {}, "doble": {}}
        entry_estudiante = {"simple": {}, "doble": {}}

        row = 1
        for tipo in precios_publico:
            for key in sorted(precios_publico[tipo].keys()):
                ttk.Label(frame_publico, text=f"{key} hojas:").grid(row=row, column=0)
                entry = ttk.Entry(frame_publico)
                entry.grid(row=row, column=1)
                entry.insert(0, precios_publico[tipo][key])
                entry_publico[tipo][key] = entry
                row += 1

        row = 1
        for tipo in precios_estudiante:
            for key in sorted(precios_estudiante[tipo].keys()):
                ttk.Label(frame_estudiante, text=f"{key} hojas:").grid(row=row, column=0)
                entry = ttk.Entry(frame_estudiante)
                entry.grid(row=row, column=1)
                entry.insert(0, precios_estudiante[tipo][key])
                entry_estudiante[tipo][key] = entry
                row += 1

        btn_guardar = ttk.Button(editar_precios_window, text="Guardar", command=guardar_precios_editar, bootstyle="success")
        btn_guardar.pack(pady=10)

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
    rutas_var = ttk.StringVar()
    entry_rutas = ttk.Entry(marco_opciones, textvariable=rutas_var, width=60, state='readonly')
    entry_rutas.grid(row=1, column=0, padx=5, pady=5, sticky="we")

    btn_seleccionar = ttk.Button(marco_opciones, text="Seleccionar Archivos", command=seleccionar_archivos)
    btn_seleccionar.grid(row=1, column=1, padx=5, pady=5)

    marco_opciones_extra = ttk.Frame(root)
    marco_opciones_extra.pack(pady=10)

    opcion_doble_faz = ttk.IntVar()
    chk_doble_faz = ttk.Checkbutton(marco_opciones_extra, text="Doble Faz", variable=opcion_doble_faz)
    chk_doble_faz.grid(row=0, column=0, padx=10, pady=5)

    opcion_usuario = ttk.IntVar()
    rad_publico = ttk.Radiobutton(marco_opciones_extra, text="Público", variable=opcion_usuario, value=0)
    rad_publico.grid(row=0, column=1, padx=10, pady=5)
    rad_estudiante = ttk.Radiobutton(marco_opciones_extra, text="Estudiante", variable=opcion_usuario, value=1)
    rad_estudiante.grid(row=0, column=2, padx=10, pady=5)

    opcion_color = ttk.IntVar()
    chk_color = ttk.Checkbutton(marco_opciones_extra, text="Color", variable=opcion_color)
    chk_color.grid(row=0, column=3, padx=10, pady=5)

    btn_calcular = ttk.Button(root, text="Calcular", command=calcular, bootstyle="primary")
    btn_calcular.pack(pady=10)

    btn_ajustes = ttk.Button(root, text="Ajustes Avanzados", command=mostrar_ventana_ajustes, bootstyle="secondary")
    btn_ajustes.pack(pady=5)

    btn_editar_precios = ttk.Button(root, text="Editar Precios", command=mostrar_ventana_editar_precios, bootstyle="secondary")
    btn_editar_precios.pack(pady=5)

    btn_salir = ttk.Button(root, text="Salir", command=salir, bootstyle="danger")
    btn_salir.pack(pady=10)

    root.mainloop()

menu_interactivo()
