import fitz  # PyMuPDF
import cv2
import numpy as np
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from PIL import Image

# Variables globales para los precios de fotocopia según la tabla
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
        for ruta_pdf in ruta_pdfs:
            doc = fitz.open(ruta_pdf)
            num_paginas = len(doc)
            detalles_archivos[ruta_pdf] = num_paginas
            total_paginas += num_paginas

            for pagina in doc:
                total_color += obtener_porcentaje_color(pagina)

        porcentaje_color_promedio = total_color / total_paginas

        if doble_faz:
            total_paginas = (total_paginas + 1) // 2  # Redondear para arriba
            total_paginas = total_paginas + (total_paginas % 2)  # Asegurar que sea par
            tipo = "doble"
        else:
            tipo = "simple"

        precio_fotocopia = obtener_precio(total_paginas, tipo, usuario, porcentaje_color_promedio, color)
        total_copias = total_paginas * precio_fotocopia
        return total_copias, detalles_archivos
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
                detalle_mensaje = "\n".join([f'Archivo: {ruta}, Hojas: {hojas}' for ruta, hojas in detalles_archivos.items()])
                messagebox.showinfo("Resultado", f'El costo total de las fotocopias es: ${total_copias}\n\nDetalles:\n{detalle_mensaje}')
        else:
            messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un archivo PDF.")

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
