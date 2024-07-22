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
from docx import Document
import docx2txt
import tempfile

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

def obtener_precio(cantidad, tipo, usuario, porcentaje_color, color, tamano_hoja):
    precios = precios_estudiante if usuario == "estudiante" else precios_publico

    if tamano_hoja == "A3":
        if color:
            precio_color = 1000  # Precio fijo para A3 a color
        else:
            precio_color = 200  # Precio fijo para A3 blanco y negro
    else:
        if color:
            precio_base_color = 100  # Precio base para A4 color
            incremento_color = porcentaje_color * 400  # Incremento por porcentaje de color
            precio_color = precio_base_color + incremento_color
            precio_color = round(precio_color / 50) * 50  # Redondear al múltiplo de 50 más cercano
            precio_color = min(precio_color, 500)  # Limitar a un máximo de $500
        else:
            precio_color = 0  # No hay ajuste por color si no es a color

    for limite, precio in sorted(precios[tipo].items(), reverse=True):
        if cantidad >= limite:
            return int(precio) + int(precio_color)
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

def obtener_porcentaje_color_docx(ruta):
    try:
        # Crear una carpeta temporal para guardar las imágenes extraídas
        temp_dir = tempfile.mkdtemp()
        images_path = os.path.join(temp_dir, "docx_images")
        
        # Asegurarse de que el directorio de imágenes existe
        os.makedirs(images_path, exist_ok=True)
        
        # Extraer todas las imágenes del DOCX en la carpeta temporal
        extraidas = docx2txt.process(ruta, images_path)
        print(f"Texto extraído: {extraidas}")  # Mensaje de depuración

        # Verificar si las imágenes han sido extraídas
        if not os.path.exists(images_path) or not os.listdir(images_path):
            print("No se encontraron imágenes.")
            return 0

        total_color_percentage = 0
        image_files = [f for f in os.listdir(images_path) if os.path.isfile(os.path.join(images_path, f))]
        
        print(f"Archivos de imagen encontrados: {image_files}")  # Mensaje de depuración

        for image_file in image_files:
            img_path = os.path.join(images_path, image_file)
            print(f"Procesando imagen: {img_path}")  # Mensaje de depuración

            if not os.path.exists(img_path):
                print(f"Error: {img_path} no existe.")
                continue

            img = Image.open(img_path)
            img_np = np.array(img)

            # Convertir la imagen a espacio de color HSV
            hsv = cv2.cvtColor(img_np, cv2.COLOR_RGB2HSV)

            # Definir umbral para detectar colores (excluyendo blanco y negro)
            lower_color = np.array([0, sensitivity, sensitivity])  # Usar sensibilidad global
            upper_color = np.array([179, 255, 255])

            # Crear una máscara para los colores
            mask = cv2.inRange(hsv, lower_color, upper_color)

            # Calcular el porcentaje de área coloreada
            color_percentage = np.sum(mask > 0) / mask.size
            print(f"Porcentaje de color para {img_path}: {color_percentage}")

            total_color_percentage += color_percentage

        # Limpiar los archivos temporales
        for image_file in image_files:
            os.remove(os.path.join(images_path, image_file))
        os.rmdir(images_path)
        os.rmdir(temp_dir)

        return total_color_percentage / len(image_files) if image_files else 0
    except Exception as e:
        print(f"Error al procesar el archivo DOCX: {e}")
        return 0





def calcular_precios(root, rutas, doble_faz=False, usuario="publico", color=False, tamano_hoja="A4"):
    detalles_archivos = {}
    try:
        total_paginas = 0
        total_costo_archivos = 0

        # Crear la ventana de progreso
        progress_window = Toplevel(root)
        progress_window.title("Procesando archivos")
        progress_window.geometry("400x100")

        # Crear la barra de progreso
        progress_bar = ttk.Progressbar(progress_window, length=300, mode='determinate')
        progress_bar.pack(pady=20)

        # Establecer el valor máximo de la barra de progreso
        progress_bar['maximum'] = len(rutas)

        for i, ruta in enumerate(rutas):
            ext = os.path.splitext(ruta)[1].lower()
            if ext == ".pdf":
                doc = fitz.open(ruta)
                num_paginas = len(doc)
                color_paginas = sum(obtener_porcentaje_color(pagina) for pagina in doc)
            elif ext == ".docx":
                doc = Document(ruta)
                num_paginas = len(doc.paragraphs)
                color_paginas = obtener_porcentaje_color_docx(ruta)
            else:
                raise ValueError(f"Formato de archivo no soportado: {ext}")

            total_paginas += num_paginas
            porcentaje_color_promedio = color_paginas / num_paginas if num_paginas else 0

            if doble_faz:
                paginas_para_precio = (num_paginas + 1) // 2  # Redondear para arriba
                tipo = "doble"
            else:
                paginas_para_precio = num_paginas
                tipo = "simple"

            precio_fotocopia = obtener_precio(paginas_para_precio, tipo, usuario, porcentaje_color_promedio, color, tamano_hoja)
            costo_archivo = paginas_para_precio * precio_fotocopia
            total_costo_archivos += costo_archivo
            detalles_archivos[ruta] = (num_paginas, costo_archivo)

            # Actualizar la barra de progreso
            progress_bar['value'] = i + 1
            progress_window.update()

        # Cerrar la ventana de progreso
        progress_window.destroy()

        return total_costo_archivos, detalles_archivos
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")
        return None, None



def menu_interactivo():
    def salir():
        root.destroy()

    def seleccionar_archivos():
        rutas = filedialog.askopenfilenames(filetypes=[("Archivos PDF y DOCX", "*.pdf;*.docx")])
        if rutas:
            rutas_var.set(", ".join(rutas))

    def calcular():
        rutas = rutas_var.get().split(", ")
        if rutas:
            doble_faz = opcion_doble_faz.get() == 1
            usuario = "estudiante" if opcion_usuario.get() == 1 else "publico"
            color = opcion_color.get() == 1
            tamano_hoja = "A3" if opcion_tamano_hoja.get() == 1 else "A4"
            total_copias, detalles_archivos = calcular_precios(root, rutas, doble_faz, usuario, color, tamano_hoja)
            if total_copias:
                opciones_seleccionadas = (
                    f"Tipo de usuario: {'Estudiante' if usuario == 'estudiante' else 'Público'}\n"
                    f"Tipo de fotocopia: {'Doble faz' if doble_faz else 'Simple'}\n"
                    f"Color: {'Sí' if color else 'No'}\n"
                    f"Tamaño de hoja: {tamano_hoja}\n"
                )
                detalle_mensaje = "\n\n".join([f'Archivo: {ruta}\nHojas: {hojas}\nPrecio: ${precio}' 
                                            for ruta, (hojas, precio) in detalles_archivos.items()])
                messagebox.showinfo("Resultado", f'{opciones_seleccionadas}\nEl costo total de las fotocopias es: ${total_copias}\n\nDetalles:\n{detalle_mensaje}')
        else:
            messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un archivo PDF o DOCX.")

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
                messagebox.showerror("Error", "Por favor, ingrese un valor válido para el umbral de sensibilidad.")

        ajustes_window = Toplevel(root)
        ajustes_window.title("Ajustes")
        ajustes_window.geometry("300x150")

        label_sensitivity = ttk.Label(ajustes_window, text="Umbral de Sensibilidad (0-255):")
        label_sensitivity.pack(pady=10)

        entry_sensitivity = ttk.Entry(ajustes_window)
        entry_sensitivity.insert(0, str(sensitivity))
        entry_sensitivity.pack(pady=5)

        boton_guardar = ttk.Button(ajustes_window, text="Guardar", command=guardar_ajustes)
        boton_guardar.pack(pady=20)

    def mostrar_ventana_precios():
        def actualizar_precios():
            try:
                nuevas_tablas = {}
                for usuario in ['publico', 'estudiante']:
                    nuevas_tablas[usuario] = {}
                    for tipo in ['simple', 'doble']:
                        nueva_tabla = {}
                        for key, entry in entries[usuario][tipo].items():
                            nueva_tabla[int(key)] = int(entry.get())
                        nuevas_tablas[usuario][tipo] = nueva_tabla
                guardar_precios(nuevas_tablas['publico'], nuevas_tablas['estudiante'])
                precios_publico.update(nuevas_tablas['publico'])
                precios_estudiante.update(nuevas_tablas['estudiante'])
                precios_window.destroy()
            except ValueError:
                messagebox.showerror("Error", "Por favor, ingrese valores numéricos válidos para los precios.")

        precios_window = Toplevel(root)
        precios_window.title("Precios")
        precios_window.geometry("500x500")

        frame = ttk.Frame(precios_window)
        frame.pack(pady=10, padx=10)

        notebook = ttk.Notebook(frame)
        notebook.pack()

        entries = {'publico': {'simple': {}, 'doble': {}}, 'estudiante': {'simple': {}, 'doble': {}}}

        for usuario in ['publico', 'estudiante']:
            user_frame = ttk.Frame(notebook)
            notebook.add(user_frame, text=usuario.capitalize())

            for tipo in ['simple', 'doble']:
                tipo_frame = ttk.Frame(user_frame)
                tipo_frame.pack(pady=10)

                ttk.Label(tipo_frame, text=f"Precios para {tipo} faz:").pack()
                precios = precios_publico if usuario == 'publico' else precios_estudiante
                for key, value in sorted(precios[tipo].items()):
                    row_frame = ttk.Frame(tipo_frame)
                    row_frame.pack(pady=2)
                    ttk.Label(row_frame, text=f"{key} páginas o más:").pack(side='left')
                    entry = ttk.Entry(row_frame, width=10)
                    entry.insert(0, str(value))
                    entry.pack(side='left')
                    entries[usuario][tipo][key] = entry

        boton_actualizar = ttk.Button(precios_window, text="Actualizar Precios", command=actualizar_precios)
        boton_actualizar.pack(pady=20)

    root = ttk.Window(themename="superhero")
    root.title("Calculadora de Precios de Fotocopias")
    root.geometry("500x400")

    rutas_var = StringVar()

    opcion_doble_faz = IntVar()
    opcion_usuario = IntVar()
    opcion_color = IntVar()
    opcion_tamano_hoja = IntVar()  # Nueva variable para el tamaño de hoja

    ttk.Label(root, text="Calculadora de Precios de Fotocopias", font=("Helvetica", 16)).pack(pady=20)
    
    ttk.Button(root, text="Seleccionar archivos PDF o DOCX", command=seleccionar_archivos).pack(pady=10)
    ttk.Entry(root, textvariable=rutas_var, state='readonly').pack(fill='x', padx=50, pady=5)

    frame_opciones = ttk.Frame(root)
    frame_opciones.pack(pady=10)

    ttk.Checkbutton(frame_opciones, text="Doble faz", variable=opcion_doble_faz).grid(row=0, column=0, padx=10)
    ttk.Radiobutton(frame_opciones, text="Público", variable=opcion_usuario, value=0).grid(row=0, column=1, padx=10)
    ttk.Radiobutton(frame_opciones, text="Estudiante", variable=opcion_usuario, value=1).grid(row=0, column=2, padx=10)
    ttk.Checkbutton(frame_opciones, text="Color", variable=opcion_color).grid(row=0, column=3, padx=10)
    ttk.Radiobutton(frame_opciones, text="A4", variable=opcion_tamano_hoja, value=0).grid(row=0, column=4, padx=10)  # Opción A4
    ttk.Radiobutton(frame_opciones, text="A3", variable=opcion_tamano_hoja, value=1).grid(row=0, column=5, padx=10)  # Opción A3

    ttk.Button(root, text="Calcular Precios", command=calcular).pack(pady=20)

    menu_bar = ttk.Menu(root)
    root.config(menu=menu_bar)
    
    archivo_menu = ttk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Archivo", menu=archivo_menu)
    archivo_menu.add_command(label="Salir", command=salir)
    
    ajustes_menu = ttk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Ajustes", menu=ajustes_menu)
    ajustes_menu.add_command(label="Configuración de Sensibilidad", command=mostrar_ventana_ajustes)
    ajustes_menu.add_command(label="Configuración de Precios", command=mostrar_ventana_precios)

    root.mainloop()

if __name__ == "__main__":
    menu_interactivo()
