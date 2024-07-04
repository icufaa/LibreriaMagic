import fitz
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Variables globales para los precios de fotocopia
precio_bn = 100
precio_color = 200

# Variable global para el precio actual de fotocopia
precio_fotocopia = precio_bn

def calcular_precios(ruta_pdf):
    try:
        doc = fitz.open(ruta_pdf)
        num_paginas = len(doc)
        total_copias = num_paginas * precio_fotocopia
        return total_copias
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo PDF: {e}")
        return None

def menu_interactivo():
    def salir():
        root.destroy()

    def seleccionar_archivo():
        ruta_pdf = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
        if ruta_pdf:
            ruta_var.set(ruta_pdf)

    def calcular():
        ruta_pdf = ruta_var.get()
        if ruta_pdf:
            total_copias = calcular_precios(ruta_pdf)
            if total_copias is not None:
                messagebox.showinfo("Resultado", f'El costo total de las fotocopias es: ${total_copias}')
        else:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo PDF.")

    def cambiar_precio():
        global precio_bn, precio_color, precio_fotocopia
        try:
            nuevo_precio_bn = int(entry_precio_bn.get())
            nuevo_precio_color = int(entry_precio_color.get())
            precio_bn = nuevo_precio_bn
            precio_color = nuevo_precio_color
            if opcion_precio.get() == 1:
                precio_fotocopia = precio_bn
            elif opcion_precio.get() == 2:
                precio_fotocopia = precio_color
            messagebox.showinfo("Precio Actualizado", f'El precio de la fotocopia ha sido actualizado a: ${precio_fotocopia} por página.')
        except ValueError:
            messagebox.showerror("Error", "Por favor, introduce un valor numérico válido para los precios.")

    root = tk.Tk()
    root.title("Calculadora de Archivos")
    root.geometry("700x500")
    root.config(background='#ACE3DF')

    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", background="#CCC")

    # Menú
    menu = tk.Menu(root)
    root.config(menu=menu)
    archivo_menu = tk.Menu(menu, tearoff=0)
    menu.add_cascade(label="Archivo", menu=archivo_menu)
    archivo_menu.add_command(label="Salir", command=salir)
    acciones_menu = tk.Menu(menu, tearoff=0)
    menu.add_cascade(label='Acciones', menu=acciones_menu)
    acciones_menu.add_command(label='Calcular Precio', command=calcular)

    # Widgets
    ruta_var = tk.StringVar()
    opcion_precio = tk.IntVar(value=1)

    label_instruccion = tk.Label(root, text="Selecciona un archivo PDF y calcula el costo de las fotocopias.", background='#ACE3DF')
    label_instruccion.pack(pady=20)

    btn_seleccionar = ttk.Button(root, text="Seleccionar Archivo PDF", command=seleccionar_archivo)
    btn_seleccionar.pack(pady=10)

    label_ruta = tk.Label(root, textvariable=ruta_var, background='#ACE3DF')
    label_ruta.pack(pady=10)

    frame_precio = tk.Frame(root, background='#ACE3DF')
    frame_precio.pack(pady=10)

    radio_bn = tk.Radiobutton(frame_precio, text="Blanco y Negro", variable=opcion_precio, value=1, background='#ACE3DF')
    radio_bn.pack(side=tk.LEFT)

    radio_color = tk.Radiobutton(frame_precio, text="Color", variable=opcion_precio, value=2, background='#ACE3DF')
    radio_color.pack(side=tk.LEFT)

    label_precio_bn = tk.Label(root, text="Precio Blanco y Negro:", background='#ACE3DF')
    label_precio_bn.pack(pady=5)
    entry_precio_bn = tk.Entry(root)
    entry_precio_bn.insert(0, str(precio_bn))
    entry_precio_bn.pack(pady=5)

    label_precio_color = tk.Label(root, text="Precio Color:", background='#ACE3DF')
    label_precio_color.pack(pady=5)
    entry_precio_color = tk.Entry(root)
    entry_precio_color.insert(0, str(precio_color))
    entry_precio_color.pack(pady=5)

    btn_cambiar_precio = ttk.Button(root, text="Cambiar Precio", command=cambiar_precio)
    btn_cambiar_precio.pack(pady=10)

    btn_calcular = ttk.Button(root, text="Calcular Precio", command=calcular)
    btn_calcular.pack(pady=10)

    root.mainloop()

if __name__ == '__main__':
    menu_interactivo()
