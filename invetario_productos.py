import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime
import os

archivo_excel = "inventario.xlsx"
productos = []


# ---------------- FUNCIONES ---------------- #

def cargar_datos():
    productos.clear()
    tabla.delete(*tabla.get_children())

    if os.path.exists(archivo_excel):
        try:
            wb = load_workbook(archivo_excel)
            ws = wb.active

            for fila in ws.iter_rows(min_row=2, values_only=True):
                codigo, nombre, cantidad, fecha = fila
                if codigo is not None:
                    producto = (str(codigo), str(nombre), cantidad, fecha)
                    productos.append(producto)
                    tabla.insert("", tk.END, values=producto)

            wb.close()
        except Exception as error:
            messagebox.showerror(
                "Error",
                f"No se pudieron cargar los datos del archivo Excel:\n{error}"
            )


def guardar_datos():
    wb = Workbook()
    ws = wb.active

    headers = ["Codigo", "Nombre", "Cantidad", "Fecha"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for producto in productos:
        ws.append(producto)

    for fila in ws.iter_rows():
        for celda in fila:
            celda.alignment = Alignment(horizontal="center")

    try:
        wb.save(archivo_excel)
    except PermissionError:
        messagebox.showerror(
            "Error",
            "No se pudo guardar el archivo. Cierra Excel e inténtalo de nuevo."
        )
    except Exception as error:
        messagebox.showerror("Error", f"No se pudieron guardar los datos:\n{error}")


def limpiar_campos():
    entrada_codigo.delete(0, tk.END)
    entrada_nombre.delete(0, tk.END)
    entrada_cantidad.delete(0, tk.END)


def agregar():
    codigo = entrada_codigo.get().strip()
    nombre = entrada_nombre.get().strip()
    cantidad_texto = entrada_cantidad.get().strip()

    if not codigo or not nombre or not cantidad_texto:
        messagebox.showwarning("Datos incompletos", "Complete todos los campos.")
        return

    if any(str(producto[0]) == codigo for producto in productos):
        messagebox.showerror("Código duplicado", "Ese código ya existe.")
        return

    try:
        cantidad = int(cantidad_texto)
        if cantidad < 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Cantidad inválida", "Ingrese una cantidad entera igual o mayor que cero.")
        return

    fecha = datetime.now().strftime("%d/%m/%Y")
    producto = (codigo, nombre, cantidad, fecha)
    productos.append(producto)
    tabla.insert("", tk.END, values=producto)

    guardar_datos()
    limpiar_campos()


def eliminar():
    seleccion = tabla.selection()
    if not seleccion:
        messagebox.showwarning("Sin selección", "Seleccione un producto.")
        return

    item = seleccion[0]
    valores = tabla.item(item, "values")
    codigo = str(valores[0])

    for indice, producto in enumerate(productos):
        if str(producto[0]) == codigo:
            productos.pop(indice)
            break

    tabla.delete(item)
    guardar_datos()


def cargar():
    seleccion = tabla.selection()
    if not seleccion:
        messagebox.showwarning("Sin selección", "Seleccione un producto.")
        return

    valores = tabla.item(seleccion[0], "values")

    limpiar_campos()
    entrada_codigo.insert(0, valores[0])
    entrada_nombre.insert(0, valores[1])
    entrada_cantidad.insert(0, valores[2])


def modificar():
    seleccion = tabla.selection()
    if not seleccion:
        messagebox.showwarning("Sin selección", "Seleccione un producto.")
        return

    item = seleccion[0]
    valores = tabla.item(item, "values")
    codigo_original = str(valores[0])

    codigo = entrada_codigo.get().strip()
    nombre = entrada_nombre.get().strip()
    cantidad_texto = entrada_cantidad.get().strip()

    if not codigo or not nombre or not cantidad_texto:
        messagebox.showwarning("Datos incompletos", "Complete todos los campos.")
        return

    if any(
        str(producto[0]) == codigo and str(producto[0]) != codigo_original
        for producto in productos
    ):
        messagebox.showerror("Código duplicado", "Ese código ya pertenece a otro producto.")
        return

    try:
        cantidad = int(cantidad_texto)
        if cantidad < 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Cantidad inválida", "Ingrese una cantidad entera igual o mayor que cero.")
        return

    fecha = datetime.now().strftime("%d/%m/%Y")
    producto_actualizado = (codigo, nombre, cantidad, fecha)

    for indice, producto in enumerate(productos):
        if str(producto[0]) == codigo_original:
            productos[indice] = producto_actualizado
            break

    tabla.item(item, values=producto_actualizado)
    guardar_datos()
    limpiar_campos()


def buscar():
    texto = entrada_buscar.get().strip().lower()
    tabla.delete(*tabla.get_children())

    for codigo, nombre, cantidad, fecha in productos:
        if (
            texto in str(codigo).lower()
            or texto in str(nombre).lower()
            or texto in str(fecha).lower()
        ):
            tabla.insert("", tk.END, values=(codigo, nombre, cantidad, fecha))


def mostrar_todo():
    entrada_buscar.delete(0, tk.END)
    tabla.delete(*tabla.get_children())

    for producto in productos:
        tabla.insert("", tk.END, values=producto)


def abrir_excel():
    if not os.path.exists(archivo_excel):
        messagebox.showwarning("Archivo no encontrado", "Todavía no existe el archivo de inventario.")
        return

    try:
        os.startfile(archivo_excel)
    except Exception as error:
        messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{error}")


# ---------------- INTERFAZ ---------------- #

ventana = tk.Tk()
ventana.title("Inventario PRO FINAL")
ventana.geometry("800x550")
ventana.configure(bg="#f0f2f5")

style = ttk.Style()
style.theme_use("clam")
style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
style.configure("Treeview", rowheight=25, font=("Segoe UI", 10))

# Título
tk.Label(
    ventana,
    text="Inventarios RJ",
    font=("Segoe UI", 18, "bold"),
    bg="#f0f2f5"
).pack(pady=10)

# Formulario
frame_form = tk.Frame(ventana, bg="#f0f2f5")
frame_form.pack(pady=10)

ttk.Label(frame_form, text="Código").grid(row=0, column=0, padx=10)
entrada_codigo = ttk.Entry(frame_form)
entrada_codigo.grid(row=1, column=0)

ttk.Label(frame_form, text="Producto").grid(row=0, column=1, padx=10)
entrada_nombre = ttk.Entry(frame_form)
entrada_nombre.grid(row=1, column=1)

ttk.Label(frame_form, text="Cantidad").grid(row=0, column=2, padx=10)
entrada_cantidad = ttk.Entry(frame_form)
entrada_cantidad.grid(row=1, column=2)

ttk.Button(
    frame_form,
    text="➕ Agregar",
    command=agregar
).grid(row=1, column=3, padx=10)

# Buscador
frame_buscar = tk.Frame(ventana, bg="#f0f2f5")
frame_buscar.pack(pady=5)

ttk.Label(frame_buscar, text="Buscar").grid(row=0, column=0)
entrada_buscar = ttk.Entry(frame_buscar, width=30)
entrada_buscar.grid(row=0, column=1, padx=5)

ttk.Button(
    frame_buscar,
    text="🔍 Buscar",
    command=buscar
).grid(row=0, column=2, padx=5)

ttk.Button(
    frame_buscar,
    text="🔄 Mostrar Todo",
    command=mostrar_todo
).grid(row=0, column=3, padx=5)

entrada_buscar.bind("<Return>", lambda evento: buscar())

# Botones
frame_botones = tk.Frame(ventana, bg="#f0f2f5")
frame_botones.pack(pady=10)

ttk.Button(
    frame_botones,
    text="🗑 Eliminar",
    command=eliminar
).grid(row=0, column=0, padx=10)

ttk.Button(
    frame_botones,
    text="📥 Cargar",
    command=cargar
).grid(row=0, column=1, padx=10)

ttk.Button(
    frame_botones,
    text="✏ Modificar",
    command=modificar
).grid(row=0, column=2, padx=10)

ttk.Button(
    frame_botones,
    text="📂 Excel",
    command=abrir_excel
).grid(row=0, column=3, padx=10)

# Tabla
frame_tabla = tk.Frame(ventana)
frame_tabla.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

tabla = ttk.Treeview(
    frame_tabla,
    columns=("Codigo", "Nombre", "Cantidad", "Fecha"),
    show="headings"
)

tabla.heading("Codigo", text="Código")
tabla.heading("Nombre", text="Producto")
tabla.heading("Cantidad", text="Cantidad")
tabla.heading("Fecha", text="Fecha")

tabla.column("Codigo", anchor="center", width=100)
tabla.column("Nombre", anchor="center", width=200)
tabla.column("Cantidad", anchor="center", width=100)
tabla.column("Fecha", anchor="center", width=120)

tabla.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scroll = ttk.Scrollbar(frame_tabla, orient="vertical", command=tabla.yview)
tabla.configure(yscrollcommand=scroll.set)
scroll.pack(side=tk.RIGHT, fill=tk.Y)

# Inicio
cargar_datos()
ventana.mainloop()