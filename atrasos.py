import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
import datetime
from escpos.printer import Usb

printer = Usb(0x04b8, 0x0202, 0, 0x82, 0x02)  # Cambia los valores según la impresora térmica que estés utilizando

def cargar_datos_desde_excel(ruta_archivo):
    try:
        # Leer todas las hojas del archivo Excel
        data_frames = pd.read_excel(ruta_archivo, sheet_name=None)

        # Inicializar una lista para almacenar los datos de cada hoja
        data_list = []

        # Agregar los datos de cada hoja a la lista
        for key, df in data_frames.items():
            data_list.append(df.to_dict(orient='list'))

        # Concatenar los datos de todas las hojas en un solo diccionario
        data_dict = {}
        for data in data_list:
            for key, value in data.items():
                if key not in data_dict:
                    data_dict[key] = []
                data_dict[key].extend(value)

        return data_dict
    except Exception as e:
        print(f"Error al cargar el archivo Excel: {e}")
        return None

def filtrar_por_curso(curso):
    global data_dict
    if data_dict:
        filtered_data = {key: [] for key in data_dict.keys()}
        for index in range(len(data_dict['Curso'])):
            if data_dict['Curso'][index] == curso:
                for key in data_dict.keys():
                    filtered_data[key].append(data_dict[key][index])
        return filtered_data
    else:
        return None

def actualizar_datos(event=None):
    print("Actualizando datos...")
    selected_rut = entry_id.get()
    selected_nombre = entry_nombre.get()
    print("RUT seleccionado:", selected_rut)
    print("Nombre seleccionado:", selected_nombre)

    if selected_rut in data_dict['RUT']:
        index = data_dict['RUT'].index(selected_rut)
        entry_nombre.delete(0, tk.END)
        entry_nombre.insert(0, data_dict['Nombre'][index])
        entry_curso.delete(0, tk.END)
        entry_curso.insert(0, data_dict['Curso'][index])
    elif selected_nombre in data_dict['Nombre']:
        index = data_dict['Nombre'].index(selected_nombre)
        entry_id.delete(0, tk.END)
        entry_id.insert(0, data_dict['RUT'][index])
        entry_curso.delete(0, tk.END)
        entry_curso.insert(0, data_dict['Curso'][index])
    else:
        entry_nombre.delete(0, tk.END)
        entry_curso.delete(0, tk.END)
        print("Nombre o RUT no encontrado en la lista.")

    print("Datos actualizados.")

def actualizar_combobox_nombre(event=None):
    selected_nombre = entry_nombre.get()
    if selected_nombre in data_dict['Nombre']:
        index = data_dict['Nombre'].index(selected_nombre)
        entry_id.delete(0, tk.END)
        entry_id.insert(0, data_dict['RUT'][index])
        entry_curso.delete(0, tk.END)
        entry_curso.insert(0, data_dict['Curso'][index])
    else:
        entry_id.delete(0, tk.END)
        entry_curso.delete(0, tk.END)
        print("Nombre no encontrado en la lista.")

def seleccionar_archivo_excel():
    ruta_archivo = filedialog.askopenfilename(title="Seleccionar Archivo Excel")
    if ruta_archivo:
        global data_dict
        data_dict = cargar_datos_desde_excel(ruta_archivo)
        if data_dict:
            cargar_comboboxes()
            with open('ultimo_archivo.txt', 'w') as f:
                f.write(ruta_archivo)

def cargar_comboboxes(event=None):
    global data_dict
    curso_seleccionado = entry_curso.get()
    if curso_seleccionado:
        filtered_data = filtrar_por_curso(curso_seleccionado)
        if filtered_data:
            entry_id.delete(0, tk.END)
            entry_nombre.delete(0, tk.END)
        else:
            entry_id.delete(0, tk.END)
            entry_nombre.delete(0, tk.END)
            entry_curso.delete(0, tk.END)
    else:
        entry_id.delete(0, tk.END)
        entry_nombre.delete(0, tk.END)
        entry_curso.delete(0, tk.END)

def cargar_ultimo_archivo():
    try:
        with open('ultimo_archivo.txt', 'r') as f:
            ruta_archivo = f.read().strip()
            global data_dict
            data_dict = cargar_datos_desde_excel(ruta_archivo)
            if data_dict:
                cargar_comboboxes()
    except FileNotFoundError:
        pass

def limpiar_consulta():
    entry_id.delete(0, tk.END)
    entry_nombre.delete(0, tk.END)
    entry_curso.delete(0, tk.END)

def generar_ticket():
    rut = entry_id.get()
    nombre = entry_nombre.get()
    curso = entry_curso.get()
    motivo = area_texto_motivo.get("1.0", tk.END).strip()
    informacion_adicional = area_texto_informacion_adicional.get("1.0", tk.END).strip()
    fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Construir la cadena de texto a imprimir
    ticket_text = f"Ticket de Atraso\n=================\nFecha y Hora: {fecha_hora}\nRUT: {rut}\nNombre: {nombre}\nCurso: {curso}\nMotivo del Atraso:\n{motivo}\nInformación Adicional:\n{informacion_adicional}\n=================\nGracias.\n"

    # Conectar a la impresora
    printer = Usb(0x04b8, 0x0202, 0, 0x82, 0x02)

    # Imprimir el ticket
    printer.text(ticket_text)
    printer.cut()
    printer.close()

# Crear la ventana principal
ventana_principal = tk.Tk()
ventana_principal.title("Aplicación de Tickets de Atraso")

# Frame para la sección de identificación
frame_identificacion = tk.Frame(ventana_principal)
frame_identificacion.pack(padx=10, pady=10, fill=tk.BOTH)

# Etiqueta para indicar el formato del RUT
etiqueta_formato_rut = tk.Label(frame_identificacion, text="El RUT debe ser escrito con puntos y guión", fg="grey")
etiqueta_formato_rut.pack(side=tk.BOTTOM, padx=(10, 0))

# Etiquetas y Entrys para RUT y Nombre
etiqueta_id = tk.Label(frame_identificacion, text="RUT:")
etiqueta_id.pack(side=tk.LEFT)
entry_id = tk.Entry(frame_identificacion)
entry_id.pack(side=tk.LEFT, fill=tk.X, expand=True)
entry_id.bind('<KeyRelease>', actualizar_datos)

etiqueta_nombre = tk.Label(frame_identificacion, text="Nombre:")
etiqueta_nombre.pack(side=tk.LEFT)
entry_nombre = tk.Entry(frame_identificacion)
entry_nombre.pack(side=tk.LEFT, fill=tk.X, expand=True)
entry_nombre.bind('<KeyRelease>', actualizar_combobox_nombre)

# Frame para la sección de información del atraso
frame_atraso = tk.Frame(ventana_principal)
frame_atraso.pack(padx=10, pady=10, fill=tk.BOTH)

# Reemplazar Combobox de curso con Entry
etiqueta_curso = tk.Label(frame_atraso, text="Curso:")
etiqueta_curso.pack(side=tk.LEFT)
entry_curso = tk.Entry(frame_atraso, width=10)
entry_curso.pack(side=tk.LEFT)
entry_curso.bind('<KeyRelease>', cargar_comboboxes)

# Text box to display the current time and date
time_textbox = tk.Text(ventana_principal, height=1, width=20)
time_textbox.pack()

def update_time_textbox():
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    time_textbox.delete(1.0, tk.END)
    time_textbox.insert(tk.END, current_time)
    ventana_principal.after(1000, update_time_textbox)

update_time_textbox()

# Botón para limpiar la consulta
boton_limpiar = tk.Button(frame_atraso, text="Limpiar Consulta", command=limpiar_consulta)
boton_limpiar.pack(side=tk.LEFT, padx=10)

etiqueta_motivo = tk.Label(frame_atraso, text="Motivo del Atraso:")
etiqueta_motivo.pack()
area_texto_motivo = tk.Text(frame_atraso, height=5, width=30)
area_texto_motivo.pack()

etiqueta_informacion_adicional = tk.Label(frame_atraso, text="Información Adicional:")
etiqueta_informacion_adicional.pack()
area_texto_informacion_adicional = tk.Text(frame_atraso, height=3, width=30)
area_texto_informacion_adicional.pack()

# Botón para seleccionar el archivo Excel
boton_seleccionar_archivo = tk.Button(ventana_principal, text="Seleccionar Archivo Excel", command=seleccionar_archivo_excel)
boton_seleccionar_archivo.pack(pady=10)

# Botón para generar el ticket
boton_generar_ticket = tk.Button(ventana_principal, text="Generar Ticket", command=generar_ticket)
boton_generar_ticket.pack(pady=10)

# Botón para salir
boton_salir = tk.Button(ventana_principal, text="Salir", command=ventana_principal.quit)
boton_salir.pack()

# Cargar el último archivo al iniciar
cargar_ultimo_archivo()

# Iniciar el bucle principal para mostrar la ventana
ventana_principal.mainloop()
