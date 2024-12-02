import ssl
import socket
from datetime import datetime
import openpyxl
from openpyxl.styles import NamedStyle
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import threading

def obtener_informacion_certificado(hostname):
    # Establecer una conexión SSL con el servidor
    contexto = ssl.create_default_context()
    with contexto.wrap_socket(socket.socket(socket.AF_INET), server_hostname=hostname) as sock:
        sock.settimeout(3)  # Tiempo de espera para la conexión
        sock.connect((hostname, 443))
        cert = sock.getpeercert()
    
    # Extraer la fecha de caducidad del certificado
    fecha_caducidad_str = cert['notAfter']
    fecha_caducidad = datetime.strptime(fecha_caducidad_str, '%b %d %H:%M:%S %Y %Z')

    # Extraer el nombre del certificado (CN)
    nombre_certificado = next((x[0][1] for x in cert['subject'] if x[0][0] == 'commonName'), 'Desconocido')

    # Extraer la entidad certificadora (O)
    entidad_certificadora = next((x[0][1] for x in cert['issuer'] if x[0][0] == 'organizationName'), 'Desconocido')

    return fecha_caducidad, nombre_certificado, entidad_certificadora

def procesar_urls_desde_excel(archivo_entrada, archivo_salida, done_event):
    try:
        # Leer el archivo Excel
        df = pd.read_excel(archivo_entrada)

        # Crear una lista para almacenar los resultados
        resultados = []

        # Configurar la barra de progreso
        total_urls = len(df)
        progress_bar['maximum'] = total_urls

        # Procesar cada URL en el archivo Excel
        for index, row in df.iterrows():
            url = row['URL']
            try:
                fecha_caducidad, nombre_certificado, entidad_certificadora = obtener_informacion_certificado(url)
                # Convertir la fecha de caducidad a cadena en el formato DD/MM/YYYY
                fecha_caducidad_str = fecha_caducidad.strftime('%d/%m/%Y')
                resultados.append({
                    'URL': url, 
                    'Fecha de Caducidad': fecha_caducidad_str, 
                    'Nombre del Certificado (CN)': nombre_certificado, 
                    'Entidad Certificadora (O)': entidad_certificadora,
                    'Area Responsable': row.iloc[1]
                })
            except Exception as e:
                resultados.append({
                    'URL': url, 
                    'Fecha de Caducidad': f'Error: {e}', 
                    'Nombre del Certificado (CN)': 'Error', 
                    'Entidad Certificadora (O)': 'Error',
                    'Area Responsable': ''
                })

            # Actualizar la barra de progreso
            progress_bar['value'] = index + 1
            root.update_idletasks()

        # Convertir los resultados a un DataFrame
        df_resultados = pd.DataFrame(resultados)

        # Convertir la columna de fechas a tipo datetime para ordenar
        df_resultados['Fecha de Caducidad'] = pd.to_datetime(df_resultados['Fecha de Caducidad'], format='%d/%m/%Y', errors='coerce')

        # Ordenar el DataFrame por la columna de fechas
        df_resultados = df_resultados.sort_values(by='Fecha de Caducidad')

        # Guardar los resultados en un nuevo archivo Excel
        df_resultados.to_excel(archivo_salida, index=False)
        
        
        # Aplicar formato de fecha corta a la columna 'Fecha de Caducidad'
        wb = openpyxl.load_workbook(archivo_salida)
        ws = wb.active

        formato = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        
        for celda in ws['B'][1:]:  
            celda.style = formato

        wb.save(archivo_salida)


        # Notificar que el proceso ha terminado
        done_event.set()
    except Exception as e:
        # Mostrar mensaje de error
        estado_var.set("Error en el proceso.")
        messagebox.showerror("Error", f"Ha ocurrido un error: {e}")
        done_event.set()

def seleccionar_archivo():
    archivo_entrada = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if archivo_entrada:
        entrada_var.set(archivo_entrada)

def ejecutar_proceso():
    archivo_entrada = entrada_var.get()
    if not archivo_entrada:
        messagebox.showerror("Error", "Por favor, selecciona un archivo de entrada.")
        return

    # Mostrar mensaje de ejecución en proceso
    estado_var.set("Ejecución en proceso...")
    root.update_idletasks()

    try:
        # Obtener la ruta de la carpeta de descargas
        carpeta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")
        archivo_salida = os.path.join(carpeta_descargas, "informacion_certificados.xlsx")

        # Crear un evento para notificar cuando el hilo termine
        done_event = threading.Event()

        # Ejecutar el proceso en un hilo separado
        threading.Thread(target=procesar_urls_desde_excel, args=(archivo_entrada, archivo_salida, done_event)).start()

        # Esperar a que el hilo termine
        root.after(100, lambda: check_thread(done_event, archivo_salida))
    except Exception as e:
        # Mostrar mensaje de error
        estado_var.set("Error en el proceso.")
        messagebox.showerror("Error", f"Ha ocurrido un error: {e}")

def check_thread(done_event, archivo_salida):
    if done_event.is_set():
        # Mostrar mensaje de éxito
        estado_var.set("Proceso completado exitosamente.")
        messagebox.showinfo("Éxito", f"El archivo se ha guardado en {archivo_salida}")
    else:
        # Volver a comprobar después de 100 ms
        root.after(100, lambda: check_thread(done_event, archivo_salida))

# Crear la ventana principal
root = tk.Tk()
root.title("Obtener Certificados SSL")

# Establecer el tamaño de la ventana principal
root.geometry("400x300")

# Variables de control
entrada_var = tk.StringVar()
estado_var = tk.StringVar()

# Crear y colocar los widgets
tk.Label(root, text="Selecciona el archivo de entrada:").pack(pady=5)
tk.Entry(root, textvariable=entrada_var, width=50).pack(pady=5)
tk.Button(root, text="Seleccionar archivo", command=seleccionar_archivo).pack(pady=5)
tk.Button(root, text="Ejecutar", command=ejecutar_proceso).pack(pady=20)
tk.Label(root, textvariable=estado_var).pack(pady=5)

# Añadir una barra de progreso
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

# Iniciar el bucle principal de la aplicación
root.mainloop()
