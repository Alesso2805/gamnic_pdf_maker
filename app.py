import locale
import os
from datetime import timedelta, datetime
from threading import Thread
from tkinter import messagebox, ttk
import tkinter as tk
import win32com.client
from pypdf import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from io import BytesIO
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pythoncom
import config
import clientes

pdfmetrics.registerFont(TTFont("Calibri-Bold", "calibrib.ttf"))
pdfmetrics.registerFont(TTFont("Calibri", "calibri.ttf"))

# ================================
# Códigos de clientes (solo numéricos)
# ================================
clientes = clientes.CLIENTES

locale.setlocale(locale.LC_TIME, "es_ES.UTF-8")

hoy = datetime.now()
primer_dia_mes_actual = hoy.replace(day=1)
mes_anterior = primer_dia_mes_actual - timedelta(days=1)

nombre_mes_anterior = mes_anterior.strftime("%B").capitalize()

numero_mes_anterior = mes_anterior.strftime("%m")

año_anterior = mes_anterior.year

# ================================
# Hojas deseadas para exportar
# ================================
hojas_deseadas = ["Indice", "Resumen", "Gráficos contribución mensual", "Resultados", "Consolidado", "Movimientos", "Global"]

# ================================
# Función para procesar un cliente
# ================================
def procesar_cliente(codigo_cliente):
    pythoncom.CoInitialize()

    try:
        codigo_padded = codigo_cliente.zfill(3)  # Asegura formato "092"
        contraseña_excel = f"gamnic{codigo_padded}"

        base_path = config.BASE_PATH
        carpeta_cliente = next(
            (nombre for nombre in os.listdir(base_path) if nombre.startswith(codigo_padded)),
            None
        )

        if not carpeta_cliente:
            print(f"❌ No se encontró carpeta para el código {codigo_padded}")
            return

        archivo_excel = os.path.join(base_path, carpeta_cliente, f"{codigo_padded} - Estado de Cuenta - Generador - copia.xlsm")
        ruta_caratula_generada = os.path.join(base_path, f"caratula_generada.pdf")
        imagen_caratula = config.IMAGEN_CARATULA
        texto_caratula = f"{codigo_padded} – Portafolio Consolidado – {nombre_mes_anterior} {año_anterior}"
        pdf_contenido = os.path.join(base_path, "temp_contenido.pdf")
        pdf_salida_sin_footer = os.path.join(base_path, carpeta_cliente, "pdf_sin_pie.pdf")
        pdf_salida_con_footer = os.path.join(base_path, carpeta_cliente, f"{codigo_padded} - {año_anterior} {numero_mes_anterior} - Estado de Cuenta.pdf")

        def generar_caratula(path_salida, imagen_path, texto):
            c = canvas.Canvas(path_salida, pagesize=landscape(letter))
            width, height = landscape(letter)
            imagen = ImageReader(imagen_path)
            c.drawImage(imagen, 4 * inch, 4 * inch, width=3.5*inch, preserveAspectRatio=True, mask='auto')
            c.setFont("Calibri-Bold", 18)
            c.drawCentredString(width / 2, height / 2 - inch, texto)
            c.save()

        generar_caratula(ruta_caratula_generada, imagen_caratula, texto_caratula)

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.AutomationSecurity = 3

        try:
            wb = excel.Workbooks.Open(archivo_excel, False, None, None, contraseña_excel)
            wb.Worksheets(hojas_deseadas).Select()
            wb.ActiveSheet.ExportAsFixedFormat(0, pdf_contenido)
            wb.Close(False)
        except Exception as e:
            print(f"❌ Error con el cliente {codigo_padded}: {e}")
            excel.Quit()
            return

        excel.Quit()

        writer = PdfWriter()
        writer.append(PdfReader(ruta_caratula_generada))
        writer.append(PdfReader(pdf_contenido))

        def agregar_imagen_a_paginas(pdf_entrada, pdf_salida, imagen_path):
            reader = PdfReader(pdf_entrada)
            writer = PdfWriter()
            imagen = ImageReader(imagen_path)

            for i, page in enumerate(reader.pages):
                if i == 0:
                    writer.add_page(page)
                    continue

                page_width = float(page.mediabox.width)
                page_height = float(page.mediabox.height)

                packet = BytesIO()
                can = canvas.Canvas(packet, pagesize=(page_width, page_height))

                image_width = 1 * inch
                image_height = image_width
                can.drawImage(
                    imagen,
                    page_width - image_width - 0.2 * inch,  # Posición X
                    0.05 * inch,  # Posición Y
                    width=image_width,
                    height=image_height,
                    preserveAspectRatio=True,
                    mask='auto'
                )
                can.save()

                packet.seek(0)
                overlay = PdfReader(packet).pages[0]
                page.merge_page(overlay)
                writer.add_page(page)

            with open(pdf_salida, "wb") as f:
                writer.write(f)

        with open(pdf_salida_sin_footer, "wb") as f:
            writer.write(f)

        reader = PdfReader(pdf_salida_sin_footer)
        final_writer = PdfWriter()

        for i, page in enumerate(reader.pages):
            if i < 2:
                final_writer.add_page(page)
                continue

            packet = BytesIO()
            can = canvas.Canvas(packet, pagesize=landscape(letter))
            footer = f"{i + 1}"
            can.setFont("Calibri", 9)
            can.drawCentredString(11 * inch / 2, 0.4 * inch, footer)
            can.save()

            packet.seek(0)
            overlay = PdfReader(packet).pages[0]
            page.merge_page(overlay)
            final_writer.add_page(page)

        pdf_con_footer = os.path.join(base_path, carpeta_cliente, "pdf_con_footer_temp.pdf")
        with open(pdf_con_footer, "wb") as f:
            final_writer.write(f)

        agregar_imagen_a_paginas(pdf_con_footer, pdf_salida_con_footer, imagen_caratula)

        reader_final = PdfReader(pdf_salida_con_footer)
        encrypted_writer = PdfWriter()

        for page in reader_final.pages:
            encrypted_writer.add_page(page)

        encrypted_writer.encrypt(
            user_password=contraseña_excel,
            owner_password=None,
            use_128bit=True
        )

        with open(pdf_salida_con_footer, "wb") as f:
            encrypted_writer.write(f)

        print(f"✅ PDF generado correctamente para cliente {codigo_padded}: {pdf_salida_con_footer}")

        if os.path.exists(pdf_con_footer):
            os.remove(pdf_con_footer)

        if os.path.exists(pdf_salida_sin_footer):
            os.remove(pdf_salida_sin_footer)

    finally:
        pythoncom.CoUninitialize()

# -------------------------------
# GUI con Tkinter
# -------------------------------
def ejecutar_seleccionados():
    seleccionados = [lista_codigos.get(i) for i in lista_codigos.curselection()]
    if not seleccionados:
        messagebox.showwarning("Aviso", "Selecciona al menos un cliente.")
        return

    def run():
        for codigo in seleccionados:
            procesar_cliente(codigo)
        messagebox.showinfo("Proceso terminado", "Se completó la generación de PDFs.")

    Thread(target=run).start()

def ejecutar_todos():
    def run():
        for codigo in clientes:
            procesar_cliente(codigo)
        messagebox.showinfo("Proceso terminado", "Se completó la generación de todos los PDFs.")

    Thread(target=run).start()

# Crear ventana principal
ventana = tk.Tk()
ventana.title("Generador de PDFs - GAMNIC")
ventana.geometry("400x500")

# Etiqueta de instrucciones
etiqueta = tk.Label(ventana, text="Selecciona uno o varios códigos de cliente:", font=("Helvetica", 11))
etiqueta.pack(pady=10)

# Lista de selección múltiple
lista_codigos = tk.Listbox(ventana, selectmode=tk.MULTIPLE, width=30, height=20, exportselection=False)
for cliente in clientes:
    lista_codigos.insert(tk.END, cliente)
lista_codigos.pack(pady=10)

# Botones
boton_seleccionados = ttk.Button(ventana, text="Procesar seleccionados", command=ejecutar_seleccionados)
boton_seleccionados.pack(pady=5)

boton_todos = ttk.Button(ventana, text="Procesar TODOS", command=ejecutar_todos)
boton_todos.pack(pady=5)

# Iniciar la GUI
ventana.mainloop()
