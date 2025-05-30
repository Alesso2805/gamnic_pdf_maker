import locale
import re
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

def obtener_hojas_deseadas(wb):
    try:
        hoja_indice = wb.Worksheets("Indice")
        contenido_indice = []

        # Leer todas las celdas de la hoja Indice
        for fila in range(1, hoja_indice.UsedRange.Rows.Count + 1):
            for columna in range(1, hoja_indice.UsedRange.Columns.Count + 1):
                valor_celda = hoja_indice.Cells(fila, columna).Value
                if valor_celda:
                    contenido_indice.append(str(valor_celda))

        # Determinar las hojas a extraer basándose en el contenido de Indice
        hojas_dinamicas = ["Indice"]
        if any("Resumen de Resultados" in texto for texto in contenido_indice):
            hojas_dinamicas.append("Resumen")
        if any("Contribuciones al Portafolio" in texto for texto in contenido_indice):
            hojas_dinamicas.append("Gráficos contribución mensual")
        if any("Detalle de Resultados" in texto for texto in contenido_indice):
            hojas_dinamicas.append("Resultados")
        if any("Portafolio Global - Composición" in texto for texto in contenido_indice):
            hojas_dinamicas.append("Consolidado")
        if any("Movimientos del Mes" in texto for texto in contenido_indice):
            hojas_dinamicas.append("Movimientos")
        if any("Portafolio Global - Detalle" in texto for texto in contenido_indice):
            hojas_dinamicas.append("Global")
        # Agrega más condiciones según sea necesario

        print(f"📄 Hojas dinámicas determinadas: {hojas_dinamicas}")
        return hojas_dinamicas

    except Exception as e:
        print(f"❌ Error al leer la hoja Indice: {e}")
        return []
# ================================
# Función para procesar un cliente
# ================================

# Python
def procesar_cliente(codigo_cliente):
    pythoncom.CoInitialize()

    try:
        codigo_padded = codigo_cliente.zfill(3)  # Format "092"
        contraseña_excel = f"gamnic{codigo_padded}"

        base_path = config.BASE_PATH
        carpeta_cliente = next(
            (nombre for nombre in os.listdir(base_path) if nombre.startswith(codigo_padded)),
            None
        )

        if not carpeta_cliente:
            print(f"❌ No se encontró carpeta para el código {codigo_padded}")
            return

        ruta_caratula_generada = os.path.join(base_path, f"caratula_generada.pdf")
        imagen_caratula = config.IMAGEN_CARATULA

        # Search for all matching files in the folder
        generadores = [
            archivo for archivo in os.listdir(os.path.join(base_path, carpeta_cliente))
            if archivo.startswith(f"{codigo_padded} - Estado de Cuenta - Generador") and archivo.endswith("- copia.xlsm")
        ]

        print(f"Archivos encontrados para el código {codigo_padded}: {generadores}")


        # Generate the caratula
        def generar_caratula(path_salida, imagen_path, texto):
            c = canvas.Canvas(path_salida, pagesize=landscape(letter))
            width, height = landscape(letter)
            imagen = ImageReader(imagen_path)
            c.drawImage(imagen, 4 * inch, 4 * inch, width=3.5 * inch, preserveAspectRatio=True, mask='auto')
            c.setFont("Calibri-Bold", 18)
            c.drawCentredString(width / 2, height / 2 - inch, texto)
            c.save()

        for archivo_excel in generadores:
            match = re.search(r"Generador\s([A-Z]|Consolidado(?:\s\(([A-Z+]+)\))?)", archivo_excel)
            if match and match.group(1):
                generador_suffix = match.group(1).strip()
            else:
                generador_suffix = ""

            texto_caratula = f"{codigo_padded} {generador_suffix} – Portafolio Consolidado – {nombre_mes_anterior} {año_anterior}"

            ruta_caratula_generada = os.path.join(base_path, carpeta_cliente, f"caratula_{generador_suffix}.pdf")

            generar_caratula(ruta_caratula_generada, imagen_caratula, texto_caratula)

            archivo_excel_path = os.path.join(base_path, carpeta_cliente, archivo_excel)
            pdf_contenido = os.path.join(base_path, carpeta_cliente, f"temp_contenido_{archivo_excel}.pdf")
            pdf_salida_sin_footer = os.path.join(base_path, carpeta_cliente, f"pdf_sin_pie_{archivo_excel}.pdf")
            pdf_salida_con_footer = os.path.join(base_path,carpeta_cliente, f"{codigo_padded} - {año_anterior} {numero_mes_anterior} - Estado de Cuenta {generador_suffix}.pdf")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Ensure Excel is not visible
            excel.DisplayAlerts = False  # Disable Excel alerts
            excel.AutomationSecurity = 3  # Disable macros (1 = msoAutomationSecurityLow, 2 = msoAutomationSecurityByUI, 3 = msoAutomationSecurityForceDisable)

            # Open the Excel file
            try:
                wb = excel.Workbooks.Open(
                    os.path.join(base_path, carpeta_cliente, archivo_excel),
                    False,  # UpdateLinks
                    False,  # ReadOnly
                    None,  # Format
                    contraseña_excel,  # Password
                    None,  # WriteResPassword
                    True,  # IgnoreReadOnlyRecommended
                    None,  # Origin
                    None,  # Delimiter
                    False,  # Editable
                    False,  # Notify
                    None,  # Converter
                    False,  # AddToMru
                    None,  # Local
                    None  # CorruptLoad
                )         # Log all available sheet names for debugging
                available_sheets = [sheet.Name for sheet in wb.Worksheets]
                print(f"📄 Hojas disponibles en '{archivo_excel}': {available_sheets}")

                hojas_dinamicas = obtener_hojas_deseadas(wb)

                # Select sheets
                first_sheet = True
                for hoja in hojas_dinamicas:
                    try:
                        if first_sheet:
                            wb.Worksheets(hoja).Select()
                            first_sheet = False
                        else:
                            wb.Worksheets(hoja).Select(Replace=False)
                    except Exception as e:
                        print(f"⚠️ Hoja '{hoja}' no encontrada: {e}")
                        continue

                # Export to PDF
                pdf_contenido = os.path.join(base_path, carpeta_cliente, f"temp_contenido_{codigo_padded}.pdf")
                wb.ActiveSheet.ExportAsFixedFormat(0, pdf_contenido)
                wb.Close(False)

            except Exception as e:
                print(f"❌ Error al procesar Excel: {e}")
                raise  # Re-raise the exception to see full details
            finally:
                excel.Quit()

            # Combine caratula and content PDFs
            writer = PdfWriter()
            writer.append(PdfReader(ruta_caratula_generada))
            writer.append(PdfReader(pdf_contenido))

            with open(pdf_salida_sin_footer, "wb") as f:
                writer.write(f)

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

            agregar_imagen_a_paginas(pdf_salida_sin_footer, pdf_salida_con_footer, imagen_caratula)

            # Encrypt the final PDF
            reader_final = PdfReader(pdf_salida_con_footer)
            encrypted_writer = PdfWriter()

            for page in reader_final.pages:
                encrypted_writer.add_page(page)

            # Determinar la contraseña a usar
            if codigo_padded == "055":
                contraseña_pdf = config.CONTRA_055
            else:
                contraseña_pdf = contraseña_excel

            # Aplicar la encriptación
            try:
                encrypted_writer.encrypt(
                    user_password=contraseña_pdf,
                    owner_password=None,
                    use_128bit=True
                )

                # Guardar el archivo encriptado
                with open(pdf_salida_con_footer, "wb") as f:
                    encrypted_writer.write(f)

                print(f"✅ PDF encriptado correctamente: {pdf_salida_con_footer}")
            except Exception as e:
                print(f"❌ Error al encriptar el PDF: {e}")

            print(f"✅ PDF generado correctamente para cliente {codigo_padded}, archivo {archivo_excel}: {pdf_salida_con_footer}")

            # Clean up temporary files
            if os.path.exists(pdf_contenido):
                os.remove(pdf_contenido)
            if os.path.exists(pdf_salida_sin_footer):
                os.remove(pdf_salida_sin_footer)
            if os.path.exists(ruta_caratula_generada):
                os.remove(ruta_caratula_generada)

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
