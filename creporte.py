import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from calendar import monthrange
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import webbrowser
import os
import tempfile
from xlsxwriter.utility import xl_rowcol_to_cell
import menu_principal
from utilerias import get_file_path
from pathlib import Path

def nombre_a_numero_mes(nombre_mes):
    nombres_meses = {
        "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6,
        "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
    }
    return nombres_meses.get(nombre_mes, 0)  # Retorna 0 si el mes no es válido

def regresar():
    root.destroy()
    menu_principal.mostrar_menu()

def reporte():
    global root
    # Meses
    nombres_meses = [
        "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]

    def obtener_nombre_mes(numero_mes):
        """Devuelve el nombre del mes en español dado su número (1-12)."""
        return nombres_meses[numero_mes]

    def contar_alimentos_por_dia(df, filtro_columna_8, dia, mes, anho):
        conteos = {"DESAYUNO": 0, "COMIDA": 0, "CENA": 0}
        df_filtrado = df[df.iloc[:, 8] == filtro_columna_8]

        for _, row in df_filtrado.iterrows():
            periodo = row.iloc[10]
            tipo_alimento = row.iloc[9]

            try:
                fecha_inicio, fecha_fin = periodo.split(' - ')
                fecha_inicio = pd.to_datetime(fecha_inicio, dayfirst=True)
                fecha_fin = pd.to_datetime(fecha_fin, dayfirst=True)

                # Verificar si el día está dentro del rango de fechas
                if fecha_inicio <= pd.Timestamp(anho, mes, dia) <= fecha_fin:
                    conteos[tipo_alimento] += 1

            except ValueError:
                continue

        return conteos

    def actualizar_tablas():
        try:
            # Limpiar tablas anteriores
            for row in tabla_grande.get_children():
                tabla_grande.delete(row)
            
            # Obtener mes, año y servicio seleccionados
            mes = nombre_a_numero_mes(mes_var.get())
            anho = int(anho_var.get())
            servicio = servicio_var.get()
            
            # Obtener días del mes
            _, num_dias = monthrange(anho, mes)
            
            # Insertar días en la tabla 
            for dia in range(1, num_dias + 1):
                tabla_grande.insert('', 'end', values=(dia, 0, 0, 0, 0))
            
            # Insertar fila de total al final de la tabla
            tabla_grande.insert('', 'end', values=('Total', '', '', '', ''), tags=('total',))
            tabla_grande.tag_configure('total', background='#f0f0f0')

            # Cargar los archivos de Excel
            medicos_df = pd.read_excel(Path.home() / "ContPerFactura" / "medicos.xlsx")
            vales_df = pd.read_excel(Path.home() / "ContPerFactura" / "vales.xlsx")
            suplencias_df = pd.read_excel(Path.home() / "ContPerFactura" / "suplencias.xlsx")
            pbase_df = pd.read_excel(Path.home() / "ContPerFactura" / "pbase.xlsx")
            pcontrato_df = pd.read_excel(Path.home() / "ContPerFactura" / "pcontrato.xlsx")
            phrf_df = pd.read_excel(Path.home() / "ContPerFactura" / "phrf.xlsx")
            imssb_df = pd.read_excel(Path.home() / "ContPerFactura" / "imssb.xlsx")
            crum_df = pd.read_excel(Path.home() / "ContPerFactura" / "crum.xlsx")
            pasantes_df = pd.read_excel(Path.home() / "ContPerFactura" / "pasantes.xlsx")

            # Procesar cada día del mes seleccionado
            for dia in range(1, num_dias + 1):
                conteos_totales = {"DESAYUNO": 0, "COMIDA": 0, "CENA": 0}
                
                if servicio == "MR":
                    # Procesar medicos.xlsx para servicio MR
                    conteos_medicos = contar_alimentos_por_dia(medicos_df, "MR", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                    conteos_vales = contar_alimentos_por_dia(vales_df, "MR", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                    conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "MR", dia, mes, anho)
                    
                    # Sumar los conteos
                    for tipo in conteos_totales:
                        conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "MIP":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(medicos_df, "MIP", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "MIP", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "MIP", dia, mes, anho)
                    
                    # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "GAB":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(pbase_df, "GAB", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "GAB", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "GAB", dia, mes, anho)
                    
                    # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "GBB":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(pbase_df, "GBB", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "GBB", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "GBB", dia, mes, anho)
                    
                    # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]
                    
                if servicio == "JADB":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(pbase_df, "JADB", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "JADB", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "JADB", dia, mes, anho)
                    
                    # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "JANB":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(pbase_df, "JANB", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "JANB", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "JANB", dia, mes, anho)
                    
                    # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "GAC":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(pcontrato_df, "GAC", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "GAC", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "GAC", dia, mes, anho)
                    
                    # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "GBC":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(pcontrato_df, "GBC", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "GBC", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "GBC", dia, mes, anho)
                    
                    # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "JADC":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(pcontrato_df, "JADC", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "JADC", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "JADC", dia, mes, anho)
                    
                    # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "JANC":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(pcontrato_df, "JANC", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "JANC", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "JANC", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "GA-HRF":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(phrf_df, "GA-HRF", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "GA-HRF", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "GA-HRF", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "GB-HRF":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(phrf_df, "GB-HRF", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "GB-HRF", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "GB-HRF", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "JAD-HRF":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(phrf_df, "JAD-HRF", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "JAD-HRF", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "JAD-HRF", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]   

                if servicio == "JAN-HRF":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(phrf_df, "JAN-HRF", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "JAN-HRF", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "JAN-HRF", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]       

                if servicio == "GA-IMSSB":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(imssb_df, "GA-IMSSB", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "GA-IMSSB", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "GA-IMSSB", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]    

                if servicio == "GB-IMSSB":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(imssb_df, "GB-IMSSB", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "GB-IMSSB", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "GB-IMSSB", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]   

                if servicio == "JAD-IMSSB":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(imssb_df, "JAD-IMSSB", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "JAD-IMSSB", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "JAD-IMSSB", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]             

                if servicio == "JAN-IMSSB":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(imssb_df, "JAN-IMSSB", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "JAN-IMSSB", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "JAN-IMSSB", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                if servicio == "JAD-PASANTES":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(pasantes_df, "JAD-PASANTES", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "JAD-PASANTES", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "JAD-PASANTES", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]            
                    
                if servicio == "CRUM":
                    # Procesar medicos.xlsx para servicio MR
                        conteos_medicos = contar_alimentos_por_dia(crum_df, "CRUM", dia, mes, anho)
                    
                    # Procesar vales.xlsx para servicio MR
                        conteos_vales = contar_alimentos_por_dia(vales_df, "CRUM", dia, mes, anho)
                    
                    # Procesar suplencias.xlsx para servicio MR
                        conteos_suplencias = contar_alimentos_por_dia(suplencias_df, "CRUM", dia, mes, anho)

                        # Sumar los conteos
                        for tipo in conteos_totales:
                            conteos_totales[tipo] = conteos_medicos[tipo] + conteos_vales[tipo] + conteos_suplencias[tipo]

                # Insertar los conteos en la tabla
                tabla_grande.item(tabla_grande.get_children()[dia-1], values=(
                    dia,
                    conteos_totales["DESAYUNO"],
                    conteos_totales["COMIDA"],
                    conteos_totales["CENA"],
                    conteos_totales["DESAYUNO"] + conteos_totales["COMIDA"] + conteos_totales["CENA"]
                ))

            # Actualizar total de columnas
            totals = {"DESAYUNO": 0, "COMIDA": 0, "CENA": 0}
            for row_id in tabla_grande.get_children():
                values = tabla_grande.item(row_id)['values']
                if values[0] == 'Total':
                    continue
                totals["DESAYUNO"] += values[1]
                totals["COMIDA"] += values[2]
                totals["CENA"] += values[3]

            # Actualizar la fila de total
            tabla_grande.item(tabla_grande.get_children()[-1], values=(
                'Total',
                totals["DESAYUNO"],
                totals["COMIDA"],
                totals["CENA"],
                totals["DESAYUNO"] + totals["COMIDA"] + totals["CENA"]
            ))

        except ValueError:
            messagebox.showerror("Error", "Asegúrese de que el mes y el año estén seleccionados correctamente.")
        except FileNotFoundError as e:
            messagebox.showerror("Error", f"No se pudo encontrar el archivo: {str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"Ha ocurrido un error: {str(e)}")

    def exportar_pdf():
        # Obtener el servicio seleccionado
        servicio = servicio_var.get()
        mes = nombre_a_numero_mes(mes_var.get())
        anho = int(anho_var.get())
        nombre_mes = obtener_nombre_mes(mes)
        
        # Crear un PDF con los datos de la tabla
        file_name = get_file_path("reporte.pdf")
        c = canvas.Canvas(file_name, pagesize=letter)
        width, height = letter

            # Posición inicial
        y = height - 50

        # Encabezado del hospital
        hospital_header = """
        HOSPITAL GENERAL DE CUERNAVACA DR. JOSE G. PARRES
        SUBDIRECCIÓN MÉDICA
        DEPARTAMENTO DE NUTRICIÓN
        """
        c.setFont("Helvetica", 10)
        for line in hospital_header.split('\n'):
                # Centrar cada línea del encabezado
            text_width = c.stringWidth(line.strip(), "Helvetica", 10)
            x = (width - text_width) / 2  # Centrar el texto
            c.drawString(x, y, line.strip())
            y -= 12  # Espacio entre líneas

        # Espacio para el salto de línea
        y -= 20

            # Títulos de columnas
        columns = ['Día', 'DESAYUNO', 'COMIDA', 'CENA', 'TOTAL']
        col_widths = [50, 100, 100, 100, 100]
        row_height = 18  # Reducir el alto de la fila
        x_start = 50  # Reiniciar posición de x para columnas

            # Escribir el título del reporte
        title_text = f"Reporte de Alimentos de {servicio} - Fecha: {nombre_mes} del {anho}"
        text_width = c.stringWidth(title_text, "Helvetica-Bold", 12)
        x = (width - text_width) / 2  # Centrar el título
        c.setFont("Helvetica-Bold", 12)
        c.drawString(x, y, title_text)

        # Espacio para el salto de línea
        y -= 40

        # Dibujar encabezados
        c.setFont("Helvetica-Bold", 10)
        x = x_start
        for col, width in zip(columns, col_widths):
            c.drawString(x, y, col)
            x += width

        # Espacio debajo de los encabezados
        y -= row_height

        # Dibujar filas
        c.setFont("Helvetica", 10)
        for row_id in tabla_grande.get_children():
            row = tabla_grande.item(row_id)['values']
            x = x_start
            for value, width in zip(row, col_widths):
                c.drawString(x, y, str(value))
                x += width
            y -= row_height

        c.save()
        print(f"PDF generado: {file_name}")

        # Abrir el PDF generado en el navegador o visor predeterminado
        if os.name == 'nt':  # Para Windows
            webbrowser.open(file_name)
        elif os.name == 'posix':  # Para macOS y Linux
            webbrowser.open(f'file://{os.path.abspath(file_name)}')

    def exportar_excel():
        servicio = servicio_var.get()
        #mes = int(mes_var.get())
        mes = nombre_a_numero_mes(mes_var.get())
        anho = int(anho_var.get())
        nombre_mes = obtener_nombre_mes(mes)

        data = []
        for row_id in tabla_grande.get_children():
            row = tabla_grande.item(row_id)['values']
            data.append(row)

        df = pd.DataFrame(data, columns=['Día', 'DESAYUNO', 'COMIDA', 'CENA', 'TOTAL'])
        
        # Crear un archivo temporal en el directorio temporal del sistema
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_file:
            file_name = tmp_file.name
            # df.to_excel(file_name, index=False)

            # Guardar el DataFrame en un archivo Excel con xlsxwriter
            with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name=f'{nombre_mes} {anho}')

                # Obtener el libro y la hoja
                workbook = writer.book
                worksheet = writer.sheets[f'{nombre_mes} {anho}']

                # Definir el formato para los encabezados
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#800080',  # Color morado
                    'font_color': 'white',
                    'text_wrap': True,  # Ajuste de texto en la celda
                    'valign': 'center'
                })

                # Aplicar el formato a los encabezados
                for col_num, value in enumerate(df.columns.values):
                    cell = xl_rowcol_to_cell(0, col_num)
                    worksheet.write(cell, value, header_format)
                    worksheet.set_column(col_num, col_num, 15)  # Ajustar el ancho de las celdas

                # Congelar la fila del encabezado para que sea estática al hacer scroll
                worksheet.freeze_panes(1, 0)

        print(f"Archivo Excel generado: {file_name}")

        if os.name == 'nt':  
            os.startfile(file_name)
        elif os.name == 'posix':  
            webbrowser.open(f'file://{os.path.abspath(file_name)}')

    # Crear ventana principal
    root = tk.Tk()
    root.title('REPORTES DEL PERSONAL')
    root.geometry("1200x700")
    root.resizable(width=False, height=False)

    # Agregar estilos al Treeview
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview",
        background="#b7eefa",
        foreground="black",
        fieldbackground="#b7eefa"
    )
    style.map("Treeview",
        background=[('selected', '#bd1323')],
        foreground=[('selected', 'white')]
    )

    # Crear frame para la selección de mes, año y servicio
    frame_seleccion = tk.Frame(root)
    frame_seleccion.grid(row=0, column=0, padx=10, pady=10, sticky='ns')

    # Estilo para las etiquetas
    label_style = {"font": ("Arial", 12, "bold"), "padx": 10, "pady": 5}

    # Etiqueta principal en el centro
    tk.Label(frame_seleccion, text="HOSPITAL GENERAL DE CUERNAVACA DR. JOSE G. PARRES \n SUBDIRECCIÓN MÈDICA \n DEPARTAMENTO DE NUTRICIÓN", 
            font=('Arial', 10), anchor="center", justify="center", bg=frame_seleccion.cget("background")).grid(row=0, column=0, columnspan=4, padx=5, pady=10, sticky="ew")

    tk.Label(frame_seleccion, text="REPORTES ESPECÍFICOS DEL PERSONAL", 
            font=('Verdana', 13, "bold"), anchor="center", justify="center", bg=frame_seleccion.cget("background")).grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky="ew")

    # Etiqueta y menú desplegable para el servicio
    tk.Label(frame_seleccion, text="Seleccione Servicio:", **label_style).grid(row=2, column=0, padx=5, pady=10, sticky="e")
    servicio_var = tk.StringVar(value='MR')
    servicio_entry = ttk.Combobox(frame_seleccion, textvariable=servicio_var, values=["MR", "MIP", "GAB", "GBB", "JADB", "JANB", "GAC", "GBC", "JADC","JANC","GA-HRF","GB-HRF","JAD-HRF","JAN-HRF","GA-IMSSB","GB-IMSSB","JAD-IMSSB","JAN-IMSSB","JAD-PASANTES","CRUM"], state="readonly")
    servicio_entry.grid(row=2, column=1, padx=5, pady=10, sticky="w")

    # Etiqueta y menú desplegable para el mes
    tk.Label(frame_seleccion, text="Seleccione Mes:", **label_style).grid(row=3, column=0, padx=5, pady=10, sticky="e")
    mes_var = tk.StringVar(value='Enero')
    mes_entry = ttk.Combobox(frame_seleccion, textvariable=mes_var, values=[
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ], state="readonly")
    mes_entry.grid(row=3, column=1, padx=5, pady=10, sticky="w")

    # Etiqueta y menú desplegable para el año
    tk.Label(frame_seleccion, text="Seleccione Año:", **label_style).grid(row=4, column=0, padx=5, pady=10, sticky="e")
    anho_var = tk.StringVar(value='2024')
    anho_entry = ttk.Combobox(frame_seleccion, textvariable=anho_var, values=[str(year) for year in range(2024, 2052)], state="readonly")
    anho_entry.grid(row=4, column=1, padx=5, pady=10, sticky="w")

    # Estilo para los botones
    boton_style = {
        "font": ("Verdana", 10, "bold"),
        "width": 18,
        "height": 2,
        "fg": "white"
    }

    # Botón para actualizar las tablas
    btn_actualizar = tk.Button(frame_seleccion, text="ACTUALIZAR", command=actualizar_tablas, bg="#6a0dad", **boton_style)
    btn_actualizar.grid(row=5, columnspan=2, pady=10)

    # Botón para exportar a PDF
    btn_exportar_pdf = tk.Button(frame_seleccion, text="EXPORTAR A PDF", command=exportar_pdf, bg="#bd1323", **boton_style)
    btn_exportar_pdf.grid(row=6, columnspan=2, pady=10)

    #Boton para exportar a EXCEL
    btn_exportar_excel = tk.Button(frame_seleccion, text="EXPORTAR A EXCEL", command=exportar_excel, bg="green", **boton_style)
    btn_exportar_excel.grid(row=7, columnspan=2, pady=10)

    btn_regresar = tk.Button(frame_seleccion, text="REGRESAR AL MENÚ", command=regresar, bg="#6a0dad", **boton_style)
    btn_regresar.grid(row=8, columnspan=2, pady=10)

    # Crear frame para la tabla
    frame_tablas = tk.Frame(root)
    frame_tablas.grid(row=0, column=1, padx=10, pady=10, sticky='nswe')

    # Crear tabla con barra de desplazamiento vertical
    frame_tabla_grande = tk.Frame(frame_tablas)
    frame_tabla_grande.pack(side='left', fill='both', expand=True)

    tabla_grande = ttk.Treeview(frame_tabla_grande, columns=('Día', 'DESAYUNO', 'COMIDA', 'CENA', 'TOTAL'), show='headings', height=32)
    tabla_grande.pack(side='left', fill='both', expand=True)

    tabla_grande.heading('Día', text='Día')
    tabla_grande.heading('DESAYUNO', text='DESAYUNO')
    tabla_grande.heading('COMIDA', text='COMIDA')
    tabla_grande.heading('CENA', text='CENA')
    tabla_grande.heading('TOTAL', text='TOTAL')

    tabla_grande.column('Día', width=80, anchor='center')
    tabla_grande.column('DESAYUNO', width=120)
    tabla_grande.column('COMIDA', width=120)
    tabla_grande.column('CENA', width=120)
    tabla_grande.column('TOTAL', width=120)
    root.mainloop()