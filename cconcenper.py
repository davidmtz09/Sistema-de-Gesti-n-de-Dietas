import tkinter as tk
from tkinter import ttk
from calendar import monthrange
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os
import tempfile
import webbrowser
from xlsxwriter.utility import xl_rowcol_to_cell
import menu_principal
from utilerias import get_file_path
from dateutil.parser import parse
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

def concenper():
    global root
    # Meses
    nombres_meses = [
        "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]

    def obtener_nombre_mes(numero_mes):
        """Devuelve el nombre del mes en español dado su número (1-12)."""
        return nombres_meses[numero_mes]

    def contar_alimentos_por_dia(mes, anho):
        user_folder = Path.home() / "ContPerFactura"  # Ruta a la carpeta ContPerFactura
        archivos_excel = [
            user_folder / 'medicos.xlsx',
            user_folder / 'pbase.xlsx',
            user_folder / 'pcontrato.xlsx',
            user_folder / 'suplencias.xlsx',
            user_folder / 'vales.xlsx',
            user_folder / 'phrf.xlsx',
            user_folder / 'imssb.xlsx',
            user_folder / 'pasantes.xlsx',
            user_folder / 'crum.xlsx'
        ]
        
        conteo_por_dia = {i: {'DESAYUNO': 0, 'COMIDA': 0, 'CENA': 0} for i in range(1, monthrange(anho, mes)[1] + 1)}

        for archivo in archivos_excel:
            df = pd.read_excel(archivo, header=None)
            
            for index, row in df.iloc[1:].iterrows():  # Comienza desde la segunda fila
                periodo = row[10].strip()  # Columna 10: Periodo_alimento
                tipo_alimento = row[9]  # Columna 9: Tipo_alimento

                if pd.isna(periodo) or pd.isna(tipo_alimento):
                    continue
                
                if not isinstance(periodo, str) or ' - ' not in periodo:
                    print(f"Formato de periodo incorrecto en archivo {archivo}, fila {index}: {periodo}")
                    continue

                try:
                    fecha_inicio, fecha_fin = periodo.split(' - ')
                    fecha_inicio = parse(fecha_inicio, dayfirst=True)
                    fecha_fin = parse(fecha_fin, dayfirst=True)

                    if pd.isna(fecha_inicio) or pd.isna(fecha_fin):
                        print(f"Fecha no válida en archivo {archivo}, fila {index}: {periodo}")
                        continue

                except Exception as e:
                    print(f"Error al procesar la fecha en archivo {archivo}, fila {index}: {periodo} - {e}")
                    continue

                for dia in pd.date_range(start=fecha_inicio, end=fecha_fin):
                    if dia.month == mes and dia.year == anho:
                        conteo_por_dia[dia.day][tipo_alimento] += 1

        return conteo_por_dia

    # Función para contar alimentos específicos para "Guardia A", "Guardia B", "J.A. Diurna", etc.
    def contar_alimentos(tipo_guardia, archivo, valor_guardia=None):
        df = pd.read_excel(archivo, header=0)  # Usar header=0 para omitir la primera fila como encabezado
        conteo_total = 0

        if valor_guardia:
            df_filtrado = df[df['MR/MIP'] == valor_guardia]  # Filtrar por el valor de la columna 'MR/MIP'
        else:
            df_filtrado = df

        for index, row in df_filtrado.iterrows():
            periodo = row['Fecha_alimento']  # Usar el nombre de la columna si ya tienes encabezados
            tipo_alimento = row['Tipo_alimento']

            if pd.isna(periodo) or pd.isna(tipo_alimento):
                continue

            # Verificar si es una cadena y si tiene el formato correcto
            if not isinstance(periodo, str) or ' - ' not in periodo:
                print(f"Formato de periodo incorrecto en archivo {archivo}, fila {index}: {periodo}")
                continue

            try:
                fecha_inicio, fecha_fin = periodo.split(' - ')
                fecha_inicio = parse(fecha_inicio, dayfirst=True)
                fecha_fin = parse(fecha_fin, dayfirst=True)

                if pd.isna(fecha_inicio) or pd.isna(fecha_fin):
                    print(f"Fecha no válida en archivo {archivo}, fila {index}: {periodo}")
                    continue

            except Exception as e:
                print(f"Error al procesar la fecha en archivo {archivo}, fila {index}: {periodo} - {e}")
                continue

            # Sumar el conteo para los días dentro del rango
            for dia in pd.date_range(start=fecha_inicio, end=fecha_fin):
                if dia.month == mes and dia.year == anho:
                    conteo_total += 1

        return conteo_total

    def actualizar_tablas():
        # Limpiar tablas anteriores
        for row in tabla_grande.get_children():
            tabla_grande.delete(row)
        for row in tabla_chica.get_children():
            tabla_chica.delete(row)
        
        # Obtener mes y año seleccionados
        global mes, anho
        mes = nombre_a_numero_mes(mes_var.get())
        anho = int(anho_var.get())
        
        # Obtener el conteo por día
        conteo = contar_alimentos_por_dia(mes, anho)
        
        # Insertar días en la tabla grande con conteos
        for dia in range(1, monthrange(anho, mes)[1] + 1):
            desayunos = conteo[dia]['DESAYUNO']
            comidas = conteo[dia]['COMIDA']
            cenas = conteo[dia]['CENA']
            total = desayunos + comidas + cenas
            tabla_grande.insert('', 'end', values=(dia, desayunos, comidas, cenas, total))
        
        # Insertar fila de total al final de la tabla grande
        total_desayunos = sum([conteo[dia]['DESAYUNO'] for dia in conteo])
        total_comidas = sum([conteo[dia]['COMIDA'] for dia in conteo])
        total_cenas = sum([conteo[dia]['CENA'] for dia in conteo])
        total_general = total_desayunos + total_comidas + total_cenas
        tabla_grande.insert('', 'end', values=('Total', total_desayunos, total_comidas, total_cenas, total_general), tags=('total',))

        # Contar alimentos para cada tipo específico
        # Define la ruta a la carpeta ContPerFactura
        user_folder = Path.home() / "ContPerFactura"

        # Ahora modificamos el diccionario conteos_guardia
        conteos_guardia = {
            'Guardia A': contar_alimentos('Guardia A', user_folder / 'pbase.xlsx', 'GAB') + contar_alimentos('Guardia A', user_folder / 'pcontrato.xlsx', 'GAC')+ contar_alimentos('Guardia A', user_folder / 'phrf.xlsx', 'GA-HRF')+ contar_alimentos('Guardia A', user_folder / 'imssb.xlsx', 'GA-IMSSB'),
            'Guardia B': contar_alimentos('Guardia B', user_folder / 'pbase.xlsx', 'GBB') + contar_alimentos('Guardia B', user_folder / 'pcontrato.xlsx', 'GBC')+ contar_alimentos('Guardia B', user_folder / 'phrf.xlsx', 'GB-HRF')+ contar_alimentos('Guardia B', user_folder / 'imssb.xlsx', 'GB-IMSSB'),
            'J.A. Diurna': contar_alimentos('J.A. Diurna', user_folder / 'pbase.xlsx', 'JADB') + contar_alimentos('J.A. Diurna', user_folder / 'pcontrato.xlsx', 'JADC')+ contar_alimentos('J.A. Diurna', user_folder / 'phrf.xlsx', 'JAD-HRF')+ contar_alimentos('J.A. Diurna', user_folder / 'imssb.xlsx', 'JAD-IMSSB')+ contar_alimentos('J.A. Diurna', user_folder / 'pasantes.xlsx', 'JAD-PASANTES'),
            'J.A. Nocturna': contar_alimentos('J.A. Nocturna', user_folder / 'pbase.xlsx', 'JANB') + contar_alimentos('J.A. Nocturna', user_folder / 'pcontrato.xlsx', 'JANC')+ contar_alimentos('J.A. Nocturna', user_folder / 'phrf.xlsx', 'JAN-HRF')+ contar_alimentos('J.A. Nocturna', user_folder / 'imssb.xlsx', 'JAN-IMSSB'),
            'CRUM': contar_alimentos('CRUM', user_folder / 'crum.xlsx', 'CRUM'),
            'Internos': contar_alimentos('Internos', user_folder / 'medicos.xlsx', 'MIP'),
            'Residentes': contar_alimentos('Residentes', user_folder / 'medicos.xlsx', 'MR'),
            'Suplencias': contar_alimentos('Suplencias', user_folder / 'suplencias.xlsx'),
            'Vales': contar_alimentos('Vales', user_folder / 'vales.xlsx')
        }

        # Insertar filas en la tabla chica
        total_guardia = 0
        for tipo, conteo in conteos_guardia.items():
            if tipo == 'Total':
                tabla_chica.insert('', 'end', values=(tipo, total_guardia))
            else:
                tabla_chica.insert('', 'end', values=(tipo, conteo))
                total_guardia += conteo

        # Insertar fila de total al final de la tabla chica
        tabla_chica.insert('', 'end', values=('Total', total_guardia), tags=('total',))

        # Aplicar estilo para la fila de total
        tabla_grande.tag_configure('total', background='#f0f0f0')
        tabla_chica.tag_configure('total', background='#f0f0f0')

    def exportar_pdf():
        # Obtener el mes y el año
        mes = nombre_a_numero_mes(mes_var.get())
        anho = int(anho_var.get())
        nombre_mes = obtener_nombre_mes(mes)
        
        # Crear un PDF con los datos de las tablas
        file_name = get_file_path("reporte.pdf")
        c = canvas.Canvas(file_name, pagesize=letter)
        width, height = letter

        # Encabezado del hospital centrado
        def dibujar_encabezado_hospital(c, y):
            hospital_header = """
            HOSPITAL GENERAL DE CUERNAVACA DR. JOSE G. PARRES
            SUBDIRECCIÓN MÉDICA
            DEPARTAMENTO DE NUTRICIÓN
            """
            c.setFont("Helvetica", 10)
            for line in hospital_header.split('\n'):
                text_width = c.stringWidth(line.strip(), "Helvetica", 10)
                x = (width - text_width) / 2  # Centrar el texto
                c.drawString(x, y, line.strip())
                y -= 12  # Espacio entre líneas
            return y

        # Función para dibujar una tabla
        def dibujar_tabla(tabla, y_start, title):
            y = y_start
            x = 50
            col_width = 100
            row_height = 20
            max_rows_per_page = int((y - 80) / row_height)  # Calcular cuántas filas caben en la página
            rows_dibujadas = 0
            
            # Escribir el título
            c.setFont("Helvetica-Bold", 12)
            c.drawString(x, y, title)
            y -= 20

            # Títulos de columnas
            columns = tabla['columns']
            c.setFont("Helvetica-Bold", 10)
            for col in columns:
                c.drawString(x, y, col)
                x += col_width
            
            # Espacio debajo de los encabezados
            y -= 20
            x = 50
            
                # Dibujar filas
            c.setFont("Helvetica", 10)
            for row in tabla['data']:
                if rows_dibujadas >= max_rows_per_page:
                    # Si se llena la página, añadir una nueva página
                    c.showPage()
                    y = height - 50
                    y = dibujar_encabezado_hospital(c, y)
                    c.setFont("Helvetica-Bold", 12)
                    c.drawString(50, y, title)  # Reescribir el título de la tabla en la nueva página
                    y -= 20
                    c.setFont("Helvetica", 10)
                    for col in columns:
                        c.drawString(x, y, col)
                        x += col_width
                    y -= 20
                    x = 50
                    rows_dibujadas = 0  # Reiniciar el conteo de filas en la nueva página
                
                for value in row:
                    c.drawString(x, y, str(value))
                    x += col_width
                y -= row_height
                x = 50
                rows_dibujadas += 1
            
            return y  # Devolver la posición Y para el próximo bloque

        # Obtener los datos de la tabla grande
        tabla_grande_data = {
            'columns': ['Día', 'DESAYUNO', 'COMIDA', 'CENA', 'TOTAL'],
            'data': [tabla_grande.item(row_id)['values'] for row_id in tabla_grande.get_children()]
        }
        
        # Obtener los datos de la tabla chica
        tabla_chica_data = {
            'columns': ['GUARDIA', 'TOTAL'],
            'data': [tabla_chica.item(row_id)['values'] for row_id in tabla_chica.get_children()]
        }
        
        # Escribir el título general para la primera página
        y_position = height - 50
        y_position = dibujar_encabezado_hospital(c, y_position)
        c.setFont("Helvetica-Bold", 12)
        title_text = f"Reporte Mensual de Alimentos del Personal - {nombre_mes} {anho}"
        text_width = c.stringWidth(title_text, "Helvetica-Bold", 12)
        x = (width - text_width) / 2  # Centrar el título
        c.drawString(x, y_position, title_text)

        # Espacio adicional para asegurar que el contenido de la tabla no se amontone
        y_position -= 40

        # Dibujar tabla grande en la primera página
        y_position = dibujar_tabla(tabla_grande_data, y_position, "CONTEO GENERAL POR DIA")

        c.showPage() # Cambiar a una nueva página

        # Escribir el título y el encabezado para la segunda página
        y_position = height - 50
        y_position = dibujar_encabezado_hospital(c, y_position)
        c.setFont("Helvetica-Bold", 12)
        title_text = f"Reporte de Alimentos - {nombre_mes} {anho}"
        text_width = c.stringWidth(title_text, "Helvetica-Bold", 12)
        x = (width - text_width) / 2  # Centrar el título
        c.drawString(x, y_position, title_text)

        # Espacio adicional para asegurar que el contenido de la tabla no se amontone
        y_position -= 40

        # Dibujar tabla chica en la segunda página
        y_position = dibujar_tabla(tabla_chica_data, y_position, "CONTEO GENERAL POR GUARDIA")

        c.save()
        print(f"PDF generado: {file_name}")

        # Abrir el PDF generado en el navegador o visor predeterminado
        if os.name == 'nt':  # Para Windows
            webbrowser.open(file_name)
        elif os.name == 'posix':  # Para macOS y Linux
            webbrowser.open(f'file://{os.path.abspath(file_name)}')

    #Funcion para exportar a Excel
    def exportar_excel():
        # Obtener mes y año seleccionados
        mes = nombre_a_numero_mes(mes_var.get())
        anho = int(anho_var.get())
        nombre_mes = obtener_nombre_mes(mes)

        # Obtener datos de la tabla grande
        data_grande = []
        for row_id in tabla_grande.get_children():
            row = tabla_grande.item(row_id)['values']
            data_grande.append(row)

        # Crear DataFrame para la tabla grande
        df_grande = pd.DataFrame(data_grande, columns=['Día', 'DESAYUNO', 'COMIDA', 'CENA', 'TOTAL'])

        # Obtener datos de la tabla chica
        data_chica = []
        for row_id in tabla_chica.get_children():
            row = tabla_chica.item(row_id)['values']
            data_chica.append(row)

        # Crear DataFrame para la tabla chica
        df_chica = pd.DataFrame(data_chica, columns=['GUARDIA', 'TOTAL'])

        # Crear un archivo temporal en el directorio temporal del sistema
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_file:
            file_name = tmp_file.name

            # Guardar ambos DataFrames en el mismo archivo Excel
            with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
                df_grande.to_excel(writer, sheet_name='Conteo Por Día', index=False)
                df_chica.to_excel(writer, sheet_name='Conteo Por Guardia', index=False)

                # Obtener el libro y las hojas
                workbook = writer.book
                worksheet_grande = writer.sheets['Conteo Por Día']
                worksheet_chica = writer.sheets['Conteo Por Guardia']

                # Definir el formato para los encabezados
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#800080',  # Color morado
                    'font_color': 'white',
                    'text_wrap': True,  # Ajuste de texto en la celda
                    'valign': 'center'
                })

                # Aplicar el formato a los encabezados en la tabla grande
                for col_num, value in enumerate(df_grande.columns.values):
                    cell = xl_rowcol_to_cell(0, col_num)
                    worksheet_grande.write(cell, value, header_format)
                    worksheet_grande.set_column(col_num, col_num, 15)  # Ajustar el ancho de las celdas

                # Aplicar el formato a los encabezados en la tabla chica
                for col_num, value in enumerate(df_chica.columns.values):
                    cell = xl_rowcol_to_cell(0, col_num)
                    worksheet_chica.write(cell, value, header_format)
                    worksheet_chica.set_column(col_num, col_num, 15) 

        print(f"Archivo Excel generado: {file_name}")

        # Abrir el archivo Excel generado en el sistema
        if os.name == 'nt':  
            os.startfile(file_name)
        elif os.name == 'posix':  
            webbrowser.open(f'file://{os.path.abspath(file_name)}')

    # Crear ventana principal
    root = tk.Tk()
    root.title('CONCENPER')
    root.geometry("1200x700")

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

    # Crear frame para selección de mes y año
    frame_seleccion = tk.Frame(root)
    frame_seleccion.pack(pady=10)

    # Etiqueta principal en el centro
    tk.Label(frame_seleccion, text="HOSPITAL GENERAL DE CUERNAVACA DR. JOSE G. PARRES \n SUBDIRECCIÓN MÈDICA \n DEPARTAMENTO DE NUTRICIÓN", 
            font=('Arial', 10), anchor="center", justify="center", bg=frame_seleccion.cget("background")).grid(row=0, column=0, columnspan=4, padx=5, pady=10, sticky="ew")

    # Etiqueta del título
    tk.Label(frame_seleccion, text="REPORTE MENSUAL DE ALIMENTOS DEL PERSONAL", 
            font=('Verdana', 12, "bold"), anchor="center", justify="center", bg=frame_seleccion.cget("background")).grid(row=1, column=0, columnspan=4, padx=5, pady=10, sticky="ew")

    # Etiqueta y menú desplegable para el mes
    tk.Label(frame_seleccion, text="Seleccione Mes:").grid(row=2, column=1, padx=5, pady=5, sticky="e")
    mes_var = tk.StringVar(value='Enero')
    mes_entry = ttk.Combobox(frame_seleccion, textvariable=mes_var, values=[
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ], state="readonly")
    mes_entry.grid(row=2, column=2, padx=5, pady=5, sticky="w")

    # Etiqueta y menú desplegable para el año
    tk.Label(frame_seleccion, text="Seleccione Año:").grid(row=3, column=1, padx=5, pady=5, sticky="e")
    anho_var = tk.StringVar(value='2024')
    anho_entry = ttk.Combobox(frame_seleccion, textvariable=anho_var, values=[str(year) for year in range(2024, 2052)], state="readonly")
    anho_entry.grid(row=3, column=2, padx=5, pady=5, sticky="w")

    # Botón para actualizar las tablas
    btn_actualizar = tk.Button(frame_seleccion, text="Actualizar", command=actualizar_tablas, bg="#6a0dad", fg="white")
    btn_actualizar.grid(row=4, column=0, padx=5, pady=10, sticky="ew")

    #Botón para exportar a pdf
    update_button = tk.Button(frame_seleccion, text="Exportar a PDF", command=exportar_pdf, bg="#bd1323", fg="white")
    update_button.grid(row=4, column=1, padx=5, pady=10, sticky="ew")

    # Botón para exportar a excel
    update_button = tk.Button(frame_seleccion, text="Exportar a EXCEL", command=exportar_excel, bg="green", fg="white")
    update_button.grid(row=4, column=2, padx=5, pady=10, sticky="ew")

    regresar_button = tk.Button(frame_seleccion, text="Regresar al Menú", command=regresar, bg="#6a0dad", fg="white")
    regresar_button.grid(row=4, column=3, padx=5, pady=10, sticky="ew")

    # Crear frame para tablas
    frame_tablas = tk.Frame(root)
    frame_tablas.pack(pady=10)

    # Crear tabla grande con barra de desplazamiento vertical
    frame_tabla_grande = tk.Frame(frame_tablas)
    frame_tabla_grande.pack(side='left')

    tabla_grande = ttk.Treeview(frame_tabla_grande, columns=('Día', 'DESAYUNO', 'COMIDA', 'CENA', 'TOTAL'), show='headings', height=32)
    tabla_grande.pack(side='left')

    scrollbar_y = tk.Scrollbar(frame_tabla_grande, orient='vertical', command=tabla_grande.yview)
    scrollbar_y.pack(side='right', fill='y')

    tabla_grande.configure(yscrollcommand=scrollbar_y.set)

    tabla_grande.heading('Día', text='Día')
    tabla_grande.heading('DESAYUNO', text='DESAYUNO')
    tabla_grande.heading('COMIDA', text='COMIDA')
    tabla_grande.heading('CENA', text='CENA')
    tabla_grande.heading('TOTAL', text='TOTAL')

    tabla_grande.column('Día', width=60, anchor='center')
    tabla_grande.column('DESAYUNO', width=100)
    tabla_grande.column('COMIDA', width=100)
    tabla_grande.column('CENA', width=100)
    tabla_grande.column('TOTAL', width=100)

    # Crear tabla chica
    frame_tabla_chica = tk.Frame(frame_tablas)
    frame_tabla_chica.pack(side='right', padx=20)

    tabla_chica = ttk.Treeview(frame_tabla_chica, columns=('GUARDIA', 'TOTAL'), show='headings', height=10)
    tabla_chica.pack()

    tabla_chica.heading('GUARDIA', text='GUARDIA')
    tabla_chica.heading('TOTAL', text='TOTAL')

    tabla_chica.column('GUARDIA', width=150)
    tabla_chica.column('TOTAL', width=100)

    # Insertar filas en la tabla chica con valores fijos
    guardias = ['Guardia A', 'Guardia B', 'J.A. Diurna', 'J.A. Nocturna', 'CRUM', 'Internos', 'Residentes', 'Suplencias', 'Vales', 'Total']
    for guardia in guardias:
        tabla_chica.insert('', 'end', values=(guardia, ''))

    # Ejecutar la aplicación
    root.mainloop()