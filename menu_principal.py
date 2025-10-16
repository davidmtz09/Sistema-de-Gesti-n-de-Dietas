import tkinter as tk
from tkinter import ttk
from PIL import ImageTk, Image
import rmedicos
import rpbase
import rpcontrato
import rvales
import rsuplencias
import cconteo
import ccontmed
import creporte
import cconcenper
import lplantilla
import rphrf
import imss_bi
import crum
import rpasantes
from utilerias import get_file_path
from rmedicos import medicos

def registro_medicos():
    root.destroy()
    rmedicos.medicos()

def registro_base():
    root.destroy()
    rpbase.pbase()

def registro_contrato():
    root.destroy()
    rpcontrato.pcontrato()

def registro_vales():
    root.destroy()
    rvales.vales()

def registro_suplencias():
    root.destroy()
    rsuplencias.suplencias()

def r_hrf():
    root.destroy()
    rphrf.phrf()

def r_crum():
    root.destroy()
    crum.crum()

def r_pasantes():
    root.destroy()
    rpasantes.pasantes()

def imss_b():
    root.destroy()
    imss_bi.imss()

def mostrar_conteo():
    root.destroy()
    cconteo.conteo()

def mostrar_contmed():
    root.destroy()
    ccontmed.contmed()

def mostrar_reporte():
    root.destroy()
    creporte.reporte()

def mostrar_concenper():
    root.destroy()
    cconcenper.concenper()

def mostrar_plantilla():
    root.destroy()
    lplantilla.plantilla()

def mostrar_menu():
    global root
    # Configuración principal de la ventana
    root = tk.Tk()
    root.title("Cont Personal Factura")
    root.geometry("1200x600")
    root.resizable(width=False, height=False)

    ruta_logo1 = get_file_path("logo1.png")
    ruta_logo2 = get_file_path("logo2.png")

    # Crear un estilo para los botones
    style = ttk.Style(root)
    style.theme_use('default')

    # Configurar el estilo para los botones
    style.configure("TButton",
                    font=('Arial', 10, 'bold'),
                    padding=6)

    style.map("TButton",
            background=[('active', '#660066'), ('!active', '#800080')],
            foreground=[('active', 'white'), ('!active', 'white')])

    # Frame principal
    main_frame = tk.Frame(root)
    main_frame.pack(fill='both', expand=True, padx=20, pady=20)

    # Frame para la parte superior que contendrá las imágenes y el texto principal
    top_frame = tk.Frame(main_frame)
    top_frame.grid(row=0, column=0, columnspan=5, pady=(10, 0))  # Añadido padding superior

    # Ajustar tamaño de las columnas y filas para centrar contenido
    main_frame.columnconfigure(0, weight=1)
    top_frame.columnconfigure(2, weight=1)  # Centrar la columna del texto principal

    # Cargar y mostrar la primera imagen
    img = Image.open(ruta_logo1)
    new_img = img.resize((300, 100))
    render = ImageTk.PhotoImage(new_img)
    img1 = tk.Label(top_frame, image=render)
    img1.image = render  # Mantener una referencia a la imagen
    img1.grid(row=0, column=0, padx=(0, 50), pady=10)  # Añadido padding derecho

    # Cargar y mostrar la segunda imagen
    img2 = Image.open(ruta_logo2)
    new_img2 = img2.resize((300, 100))
    render2 = ImageTk.PhotoImage(new_img2)
    img2_label = tk.Label(top_frame, image=render2)
    img2_label.image = render2
    img2_label.grid(row=0, column=4, padx=(50, 0), pady=10)

    # Etiqueta principal en el centro
    label1 = tk.Label(top_frame, text="SERVICIOS DE SALUD DE MORELOS \n H G DE CUERNAVACA DR. JOSE G. PARRES \n SUBDIRECCIÓN MÈDICA", 
                    font=('Arial', 12), anchor="center", justify="center", bg=top_frame.cget("background"))
    label1.grid(row=0, column=2, padx=20, pady=10)

    # Etiqueta de departamento de nutrición
    label2 = tk.Label(main_frame, text="DEPARTAMENTO DE NUTRICIÓN",
                    font=('Arial', 12, 'bold'), anchor="center", justify="center", bg=main_frame.cget("background"))
    label2.grid(row=1, column=0, columnspan=5, pady=10)

    # Etiqueta de menú
    label3 = tk.Label(main_frame, text="MENÚ",
                    font=('Arial', 16, 'bold'), anchor="center", justify="center", foreground="#722FC9", bg=main_frame.cget("background"))
    label3.grid(row=2, column=0, columnspan=5, pady=10)

    # Botón de registrar médicos
    button_med = ttk.Button(main_frame, text="REGISTRAR MEDICOS (MR,MIP)", style="TButton", command=registro_medicos)
    button_med.grid(row=3, column=0, columnspan=5, pady=30)

    # Frame para las secciones
    sections_frame = tk.Frame(main_frame)
    sections_frame.grid(row=4, column=0, columnspan=5)

    # Ajustar tamaño de las columnas para centrar contenido
    sections_frame.columnconfigure(0, weight=1)
    sections_frame.columnconfigure(1, weight=1)
    sections_frame.columnconfigure(2, weight=1)
    sections_frame.columnconfigure(3, weight=1)
    sections_frame.columnconfigure(4, weight=1)

    # Sección PERSONAL DE BASE
    lbl_base = tk.Label(sections_frame, text="PERSONAL DE BASE", font=('Arial', 12), bg=sections_frame.cget("background"))
    lbl_base.grid(row=0, column=0, padx=20, pady=10)

    button_gab = ttk.Button(sections_frame, text="REGISTRAR", style="TButton", command=registro_base)
    button_gab.grid(row=1, column=0, padx=20, pady=10)

    # Sección PERSONAL DE CONTRATO
    lbl_contrato = tk.Label(sections_frame, text="PERSONAL DE CONTRATO", font=('Arial', 12), bg=sections_frame.cget("background"))
    lbl_contrato.grid(row=0, column=1, padx=20, pady=10)

    button_gac = ttk.Button(sections_frame, text="REGISTRAR", style="TButton", command=registro_contrato)
    button_gac.grid(row=1, column=1, padx=20, pady=10)

    #Seccion IMSS BIENESTAR
    lbl_imss = tk.Label(sections_frame, text="IMSS Bienestar", font=('Arial', 12), bg=sections_frame.cget("background"))
    lbl_imss.grid(row=2, column=0, padx=20, pady=10)
    button_imss = ttk.Button(sections_frame, text="REGISTRAR", style="TButton", command=imss_b)
    button_imss.grid(row=3, column=0, padx=20, pady=10)

    #Seccion HRF
    lbl_hrf = tk.Label(sections_frame, text="HRF", font=('Arial', 12), bg=sections_frame.cget("background"))
    lbl_hrf.grid(row=2, column=1, padx=20, pady=10)
    button_hrf = ttk.Button(sections_frame, text="REGISTRAR", style="TButton", command=r_hrf)
    button_hrf.grid(row=3, column=1, padx=20, pady=10)

    button_vales = ttk.Button(sections_frame, text="VALES", style="TButton", command=registro_vales)
    button_vales.grid(row=2, column=2, padx=20, pady=10)

    button_supl = ttk.Button(sections_frame, text="SUPLENCIAS", style="TButton", command=registro_suplencias)
    button_supl.grid(row=2, column=3, padx=20, pady=10)

    button_supl = ttk.Button(sections_frame, text="CRUM", style="TButton", command=r_crum)
    button_supl.grid(row=0, column=2, padx=20, pady=10)

    button_supl = ttk.Button(sections_frame, text="PASANTES", style="TButton", command=r_pasantes)
    button_supl.grid(row=1, column=2, padx=20, pady=10)

    # Sección 4
    button_conteo = ttk.Button(sections_frame, text="CONTEO", style="TButton", command=mostrar_conteo)
    button_conteo.grid(row=0, column=3, padx=20, pady=10)

    button_mrmi = ttk.Button(sections_frame, text="M.R. M.I.P.", style="TButton", command=mostrar_contmed)
    button_mrmi.grid(row=1, column=3, padx=20, pady=10)

    button_rep = ttk.Button(sections_frame, text="REPORTE", style="TButton", command=mostrar_reporte)
    button_rep.grid(row=1, column=4, padx=20, pady=10)

    # Sección 5
    button_concenper = ttk.Button(sections_frame, text="CONCENPER", style="TButton", command=mostrar_concenper)
    button_concenper.grid(row=0, column=4, padx=20, pady=10)

    button_plantilla = ttk.Button(sections_frame, text="PLANTILLA", style="TButton", command=mostrar_plantilla)
    button_plantilla.grid(row=2, column=4, padx=20, pady=10)

    # Iniciar el loop de la aplicación
    root.mainloop()