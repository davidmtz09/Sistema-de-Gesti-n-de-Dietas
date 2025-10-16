from tkinter import *
from tkinter import filedialog, messagebox
import pandas as pd
from pandastable import Table, TableModel
import menu_principal
from utilerias import get_file_path
from pathlib import Path

def regresar():
    global app
    app.destroy()
    menu_principal.mostrar_menu()

def plantilla():
    class Tables(Tk):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.frame = Frame(self)
            self.frame.pack(fill=BOTH, expand=1)

            self.data = self.lectura(Path.home() / "ContPerFactura" / "plantilla.xlsx")
            self.table = Table(
                self.frame,
                dataframe=self.data,
                showtoolbar=False,
                showstatusbar=True,
                editable=True
            )
            self.table.show()

            # Frame para los botones
            self.button_frame = Frame(self)
            self.button_frame.pack()

            self.load_button = Button(self.button_frame, text="Reemplazar Plantilla", command=self.cargar_archivo, bg="purple", fg="white")
            self.load_button.pack(side=LEFT, padx=10)

            self.save_button = Button(self.button_frame, text="Guardar Cambios", command=self.guardar, bg="purple", fg="white")
            self.save_button.pack(side=LEFT, padx=10)

            self.regresar_button = Button(self.button_frame, text="Regresar al Menú", command=regresar, bg="purple", fg="white")
            self.regresar_button.pack(side=LEFT, padx=10)

        def lectura(self, filename):
            data = pd.read_excel(filename)
            return data

        def guardar(self):
            data = self.table.model.df
            # Cambia la ruta a la carpeta ContPerFactura
            data.to_excel(Path.home() / "ContPerFactura" / "plantilla.xlsx", index=False)
            messagebox.showinfo("Información", "Archivo guardado correctamente")

        def cargar_archivo(self):
            filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if filename:
                self.data = self.lectura(filename)
                self.table.updateModel(TableModel(dataframe=self.data))
                self.table.redraw()
                messagebox.showinfo("Información", f"Archivo {filename} cargado correctamente")

    # Instanciar la clase Tables después de definirla
    global app
    app = Tables()
    app.mainloop()