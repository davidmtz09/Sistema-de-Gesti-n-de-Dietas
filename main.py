import os
import shutil
from pathlib import Path
import menu_principal
from utilerias import get_resource_path

# Obtener la ruta del directorio de usuario y crear la carpeta ContPerFactura si no existe
user_folder = Path.home() / "ContPerFactura"
if not user_folder.exists():
    user_folder.mkdir(parents=True, exist_ok=True)  # Crea la carpeta automáticamente

# Definir rutas de archivos dentro de la carpeta ContPerFactura
plantilla_path = user_folder / "plantilla.xlsx"
reporte_path = user_folder / "reporte.pdf"
imss_path = user_folder / "imssb.xlsx"
medicos_path = user_folder / "medicos.xlsx"
pbase_path = user_folder / "pbase.xlsx"
pcontrato_path = user_folder / "pcontrato.xlsx"
phrf_path = user_folder / "phrf.xlsx"
suplencias_path = user_folder / "suplencias.xlsx"
vales_path = user_folder / "vales.xlsx"
scrum_path = user_folder / "crum.xlsx"
pasantes_path = user_folder / "pasantes.xlsx"

# Verificar si los archivos ya existen y si no, copiarlos desde su ubicación original
def copy_if_not_exists(src, dst):
    if not dst.exists():
        shutil.copy(src, dst)

# Copiar archivos desde la carpeta del proyecto
copy_if_not_exists(get_resource_path("plantilla.xlsx"), plantilla_path)
copy_if_not_exists(get_resource_path("reporte.pdf"), reporte_path)
copy_if_not_exists(get_resource_path("imssb.xlsx"), imss_path)
copy_if_not_exists(get_resource_path("medicos.xlsx"), medicos_path)
copy_if_not_exists(get_resource_path("pbase.xlsx"), pbase_path)
copy_if_not_exists(get_resource_path("pcontrato.xlsx"), pcontrato_path)
copy_if_not_exists(get_resource_path("phrf.xlsx"), phrf_path)
copy_if_not_exists(get_resource_path("suplencias.xlsx"), suplencias_path)
copy_if_not_exists(get_resource_path("vales.xlsx"), vales_path)
copy_if_not_exists(get_resource_path("crum.xlsx"), scrum_path)
copy_if_not_exists(get_resource_path("pasantes.xlsx"), pasantes_path)

menu_principal.mostrar_menu()