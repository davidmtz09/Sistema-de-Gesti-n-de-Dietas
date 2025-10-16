# utilerias.py
import os
import sys

def get_resource_path(relative_path):
    """Devuelve la ruta correcta de un archivo empaquetado por PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_file_path(file_name):
    """Devuelve la ruta de un archivo dado su nombre.

    Args:
        file_name (str): El nombre del archivo para obtener su ruta.

    Returns:
        str: Ruta completa del archivo.
    """
    return get_resource_path(file_name)
