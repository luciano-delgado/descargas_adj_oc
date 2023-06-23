import os
from getpass import getuser
from datetime import datetime

now = datetime.now()

#USUARIO LOCAL
usuario = getuser()

# --> obtenemos la ruta del directorio actual.
directorio_raiz = f"C:/Users/{usuario}/Documents/Descarga_automática_legajos"

# --> Armamos la ruta del directorio de archivos.
directorio_archivos = directorio_raiz + "/" + "archivos"

# --> Armamos la ruta del Excel.
archivo_excel = directorio_archivos + "/" + "BaseLegajos.xlsx"
archivo_excel_trabajo = directorio_archivos + "/" + "BaseLegajosDescargar -- " + now.strftime('%m-%d-%Y') + ".xlsx"



