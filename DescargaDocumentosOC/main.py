from extraerDatosExcel import leer_excel_pandas
from descargar_doc import me23n


def main():
    
    l_oc = leer_excel_pandas()
    # l_oc = [4300012625]

    for oc in l_oc:
        me23n(0, str(oc))

if __name__ == '__main__':
    main()


# Notas:
# Caso 1: 4300012625 - 4300012630 (con duplicados) - 4300012648 (archivo raro) 
# Hay que tener SAP Abierto
# Hay que hacer pruebas de stress de Usuario
# Ver casos que tienen adjuntos duplicados
# Verificar que usuarios tiene acceso al FS