from extraerDatosExcel import leer_excel_pandas
from descargar_doc import me23n


def main():
    l_oc = leer_excel_pandas()

    for oc in l_oc:
        me23n(0, str(oc))

if __name__ == '__main__':
    main()

# Notas:
# Hay que hacer pruebas de stress
# Ver casos que tienen adjuntos duplicados
# Verificar que usuarios tiene acceso al FS