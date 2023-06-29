from extraerDatosExcel import leer_excel_pandas
from descargar_doc import me23n
from sql import Hana
import tkinter as tk


def main():
    """
    29/6: Verlo con Joshu
    """
    ruta = r"C:\Users\ldelgado\Desktop\Base Extras V1.7-I&P.xlsm"
    l_facturas = leer_excel_pandas(ruta)
    # l_facturas=["00003A00000386","00011A00109524","00003A00000254"]
    l_fc_sin_oc = []
    
    for fc in l_facturas:

        datos = Hana("PRD")
        l_oc = datos.oc_de_factura(fc)
        
        if len(l_oc) ==0: 
            print(f'No se encontro OC para FC {fc}')
        l_fc_sin_oc.append(fc)
        
        for oc in l_oc:
            me23n(0, str(oc), str(fc))
            print(f"\t\t  OC {oc} - se descargo legajo para FC {fc}")



    return True 

# -----------------
# main()

## -- INTERFAZ GRAFICA -- ##
if __name__ == '__main__':
    
    root = tk.Tk()
    root.geometry('450x90')
    root.title('I&P Solutions - Descarga adjunta de OC v1.0')
    boton_leer = tk.Button(root,text="Iniciar descarga",command=lambda: main(),bg='lightblue',font =('calibri', 12)) 
    boton_leer.pack()
    root.mainloop()


# Notas:
# Caso 1: 4300012625 - 4300012630 (con duplicados) - 4300012648 (archivo raro) 
# Hay que tener SAP Abierto
# Verificar FC que no respecten nomenclatura

# 00003A00000386 Chalten -  00011A00109524  Marken - 00003A00000254 Mazza
# Ruta Base Extras: \\SCZ-G1-FS\Scienza_fs\Administracion_y_Finanzas\Cuentas_a_Pagar\000 HOME\02 Extras
# C:\Users\ldelgado\Desktop\Base Extras V1.7-I&P.xlsm