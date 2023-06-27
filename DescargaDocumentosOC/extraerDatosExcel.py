import pandas as pd


def leer_excel_pandas(ruta=''):
    
    """EXCEL DONDE ESTARAN LA OC, SE DEBE INDICAR LA RUTA DESEADA DONDE SE ENCUENTRA EL MISMO"""
    
    try:
        df = pd.read_excel(r'C:\Users\ldelgado\Desktop\0_proyectos\AYF\descargas_adj_oc\descargas_adj_oc\DescargaDocumentosOC\oc_pruebas.xlsx')
        # print(df)

        lista_oc = df["OC"].tolist()
        print(f' \nLista de OC: {lista_oc}\n')
        
        
        return lista_oc

    except: 


        return False
# ---------------
# leer_excel_pandas('')