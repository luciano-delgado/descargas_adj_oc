import pandas as pd


def leer_excel_pandas(ruta):
    
    """
    EXCEL  BASE EXTRAS DONDE ESTARAN LA FC, SE DEBE INDICAR LA RUTA DESEADA DONDE SE ENCUENTRA EL MISMO
    """
    
    try:
        df = pd.read_excel(ruta, sheet_name='Base', header=3)
        # lista_fc = df["OC"].tolist()
        df_filtrado = df[df['Fecha de factura'] == pd.Timestamp(2022, 6, 30)]
        df_filtrado.head(10)
        print(df_filtrado)
        l_facturas = df_filtrado['N° de Factura'].tolist()
        return l_facturas

    except: 


        return False
# ---------------
# ruta =  r"C:\Users\ldelgado\Desktop\Base Extras V1.7-I&P.xlsm"
# leer_excel_pandas(ruta)