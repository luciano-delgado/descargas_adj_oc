        
from decouple import config
from hdbcli import dbapi
import pandas as pd, urllib3, warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning);  warnings.simplefilter(action='ignore', category=FutureWarning);  warnings.simplefilter(action='ignore', category=UserWarning);


class Hana():
    
    def __init__(self, ambiente):

        if ambiente == 'QAS':
            self.__host = config('HANA_HOST_QAS')       
            self.__password = config('HANA_PASSWORD_QAS')
        elif ambiente == 'PRD':
            self.__host = config('HANA_HOST_PRD')       
            self.__password = config('HANA_PASSWORD_PRD')

        self.port = config('HANA_PORT')
        self.__user = config('HANA_USER')
        self.volumen = ""
        self.__connection = dbapi.connect(address=self.__host, port=self.port, user=self.__user, password=self.__password)
        self.__cursor =self. __connection.cursor()
        self.__cursor.execute("SET SCHEMA SAPABAP1")

    def oc_de_factura(self, factura: str ):

        self.q=f""" 
        SELECT 
        DISTINCT EBELN AS OC
        FROM EKBE 
        WHERE 
            XBLNR IN ('{factura}') 
            -- ('00003A00000386') -- #MAS DE UNA OC?
            -- XBLNR IN('00011A00109524')
            -- XBLNR IN('00003A00000254')
    """
        self.__cursor.execute(self.q)
        df = pd.read_sql_query(self.q,self.__connection)
        l_oc = df["OC"].tolist()

        return l_oc

# -----------------------------

# obj = Hana('PRD')
# oc = obj.oc_de_factura("00011A00109524")