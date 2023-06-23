import win32com.client as win32
import pythoncom
import win32com.client
import os, shutil, time


def me23n(sesionsap, oc):

     """DESCARGAR ADJUNTOS POR CADA OC """
     
     ruta_destino = f"U:/Aplicaciones_procesos/COMPRAS/descarga_adj_oc/{oc}"
     
     flag = True
     if flag:     
          pythoncom.CoInitialize()
          SapGuiAuto = win32com.client.GetObject('SAPGUI')
          if not type(SapGuiAuto) == win32com.client.CDispatch:
               return
          application = SapGuiAuto.GetScriptingEngine
          if not type(application) == win32com.client.CDispatch:
               SapGuiAuto = None
               return
          connection = application.Children(0)
          if not type(connection) == win32com.client.CDispatch:
               application = None
               SapGuiAuto = None
               return
          session = connection.Children(sesionsap)
          if not type(session) == win32com.client.CDispatch:
               connection = None
               application = None
               SapGuiAuto = None
               return

     try:
          session.findById("wnd[0]/tbar[0]/okcd").text = "/NME23N"
          session.findById("wnd[0]").sendVKey(0)
          session.findById("wnd[0]/tbar[1]/btn[17]").press()
          session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = f"{oc}" 
          session.findById("wnd[1]").sendVKey(0)
          session.findById("wnd[0]/titl/shellcont/shell").pressButton("%GOS_TOOLBOX")
          session.findById("wnd[0]/shellcont/shell").pressButton("VIEW_ATTA")
          try: 
               os.mkdir(ruta_destino)
          except FileExistsError: 
               print(f'Ya existe la carpeta para la oc {oc}')
               return True
          for i in range (0,50): 
               try:
                    session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = f"{i}" 
                    session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").pressToolbarButton("%ATTA_EXPORT")
                    time.sleep(0.2)
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta_destino
                    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
                    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 32
                    session.findById("wnd[1]").sendVKey (0)
                    time.sleep(0.2)
               except: 
                    # Esta parte es por adjunto duplicado, intento darle al boton cerrar 
                    try: session.findById("wnd[1]/tbar[0]/btn[12]").press()
                    except: pass
                    break

          try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
          except: pass
          session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
          session.findById("wnd[0]").sendVKey(0)
          print(f'Se proceso correctamente oc {oc}')
          
          
          return True

     except Exception as e:
          print(f'Error sin capturar en oc {oc}{e}')
         
         
          return False
     

     
# # ---------------------------------------------------------------------------

# l_oc = ["4300012630","4300012629"]

# for oc in l_oc:
#      me23n(0, oc)
