import win32com.client as win32
import pythoncom
import win32com.client
import time

def va01_2(sesionsap, ped_ext, dispone,
           fecha_entrega, lista_id_productos, cantidades,
           convenio, lista_mat_cl, tipo_dispone, ce_asistencial, canal, sector):

     #----------------------------#
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
#----------------------------#
     descargar_documentos_adjuntos = False

     try:
          # session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA01"
          # session.findById("wnd[0]").sendVKey(0)
          # session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = "ZTER"
          # session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = "SC10"
          # session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = canal
          # session.findById("wnd[0]/usr/ctxtVBAK-SPART").text = sector
          # session.findById("wnd[0]/usr/ctxtVBAK-SPART").setFocus()
          # session.findById("wnd[0]/usr/ctxtVBAK-SPART").caretPosition = 2
          # session.findById("wnd[0]").sendVKey(0)
          # session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = ped_ext
          # session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = "10000029"
          # session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").text = dispone
          # session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").setFocus()
          # session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 8
          # session.findById("wnd[0]").sendVKey(0)
          # session.findById("wnd[0]").sendVKey(0)


          session.findById("wnd[0]/tbar[0]/okcd").text = "/NME23N"
          session.findById("wnd[0]").sendVKey(0)
          session.findById("wnd[0]/tbar[1]/btn[17]").press
          session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = oc_descargar ## "4300012625"
          session.findById("wnd[1]").sendVKey(0)
          session.findById("wnd[0]/titl/shellcont/shell").pressButton "%GOS_TOOLBOX"
          session.findById("wnd[0]/shellcont/shell").pressButton "VIEW_ATTA"
          session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "0"
          session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").doubleClickCurrentCell
          session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").currentCellRow = 1
          session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "1"
          session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").doubleClickCurrentCell
          session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").currentCellRow = 2
          session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "2"
          session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").doubleClickCurrentCell
          session.findById("wnd[1]/tbar[0]/btn[0]").press
          session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
          session.findById("wnd[0]").sendVKey(0)
          # VALIDACION PARA LAS FARMACIAS QUE SE ENCUENTRAN BLOQUEADAS POR MOTIVO DE VACACIONES
          # try:
          #      info_dispone_bloqueado = session.findById("wnd[0]/sbar").text
          #      lista_palabras_claves = info_dispone_bloqueado.split()
          #      print(info_dispone_bloqueado)
          #      if "bloqueo" in lista_palabras_claves:
          #           session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").text = "84005541"
          #           bandera_dispone_bloqueado = True
          #      session.findById("wnd[0]").sendVKey(0)
          #      session.findById("wnd[0]").sendVKey(0)
          #      session.findById("wnd[0]").sendVKey(0)
          # except:
          #      session.findById("wnd[0]").sendVKey(0)
          #      session.findById("wnd[0]").sendVKey(0)
          # validar si el dispone esta bloqueado


          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").key = "NT"
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT").text = fecha_entrega
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT").setFocus()
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT").caretPosition = 2
          # session.findById("wnd[0]").sendVKey(0)

          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02").select()

          # try:
          #      for i in range(0, len(lista_id_productos)):
          #           session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,{i}]").text = lista_id_productos[i]
          #           session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,{i}]").text = cantidades[i]
          #           session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-KDMAT[6,{i}]").text = lista_mat_cl[i]
          #           session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[12,{i}]").text = "TOSD"
          #           session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[65,{i}]").text = "ALMA"
          #           # session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[12,{i}]").text = "HT01"
          #           # session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[65,{i}]").text = "1104"
          #           session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[65,{i}]").setFocus()
          #           session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[65,{i}]").caretPosition = 4
          #           session.findById("wnd[0]").sendVKey(0)
          # except Exception as a:
          #      print("Linea 72 VA01-2", a)
          #      return

          # session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()

          # # --- LOGICA PARA PEDIDOS EN CENTROS ASISNTENCIALES --- #
          # if tipo_dispone == 0:
          #      session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").select()
          #      session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,4]").key = "ZD"
          #      session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").text = dispone
          #      session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").setFocus()
          #      session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").caretPosition = 8
          #      session.findById(r"wnd[0]").sendVKey(0)
          #      session.findById(r"wnd[0]").sendVKey(0)
          #      session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,8]").key = "ZC"
          #      session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").text = str(ce_asistencial)
          #      session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").setFocus()
          #      session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,8]").caretPosition = 8
          #      session.findById("wnd[0]").sendVKey(0)
          #      session.findById("wnd[0]").sendVKey(0)

          # # ----------------------------------------------------- #
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10").select()
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\12").select()
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13").select()
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZCONVENIO").text = convenio
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZCONVENIO").caretPosition = 2
          # session.findById("wnd[0]").sendVKey(0)
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZTURNO").text = "MAN"
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZTURNO").setFocus()
          # session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZTURNO").caretPosition = 3
          # #Hace falta agregar una excepcion en este punto?
          # session.findById("wnd[0]/tbar[0]/btn[11]").press()
          # time.sleep(3)
          # ped = session.findById("wnd[0]/sbar").text
          # ped_final = ped[21:29]
          # return ped_final, bandera_dispone_bloqueado

     except Exception as e:
          print("Linea 111 VA01-2", e)
          return -1, descargar_documentos_adjuntos

# va01_2(0, "123", "84000977", "23123", "23123", "23123", "23123" ,"23123" ,"23123", "23123", "02", "06")