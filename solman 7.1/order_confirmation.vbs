'@description:	Script uses transaction CRMD_ORDER, and changes status of CRM Incidents to "Confirmed"
'				Need to put list of CRM Incidents ID's to txt-file, and drag-&-drop script to SapGuiMenu
'@author:		Stepan Kulchanovskiy (proletarius)

navigateWnd = "wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0130" & _
			  "/subSUBSCREEN_1O_NAVIG:SAPLCRM_1O_LOCATOR:0110/ssubCRM_BUS_LOCATOR:SAPLBUS_LOCATOR:3101" & _
			  "/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202"
workWnd = "wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120" & _
		  "/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100"

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on" 
End If

Function BrowseForFile()
'@description: Browse for file dialog.
'@author: Jeremy England (SimplyCoded)
  BrowseForFile = CreateObject("WScript.Shell").Exec( _
    "mshta.exe ""about:<input type=file id=f>" & _
    "<script>resizeTo(0,0);f.click();new ActiveXObject('Scripting.FileSystemObject')" & _
    ".GetStandardStream(1).WriteLine(f.value);close();</script>""" _
  ).StdOut.ReadLine()
End Function

Set fso = CreateObject("Scripting.FileSystemObject")
filename = BrowseForFile()
Set f = fso.OpenTextFile(filename)
num = f.ReadLine

session.findById("wnd[0]/tbar[0]/okcd").text = "CRMD_ORDER"
session.findById("wnd[0]").sendVKey 0
session.findById(navigateWnd+"/cmbBUS_LOCA_SRCH01-SEARCH_TYPE").key = "1"
session.findById(navigateWnd+"/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3213/" & _
				"subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLCRM_1O_LOCATOR:6020/ctxtCRMT_SEARCH_LOC-OBJECT_ID").text = num
session.findById(navigateWnd+"/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3213" & _
				"/subSCREEN_3200_SEARCH_BUTTON_AREA:SAPLBUS_LOCATOR:3240/btnBUS_LOCA_SRCH01-GO").press
session.findById(navigateWnd+"/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3213" & _
				"/subSCREEN_3200_RESULT_AREA:SAPLBUS_LOCATOR:3250/cntlSCREEN_3210_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById(navigateWnd+"/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3213" & _
				"/subSCREEN_3200_RESULT_AREA:SAPLBUS_LOCATOR:3250/cntlSCREEN_3210_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById(workWnd+"/subSCR_1O_COMMON:SAPLCRM_1O_UI:3150/subSCR_1O_TT:SAPLCRM_1O_UI:2600/btnGV_TOGGTRANS").press
session.findById(workWnd+"/subSCR_1O_MAINTAIN:SAPLCRM_SALES_UI:0300/subMAINSCR0:SAPLCRM_SALES_UI:3010" & _
				"/subSCRAREA1:SAPLCRM_SALES_UI:3101/tabsTABSTRIP_HEADER/tabpT\SALS_HD01/ssubHEADER_DETAIL:SAPLCRM_SALES_UI:2140" & _
				"/subSCRAREA1:SAPLCRM_SALES_UI:7001/subSTATUS:SAPLCRM_SALES_UI:7003/subSTATUS:SAPLCRM_STATUS_UI:0130" & _
				"/cntlSTATUSCONT_0130/shellcont/shell").pressContextButton "BT_STATUS_EXTERN"
session.findById(workWnd+"/subSCR_1O_MAINTAIN:SAPLCRM_SALES_UI:0300/subMAINSCR0:SAPLCRM_SALES_UI:3010" & _
				"/subSCRAREA1:SAPLCRM_SALES_UI:3101/tabsTABSTRIP_HEADER/tabpT\SALS_HD01/ssubHEADER_DETAIL:SAPLCRM_SALES_UI:2140" & _
				"/subSCRAREA1:SAPLCRM_SALES_UI:7001/subSTATUS:SAPLCRM_SALES_UI:7003/subSTATUS:SAPLCRM_STATUS_UI:0130" & _
				"/cntlSTATUSCONT_0130/shellcont/shell").selectContextMenuItem "E0008"
session.findById("wnd[0]/tbar[0]/btn[11]").press

Do Until f.AtEndOfStream
	session.findById("wnd[0]/tbar[1]/btn[17]").press
	session.findById("wnd[1]/usr/ctxtGV_OBJECT_ID").text = f.ReadLine
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById(workWnd+"/subSCR_1O_COMMON:SAPLCRM_1O_UI:3150/subSCR_1O_TT:SAPLCRM_1O_UI:2600/btnGV_TOGGTRANS").press
	session.findById(workWnd+"/subSCR_1O_MAINTAIN:SAPLCRM_SALES_UI:0300/subMAINSCR0:SAPLCRM_SALES_UI:3010" & _
					"/subSCRAREA1:SAPLCRM_SALES_UI:3101/tabsTABSTRIP_HEADER/tabpT\SALS_HD01/ssubHEADER_DETAIL:SAPLCRM_SALES_UI:2140" & _
					"/subSCRAREA1:SAPLCRM_SALES_UI:7001/subSTATUS:SAPLCRM_SALES_UI:7003/subSTATUS:SAPLCRM_STATUS_UI:0130/cntlSTATUSCONT_0130" & _
					"/shellcont/shell").pressContextButton "BT_STATUS_EXTERN"
	session.findById(workWnd+"/subSCR_1O_MAINTAIN:SAPLCRM_SALES_UI:0300/subMAINSCR0:SAPLCRM_SALES_UI:3010" & _
					"/subSCRAREA1:SAPLCRM_SALES_UI:3101/tabsTABSTRIP_HEADER/tabpT\SALS_HD01/ssubHEADER_DETAIL:SAPLCRM_SALES_UI:2140" & _
					"/subSCRAREA1:SAPLCRM_SALES_UI:7001/subSTATUS:SAPLCRM_SALES_UI:7003/subSTATUS:SAPLCRM_STATUS_UI:0130" & _
					"/cntlSTATUSCONT_0130/shellcont/shell").selectContextMenuItem "E0008"
	session.findById("wnd[0]/tbar[0]/btn[11]").press
Loop

session.findById("wnd[0]/tbar[0]/btn[3]").press