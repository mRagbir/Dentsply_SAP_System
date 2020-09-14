Public Class CableKits


    Public Shared Function SAP_CableKits(_materialNumber As String, _orderNumber As String, _location As String, _serialList As List(Of String)) As Collection

        Dim var As New Collection 'used to hold error messages


        Dim serialNumber As String
        Dim boolKey As Boolean


        On Error GoTo SAPerrorHandler_Connection
        'Connect to SAP GUI
        Dim App, Connection, session As Object
        Dim SapGuiAuto = GetObject("SAPGUI")
        App = SapGuiAuto.GetScriptingEngine
        Connection = App.Children(0)
        session = Connection.Children(0) ' grabs the 1st sap window.

        For Each pair In _serialList


            serialNumber = pair


            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzus_hierarchy"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtP_AUFNR").Text = _orderNumber
            session.findById("wnd[0]/usr/ctxtP_MATNR").Text = _materialNumber ' <~~~ PARENT MATERIAL NUMBER
            session.findById("wnd[0]/usr/ctxtP_SERNR").Text = serialNumber 'Parent serial number
            session.findById("wnd[0]/usr/cmbP_CHECKP").Key = 2
            session.findById("wnd[0]/usr/ctxtP_LGPRO").Text = "NY01"
            session.findById("wnd[0]/usr/cmbP_CHECKP").SetFocus
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/cmbP_CHECKP").Key = 2
            session.findById("wnd[0]/usr/ctxtP_LGPLA").Text = _location
            session.findById("wnd[0]/usr/cmbP_CHECKP").Key = 2
            session.findById("wnd[0]/tbar[1]/btn[8]").press

            session.findById("wnd[0]").sendVKey(0) 'Enter key

        Next
        'success
        var.Add(0, "Passed")
        Return var

SAPerrorHandler_Connection:

        var.Add(3, "Error")

        Return var


    End Function




End Class
