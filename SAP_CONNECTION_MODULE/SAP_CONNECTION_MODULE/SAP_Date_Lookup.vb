Public Class SAP_Date_Lookup

    Public Shared Function Date_Lookup(_materialNumber As String, _kitSerials As List(Of String)) As Dictionary(Of String, String)

        Dim obj As New Dictionary(Of String, String)

        Dim foundDate As String
        Dim globalSerial As String


        On Error GoTo errorSapConnection
        'connect to SAP
        Dim App, Connection, session As Object
        Dim SapGuiAuto = GetObject("SAPGUI")
        App = SapGuiAuto.GetScriptingEngine
        Connection = App.Children(0)
        session = Connection.Children(0) ' grabs the 1st sap window.


        On Error GoTo errorSerial
        'check each serial
        For Each serial In _kitSerials

            globalSerial = serial


            '~~~> t-code
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nIQ03"
            session.findById("wnd[0]").sendVKey(0) ' enter

            session.findById("wnd[0]/usr/ctxtRISA0-MATNR").text = _materialNumber
            session.findById("wnd[0]/usr/ctxtRISA0-SERNR").text = serial
            session.findById("wnd[0]").sendVKey(0) ' enter



            'click the HISTORY button
            session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\05/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122B:SAPLITO0:1221/btn%_AUTOTEXT002").press

            session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectItem("          4", "1")
            session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").ensureVisibleHorizontalItem("          4", "1")
            session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").doubleClickItem("          4", "1")
            foundDate = session.findById("wnd[0]/usr/ctxtMKPF-BUDAT").Text

            'get manufacturing date and add to the collection
            obj.Add(serial, foundDate)

        Next
        'default back to mmbe screen just to get off the iq03 screen
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMMBE"
        session.findById("wnd[0]").sendVKey(0) ' enter


        'everything passed
        Return obj


errorSapConnection:
        obj.Add(19, "Error")
        Return obj

errorSerial:
        obj.Add(1, globalSerial)
        Return obj


    End Function



End Class
