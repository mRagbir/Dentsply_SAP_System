Public Class Inspection

    Public Shared Function PerformSAPInspection(lotNumber As String, qty As String)

        Dim App, Connection, session As Object
        Dim SapGuiAuto As Object

        Try

            SapGuiAuto = GetObject("SAPGUI")
            App = SapGuiAuto.GetScriptingEngine
            Connection = App.Children(0)
            session = Connection.Children(0) ' grabs the 1st sap session window.

            '~~~> t-code
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nQE51N" 'navigate to the hierarchy screen
            session.findById("wnd[0]").sendVKey(0) ' enter

            session.findById("wnd[0]/usr/tabsTABSTRIP_OBJECT/tabpTAB_LOS").select
            session.findById("wnd[0]/usr/tabsTABSTRIP_OBJECT/tabpTAB_LOS/ssub%_SUBSCREEN_OBJECT:SAPLQEES:0510/ctxtQL_PRLOS-LOW").text = lotNumber

            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[1]/btn[8]").press
            session.findById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/ctxtQAQEE-SUMPLUS[9,0]").text = "p"
            session.findById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/ctxtQAQEE-SUMPLUS[9,0]").caretPosition = 1
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/ctxtQAQEE-SUMPLUS[9,1]").text = "p"
            session.findById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/ctxtQAQEE-SUMPLUS[9,1]").caretPosition = 1
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/ctxtQAQEE-SUMPLUS[9,2]").text = "p"
            session.findById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/ctxtQAQEE-SUMPLUS[9,2]").caretPosition = 1
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/ctxtQAQEE-SUMPLUS[9,3]").text = "p"
            session.findById("wnd[0]/usr/tabsEE_DATEN/tabpSISP/ssubSUB_EE_DATEN:SAPLQEEM:0202/tblSAPLQEEMSUMPLUS/ctxtQAQEE-SUMPLUS[9,3]").caretPosition = 1
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell").selectItem("          2", "Column01")
            session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem("          1", "Column01")
            session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickItem("          1", "Column01")
            session.findById("wnd[0]/usr/tabsUD_DATA/tabpPLMK/ssubSUB_UD_DATA:SAPMQEVA:0101/ssubUD_DATA:SAPMQEVA:1103/ctxtRQEVA-VCODE").text = "a"
            session.findById("wnd[0]/usr/tabsUD_DATA/tabpPLMK/ssubSUB_UD_DATA:SAPMQEVA:0101/ssubUD_DATA:SAPMQEVA:1103/ctxtRQEVA-VCODE").caretPosition = 1
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/tabsUD_DATA/tabpBB").select
            session.findById("wnd[0]/usr/tabsUD_DATA/tabpBB/ssubSUB_UD_DATA:SAPMQEVA:0102/txtRQEVA-VMENGE01").text = qty
            session.findById("wnd[0]/usr/tabsUD_DATA/tabpBB/ssubSUB_UD_DATA:SAPMQEVA:0102/ctxtRQEVA-QLGO_VM01").text = "sr04"
            session.findById("wnd[0]/usr/tabsUD_DATA/tabpBB/ssubSUB_UD_DATA:SAPMQEVA:0102/ctxtRQEVA-QLGO_VM01").setFocus
            session.findById("wnd[0]/usr/tabsUD_DATA/tabpBB/ssubSUB_UD_DATA:SAPMQEVA:0102/ctxtRQEVA-QLGO_VM01").caretPosition = 4
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

            Return True

        Catch ex As Exception

            Dim errorFound = ex.InnerException.ToString()
            Return False
        End Try











    End Function
End Class
