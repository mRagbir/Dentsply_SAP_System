Module Connect

    Sub Main()

        Dim OrderNumber = "" 'Order_Number_Global
        Dim KitParentMaterial = "" 'Material_Number_Global
        Dim KitSerial = "" 'KitSerial_Global
        Dim SensorSerial = "" 'SensorSerial_Global
        Dim CableSerial = "" 'CableSerial_Global
        Dim SpareCableSerial = "" 'SpareCableSerial_Global
        Dim SensorType = "" 'SensorType_Global
        Dim StorageLocation = "" 'Storage_Location_Global


        Dim SubKitSerial = "" 'SubKitSerial
        Dim SubSensorSerial = "" 'SubSensorSerial
        Dim SubCableSerial = "" 'SubCableSerial
        Dim SubSpareCableSerial = "" 'SubSpareCableSerial




        On Error GoTo errorHandler1
        'Connect to SAP GUI
        Dim App, Connection, session As Object
        Dim SapGuiAuto = GetObject("SAPGUI")
        App = SapGuiAuto.GetScriptingEngine
        Connection = App.Children(0)
        session = Connection.Children(0) ' grabs the 1st sap window.

        On Error GoTo errorHandler2
        'Input text into SAP text boxes
        'session.createSession

        '~~~> t-code
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzus_hierarchy"
        session.findById("wnd[0]").sendVKey(0) ' enter
        session.findById("wnd[0]").resizeWorkingPane(123, 36, False) ' window size is optional

        '~~~> if its a fona elite then skip to an alternate sap hierarchy input
        If KitParentMaterial = "B1211045" Or KitParentMaterial = "B1111045" Then
            GoTo sap_for_eliteFona
        End If



        '************************************************************************

        ' depending on cable identifier then populate sap cable parent text

        '*************************************************************************

        '6FT***** Rma ****** Only one cable
        If SubKitSerial.Contains("S") And SubCableSerial.Contains("S") And SpareCableSerial = String.Empty Then
            session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = "B1209155"
            session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = ""
            GoTo assignSAPValues
        End If

        '6FT normal Kit **** 2 Cables
        If SubKitSerial.Contains("S") And SubCableSerial.Contains("S") And SubSpareCableSerial.Contains("S") Then
            session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = "B1209155"
            session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = "B1209155"
            GoTo assignSAPValues
        End If


        '9Ft***** Rma **** Only one cable
        If SubKitSerial.Contains("L") And SubCableSerial.Contains("L") And SubSpareCableSerial = String.Empty Then
            session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = "B1209156"
            session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = ""
            GoTo assignSAPValues
        End If

        '9Ft*****Kit**** 2 Cables
        If SubKitSerial.Contains("L") And SubCableSerial.Contains("L") And SubSpareCableSerial.Contains("L") Then
            session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = "B1209156"
            session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = "B1209156"
            GoTo assignSAPValues
        End If

        '3FT***** NORMAL ****** 2 cableS
        If SubKitSerial.Contains("M") And SubCableSerial.Contains("M") And SubSpareCableSerial.Contains("M") Then
            session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = "B1209157"
            session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = "B1209157"
            GoTo assignSAPValues
        End If

        '3FT***** RMA ****** 1 cable
        If SubKitSerial.Contains("M") And SubCableSerial.Contains("M") And SubSpareCableSerial = String.Empty Then
            session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = "B1209157"
            session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = "B1209157"
            GoTo assignSAPValues
        End If


        '  ***********  Assign SAP VALUES   **********

assignSAPValues:

        session.findById("wnd[0]/usr/chkP_CHECK1").Selected = False ' uncheck the sensor child material
        session.findById("wnd[0]/usr/ctxtP_AUFNR").Text = OrderNumber ' order number
        session.findById("wnd[0]/usr/ctxtP_MATNR").Text = KitParentMaterial 'Parent Material
        session.findById("wnd[0]/usr/ctxtP_SERNR").Text = KitSerial 'Parent serial number
        session.findById("wnd[0]/usr/ctxtP_MATNR1").Text = SensorType ' sensor parent depends on sensor identifier
        session.findById("wnd[0]/usr/ctxtP_SERNR1").Text = SensorSerial 'sensor serial
        session.findById("wnd[0]/usr/ctxtP_SERNR2").Text = CableSerial 'cable 1


        'If spare cable text is NOT BLANK - then ASSIGN VALUES
        If SpareCableSerial <> String.Empty Then
            session.findById("wnd[0]/usr/ctxtP_SERNR3").Text = SpareCableSerial 'spare cable
            session.findById("wnd[0]/usr/txtP_REM3").Text = "SPARE CABLE" 'spare cable remark

            'if it is BLANK make all spare cable values in sap blank
        Else
            session.findById("wnd[0]/usr/ctxtP_SERNR3").Text = String.Empty 'spare cable
            session.findById("wnd[0]/usr/txtP_REM3").Text = String.Empty 'spare cable remark
            session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = String.Empty   'spare cable material number
        End If

        session.findById("wnd[0]/usr/ctxtP_LGPRO").SetFocus 'select storage location window
        session.findById("wnd[0]/usr/ctxtP_LGPRO").caretPosition = 2
        session.findById("wnd[0]/usr/ctxtP_LGPRO").Text = "ny01" '1st storage location
        session.findById("wnd[0]/usr/cmbP_CHECKP").Key = "1" ' Picks 1st option for GR (needs to be picked twice)
        session.findById("wnd[0]").sendVKey(0) ' enter

        session.findById("wnd[0]/usr/ctxtP_LGPLA").Text = StorageLocation 'storage location
        session.findById("wnd[0]/usr/cmbP_CHECKP").Key = "1" ' Picks 1st option for GR
        session.findById("wnd[0]/tbar[1]/btn[8]").press ' SAP F8 BUTTON
        session.findById("wnd[0]").sendVKey(0) 'Enter key

        GoTo assignValuesAfterSapInputs


#Region "sap_for_eliteFona"



sap_for_eliteFona:
        session.findById("wnd[0]/usr/ctxtP_AUFNR").Text = OrderNumber ' order number
        session.findById("wnd[0]/usr/ctxtP_MATNR").Text = KitParentMaterial 'Parent Material
        session.findById("wnd[0]/usr/cmbP_CHECKP").Key = "1"
        session.findById("wnd[0]/usr/ctxtP_SERNR").Text = KitSerial
        session.findById("wnd[0]/usr/ctxtP_LGPRO").Text = "ny01"
        session.findById("wnd[0]/usr/ctxtP_MATNR1").Text = "B1209155"
        session.findById("wnd[0]/usr/ctxtP_SERNR1").Text = CableSerial
        session.findById("wnd[0]/usr/ctxtP_LGPRO").SetFocus
        session.findById("wnd[0]/usr/ctxtP_LGPRO").caretPosition = 4
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtP_LGPRO").SetFocus 'select storage location window
        session.findById("wnd[0]/usr/ctxtP_LGPRO").caretPosition = 2
        session.findById("wnd[0]/usr/ctxtP_LGPRO").Text = "ny01" '1st storage location
        session.findById("wnd[0]/usr/cmbP_CHECKP").Key = "1" ' Picks 1st option for GR (needs to be picked twice)
        session.findById("wnd[0]").sendVKey(0) ' enter
        session.findById("wnd[0]/usr/ctxtP_LGPLA").Text = StorageLocation 'storage location
        session.findById("wnd[0]/usr/cmbP_CHECKP").Key = "1" ' Picks 1st option for GR
        session.findById("wnd[0]/tbar[1]/btn[8]").press ' SAP F8 BUTTON
        session.findById("wnd[0]").sendVKey(0) 'Enter key

#End Region


assignValuesAfterSapInputs:

        Exit Sub











errorHandler1:
        Console.WriteLine("Make sure you are logged into SAP and have at least 3 windows open")
        Exit Sub

errorHandler2:
        Console.WriteLine("Error populating SAP text boxes , check SAP screen for error details")
        Exit Sub


    End Sub

End Module
