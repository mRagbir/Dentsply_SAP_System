Public Class SAP_Hierarchy

    Public Shared Function SapHierarchyTransactions(args As String())

        'initialize the params
        Dim Order_Number_Global As String = args(0).ToString() 'Order #
        Dim Material_Number_Global As String = args(1).ToString() 'Kit #
        Dim KitSerial_Global As String = args(2).ToString() 'Kit serial #
        Dim Storage_Location_Global As String = args(3).ToString() 'Storage location
        Dim TypeOfGR As String = args(4).ToString() 'Determines the selection in the dropdown
        Dim bIsSensorNeeded As Boolean = Convert.ToBoolean(args(5)) 'Determines is a sensor is needed in kit
        Dim bIsSpareCableNeeded As Boolean = Convert.ToBoolean(args(6)) 'Determines is a spare cable is needed in kit
        Dim bIsStarterKit As Boolean = Convert.ToBoolean(args(7)) 'Determines if a starter kit
        Dim bIsBasicKit As Boolean = Convert.ToBoolean(args(8)) 'Determines if a basic kit
        Dim bIsCableNeededOnSensor As Boolean = Convert.ToBoolean(args(9)) 'Determines is a cable is needed on the sensor
        Dim bIsChild1Checked As Boolean = Convert.ToBoolean(args(10)) 'does checkbox 1 need to be checked
        Dim bIsChild2Checked As Boolean = Convert.ToBoolean(args(11)) 'does checkbox 2 need to be checked
        Dim bIsChild3Checked As Boolean = Convert.ToBoolean(args(12)) 'does checkbox 3 need to be checked
        Dim Sensor_MaterialNumber As String = args(13)
        Dim Sensor_SerialNumber As String = args(14) 'Sensor serial #
        Dim Cable_MaterialNumber As String = args(15) 'Cable material #
        Dim CableSerial_Global As String = args(16) 'Cable serial #
        Dim SpareCableSerial_Global As String = args(17)
        Dim RemotePartNumber As String = args(18) 'Remote Part Number
        Dim Remote_SerialNumber As String = args(19)


        'Assign initial form fields that aren't conditional
        Dim OrderNumber = Order_Number_Global 'Order_Number_Global
        Dim KitParentMaterial = Material_Number_Global 'Material_Number_Global
        Dim KitSerial = KitSerial_Global 'KitSerial_Global
        Dim StorageLocation = Storage_Location_Global 'Storage_Location_Global


        If bIsSensorNeeded Then
            Dim SensorSerial = Sensor_SerialNumber 'SensorSerial_Global
        End If

        If CableSerial_Global <> String.Empty Then
            Dim CableSerial = CableSerial_Global 'CableSerial_Global
        End If

        If SpareCableSerial_Global <> String.Empty Then
            Dim SpareCableSerial = SpareCableSerial_Global 'SpareCableSerial_Global
        End If



        '  ***********  Assign SAP VALUES   **********



        'Connect to SAP GUI
        Try
            Dim App, Connection, session As Object

            Try

                Dim SapGuiAuto = GetObject("SAPGUI")
                App = SapGuiAuto.GetScriptingEngine
                Connection = App.Children(0)
                session = Connection.Children(0) ' grabs the 1st sap session window.

            Catch ex As Exception

                Console.WriteLine("Error : Cannot connect to SAP in Hierarchy transaction..!!")
                Return False
            End Try



            'Input text into SAP text boxes
            'session.createSession

            '~~~> t-code
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzus_hierarchy" 'navigate to the hierarchy screen
            session.findById("wnd[0]").sendVKey(0) ' enter
            session.findById("wnd[0]").resizeWorkingPane(123, 36, False) ' window size is optional

            'Initial values needed on Hierarchy page
            session.findById("wnd[0]/usr/ctxtP_AUFNR").Text = OrderNumber ' order number
            session.findById("wnd[0]/usr/ctxtP_MATNR").Text = KitParentMaterial 'Parent Material
            session.findById("wnd[0]/usr/ctxtP_SERNR").Text = KitSerial 'Parent serial number


            'Optional requirements depending on bool flags

            'Starter kit *************************
            'only needs Sensor Kit # And a Remote
            If bIsStarterKit Then
                'sensor
                session.findById("wnd[0]/usr/ctxtP_MATNR1").Text = Sensor_MaterialNumber ' sensor parent material number
                session.findById("wnd[0]/usr/ctxtP_SERNR1").Text = Sensor_SerialNumber 'sensor serial
                session.findById("wnd[0]/usr/chkP_CHECK1").Selected = bIsChild1Checked ' Determine if the child material needs to be checked
                'remote
                session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = RemotePartNumber ' remote part number
                session.findById("wnd[0]/usr/ctxtP_SERNR2").Text = Remote_SerialNumber 'remote serial number
                session.findById("wnd[0]/usr/chkP_CHECK2").Selected = bIsChild2Checked ' Determine if the child material needs to be checked

                'complete
                GoTo completeSapTransaction

            End If

            'Basic kit ************************
            'only needs a cable as child material, the sensor serial is the parent material serial number
            If bIsBasicKit Then
                session.findById("wnd[0]/usr/cmbP_CHECKP").Key = TypeOfGR 'Determines the selection in the dropdown
                session.findById("wnd[0]/usr/ctxtP_LGPRO").Text = "NY01"
                'cable
                session.findById("wnd[0]/usr/ctxtP_MATNR1").Text = Cable_MaterialNumber
                session.findById("wnd[0]/usr/ctxtP_SERNR1").Text = CableSerial_Global
                session.findById("wnd[0]/usr/chkP_CHECK1").Selected = bIsChild1Checked ' Determine if the child material needs to be checked
                'set focus on the dropdown
                session.findById("wnd[0]/usr/cmbP_CHECKP").SetFocus
                session.findById("wnd[0]").sendVKey(0) ' enter
                session.findById("wnd[0]/usr/cmbP_CHECKP").Key = TypeOfGR 'Picks 1st option for GR (needs to be picked twice)
                session.findById("wnd[0]/usr/ctxtP_LGPLA").Text = StorageLocation
                session.findById("wnd[0]/usr/cmbP_CHECKP").Key = TypeOfGR 'Picks 1st option for GR (needs to be picked again after selecting the storage location)

                'complete
                GoTo completeSapTransaction
            End If

            'Rma kit only needs a sensor and one cable
            If bIsCableNeededOnSensor And bIsSpareCableNeeded = False Then
                'sensor
                session.findById("wnd[0]/usr/ctxtP_MATNR1").Text = Sensor_MaterialNumber ' sensor parent depends on sensor identifier
                session.findById("wnd[0]/usr/ctxtP_SERNR1").Text = Sensor_SerialNumber 'sensor serial
                session.findById("wnd[0]/usr/chkP_CHECK1").Selected = bIsChild1Checked ' Determine if the child material needs to be checked
                'clear fields if populated
                session.findById("wnd[0]/usr/ctxtP_SERNR3").Text = String.Empty 'spare cable
                session.findById("wnd[0]/usr/txtP_REM3").Text = String.Empty 'spare cable remark
                session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = String.Empty   'spare cable material number
                'cable
                session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = Cable_MaterialNumber
                session.findById("wnd[0]/usr/ctxtP_SERNR2").Text = CableSerial_Global
                session.findById("wnd[0]/usr/chkP_CHECK2").Selected = bIsChild2Checked ' Determine if the child material needs to be checked
                'complete
                GoTo completeSapTransaction
            End If

            'Full kit needs sensor and 2 cables
            If bIsSpareCableNeeded Then
                'sensor
                session.findById("wnd[0]/usr/ctxtP_MATNR1").Text = Sensor_MaterialNumber ' sensor parent depends on sensor identifier
                session.findById("wnd[0]/usr/ctxtP_SERNR1").Text = Sensor_SerialNumber 'sensor serial
                session.findById("wnd[0]/usr/chkP_CHECK1").Selected = bIsChild1Checked ' Determine if the child material needs to be checked
                'cable 1
                session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = Cable_MaterialNumber
                session.findById("wnd[0]/usr/ctxtP_SERNR2").Text = CableSerial_Global
                session.findById("wnd[0]/usr/chkP_CHECK2").Selected = bIsChild2Checked ' Determine if the child material needs to be checked
                ' cable 2
                session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = Cable_MaterialNumber
                session.findById("wnd[0]/usr/ctxtP_SERNR3").Text = SpareCableSerial_Global 'spare cable
                session.findById("wnd[0]/usr/txtP_REM3").Text = "SPARE CABLE" 'spare cable remark
                session.findById("wnd[0]/usr/chkP_CHECK3").Selected = bIsChild3Checked ' Determine if the child material needs to be checked
                'complete
                GoTo completeSapTransaction
            End If



            'Double check if drop down is selected and  complete transaction
completeSapTransaction:

            session.findById("wnd[0]/usr/ctxtP_LGPRO").SetFocus 'select storage location window
            session.findById("wnd[0]/usr/ctxtP_LGPRO").caretPosition = 2
            session.findById("wnd[0]/usr/ctxtP_LGPRO").Text = "NY01" '1st storage location
            session.findById("wnd[0]/usr/cmbP_CHECKP").Key = TypeOfGR ' Picks 1st option for GR (needs to be picked twice)
            session.findById("wnd[0]").sendVKey(0) ' enter

            session.findById("wnd[0]/usr/ctxtP_LGPLA").Text = StorageLocation 'storage location
            session.findById("wnd[0]/usr/cmbP_CHECKP").Key = TypeOfGR  ' Picks 1st option for GR

            'Final execute!
            '
#Region "uncomment for production"



            'session.findById("wnd[0]/tbar[1]/btn[8]").press ' SAP F8 BUTTON - execute the transaction **********
            'session.findById("wnd[0]").sendVKey(0) 'Enter key


            ''***************************************************************************************************************************************
            ''If the transaction was completed successfully then the following serial field will be empty.
            ''If an error has occurred this field will still have text and the user will be prompted to look at the SAP screen to correct the error.
            ''Error usually is the order number entered doesn't match the material being received 99% of the time because all other issues_
            ''_were already checked ahead of time.


            'Dim checkforSapError = session.findById("wnd[0]/usr/ctxtP_SERNR1").Text

            'If checkforSapError <> "" Then
            '    Return False ' error was found
            'End If

#End Region

            'for test


            Return True ' no error was found

        Catch ex As Exception

            Console.WriteLine("SAP Transaction error : " + ex.Message)
            Return False
        End Try





#Region "Original code"

        ''9Ft***** Rma **** Only one cable
        'If SubKitSerial.Contains("L") And SubCableSerial.Contains("L") And SubSpareCableSerial = String.Empty Then
        '    session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = "B1209156"
        '    session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = ""
        '    GoTo assignSAPValues
        'End If

        ''9Ft*****Kit**** 2 Cables
        'If SubKitSerial.Contains("L") And SubCableSerial.Contains("L") And SubSpareCableSerial.Contains("L") Then
        '    session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = "B1209156"
        '    session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = "B1209156"
        '    GoTo assignSAPValues
        'End If

        ''3FT***** NORMAL ****** 2 cableS
        'If SubKitSerial.Contains("M") And SubCableSerial.Contains("M") And SubSpareCableSerial.Contains("M") Then
        '    session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = "B1209157"
        '    session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = "B1209157"
        '    GoTo assignSAPValues
        'End If

        ''3FT***** RMA ****** 1 cable
        'If SubKitSerial.Contains("M") And SubCableSerial.Contains("M") And SubSpareCableSerial = String.Empty Then
        '    session.findById("wnd[0]/usr/ctxtP_MATNR2").Text = "B1209157"
        '    session.findById("wnd[0]/usr/ctxtP_MATNR3").Text = "B1209157"
        '    GoTo assignSAPValues
        'End If

#End Region


#Region "sap_for_eliteFona"



        'sap_for_eliteFona:
        '        session.findById("wnd[0]/usr/ctxtP_AUFNR").Text = OrderNumber ' order number
        '        session.findById("wnd[0]/usr/ctxtP_MATNR").Text = KitParentMaterial 'Parent Material
        '        session.findById("wnd[0]/usr/cmbP_CHECKP").Key = "1"
        '        session.findById("wnd[0]/usr/ctxtP_SERNR").Text = KitSerial
        '        session.findById("wnd[0]/usr/ctxtP_LGPRO").Text = "ny01"
        '        session.findById("wnd[0]/usr/ctxtP_MATNR1").Text = "B1209155"
        '        session.findById("wnd[0]/usr/ctxtP_SERNR1").Text = CableSerial_Global
        '        session.findById("wnd[0]/usr/ctxtP_LGPRO").SetFocus
        '        session.findById("wnd[0]/usr/ctxtP_LGPRO").caretPosition = 4
        '        session.findById("wnd[0]").sendVKey(0)
        '        session.findById("wnd[0]/usr/ctxtP_LGPRO").SetFocus 'select storage location window
        '        session.findById("wnd[0]/usr/ctxtP_LGPRO").caretPosition = 2
        '        session.findById("wnd[0]/usr/ctxtP_LGPRO").Text = "ny01" '1st storage location
        '        session.findById("wnd[0]/usr/cmbP_CHECKP").Key = "1" ' Picks 1st option for GR (needs to be picked twice)
        '        session.findById("wnd[0]").sendVKey(0) ' enter
        '        session.findById("wnd[0]/usr/ctxtP_LGPLA").Text = StorageLocation 'storage location
        '        session.findById("wnd[0]/usr/cmbP_CHECKP").Key = "1" ' Picks 1st option for GR
        '        session.findById("wnd[0]/tbar[1]/btn[8]").press ' SAP F8 BUTTON
        '        session.findById("wnd[0]").sendVKey(0) 'Enter key

#End Region






    End Function



End Class
