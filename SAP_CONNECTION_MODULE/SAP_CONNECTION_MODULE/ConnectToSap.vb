
Imports System.Text

Public Module ConnectToSap

    Dim sb As New StringBuilder

    ''' <summary>
    ''' API needs 20 arguments to complete SAP hierarchy transaction
    ''' Order_Number_Global As String,'Order #
    ''' Material_Number_Global As String,'Kit #
    ''' KitSerial_Global As String, 'Kit serial #
    ''' Storage_Location_Global As String, 'Storage location
    ''' TypeOfGR As String,'Determines the selection in the dropdown
    ''' bIsSensorNeeded As Boolean,'Determines is a sensor is needed in kit
    ''' bIsSpareCableNeeded As Boolean,'Determines is a spare cable is needed in kit
    ''' bIsStarterKit As Boolean,'Determines if a starter kit
    ''' bIsBasicKit As Boolean,'Determines if a basic kit
    ''' bIsCableNeededOnSensor As Boolean,'Determines is a cable is needed on the sensor
    ''' bIsChild1Checked As Boolean,'does checkbox 1 need to be checked
    ''' bIsChild2Checked As Boolean,'does checkbox 2 need to be checked
    ''' bIsChild3Checked As Boolean,'does checkbox 3 need to be checked
    ''' Optional Sensor_MaterialNumber As String = "", 'Sensor Material #
    ''' Optional Sensor_SerialNumber As String = "", 'Sensor serial #
    ''' Optional Cable_MaterialNumber As String = "", 'Cable material #
    ''' Optional CableSerial_Global As String = "", 'Cable serial #
    ''' Optional SpareCableSerial_Global As String = "", 'Spare cable serial #
    ''' Optional RemotePartNumber As String = "",'Remote Part Number
    ''' Optional Remote_SerialNumber As String = ""
    ''' </summary>

    Public Sub Main()

#Region "All these tests passed"


        'TEST RMA KIT
        'Dim results = SAP_Hierarchy.SapHierarchyTransactions("7525043", "B1218051", "JS20002958", "FG-S11",
        '                                                     True, False, False, False, True, "B1218200", "25017924",
        '                                                     "B1209155", "US31599",,)
        'TEST FULL KIT
        'Dim results2 = SAP_Hierarchy.SapHierarchyTransactions("7525042", "B1218000", "IS20012120",
        '                                                     "FG-S11", True, True, False, False, True, "B1218200",
        '                                                     "25017922", "B1209155", "US774255", "US759499",)
        'test starter kit
        'Dim results3 = SAP_Hierarchy.SapHierarchyTransactions("7525044", "100007345", "QS2002250", "FG-S11",
        '                                                     True, False, True, False, False, "B1218000", "IS20012120",,,, "125055125055")



        'Dim serials As New List(Of String)
        'serials.Add("IS20013432")
        'serials.Add("IS20013431")
        'serials.Add("IS20013418")
        'serials.Add("IS20013414")
        'serials.Add("IS20013402")

        'Dim results = SAP_CONNECTION.SAP_Date_Lookup.Date_Lookup("B1218000", serials)

        'Dim results = PerformSapActions("B1218000", "IS20013206", False)
        'Dim results = PerformSapActions("100007349", "QL1001054", True)

        'Dim serials As New Dictionary(Of String, Boolean)
        'serials.Add("US45254", True)
        'serials.Add("US45255", True)
        'serials.Add("US45256", False)
        'serials.Add("US45257", False)
        'serials.Add("US45258", False)
        'serials.Add("US45259", True)
        'serials.Add("US45260", False)

        'CableKits.SAP_CableKits("B1209120", "12545", "FG-S11", serials)


        'double check the sap call **************
        'TEST BASIC KIT
        'Dim results4 = SAP_Hierarchy.SapHierarchyTransactions("123456", "Kit_PN", "KIT_serial", "FG-S11", True, False,
        '                                                     False, True, True, "SENSOR_MATERIAL", "SENSOR_SERIAL", "B1209155", "CABLE1_SERIAL",,)
        'Dim argsForTest As String() = {
        '    "1",
        '    "7647004",
        '    "B1318001",
        '    "IL30001934",
        '    "FG-S11",
        '    "1",
        '    "true",
        '    "true",
        '    "false",
        '    "false",
        '    "true",
        '    "false",
        '    "true",
        '    "true",
        '    "B1318200",
        '    "26005083",
        '    "B1209156",
        '    "UL56984",
        '    "UL78545",
        '    String.Empty,'"RemotePartNumber"
        '    String.Empty '"Remote_SerialNumber"
        '}


        'args = argsForTest


#End Region

        Dim args As String() = System.Environment.GetCommandLineArgs()
        args = args.Where(Function(s) s <> args(0)).ToArray

        ' Dim person = (FirstName:="", LastName:="")

        Try

            'determine if all arguments are supplied before calling API
            ' Dim argsLength = argsIntoArray.Length
            Dim argsLength = args.Length

            If argsLength > 0 Then

                'parse the args and use the first argument as the switch
                Dim WhichApiToUse = args(0).ToString()

                Select Case WhichApiToUse

                    Case "PerformHierarchy"
                        'Call Hierarchy
                        PerformHierarchy(args)

                    Case "PerformConnectionTest"
                        PerformConnectionTest()

                    Case "PerformInspection"
                        PerformInspection(args)
                    Case Else

                End Select

            Else
                'no arguments
                'PerformConnectionTest()
                'RunTestClass()
            End If


#Region "Test"


            'determine if all arguments are supplied before calling API
            'Dim argsLength = args.Length

            'for testing
            'bTransactionComplete = TestConnection.TestSAPConnection() 'SAP_Hierarchy.SapHierarchyTransactions(args) '

            'used for hierarchy only
            'If argsLength = 20 Then

            '    bTransactionComplete = TestConnection.TestSAPConnection() 'SAP_Hierarchy.SapHierarchyTransactions(args) '
            'Else
            '    'not enough arguments so return out of dll
            '    bTransactionComplete = False
            'End If

            ' Console.WriteLine("Was SAP Transaction Performed ? :  " + bTransactionComplete.ToString)
            'Console.ReadLine()

#End Region

        Catch ex As Exception
            Console.WriteLine(ex.StackTrace)
        End Try


    End Sub
    Public Function PerformConnectionTest()

        Dim bTransactionComplete = False

        Try
            bTransactionComplete = TestConnection.TestSAPConnection() 'SAP_Hierarchy.SapHierarchyTransactions(args) '

            Console.WriteLine("Was SAP Connection Test successful ? : " + bTransactionComplete.ToString)

        Catch ex As Exception

        End Try

        Return bTransactionComplete

    End Function

    Public Function PerformHierarchy(args As String())

        Dim index1 = 0
        Dim arrIndex = args.Length - 1

        Try
            'delete first argument
            args = args.Where(Function(s) s <> args(index1)).ToArray

            'call hierarchy api
            Dim returnVal = SAP_Hierarchy.SapHierarchyTransactions(args)

            Console.WriteLine("SAP Hierarchy transaction completed ? ~>  " + returnVal)

            Return returnVal

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return False
        End Try





    End Function

    Public Function PerformInspection(args As String())



        Try
            'delete first argument
            args = args.Where(Function(s) s <> args(0)).ToArray

            'grab args into params for method
            Dim lotNumber = args(0).ToString()
            Dim qty = args(1).ToString()

            'call Inspection function
            Dim returnVal As Boolean = Inspection.PerformSAPInspection(lotNumber, qty)

            Return returnVal

        Catch ex As Exception

            Return False

        End Try



    End Function

    Public Function RunTestClass()



        Try
            Dim test = ("Mitch", "Ragbir", "DoingTest", 20)

            Dim t = TestClass.TryTupleAsArg(test)

            Return True


        Catch ex As Exception
            Return False
        End Try



    End Function

    Public Function DoWork(Order_Number_Global As String, Material_Number_Global As String, KitSerial_Global As String,
                      SensorSerial_Global As String, CableSerial_Global As String,
                      SpareCableSerial_Global As String, SensorType_Global As String,
                      Storage_Location_Global As String, SubKitSer As String, SubSensorSer As String,
                      SubCableSer As String, SubSpareCableSer As String)

        Dim OrderNumber = Order_Number_Global 'Order_Number_Global
        Dim KitParentMaterial = Material_Number_Global 'Material_Number_Global
        Dim KitSerial = KitSerial_Global 'KitSerial_Global
        Dim SensorSerial = SensorSerial_Global 'SensorSerial_Global
        Dim CableSerial = CableSerial_Global 'CableSerial_Global
        Dim SpareCableSerial = SpareCableSerial_Global 'SpareCableSerial_Global
        Dim SensorType = SensorType_Global 'SensorType_Global
        Dim StorageLocation = Storage_Location_Global 'Storage_Location_Global


        Dim SubKitSerial = SubKitSer 'SubKitSer
        Dim SubSensorSerial = SubSensorSer 'SubSensorSer
        Dim SubCableSerial = SubCableSer 'SubCableSer
        Dim SubSpareCableSerial = SubSpareCableSer 'SubSpareCableSer




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
        session.findById("wnd[0]/usr/ctxtP_LGPRO").Text = "NY01" '1st storage location
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
        Return True
        'Exit Sub

errorHandler1:
        Return False
        ' Console.WriteLine("Make sure you are logged into SAP and have at least 3 windows open")
        'Exit Sub

errorHandler2:
        Return False
        ' Console.WriteLine("Error populating SAP text boxes , check SAP screen for error details")
        ' Exit Sub



    End Function

End Module
