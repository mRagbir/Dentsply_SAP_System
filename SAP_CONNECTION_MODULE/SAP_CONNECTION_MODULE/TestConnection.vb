Public Class TestConnection

    Public Shared Function TestSAPConnection()

        Dim bIsConnected As Boolean = False
        Dim App, Connection, session As Object
        Dim SapGuiAuto As Object

        Try

            SapGuiAuto = GetObject("SAPGUI")

            ' Console.WriteLine("Found SAP object : SAPGUI")

            App = SapGuiAuto.GetScriptingEngine

            Try
                Try
                    Connection = App.Children(0)

                    session = Connection.Children(0) ' grabs the 1st sap session window.

                    'Console.WriteLine("Found SAP session")

                    '~~~> t-code
                    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmmbe" 'navigate to the hierarchy screen
                    session.findById("wnd[0]").sendVKey(0) ' enter

                    'test input of a part number
                    session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").text = "B1218000"
                    session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").caretPosition = 9
                    session.findById("wnd[0]/tbar[1]/btn[8]").press

                    bIsConnected = True

                Catch
                    Console.WriteLine("Error: You are NOT LOGGED into SAP!..PLease log into system , open 2 sessions then try again")
                End Try


            Catch ex As Exception
                Console.WriteLine("Error : " + ex.InnerException.ToString())
                bIsConnected = False
            End Try

            Return bIsConnected

        Catch ex As Exception
            Console.WriteLine("Error : " + ex.InnerException.ToString())
            session = Nothing
            Connection = Nothing
            App = Nothing
            SapGuiAuto = Nothing
            ' errorFound = ex.InnerException.ToString()
            Return bIsConnected

        Finally

            'clean up objects in reverse order
            session = Nothing
            Connection = Nothing
            App = Nothing
            SapGuiAuto = Nothing

        End Try

    End Function

End Class
