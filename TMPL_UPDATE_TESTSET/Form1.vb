Imports System.Net
Imports System.Data.SqlClient

Public Class Form1
    Public clTDConnectivity As New TDConnectivity
    Sub GetTsTcFromTestLab()

        On Error GoTo ErrorHandler

        Dim tsTreeMgr, theTestSet
        Dim separateChars1 As String = "$"
        Dim separateChars2 As String = "<"
        Dim sLog As String, dDateToday As Date = Today

        Dim sTS_Name As String, sTS_Path As String, sTS_Status As String, sTS_ALM_ID As Long
        Dim sTS_Group As String, sTSPriorytet As String, sTSEmail As String, sTSEnvirnoment As String, sTSVDINum As String, sTSSystem As String
        Dim sTSLogLevel As String, sTSVDIName As String, sTSDataSource As String, sTS_Param As String, sTSReUseData As String
        Dim sTSRepeating As String, aRepeatingDate As Array, sRepeatingDate As String = ""

        Dim sTsTcCorrectDownload As String = "Not Completed"
        Dim sTsTcIncorrectTechnicalDownload As String = "Reject"
        Dim sTsTcUserInterrupt As String = "Canceled by user"
        Dim sTsTcProcessing As String = "Processing"


        Dim sTC_Name As String, sTC_Path As String, sTC_Param As String, sTC_Status As String, sTC_ALM_ID As Long
        Dim TSetFact, tsFolder, tsList, tsTestFactory, tsTestList
        Dim aTCList As Array, iTCNumber As Integer

        'Podłaczenie pod klasy łączące z DB MSSQL
        Dim dbConn As New dbConnectivity
        Dim sqlConnScenarioID As String, sqlConnCaseID As String

        'Podłaczenie pod klasy OTA zmieniające statusy TS / TC
        Dim clStatusChange As New StatusChange
        Dim clStatusChangeCase As New StatusChangeCase

        'Połącz z ALM przez OTA
        Call clTDConnectivity.ConnectToTD()

        'Jeśli mamy określoną ściężkę - określamy w TDConnetivity nPath = Trim("Root\Automaty") ' test set folder
        TSetFact = clTDConnectivity.tdConnection.TestSetFactory
        tsTreeMgr = clTDConnectivity.tdConnection.testsettreemanager
        tsFolder = tsTreeMgr.NodeByPath(clTDConnectivity.nPath)

        'Sprawdzenie czy istnieje lokalizacja
        If tsFolder Is Nothing Then
            Err.Number = 100
            GoTo ErrorHandler
        End If

        ' Sprawdzenie czy istnieje Test Set
        tsList = tsFolder.FindTestSets("")
        'Sprawdzenie czy istnieją Test Sety w określonej lokalizacji
        If tsList Is Nothing Then
            Err.Number = 101
            GoTo ErrorHandler
        End If

        'Lista Test Scenariuszy
        theTestSet = tsList.Item(1)

        For Each TestSetFound In tsList
            tsFolder = TestSetFound.TestSetFolder
            tsTestFactory = TestSetFound.tsTestFactory
            tsTestList = tsTestFactory.NewList("")
            'Debug.Print(StrTok(testsetfound.Field("CY_COMMENT"), separateChars1, separateChars2))

            sTS_Status = TestSetFound.Status
            If UCase(sTS_Status) = UCase("Ready to start") Then
                sTS_Name = TestSetFound.Name
                sTS_Path = tsFolder.path
                sTS_ALM_ID = TestSetFound.ID

                'Potwierdzenie rozpoczęcia procesowania TS
                sLog = "###POBRANIE DO KOLEJKI WYKONANIA - START###" & vbCrLf
                sLog = sLog & "Data: " & dDateToday & " " & Now.ToLongTimeString
                sLog = sLog & " | " & "ID_TS: " & sTS_ALM_ID
                sLog = sLog & " | " & "Status zmieniany z: " & sTS_Status & " na: " & sTsTcProcessing

                clStatusChange.ChangeStatusScenario(sTS_ALM_ID, sTsTcProcessing, sLog)

                'Grupa Testów	Cycle	CY_USER_06 'QTP UFT SELENIUM
                If IsNothing(TestSetFound.Field("CY_USER_06")) Then
                    Err.Number = 200
                    GoTo ErrorHandler
                Else
                    sTS_Group = "$group=" & TestSetFound.Field("CY_USER_06")
                End If

                'Priorytet Uruchomienia	Cycle	CY_USER_05 Wartości: 1;2;3;4;5  Domyślny: 3 Opis: (1-Najwyższy, 5-Najniższy)
                If IsNothing(TestSetFound.Field("CY_USER_05")) Then
                    sTSPriorytet = "$priorityId=3"
                Else
                    sTSPriorytet = "$priorityId=" & TestSetFound.Field("CY_USER_05")
                End If

                'Wysyłka email: Cycle   CY_USER_08
                If IsNothing(TestSetFound.Field("CY_USER_08")) Then
                    sTSEmail = "$sendEmail=N"
                Else
                    sTSEmail = "$sendEmail=" & TestSetFound.Field("CY_USER_08")
                End If

                'Środowisko  Cycle	CY_USER_09
                If IsNothing(TestSetFound.Field("CY_USER_09")) Then
                    Err.Number = 201
                    GoTo ErrorHandler
                Else
                    sTSEnvirnoment = "$envirnoment=" & TestSetFound.Field("CY_USER_09")
                End If

                'Log Level  Cycle	CY_USER_10
                If IsNothing(TestSetFound.Field("CY_USER_10")) Then
                    sTSLogLevel = "$logLevel=1"
                Else
                    sTSLogLevel = "$logLevel=" & TestSetFound.Field("CY_USER_10")
                End If

                'Wybór VDI	Cycle	CY_USER_01
                If IsNothing(TestSetFound.Field("CY_USER_01")) Then
                    sTSVDIName = "$machinesName=Wszystkie_VDI"
                    sTSVDINum = "$machines=999"
                Else
                    sTSVDIName = "$machinesName=" & TestSetFound.Field("CY_USER_01")
                    If LCase(sTSVDIName) = LCase("Wszystkie_VDI") Then
                        'Maksymalna ilość VDI jeśli nie jest wybrana określona maszyna	Cycle	CY_USER_02
                        sTSVDINum = "$machines=" & TestSetFound.Field("CY_USER_02")
                    Else
                        sTSVDINum = "$machines=1"
                    End If
                End If

                'Źródło danych	Cycle	CY_USER_04
                If IsNothing(TestSetFound.Field("CY_USER_10")) Then
                    sTSDataSource = "$sourceData=MSSQL"
                Else
                    sTSDataSource = "$sourceData=" & TestSetFound.Field("CY_USER_04")
                End If

                'Cykliczność Cycle	CY_USER_11
                sTSRepeating = TestSetFound.Field("CY_USER_11")

                If sTSRepeating = "Y" Then

                    clTDConnectivity.tdConnection.IgnoreHtmlFormat = True
                    aRepeatingDate = Split(TestSetFound.Field("CY_COMMENT"), Chr(10))
                    clTDConnectivity.tdConnection.IgnoreHtmlFormat = False

                    For i = LBound(aRepeatingDate) To UBound(aRepeatingDate)
                        If sRepeatingDate = "" And IsDate(Trim(aRepeatingDate(i))) Then
                            sRepeatingDate = "$runDate=" & Replace(Trim(aRepeatingDate(i)), vbCr, "")
                        ElseIf IsDate(Trim(aRepeatingDate(i))) Then
                            sRepeatingDate = sRepeatingDate & ",$runDate=" & Replace(Trim(aRepeatingDate(i)), vbCr, "")
                        ElseIf Trim(aRepeatingDate(i)) = "" Then

                        Else
                            Debug.Print("!" & aRepeatingDate(i) & "!")
                            'jeśli ktoś źle wprowadził daty. Zwrot na TS: Reject + dodanie logu
                            Err.Number = 202
                            GoTo ErrorHandler

                        End If
                    Next
                Else 'jeśli nie jest cykliczność sprawdzenie czy uruchomienie w dacie w przyszłości
                    If IsNothing(TestSetFound.Field("CY_USER_03")) Then
                        sRepeatingDate = "$runDate=" & Now.ToLongTimeString
                    Else 'Planowana Data uruchomienia jeśli nie cykliczna	CY_USER_03
                        sRepeatingDate = "$runDate=" & TestSetFound.Field("CY_USER_03")
                    End If
                End If

                'Reużywalność danych: Cycle   CY_USER_12
                If IsNothing(TestSetFound.Field("CY_USER_12")) Then
                    sTSReUseData = "$useCreatingData=N"
                Else
                    sTSReUseData = "$useCreatingData=" & TestSetFound.Field("CY_USER_12")
                End If

                'Testowany System  Cycle	CY_USER_12
                If IsNothing(TestSetFound.Field("CY_USER_12")) Then
                    sTSSystem = "$system=Brak_Danych"
                Else
                    sTSSystem = "$system=" & TestSetFound.Field("CY_USER_12")
                End If

                'Zebranie parametrów do przekazania do DB MSSQL do dalszej obróbki
                sTS_Param = sTS_Group
                sTS_Param = sTS_Param & "," & sTSPriorytet
                sTS_Param = sTS_Param & "," & sTSEmail
                sTS_Param = sTS_Param & "," & sTSEnvirnoment
                sTS_Param = sTS_Param & "," & sTSLogLevel
                sTS_Param = sTS_Param & "," & sTSVDIName
                sTS_Param = sTS_Param & "," & sTSVDINum
                sTS_Param = sTS_Param & "," & sTSDataSource
                sTS_Param = sTS_Param & "," & sTSReUseData
                sTS_Param = sTS_Param & "," & sTSSystem
                sTS_Param = sTS_Param & "," & sRepeatingDate

                'Dodaj TS do DB i pobierz unikalne ID z MSSQL
                sqlConnScenarioID = dbConn.sqlConnScenario(sTS_Name, sTS_Path, sTS_Param, sTS_Status, sTS_ALM_ID)

                'Rozpocznij wyszukiwanie TC
                For Each tsTest In tsTestList
                    sTC_Name = tsTest.Name
                    sTC_Path = tsFolder.path
                    clTDConnectivity.tdConnection.IgnoreHtmlFormat = True
                    If IsNothing(tsTest.field("TS_DESCRIPTION")) Or tsTest.field("TS_DESCRIPTION") = "" Then
                        sTC_Param = ("$repeat=1")
                    Else
                        sTC_Param = tsTest.field("TS_DESCRIPTION")
                        'sTC_Param = StrTok(tsTest.field("TS_DESCRIPTION"), separateChars1, separateChars2) '<-- jeśli format html
                    End If
                    clTDConnectivity.tdConnection.IgnoreHtmlFormat = False

                    sTC_Status = tsTest.Status
                    sTC_ALM_ID = tsTest.ID

                    sLog = "Data: " & dDateToday & " " & Now.ToLongTimeString
                    sLog = sLog & " | " & "ID_TS: " & sTS_ALM_ID
                    sLog = sLog & " | " & "ID_TC: " & sTC_ALM_ID
                    sLog = sLog & " | " & "Nazwa_TC: " & sTC_Name
                    sLog = sLog & " | " & "Status zmieniany z: " & sTC_Status & " na: " & sTsTcCorrectDownload

                    'Zapis TC do DB MSSQL
                    sqlConnCaseID = dbConn.sqlConnCase(sqlConnScenarioID, sTC_Name, sTC_Param, sTC_ALM_ID)
                    'Potwierdzenie wrzucenia TC 
                    clStatusChangeCase.ChangeStatusCase(sTS_ALM_ID, sTC_ALM_ID, sTsTcCorrectDownload, sLog)
                Next tsTest

                'Potwierdzenie zakańczania wrzucenia TS i TC do DB
                sqlConnCaseID = dbConn.sqlConnScenarioCompleted(sqlConnScenarioID)


                sLog = "Data: " & dDateToday & " " & Now.ToLongTimeString
                sLog = sLog & " | " & "ID_TS: " & sTS_ALM_ID
                sLog = sLog & " | " & "Status zmieniany z: " & sTsTcProcessing & " na: " & sTsTcCorrectDownload
                sLog = sLog & Chr(10) & "###POBRANIE DO KOLEJKI WYKONANIA - KONIEC###"
                'Potwierdzenie zakończenia pobierania całego TS
                clStatusChange.ChangeStatusScenario(sTS_ALM_ID, sTsTcCorrectDownload, sLog)
                Debug.Print(sqlConnScenarioID)
            End If
        Next TestSetFound

        'Zamknięcie połączenia z DB
        Call clTDConnectivity.DisconnectfromTD()

        Exit Sub

ErrorHandler:
        Dim bBoolKoniec As Boolean = True
        Select Case Err.Number
            Case 100
                Err.Description = ("Brak Folderu: " & clTDConnectivity.nPath)
                bBoolKoniec = True
            Case 101
                Err.Description = ("Brak TestSet w Folderze: " & clTDConnectivity.nPath)
                bBoolKoniec = True
            Case 200
                Err.Description = "Brak wybranej grupy testów (QTP/UFT/SELENIUM)"
                bBoolKoniec = False
            Case 201
                Err.Description = "Brak wybranego środowiska w Test Scenario"
                bBoolKoniec = False
            Case 202
                Err.Description = "Nieprawidłowa data cykliczności"
                bBoolKoniec = False
        End Select

        If bBoolKoniec = False Then
            sLog = "Data: " & dDateToday & " " & Now.ToLongTimeString
            sLog = sLog & " | " & "ID_TS: " & sTS_ALM_ID
            sLog = sLog & " | " & "Status zmieniany z: " & sTsTcProcessing & " na: " & sTsTcIncorrectTechnicalDownload
            sLog = sLog & Chr(10) & "Błąd: " & Err.Number & " | " & "Opis błędu: " & Err.Description
            sLog = sLog & Chr(10) & "###POBRANIE DO KOLEJKI WYKONANIA - KONIEC###"
            Err.Clear()
            'Potwierdzenie błędnego zakończenia pobierania TS do QC
            clStatusChange.ChangeStatusScenario(sTS_ALM_ID, sTsTcCorrectDownload, sLog)

            'TUTAJ JAKIŚ ZAPIS DO DB MSSQL
            'TUTAJ JAKIŚ ZAPIS DO DB MSSQL
            'TUTAJ JAKIŚ ZAPIS DO DB MSSQL
            'TUTAJ JAKIŚ ZAPIS DO DB MSSQL

            Call clTDConnectivity.DisconnectfromTD()
            'Uruchom ponownie pobieranie bez już pobranych TS/TC
            GetTsTcFromTestLab()
        Else
            sLog = "Data: " & dDateToday & " " & Now.ToLongTimeString
            sLog = sLog & " | " & "ID_TS: " & sTS_ALM_ID
            sLog = sLog & " | " & "Status zmieniany z: " & sTsTcProcessing & " na: " & sTsTcIncorrectTechnicalDownload
            sLog = sLog & Chr(10) & "Błąd: " & Err.Number & " | " & "Opis błędu: " & Err.Description
            sLog = sLog & Chr(10) & "###POBRANIE DO KOLEJKI WYKONANIA - KONIEC###"
            Err.Clear()
            'Potwierdzenie błędnego zakończenia pobierania TS do QC
            clStatusChange.ChangeStatusScenario(sTS_ALM_ID, sTsTcCorrectDownload, sLog)

            'TUTAJ JAKIŚ ZAPIS DO DB MSSQL
            'TUTAJ JAKIŚ ZAPIS DO DB MSSQL
            'TUTAJ JAKIŚ ZAPIS DO DB MSSQL
            'TUTAJ JAKIŚ ZAPIS DO DB MSSQL

            Call clTDConnectivity.DisconnectfromTD()
            Exit Sub
        End If

    End Sub

    Function StrTok(ByVal str As String,
                ByVal separateChars1 As String, ByVal separateChars2 As String) As String
        Dim i As Long
        Dim pos As Long, posParam As Long
        Dim iDlugoscParam As Integer
        Dim lDlugoscSTR As Long
        StrTok = ""

        For i = 1 To Len(str)
            pos = InStr(1, str, Strings.Mid(separateChars1, 1, 1), vbTextCompare)
            If pos > 0 Then
                'MsgBox(str)
                posParam = InStr(pos, str, Strings.Mid(separateChars2, 1, 1), vbTextCompare)
                If Len(StrTok) = 0 Then
                    StrTok = Strings.Mid(str, pos, posParam - pos)
                Else
                    StrTok = StrTok & "," & Strings.Mid(str, pos, posParam - pos)
                End If
                StrTok = Replace(StrTok, "&quot;", "")
                StrTok = Replace(StrTok, "&nbsp;", "")
                StrTok = Replace(StrTok, """", "")

                lDlugoscSTR = Len(str)
                iDlugoscParam = lDlugoscSTR - posParam
                str = Strings.Right(str, iDlugoscParam)
            Else
                Exit For
            End If
        Next i

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GetTsTcFromTestLab()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim clStatusChangeCase As New StatusChangeCase
        'Call clStatusChangeCase.ChangeStatusCase("101", "5", "Failed")
    End Sub
End Class

Public Class TDConnectivity
    Public tdConnection As Object
    Public nPath As String

    Public Sub ConnectToTD()
        Dim qcURL As String
        Dim qcID As String
        Dim qcPWD As String
        Dim qcDomain As String
        Dim qcProject As String

        qcURL = "http://profive1v5.radom.tsunami:8080/qcbin"
        qcID = "mantoniak"
        qcPWD = "mantoniak"
        qcDomain = "UFT"
        qcProject = "UFT"
        nPath = Trim("Root\Automaty") ' test set folder
        '-----------------------------------------------------Connect to Quality Center --------------------------------------------------------

        'Create a Connection object to connect to Quality Center
        tdConnection = CreateObject("TDApiOle80.TDConnection")
        'Initialise the Quality center connection
        tdConnection.InitConnectionEx(qcURL)
        'Authenticating with username and password
        tdConnection.Login(qcID, qcPWD)
        'connecting to the domain and project
        tdConnection.Connect(qcDomain, qcProject)


    End Sub

    Public Sub DisconnectfromTD()
        '------------------------------------------------------Disconnect Quality Center -----------------------------------------------------------------
        On Error Resume Next
        tdConnection.Disconnect
        tdConnection.Logout
        tdConnection.ReleaseConnection
        On Error GoTo 0
    End Sub
End Class

Public Class dbConnectivity
    Public Function sqlConnScenario(ByVal sName As String, ByVal sPath As String,
                       ByVal sParam As String, ByVal sStatus As String, ByVal sALM_ID As Long) As String
        Dim connectionString As String
        Dim connection As SqlConnection
        Dim command As SqlCommand
        'Dim sql As String
        Dim sIDMssql As String = ""
        ' ***************************
        ' string używany do uruchomienia SP (Store Procedure) na bazie
        '**************************
        connectionString = "Data Source=profive2.radom.tsunami;Initial Catalog=auto;User ID=autoUser;Password=xyz1234#"
        connection = New SqlConnection(connectionString)

        Try
            connection.Open()
            command = New SqlCommand()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "conf.addScenario"
            command.Parameters.AddWithValue("scenarioName", sName)
            command.Parameters.AddWithValue("scenarioHPId", sALM_ID)
            command.Parameters.AddWithValue("path", sPath)
            command.Parameters.AddWithValue("statusHP", sStatus)
            command.Parameters.AddWithValue("param", sParam)
            Dim sqlReader As SqlDataReader = command.ExecuteReader()
            While sqlReader.Read()
                sIDMssql = (sqlReader.Item(0).ToString())
            End While
            sqlReader.Close()
            connection.Close()

        Catch ex As Exception
            MsgBox("Can not open connection ! " & ex.Message.ToString())
        End Try
        sqlConnScenario = sIDMssql
    End Function

    Public Function sqlConnCase(ByVal sTestScenarioID As String, ByVal sName As String, ByVal sParam As String,
                                ByVal sALM_ID As Long) As String
        Dim connectionString As String
        Dim connection As SqlConnection
        Dim command As SqlCommand
        Dim sIDMssql As String = ""
        Dim sParameterToDB As String
        ' ***************************
        ' string używany do odpalenia SP na bazie
        '**************************
        connectionString = "Data Source=profive2.radom.tsunami;Initial Catalog=auto;User ID=autoUser;Password=xyz1234#"
        connection = New SqlConnection(connectionString)

        sParameterToDB = ("$testCaseName=" & sName & "," & "$scenarioCaseId=" & sALM_ID & "," & sParam)
        Try
            connection.Open()
            command = New SqlCommand()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "conf.addScenarioCase"
            command.Parameters.AddWithValue("testScenarioId", sTestScenarioID)
            command.Parameters.AddWithValue("param", sParameterToDB)
            'command.Parameters.AddWithValue("testCaseName", sName)
            '  command.Parameters.AddWithValue("testCaseName", sName)
            ' command.Parameters.AddWithValue("repeat", sParam)
            'command.Parameters.AddWithValue("scenarioCaseId", sALM_ID)

            Dim sqlReader As SqlDataReader = command.ExecuteReader()
            While sqlReader.Read()
                sIDMssql = (sqlReader.Item(0).ToString())
            End While
            sqlReader.Close()
            connection.Close()

        Catch ex As Exception
            MsgBox("Can not open connection ! " & ex.Message.ToString())
        End Try
        sqlConnCase = sIDMssql
    End Function

    Public Function sqlConnScenarioCompleted(ByVal sTestScenarioID As String) As String
        Dim connectionString As String
        Dim connection As SqlConnection
        Dim command As SqlCommand
        Dim sIDMssql As String = ""
        ' ***************************
        ' string używany do odpalenia SP na bazie
        '**************************
        connectionString = "Data Source=profive2.radom.tsunami;Initial Catalog=auto;User ID=autoUser;Password=xyz1234#"
        connection = New SqlConnection(connectionString)


        Try
            connection.Open()
            command = New SqlCommand()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "conf.addScenarioCompleted"
            command.Parameters.AddWithValue("testScenarioId", sTestScenarioID)

            Dim sqlReader As SqlDataReader = command.ExecuteReader()
            While sqlReader.Read()
                sIDMssql = (sqlReader.Item(0).ToString())
            End While
            sqlReader.Close()
            connection.Close()

        Catch ex As Exception
            MsgBox("Can not open connection ! " & ex.Message.ToString())
        End Try
        sqlConnScenarioCompleted = sIDMssql
    End Function



End Class


Public Class StatusChange
    Public clTDConnectivity As New TDConnectivity

    Sub ChangeStatusScenario(ByVal qcTSId As Long, ByVal qcStatus As String, ByVal sLog As String)
        Call clTDConnectivity.ConnectToTD()

        Dim TSetFact, TestSetFilter, TestSetList, myTestSet
        TSetFact = clTDConnectivity.tdConnection.TestSetFactory
        TestSetFilter = TSetFact.Filter
        TestSetFilter.Filter("CY_CYCLE_ID") = qcTSId

        TestSetList = TestSetFilter.NewList
        myTestSet = TestSetList.Item(1)

        clTDConnectivity.tdConnection.IgnoreHtmlFormat = False

        myTestSet.Field("CY_STATUS") = (qcStatus)
        myTestSet.Field("CY_USER_07") = myTestSet.Field("CY_USER_07") & vbCr & sLog
        myTestSet.Post

        'clTDConnectivity.tdConnection.IgnoreHtmlFormat = True

        Call clTDConnectivity.DisconnectfromTD()
        TestSetFilter = Nothing
        myTestSet = Nothing
        TSetFact = Nothing
    End Sub
End Class

Public Class StatusChangeCase
    Public clTDConnectivity As New TDConnectivity

    Sub ChangeStatusCase(ByVal qcTSId As Long, ByVal qcTCId As Long, ByVal qcStatus As String, ByVal sLog As String)
        Call clTDConnectivity.ConnectToTD()

        Dim TSetFact, TestSetFilter, TestSetList, myTestSet, TSTestFactory
        Dim tcStepID, theRun

        TSetFact = clTDConnectivity.tdConnection.TestSetFactory
        TestSetFilter = TSetFact.Filter
        TestSetFilter.Filter("CY_CYCLE_ID") = qcTSId

        TestSetList = TestSetFilter.NewList
        myTestSet = TestSetList.Item(1)
        TSTestFactory = myTestSet.TSTestFactory

        clTDConnectivity.tdConnection.IgnoreHtmlFormat = False

        myTestSet.Field("CY_USER_07") = myTestSet.Field("CY_USER_07") & vbCr & sLog
        myTestSet.post

        'clTDConnectivity.tdConnection.IgnoreHtmlFormat = False

        For Each qtTest In TSTestFactory.NewList("")
            If qtTest.id = qcTCId Then
                qtTest.Field("TC_STATUS") = qcStatus
                qtTest.Post
                tcStepID = qtTest.RunFactory
                For Each theRun In tcStepID.newlist("")
                    theRun.Status = qcStatus
                    ' theRun.CopyDesignSteps
                    theRun.Post
                Next
            End If
        Next
        Call clTDConnectivity.DisconnectfromTD()

    End Sub

End Class


