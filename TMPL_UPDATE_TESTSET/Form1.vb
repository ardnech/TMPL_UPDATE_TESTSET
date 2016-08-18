Imports System.Net
Imports System.Data.SqlClient
Imports TDAPIOLELib

Public Class Form1
    Public clTDConnectivity As New TDConnectivity
    Sub GetTsTcFromTestLab()

        '  On Error GoTo ErrorHandler

        Dim tsTreeMgr, theTestSet
        Dim separateChars1 As String = "$"
        Dim separateChars2 As String = "<"
        Dim sLog As String, dDateToday As Date = Today

        Dim sTS_Name As String, sTS_Path As String, sTS_Status As String, sTS_ALM_ID As Long
        Dim sTS_Group As String, sTSPriorytet As String, sTSEmail As String, sTSEnvirnoment As String, sTSVDINum As String, sTSSystem As String
        Dim sTSLogLevel As String, sTSVDIName As String, sTSDataSource As String, sTS_Param As String, sTSReUseData As String
        Dim sTSRepeating As String, aRepeatingDate As Array, sRepeatingDate As String = "", sTSJira As String

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
            If UCase(sTS_Status) = UCase("Ready to start") And TestSetFound.IsLocked = False Then
                TestSetFound.LockObject
                sTS_Name = TestSetFound.Name
                sTS_Path = tsFolder.path
                sTS_ALM_ID = TestSetFound.ID

                'Potwierdzenie rozpoczęcia procesowania TS
                sLog = "###POBRANIE DO KOLEJKI WYKONANIA - START###" & vbCrLf
                sLog = sLog & "Data: " & dDateToday & " " & Now.ToLongTimeString
                sLog = sLog & " | " & "ID_TS: " & sTS_ALM_ID
                sLog = sLog & " | " & "Status zmieniany z: " & sTS_Status & " na: " & sTsTcProcessing

                clStatusChange.ChangeStatusScenario(sTS_ALM_ID, sTsTcProcessing, sLog)
                TestSetFound.UnlockObject
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
                    sTSRepeating = "N" 'przypisanie wartości jeśli nie jest Y
                    If IsNothing(TestSetFound.Field("CY_USER_03")) Then
                        sRepeatingDate = "$runDate=" & Now.ToLongTimeString
                    Else 'Planowana Data uruchomienia jeśli nie cykliczna	CY_USER_03
                        sRepeatingDate = "$runDate=" & TestSetFound.Field("CY_USER_03")
                    End If
                End If

                'Reużywalność danych: Cycle   CY_USER_14
                If IsNothing(TestSetFound.Field("CY_USER_14")) Then
                    sTSReUseData = "$useCreatingData=N"
                Else
                    sTSReUseData = "$useCreatingData=" & TestSetFound.Field("CY_USER_14")
                End If

                'Testowany System  Cycle	CY_USER_12
                If IsNothing(TestSetFound.Field("CY_USER_12")) Then
                    sTSSystem = "$system=Brak_Danych"
                Else
                    sTSSystem = "$system=" & TestSetFound.Field("CY_USER_12")
                End If
                'Wysyłka Jira CY_USER_15
                If IsNothing(TestSetFound.Field("CY_USER_15")) Then
                    sTSJira = "$sendJira=N"
                Else
                    sTSJira = "$sendJira=" & TestSetFound.Field("CY_USER_15")
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
                sTS_Param = sTS_Param & "," & sTSRepeating
                sTS_Param = sTS_Param & "," & sTSJira

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
        MsgBox("Done")
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
        Call clStatusChangeCase.ChangeStatusCase("101", "5", "Failed", "AA")
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Dim clAddAttachment As New AddAttachments
        Call clAddAttachment.AddAttachmentTS_TC("108", "1002", "Jakiś opis do dodawanego załącznika", "C:\\test.txt", 0)
        'TS ID, TC ID, Opis załącznika, Ścieżka załącznika, parametr: 0 - dodawanie do TS lub 1 - dodawanie do TS i TC)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ClCopyTP As New Copy_TPtoTL
        Call ClCopyTP.CopyTPlanToTLab()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'Dim tdConnection As Object
        On Error Resume Next
        Dim nPath As String
        Dim qcURL As String
        Dim qcID As String
        Dim qcPWD As String
        Dim qcDomain As String
        Dim qcProject As String
        Dim TestSetFound, strMsg, iFolderID

        Dim tdConnection = New TDAPIOLELib.TDConnection
        Dim TSetFact As TestSetFactory
        Dim tsTreeMgr As TestSetTreeManager
        Dim tsFolder As TestSetFolder
        Dim tsList
        iFolderID = 65

        qcURL = "http://profive1v5.radom.tsunami:8080/qcbin"
        qcID = "uft"
        qcPWD = "uft"
        qcDomain = "UFT"
        qcProject = "UFT"
        nPath = Trim("Root") ' test set folder
        '-----------------------------------------------------Connect to Quality Center --------------------------------------------------------
        'tdConnection
        'Create a Connection object to connect to Quality Center
        'tdConnection = CreateObject("TDApiOle80.TDConnection")


        'Initialise the Quality center connection
        TDConnection.InitConnectionEx(qcURL)
        'Authenticating with username and password
        tdConnection.Login(qcID, qcPWD)
        'connecting to the domain and project
        tdConnection.Connect(qcDomain, qcProject)

        tsTreeMgr = tdConnection.TestSetTreeManager
        tsFolder = tsTreeMgr.NodeById(iFolderID)
        'MsgBox(tsFolder.Path)
        tsList = tsFolder.FindTestSets("")
        For Each TestSetFound In tsList

            TestSetFound.Status = "Ready to start"
            TestSetFound.Post
            TestSetFound.Refresh
        Next

        '        / where 'tdc' is a valid TDConnection object logged in to DEFAULT.QualityCenter_Demo
        '// string testSetFolderPath = @"Root\Mercury Tours Web Site";                          // 0 test sets (Method 1 And 2 return 5)
        'String testSetFolderPath = @"Root\Mercury Tours Web Site\Functionality And UI";     // 3 test sets

        '// Method 1 TestSetFolder.FindTestSets()
        '        var TestSetTreeManager = (TestSetTreeManager)tdc.TestSetTreeManager;
        'var TestSetFolder = (TestSetFolder)testSetTreeManager.get_NodeByPath(testSetFolderPath);
        'var testSets = TestSetFolder.FindTestSets("", False, "");
        'Console.WriteLine("Folder {0} contains {1} testsets", testSetFolderPath, testSets.Count);

        '// Method 2 NewList() With filter
        'var TestSetFactory = (TestSetFactory)tdc.TestSetFactory;
        'var Filter() = (TDFilter)testSetFactory.Filter;
        'Filter()["CY_FOLDER_ID"] = "^" + testSetFolderPath + "^";
        'testSets = (List)testSetFactory.NewList(filter.Text);
        'Console.WriteLine("Folder {0} contains {1} testsets", testSetFolderPath, testSets.Count);

        '// Method 3 Sql Query using Command object
        'var Command = tdc.Command;
        'Command.CommandText = "select CY_CYCLE as TestSet from CYCLE where CY_FOLDER_ID = " + TestSetFolder.NodeID;
        'Recordset records = Command.Execute();
        'Console.WriteLine("Folder {0} contains {1} testsets", testSetFolderPath, records.RecordCount);


        'are you sure you want to copy?'
        strMsg = "Wykryte TestScenariusze: " & tsList.Count & Chr(10)
        strMsg = strMsg & "w katalogu: " & tsFolder.Path & Chr(10)
        strMsg = strMsg & "Czy zmienić statun na 'Ready to start'?"
        'Result = MsgBox(strMsg, vbYesNo, "Zmiana statusu na 'Ready to start'")

        ' If Result = vbNo Then
        ' Exit Function
        ' End If



        '-------------------------------------------------?-----
        'Cleanup for objects (just to be sure)
        TSetFact = Nothing
        tsTreeMgr = Nothing
        tsFolder = Nothing
        tsList = Nothing
        TestSetFound = Nothing


        tdConnection.Disconnect()
        tdConnection.Logout()
        tdConnection.ReleaseConnection()
        On Error GoTo 0



    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        On Error Resume Next
        Dim ClassCustomizationList As New CustomizationList
        Call ClassCustomizationList.GetCustomizationList("PR_Envirnoment")
        '        '#Lista List do wywołań które można modyfikować
        '        PR_Envirnoment
        '        PR_ListaVDI
        '        PR_TestGroup
        '        PR_System

        '        '#Tych nie implementujemy Do zmian, ale są
        '        Plan Status
        '        PR_DataSource
        '        PR_LogLevel
        '        PR_RunPriority
        '        Test Running Status
        On Error GoTo 0
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim ClassCustomizationList As New CustomizationList
        On Error Resume Next
        Call ClassCustomizationList.AddItemToList("PR_ListaVDI", "VDI-10") '1 wartość  z listy poniżej, 2 wartość Input Box
        On Error GoTo 0
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim PeselGenerator As New GeneratoryDanych
        On Error Resume Next
        MsgBox(PeselGenerator.GetPesel())
        On Error GoTo 0
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim NipGenerator As New GeneratoryDanych
        On Error Resume Next
        MsgBox(NipGenerator.GetNIP())
        On Error GoTo 0
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim RegonGenerator As New GeneratoryDanych
        ' On Error Resume Next
        MsgBox(RegonGenerator.GetREGON())
        On Error GoTo 0
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim DowodGenerator As New GeneratoryDanych
        ' On Error Resume Next
        MsgBox(DowodGenerator.GetDowod())
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
        qcID = "uft"
        qcPWD = "uft"
        qcDomain = "UFT"
        qcProject = "UFT"
        nPath = Trim("Root\NoweTesty") ' test set folder
        '-----------------------------------------------------Connect to Quality Center --------------------------------------------------------
        'tdConnection
        'Create a Connection object to connect to Quality Center
        tdConnection = CreateObject("TDApiOle80.TDConnection")
        'tdConnection = New TDAPIOLELib.TDConnection
        ' tdConnection.
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

    Sub ChangeStatusScenario(ByVal qcTSId As Long, ByVal qcStatus As String, ByVal sLog As String,
                             Optional ByVal sAttachment As String = "C:\\test.txt")
        'Call clTDConnectivity.ConnectToTD()
        Dim dDateToday As String = Today & "_" & Now.ToLongTimeString
        Dim TSetFact, TestSetFilter, TestSetList, myTestSet
        Dim attachF, theAttachment
        TSetFact = clTDConnectivity.tdConnection.TestSetFactory
        TestSetFilter = TSetFact.Filter
        TestSetFilter.Filter("CY_CYCLE_ID") = qcTSId

        TestSetList = TestSetFilter.NewList
        myTestSet = TestSetList.Item(1)

        clTDConnectivity.tdConnection.IgnoreHtmlFormat = False

        myTestSet.Field("CY_STATUS") = (qcStatus)
        myTestSet.Field("CY_USER_07") = myTestSet.Field("CY_USER_07") & vbCr & sLog
        myTestSet.Post


        If sAttachment <> "" Then
            attachF = myTestSet.Attachments
            theAttachment = attachF.AddItem(System.DBNull.Value)
            theAttachment.Description = dDateToday & " Załącznik dodany: " & sAttachment
            theAttachment.Type = 1
            theAttachment.FileName = sAttachment

            theAttachment.Post
        End If
        ' theAttachment.Filename = "123"
        ' theAttachment.Type = "TDATT_FILE"

        'End If
        'clTDConnectivity.tdConnection.IgnoreHtmlFormat = True

        Call clTDConnectivity.DisconnectfromTD()
        TestSetFilter = Nothing
        myTestSet = Nothing
        TSetFact = Nothing
    End Sub
End Class

Public Class StatusChangeCase
    Public clTDConnectivity As New TDConnectivity

    Sub ChangeStatusCase(ByVal qcTSId As Long, ByVal qcTCId As Long, ByVal qcStatus As String,
                         ByVal sLog As String, Optional ByVal sAttachment As String = "C:\\test.txt")
        Call clTDConnectivity.ConnectToTD()

        Dim TSetFact, TestSetFilter, TestSetList, myTestSet, TSTestFactory
        Dim tcStepID, theRun, attachF, theAttachment
        Dim dDateToday As String = Today & "_" & Now.ToLongTimeString


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
                'qtTest.Field("TC_STATUS") = "No Run"
                qtTest.Post
                tcStepID = qtTest.RunFactory
                ' RunF = tstInstance.RunFactory 
                'For Each theRun In tcStepID.newlist("")
                theRun = tcStepID.AddItem("Automated_" & dDateToday)
                theRun.Status = qcStatus
                ' theRun.CopyDesignSteps
                theRun.Post

                If sAttachment <> "" Then
                    attachF = theRun.Attachments
                    theAttachment = attachF.AddItem(System.DBNull.Value)
                    theAttachment.Description = dDateToday & " Załącznik dodany: " & sAttachment
                    theAttachment.Type = 1
                    theAttachment.FileName = sAttachment
                    theAttachment.Post
                End If
                theRun.Refresh
                ' Next
            End If
        Next
        Call clTDConnectivity.DisconnectfromTD()

    End Sub



End Class

Public Class AddAttachments
    Public clTDConnectivity As New TDConnectivity
    Sub AddAttachmentTS_TC(ByVal qcTSId As Long, ByVal qcTCId As Long, ByVal sLog As String, ByVal sAttachment As String, Optional ByVal iAttachment As Integer = 1)
        Call clTDConnectivity.ConnectToTD()

        Dim TSetFact, TestSetFilter, TestSetList, myTestSet, TSTestFactory
        Dim tcStepID, theRun, attachF, theAttachment
        Dim dDateToday As String = Today & "_" & Now.ToLongTimeString


        TSetFact = clTDConnectivity.tdConnection.TestSetFactory
        TestSetFilter = TSetFact.Filter
        TestSetFilter.Filter("CY_CYCLE_ID") = qcTSId

        TestSetList = TestSetFilter.NewList
        myTestSet = TestSetList.Item(1)
        TSTestFactory = myTestSet.TSTestFactory

        clTDConnectivity.tdConnection.IgnoreHtmlFormat = False

        Select Case iAttachment
            Case 0 'Dodanie załącznika do TS i TC
                myTestSet.Field("CY_USER_07") = myTestSet.Field("CY_USER_07") & vbCr & sLog
                myTestSet.post
                attachF = myTestSet.Attachments
                theAttachment = attachF.AddItem(System.DBNull.Value)
                theAttachment.Description = dDateToday & " Załącznik dodany: " & sAttachment
                theAttachment.Type = 1
                theAttachment.FileName = sAttachment
                theAttachment.Post

                For Each qtTest In TSTestFactory.NewList("")
                    If qtTest.id = qcTCId Then
                        qtTest.Post
                        tcStepID = qtTest.RunFactory
                        theRun = tcStepID.AddItem("Automated_" & dDateToday)
                        attachF = theRun.Attachments
                        theAttachment = attachF.AddItem(System.DBNull.Value)
                        theAttachment.Description = dDateToday & " Załącznik dodany: " & sAttachment
                        theAttachment.Type = 1
                        theAttachment.FileName = sAttachment
                        theAttachment.Post
                        theRun.Refresh
                        ' Next
                    End If
                Next
            Case 1 'Dodanie załącznika do TS 
                myTestSet.Field("CY_USER_07") = myTestSet.Field("CY_USER_07") & vbCr & sLog
                myTestSet.post
                attachF = myTestSet.Attachments
                theAttachment = attachF.AddItem(System.DBNull.Value)
                theAttachment.Description = dDateToday & " Załącznik dodany: " & sAttachment
                theAttachment.Type = 1
                theAttachment.FileName = sAttachment
                theAttachment.Post
        End Select
        Call clTDConnectivity.DisconnectfromTD()

    End Sub
End Class

Public Class Copy_TPtoTL

    Public clTDConnectivity As New TDConnectivity
    Sub CopyTPlanToTLab()
        Call clTDConnectivity.ConnectToTD()
        Dim Treemgr, myTestFact, myTestFilter, TestFactory, tdc, treemanager
        Dim myTestList, strMsg, myFolderPath, myFolderID
        Dim Result

        'Dim treemanager As TDAPIOLELib.TreeManager


        myFolderID = "1001"
        myFolderPath = "Subject\test"
        tdc = clTDConnectivity.tdConnection

        Treemgr = tdc.treemanager
        myTestFact = tdc.TestFactory
        myTestFilter = myTestFact.Filter

        ' build filter regarding the last known folder'
        myTestFilter.Filter("TS_SUBJECT") = "^\" & myFolderPath & "^"
        myTestFilter.Order("TS_SUBJECT") = 1
        myTestFilter.Order("TS_NAME") = 2
        myTestList = myTestFact.NewList(myTestFilter.Text)

        'are you sure you want to copy?'
        strMsg = "Wykryte TestCase: " & myTestList.Count & Chr(10)
        strMsg = strMsg & "w katalogu: " & myFolderPath & Chr(10)
        strMsg = strMsg & "Czy skopiować do TestLab?"
        Result = MsgBox(strMsg, vbYesNo, "Kopiowanie TestPlan do TestLab")

        If Result = vbNo Then
            Exit Sub
        End If
        Dim actTest
        Dim mySNode, myPath, sTSVDIName, sTSVDINum, sStartDate, sTSDataSource, sTSPriorytet, sTestGroup
        Dim sTSEmail, sTSEnvirnoment, sTSLogLevel, sPeriodic, sTSReUseData, sRepeatingDate, sTSSystem, sName
        For Each actTest In myTestList
            'Node of Subject-Folder
            mySNode = actTest.Field("TS_Subject")
            myPath = mySNode.Path
            sTSVDIName = actTest.Field("TS_USER_01")           'Nazwa maszyny
            sTSVDINum = actTest.Field("TS_USER_02")      'Maksymalna ilość maszyn
            sStartDate = actTest.Field("TS_USER_03")              'Data uruchomienia
            sTSDataSource = actTest.Field("TS_USER_04")           'Wybór źródła danych
            sTSPriorytet = actTest.Field("TS_USER_05")               'Priorytet uruchomienia
            sTestGroup = actTest.Field("TS_USER_06")              'Grupa Testów
            sTSEmail = actTest.Field("TS_USER_07")              'Wyślij mail
            sTSEnvirnoment = actTest.Field("TS_USER_08")              'Środowisko
            sTSLogLevel = actTest.Field("TS_USER_09")                'Poziom logowania
            sPeriodic = actTest.Field("TS_USER_10")                'Cykliczność
            sTSReUseData = actTest.Field("TS_USER_11")              'Reużyj dane
            sRepeatingDate = actTest.Field("TS_Description")      'DatyCykliczności
            sTSSystem = actTest.Field("TS_USER_12")               'Testowany System
            sName = actTest.Field("TS_NAME")

            strMsg = "sTSVDIName: " & sTSVDIName & Chr(10)
            strMsg = strMsg & "sTSVDINum: " & sTSVDINum & Chr(10)
            strMsg = strMsg & "sStartDate: " & sStartDate & Chr(10)
            strMsg = strMsg & "sTSDataSource: " & sTSDataSource & Chr(10)
            strMsg = strMsg & "sTSPriorytet: " & sTSPriorytet & Chr(10)
            strMsg = strMsg & "sTestGroup: " & sTestGroup & Chr(10)
            strMsg = strMsg & "sTSEmail: " & sTSEmail & Chr(10)
            strMsg = strMsg & "sTSEnvirnoment: " & sTSEnvirnoment & Chr(10)
            strMsg = strMsg & "sTSLogLevel: " & sTSLogLevel & Chr(10)
            strMsg = strMsg & "sPeriodic: " & sPeriodic & Chr(10)
            strMsg = strMsg & "sTSReUseData: " & sTSReUseData & Chr(10)
            strMsg = strMsg & "sRepeatingDate: " & sRepeatingDate & Chr(10)
            strMsg = strMsg & "sTSSystem: " & sTSSystem & Chr(10)
            strMsg = strMsg & "myPath: " & myPath & Chr(10)
            'strMSG = strMSG & "mySNode: " & mySNode & Chr(10)
            strMsg = strMsg & "myName: " & sName

            MsgBox(strMsg)

            'build testset and add testinstance
            Result = StworzTestCase(myPath, actTest, sTSVDIName, sTSVDINum, sStartDate, sTSDataSource, sTSPriorytet, sTestGroup, sTSEmail,
            sTSEnvirnoment, sTSLogLevel, sPeriodic, sTSReUseData, sRepeatingDate, sTSSystem)
        Next 'Testcase

        'now the end is near
        MsgBox("Kopiowanie zakończone", vbOKOnly)
        Call clTDConnectivity.DisconnectfromTD()
        myTestList = Nothing
        myTestFilter = Nothing
        myTestFact = Nothing
    End Sub


    Function StworzTestCase(sPath, sTestCase, sTSVDIName, sTSVDINum, sStartDate, sTSDataSource, sTSPriorytet, sTestGroup, sTSEmail,
    sTSEnvirnoment, sTSLogLevel, sPeriodic, sTSReUseData, sRepeatingDate, sTSSystem)

        Dim tdcF
        Dim TStmgr
        Dim myRoot
        Dim newTSTest
        Dim myTSTest, testSetFilter, TSTestF, TSTestList, testSetF, folder, newNode, build_case
        Dim subjectArray, NewPath, OldPath, CurrentSubName
        Dim TSList, testSet1, foundTS, currentPath

        Dim iDateFolder = 1 'jeśli 1 to kreowanie folderów z NOW
        'On Error Resume Next
        ' Preparation
        tdcF = clTDConnectivity.tdConnection
        TStmgr = tdcF.TestSetTreeManager
        ' Split path for loop
        subjectArray = Split(sPath, "\")

        ' initialize variable for path
        ' Remember: Test Plan begins with Subject and Test Lab with Root!
        NewPath = "Root"
        OldPath = ""

        If iDateFolder = 1 Then
            Dim NewFolder = Now()
            NewPath = Trim(NewPath) & "\" & NewFolder
            On Error Resume Next
            'search Folder
            newNode = TStmgr.NodeByPath(NewPath)
            On Error GoTo 0
            If newNode Is Nothing Then
                TStmgr = Nothing
                TStmgr = tdcF.TestSetTreeManager
                myRoot = TStmgr.Root
                newNode = myRoot.addNode(NewFolder)
                newNode.post
                newNode.Refresh

            End If 'new Node
        End If
        For idx = 1 To UBound(subjectArray)
            'save path
            OldPath = NewPath

            'get new folder
            CurrentSubName = subjectArray(idx)
            'build new path
            NewPath = Trim(NewPath) & "\" & CurrentSubName
            On Error Resume Next
            newNode = Nothing
            'search Folder
            newNode = TStmgr.NodeByPath(NewPath)
            On Error GoTo 0
            'create folder if it does not exist

            If newNode Is Nothing Then
                TStmgr = Nothing
                TStmgr = tdcF.TestSetTreeManager

                If idx = 1 And iDateFolder <> 1 Then
                    myRoot = TStmgr.Root
                Else
                    myRoot = TStmgr.NodeByPath(OldPath)
                End If ' idx'

                newNode = myRoot.addNode(CurrentSubName)
                newNode.post
            End If 'new Node

            ' if the current folder is the last folder of the array
            ' create a testset (if necessary) and add the current test

            If idx = UBound(subjectArray) Then
                'Check: Does the testset exist?    ' create a filter with Folder-id and -name
                testSetF = newNode.TestSetFactory
                testSetFilter = testSetF.Filter
                testSetFilter.Filter("CY_FOLDER_ID") = newNode.Nodeid
                testSetFilter.Filter("CY_CYCLE") = CurrentSubName
                TSList = testSetF.newList(testSetFilter.Text)

                'Add Testset only if necessary

                If TSList.Count = 0 Then
                    'MsgBox("Add Testset")
                    'nothing found'
                    testSet1 = testSetF.AddItem(System.DBNull.Value)
                    testSet1.Name = CurrentSubName
                    testSet1.Status = "Open"

                    testSet1.Field("CY_USER_01") = sTSVDIName         'Nazwa maszyny
                    testSet1.Field("CY_USER_02") = sTSVDINum    'Maksymalna ilość maszyn
                    testSet1.Field("CY_USER_03") = sStartDate            'Data uruchomienia
                    testSet1.Field("CY_USER_04") = sTSDataSource         'Wybór źródła danych
                    testSet1.Field("CY_USER_05") = sTSPriorytet             'Priorytet uruchomienia
                    testSet1.Field("CY_USER_06") = sTestGroup            'Grupa Testów
                    testSet1.Field("CY_USER_08") = sTSEmail            'Wyślij mail
                    testSet1.Field("CY_USER_09") = sTSEnvirnoment            'Środowisko
                    testSet1.Field("CY_USER_10") = sTSLogLevel              'Poziom logowania
                    testSet1.Field("CY_USER_11") = sPeriodic              'Cykliczność
                    testSet1.Field("CY_COMMENT") = sRepeatingDate    'DatyCykliczności
                    testSet1.Field("CY_USER_12") = sTSSystem               'Testowany System
                    testSet1.Field("CY_USER_14") = sTSReUseData            'Reużyj dane
                    testSet1.Post
                    testSet1.refresh
                Else
                    'else get it
                    testSet1 = TSList.Item(1)
                End If 'TSList

                'Check: testinstance
                'DO not use FindTestInstance (way too much overhead)
                TSTestF = testSet1.TSTestFactory
                TSTestList = TSTestF.newList("")

                'initialize marker
                foundTS = 0

                If TSTestList.Count > 0 Then

                    For Each myTSTest In TSTestList
                        If myTSTest.testId = Trim(sTestCase.ID & " ") Then
                            foundTS = 1
                        End If
                    Next ' myTSTest
                End If ' TSTestList

                'Add Test if necessary

                If foundTS = 0 Then
                    'nothing found => add test to testset
                    newTSTest = TSTestF.AddItem(sTestCase.ID)
                    newTSTest.Post
                End If ' foundTS
            End If ' idx

            '-------------------------------------------------?-----
            'Cleanup for objects (just to be sure)
            newTSTest = Nothing
            myTSTest = Nothing
            testSetFilter = Nothing
            TSTestF = Nothing
            TSTestList = Nothing
            testSetFilter = Nothing
            testSetF = Nothing
            folder = Nothing
            newNode = Nothing

        Next 'idx

        On Error GoTo 0
        build_case = True
    End Function

End Class

Public Class CustomizationList
    Sub GetCustomizationList(ListName As String)
        On Error Resume Next
        Dim qcURL As String
        Dim qcID As String
        Dim qcPWD As String
        Dim qcDomain As String
        Dim qcProject As String
        Dim TestSetFound, strMsg, iFolderID

        Dim tdConnection = New TDAPIOLELib.TDConnection
        Dim ret

        qcURL = "http://profive1v5.radom.tsunami:8080/qcbin"
        qcID = "uft"
        qcPWD = "uft"
        qcDomain = "UFT"
        qcProject = "UFT"
        tdConnection.InitConnectionEx(qcURL)
        tdConnection.Login(qcID, qcPWD)
        tdConnection.Connect(qcDomain, qcProject)

        Dim cust As Customization
        Dim custFields As CustomizationFields
        Dim aCustField As CustomizationField
        Dim custlists As CustomizationLists
        Dim aCustList 'As CustomizationLists
        Dim aListNode As TDAPIOLELib.CustomizationListNode
        'Dim listName$, i%
        Dim msg As String
        Dim c As CustomizationListNode

        cust = tdConnection.Customization
        cust.Load()
        custlists = cust.Lists
        aCustList = custlists.List(ListName)
        aListNode = aCustList.RootNode

        For Each c In aListNode.Children
            MsgBox(c.Name) 'tutaj dodać jakaś funkcję ładującą i wyświetlającą w MSSQL, tudzież zwrotka jako lista
        Next
        'cust.Commit()

        tdConnection.Disconnect()
        tdConnection.Logout()
        tdConnection.ReleaseConnection()
        On Error GoTo 0
    End Sub

    Function AddItemToList(ListName As String, ItemName As String) _
      As CustomizationListNode
        On Error Resume Next
        Dim qcURL As String
        Dim qcID As String
        Dim qcPWD As String
        Dim qcDomain As String
        Dim qcProject As String
        Dim TestSetFound, strMsg, iFolderID
        Dim tdConnection = New TDAPIOLELib.TDConnection
        Dim ret

        qcURL = "http://profive1v5.radom.tsunami:8080/qcbin"
        qcID = "uft"
        qcPWD = "uft"
        qcDomain = "UFT"
        qcProject = "UFT"
        tdConnection.InitConnectionEx(qcURL)
        tdConnection.Login(qcID, qcPWD)
        tdConnection.Connect(qcDomain, qcProject)

        Dim cust As Customization
        Dim custFields As CustomizationFields
        Dim aCustField As CustomizationField
        Dim custlists As CustomizationLists
        Dim aCustList 'As CustomizationLists
        Dim aListNode As TDAPIOLELib.CustomizationListNode
        Dim msg As String
        Dim c As CustomizationListNode

        cust = tdConnection.Customization
        cust.Load()
        custlists = cust.Lists
        aCustList = custlists.List(ListName)
        aListNode = aCustList.RootNode
        AddItemToList = aListNode.AddChild(ItemName)
        cust.Commit()

        tdConnection.Disconnect()
        tdConnection.Logout()
        tdConnection.ReleaseConnection()
        On Error GoTo 0
    End Function

End Class

Public Class GeneratoryDanych
    Function GetPesel() As Long
        'cyfry [1-6] – data urodzenia : Kolejne pary cyfr oznaczają kolejno rok, miesiąc i dzień urodzenia.
        'cyfry [7-9] – numer serii(możemy go traktować jako pojedyncze cyfry). Mogą być to dowolne cyfry. Mają one znaczenie tylko przy obliczaniu cyfry kontrolnej.
        'cyfra [10] – płeć. Cyfry parzyste wraz z zerem oznaczają płeć żenską, natomiast wszystkie cyfry nieparzyste oznaczają płeć męską.
        'cyfra [11] – cyfra kontrolna, służąca do weryfikacji numeru PESEL
        Dim sPESEL As String
        Dim dDataUrodzenia
        Dim iWypelnienie1 As Integer
        Dim iPlec As Integer

        dDataUrodzenia = Right(Replace(RandomDate("1900-01-01", "1999-12-31"), "-", ""), 6) 'GetBirthDate()
        iWypelnienie1 = Randomize("100", "999") 'środek bez weryfikacji 3 cyfry
        iPlec = Randomize("0", "9") 'płeć
        sPESEL = dDataUrodzenia & iWypelnienie1 & iPlec

        Dim iSumaKon As Integer
        Dim iWagi = New Integer() {1, 3, 7, 9, 1, 3, 7, 9, 1, 3}
        For i = LBound(iWagi) To UBound(iWagi)
            iSumaKon = iSumaKon + (iWagi(i) * CInt(Mid(sPESEL, i + 1, 1)))
        Next i
        iSumaKon = iSumaKon Mod 10
        iSumaKon = 10 - iSumaKon
        sPESEL = sPESEL & iSumaKon

        Return sPESEL
    End Function

    Function GetNIP() As Long
        Dim sNIP As String, iLen As Integer
        Dim iWypelnienie1 As Integer
        iWypelnienie1 = Randomize("0", "999999999") 'środek bez weryfikacji
        For i = 1 To 9
            iLen = (CStr(iWypelnienie1)).Length
            If iLen = 9 Then Exit For
            iWypelnienie1 = iWypelnienie1 & "0"
        Next

        sNIP = iWypelnienie1
        Dim iSumaKon As Integer
        Dim iWagi = New Integer() {6, 5, 7, 2, 3, 4, 5, 6, 7}
        For i = LBound(iWagi) To UBound(iWagi)
            iSumaKon = iSumaKon + (iWagi(i) * CInt(Mid(sNIP, i + 1, 1)))
        Next i
        iSumaKon = iSumaKon Mod 11
        If iSumaKon > 9 Then GetNIP() 'jesli z modulo wychodzi 10 powtórz

        sNIP = sNIP & iSumaKon
        Return sNIP
    End Function

    Function GetREGON() As Long
        Dim sRegon As String, iLen As Integer
        Dim iWypelnienie1 As Integer
        iWypelnienie1 = Randomize("0", "99999999") 'środek bez weryfikacji
        For i = 1 To 8
            iLen = (CStr(iWypelnienie1)).Length
            If iLen = 8 Then Exit For
            iWypelnienie1 = iWypelnienie1 & "0"
        Next

        Dim iSumaKon As Integer
        Dim iWagi = New Integer() {8, 9, 2, 3, 4, 5, 6, 7}
        sRegon = iWypelnienie1

        For i = LBound(iWagi) To UBound(iWagi)
            iSumaKon = iSumaKon + (iWagi(i) * CInt(Mid(sRegon, i + 1, 1)))
        Next i
        iSumaKon = iSumaKon Mod 11
        If iSumaKon = 10 Then iSumaKon = 0
        sRegon = sRegon & iSumaKon

        Return sRegon
    End Function

    Function GetDowod() As String
        Dim sDowod As String
        Dim iSumaKon As Integer
        Dim iWagi = New Integer() {7, 3, 1, 7, 3} '7, 3, 1
        Dim sChar1 As String() = GetChar(Randomize("1", "26")).Split(",")
        Dim sChar2 As String() = GetChar(Randomize("1", "26")).Split(",")
        Dim sChar3 As String() = GetChar(Randomize("1", "26")).Split(",")
        sDowod = Randomize("10000", "99999")

        iSumaKon = sChar1(1) * 7
        iSumaKon = iSumaKon + (sChar2(1) * 3)
        iSumaKon = iSumaKon + (sChar3(1) * 1)

        For i = LBound(iWagi) To UBound(iWagi)
            iSumaKon = iSumaKon + (iWagi(i) * CInt(Mid(sDowod, i + 1, 1)))
        Next i
        iSumaKon = iSumaKon Mod 10

        sDowod = sChar1(0) & sChar2(0) & sChar3(0) & iSumaKon & sDowod

        Return sDowod
    End Function

    Public Function RandomDate(ByVal StartDate, ByVal EndDate) As Date
        'returns random date between start date and now

        If Not IsDate(StartDate) Then Exit Function
        If Not IsDate(EndDate) Then Exit Function
        Dim dtStartDate = CDate(StartDate)
        Dim dtEndDate = CDate(EndDate)
        Dim iDifferential = DateDiff(DateInterval.Day, dtStartDate, dtEndDate)
        iDifferential = New Random(System.DateTime.Now.Millisecond).Next(0, iDifferential)
        dtStartDate = DateAdd(DateInterval.Day, iDifferential, dtStartDate)
        Return dtStartDate
    End Function

    Public Function Randomize(ByVal liczbaOd As Long, ByVal liczbaDo As Long) As Long
        Static Generator As System.Random = New System.Random()
        Return Generator.Next(liczbaOd, liczbaDo)
    End Function

    Function GetChar(iZnak) As String
        ' W literach serii nie używa się liter 'O' i 'Q'. Jeżeli więc ktoś poda takie litery w serii to z pewnością dane są błędne.
        ' http://zylla.wipos.p.lodz.pl/ut/paszport.html#dowodosobisty
        Dim sZnak As String = "A,10"
        Select Case iZnak
            Case 1
                sZnak = ("A,10")
            Case 2
                sZnak = ("B,11")
            Case 3
                sZnak = ("C,12")
            Case 4
                sZnak = ("D,13")
            Case 5
                sZnak = ("E,14")
            Case 6
                sZnak = ("F,15")
            Case 7
                sZnak = ("G,16")
            Case 8
                sZnak = ("H,17")
            Case 9
                sZnak = ("I,18")
            Case 10
                sZnak = ("J,19")
            Case 11
                sZnak = ("K,20")
            Case 12
                sZnak = ("L,21")
            Case 13
                sZnak = ("M,22")
            Case 14
                sZnak = ("N,23")
            Case 15
             '   sZnak = ("O,24") 'pobierze wartość domyślną A10
            Case 16
                sZnak = ("P,25")
            Case 17
               ' sZnak = ("Q,26")'pobierze wartość domyślną A10
            Case 18
                sZnak = ("R,27")
            Case 19
                sZnak = ("S,28")
            Case 20
                sZnak = ("T,29")
            Case 21
                sZnak = ("U,30")
            Case 22
                sZnak = ("V,31")
            Case 23
                sZnak = ("W,32")
            Case 24
                sZnak = ("X,33")
            Case 25
                sZnak = ("Y,34")
            Case 26
                sZnak = ("Z,35")
        End Select
        Return sZnak
    End Function
End Class