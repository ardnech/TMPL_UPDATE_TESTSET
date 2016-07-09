Imports System.Net
Imports System.Data.SqlClient

Public Class Form1
    Public clTDConnectivity As New TDConnectivity
    Sub GetTsTcFromTestLab()

        Call clTDConnectivity.ConnectToTD()

        Dim tsTreeMgr, theTestSet
        Dim separateChars1 As String = "$"
        Dim separateChars2 As String = "<"
        '  On Error GoTo err

        Dim TSetFact
        Dim tsFolder
        Dim tsList
        ' Get the test set tree manager from the test set factory
        'tdconnection is the global TDConnection object.
        TSetFact = clTDConnectivity.tdConnection.TestSetFactory
        tsTreeMgr = clTDConnectivity.tdConnection.testsettreemanager

        tsFolder = tsTreeMgr.NodeByPath(clTDConnectivity.nPath)
        '--------------------------------Check if the Path Exists Or Not ---------------------------------------------------------------------
        If tsFolder Is Nothing Then
            Debug.Print("Brak Folderu: " & clTDConnectivity.nPath)
        End If

        ' Search for the test set passed as an argument to the example code
        tsList = tsFolder.FindTestSets("")
        '----------------------------------Check if the Test Set Exists --------------------------------------------------------------------
        If tsList Is Nothing Then
            Debug.Print("Brak TestSet w Folderze: " & clTDConnectivity.nPath)
        End If
        '-------------------------------------------Access the Test Cases inside the Test SEt -------------------------------------------------

        Dim tsTestFactory
        Dim tsTestList

        theTestSet = tsList.Item(1)


        Dim sTS_Name As String, sTS_Path As String, sTS_Param As String, sTS_Status As String, sTS_ALM_ID As Long
        Dim sTC_Name As String, sTC_Path As String, sTC_Param As String, sTC_Status As String, sTC_ALM_ID As Long
        Dim sqlConnScenarioID As String, sqlConnCaseID As String

        Dim dbConn As New dbConnectivity

        For Each testsetfound In tsList
            tsFolder = testsetfound.TestSetFolder
            tsTestFactory = testsetfound.tsTestFactory
            tsTestList = tsTestFactory.NewList("")
            'Debug.Print(StrTok(testsetfound.Field("CY_COMMENT"), separateChars1, separateChars2))

            sTS_Status = testsetfound.Status
            If UCase(sTS_Status) = UCase("Open") Then
                sTS_Name = testsetfound.Name
                sTS_Path = tsFolder.path
                sTS_Param = (StrTok(testsetfound.Field("CY_COMMENT"), separateChars1, separateChars2))
                sTS_ALM_ID = testsetfound.ID


                'Rozpoczęcie wrzucania TSTS do DB
                sqlConnScenarioID = dbConn.sqlConnScenario(sTS_Name, sTS_Path, sTS_Param, sTS_Status, sTS_ALM_ID)

                'Wrzucenie TC
                For Each tsTest In tsTestList
                    sTC_Name = tsTest.Name
                    sTC_Path = tsFolder.path
                    sTC_Param = (StrTok(tsTest.field("TS_DESCRIPTION"), separateChars1, separateChars2))
                    If sTC_Param = "" Then sTC_Param = ("$repeat=1")
                    sTC_Status = tsTest.Status
                    sTC_ALM_ID = tsTest.ID

                    'Zapis TC do DB MSSQL
                    sqlConnCaseID = dbConn.sqlConnCase(sqlConnScenarioID, sTC_Name, sTC_Param, sTC_ALM_ID)
                Next tsTest

                'Potwierdzenie wrzucenie TS i TC
                sqlConnCaseID = dbConn.sqlConnScenarioCompleted(sqlConnScenarioID)
                Debug.Print(sqlConnScenarioID)
            End If


        Next testsetfound
        Call clTDConnectivity.DisconnectfromTD()

        Exit Sub
err:
        MsgBox(Err.Description)
        Call clTDConnectivity.DisconnectfromTD()
        Exit Sub
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
        Dim clStatusChange As New StatusChange
        Call clStatusChange.ChangeStatusScenario("101", "Not Completed")
        Call clStatusChange.ChangeStatusCase("4", "Not Completed")
        'GetTsTcFromTestLab()
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
        'On successfull login display message in Status bar
        ' Application.StatusBar = "........QC Connection is done Successfully"
        '  MsgBox("Connection Established")

        '---------------------------------------Connection Established --------------------------------------------------------------------------
        'ConnectToQualityCenter(qcURL, qcID, qcPWD, qcDomain, qcProject, nPath)
    End Sub

    Public Sub DisconnectfromTD()
        '------------------------------------------------------Disconnect Quality Center -----------------------------------------------------------------
        tdConnection.Disconnect
        tdConnection.Logout
        tdConnection.ReleaseConnection
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
        ' string używany do odpalenia SP na bazie
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
    Sub ChangeStatusScenario(ByVal qcTSId As Long, ByVal qcStatus As String)
        Call clTDConnectivity.ConnectToTD()

        Dim TSetFact, TestSetFilter, TestSetList, myTestSet
        TSetFact = clTDConnectivity.tdConnection.TestSetFactory
        TestSetFilter = TSetFact.Filter
        TestSetFilter.Filter("CY_CYCLE_ID") = qcTSId

        TestSetList = TestSetFilter.NewList
        myTestSet = TestSetList.Item(1)
        myTestSet.Field("CY_STATUS") = (qcStatus)
        myTestSet.Post

        Call clTDConnectivity.DisconnectfromTD()
        TestSetFilter = Nothing
        myTestSet = Nothing
        TSetFact = Nothing
        'data
    End Sub

    Sub ChangeStatusCase(ByVal qcTCId As Long, ByVal qcStatus As String)
        Call clTDConnectivity.ConnectToTD()

        Dim TTestFact, TestFilter, TestList, myTest
        'TSetFact = clTDConnectivity.tdConnection.TestSetFactory
        TTestFact = clTDConnectivity.tdConnection.TestFactory

        'TestSetFilter = TSetFact.Filter
        TestFilter = TTestFact.Filter

        'TestSetFilter.Filter("CY_CYCLE_ID") = qcTCId
        TestFilter.Filter("TC_TESTCYCL_ID") = qcTCId

        'TestSetList = TestSetFilter.NewList
        TestList = TestFilter.NewList

        'myTestSet = TestSetList.Item(1)
        myTest = TestList.Item(1)
        MsgBox(myTest.name)
        'myTestSet.Field("TS_TEST_ID") = (qcStatus)
        myTest.Field("TC_STATUS") = (qcStatus)

        'myTestSet.Post
        myTest.Post

        Call clTDConnectivity.DisconnectfromTD()
        TestFilter = Nothing
        myTest = Nothing
        TTestFact = Nothing
    End Sub
End Class