'Imports System
'Imports System.Data.SqlClient

''<ComClass(JConnection.ClassId, JConnection.InterfaceId, JConnection.EventsId)> _
'Public Class JConnection

'#Region "COM GUIDs"
'    ' These  GUIDs provide the COM identity for this class and its COM interfaces. 
'    ' If you change them, existing clients will no longer be able to access the class.
'Public Const ClassId As String = "025DC407-3C8B-4590-8B19-CAF4C95333F1"
'Public Const InterfaceId As String = "C232EE0E-31B5-4460-9452-188E8054C044"
'Public Const EventsId As String = "2293686F-0014-464E-957D-133A1D792553"
'#End Region

'    ' A creatable COM class must have a Public Sub New() with no parameters, 
'    ' otherwise, the class will not be registered in the COM registry and cannot be created via CreateObject.
'Public moConnection As ADODB.Connection
'Public cFormLogin As Object


'Public Sub New()
'    MyBase.New()
'End Sub
''Public ReadOnly Property SQLCon() As SqlConnection
''    Get
''        Return SQLConn.GetInstance()
''    End Get
''End Property

'Public Sub setLogin(ByVal frmLogin As Object)
'    cFormLogin = frmLogin
'End Sub

'Public Function InitializeSQLConnection() As Boolean
'    Dim uName As String = String.Empty
'    Dim pWord As String = String.Empty
'    Dim cServer As String = String.Empty
'    Dim cDb As String = String.Empty
'    Dim ctl As Control

'    Try
'        'With cFormLogin.CInstance
'        'For Each ctl In cFormLogin.CInstance.Controls
'        For Each ctl In cFormLogin.controls
'            If (TypeOf ctl Is TextBox) Then
'                Select Case ctl.Name.ToString
'                    Case "txtPassword"
'                        pWord = ctl.Text
'                        Case "txtUserName"
'                            uName = ctl.Text
'                    End Select
'            ElseIf (TypeOf ctl Is ComboBox) Then
'                Select Case ctl.Name.ToString
'                    Case "cbServer"
'                        cServer = ctl.Text
'                    Case "cbDatabase"
'                        cDb = ctl.Text
'                End Select
'            End If
'        Next
'        'End With
'    Catch ex As Exception
'        Trace.WriteLine(ex.Source + ":" + ex.Message)
'        Throw ex
'    End Try

'    Dim sConnection As String
'    sConnection = ""

'    If Trim(uName) <> "" Then
'        sConnection = sConnection & "User ID=" & uName & ";"
'        sConnection = sConnection & "Password=" & pWord & ";"
'    Else
'        sConnection = sConnection & "Trusted_Connection=Yes" & ";"
'    End If
'    sConnection = sConnection & "Data Source=" & cServer & ";"
'    sConnection = sConnection & "Initial Catalog=" & cDb
'    'SQLConn.GetInstance(sConnection)

'End Function

'Public Function WMS() As ADODB.Connection
'    Dim uName As String = String.Empty
'    Dim pWord As String = String.Empty
'    Dim cServer As String = String.Empty
'    Dim cDb As String = String.Empty
'    Dim ctl As Control

'    On Error GoTo Err_Sub
'    Err.Clear()
'    If (moConnection Is Nothing) Then
'        moConnection = New ADODB.Connection
'        'With cFormLogin.CInstance
'        'For Each ctl In cFormLogin.CInstance.Controls
'        For Each ctl In cFormLogin.Controls
'            If (TypeOf ctl Is TextBox) Then
'                Select Case ctl.Name.ToString
'                    Case "txtPassword"
'                        pWord = ctl.Text
'                        Case "txtUserName"
'                            uName = ctl.Text
'                    End Select
'            ElseIf (TypeOf ctl Is ComboBox) Then
'                Select Case ctl.Name.ToString
'                    Case "cbServer"
'                        cServer = ctl.Text
'                    Case "cbDatabase"
'                        cDb = ctl.Text
'                End Select
'            End If
'        Next
'        Call OpenConnection(uName, pWord, cServer, cDb)
'            '            Call OpenConnection(.txtUserName.Text, .txtPassword.Text, .cbServer.Text, .cbDatabase.Text)
'        'End With
'    End If

'    Return moConnection

'Err_Sub:

'    If Err.Number <> 0 Then
'        '            Dim moMSCErrMsc As JCommonLib = New JCommonLib
'        'Call moMSCErrMsc.MSCUserMessage.ErrorHandler("WMS", Err.Description, Err.Number, True)
'        Err.Clear()
'    End If

'End Function

'Public Function OpenConnection(ByRef sUserName As String, ByRef sPassword As String, ByRef sServer As String, ByRef sDatabase As String) As Boolean
'    Dim sConnection As String
'    sConnection = ""
'    OpenConnection = True
'    If Trim(sUserName) <> "" Then
'        sConnection = sConnection & "User ID=" & sUserName & ";"
'        sConnection = sConnection & "Password=" & sPassword & ";"
'    Else
'        sConnection = sConnection & "Trusted_Connection=Yes" & ";"
'    End If
'    sConnection = sConnection & "Data Source=" & sServer & ";"
'    sConnection = sConnection & "Initial Catalog=" & sDatabase

'    On Error GoTo Err_Sub
'    Err.Clear()

'    '      moConnection = New ADODB.Connection

'    With moConnection
'        .Provider = "SQLOLEDB"
'        .ConnectionString = sConnection
'        .CursorLocation = ADODB.CursorLocationEnum.adUseClient
'    End With
'    moConnection.Open()
'Err_Sub:

'    If Err.Number <> 0 Then
'        Dim sErrDescription As String
'        Dim iErrNumber As Int32
'        iErrNumber = Err.Number
'        sErrDescription = Err.Description
'        'Dim moMSCErrMsc As JCommonLib = New JCommonLib
'        'Call moMSCErrMsc.MSCUserMessage.ErrorHandler("OpenConnection", sErrDescription, iErrNumber, True)
'        OpenConnection = False
'        Err.Clear()
'    End If

'End Function

'Public Sub CloseConnection()

'    On Error Resume Next
'    Err.Clear()

'    If moConnection Is Nothing Then
'        moConnection.Close()
'    End If
'    moConnection = Nothing

'End Sub
'End Class


