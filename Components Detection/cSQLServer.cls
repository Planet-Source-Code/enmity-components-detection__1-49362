VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum SQLServerLoginMode
    sslmSystemDefault
    sslmWindowsIV
    sslmMixMode
End Enum

Private Const MODULE_NAME As String = "mdEnumServers"

'--- retvals
Private Const SQL_ERROR                     As Integer = -1
Private Const SQL_INVALID_HANDLE            As Integer = -2
Private Const SQL_NEED_DATA                 As Integer = 99
Private Const SQL_NO_DATA_FOUND             As Integer = 100
Private Const SQL_SUCCESS                   As Integer = 0
Private Const SQL_SUCCESS_WITH_INFO         As Integer = 1
'--- for SQLSetConnectOption
Private Const SQL_ATTR_LOGIN_TIMEOUT        As Long = 103
Private Const SQL_ATTR_CONNECTION_TIMEOUT   As Long = 113
Private Const SQL_ATTR_QUERY_TIMEOUT        As Long = 0
Private Const SQL_COPT_SS_BASE              As Long = 1200
Private Const SQL_COPT_SS_INTEGRATED_SECURITY As Long = (SQL_COPT_SS_BASE + 3) ' Force integrated security on login
Private Const SQL_COPT_SS_BASE_EX           As Long = 1240
Private Const SQL_COPT_SS_BROWSE_CACHE_DATA As Long = (SQL_COPT_SS_BASE_EX + 5) ' Determines if we should cache browse info. Used when returned buffer is greater then ODBC limit (32K)
'--- param type
Private Const SQL_IS_UINTEGER               As Integer = (-5)
Private Const SQL_IS_INTEGER                As Integer = (-6)
'--- for SQL_COPT_SS_INTEGRATED_SECURITY
Private Const SQL_IS_OFF                    As Long = 0
Private Const SQL_IS_ON                     As Long = 1
'--- for SQL_COPT_SS_BROWSE_CACHE_DATA
Private Const SQL_CACHE_DATA_NO             As Long = 0
Private Const SQL_CACHE_DATA_YES            As Long = 1
'--- for SQLSetEnvAttr
Private Const SQL_ATTR_ODBC_VERSION         As Long = 200
Private Const SQL_OV_ODBC3                  As Long = 3

Private Declare Function SQLAllocEnv Lib "odbc32.dll" (phEnv As Long) As Integer
Private Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal hEnv As Long, phDbc As Long) As Integer
Private Declare Function SQLSetEnvAttr Lib "odbc32" (ByVal EnvironmentHandle As Long, ByVal Attrib As Long, Value As Any, ByVal StringLength As Long) As Integer
Private Declare Function SQLBrowseConnect Lib "odbc32.dll" (ByVal hDbc As Long, ByVal szConnStrIn As String, ByVal cbConnStrIn As Integer, ByVal szConnStrOut As String, ByVal cbConnStrOutMax As Integer, pcbConnStrOut As Integer) As Integer
Private Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hDbc As Long) As Integer
Private Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hDbc As Long) As Integer
Private Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As Long) As Integer
Private Declare Function SQLSetConnectOption Lib "odbc32.dll" (ByVal ConnectionHandle As Long, ByVal Option_ As Integer, ByVal Value As Long) As Integer
Private Declare Function SQLGetConnectOption Lib "odbc32.dll" (ByVal ConnectionHandle As Long, ByVal Option_ As Integer, Value As Long) As Integer
Private Declare Function SQLError Lib "odbc32.dll" (ByVal EnvironmentHandle As Long, ByVal ConnectionHandle As Long, ByVal StatementHandle As Long, ByVal Sqlstate As String, NativeError As Long, ByVal MessageText As String, ByVal BufferLength As Integer, TextLength As Integer) As Integer
'--- ODBC 3.0
Private Declare Function SQLSetConnectAttr Lib "odbc32" Alias "SQLSetConnectAttrA" (ByVal ConnectionHandle As Long, ByVal Attrib As Long, Value As Any, ByVal StringLength As Long) As Integer
Private Declare Function SQLGetConnectAttr Lib "odbc32" Alias "SQLGetConnectAttrA" (ByVal ConnectionHandle As Long, ByVal Attrib As Long, Value As Any, ByVal BufferLength As Long, StringLength As Long) As Integer

Private Const STR_NO_USER_DBS           As String = "<No user databases>"

Private m_bEnumSrv As Boolean




Private Sub RaiseError(sFunc As String)
    Err.Raise Err.Number, MODULE_NAME & "." & sFunc & vbCrLf & Err.Source, Err.Description
End Sub


Public Function EnumSqlServers() As Variant
    
    Const FUNC_NAME     As String = "EnumSqlServers"
    Const CONN_STR      As String = "DRIVER={SQL Server}"
    Const PREFIX        As String = "Server={"
    Const SUFFIX        As String = "}"
    
    On Error GoTo HandleError
        
        EnumSqlServers = pvBrowseConnect(CONN_STR, PREFIX, SUFFIX)
        
    Exit Function
    
HandleError:

    'RaiseError FUNC_NAME

End Function


Public Function EnumSqlDbs(ByVal sServer As String, _
                           Optional ByVal sUser As String, _
                           Optional ByVal sPass As String) As Variant
    
    Const FUNC_NAME     As String = "EnumSqlDbs"
    Const CONN_STR      As String = "DRIVER={SQL Server};SERVER=%1;UID=%2;PWD=%3;"
    Const PREFIX        As String = "Database={"
    Const SUFFIX        As String = "}"
    Dim sConnStr        As String
    
    On Error GoTo HandleError
    
        EnumSqlDbs = pvBrowseConnect( _
                Replace(Replace(Replace(CONN_STR, _
                        "%1", sServer), _
                        "%2", sUser), _
                        "%3", sPass), _
                PREFIX, SUFFIX, Len(sUser) = 0)
    
    Exit Function

HandleError:
    
    'RaiseError FUNC_NAME

End Function


Private Function pvBrowseConnect(ByVal sConnStr As String, _
                                 ByVal sPrefix As String, _
                                 ByVal sSuffix As String, _
                                 Optional ByVal bItegrated As Boolean)
    Const FUNC_NAME     As String = "pvBrowseConnect"
    Dim rc              As Integer
    Dim hEnv            As Long
    Dim hDbc            As Long
    Dim sBuffer         As String
    Dim nReqBufSize     As Integer
    Dim lStart          As Long
    Dim lEnd            As Long
    Dim dwSec           As Long
    Dim lstrlen         As Long

    On Error GoTo HandleError
    '--- init environment
    rc = SQLAllocEnv(hEnv)
    rc = SQLSetEnvAttr(hEnv, SQL_ATTR_ODBC_VERSION, ByVal SQL_OV_ODBC3, SQL_IS_INTEGER)
    '--- init conn
    rc = SQLAllocConnect(hEnv, hDbc)
    '--- timeouts to ~5 secs
    rc = SQLSetConnectOption(hDbc, SQL_ATTR_CONNECTION_TIMEOUT, 3)
    rc = SQLSetConnectOption(hDbc, SQL_ATTR_LOGIN_TIMEOUT, 3)
    '--- integrated security
    If bItegrated Then
        rc = SQLSetConnectOption(hDbc, SQL_COPT_SS_INTEGRATED_SECURITY, SQL_IS_ON)
    End If
    '--- improve performance
    rc = SQLSetConnectOption(hDbc, SQL_COPT_SS_BROWSE_CACHE_DATA, SQL_CACHE_DATA_YES)
    '--- initial buffer size
    nReqBufSize = 1000
    '--- repeat getting info until buffer gets large enough
    Do
        sBuffer = String(nReqBufSize + 1, 0)
        rc = SQLBrowseConnect(hDbc, sConnStr, Len(sConnStr), sBuffer, Len(sBuffer), nReqBufSize)
    Loop While rc = SQL_NEED_DATA And nReqBufSize >= Len(sBuffer)
    '--- if ok -> parse buffer
    If rc = SQL_SUCCESS Or rc = SQL_NEED_DATA Then
        '--- find prefix
        lStart = InStr(1, sBuffer, sPrefix)
        If lStart > 0 Then
            lStart = lStart + Len(sPrefix)
            '--- find suffix
            lEnd = InStr(lStart, sBuffer, sSuffix)
            If lEnd > 0 Then
                lEnd = lEnd - Len(sSuffix) + 1
                '--- success
                pvBrowseConnect = Split(Mid(sBuffer, lStart, lEnd - lStart), ",")
            End If
        Else
            Err.Raise vbObjectError, "ODBC", pvGetError(rc, hEnv, hDbc, 0)
        End If
    End If
    '--- disconnect
    rc = SQLDisconnect(hDbc)
    '--- free handles
    rc = SQLFreeConnect(hDbc)
    rc = SQLFreeEnv(hEnv)
    '--- on failure -> return Array(0 To -1)
    If Not IsArray(pvBrowseConnect) Then
        pvBrowseConnect = Split("")
    End If
    
    Exit Function
    
HandleError:
    'RaiseError FUNC_NAME
    
End Function


Public Function DoesSQLServer2000Exist2() As Boolean
        
    Dim vSrv            As Variant
    Dim o_lngItems As Long
    Dim o_blnRet As Boolean
    
    On Error GoTo HandleError
        
        o_blnRet = False
        
        If Not m_bEnumSrv Then
            Screen.MousePointer = vbHourglass
            For Each vSrv In EnumSqlServers
                If LCase(CStr(vSrv)) = LCase("(local)") Then
                    o_blnRet = True
                    Exit For
                Else
                End If
            Next
            m_bEnumSrv = True
            Screen.MousePointer = vbDefault
        End If
        
        DoesSQLServer2000Exist2 = o_blnRet
    
    Exit Function
    
HandleError:

    DoesSQLServer2000Exist2 = False

End Function


Public Function GetSQLServerLoginMode() As SQLServerLoginMode
    
    Dim strRet As String
    
    With New cRegistry
        .hKey = HKEY_LOCAL_MACHINE
        
        .KeyPath = "\SOFTWARE\Microsoft\MSSQLServer\MSSQLServer"
        
        strRet = GetRegValue(.GetRegistryValue("LoginMode"))
        
        Select Case strRet
            Case "0x1"
                GetSQLServerLoginMode = sslmWindowsIV
            Case "0x2"
                GetSQLServerLoginMode = sslmMixMode
            Case Else
                GetSQLServerLoginMode = sslmSystemDefault
        End Select
    End With
    
End Function


Private Function GetRegValue(ByVal strValue As String) As String
    
    On Error Resume Next
    
    If strValue <> "" Then
        If InStr(1, strValue, ":") > 0 Then
            GetRegValue = Right(strValue, Len(strValue) - InStr(1, strValue, ":") - 1)
        Else
            GetRegValue = strValue
        End If
    Else
        GetRegValue = ""
    End If
    
End Function


Public Function SetSQLServerLoginMode(ByVal udeLoginMode As SQLServerLoginMode) As Boolean
    
    Dim intMode As Integer
    
    With New cRegistry
    
        .hKey = HKEY_LOCAL_MACHINE
        
        .KeyPath = "\SOFTWARE\Microsoft\MSSQLServer\MSSQLServer"
        
        Select Case udeLoginMode
            Case sslmWindowsIV
                intMode = 1
            Case sslmMixMode
                intMode = 2
            Case Else 'sslmSystemDefault
                intMode = 0
        End Select
        
        SetSQLServerLoginMode = .SetRegistryValue("LoginMode", intMode, REG_DWORD)
        
    End With

End Function


Public Function DoesSQLServer2000Exist() As Boolean
        
    Dim vSrv            As Variant
    Dim o_lngItems As Long
    
    On Error GoTo HandleError
    
        If Not m_bEnumSrv Then
            Screen.MousePointer = vbHourglass
            For Each vSrv In EnumSqlServers
                o_lngItems = o_lngItems + 1
            Next
            m_bEnumSrv = True
            Screen.MousePointer = vbDefault
        End If
        
        DoesSQLServer2000Exist = (o_lngItems > 0)
    
    Exit Function
    
HandleError:

    DoesSQLServer2000Exist = False

End Function


Public Function DoesSQLServer2000ExistEx(ByVal strServer As String, _
                                         ByVal strDatabase As String, _
                                         ByVal strUser As String, _
                                         ByVal strPassword As String) As Boolean

    On Error GoTo HandleError
        
        Dim o_blnRet As Boolean
        
        With New cDBConnection
            .Server = strServer
            .Database = strDatabase
            .UserName = strUser
            .Password = strPassword
            .IsAccessMDB = False
            o_blnRet = .StartConnection(False)
            .StopConnection
        End With
        
        DoesSQLServer2000ExistEx = o_blnRet
        
    Exit Function
    
HandleError:

    DoesSQLServer2000ExistEx = False

End Function


Public Function GetMSSQLServerVersion(ByVal strServer As String, _
                                      ByVal strUserName As String, _
                                      ByVal strPassword As String) As String
    
    Dim cn As Object
    Dim strRet As String
    Dim strSQL As String
    
    On Error GoTo HandleError
    
    Set cn = CreateObject("ADODB.CONNECTION")
    
    If Not (cn Is Nothing) Then
        cn.ConnectionTimeout = 3
        
        cn.CursorLocation = 3 'adUseClient
        
        'cn.Open "Provider=SQLOLEDB;Persist Security Info=false;Integrated Security=SSPI;Data Source=" & strServer & ""
        cn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & strUserName & ";Password=" & strPassword & ";Initial Catalog=master;Data Source=" & strServer
        If Err.Number = 0 Then
            strSQL = "select @@version"
            With cn.Execute(strSQL)
                strRet = .fields(0).Value
            End With
        Else
            strRet = ""
        End If
    Else
        strRet = ""
    End If
    
    GetMSSQLServerVersion = strRet
    
    Exit Function
    
HandleError:

    'MsgBox Err.Description

    GetMSSQLServerVersion = ""

End Function


Private Function pvGetError(ByVal rc As Long, ByVal hEnv As Long, ByVal hDbc As Long, ByVal hStm As Long) As String
    
    Const FUNC_NAME     As String = "pvGetError"
    Dim sSqlState       As String * 5
    Dim lNativeError    As Long
    Dim sMsg            As String * 512
    Dim nTextLength     As Integer
    
    On Error GoTo HandleError
    
        SQLError hEnv, hDbc, hStm, sSqlState, lNativeError, sMsg, Len(sMsg), nTextLength
        
        pvGetError = "ODBC Result: 0x" & Hex(rc) & vbCrLf & vbCrLf & Left(sMsg, nTextLength)
    
    Exit Function

HandleError:
    
    'RaiseError FUNC_NAME

End Function


Public Function EnumSqlDbAdo( _
            sServer As String, _
            Optional sUser As String, _
            Optional sPass As String) As Variant
    
    Const FUNC_NAME     As String = "EnumSqlDbAdo"
    Dim cn              As Object
    Dim vRet            As Variant
    Dim lIdx            As Long
    Dim sSQL            As String
    
    On Error Resume Next
    
    Set cn = CreateObject("ADODB.CONNECTION")
    
    If Not (cn Is Nothing) Then
        cn.ConnectionTimeout = 3
        
        cn.CursorLocation = 3 'adUseClient
        
        For lIdx = 1 To 3
            If lIdx = 3 Then
                On Error GoTo HandleError
            Else
                Err.Clear
            End If
            If sUser <> "" Then
                cn.Open "Provider=SQLOLEDB;Persist Security Info=false;Data Source=" & sServer, sUser, sPass
            Else
                cn.Open "Provider=SQLOLEDB;Persist Security Info=false;Integrated Security=SSPI;Data Source=" & sServer
            End If
            If Err.Number = 0 Then
                sSQL = "select  name" & vbCrLf & _
                       "from    sysdatabases " & vbCrLf & _
                       "where   (status & 512) = 0" & vbCrLf & _
                       "        and charindex('|' + name + '|', " & vbCrLf & _
                       "            '|master|model|tempdb|msdb|distribution|') = 0"
                With cn.Execute(sSQL)
                    ReDim vRet(.RecordCount - 1)
                    lIdx = 0
                    Do While Not .EOF
                        vRet(lIdx) = !Name
                        lIdx = lIdx + 1
                        .MoveNext
                    Loop
                    If lIdx = 0 Then
                        ReDim vRet(0)
                        vRet(0) = STR_NO_USER_DBS
                    Else
                        ReDim Preserve vRet(lIdx - 1)
                    End If
                End With
                EnumSqlDbAdo = vRet
                Exit Function
            End If
        Next
        
        '--- on failure -> return Array(0 To -1)
        
        If Not IsArray(EnumSqlDbAdo) Then
            EnumSqlDbAdo = Split("")
        End If
        
    Else
        EnumSqlDbAdo = Split("")
    End If
        
    Exit Function

HandleError:
    
    'RaiseError FUNC_NAME

End Function

