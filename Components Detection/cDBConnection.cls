VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'*
'*  Ãû³Æ£ºcDBConnection
'*
'*  ¹¦ÄÜ£ºÊý¾Ý¿âÁ¬½Ó·â×°Àà
'*
'*  ×÷Õß£ºchs(wilson chan)
'*
'*  ´´½¨£º2003-04-14
'*
'*  ¸üÐÂ£º2003-04-16 10:29
'*
'*************************************************************************
Option Explicit

Private m_cnnDatabase As Object 'ADODB.Connection

Private m_strConnectionFile As String
Private m_strServer As String
Private m_strDatabase As String
Private m_strUserName As String
Private m_strPassword As String
Private m_blnIsAccessMDB As Boolean



Public Property Get ConnectionFile() As String
        
        ConnectionFile = m_strConnectionFile
        
End Property

Public Property Let ConnectionFile(ByVal strConnectionFile As String)
        
        m_strConnectionFile = strConnectionFile
        
End Property


Public Property Get Server() As String
        
        Server = m_strServer
        
End Property

Public Property Let Server(ByVal strServer As String)
        
        m_strServer = strServer
        
End Property


Public Property Get Database() As String
        
        Database = m_strDatabase
        
End Property

Public Property Let Database(ByVal strDatabase As String)
        
        m_strDatabase = strDatabase
        
End Property


Public Property Get UserName() As String
        
        UserName = m_strUserName
        
End Property

Public Property Let UserName(ByVal strUserName As String)
        
        m_strUserName = strUserName
        
End Property


Public Property Get Password() As String
        
        Password = m_strPassword
        
End Property

Public Property Let Password(ByVal strPassword As String)
        
        m_strPassword = strPassword
        
End Property


Public Property Get IsAccessMDB() As Boolean
        
        IsAccessMDB = m_blnIsAccessMDB
        
End Property

Public Property Let IsAccessMDB(ByVal blnIsAccessMDB As Boolean)
        
        m_blnIsAccessMDB = blnIsAccessMDB
        
End Property


Public Property Get DBConnection() As Object 'ADODB.Connection
        
        Set DBConnection = m_cnnDatabase
        
End Property


Private Function ReadConnectionInfo(ByVal strDBConnection As String) As Boolean
        
        Dim o_strFile As String
        Dim o_ucmFileFuncs As New cFileFuncs
        
        With o_ucmFileFuncs
            o_strFile = ConnectionFile()
            
            If .DoesFileExist(o_strFile) Then
                With New cINI
                    .Path = o_strFile
                    .Section = strDBConnection
                    
                    .Key = "SERVER"
                    Server() = .Value
                    
                    .Key = "DATABASE"
                    Database = .Value
                    
                    .Key = "USERNAME"
                    UserName = .Value
                    
                    .Key = "PASSWORD"
                    Password = .Value
                    
                    .Key = "ISAccessMDB"
                    IsAccessMDB = CBool(Val(.Value) = 1)
                End With
                
                ReadConnectionInfo = True
            Else
                ReadConnectionInfo = False
            End If
        End With
        
        Set o_ucmFileFuncs = Nothing
                
End Function


Public Function StartConnection(Optional ByVal blnUseConnectionFile As Boolean = True, _
                                Optional ByVal strDBConnection As String = "DBConnection") As Boolean
        
    On Error GoTo HandleError
        
        Dim o_blnRet As Boolean
        
        StopConnection
        
        Set m_cnnDatabase = CreateObject("ADODB.Connection") 'New ADODB.Connection
                
        With m_cnnDatabase
            
            If blnUseConnectionFile Then
                o_blnRet = ReadConnectionInfo(strDBConnection)
            Else
                o_blnRet = True
            End If
            
            If o_blnRet Then
                If IsAccessMDB() Then
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;User ID=" & UserName() & ";Password=" & Password() & ";Data Source=" & Database
                Else
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UserName() & ";Password=" & Password() & ";Initial Catalog=" & Database() & ";Data Source=" & Server()
                End If
                
                .Open
                
                o_blnRet = True
            Else
                o_blnRet = False
            End If
        End With
        
        StartConnection = o_blnRet
        
    Exit Function
        
HandleError:

    StartConnection = False

End Function


Public Function StopConnection() As Boolean
        
        If Not (m_cnnDatabase Is Nothing) Then
            With m_cnnDatabase
                If .State <> 0 Then 'adStateClosed Then
                    .Close
                Else
                End If
            End With
            
            Set m_cnnDatabase = Nothing
        Else
        End If
        
        StopConnection = True
        
End Function


Private Sub Class_Initialize()
        
        m_blnIsAccessMDB = False
        
End Sub


Private Sub Class_Terminate()
        
        StopConnection
        
End Sub

