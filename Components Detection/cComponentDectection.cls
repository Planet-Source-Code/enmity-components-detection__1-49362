VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComponentDectection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum DotNETFrameworkVersions
    dnfvV1
    dnfvV1_1
    dnfvVAny
End Enum



Public Function DoesDotNETFrameworkExist(ByVal udeVersion As DotNETFrameworkVersions) As Boolean
        
        Dim o_blnRet As Boolean
        Dim o_strRet As String
        
        With New cRegistry
            .hKey = HKEY_LOCAL_MACHINE
            .KeyPath = "Software\Microsoft\.NETFramework"
            If .DoesKeyExist() Then
                .KeyPath = "Software\Microsoft\.NETFramework"
                o_strRet = GetRegValue(.GetRegistryValue("InstallRoot", ""))
                o_blnRet = (o_strRet <> "")
                If o_blnRet Then
                    With New cFileFuncs
                        Select Case udeVersion
                            Case dnfvV1
                                o_blnRet = .DoesFileExistEx(o_strRet & "v1.0.3705\mscorlib.dll")
                            Case dnfvV1_1
                                o_blnRet = .DoesFileExistEx(o_strRet & "v1.1.4322\mscorlib.dll")
                            Case dnfvVAny
                                o_blnRet = .DoesFileExistEx(o_strRet & "v1.0.3705\mscorlib.dll")
                                If o_blnRet Then
                                Else
                                    o_blnRet = .DoesFileExistEx(o_strRet & "v1.1.4322\mscorlib.dll")
                                End If
                        End Select
                    End With
                Else
                
                End If
            Else
                o_blnRet = False
            End If
        End With
        
        DoesDotNETFrameworkExist = o_blnRet
        
End Function


Private Function GetRegValue(ByVal strValue As String) As String
    
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


Public Function DoesIISExist() As Boolean

        Dim objIIS As Object                    'ADSI IIS Object
        
        Err.Clear
        
        On Error Resume Next
        
        Set objIIS = GetObject("IIS://localhost")
        
        If Err.Number <> 0 Then
            DoesIISExist = False 'MsgBox "IIS ²»´æÔÚ"
        Else
            DoesIISExist = True 'MsgBox "IIS ´æÔÚ"
        End If

        Set objIIS = Nothing
        
End Function


Public Function DoesIE6Exist() As Boolean
            
        With New cRegistry
            .hKey = HKEY_LOCAL_MACHINE
            .KeyPath = "SOFTWARE\Microsoft\Internet Explorer"
            DoesIE6Exist = (Left(.GetRegistryValue("Version"), 10) = "REG_SZ : " & "6")
        End With
            
End Function

'
'Public Function DoesSQLServer2kSP3Exist() As Boolean
'
'        With New cSQLServer
'            'DoesSQLServer2kExist = .DoesSQLServer2kSP3Exist()
'        End With
'
'End Function
'

Public Function DoesMSSQLServerSP3Exist(ByVal strServer As String, _
                                        ByVal strUserName As String, _
                                        ByVal strPassword As String) As Boolean

    Dim strRet As String
    
    With New cSQLServer
        strRet = .GetMSSQLServerVersion(strServer, strUserName, strPassword)
    End With
    
    DoesMSSQLServerSP3Exist = InStr(1, strRet, "Microsoft SQL Server  2000 - 8.00.760", vbTextCompare) > 0

End Function


Public Function DoesSQLServer2kExistEx(ByVal strServer As String, _
                                       ByVal strDatabase As String, _
                                       ByVal strUser As String, _
                                       ByVal strPassword As String) As Boolean

        With New cSQLServer
            DoesSQLServer2kExistEx = .DoesSQLServer2000ExistEx(strServer, strDatabase, strUser, strPassword)
        End With

End Function


Public Function DoesSQLServer2kExist() As Boolean

        With New cSQLServer
            DoesSQLServer2kExist = .DoesSQLServer2000Exist()
        End With

End Function


Public Function DoesWindows2kSP2Exist() As Boolean
            
        With New cSystemInfoEx
            DoesWindows2kSP2Exist = (.ServicePackMajorVersion = 2)
        End With
            
End Function


Public Function DoesWindows2kSP3Exist() As Boolean
            
        With New cSystemInfoEx
            DoesWindows2kSP3Exist = (.ServicePackMajorVersion = 3)
        End With
            
End Function


Public Function DoesWindows2kSP4Exist() As Boolean
            
        With New cSystemInfoEx
            DoesWindows2kSP4Exist = (.ServicePackMajorVersion = 4)
        End With
            
End Function


Public Function DoesMDAC27Exist() As Boolean
        
        Dim o_objCNN As Object
        
        On Error GoTo HandleError
            
            Set o_objCNN = CreateObject("ADODB.Connection")
            
            DoesMDAC27Exist = (o_objCNN.Version = "2.7")
        
            Set o_objCNN = Nothing
        
        Exit Function

HandleError:
    
    'MsgBox Err.Description
    
    DoesMDAC27Exist = False
        
End Function


Public Function DoesWindowsInstaller2Exist() As Boolean
        
        Dim o_objWI As Object
        
        On Error GoTo HandleError
        
            Set o_objWI = CreateObject("WindowsInstaller.Installer")
            
            DoesWindowsInstaller2Exist = CBool(Val(Left(o_objWI.Version, 1)) = 2)
            
            Set o_objWI = Nothing
        
        Exit Function
        
HandleError:
    
    'MsgBox Err.Description
    
    DoesWindowsInstaller2Exist = False
                
End Function


