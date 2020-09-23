VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Private variables to store the setting
'     s made:
Private m_strPath As String
Private m_strKey As String
Private m_strSection As String
Private m_strDefault As String
Private m_lngLastReturnCode As Long
Private m_lngCacheSize As Long

' Declares for cIniFile:
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



Property Get CacheSize() As Long
    CacheSize = m_lngCacheSize
End Property

Property Let CacheSize(ByVal lngSize As Long)
    m_lngCacheSize = lngSize
End Property


Property Get LastReturnCode() As Long
    ' Did the last call succeed?
    ' 0 if not!
    LastReturnCode = m_lngLastReturnCode
End Property


Property Let Default(sDefault As String)
    ' What to return if something goes wrong
    '     :
    m_strDefault = sDefault
End Property


Property Get Default() As String
    ' What to return if something goes wrong
    '     :
    Default = m_strDefault
End Property


Property Let Path(sPath As String)
    ' The filename of the INI file:
    m_strPath = sPath
End Property


Property Get Path() As String
    ' The filename of the INI file:
    Path = m_strPath
End Property


Property Let Key(sKey As String)
    ' The KEY= bit to look for
    m_strKey = sKey
End Property


Property Get Key() As String
    ' The KEY= bit to look for
    Key = m_strKey
End Property


Property Let Section(sSection As String)
    ' The [SECTION] bit to look for
    m_strSection = sSection
End Property


Property Get Section() As String
    ' The [SECTION] bit to look for
    Section = m_strSection
End Property


Property Get Value() As String
    ' Get the value of the current Key within Section of Path
    Dim o_strBuf As String
    Dim o_lngSize As Long
    Dim o_lngRet As Long
    
    o_strBuf = Space$(m_lngCacheSize)
    o_lngSize = Len(o_strBuf)
    o_lngRet = GetPrivateProfileString(m_strSection, m_strKey, m_strDefault, o_strBuf, o_lngSize, m_strPath)

    If (o_lngSize > 0) Then
        If InStr(o_strBuf, vbNullChar) > 0 Then
            Value = Left(o_strBuf, InStr(o_strBuf, vbNullChar) - 1) 'Left$(o_strBuf, o_lngRet)
        Else
            Value = ""
        End If
    Else
        Value = ""
    End If

End Property


Property Let Value(sValue As String)
    ' Set the value of the current Key withi
    '     n Section of Path
    Dim o_intPos As Integer
    
    ' Strip chr$(0):
    o_intPos = InStr(sValue, Chr$(0))

    Do While o_intPos <> 0
        sValue = Left$(sValue, (o_intPos - 1)) & Mid$(sValue, (o_intPos + 1))
        o_intPos = InStr(sValue, Chr$(0))
    Loop
    m_lngLastReturnCode = WritePrivateProfileString(m_strSection, m_strKey, sValue, m_strPath)
End Property


Public Sub DeleteValue()
    ' Delete the value at Key within Section
    '     of Path
    m_lngLastReturnCode = WritePrivateProfileString(m_strSection, m_strKey, 0&, m_strPath)
End Sub


Public Sub DeleteSection()
    ' Delete the Section in Path
    m_lngLastReturnCode = WritePrivateProfileString(m_strSection, 0&, 0&, m_strPath)
End Sub


Property Get INISection() As String
    ' Return all the keys and values within
    '     the current
    ' section, separated by chr$(0):
    Dim o_strBuf As String
    Dim o_lngSize As Long
    Dim o_lngRet As Long
    
    o_strBuf = Space$(m_lngCacheSize)
    o_lngSize = Len(o_strBuf)
    o_lngRet = GetPrivateProfileString(m_strSection, 0&, m_strDefault, o_strBuf, o_lngSize, m_strPath)

    If (o_lngSize > 0) Then
        INISection = Left$(o_strBuf, o_lngRet)
    Else
        INISection = ""
    End If
End Property


Property Let INISection(sSection As String)
    ' Set one or more the keys within the cu
    '     rrent section.
    ' Keys and Values should be separated by
    '     chr$(0):
    m_lngLastReturnCode = WritePrivateProfileString(m_strSection, 0&, sSection, m_strPath)

End Property

Private Sub Class_Initialize()
    m_lngCacheSize = 1024
End Sub
