VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileFuncs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH As Long = 260

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Const INVALID_HANDLE_VALUE = -1

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Private Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4

'SHFileOperation declarations
Public Enum FO_Operation
    FO_MOVE = 1
    FO_COPY = 2
    FO_DELETE = 3
    FO_RENAME = 4
End Enum

Public Enum FOFlags
    FOF_MULTIDESTFILES = &H1      'Destination specifies multiple files
    FOF_SILENT = &H4              'Don't display progress dialog
    FOF_RENAMEONCOLLISION = &H8   'Rename if destination already exists
    FOF_NOCONFIRMATION = &H10     'Don't prompt user
    FOF_WANTMAPPINGHANDLE = &H20  'Fill in hNameMappings member
    FOF_ALLOWUNDO = &H40          'Store undo information if possible
    FOF_FILESONLY = &H80          'On *.*, don't copy directories
    FOF_SIMPLEPROGRESS = &H100    'Don't show name of each file
    FOF_NOCONFIRMMKDIR = &H200    'Don't confirm making any needed dirs
End Enum

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private m_opFileOperation As SHFILEOPSTRUCT

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Const TOKEN_QUERY = (&H8)
Private Declare Function GetAllUsersProfileDirectory Lib "userenv.dll" Alias "GetAllUsersProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetDefaultUserProfileDirectory Lib "userenv.dll" Alias "GetDefaultUserProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetProfilesDirectory Lib "userenv.dll" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

Private m_strLastErrorDescription As String
Private m_lngLastErrorNumber As Long



Public Property Get LastErrorDescription() As String
    
        LastErrorDescription = m_strLastErrorDescription
    
End Property

Public Property Let LastErrorDescription(ByVal strLastErrorDescription As String)
    
        m_strLastErrorDescription = strLastErrorDescription
    
End Property


Public Property Get LastErrorNumber() As Long
    
        LastErrorNumber = m_lngLastErrorNumber
    
End Property

Public Property Let LastErrorNumber(ByVal lngLastErrorNumber As Long)
    
        m_lngLastErrorNumber = lngLastErrorNumber
    
End Property


Public Function DeleteFilesAndFolders(strDeleteFileFolder As String, _
                                      ByRef strRet As String, _
                                      Optional blnInterface As Boolean = False) As Boolean

    On Error GoTo HandleError

        With m_opFileOperation
            .wFunc = FO_DELETE
            .pFrom = strDeleteFileFolder
            .fFlags = IIf(blnInterface = False, FOF_NOCONFIRMATION, FOF_NOCONFIRMATION And FOF_SILENT)
        End With
        
        SHFileOperation m_opFileOperation
        
        strRet = ""
        
        DeleteFilesAndFolders = True
    
    Exit Function
    
HandleError:
    
    strRet = Err.Description
    
    DeleteFilesAndFolders = False

End Function


Public Function GetToken(ByVal strText As String) As String
        
  Dim intPos As Integer
  
  On Error GoTo PROC_ERR

  intPos = InStr(strText, vbNullChar)
  
  If intPos = 0 Then
    ' No nulls in the strTextg, just return it as is
    GetToken = strText
  Else
    If intPos = 1 Then
      ' If the null character is at the first position, the
      ' entire string is a null string, so return a zero-length string
      GetToken = ""
    Else
      ' Not at the first position, so return the contents up
      ' to the occurrence of the null character
      GetToken = VBA.Left$(strText, intPos - 1)
    End If
  End If
    
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Resume PROC_EXIT
        
End Function


Public Property Get AllUsersProfileDirectory() As String
        
        Dim o_strRet As String
        
        o_strRet = Space(256)
        
        GetAllUsersProfileDirectory o_strRet, Len(o_strRet)
        
        AllUsersProfileDirectory = GetToken(o_strRet)
        
End Property


Public Property Get DefaultUserProfileDirectory() As String
        
        Dim o_strRet As String
        
        o_strRet = Space(256)
        
        GetDefaultUserProfileDirectory o_strRet, Len(o_strRet)
        
        DefaultUserProfileDirectory = GetToken(o_strRet)
        
End Property


Public Property Get ProfilesDirectory() As String
        
        Dim o_strRet As String
        
        o_strRet = Space(256)
        
        GetProfilesDirectory o_strRet, Len(o_strRet)
        
        ProfilesDirectory = GetToken(o_strRet)
        
End Property


Public Property Get UserProfileDirectory() As String
        
        Dim o_strRet As String
        Dim hToken As Long
        
        o_strRet = Space(256)
        
        OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
        GetUserProfileDirectory hToken, o_strRet, Len(o_strRet)
        
        UserProfileDirectory = GetToken(o_strRet)
        
End Property


Public Function GetWindowsPath2() As String
        
        Dim o_strRet As String
        
        o_strRet = Space(256)
        
        GetWindowsDirectory o_strRet, Len(o_strRet)
        
        'GetWindowsPath2 = Left(o_strRet, InStr(1, o_strRet, vbNullChar, vbTextCompare) - 1)
        GetWindowsPath2 = GetToken(o_strRet)
        
End Function


Public Function GetWindowsSystemPath() As String
        
        Dim o_strRet As String
        
        o_strRet = Space(256)
        
        GetSystemDirectory o_strRet, Len(o_strRet)
        
        'GetWindowsSystemPath = Left(o_strRet, InStr(1, o_strRet, vbNullChar, vbTextCompare) - 1)
        GetWindowsSystemPath = GetToken(o_strRet)
        
End Function


Public Function GetFileName(ByVal strFilePath As String) As String
        
        GetFileName = Right(strFilePath, Len(strFilePath) - InStrRev(strFilePath, "\"))
        
End Function


Public Function GetWindowsPath() As String
        
        Dim o_strRet As String
        
        o_strRet = Space(256)
        
        GetWindowsDirectory o_strRet, Len(o_strRet)
        
        'GetWindowsPath = Left(o_strRet, InStr(1, o_strRet, vbNullChar, vbTextCompare) - 1)
        GetWindowsPath = GetToken(o_strRet)
        
End Function


Public Function GetWindowsTempPath() As String
        
        Dim o_strRet As String
        
        o_strRet = Space(256)
        
        GetTempPath Len(o_strRet), o_strRet
        
        GetWindowsTempPath = Left(o_strRet, InStr(1, o_strRet, vbNullChar, vbTextCompare) - 1)
        
        GetWindowsTempPath = GetToken(o_strRet)
        
End Function


Public Sub ShellOpenFile(ByVal strFile As String)
        
        ShellExecute 0, "Open", strFile, vbNullString, vbNullString, vbNormalFocus
        
End Sub


Public Sub KillSelf(ByVal strFile As String)

        MoveFileEx strFile, vbNullString, MOVEFILE_DELAY_UNTIL_REBOOT

End Sub


Public Function IsCDROMDrive(ByVal strFilePath As String) As Boolean
        
'        Dim o_strDrives As String
'        Dim o_strItems() As String
'        Dim o_lngRet As Long
        Dim o_strDrive As String
'
'        o_strDrives = Space(256)
'
'        o_lngRet = GetLogicalDriveStrings(Len(o_strDrives), o_strDrives)
'
'        o_strDrives = Left(o_strDrives, InStr(1, o_strDrives, vbNullChar) - 1)
'
'        Debug.Print o_strDrives
'
'        o_strItems() = Split(o_strDrives, vbNullChar)

        If Len(strFilePath) > 2 Then
            If InStr(1, strFilePath, ":\") = 2 Then
                o_strDrive = Left(strFilePath, 3)
            Else
                o_strDrive = ""
            End If
        Else
            o_strDrive = ""
        End If
        
        If Len(o_strDrive) > 0 Then
            IsCDROMDrive = (GetDriveType(o_strDrive) = DRIVE_CDROM)
        Else
            IsCDROMDrive = False
        End If
        
End Function


Public Function GetCDROMDrive() As String
        
        Dim o_strDrives As String
        Dim o_strItems() As String
        Dim o_lngRet As Long
        Dim o_strDrive As Variant

        o_strDrives = Space(256)

        o_lngRet = GetLogicalDriveStrings(Len(o_strDrives), o_strDrives)

        o_strDrives = Left(o_strDrives, InStrRev(o_strDrives, "\"))

        Debug.Print o_strDrives

        o_strItems() = Split(o_strDrives, vbNullChar)

        For Each o_strDrive In o_strItems
            If (GetDriveType(o_strDrive) = DRIVE_CDROM) Then
                GetCDROMDrive = o_strDrive
                
                Exit Function
            Else
            End If
        Next
        
End Function


Public Function GetHardDiskDrives() As String()
        
        Dim o_strDrives As String
        Dim o_strItems() As String
        Dim o_strDriveItems() As String
        Dim o_lngRet As Long
        Dim o_strDrive As Variant
        Dim o_lngDrives As Long

        o_strDrives = Space(256)

        o_lngRet = GetLogicalDriveStrings(Len(o_strDrives), o_strDrives)

        o_strDrives = Left(o_strDrives, InStrRev(o_strDrives, "\"))

        Debug.Print o_strDrives

        o_strItems() = Split(o_strDrives, vbNullChar)

        For Each o_strDrive In o_strItems
            If (GetDriveType(o_strDrive) = DRIVE_FIXED) Then
                o_lngDrives = o_lngDrives + 1
                ReDim Preserve o_strDriveItems(0 To o_lngDrives - 1)
                o_strDriveItems(o_lngDrives - 1) = o_strDrive
            Else
            End If
        Next
        
        GetHardDiskDrives = o_strDriveItems
        
End Function


Public Sub SafeKill(ByVal strFile As String)
        
        On Error Resume Next
        
        If DoesFileExist(strFile) Then
            SetAttr strFile, vbNormal
            Kill strFile
        Else
        End If
        
End Sub


Public Function GetDriveName(ByVal strPath As String) As String
        
        If Len(strPath) >= 2 Then
            GetDriveName = Left(strPath, 2)
        Else
            GetDriveName = ""
        End If

End Function


Public Function GetFilePath(ByVal strFile As String) As String
        
        GetFilePath = Left(strFile, InStrRev(strFile, "\") - 1)
        
End Function


Public Function FileCopyEx(ByVal strSrcFile As String, ByVal strDestFile As String) As Boolean
            
        On Error GoTo HandleError
            
            FileCopy strSrcFile, strDestFile
            
            FileCopyEx = True
        
        Exit Function
        
HandleError:
    
    LastErrorDescription = Err.Description
    LastErrorNumber = Err.Number
    
    FileCopyEx = False
            
End Function


Public Function GetParentFolder(ByVal strFolder As String) As String
        
        Dim o_strRet As String
        
        If Len(Trim(strFolder)) Then
            If Right(strFolder, 1) = "\" Then
                o_strRet = Left(strFolder, Len(strFolder) - 1)
            Else
                o_strRet = strFolder
            End If
            
            o_strRet = Left(o_strRet, InStrRev(o_strRet, "\") - 1)
            
        Else
            o_strRet = vbNullString
        End If
        
        GetParentFolder = o_strRet
        
End Function


Public Function LoadTextContent(ByVal strFile As String) As String
        
        Dim o_intFileNum As Integer
        Dim o_strRet As String
        
        o_intFileNum = FreeFile()
        
        Open strFile For Binary As #o_intFileNum
        
        o_strRet = Space(LOF(o_intFileNum))
        
        Get #o_intFileNum, , o_strRet
        
        Close #o_intFileNum
        
        LoadTextContent = o_strRet
        
End Function


Public Function SaveTextToFile(ByVal strFile As String, _
                               ByVal strText As String, _
                               Optional ByVal blnOverwrite As Boolean = False) As Boolean
        
        Dim o_intFileNum As Integer
        
        o_intFileNum = FreeFile()
        
        If blnOverwrite Then
            Open strFile For Output As #o_intFileNum
        Else
            Open strFile For Append As #o_intFileNum
        End If
        
        Print #o_intFileNum, strText
        
        Close #o_intFileNum
        
        SaveTextToFile = True
        
End Function


Public Function GetNormalPath(ByVal strPath As String) As String

        GetNormalPath = IIf(Right(strPath, 1) = "\", Left(strPath, Len(strPath) - 1), strPath)
                
End Function


Public Function GetAppPath() As String

        GetAppPath = IIf(Right(App.Path, 1) = "\", Left(App.Path, Len(App.Path) - 1), App.Path)
                
End Function



Public Function DoesFileExist(ByVal sFIle As String) As Boolean
    
    Dim FindData As WIN32_FIND_DATA
    Dim GetFile As Long
       
    GetFile = FindFirstFile(sFIle, FindData)
        
    If GetFile <> INVALID_HANDLE_VALUE Then
        If (FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            DoesFileExist = False 'Ä¿Â¼
        Else
            DoesFileExist = True
        End If
    Else
        DoesFileExist = False
    End If
    
    FindClose GetFile
 
End Function


Public Function DoesFolderExist(ByVal sFIle As String) As Boolean
    
    Dim FindData As WIN32_FIND_DATA
    Dim GetFile As Long
       
    GetFile = FindFirstFile(sFIle, FindData)
        
    If GetFile <> INVALID_HANDLE_VALUE Then
        If (FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            DoesFolderExist = True 'Ä¿Â¼
        Else
            DoesFolderExist = False
        End If
    Else
        DoesFolderExist = False
    End If
    
    FindClose GetFile
 
End Function


Public Function DoesFileExistEx(ByVal strFile As String) As Boolean
        
        Dim o_intFileNum As Integer
        
        On Error GoTo HandleError
        
            o_intFileNum = FreeFile()
            
            Open strFile For Input As #o_intFileNum
            
            Close #o_intFileNum
            
            DoesFileExistEx = True
        
        Exit Function
        
HandleError:

    DoesFileExistEx = False
        
End Function


Public Function IsCDROMReady(ByVal strDrive As String) As Boolean
        
        On Error GoTo HandleError
            
            IsCDROMReady = CBool(Dir(strDrive & "*.*") <> "")
        
        Exit Function
        
HandleError:
    
    IsCDROMReady = False

End Function
