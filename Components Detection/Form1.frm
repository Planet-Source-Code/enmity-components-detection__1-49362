VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   5715
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "&Detect"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
        
        With New cComponentDectection
            Debug.Print "DoesDotNETFrameworkExist:" & .DoesDotNETFrameworkExist(dnfvVAny)
            Debug.Print "DoesIE6Exist" & .DoesIE6Exist
            Debug.Print "DoesIISExist" & .DoesIISExist
            Debug.Print "DoesMDAC27Exist" & .DoesMDAC27Exist
            Debug.Print "DoesMSSQLServerSP3Exist:" & .DoesMSSQLServerSP3Exist("(local)", "sa", "")
            Debug.Print "DoesSQLServer2kExist:" & .DoesSQLServer2kExist
            Debug.Print "DoesSQLServer2kExistEx:" & .DoesSQLServer2kExistEx("(local)", "master", "sa", "")
            Debug.Print "DoesWindows2kSP2Exist:" & .DoesWindows2kSP2Exist
            Debug.Print "DoesWindows2kSP3Exist:" & .DoesWindows2kSP3Exist
            Debug.Print "DoesWindows2kSP4Exist:" & .DoesWindows2kSP4Exist
            Debug.Print "DoesWindowsInstaller2Exist:" & .DoesWindowsInstaller2Exist
        End With
        
        MsgBox "done!"
        
End Sub
