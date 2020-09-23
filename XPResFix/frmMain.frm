VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supercalifragilistic XP Alidocious"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optProcess 
      Caption         =   "Add"
      Height          =   255
      Index           =   0
      Left            =   1245
      TabIndex        =   4
      Top             =   510
      Width           =   1215
   End
   Begin VB.OptionButton optProcess 
      Caption         =   "Remove"
      Height          =   255
      Index           =   1
      Left            =   1245
      TabIndex        =   3
      Top             =   870
      Width           =   1215
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Executables|*.exe"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Function"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.Image Image1 
         Height          =   435
         Left            =   240
         Picture         =   "frmMain.frx":628A
         Top             =   420
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" _
        (ByVal pFileName As String, _
        ByVal bDeleteExistingResources As Long) As Long

Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" _
        (ByVal hUpdate As Long, _
        ByVal lpType As Integer, _
        ByVal lpName As Integer, _
        ByVal wLanguage As Long, _
        lpData As Any, _
        ByVal cbData As Long) As Long

Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" _
        (ByVal hUpdate As Long, _
        ByVal fDiscard As Long) As Long

Private strManifest As String, strFile As String

Private Sub cmdProcess_Click()

    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    strFile = CommonDialog1.FileName
    
    If strFile = "" Then Exit Sub
    
    ModifyResource strManifest, Len(strManifest)

End Sub

Private Sub Form_Load()
    optProcess(0).Value = True
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub ModifyResource(lpData As String, ByVal cbData As Long)

    Dim hUpdateRes As Long, lRet As Long
    
    'get handle for UpdateResource. strFile must be an executable.
    hUpdateRes = BeginUpdateResource(strFile, False)

    If hUpdateRes = 0 Then GoTo FileError
    
    'modify the resource.
    lRet = UpdateResource(hUpdateRes, 24, 1, 1033, ByVal lpData, cbData)
    
    'commit the changes to the executable file.
    lRet = EndUpdateResource(hUpdateRes, False)

    MsgBox "File Updated!"
    
Exit Sub

FileError:
        MsgBox "Could not open file for writing!"
        
End Sub


Private Sub optProcess_Click(Index As Integer)

    Select Case Index
        Case 0
            strManifest = LoadResString(101)
        Case Else
            strManifest = ""
    End Select

End Sub
