Attribute VB_Name = "XPInit"
Private Declare Function InitCommonControls Lib "COMCTL32" () As Long

Sub Main()
    InitCommonControls
    frmMain.Show
End Sub
