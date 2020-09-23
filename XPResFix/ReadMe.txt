Unfortunately, VB6's Resource compiler does not natively support XML type resource files. I have made a res file with a manifest and have included it in the Test Project. I do not know why the project will not run when this resource is added. I imagine that VB compiles it in the wrong format, but that is just a guess (maybe it's at the wrong address).

As a workaround I have made this small project which will insert a manifest file into your VB executable.

To use WindowsXP Visual Styles you must call the InitCommonControls() function as the program Initialises.

i.e.

'********************************************************************
Private Declare Function InitCommonControls Lib "COMCTL32" () As Long

Private Sub Form_Initialize()
    InitCommonControls
End Sub
'*********************************************************************

I have found that putting controls inside frames causes a Black border to be shown around the control when using Visual Styles. This is why I have put command buttons and Option boxes on top of the framecontrol.

This project will add a manifest resource to any exe file and if the file becomes unusable can also remove it. Occasionally a file may become unusable even after removing the resource so always make a backup of any exe's before you modify them ;-)