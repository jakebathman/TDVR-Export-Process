Attribute VB_Name = "z_btn"
Option Explicit

Public Sub btnProcessExportRibbonButton(control As IRibbonControl)
Attribute btnProcessExportRibbonButton.VB_ProcData.VB_Invoke_Func = " \n14"
    Call mProcessTDVRExportSheet
End Sub

' Using program Custom UI Editor for Microsoft Office
' More info: http://stackoverflow.com/questions/8850836/how-to-add-a-custom-ribbon-tab-using-vba
' Image Mso gallery: http://soltechs.net/CustomUI/imageMso01.asp
