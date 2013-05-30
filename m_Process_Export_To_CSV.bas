Attribute VB_Name = "m_Process_Export_To_CSV"
Option Explicit

Public Sub mProcessTDVRExportSheet()
    Dim i%, j%
    Dim intLastCol
    Dim intLastRow
    Dim shtExport As Worksheet
    Dim shtSQL As Worksheet

    ' Duplicate sheet (preserving original sheet)
    On Error Resume Next
    Set shtSQL = ThisWorkbook.Sheets("SQL")
    'Debug.Print Err.Number
    'Debug.Print Err.Description
    Application.DisplayAlerts = False
    If Err.Number = 0 Then
        If MsgBox("Overwrite existing SQL sheet?", vbOKCancel) <> vbOK Then Exit Sub
        shtSQL.Delete
    End If
    Set shtExport = ActiveSheet
    shtExport.Copy after:=shtExport
    Set shtSQL = ActiveSheet
    shtSQL.Name = "SQL"
    Err.Number = 0
    Application.DisplayAlerts = True
    'Debug.Print Err.Number
    On Error GoTo 0

    For j = 1 To 300
        If Cells(1, j).Value = vbNullString Then intLastCol = j - 1: Exit For
    Next j

    For j = 1 To intLastCol
        Select Case UCase$(Cells(1, j).Value)
            Case Is = UCase$("Date Registered")
            Case Is = UCase$("Email")
            Case Is = UCase$("Email 2")
            Case Is = UCase$("First Name")
            Case Is = UCase$("Gender")
            Case Is = UCase$("Id")
            Case Is = UCase$("Last Name")
            Case Is = UCase$("Last Updated")
            Case Is = UCase$("Login Date")
            Case Is = UCase$("Profile % Complete")
            Case Is = UCase$("UserName")
            Case Is = vbNullString: Exit For
            Case Else
                Columns(j).Delete
                j = j - 1
        End Select
    Next j

Cells.EntireColumn.AutoFit

End Sub
