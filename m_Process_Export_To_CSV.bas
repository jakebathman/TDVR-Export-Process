Attribute VB_Name = "m_Process_Export_To_CSV"
Option Explicit

'   MAIN SUB

Public Sub mProcessTDVRExportSheet()
    Dim i%, j%, c%
    Dim intLastCol
    Dim intLastRow
    Dim shtExport As Worksheet
    Dim shtSQL As Worksheet
    Dim boolDoneSorting As Boolean
    Dim intCol1%, intCol2%, intCol3%, intCol4%, intCol5%, intCol6%, intCol7%, intCol8%, intCol9%, intCol10%
    Dim t

    Application.ScreenUpdating = False

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

    t = Timer
    While Timer < t + 0.1
        DoEvents
    Wend


    Cells.EntireColumn.AutoFit

    boolDoneSorting = False

    ' re-order to match SQL database
    c = 0
    For j = 1 To 15
        Select Case UCase$(Cells(1, j).Value)
            Case Is = UCase$("Id"): intCol1 = j
            Case Is = UCase$("Last Updated"): intCol2 = j
            Case Is = UCase$("First Name"): intCol3 = j
            Case Is = UCase$("Last Name"): intCol4 = j
            Case Is = UCase$("Email"): intCol5 = j
            Case Is = UCase$("Email 2"): intCol6 = j
            Case Is = UCase$("Date Registered"): intCol7 = j
            Case Is = UCase$("Login Date"): intCol8 = j
            Case Is = UCase$("Profile % Complete"): intCol9 = j
            Case Is = UCase$("UserName"): intCol10 = j
            Case Is = vbNullString: Exit For
        End Select
        If c = 10 Then boolDoneSorting = True
    Next j
    
    'Columns("A:J").Insert shift:=xlToRight
    Call MoveCol(intCol1, 11)
    Call MoveCol(intCol2, 12)
    Call MoveCol(intCol3, 13)
    Call MoveCol(intCol4, 14)
    Call MoveCol(intCol5, 15)
    Call MoveCol(intCol6, 16)
    Call MoveCol(intCol7, 17)
    Call MoveCol(intCol8, 18)
    Call MoveCol(intCol9, 19)
    Call MoveCol(intCol10, 20)
    Application.CutCopyMode = False
    Columns("A:J").Delete
    Rows(1).Delete

    Application.ScreenUpdating = True

    t = Timer
    While Timer < t + 0.25
        DoEvents
    Wend


    ' Save as CSV
    Call ExportToCSV(ThisWorkbook.Name, shtSQL.Name, "CSVforSQL")

End Sub


Public Sub MoveCol(curCol As Integer, newCol As Integer)
    Columns(curCol).Copy
    Columns(newCol).Insert shift:=xlToRight
    Dim t
    t = Timer
    While Timer < t + 0.05
        DoEvents
    Wend
End Sub
