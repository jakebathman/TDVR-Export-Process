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
    Dim intCol1%, intCol2%, intCol3%, intCol4%, intCol5%, intCol6%, intCol7%, intCol8%, intCol9%, intCol10%, intCol11%
    Dim t, w, ss, wb, sh

    Application.ScreenUpdating = False

    Set wb = ThisWorkbook

    For Each sh In wb.Sheets
        If sh.Name = "Old TDVR Export" Then
            Application.DisplayAlerts = False
            sh.Delete
            Application.DisplayAlerts = True
        End If
    Next sh




    'Get export sheet if it looks like it's open
    For Each w In Application.Workbooks
        If InStr(1, w.FullName, "Downloads\responder", vbTextCompare) > 0 Then
            Set ss = w.Sheets(1)
            ss.Copy Before:=wb.Sheets(1)
            t = Timer
            While Timer < t + 1
                DoEvents
            Wend
            w.Close savechanges:=False
        End If
    Next w

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
            Case Is = UCase$("Training Courses")
            Case Is = vbNullString: Exit For
            Case Else
                Columns(j).Delete
                j = j - 1
        End Select
    Next j

    For i = 1 To 5000
        If Cells(i, 1).Value = vbNullString Then
            intLastRow = i - 1
            Exit For
        End If
    Next i

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
            Case Is = UCase$("Id"): intCol1 = j: Cells(1, j).Value = "Id"
            Case Is = UCase$("Last Updated"): intCol2 = j: Cells(1, j).Value = "LastUpdated"
            Case Is = UCase$("First Name"): intCol3 = j: Cells(1, j).Value = "FirstName"
            Case Is = UCase$("Last Name"): intCol4 = j: Cells(1, j).Value = "LastName"
            Case Is = UCase$("Email"): intCol5 = j: Cells(1, j).Value = "Email"
            Case Is = UCase$("Email 2"): intCol6 = j: Cells(1, j).Value = "Email2"
            Case Is = UCase$("Date Registered"): intCol7 = j: Cells(1, j).Value = "DateRegistered"
            Case Is = UCase$("Login Date"): intCol8 = j: Cells(1, j).Value = "LoginDate"
            Case Is = UCase$("Profile % Complete"): intCol9 = j: Cells(1, j).Value = "ProfileComplete"
            Case Is = UCase$("UserName"): intCol10 = j: Cells(1, j).Value = "Username"
            Case Is = UCase$("Training Courses"): intCol11 = j: Cells(1, j).Value = "TrainingCourses"
            Case Is = vbNullString: Exit For
        End Select
        If c = 10 Then boolDoneSorting = True
    Next j

    'Columns("A:J").Insert shift:=xlToRight
    Call MoveCol(intCol1, 12)
    Call MoveCol(intCol2, 13)
    Call MoveCol(intCol3, 14)
    Call MoveCol(intCol4, 15)
    Call MoveCol(intCol5, 16)
    Call MoveCol(intCol6, 17)
    Call MoveCol(intCol7, 18)
    Call MoveCol(intCol8, 19)
    Call MoveCol(intCol9, 20)
    Call MoveCol(intCol10, 21)
    Call MoveCol(intCol11, 22)
    Application.CutCopyMode = False
    Columns("A:K").Delete
    'Rows(1).Delete

    Application.ScreenUpdating = True

    t = Timer
    While Timer < t + 0.25
        DoEvents
    Wend

    'Level 1
    Cells(1, 12).Value = "bool100Percent"
    Cells(1, 13).Value = "boolOrientation"
    Cells(1, 14).Value = "boolICS100"
    Cells(1, 15).Value = "boolICS700"

    'Level 2
    Cells(1, 16).Value = "boolCPR"
    Cells(1, 17).Value = "boolPDP"
    Cells(1, 18).Value = "boolDeployment101"
    Cells(1, 19).Value = "boolICS200"
    Cells(1, 20).Value = "boolICS800"

    Cells(1, 21).Value = "boolLevel1"
    Cells(1, 22).Value = "boolLevel2"

    Cells(1, 34).Value = "boolCloseToLevel1"
    Cells(1, 35).Value = "boolCloseToLevel2"
    Cells(1, 36).Value = "boolCloseToLevel3"

    Dim b100Perc As Boolean, bOrient As Boolean, b100 As Boolean, b700 As Boolean
    Dim bLevel1 As Boolean, bLevel2 As Boolean
    Dim bCPR As Boolean, bPDP As Boolean, bDep101 As Boolean, b200 As Boolean, b800 As Boolean
    Dim bClose1 As Boolean, bClose2 As Boolean, bClose3 As Boolean

    Dim tmpTraining$

    For i = 2 To 1000
        If Cells(i, 1).Value = vbNullString Then Exit For
        b100Perc = False: bOrient = False: b100 = False: b700 = False
        bLevel1 = False: bLevel2 = False
        bCPR = False: bPDP = False: bDep101 = False: b200 = False: b800 = False

        tmpTraining = Cells(i, 11).Value

        If Cells(i, 9).Value = "100" Then b100Perc = True: Cells(i, 12).Value = 1 Else Cells(i, 12).Value = 0
        If InStr(1, tmpTraining, "Orientation", vbTextCompare) > 0 Then bOrient = True: Cells(i, 13).Value = 1 Else Cells(i, 13).Value = 0
        If InStr(1, tmpTraining, "ICS-100", vbTextCompare) > 0 Then b100 = True: Cells(i, 14).Value = 1 Else Cells(i, 14).Value = 0
        If InStr(1, tmpTraining, "ICS-700", vbTextCompare) > 0 Then b700 = True: Cells(i, 15).Value = 1 Else Cells(i, 15).Value = 0

        If b100Perc And bOrient And b100 And b700 Then
            bLevel1 = True
            Cells(i, 21).Value = 1
            Cells(i, 34).Value = 0
        Else
            Cells(i, 21).Value = 0
            If Cells(i, 12).Value + Cells(i, 13).Value + Cells(i, 14).Value + Cells(i, 15).Value = 3 Then
                Cells(i, 34).Value = 1
            Else
                Cells(i, 34).Value = 0
            End If
        End If

        If (InStr(1, tmpTraining, "Cardio", vbTextCompare) > 0 Or InStr(1, tmpTraining, "CPR", vbTextCompare) > 0) Then bCPR = True: Cells(i, 16).Value = 1 Else Cells(i, 16).Value = 0
        If InStr(1, tmpTraining, "Personal Disaster", vbTextCompare) > 0 Then bPDP = True: Cells(i, 17).Value = 1 Else Cells(i, 17).Value = 0
        If InStr(1, tmpTraining, "Deployment", vbTextCompare) > 0 Then bDep101 = True: Cells(i, 18).Value = 1 Else Cells(i, 18).Value = 0
        If InStr(1, tmpTraining, "ICS-200", vbTextCompare) > 0 Then b200 = True: Cells(i, 19).Value = 1 Else Cells(i, 19).Value = 0
        If InStr(1, tmpTraining, "ICS-800", vbTextCompare) > 0 Then b800 = True: Cells(i, 20).Value = 1 Else Cells(i, 20).Value = 0

        If bCPR And bPDP And bDep101 And b200 And b800 And bLevel1 Then
            bLevel2 = True
            Cells(i, 22).Value = 1
        Else
            Cells(i, 22).Value = 0
        End If

        ' Close to Level 2 (but already Level 1)
        If (Cells(i, 16).Value + Cells(i, 17).Value + Cells(i, 18).Value + Cells(i, 19).Value + Cells(i, 20).Value = 4) And bLevel1 Then
            Cells(i, 35).Value = 1
        Else
            Cells(i, 35).Value = 0
        End If


        ' Close to Level 3
        Cells(i, 36).Value = 0

    Next i


    ' Additional columns that are updated elsewhere
    Cells(1, 23).Value = "boolLevel3"
    Cells(1, 24).Value = "boolHasBadge"
    Cells(1, 25).Value = "boolHasBlackBag"
    Cells(1, 26).Value = "boolHasShirt"
    Cells(1, 27).Value = "boolHasHat"
    Cells(1, 28).Value = "boolAttendedCeremony"
    Cells(1, 29).Value = "boolBackgroundCheck"
    Cells(1, 30).Value = "boolICS300"
    Cells(1, 31).Value = "boolICS400"
    Cells(1, 32).Value = "boolPFA"
    Cells(1, 33).Value = "boolPODLeadership"
    
    ' Put current timestamp in last column so we know how fresh this data is
    Cells(1, 37).Value = "dateLastTdvrUpdate"


    ' Give these cells a FALSE value, so the SQL import doesn't choke
    Range(Cells(2, 37), Cells(intLastRow, 37)).Value = Format$(Now(), "yyyy-mm-ddTHH:MM:SS")

    ' Save as CSV
    Call ExportToCSV(ThisWorkbook.Name, shtSQL.Name, "CSVforSQL")
    shtExport.Name = "Old TDVR Export"



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
