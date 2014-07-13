Attribute VB_Name = "m_Export_To_CSV"
Option Explicit

Public Sub ExportToCSV(strCurWBName As String, strCurWSName As String, Optional strFilePrefix As String)

    Dim strfilename As String
    Dim strOutputPrefix$
    Dim strNow As String
    Dim strDate As String
    Dim strTime As String
    Dim strMin As String
    Dim strCurDir As String
    Dim strAWName As String
    Dim strNewWBName As String
    Dim strTempNewWBName As String
    Dim vbOpenFolder
    Dim strPath As String
    Dim intLastRow%, i%
    Dim wbCSVWorkbook As Workbook
    Dim wbOrigBook As Workbook
    Dim t

    Set wbOrigBook = ThisWorkbook
    strDate = Date
    If Len(Minute(Time)) = 1 Then strMin = "0" & Minute(Time) Else strMin = Minute(Time)
    strTime = Hour(Time) & strMin
    'strTime = Left(Time, Len(Time) - 6)
    'strTime = Replace(strTime, ":", "")
    If Len(strTime) = 3 Then strTime = "0" & strTime
    strDate = Replace(strDate, "/", "")

    If strFilePrefix = vbNullString Then strOutputPrefix = strCurWBName Else strOutputPrefix = strFilePrefix
    strNewWBName = strOutputPrefix & strDate & "at" & strTime & ".csv"

    strAWName = ActiveWorkbook.Name

    t = Timer
    While Timer < t + 0.1
        DoEvents
    Wend


    strTempNewWBName = "CSVforSQL"
    Set wbCSVWorkbook = Workbooks.Add
    strCurDir = ThisWorkbook.Path

    Application.DisplayAlerts = False
    wbOrigBook.Sheets("SQL").Copy Before:=wbCSVWorkbook.Sheets(1)

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''


    strPath = strCurDir & "\CSV EXPORT\"
    strfilename = strPath & strNewWBName
    If Len(Dir(strPath, vbDirectory)) = 0 Then
        MkDir strPath
    End If
    t = Timer
    While Timer < t + 0.5
        DoEvents
    Wend
    wbCSVWorkbook.SaveAs FileName:=strfilename, FileFormat:=xlCSV, CreateBackup:=False
    t = Timer
    While Timer < t + 1
        DoEvents
    Wend

    wbCSVWorkbook.Close
    Set wbCSVWorkbook = Nothing
    t = Timer
    While Timer < t + 0.5
        DoEvents
    Wend

    Application.DisplayAlerts = True

    vbOpenFolder = MsgBox("The file was exported successfully. You may find it in the same directory as this workbook." & vbCrLf & vbCrLf _
                        & "Open the folder location now?", vbYesNo)

    If vbOpenFolder = vbYes Then
        Shell "explorer.exe " & strPath, vbNormalFocus
    End If

End Sub
