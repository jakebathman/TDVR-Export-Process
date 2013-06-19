Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Columns("E:E").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert shift:=xlToRight
    Range("G18").Select
End Sub
