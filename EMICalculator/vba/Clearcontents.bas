Attribute VB_Name = "Module_ClearContents"
Option Explicit
'====================================================
' Project     : EMI Calculator
' Module      : ClearContents
' Description : Clears EMI output range and borders
' Author      : Balaji
' Created     : 2024-11-10
' Version     : 1.0.0
' Last Update : 2025-12-25
'====================================================
Sub Clearcontents()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableCancelKey = xlDisabled

Dim lastrw As Long, Rng As Range
lastrw = Sheet1.Range("B" & Rows.Count).End(xlUp).Row

If lastrw > 9 Then
    Set Rng = Sheet1.Range("B" & 10, "G" & lastrw)
    With Rng
    .Clearcontents
    .Borders.LineStyle = xlNone
    End With
End If

End Sub
