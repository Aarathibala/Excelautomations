/*
    Developer Name : Balaji
    Used to Clear the existing details on the input sheet

*/

Option Explicit
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
