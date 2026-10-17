Attribute VB_Name = "Module_CalculateEMI"
Option Explicit
'====================================================
' Project     : EMI Calculator
' Module      : ClearContents
' Description : Clears EMI output range and borders
' Author      : Blaaji
' Created     : 2024-11-10
' Version     : 1.0.0
' Last Update : 2025-12-25
'====================================================
Sub CalculateEMI()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableCancelKey = xlDisabled

Dim lastrw As Long, tenure As Long, i As Long, fr As Long, I_IEMIAMT As Double, inc As Long, Rng As Range
Dim Op_bal As Double, cl_bal As Double, iPrinciple As Double, iInterest As Double, iRate As Double, endrw As Long

Clearcontents
lastrw = Sheet1.Range("B" & Rows.Count).End(xlUp).Row
tenure = Sheet1.Range("D5")
endrw = lastrw + tenure
fr = lastrw + 1
Op_bal = Sheet1.Range("D3")
iRate = Sheet1.Range("D4")
inc = 0
I_IEMIAMT = -(Application.WorksheetFunction.Pmt(iRate / 1200, tenure, Op_bal))
For i = fr To endrw
 On Error Resume Next
   
    iInterest = Op_bal * (iRate / 1200)
    iPrinciple = I_IEMIAMT - iInterest
    cl_bal = Op_bal - iPrinciple
    
    inc = inc + 1
    Sheet1.Range("B" & i) = inc
    Sheet1.Range("C" & i) = Op_bal
    Sheet1.Range("D" & i) = I_IEMIAMT
    Sheet1.Range("E" & i) = iInterest
    Sheet1.Range("F" & i) = iPrinciple
    Sheet1.Range("G" & i) = cl_bal
    Sheet1.Range("C" & i, "G" & i).Style = "comma"
    Op_bal = cl_bal
    Set Rng = Sheet1.Range(Cells(i, 2), Cells(i, 7))
    
    With Rng.Borders
        .LineStyle = "xlContinous"
         .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
  End With
Next i
 

End Sub
