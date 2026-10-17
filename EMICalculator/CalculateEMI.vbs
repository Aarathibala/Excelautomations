/*
    Developer Name : Balaji

    First Commit Date: 2023-04-18
    To Calculate the EMI Amount for the Entire Tenure
    with  Principle Amount and Interest Break up for Each month


*/
Option Explicit
Sub CalculateEMI()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableCancelKey = xlDisabled

Dim lastrw As Long, tenure As Long, i As Long, fr As Long, I_IEMIAMT As Double, inc As Long, Rng As Range
Dim Op_bal As Double, cl_bal As Double, iPrinciple As Double, iInterest As Double, iRate As Double, endrw As Long
Dim cum_interest As Double

Clearcontents
lastrw = Sheet1.Range("B" & Rows.Count).End(xlUp).Row
tenure = Sheet1.Range("D5")
endrw = lastrw + tenure
fr = lastrw + 1
Op_bal = Sheet1.Range("D3")
iRate = Sheet1.Range("D4")
inc = 0
cum_interest = 0
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
  cum_interest = cum_interest + iInterest
Next i
Sheet1.Range("H3").Value = cum_interest

End Sub
