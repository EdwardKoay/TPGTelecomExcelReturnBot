Attribute VB_Name = "Module24"
Sub a_GetCountt()


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim LastRow As Integer

    Range("M65535").Select
    Selection.End(xlUp).Select
    LastRow = Selection.Row
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=HOUR(RC[-3])"
    Range("P4").Select
    Selection.Copy
    Range(Cells(4, 16), Cells(LastRow, 16)).Select
    ActiveSheet.Paste
    Range("Q4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "9"
    Range("Q5").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("Q4:Q5").Select
    Selection.AutoFill Destination:=Range("Q4:Q12"), Type:=xlFillDefault
    Range("Q4:Q12").Select
    Range("R4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C16,RC[-1])"
    Range("R4").Select
    Selection.AutoFill Destination:=Range("R4:R12"), Type:=xlFillDefault
    Range("R4:R12").Select
    
    
    '' nextange("T1").Select
    Range("T1").FormulaR1C1 = "0"
    Range("T2").FormulaR1C1 = "STAFF"
    Range("T3").Select
    ActiveCell.FormulaR1C1 = _
        "=LOOKUP(2,1/(COUNTIF(R1C20:R[-1]C,R4C11:R1048576C11)=0),R4C11:R1048576C11)"
    Selection.AutoFill Destination:=Range("T3:T6"), Type:=xlFillDefault
    Range("T3:T6").Select
    Range("U3").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C11,RC[-1])"
    Range("U3").Select
    Selection.AutoFill Destination:=Range("U3:U6"), Type:=xlFillDefault
    Range("U3:U6").Select
    Range("T1").Select
    Columns("T:T").Select
    ActiveSheet.Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("T1").Select
    Application.CutCopyMode = False
    
    ActiveSheet.Calculate
Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Range("V3").Formula = "=U3/'Time Attack'!J11"
    
End Sub
