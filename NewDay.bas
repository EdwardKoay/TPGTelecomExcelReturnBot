Attribute VB_Name = "Module1"
Sub aa_NewDay()
Attribute aa_NewDay.VB_ProcData.VB_Invoke_Func = " \n14"
'

'
' TimeLap Macro
'

'

    Application.ScreenUpdating = False

    Worksheets("Time Attack").Activate
    Application.CutCopyMode = False
    Range("B3:C20").ClearContents
    Range("B4:B20").Value = "1"
    Range("B7:B7").Value = "Load/Unload"
    Range("B10:B11").Value = "Lunch"
    Range("B12:B12").Value = "Cleanup/Sorting"
    Range("B15:B15").Value = "Load/Unload"
    Range("B18:B18").Value = "Cleanup/Sorting"
    Range("B20:B20").Value = "Reports"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "1"

' NewDay Macro
'

    Worksheets("Returns").Activate
    Rows("4:300").ClearContents
    Range("G3:J3").Copy
    Range("G4:J300").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("E3").Copy
    Range("E4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E4").Value = "" & WorksheetFunction.Text(Range("E3").Value, "d-MMM-YY")
    Application.CutCopyMode = False
    Selection.Copy
    Range("E4:E300").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A4").Select
    
    Sheets("Time Attack").Select
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "0"
    
    Sheets("Returns").Select
    
    
    Range("A3:A1000").NumberFormat = "@"
    Range("A2:D1000").NumberFormat = "@"
    Range("F2:H1000").NumberFormat = "@"
    Range("K:K").NumberFormat = "@"
    
    
    Rows("3:3").Copy
    Rows("3:300").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True
    Range("A4").Select
    
    
    
End Sub

