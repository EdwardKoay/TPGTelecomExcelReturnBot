Attribute VB_Name = "Module2"

Sub aa_Sort()
'
' Sort Macro
'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    LastRow = Range("A1048576").End(xlUp).Row()
    Dim n As Long
    Range(Cells(LastRow, "K"), Cells(LastRow, "K")).Select
    n = Selection.Row
    Range(Cells(4, 1).Address(), Cells(n, 11).Address()).Select
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add(Range(Cells(4, "J"), Cells(LastRow, "J")) _
    , xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(198, 239, 206)
    ActiveSheet.Sort.SortFields.Add(Range(Cells(4, "F"), Cells(LastRow, "F")), _
        xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(177, _
        160, 199)
    ActiveSheet.Sort.SortFields.Add(Range(Cells(4, "F"), Cells(LastRow, "F")), _
        xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, _
        255, 0)
    ActiveSheet.Sort.SortFields.Add(Range(Cells(4, "F"), Cells(LastRow, "F")), _
        xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(155, _
        187, 89)
    ActiveSheet.Sort.SortFields.Add(Range(Cells(4, "F"), Cells(LastRow, "F")), _
        xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(79, 129 _
        , 189)
    ActiveSheet.Sort.SortFields.Add Key:=Range(Cells(4, "F"), Cells(LastRow, "F")), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(Cells(4, "A"), Cells(LastRow, "K"))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'Range("F3").Copy
    'Columns("F:F").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("A4").Select
    ActiveSheet.Calculate
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    
End Sub


