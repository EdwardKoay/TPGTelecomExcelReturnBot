Attribute VB_Name = "Module18"
Sub REFUNDREPORT_Ctrl_R()
Attribute REFUNDREPORT_Ctrl_R.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+r
'
    ' Open the data workbook
    
    Dim n As Integer
    Dim n2 As Integer
    Dim wb As Workbook
    Dim eApp As Excel.Application
    
    Application.ScreenUpdating = False
    Selection.Copy
    Windows("Equipment Returned.xls").Activate
    'Workbooks.Open ("C:\Users\T6433677\Desktop\Reports\Equipment Returned.xls")
    
    'For Each wb In Workbooks
        'If wb.Name = "Equipment Returned.xls" Then
            'Exit For
        'End If
    'Next
    
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    
    n = Selection.Rows.Count
    n2 = n + 2
    
    Range(Cells(n2 + 1, "A"), Cells(65536, "Z")).Delete Shift:=xlUp
    Range("I:K").Delete Shift:=xlLeft
    
    Range("E3:F3").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("A2").Select
    Application.ScreenUpdating = True
    
    
End Sub
Sub RETURNEDRTSREPORT_Ctrl_T()
Attribute RETURNEDRTSREPORT_Ctrl_T.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+t
'

    Application.ScreenUpdating = False
    Selection.Copy
    Windows("Modems - RTS.xls").Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    
    Dim n As Integer
    Dim n2 As Integer
    
    n = Selection.Rows.Count
    n2 = n + 2
    
    Range(Cells(n2 + 1, "A"), Cells(65536, "Z")).Delete Shift:=xlUp
    Range("I:K").Delete Shift:=xlLeft
    
    Range("E3:F3").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("A2").Select
    Application.ScreenUpdating = True
    
End Sub
Sub RETURNEDLOANREPORT_Ctrl_E()
Attribute RETURNEDLOANREPORT_Ctrl_E.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+e
'

    Application.ScreenUpdating = False
    Selection.Copy
    Windows("LMAR Returns.xls").Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    
    Dim n As Integer
    Dim n2 As Integer
    
    n = Selection.Rows.Count
    n2 = n + 2
    
    Range(Cells(n2 + 1, "A"), Cells(65536, "Z")).Delete Shift:=xlUp
    Range("I:K").Delete Shift:=xlLeft
    
    Range("E3:F3").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("A2").Select
    Application.ScreenUpdating = True
    
End Sub

' Sub RETURNEDRTSREPORT()
Sub RETURNEDiiNETREPORT_Ctrl_Y()
Attribute RETURNEDiiNETREPORT_Ctrl_Y.VB_ProcData.VB_Invoke_Func = "Y\n14"

'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+y
'

    Application.ScreenUpdating = False
    Selection.Copy
    Windows("iiNet Returns.xls").Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    
    Dim n As Integer
    Dim n2 As Integer
    
    n = Selection.Rows.Count
    n2 = n + 2
    
    Range(Cells(n2 + 1, "A"), Cells(65536, "Z")).Delete Shift:=xlUp
    'Range("I:K").Delete Shift:=xlLeft
    
    Range("E3:F3").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("A2").Select
    
    Application.ScreenUpdating = True
End Sub

