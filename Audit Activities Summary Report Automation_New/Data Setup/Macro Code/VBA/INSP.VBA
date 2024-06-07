Sub Macro1()
'
' Macro1 Macro
'

'
    Cells.Select
    Range("C1").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("N3").Select
    Application.CutCopyMode = False
    Cells.Select
    Range("C1").Activate
    Cells.EntireColumn.AutoFit
    ActiveWorkbook.Save
End Sub
