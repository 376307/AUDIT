Sub Macro1()
'
' Macro1 Macro
'

'
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("G2").Select
    ActiveWorkbook.Save
End Sub
