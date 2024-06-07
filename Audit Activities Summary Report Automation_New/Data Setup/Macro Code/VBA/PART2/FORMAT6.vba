Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("INSP").Select
    Columns("O:O").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("O2").Select
    Application.CutCopyMode = False
 ActiveWorkbook.Save
End Sub
