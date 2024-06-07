Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Region").Select
    Columns("A:J").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("D7").Select
    Sheets("Zone").Select
    Columns("A:J").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("G6").Select
    Application.CutCopyMode = False
    Sheets("Zone").Select
    ActiveWorkbook.Save
End Sub