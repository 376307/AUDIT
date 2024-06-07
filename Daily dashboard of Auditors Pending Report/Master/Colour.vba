Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Summary").Select
    Sheets("Summary").Move Before:=Sheets(1)
    Cells.Select
    Range("D1").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("F5").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    Range("H6:H7").Select
    ActiveWindow.SmallScroll Down:=-9
End Sub
