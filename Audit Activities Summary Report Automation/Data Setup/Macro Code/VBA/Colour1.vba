Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Summary").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("G7:P17").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("P17").Select
    Sheets("Count").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Summary").Select
    Range("M23").Select
    ActiveWindow.SmallScroll Down:=-9
End Sub
