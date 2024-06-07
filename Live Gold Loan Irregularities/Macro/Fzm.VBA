Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Daily Report").Select
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    Columns("AH:AH").Select
    Selection.EntireColumn.Hidden = True
    Sheets("GoldLoan_Irregularities").Select
    Columns("AH:AH").Select
    Selection.EntireColumn.Hidden = True
    ActiveWorkbook.Save
End Sub
