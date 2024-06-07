Sub Macro1()
'
' Macro1 Macro
'

'
    Columns("E:E").Select
    Selection.NumberFormat = "mm/dd/yyyy"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "AUD_NAME"
    Columns("H:H").Select
    Columns("H:H").EntireColumn.AutoFit
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "STATUS"
    Range("B1").Select
 ActiveWorkbook.Save
End Sub
