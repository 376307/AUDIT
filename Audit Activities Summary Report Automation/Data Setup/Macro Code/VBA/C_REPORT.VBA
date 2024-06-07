Sub Macro1()
'
' Macro1 Macro
'

'
    Cells.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Columns("C:C").ColumnWidth = 8.29
    Columns("C:C").EntireColumn.AutoFit
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Range("A1").Select
End Sub
