Sub Macro1()
'
' Macro1 Macro
'

'
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("F:F").Select
    Selection.AutoFilter
    ActiveSheet.Range("$F$1:$F$9888").AutoFilter Field:=1, Criteria1:= _
        "IN PROGRESS"
    Cells.Select
    Range("F1").Activate
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Columns("D:D").EntireColumn.AutoFit
    Range("F2").Select
    Columns("F:F").EntireColumn.AutoFit
    Range("G2").Select
    Columns("B:B").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
 ActiveWorkbook.Save
End Sub
