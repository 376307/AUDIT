Sub Macro3()
'
' Macro3 Macro
'

'
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("F1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$F$7601").AutoFilter Field:=6, Criteria1:= _
        "IN PROGRESS"
ActiveWorkbook.Save
End Sub
