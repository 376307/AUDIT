Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("INSP").Select
    Columns("J:J").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("INSP").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("INSP").Sort.SortFields.Add Key:=Range("J1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("INSP").Sort
        .SetRange Range("A2:N3540")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
 ActiveWorkbook.Save
    End With
End Sub
