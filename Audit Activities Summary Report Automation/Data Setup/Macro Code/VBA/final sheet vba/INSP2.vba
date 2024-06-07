Sub Macro1()
'
' Macro1 Macro
'

'
    ActiveWindow.SmallScroll Down:=-3
    Columns("J:J").Select
    ActiveWorkbook.Worksheets("INSP").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("INSP").Sort.SortFields.Add Key:=Range("J1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("INSP").Sort
        .SetRange Range("A2:N3451")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
