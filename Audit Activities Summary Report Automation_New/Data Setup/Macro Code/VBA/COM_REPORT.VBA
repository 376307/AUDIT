Sub Macro1()
'
' Macro1 Macro
'

'
    Range("E2").Select
    Range("E2:E3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "m/d/yyyy"
End Sub
