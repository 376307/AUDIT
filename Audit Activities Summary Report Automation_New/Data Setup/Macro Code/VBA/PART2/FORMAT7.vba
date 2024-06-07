Sub Macro2()
'
' Macro2 Macro
'

'
    Sheets("INSP").Select
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=27
    Range("A3527").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=-27
    ActiveWindow.ScrollRow = 3478
    ActiveWindow.ScrollRow = 3462
    ActiveWindow.ScrollRow = 3445
    ActiveWindow.ScrollRow = 3429
    ActiveWindow.ScrollRow = 3421
    ActiveWindow.ScrollRow = 3209
    ActiveWindow.ScrollRow = 3136
    ActiveWindow.ScrollRow = 3062
    ActiveWindow.ScrollRow = 2965
    ActiveWindow.ScrollRow = 2069
    ActiveWindow.ScrollRow = 2004
    ActiveWindow.ScrollRow = 1817
    ActiveWindow.ScrollRow = 1784
    ActiveWindow.ScrollRow = 1662
    ActiveWindow.ScrollRow = 1067
    ActiveWindow.ScrollRow = 1002
    ActiveWindow.ScrollRow = 961
    ActiveWindow.ScrollRow = 888
    ActiveWindow.ScrollRow = 636
    ActiveWindow.ScrollRow = 611
    ActiveWindow.ScrollRow = 571
    ActiveWindow.ScrollRow = 554
    ActiveWindow.ScrollRow = 530
    ActiveWindow.ScrollRow = 424
    ActiveWindow.ScrollRow = 408
    ActiveWindow.ScrollRow = 375
    ActiveWindow.ScrollRow = 359
    ActiveWindow.ScrollRow = 343
    ActiveWindow.ScrollRow = 326
    ActiveWindow.ScrollRow = 269
    ActiveWindow.ScrollRow = 253
    ActiveWindow.ScrollRow = 245
    ActiveWindow.ScrollRow = 237
    ActiveWindow.ScrollRow = 229
    ActiveWindow.ScrollRow = 212
    ActiveWindow.ScrollRow = 204
    ActiveWindow.ScrollRow = 196
    ActiveWindow.ScrollRow = 180
    ActiveWindow.ScrollRow = 163
    ActiveWindow.ScrollRow = 155
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 1
    Range("A1").Select
 ActiveWorkbook.Save
End Sub
