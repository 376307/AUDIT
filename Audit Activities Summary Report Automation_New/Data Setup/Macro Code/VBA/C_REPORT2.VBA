Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("com_report").Select
    Range("D55").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
End Sub
