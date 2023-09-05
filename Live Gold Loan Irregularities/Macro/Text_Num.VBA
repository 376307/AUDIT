Sub Text_Num()

    Sheets("GoldLoan_Irregularities").Select
    Columns("U:U").Select
    Selection.TextToColumns Destination:=Range("U1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
End Sub

