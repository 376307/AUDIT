Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Region").Select
    Sheets("Region").Move Before:=Sheets(2)
    Sheets("Zone").Select
    Sheets("Zone").Move Before:=Sheets(3)
    Sheets("BANK VERIFICATION").Select
    Sheets("BANK VERIFICATION").Move Before:=Sheets(9)
    Sheets("Region").Select
    ActiveWindow.SmallScroll Down:=-42
    Range("C2").Select
    Sheets("Zone").Select
    ActiveWindow.SmallScroll Down:=-15
    Sheets("Summary").Select
    Range("E5:S6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("G5:L6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("E7:S21").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("E7:S21").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("E4:S4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Range("I14").Select
    ActiveWorkbook.Save
End Sub
