Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("INSP").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("N:N").Select
    Selection.Delete Shift:=xlToLeft
    Range("A2:N2").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.ScrollRow = 3461
    ActiveWindow.ScrollRow = 3441
    ActiveWindow.ScrollRow = 3396
    ActiveWindow.ScrollRow = 3243
    ActiveWindow.ScrollRow = 3124
    ActiveWindow.ScrollRow = 3059
    ActiveWindow.ScrollRow = 2053
    ActiveWindow.ScrollRow = 1983
    ActiveWindow.ScrollRow = 918
    ActiveWindow.ScrollRow = 823
    ActiveWindow.ScrollRow = 387
    ActiveWindow.ScrollRow = 357
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 1
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$N2=""NOT STARTED"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 6053114
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$N2=""STARTED"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 427
    ActiveWindow.ScrollRow = 506
    ActiveWindow.ScrollRow = 536
    ActiveWindow.ScrollRow = 744
    ActiveWindow.ScrollRow = 769
    ActiveWindow.ScrollRow = 1017
    ActiveWindow.ScrollRow = 1106
    ActiveWindow.ScrollRow = 1141
    ActiveWindow.ScrollRow = 1483
    ActiveWindow.ScrollRow = 1488
    ActiveWindow.ScrollRow = 1498
    ActiveWindow.ScrollRow = 1503
    ActiveWindow.ScrollRow = 1498
    ActiveWindow.ScrollRow = 1483
    ActiveWindow.ScrollRow = 1473
    ActiveWindow.ScrollRow = 1468
    ActiveWindow.ScrollRow = 1463
    ActiveWindow.ScrollRow = 1458
    ActiveWindow.ScrollRow = 1453
    ActiveWindow.ScrollRow = 1458
    ActiveWindow.ScrollRow = 1468
    ActiveWindow.ScrollRow = 1478
    ActiveWindow.ScrollRow = 1493
    ActiveWindow.ScrollRow = 1508
    ActiveWindow.ScrollRow = 1522
    ActiveWindow.ScrollRow = 1527
    ActiveWindow.ScrollRow = 1532
    ActiveWindow.ScrollRow = 1537
    ActiveWindow.ScrollRow = 1542
    ActiveWindow.ScrollRow = 1547
    ActiveWindow.ScrollRow = 1552
    ActiveWindow.ScrollRow = 1567
    ActiveWindow.ScrollRow = 1572
    ActiveWindow.ScrollRow = 1577
    ActiveWindow.ScrollRow = 1582
    ActiveWindow.ScrollRow = 1592
    ActiveWindow.ScrollRow = 1602
    ActiveWindow.ScrollRow = 1612
    ActiveWindow.ScrollRow = 1617
    ActiveWindow.ScrollRow = 1661
    ActiveWindow.ScrollRow = 1666
    ActiveWindow.ScrollRow = 1671
    ActiveWindow.ScrollRow = 1681
    ActiveWindow.ScrollRow = 1686
    ActiveWindow.ScrollRow = 1691
    ActiveWindow.ScrollRow = 1701
    ActiveWindow.ScrollRow = 1706
    ActiveWindow.ScrollRow = 1711
    ActiveWindow.ScrollRow = 1716
    ActiveWindow.ScrollRow = 1721
    ActiveWindow.ScrollRow = 1726
    ActiveWindow.ScrollRow = 1731
    ActiveWindow.ScrollRow = 1736
    ActiveWindow.ScrollRow = 1741
    ActiveWindow.ScrollRow = 1746
    ActiveWindow.ScrollRow = 1750
    ActiveWindow.ScrollRow = 1760
    ActiveWindow.ScrollRow = 1765
    ActiveWindow.ScrollRow = 1800
    ActiveWindow.ScrollRow = 1805
    ActiveWindow.ScrollRow = 1810
    ActiveWindow.ScrollRow = 1815
    ActiveWindow.ScrollRow = 1820
    ActiveWindow.ScrollRow = 1830
    ActiveWindow.ScrollRow = 1840
    ActiveWindow.ScrollRow = 1855
    ActiveWindow.ScrollRow = 1894
    ActiveWindow.ScrollRow = 1914
    ActiveWindow.ScrollRow = 1924
    ActiveWindow.ScrollRow = 1929
    ActiveWindow.ScrollRow = 1934
    ActiveWindow.ScrollRow = 1939
    ActiveWindow.ScrollRow = 2008
    ActiveWindow.ScrollRow = 2028
    ActiveWindow.ScrollRow = 2048
    ActiveWindow.ScrollRow = 2058
    ActiveWindow.ScrollRow = 2107
    ActiveWindow.ScrollRow = 2112
    ActiveWindow.ScrollRow = 2117
    ActiveWindow.ScrollRow = 2132
    ActiveWindow.ScrollRow = 2147
    ActiveWindow.ScrollRow = 2162
    ActiveWindow.ScrollRow = 2484
    ActiveWindow.ScrollRow = 2509
    ActiveWindow.ScrollRow = 2544
    ActiveWindow.ScrollRow = 2613
    ActiveWindow.ScrollRow = 2618
    ActiveWindow.ScrollRow = 2608
    ActiveWindow.ScrollRow = 2588
    ActiveWindow.ScrollRow = 2529
    ActiveWindow.ScrollRow = 2509
    ActiveWindow.ScrollRow = 2479
    ActiveWindow.ScrollRow = 2444
    ActiveWindow.ScrollRow = 2093
    ActiveWindow.ScrollRow = 1988
    ActiveWindow.ScrollRow = 1929
    ActiveWindow.ScrollRow = 1582
    ActiveWindow.ScrollRow = 1552
    ActiveWindow.ScrollRow = 1195
    ActiveWindow.ScrollRow = 1165
    ActiveWindow.ScrollRow = 1086
    ActiveWindow.ScrollRow = 848
    ActiveWindow.ScrollRow = 769
    ActiveWindow.ScrollRow = 759
    ActiveWindow.ScrollRow = 561
    ActiveWindow.ScrollRow = 531
    ActiveWindow.ScrollRow = 328
    ActiveWindow.ScrollRow = 303
    ActiveWindow.ScrollRow = 268
    ActiveWindow.ScrollRow = 174
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 129
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 1
    Range("A1:N1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1:N1").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("F5").Select
    Columns("N:N").EntireColumn.AutoFit
    ActiveWindow.SmallScroll Down:=-12
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("N2").Select
    ActiveWindow.SmallScroll Down:=-57
    Range("F2").Select
    ActiveWindow.SmallScroll Down:=-9
End Sub
