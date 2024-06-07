Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("INSP").Select
    Cells.Select
    Range("C1").Activate
    Cells.EntireColumn.AutoFit
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("N1").Select
    Application.CutCopyMode = False
    Range("N2").Select
    Selection.AutoFilter
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveSheet.Range("$A$1:$N$3525").AutoFilter Field:=14, Criteria1:= _
        "COMPLETED"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "NOT STARTED"
    Selection.FillDown
    ActiveWindow.SmallScroll Down:=-6
    ActiveSheet.Range("$A$1:$N$3525").AutoFilter Field:=14, Criteria1:= _
        "IN PROGRESS"
    ActiveCell.FormulaR1C1 = "STARTED"
    Range("N2").Select
    Selection.FillDown
    ActiveWindow.SmallScroll Down:=-33
    Selection.AutoFilter
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWorkbook.Save
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Range("S17").Select
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=45
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 73
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 163
    ActiveWindow.ScrollRow = 187
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 341
    ActiveWindow.ScrollRow = 349
    ActiveWindow.ScrollRow = 365
    ActiveWindow.ScrollRow = 390
    ActiveWindow.ScrollRow = 414
    ActiveWindow.ScrollRow = 568
    ActiveWindow.ScrollRow = 576
    ActiveWindow.ScrollRow = 600
    ActiveWindow.ScrollRow = 641
    ActiveWindow.ScrollRow = 673
    ActiveWindow.ScrollRow = 698
    ActiveWindow.ScrollRow = 852
    ActiveWindow.ScrollRow = 860
    ActiveWindow.ScrollRow = 868
    ActiveWindow.ScrollRow = 892
    ActiveWindow.ScrollRow = 917
    ActiveWindow.ScrollRow = 941
    ActiveWindow.ScrollRow = 1111
    ActiveWindow.ScrollRow = 1127
    ActiveWindow.ScrollRow = 1144
    ActiveWindow.ScrollRow = 1152
    ActiveWindow.ScrollRow = 1168
    ActiveWindow.ScrollRow = 1192
    ActiveWindow.ScrollRow = 1216
    ActiveWindow.ScrollRow = 1338
    ActiveWindow.ScrollRow = 1354
    ActiveWindow.ScrollRow = 1362
    ActiveWindow.ScrollRow = 1371
    ActiveWindow.ScrollRow = 1379
    ActiveWindow.ScrollRow = 1403
    ActiveWindow.ScrollRow = 1427
    ActiveWindow.ScrollRow = 1476
    ActiveWindow.ScrollRow = 1492
    ActiveWindow.ScrollRow = 1500
    ActiveWindow.ScrollRow = 1508
    ActiveWindow.ScrollRow = 1516
    ActiveWindow.ScrollRow = 1533
    ActiveWindow.ScrollRow = 1557
    ActiveWindow.ScrollRow = 1654
    ActiveWindow.ScrollRow = 1679
    ActiveWindow.ScrollRow = 1711
    ActiveWindow.ScrollRow = 1735
    ActiveWindow.ScrollRow = 1752
    ActiveWindow.ScrollRow = 1776
    ActiveWindow.ScrollRow = 1914
    ActiveWindow.ScrollRow = 1922
    ActiveWindow.ScrollRow = 1946
    ActiveWindow.ScrollRow = 1970
    ActiveWindow.ScrollRow = 1987
    ActiveWindow.ScrollRow = 1995
    ActiveWindow.ScrollRow = 2003
    ActiveWindow.ScrollRow = 2092
    ActiveWindow.ScrollRow = 2100
    ActiveWindow.ScrollRow = 2116
    ActiveWindow.ScrollRow = 2149
    ActiveWindow.ScrollRow = 2189
    ActiveWindow.ScrollRow = 2230
    ActiveWindow.ScrollRow = 2327
    ActiveWindow.ScrollRow = 2335
    ActiveWindow.ScrollRow = 2343
    ActiveWindow.ScrollRow = 2359
    ActiveWindow.ScrollRow = 2368
    ActiveWindow.ScrollRow = 2376
    ActiveWindow.ScrollRow = 2384
    ActiveWindow.ScrollRow = 2392
    ActiveWindow.ScrollRow = 2400
    ActiveWindow.ScrollRow = 2408
    ActiveWindow.ScrollRow = 2416
    ActiveWindow.ScrollRow = 2424
    ActiveWindow.ScrollRow = 2432
    ActiveWindow.ScrollRow = 2441
    ActiveWindow.ScrollRow = 2457
    ActiveWindow.ScrollRow = 2465
    ActiveWindow.ScrollRow = 2473
    ActiveWindow.ScrollRow = 2489
    ActiveWindow.ScrollRow = 2497
    ActiveWindow.ScrollRow = 2505
    ActiveWindow.ScrollRow = 2514
    ActiveWindow.ScrollRow = 2538
    ActiveWindow.ScrollRow = 2554
    ActiveWindow.ScrollRow = 2570
    ActiveWindow.ScrollRow = 2578
    ActiveWindow.ScrollRow = 2586
    ActiveWindow.ScrollRow = 2603
    ActiveWindow.ScrollRow = 2611
    ActiveWindow.ScrollRow = 2619
    ActiveWindow.ScrollRow = 2627
    ActiveWindow.ScrollRow = 2643
    ActiveWindow.ScrollRow = 2651
    ActiveWindow.ScrollRow = 2659
    ActiveWindow.ScrollRow = 2676
    ActiveWindow.ScrollRow = 2684
    ActiveWindow.ScrollRow = 2692
    ActiveWindow.ScrollRow = 2700
    ActiveWindow.ScrollRow = 2724
    ActiveWindow.ScrollRow = 2732
    ActiveWindow.ScrollRow = 2740
    ActiveWindow.ScrollRow = 2749
    ActiveWindow.ScrollRow = 2757
    ActiveWindow.ScrollRow = 2765
    ActiveWindow.ScrollRow = 2773
    ActiveWindow.ScrollRow = 2781
    ActiveWindow.ScrollRow = 2789
    ActiveWindow.ScrollRow = 2805
    ActiveWindow.ScrollRow = 2813
    ActiveWindow.ScrollRow = 2822
    ActiveWindow.ScrollRow = 2830
    ActiveWindow.ScrollRow = 2846
    ActiveWindow.ScrollRow = 2854
    ActiveWindow.ScrollRow = 2862
    ActiveWindow.ScrollRow = 2870
    ActiveWindow.ScrollRow = 2878
    ActiveWindow.ScrollRow = 2886
    ActiveWindow.ScrollRow = 2895
    ActiveWindow.ScrollRow = 2903
    ActiveWindow.ScrollRow = 2927
    ActiveWindow.ScrollRow = 2967
    ActiveWindow.ScrollRow = 3040
    ActiveWindow.ScrollRow = 3097
    ActiveWindow.ScrollRow = 3113
    ActiveWindow.ScrollRow = 3138
    ActiveWindow.ScrollRow = 3146
    ActiveWindow.ScrollRow = 3154
    ActiveWindow.ScrollRow = 3162
    ActiveWindow.ScrollRow = 3170
    ActiveWindow.ScrollRow = 3186
    ActiveWindow.ScrollRow = 3194
    ActiveWindow.ScrollRow = 3203
    ActiveWindow.ScrollRow = 3211
    ActiveWindow.ScrollRow = 3219
    ActiveWindow.ScrollRow = 3227
    ActiveWindow.ScrollRow = 3235
    ActiveWindow.ScrollRow = 3243
    ActiveWindow.ScrollRow = 3251
    ActiveWindow.ScrollRow = 3259
    ActiveWindow.ScrollRow = 3267
    ActiveWindow.ScrollRow = 3284
    ActiveWindow.ScrollRow = 3292
    ActiveWindow.ScrollRow = 3300
    ActiveWindow.ScrollRow = 3308
    ActiveWindow.ScrollRow = 3316
    ActiveWindow.ScrollRow = 3324
    ActiveWindow.ScrollRow = 3332
    ActiveWindow.ScrollRow = 3340
    ActiveWindow.ScrollRow = 3348
    ActiveWindow.ScrollRow = 3357
    ActiveWindow.ScrollRow = 3373
    ActiveWindow.ScrollRow = 3381
    ActiveWindow.ScrollRow = 3389
    ActiveWindow.ScrollRow = 3405
    ActiveWindow.ScrollRow = 3421
    ActiveWindow.ScrollRow = 3438
    ActiveWindow.ScrollRow = 3446
    ActiveWindow.ScrollRow = 3454
    ActiveWindow.ScrollRow = 3462
    ActiveWindow.ScrollRow = 3470
    ActiveWindow.ScrollRow = 3478
    ActiveWindow.ScrollRow = 3494
    ActiveWindow.ScrollRow = 3503
    ActiveWindow.ScrollRow = 3494
    ActiveWindow.ScrollRow = 3478
    ActiveWindow.ScrollRow = 3470
    ActiveWindow.ScrollRow = 3462
    ActiveWindow.ScrollRow = 3454
    ActiveWindow.ScrollRow = 3438
    ActiveWindow.ScrollRow = 3430
    ActiveWindow.ScrollRow = 3421
    ActiveWindow.ScrollRow = 3413
    ActiveWindow.ScrollRow = 3397
    ActiveWindow.ScrollRow = 3389
    ActiveWindow.ScrollRow = 3381
    ActiveWindow.ScrollRow = 3373
    ActiveWindow.ScrollRow = 3381
    ActiveWindow.ScrollRow = 3397
    ActiveWindow.ScrollRow = 3405
    ActiveWindow.ScrollRow = 3413
    ActiveWindow.ScrollRow = 3430
    ActiveWindow.ScrollRow = 3438
    ActiveWindow.ScrollRow = 3446
    ActiveWindow.ScrollRow = 3454
    ActiveWindow.ScrollRow = 3470
    ActiveWindow.ScrollRow = 3478
    ActiveWindow.ScrollRow = 3486
    ActiveWindow.ScrollRow = 3494
    ActiveWindow.ScrollRow = 3486
    ActiveWindow.ScrollRow = 3478
    ActiveWindow.ScrollRow = 3470
    ActiveWindow.ScrollRow = 3462
    ActiveWindow.ScrollRow = 3454
    ActiveWindow.ScrollRow = 3446
    ActiveWindow.ScrollRow = 3438
    ActiveWindow.ScrollRow = 3430
    ActiveWindow.ScrollRow = 3446
    Range("A3452").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.ScrollRow = 3438
    ActiveWindow.ScrollRow = 3430
    ActiveWindow.ScrollRow = 3389
    ActiveWindow.ScrollRow = 3332
    ActiveWindow.ScrollRow = 3065
    ActiveWindow.ScrollRow = 3040
    ActiveWindow.ScrollRow = 3000
    ActiveWindow.ScrollRow = 2959
    ActiveWindow.ScrollRow = 2497
    ActiveWindow.ScrollRow = 2343
    ActiveWindow.ScrollRow = 2262
    ActiveWindow.ScrollRow = 2051
    ActiveWindow.ScrollRow = 1614
    ActiveWindow.ScrollRow = 1476
    ActiveWindow.ScrollRow = 1452
    ActiveWindow.ScrollRow = 1411
    ActiveWindow.ScrollRow = 1192
    ActiveWindow.ScrollRow = 1168
    ActiveWindow.ScrollRow = 1144
    ActiveWindow.ScrollRow = 1119
    ActiveWindow.ScrollRow = 1071
    ActiveWindow.ScrollRow = 941
    ActiveWindow.ScrollRow = 917
    ActiveWindow.ScrollRow = 892
    ActiveWindow.ScrollRow = 876
    ActiveWindow.ScrollRow = 868
    ActiveWindow.ScrollRow = 844
    ActiveWindow.ScrollRow = 795
    ActiveWindow.ScrollRow = 746
    ActiveWindow.ScrollRow = 730
    ActiveWindow.ScrollRow = 722
    ActiveWindow.ScrollRow = 714
    ActiveWindow.ScrollRow = 698
    ActiveWindow.ScrollRow = 690
    ActiveWindow.ScrollRow = 600
    ActiveWindow.ScrollRow = 552
    ActiveWindow.ScrollRow = 544
    ActiveWindow.ScrollRow = 527
    ActiveWindow.ScrollRow = 511
    ActiveWindow.ScrollRow = 495
    ActiveWindow.ScrollRow = 487
    ActiveWindow.ScrollRow = 446
    ActiveWindow.ScrollRow = 438
    ActiveWindow.ScrollRow = 414
    ActiveWindow.ScrollRow = 406
    ActiveWindow.ScrollRow = 390
    ActiveWindow.ScrollRow = 382
    ActiveWindow.ScrollRow = 373
    ActiveWindow.ScrollRow = 365
    ActiveWindow.ScrollRow = 341
    ActiveWindow.ScrollRow = 284
    ActiveWindow.ScrollRow = 276
    ActiveWindow.ScrollRow = 268
    ActiveWindow.ScrollRow = 252
    ActiveWindow.ScrollRow = 244
    ActiveWindow.ScrollRow = 236
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 203
    ActiveWindow.ScrollRow = 171
    ActiveWindow.ScrollRow = 146
    ActiveWindow.ScrollRow = 138
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 1
    Range("A1").Select
    ActiveWorkbook.Save
End Sub
