Attribute VB_Name = "CCH_Report_Macros"
Sub cleanELF()
    Columns("A:J").EntireColumn.AutoFit
    Columns("E:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").FormulaR1C1 = "ClientID"
    Range("F1").FormulaR1C1 = "FBAR?"
    Range("E2").FormulaR1C1 = _
        "=IFERROR(VALUE(MID(RC[-1],5,LEN(RC[-1])-7)),MID(RC[-1],5,LEN(RC[-1])-7))"
    Range("F2").FormulaR1C1 = "=IF(LEFT(RC[4],4)=""FBAR"", ""FBAR"", ""Tax Return"")"
    Range("E2:F2").Copy
    Range("E2:F" & lastRow(ActiveSheet.Name, "A")).PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    Calculate
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveWorkbook.Worksheets("Silverlight Export").Sort.SortFields.Clear
    
    Columns("I:I").Select
    Selection.TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("L:L").Select
    Selection.TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveWorkbook.Worksheets("Silverlight Export").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Silverlight Export").Sort.SortFields.Add Key:= _
        Range("F2:F" & lastRow(ActiveSheet.Name, "A")), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets("Silverlight Export").Sort.SortFields.Add Key:= _
        Range("E2:E" & lastRow(ActiveSheet.Name, "A")), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets("Silverlight Export").Sort.SortFields.Add Key:= _
        Range("I2:I" & lastRow(ActiveSheet.Name, "A")), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Silverlight Export").Sort.SortFields.Add Key:= _
        Range("L2:L" & lastRow(ActiveSheet.Name, "A")), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Silverlight Export").Sort
        .SetRange Range("A1:L" & lastRow(ActiveSheet.Name, "A"))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
End Sub


