Attribute VB_Name = "dataSplitterSameSheet"
Dim wbk As Workbook
Dim srcWs As Worksheet
Dim newWs As Worksheet
Dim headerRow As Range, sheetNames As Range, r As Range
Dim newSheet As String
Dim newLastRow As Long
Sub splitInFile()
    Set wbk = ActiveWorkbook
    Set srcWs = wbk.ActiveSheet
    Set headerRow = Application.InputBox("Please select the row(s) with your column headers", Type:=8)
        Range("a1").Select
    Set sheetNames = Application.InputBox("Please select the range of split variables (don't include header row)", Type:=8)
        Range("a1").Select
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    colCount = headerRow.Columns.count
    rowCount = headerRow.Rows.count
    colLeft = headerRow.Column
    colRight = colLeft + colCount - 1
    For Each r In sheetNames
        newSheet = Left(r.Value, 31)
        On Error Resume Next
        Set newWs = Sheets(newSheet)
        If newWs Is Nothing Then
            Sheets.Add.Name = newSheet
            Set newWs = ActiveSheet
            ActiveSheet.Move After:=Sheets(wbk.Sheets.count)
            headerRow.Copy
            Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            ActiveSheet.Range("A1").Select
            ActiveSheet.Paste
            ActiveSheet.Range("A1").Select
        End If
        srcWs.Select
        Range(Cells(r.row, colLeft), Cells(r.row, colRight)).Copy
        newWs.Select
        newLastRow = newWs.UsedRange.Rows.count + 1
        Cells(newLastRow, 1).PasteSpecial xlPasteValues, xlPasteSpecialOperationNone, False, False
        srcWs.Select
        Set newWs = Nothing
    Next r
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub fixit()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub
