Attribute VB_Name = "dataSwing"
Sub DataSwing()
Dim lastRow, lastRowA, count, rowA, colA, repeatCols, runLoops, headerRows As Long
Dim headerRepeat, headerSwing, rngA, rngB, rngC, rngD, delRng As Range

Set headerRepeat = Application.InputBox("Please select cells in your header row with repeating data", Type:=8)
    Range("a1").Select
Set headerSwing = Application.InputBox("Please select cells in your header row with swing data", Type:=8)
    Range("a1").Select
repeatCols = InputBox("How many columns are in the grouping to repeat on each swing")
repeatCols = repeatCols + 0

count = 1

runLoops = headerSwing.Columns.count / repeatCols

Do While count < runLoops
    On Error Resume Next
    lastRowA = fnLastRow(1, Columns(headerRepeat.Column))
    lastRow = fnLastRow(1, Range(Cells(headerSwing.Rows.count + 1, headerSwing.Column), Cells(lastRowA, headerSwing.Column)))
    lastRowA = lastRowA + 1
    colA = headerSwing.Column + repeatCols + repeatCols - 1
    headerRows = headerSwing.Rows.count
    Set rngA = Cells(headerRows, colA)
    Set rngB = Range(Cells(headerRepeat.row, headerRepeat.Column), Cells(lastRowA, headerRepeat.Columns.count + headerSwing.Columns.count))
    
    rngB.AutoFilter Field:=colA, Criteria1:="<>"
    
    rngA.Select
    NextVisibleRow
    rowA = ActiveCell.row
    If rowA > lastRow Then
        Set delRng = Range(Columns(headerSwing.Column + repeatCols), Columns(headerSwing.Columns.count + headerSwing.Column - count + 2))
        count = runLoops
        GoTo deleteColumns
    End If
    Set rngC = Range(Cells(rowA, headerRepeat.Column), Cells(lastRow, headerRepeat.Columns.count))
    Set rngD = Range(Cells(rowA, headerSwing.Column + repeatCols), Cells(lastRow, colA))
    rngC.Select
    rngC.Copy
    Cells(lastRowA, headerRepeat.Column).Select
    ActiveSheet.Paste
    rngD.Copy
    Cells(lastRowA, headerSwing.Column).Select
    ActiveSheet.Paste
    Set delRng = Range(Columns(headerSwing.Column + repeatCols), Columns(colA))
    GoTo deleteColumns
    
deleteColumns:
    delRng.Select
    delRng.Delete
    Range("A1").Select
    count = count + 1
Loop
End Sub
Private Sub NextVisibleRow()
Dim x As Long, y As Long
x = ActiveCell.row
y = ActiveCell.Column
Do
    x = x + 1
Loop Until Cells(x, y).EntireRow.Hidden = False
Cells(x, y).Select
End Sub
