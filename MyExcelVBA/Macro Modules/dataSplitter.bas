Attribute VB_Name = "dataSplitter"
Option Explicit
Dim mydir As String
Dim wbk As Workbook
Sub dataSplitter()
Dim colCount, rowCount, lastRow, colLeft, colRight, rowFirst, rowLast As Long
Dim ws, wsNew As Worksheet
Dim sheetNames, headerRow, b, c, copyRange As Range
Dim rowAbove, rowBelow, newSheet As String

Application.ScreenUpdating = True
Application.DisplayAlerts = False

Set wbk = ActiveWorkbook
Set ws = ActiveSheet
Set headerRow = Application.InputBox("Please select the row(s) with your column headers", Type:=8)
    Range("a1").Select
Set sheetNames = Application.InputBox("Please select the range of split variables", Type:=8)
    Range("a1").Select

With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    If .Show <> -1 Then MsgBox "No folder selected! Exiting sub...": Exit Sub
    mydir = .SelectedItems(1)
End With

Application.ScreenUpdating = False

colCount = headerRow.Columns.count
rowCount = headerRow.Rows.count
colLeft = headerRow.Column
colRight = colLeft + colCount - 1
    For Each b In sheetNames
        rowAbove = ws.Cells(b.row - 1, b.Column).Value
        rowBelow = ws.Cells(b.row + 1, b.Column).Value
        If b.Value <> rowAbove Then
            rowFirst = b.row
            newSheet = Left(b.Value, 31)
            Sheets.Add.Name = newSheet
            Set wsNew = ActiveSheet
            ActiveSheet.Move After:=Sheets(wbk.Sheets.count)
            headerRow.Copy
            Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            ActiveSheet.Paste
            ActiveSheet.Range("A1").Select
        End If
        ws.Select
        If b.Value = rowAbove Then rowLast = b.row
        If b.Value <> rowBelow Then
            Set copyRange = ws.Range(Cells(rowFirst, colLeft), Cells(rowLast, colRight))
            copyRange.Copy
            wsNew.Select
            lastRow = wsNew.Cells(wsNew.Rows.count, "A").End(xlUp).row + 1
            wsNew.Cells(lastRow, 1).Select
            ActiveSheet.Paste
            Cells.Select
            Application.CutCopyMode = False
            With Selection
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.AutoFilter
            Cells.EntireColumn.AutoFit
            Range("A1").Select
            ws.Select
            copySheet fileName:=newSheet, saveDir:=mydir
            Sheets(newSheet).Delete
        End If
        ws.Select
    Next b
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Private Function copySheet(ByVal fileName As String, ByVal saveDir As String)
Dim wbkA As Workbook
Dim Sh As Worksheet
Set wbkA = Workbooks.Add
    DoEvents
    wbk.Worksheets(fileName).Copy wbkA.Sheets(1) ' Copy activesheet before the first sheet of wbk
    For Each Sh In wbkA.Sheets
        If Sh.Name <> fileName Then Sh.Delete
    Next Sh
    With wbkA
        .ActiveSheet.SaveAs saveDir & "\" & fileName & ".xlsx"
        .Close SaveChanges:=False
    End With
    Set wbkA = Nothing
End Function
