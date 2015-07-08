Attribute VB_Name = "unpivot"
'unpivot a table
Sub unpivot()
Dim wbk As Workbook
Dim b, c, rowRng, colRng As Range
Dim rowVal, colVal, amount, destSheet As String
Dim rowCounter As Long
Dim wsOrig, wsNew As Worksheet

Set wbk = ActiveWorkbook
Set wsOrig = ActiveSheet
Set rowRng = Application.InputBox("Please select your column headers (exclude Row Labels, Blank, and Grand Totals)", Type:=8)
Range("a1").Select
Set colRng = Application.InputBox("Please select the range of Row Labels (exclude headers, blanks, and totals)", Type:=8)
Range("a1").Select
destSheet = InputBox("Please define the destination sheet name", "New Sheet Name")
If destSheet = "" Then destSheet = "NewSheet"

Sheets.Add.Name = destSheet
Set wsNew = ActiveSheet
ActiveSheet.Move After:=Sheets(wbk.Sheets.count)
Range("A1").Value = "Row Label"
Range("B1").Value = "Column Label"
Range("C1").Value = "Amount"
Sheets(wsOrig.Name).Select

Application.Calculation = xlCalculationManual
rowCounter = 2

For Each b In rowRng
    colVal = b.Value
    For Each c In colRng
        If Cells(c.row, b.Column).Formula <> vbNullString Then
            rowVal = c.Value
            amount = Cells(c.row, b.Column).Value
            Sheets(wsNew.Name).Range("A" & rowCounter).Value = rowVal
            Sheets(wsNew.Name).Range("B" & rowCounter).Value = colVal
            Sheets(wsNew.Name).Range("C" & rowCounter).Value = amount
            rowCounter = rowCounter + 1
        End If
    Next
Next
Application.Calculation = xlCalculationAutomatic
End Sub
