Attribute VB_Name = "Macros"
' Split File into many different sheets based on ColA value already being equal to an existing Sheet Name
Private Sub split()
Dim emplNum As String
Dim ws As Worksheet
Dim lastRow As Long
Dim c, emplNumRng, dataRng As Range
Set emplNumRng = Range("A2:A" & lastRow(ActiveSheet.Name, "A"))

For Each c In emplNumRng
    emplNum = c.Value
    Set dataRng = Range("A" & c.Row & ":BJ" & c.Row)
    With Sheets(emplNum)
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        .Range("A" & lastRow & ":BJ" & lastRow).Value = dataRng.Value
    End With
Next
End Sub

'unpivot a table
Private Sub unpivot()
Dim b, c, rowRng, colRng As Range
Dim rowVal, colVal, amount As String
Dim rowCounter As Long
Application.Calculation = xlCalculationManual
rowCounter = 2
'set number of rows to first row of pivot data to last without blanks or grand total
Set rowRng = Range("A5:A55")
'set column data points to header row without column A through last row of pivot without blanks or grand total
Set colRng = Range("B4:bm4")
For Each b In rowRng
    rowVal = b.Value
    For Each c In colRng
        If Cells(b.Row, c.Column).Formula <> vbNullString Then
            colVal = c.Value
            amount = Cells(b.Row, c.Column).Value
            'Sheet3 is whatever your destination sheet should be
            Sheets("Sheet3").Range("A" & rowCounter).Value = rowVal
            Sheets("Sheet3").Range("B" & rowCounter).Value = colVal
            Sheets("Sheet3").Range("D" & rowCounter).Value = amount
            rowCounter = rowCounter + 1
        End If
    Next
Next
Application.Calculation = xlCalculationAutomatic
End Sub

Sub ListWindows()
Dim wn As Excel.Window
For Each wn In Application.Windows
    Debug.Print wn.Caption
    If wn.Caption = "PERSONAL.XLSB" Then
        wn.Visible = True
        wn.Activate
        wn.Visible = False
    End If
Next wn
End Sub
Sub listFiles()
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim i As Integer

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder(GetFolder("Z:\"))
i = 2
'loops through each file in the directory and prints their names and path
For Each objFile In objFolder.Files
    'print file name
    Cells(i + 1, 2) = objFile.Name
    i = i + 1
Next objFile
End Sub
