Attribute VB_Name = "insertColRight"
Sub insertColRight()
    Dim colA, colN, t As Long
    Dim r, n As Range
    Dim wks As Worksheet
    Set wks = ActiveSheet
    
    t = wks.UsedRange.Rows.count
    colA = ActiveCell.Column
    colN = colA + 1
    
    Set r = Range(Cells(1, colA), Cells(t, colA))
    Set n = Range(Cells(1, colN), Cells(t, colN))
    
    r.Select
    Selection.Copy
    n.Select
    ActiveSheet.Paste
    Range(Cells(1, 1), Cells(t, colN)).Select
    If ActiveSheet.AutoFilterMode Then
        Selection.AutoFilter
    End If
    Application.CutCopyMode = False
    Selection.AutoFilter

    Range(Cells(1, colN), Cells(t, colN)).ClearContents
    Cells(1, colN).Select
End Sub
