Attribute VB_Name = "Tools"
Private Sub clearBlanks()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Dim rng As Range
Set rng = Selection
For Each r In rng
    If r.Value = "" Then r.ClearContents
    If r.Column = rng.Column Then
        Application.StatusBar = r.Row & " of " & rng.Rows.Count
    End If
Next
Application.StatusBar = False
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
Private Sub SaveAsXLSX()
    FName = ActiveWorkbook.Name
    FPath = ActiveWorkbook.Path
    OfficeSaveAsDialog (FPath & "\" & FName)
End Sub
Private Sub OfficeSaveAsDialog(vInitialFilename As String)
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .InitialView = msoFileDialogViewDetails
        .InitialFileName = vInitialFilename
        If .Show = -1 Then
            .Execute
        Else
        End If
    End With
End Sub
Private Sub myTextToColumns()

Dim r, rng As Range
Set rng = ActiveSheet.UsedRange
rng.Select
Selection.AutoFilter
Selection.NumberFormat = "General"
For Each r In rng.Columns
    Columns(r.Column).Select
    
    Selection.TextToColumns Destination:=Cells(1, r.Column), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Debug.Print r.Column
Next
    
    Range("A1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Cells.EntireColumn.AutoFit
    Range("A2").Select
End Sub
Sub HideBlankWorkbook()
' HideBlankWorkbook Macro
    Windows("PERSONAL.XLSB").Visible = True
    Windows("PERSONAL.XLSB").Visible = False
End Sub
