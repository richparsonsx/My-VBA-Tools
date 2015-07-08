Attribute VB_Name = "GlobalFunctions"
'fnLastRow and fnColLtr included here
'
'fnLastRow returns the last row of each row/column/cell
'type =fnLastRow(1, 'Sheet1'!$A$1:$G:$100000) for example
'to find the last row used in that range
'
'fnColLtr takes a number and converts it to a letter for column references
'so =fnColLtr(1) will return A, fnColLtr(20) will return T

Function fnLastRow(ByVal choice As Long, ByVal rng As Range)
' 1 = last row
' 2 = last column
' 3 = last cell
Dim lrw As Long
Dim lcol As Long
    Select Case choice
    Case 1:
        On Error Resume Next
        fnLastRow = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).row
        On Error GoTo 0
    Case 2:
        On Error Resume Next
        fnLastRow = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0
    Case 3:
        On Error Resume Next
        lrw = rng.Find(What:="*", _
                       After:=rng.Cells(1), _
                       Lookat:=xlPart, _
                       LookIn:=xlFormulas, _
                       SearchOrder:=xlByRows, _
                       SearchDirection:=xlPrevious, _
                       MatchCase:=False).row
        On Error GoTo 0
        On Error Resume Next
        lcol = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0
        On Error Resume Next
        fnLastRow = rng.Parent.Cells(lrw, lcol).Address(False, False)
        If Err.Number > 0 Then
            fnLastRow = rng.Cells(1).Address(False, False)
            Err.Clear
        End If
        On Error GoTo 0
    End Select
End Function

Function fnColLtr(iCol As Long) As String
  If iCol > 0 And iCol <= Columns.count Then fnColLtr = Replace(Cells(1, iCol).Address(0, 0), 1, "")
End Function
Function fnColLtr(iCol As Long) As String
If iCol > 0 And iCol <= Columns.count Then fnColLtr = Replace(Cells(1, iCol).Address(0, 0), 1, "")
End Function
