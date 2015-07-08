Attribute VB_Name = "Functions"
Public Function lastRow(Optional sheetName As String, Optional columnLetter As String) As Double
    If IsMissing(sheetName) = True Then
        sheetName = ActiveSheet.Name
    Else
    
    If IsMissing(columnLetter) = True Then
        columnLetter = "A"
    Else
    
    lastRow = Worksheets(sheetName).Range(columnLetter & Rows.Count).End(xlUp).Row
End Function
Public Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
Public Function lastColumn(Optional sheetName As String, Optional rowNumber As Long) As Long
    If IsMissing(sheetName) = True Then
        sheetName = ActiveSheet.Name
    Else
    
    If IsMissing(rowNumber) = True Then
        columnLetter = 1
    Else
    
    lastColumn = Worksheets(sheetName).Cells(1, Columns.Count).End(xlToLeft).Column
End Function


Function GetFolder(strPath As String) As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function
Function G_DISTANCE(Origin As String, Destination As String) As Double
' Requires a reference to Microsoft XML, v6.0
' Draws on the stackoverflow answer at bit.ly/parseXML
Dim myRequest As XMLHTTP60
Dim myDomDoc As DOMDocument60
Dim distanceNode As IXMLDOMNode
    G_DISTANCE = 0
    ' Check and clean inputs
    On Error GoTo exitRoute
    Origin = WorksheetFunction.EncodeURL(Origin)
    Destination = WorksheetFunction.EncodeURL(Destination)
    ' Read the XML data from the Google Maps API
    Set myRequest = New XMLHTTP60
    myRequest.Open "GET", "http://maps.googleapis.com/maps/api/directions/xml?origin=" _
        & Origin & "&destination=" & Destination & "&sensor=false", False
    myRequest.send
    ' Make the XML readable usign XPath
    Set myDomDoc = New DOMDocument60
    myDomDoc.LoadXML myRequest.responseText
    ' Get the distance node value
    Set distanceNode = myDomDoc.SelectSingleNode("//leg/distance/value")
    If Not distanceNode Is Nothing Then G_DISTANCE = distanceNode.Text / 1000
exitRoute:
    ' Tidy up
    Set distanceNode = Nothing
    Set myDomDoc = Nothing
    Set myRequest = Nothing
End Function

Function lat_lon(a_t As String, c_t As String, s_t As String, co_t As String, z_t As String)
Dim sURL As String
Dim BodyTxt As String
Dim apan As String, la_t As String, lo_g As String
Dim oXH As Object
    
    'create web url
    sURL = "http://maps.googleapis.com/maps/api/geocode/xml?address="""
    sURL = sURL & Replace(a_t, " ", "+") & ",+" & Replace(c_t, " ", "+") & ",+" & Replace(s_t, " ", "+") & _
    ",+" & Replace(co_t, " ", "+") & ",+" & ",+" & Replace(z_t, " ", "+") & ",+" & _
    "&sensor=false"""
    ' browse url
    Set oXH = CreateObject("msxml2.xmlhttp")
    With oXH
    .Open "get", sURL, False
    .send
    BodyTxt = .responseText
    End With
    apan = Application.WorksheetFunction.Trim(BodyTxt)
    'Latitude
    apan = Right(apan, Len(apan) - InStr(1, apan, "<lat>") - 4)
    la_t = Left(apan, InStr(1, apan, "</lat>") - 1)
    'Longitude
    apan = Right(apan, Len(apan) - InStr(1, apan, "<lng>") - 4)
    lo_g = Left(apan, InStr(1, apan, "</lng>") - 1)
    lat_lon = "Lat:" & la_t & " Lng:" & lo_g
End Function

Function DepHdg(ByVal lat1 As Double, ByVal lon1 As Double, _
                ByVal lat2 As Double, ByVal lon2 As Double) As Double
    Const pi        As Double = 3.14159265358979
    Const D2R       As Double = pi / 180#
 
    lat1 = D2R * lat1
    lat2 = D2R * lat2
    lon1 = D2R * lon1
    lon2 = D2R * lon2
 
    DepHdg = WorksheetFunction.Atan2(Cos(lat1) * Sin(lat2) - Sin(lat1) * Cos(lat2) * Cos(lon1 - lon2), _
                                     Sin(lon2 - lon1) * Cos(lat2)) / D2R
    If DepHdg < 0 Then DepHdg = DepHdg + 360
End Function
Function ImportVariable(strFile As String) As String
    strFile = "C:\Users\riparsons\Desktop\GS2013\" & strFile
    Open strFile For Input As #1
    Line Input #1, ImportVariable
    Close #1
 
End Function
Function OverWriteVariable(strFile As String, strNewName As String, strRewrite As String) As String
Dim strFinal As String, strInput As String, strOutput As String, strLine As String
Dim LineNum As Long
    
    LineNum = 0
    strFinal = strRewrite & vbCrLf
    
    strInput = "C:\Users\riparsons\Desktop\GS2013\" & strFile
    strOutput = "C:\Users\riparsons\Desktop\GS2013\out\" & strNewName
    
    Open strInput For Input As #1
    While EOF(1) = False
        LineNum = LineNum + 1
        If LineNum > 2 Then
            Line Input #1, strLine
            'Debug.Print strLine
            If Left(strLine, 8) <> "**BEGIN," Then
                strFinal = strFinal + strLine + vbCrLf
            End If
        End If
    Wend
    strFinal = Left(strFinal, Len(strFinal) - 2)
    Close #1
    
    Open strOutput For Output As #1
    Print #1, strFinal
    Close #1
    OverWriteVariable = "done"
End Function

