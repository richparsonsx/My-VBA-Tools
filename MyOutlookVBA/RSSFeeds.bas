Attribute VB_Name = "RSSFeeds"
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Sub RunRSS()
Dim oApp As New Outlook.Application
Dim oNS As Outlook.NameSpace
Dim olParentFolder As Outlook.MAPIFolder
Dim f As Outlook.Folder, oFld As Outlook.Folder
Dim oPosts As Outlook.Items
Dim oPostItem As Outlook.PostItem
Dim oProp As Outlook.PropertyPage
Dim cl, c2, c3, count, m, ret As Long
Dim sSubject As String, sBody As String, displayMsg As String, fileName As String, fullAddress As String, strFolderpath As String
Dim WinHttpReq As Object

'On Error GoTo Err_OL
Set oNS = oApp.GetNamespace("MAPI")
Set olParentFolder = oNS.GetDefaultFolder(olFolderInbox)
For Each f In olParentFolder.Folders
    If f.Name = ".ChangeDetection feed" Then
        Set oFld = f
    End If
Next
strFolderpath = "\\deloitteteams.deloittenet.com@SSL\DavWWWRoot\sites\PfxTaxSaaS\IASTest\Shared Documents"
Set oPosts = oFld.Items
m = 1
count = oPosts.count
Do Until m > count
    Set oPostItem = oPosts.Item(m)
    c1 = 1
    sBody = oPosts.Item(m).Body
    sSubject = "http://www.irs.gov" & Right(oPosts.Item(m).Subject, Len(oPosts.Item(m).Subject) - 9) & "/"
    If oPostItem.UnRead = True Then
        Do While InStr(c1, sBody, ".pdf") <> 0
            c3 = InStr(c1, sBody, ".pdf")
            For i = c3 To 1 Step -1
                If Right(Left(sBody, i), 1) = " " Then
                    c2 = i
                    i = 1
                End If
            Next
            fileName = Left(Right(sBody, Len(sBody) - c2), c3 + 3 - c2)
            fullAddress = sSubject & fileName
            ret = URLDownloadToFile(0, fullAddress, strFolderpath & "\IRS-Updates\" & fileName, 0, 0)
            c1 = c3 + 1
        Loop
        Do While InStr(c1, sBody, ".xls") <> 0
            c3 = InStr(c1, sBody, ".xls")
            For i = c3 To 1 Step -1
                If Right(Left(sBody, i), 1) = " " Then
                    c2 = i
                    i = 1
                End If
            Next
            fileName = Left(Right(sBody, Len(sBody) - c2), c3 + 3 - c2)
            fullAddress = sSubject & fileName
            ret = URLDownloadToFile(0, fullAddress, strFolderpath & "\IRS-Updates\" & fileName, 0, 0)
            'downloadString = downloadString & sSubject & Left(Right(sBody, Len(sBody) - c2), c3 + 3 - c2) & vbCrLf
            c1 = c3 + 1
        Loop
        Do While InStr(c1, sBody, ".zip") <> 0
            c3 = InStr(c1, sBody, ".zip")
            For i = c3 To 1 Step -1
                If Right(Left(sBody, i), 1) = " " Then
                    c2 = i
                    i = 1
                End If
            Next
            fileName = Left(Right(sBody, Len(sBody) - c2), c3 + 3 - c2)
            fullAddress = sSubject & fileName
            ret = URLDownloadToFile(0, fullAddress, strFolderpath & "\IRS-Updates\" & fileName, 0, 0)
            'downloadString = downloadString & sSubject & Left(Right(sBody, Len(sBody) - c2), c3 + 3 - c2) & vbCrLf
            c1 = c3 + 1
        Loop
        oPostItem.UnRead = False
    End If
    m = m + 1
Loop
End Sub
