VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_ItemSend(ByVal item As Object, Cancel As Boolean)
    RunAllInboxRules
    cleanDeleted
End Sub

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    RunAllInboxRules
    cleanDeleted
End Sub
Private Sub RunAllInboxRules()
Dim st As Outlook.Store
Dim myRules As Outlook.Rules
Dim rl As Outlook.Rule
Dim count As Integer
Dim ruleList As String
' get default store (where rules live)
Set st = Application.session.DefaultStore
' get rules
If Application.session.Offline = True Then Exit Sub
Set myRules = st.GetRules
' iterate all the rules
For Each rl In myRules
    ' determine if it's an Inbox rule
    If rl.RuleType = olRuleReceive Then
        If rl.Name <> "unfiled" Then
            ' if so, run it
            
            On Error Resume Next
            rl.Execute ShowProgress:=False, RuleExecuteOption:=1
            count = count + 1
            ruleList = ruleList & vbCrLf & rl.Name
        End If
    End If
Next
For Each rl In myRules
    ' determine if it's an Inbox rule
    If rl.RuleType = olRuleReceive Then
        If rl.Name = "unfiled" Then
            ' if so, run it
            rl.Execute ShowProgress:=False, RuleExecuteOption:=1
            count = count + 1
            ruleList = ruleList & vbCrLf & rl.Name
        End If
    End If
Next
Set rl = Nothing
Set st = Nothing
Set myRules = Nothing
End Sub
Private Sub cleanDeleted()
Dim olApp As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim oFolder As Outlook.MAPIFolder
Dim olFolder As Outlook.MAPIFolder
Dim oMsg As Object

    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    Set olFolder = objNS.GetDefaultFolder(olFolderDeletedItems)
    If olFolder.Items.count = olFolder.UnReadItemCount Then Exit Sub
    For Each oFolder In olFolder.Folders
        If oFolder.Name = "Read" Then Exit For
    Next
    
    For Each oMsg In olFolder.Items
        If oMsg.UnRead = False Then
            oMsg.Move oFolder
        End If
    Next
    
    Set olApp = Nothing
    Set objNS = Nothing
    Set oFolder = Nothing
    Set olFolder = Nothing
    Set oMsg = Nothing
End Sub

Public Function MyDocsPath() As String
    MyDocsPath = Environ$("USERPROFILE") & "\" & "My Documents\"
End Function

Private Sub GetALLEmailAddresses()
Dim objFolder As MAPIFolder
Dim strEmail As String
Dim strEmails As String
Dim objDic As Object
Dim objItem As Object
Dim objFSO As Object
Dim objTF As Object
Dim folderPath As String
Dim mailCount As Long

folderPath = MyDocsPath()
mailCount = 0

Set objDic = CreateObject("scripting.dictionary")
Set objFSO = CreateObject("scripting.filesystemobject")
Set objFolder = Application.GetNamespace("Mapi").PickFolder
folderPath = MyDocsPath() & "emails - " & objFolder.Name & ".csv"
If objFolder.Name = "Sent Items" Then
    Set objTF = objFSO.createtextfile(folderPath, 2)
    For Each objItem In objFolder.Items
        If objItem.Class = olMail Then
            mailCount = mailCount + 1
            totalCount = objItem.Recipients.count
            For i = 1 To totalCount
                If InStr(1, objItem.Recipients.item(i).Address, "@") > 1 Then
                    strEmail = objItem.Recipients.item(i).Address
                    If Not objDic.Exists(strEmail) Then
                        objTF.writeline strEmail
                        objDic.Add strEmail, ""
                    End If
                End If
            Next
        End If
    Next
Else
    Dim check1 As String
    check1 = InputBox("Input sender or body", "sender or body")
    If check1 = "sender" Then
        Set objTF = objFSO.createtextfile(folderPath, 2)
        For Each objItem In objFolder.Items
            If objItem.Class = olMail Then
                mailCount = mailCount + 1
                If InStr(1, objItem.sender.Address, "@") > 1 Then
                    strEmail = objItem.sender.Address
                    If Not objDic.Exists(strEmail) Then
                        objTF.writeline strEmail
                        objDic.Add strEmail, ""
                    End If
                End If
            End If
        Next
    Else
    Set objTF = objFSO.createtextfile(folderPath, 2)
        For Each objItem In objFolder.Items
            If objItem.Class = olMail Then
                mailCount = mailCount + 1
                If InStr(1, objItem.Body, "@") > 1 Then
                    strEmail = objItem.Body
                    If Not objDic.Exists(strEmail) Then
                        objTF.writeline strEmail
                        objDic.Add strEmail, ""
                    End If
                End If
            End If
        Next
    End If
End If
MsgBox mailCount & " total emails in folder " & objFolder.Name
objTF.Close
End Sub

 
Function getAgingProperties(oFolder As Outlook.Folder) As String
'6 MAPI properties for aging items in a folder
Const PR_AGING_AGE_FOLDER = "http://schemas.microsoft.com/mapi/proptag/0x6857000B"
Const PR_AGING_DELETE_ITEMS = "http://schemas.microsoft.com/mapi/proptag/0x6855000B"
Const PR_AGING_FILE_NAME_AFTER9 = "http://schemas.microsoft.com/mapi/proptag/0x6859001E"
Const PR_AGING_GRANULARITY = "http://schemas.microsoft.com/mapi/proptag/0x36EE0003"
Const PR_AGING_PERIOD = "http://schemas.microsoft.com/mapi/proptag/0x36EC0003"
Const PR_AGING_DEFAULT = "http://schemas.microsoft.com/mapi/proptag/0x685E0003"
Dim oStorage As StorageItem
Dim oPA As PropertyAccessor
Dim strAge, strGran, strDel, strPer, strNam, strDef As String

On Error GoTo Aging_ErrTrap
If oFolder Is Nothing Then
    getAgingProperties = oFolder.Name & "; Nothing"
    Exit Function
End If

    Set oStorage = oFolder.GetStorage("IPC.MS.Outlook.AgingProperties", olIdentifyByMessageClass)
    Set oPA = oStorage.PropertyAccessor
    strAge = oPA.GetProperty(PR_AGING_AGE_FOLDER)
    strGran = oPA.GetProperty(PR_AGING_GRANULARITY)
    strDel = oPA.GetProperty(PR_AGING_DELETE_ITEMS)
    strPer = oPA.GetProperty(PR_AGING_PERIOD)
    strNam = oPA.GetProperty(PR_AGING_FILE_NAME_AFTER9)
    strDef = oPA.GetProperty(PR_AGING_DEFAULT)
    
    getAgingProperties = strAge & "; " & strGran & "; " & strDel & "; " & strPer & "; " & strNam & vbCrLf
    Exit Function

Aging_ErrTrap:
'Debug.Print Err.Number, Err.Description
getAgingProperties = ""

End Function

Private Sub TestAgingProps()
    Dim objNS As NameSpace
    Dim objFolder As Folder
    Dim olFolder As Folder
    
    Set objNS = Application.GetNamespace("MAPI")
    Set objFolder = objNS.PickFolder
    
    Debug.Print objFolder.Name & "; " & getAgingProperties(objFolder)
    objFolder.Delete
'    For Each olFolder In objFolder.Folders
'        Debug.Print olFolder.Name & "; " & getAgingProperties(objFolder)
'    Next
End Sub
