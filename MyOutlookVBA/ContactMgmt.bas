Attribute VB_Name = "ContactMgmt"
Private Sub ChangeEmailDisplayName()
    Dim objOL As Outlook.Application
    Dim objNS As Outlook.NameSpace
    Dim objContact As Outlook.ContactItem
    Dim objItems As Outlook.Items
    Dim objContactsFolder As Outlook.MAPIFolder
    Dim obj As Object
    Dim strName As String
    Dim strLastName As String
    Dim strFileAs As String
 
    On Error Resume Next
 
    Set objOL = CreateObject("Outlook.Application")
    Set objNS = objOL.GetNamespace("MAPI")
    Set objContactsFolder = objNS.GetDefaultFolder(olFolderContacts)
    Set objItems = objContactsFolder.Items

    For Each obj In objItems
        'Test for contact and not distribution list
        If obj.Class = olContact Then
            Set objContact = obj
            With objContact
            ' If .LastNameAndFirstName <> "" Then
                ' Debug.Print vbCrLf & "File As: " & .LastNameAndFirstName
                ' '.FileAs = .LastNameAndFirstName
                
                ' If .Email1Address <> "" Then
                    ' 'Lastname, Firstname format
                    ' strFileAs = .LastNameAndFirstName & " (" & .Email1Address & ")"
                    ' Debug.Print "Email 1: " & strFileAs
                    
                    ' .Email1DisplayName = strFileAs
                ' End If
                ' If .Email2Address <> "" Then
                    ' 'Lastname, Firstname format
                    ' strFileAs = .LastNameAndFirstName & " (" & .Email2Address & ")"
                    ' Debug.Print "Email 2: " & strFileAs
                    ' .Email2DisplayName = strFileAs
                ' End If
                
                ' If .Email3Address <> "" Then
                    ' 'Lastname, Firstname format
                    ' strFileAs = .LastNameAndFirstName & " (" & .Email3Address & ")"
                    ' Debug.Print "Email 3: " & strFileAs
                    ' .Email3DisplayName = strFileAs
                ' End If
                ' .Save
            ' End If
            End With
            
        End If
        Err.Clear
    Next
    Close #2
    Set objOL = Nothing
    Set objNS = Nothing
    Set obj = Nothing
    Set objContact = Nothing
    Set objItems = Nothing
    Set objContactsFolder = Nothing
End Sub
Private Sub ChangeEmailDisplayName2()
    Dim objOL As Outlook.Application
    Dim objNS As Outlook.NameSpace
    Dim objContact As Outlook.ContactItem
    Dim objItems As Outlook.Items
    Dim objContactsFolder As Outlook.MAPIFolder
    Dim obj As Object
    Dim strName As String
    Dim strLastName As String
    Dim strFileAs As String
    Dim filepath As String
    Dim currEmailDisplay As String
    Dim newEmailDisplay As String
 
    On Error Resume Next
 
    Set objOL = CreateObject("Outlook.Application")
    Set objNS = objOL.GetNamespace("MAPI")
    Set objContactsFolder = objNS.GetDefaultFolder(olFolderContacts)
    Set objItems = objContactsFolder.Items
    filepath = "c:\users\riparsons\Desktop\file.txt"
    Open filepath For Output As #2
    For Each obj In objItems
        'Test for contact and not distribution list
        If obj.Class = olContact Then
            Set objContact = obj
            With objContact
                If .FileAs <> "" And .Email1Address <> "" Then
                    currEmailDisplay = .Email1DisplayName
                    newEmailDisplay = .FileAs & " (" & .Email1Address & ")"
                    If currEmailDisplay <> newEmailDisplay Then .Email1DisplayName = newEmailDisplay
                    currEmailDisplay = ""
                    newEmailDisplay = ""
                End If
                If .FileAs <> "" And .Email2Address <> "" Then
                    currEmailDisplay = .Email2DisplayName
                    newEmailDisplay = .FileAs & " (" & .Email2Address & ")"
                    If currEmailDisplay <> newEmailDisplay Then .Email2DisplayName = newEmailDisplay
                    currEmailDisplay = ""
                    newEmailDisplay = ""
                End If
                If .FileAs <> "" And .Email3Address <> "" Then
                    currEmailDisplay = .Email3DisplayName
                    newEmailDisplay = .FileAs & " (" & .Email3Address & ")"
                    If currEmailDisplay <> newEmailDisplay Then .Email3DisplayName = newEmailDisplay
                    currEmailDisplay = ""
                    newEmailDisplay = ""
                End If
                .Save
            End With
            
        End If
        Err.Clear
    Next
    Close #2
    Set objOL = Nothing
    Set objNS = Nothing
    Set obj = Nothing
    Set objContact = Nothing
    Set objItems = Nothing
    Set objContactsFolder = Nothing
End Sub
Private Sub ChangeEmailString()
    Dim objOL As Outlook.Application
    Dim objNS As Outlook.NameSpace
    Dim objContact As Outlook.ContactItem
    Dim objItems As Outlook.Items
    Dim objContactsFolder As Outlook.MAPIFolder
    Dim obj As Object
    Dim strFirstName As String
    Dim strLastName As String
    Dim strEmailAs As String
    Dim strEmailAs2 As String
 
    On Error Resume Next
 
    Set objOL = CreateObject("Outlook.Application")
    Set objNS = objOL.GetNamespace("MAPI")
    Set objContactsFolder = objNS.GetDefaultFolder(olFolderContacts)
    Set objItems = objContactsFolder.Items
 
    For Each obj In objItems
        'Test for contact and not distribution list
        If obj.Class = olContact Then
            Set objContact = obj
            With objContact
                If Left(.Email1Address, 2) = "/o" Then
                    'Lastname, Firstname format
                    strEmailAs2 = ResolveDisplayNameToSMTP(.Email1Address)
                    If Len(strEmailAs2) < 2 Then strEmailAs2 = "FIX ME"
                    strEmailAs = .LastNameAndFirstName & " (" & strEmailAs2 & ")"
                    .Email1DisplayName = strEmailAs
                    .Email1Address = strEmailAs2
                    .Save
                End If
                If Left(.Email2Address, 2) = "/o" Then
                    strEmailAs2 = ResolveDisplayNameToSMTP(.Email2Address)
                    If Len(strEmailAs2) < 2 Then strEmailAs2 = "FIX ME"
                    strEmailAs = .LastNameAndFirstName & " (" & strEmailAs2 & ")"
                    Debug.Print strEmailAs
                    .Email2DisplayName = strEmailAs
                    .Email2Address = strEmailAs2
                    .Save
                End If
                
                If Left(.Email3Address, 2) = "/o" Then
                    strEmailAs2 = ResolveDisplayNameToSMTP(.Email3Address)
                    If Len(strEmailAs2) < 2 Then strEmailAs2 = "FIX ME"
                    strEmailAs = .LastNameAndFirstName & " (" & strEmailAs2 & ")"
                    Debug.Print strEmailAs
                    .Email3DisplayName = strEmailAs
                    .Email3Address = strEmailAs2
                    .Save
                End If
            End With
        End If
        Err.Clear
    Next
 
    Set objOL = Nothing
    Set objNS = Nothing
    Set obj = Nothing
    Set objContact = Nothing
    Set objItems = Nothing
    Set objContactsFolder = Nothing
End Sub


 
Function ResolveDisplayNameToSMTP(sFromName)
  Dim oRecip As Outlook.Recipient
  Dim oEU As Outlook.exchangeUser
  Dim oEDL As Outlook.ExchangeDistributionList
  
  Set oRecip = Application.session.CreateRecipient(sFromName)
  oRecip.Resolve
  If oRecip.Resolved Then
    Select Case oRecip.AddressEntry.AddressEntryUserType
      Case OlAddressEntryUserType.olExchangeUserAddressEntry
        Set oEU = oRecip.AddressEntry.GetExchangeUser
        If Not (oEU Is Nothing) Then
          ResolveDisplayNameToSMTP = oEU.PrimarySmtpAddress
        End If
      Case OlAddressEntryUserType.olOutlookContactAddressEntry
        Set oEU = oRecip.AddressEntry.GetExchangeUser
        If Not (oEU Is Nothing) Then
          ResolveDisplayNameToSMTP = oEU.PrimarySmtpAddress
        End If
      Case OlAddressEntryUserType.olExchangeDistributionListAddressEntry
        Set oEDL = oRecip.AddressEntry.GetExchangeDistributionList
        If Not (oEDL Is Nothing) Then
          ResolveDisplayNameToSMTP = oEU.PrimarySmtpAddress
        End If
    End Select
  End If
End Function  ' ResolveDisplayNameToSMTP
