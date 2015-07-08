Attribute VB_Name = "ListAllFolders"
Public strFolders As String
Private Sub GetFolderNames()
    Dim olApp As Outlook.Application
    Dim olSession As Outlook.NameSpace
    Dim olStartFolder As Outlook.Folder
    Dim lCountOfFound As Long
 
    lCountOfFound = 0
      
    Set olApp = New Outlook.Application
    Set olSession = olApp.GetNamespace("MAPI")
      
     ' Allow the user to pick the folder in which to start the search.
    Set olStartFolder = olSession.PickFolder
      
     ' Check to make sure user didn't cancel PickFolder dialog.
    If Not (olStartFolder Is Nothing) Then
         ' Start the search process.
        processFolder olStartFolder
    End If
     
' Create a new mail message with the folder list inserted
'Set ListFolders = Application.CreateItem(olMailItem)
'ListFolders.Body = strFolders
'  ListFolders.Display
      
' clear the string so you can run it on another folder
  strFolders = ""
End Sub
Private Sub processFolder(CurrentFolder As Outlook.Folder)
    Dim i As Long
    Dim olNewFolder As Outlook.Folder
    Dim olTempFolder As Outlook.Folder
    Dim olTempFolderPath As String
     ' Loop through the items in the current folder.
    For i = CurrentFolder.Folders.count To 1 Step -1
        Set olTempFolder = CurrentFolder.Folders(i)
         'olTempFolderPath = olTempFolder.folderPath
         'prints the folder path and name in the VB Editor's Immediate window
         'Debug.Print olTempFolderPath
         'prints the folder name only
         'Debug.Print olTempFolder
         'create a string with the folder names.  use olTempFolder if you want foldernames only
         'strFolders = strFolders & vbCrLf & olTempFolder
         ThisOutlookSession.TestAgingProps olTempFolder
        lCountOfFound = lCountOfFound + 1
          
    Next
     ' Loop through and search each subfolder of the current folder.
    For Each olNewFolder In CurrentFolder.Folders
          
         'Don't need to process the Deleted Items folder
        If olNewFolder.Name <> "Deleted Items" Then
            processFolder olNewFolder
        End If
          
    Next
      
End Sub
