Attribute VB_Name = "Usability"
Option Explicit

Private Sub External()
    Set newItem = Application.CreateItem(olMailItem)
    newItem.Display
    Set newItem = Nothing
End Sub
Private Sub Internal()
    Set newItem = Application.CreateItemFromTemplate("C:\Users\riparsons\Documents\ShareFile\My Files & Folders\Templates\Outlook\InternalMessage.oft")
    newItem.Display
    Call DeleteSig(newItem)
    Set newItem = Nothing
End Sub
Private Sub DeleteSig(msg As Outlook.MailItem)
    Dim objDoc As Word.Document
    Dim objBkm As Word.Bookmark
    On Error Resume Next
    Set objDoc = msg.GetInspector.WordEditor
    Set objBkm = objDoc.Bookmarks("_MailAutoSig")
    If Not objBkm Is Nothing Then
        objBkm.Select
        objDoc.Windows(1).Selection.Delete
    End If
    Set objDoc = Nothing
    Set objBkm = Nothing
End Sub
Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
            
    Set objApp = Application
    On Error Resume Next
    Select Case typeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
        
    Set objApp = Nothing
End Function
