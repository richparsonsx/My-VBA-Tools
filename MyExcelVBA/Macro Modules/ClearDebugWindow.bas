Attribute VB_Name = "ClearDebugWindow"
Private Declare Function GetWindow _
Lib "user32" ( _
ByVal hWnd As Long, _
ByVal wCmd As Long) As Long
Private Declare Function FindWindow _
Lib "user32" Alias "FindWindowA" ( _
ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx _
Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long
Private Declare Function GetKeyboardState _
Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState _
Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function PostMessage _
Lib "user32" Alias "PostMessageA" ( _
ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long _
) As Long


Private Const WM_KEYDOWN As Long = &H100
Private Const KEYSTATE_KEYDOWN As Long = &H80


Private savState(0 To 255) As Byte

Public Sub ClearImmediateWindow()

Dim hPane As Long
Dim tmpState(0 To 255) As Byte


hPane = GetImmHandle
If hPane = 0 Then MsgBox "Immediate Window not found."
If hPane < 1 Then Exit Sub


'Save the keyboardstate
GetKeyboardState savState(0)


'Sink the CTRL (note we work with the empty tmpState)
tmpState(vbKeyControl) = KEYSTATE_KEYDOWN
SetKeyboardState tmpState(0)
'Send CTRL+End
PostMessage hPane, WM_KEYDOWN, vbKeyEnd, 0&
'Sink the SHIFT
tmpState(vbKeyShift) = KEYSTATE_KEYDOWN
SetKeyboardState tmpState(0)
'Send CTRLSHIFT+Home and CTRLSHIFT+BackSpace
PostMessage hPane, WM_KEYDOWN, vbKeyHome, 0&
PostMessage hPane, WM_KEYDOWN, vbKeyBack, 0&


'Schedule cleanup code to run
DoCleanUp


End Sub


Sub DoCleanUp()
' Restore keyboard state
SetKeyboardState savState(0)
End Sub


Function GetImmHandle() As Long
'This function finds the Immediate Pane and returns a handle.
'Docked or MDI, Desked or Floating, Visible or Hidden


Dim oWnd As Object, bDock As Boolean, bShow As Boolean
Dim sMain$, sDock$, sPane$
Dim lMain&, lDock&, lPane&


On Error Resume Next
sMain = Application.VBE.MainWindow.Caption
If Err <> 0 Then
MsgBox "No Access to Visual Basic Project"
GetImmHandle = -1
Exit Function
' Excel2003: Registry Editor (Regedit.exe)
'    HKLM\SOFTWARE\Microsoft\Office\11.0\Excel\Security
'    Change or add a DWORD called 'AccessVBOM', set to 1
' Excel2002: Tools/Macro/Security
'    Tab 'Trusted Sources', Check 'Trust access..'
End If


For Each oWnd In Application.VBE.Windows
If oWnd.Type = 5 Then
bShow = oWnd.Visible
sPane = oWnd.Caption
If Not oWnd.LinkedWindowFrame Is Nothing Then
bDock = True
sDock = oWnd.LinkedWindowFrame.Caption
End If
Exit For
End If
Next
lMain = FindWindow("wndclass_desked_gsk", sMain)
If bDock Then
'Docked within the VBE
lPane = FindWindowEx(lMain, 0&, "VbaWindow", sPane)
If lPane = 0 Then
'Floating Pane.. which MAY have it's own frame
lDock = FindWindow("VbFloatingPalette", vbNullString)
lPane = FindWindowEx(lDock, 0&, "VbaWindow", sPane)
While lDock > 0 And lPane = 0
lDock = GetWindow(lDock, 2) 'GW_HWNDNEXT = 2
lPane = FindWindowEx(lDock, 0&, "VbaWindow", sPane)
Wend
End If
ElseIf bShow Then
lDock = FindWindowEx(lMain, 0&, "MDIClient", _
vbNullString)
lDock = FindWindowEx(lDock, 0&, "DockingView", _
vbNullString)
lPane = FindWindowEx(lDock, 0&, "VbaWindow", sPane)
Else
lPane = FindWindowEx(lMain, 0&, "VbaWindow", sPane)
End If


GetImmHandle = lPane


End Function

