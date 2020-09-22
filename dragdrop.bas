Attribute VB_Name = "Module1"
Private Const WM_DROPFILES = &H233
'&H233 is the windows message id for the drop files message.
'It is the value of the uMsg parameter in the window procedure call.

Private Const GWL_WNDPROC = (-4)
'The index parameter to the SetWindowLong function
'that specifies to change a windows message handler procedure.

Private Declare Sub DragAcceptFiles Lib "shell32.dll" _
(ByVal hwnd As Long, ByVal fAccept As Long)
'DragAcceptFiles enables or disables a form or window to accept files.
'fAccept = 1 Enables.

Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" _
(ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
'DragQueryFile gives the information to us about the dropped file.
'lpStr outputs the filename.

Private Declare Sub DragFinish Lib "shell32.dll" _
(ByVal HDROP As Long)
'This function frees the resources used during the drag operation

Private PrevProc As Long
'Variable to store the address of the default window procedure

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
ByVal msg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long

Private Function HookForm(ByVal hwnd As Long)
    PrevProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
 'Setting our new windowProc function, now all message to window goes through WindowProc.
 'Return value is the address of the previous function. ie,
 'the AddressOf default window proc function
End Function
'Our Custom WindowProc Function
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_DROPFILES Then 'If we have got a drop
        Dropped wParam 'wparam stores the Hdrop handle
    End If
    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
'Call the default window procedure !IMPORTANT
End Function

'Remove our default window procedure.
Private Function UnHookForm(ByVal hwnd As Long)
    If PrevProc <> 0 Then
        SetWindowLong hwnd, GWL_WNDPROC, PrevProc
        PrevProc = 0
    End If
End Function

''' interface api '''
Public Sub EnableDragDrop(ByVal hwnd As Long)
    DragAcceptFiles hwnd, 1
    HookForm (hwnd)
End Sub

Public Sub DisableDragDrop(ByVal hwnd As Long)
    DragAcceptFiles hwnd, 0
    UnHookForm hwnd
End Sub

Public Sub Dropped(ByVal HDROP As Long)
    Dim strFilename As String * 511
    Call DragQueryFile(HDROP, 0, strFilename, 511) 'Get the filename.
    
    '!! replace with your function below ....
    Form1.GotADrop (strFilename)
        Call DragQueryFile(HDROP, 2, strFilename, 511)
End Sub



