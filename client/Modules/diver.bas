Attribute VB_Name = "diver"
Option Explicit

'Public Const HWND_BOTTOM = 0
'Public Const HWND_TOPMOST = -1
'Private Const SWP_NOMOVE = &H2
'Private Const SWP_NOSIZE = &H1
'Private Const FIXED = SWP_NOSIZE Or SWP_NOMOVE
'
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'
'Public Function topMost(where As Long, frm As Form)
'SetWindowPos frm.hWnd, where, 0&, 0&, 0&, 0&, FIXED
'End Function


Public Function windowProc(ByVal hwnd&, ByVal msg&, ByVal wparam&, ByVal lparam&) As Long
Debug.Print "They send me " & msg
End Function

