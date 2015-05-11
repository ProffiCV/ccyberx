Attribute VB_Name = "MODmSN"
Option Explicit

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function BlockInput Lib "user32" (ByVal lok As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function ShowWindowAsync Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Const WM_MOUSEMOVE = &H200
Public Const PROCESS_TERMINATE = &H1
Public Const SC_CLOSE = &HF060&
Public Const WM_SYSCOMMAND = &H112

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal lparam As Long) As Long
Private myhw As Long


Public Function logOffMSN()

If debugMode = True Then
'    Debug.Print "logOffMSN On debugmode "
    Exit Function
End If


Dim buffer$
buffer = Space(255)
Dim pid As Long
Const MSNmain = "MSNMSBLGeneric"
Const MSNtray = "MSNHiddenWindowClass"
Const MSNmsgr = "MSBLWindowClass"
Dim msnhnd&


'msnhnd = FindWindow("IME", vbNullString)
'GetWindowThreadProcessId msnhnd, pid
'CloseThis msnhnd, pid
tell "Refreshing Icons..."

msnhnd = FindWindow("ATL:0084B4E8", vbNullString)
GetWindowThreadProcessId msnhnd, pid
hardKill msnhnd

msnhnd = FindWindow("MSGRIMEWINDOWCLASS", vbNullString)
GetWindowThreadProcessId msnhnd, pid
hardKill msnhnd

'''''''''''''''''''''''''''''
msnhnd = FindWindow(MSNtray, vbNullString)
GetWindowThreadProcessId msnhnd, pid
hardKill msnhnd

''''''''''''''''''''''''''''''''''''''''''
msnhnd = FindWindow(MSNmain, vbNullString)
GetWindowThreadProcessId msnhnd, pid
hardKill msnhnd

msnhnd = FindWindow(MSNmsgr, vbNullString)
GetWindowThreadProcessId msnhnd, pid
hardKill msnhnd

'''''''''''''''''''''''
'Unload frmm
'Shell Environ("PROGRAMFILES") & "\msn messenger\msnmsgr.exe", vbHide

Pause 0.4
msnhnd = FindWindow(MSNmsgr, vbNullString)
GetWindowThreadProcessId msnhnd, pid
CloseThis msnhnd, -1

msnhnd = FindWindow("Shell_TrayWnd", vbNullString)
msnhnd = FindWindowEx(msnhnd, 0, "TrayNotifyWnd", vbNullString)
mouseMove msnhnd
clientOK
End Function


Private Function mouseMove(ByVal hwnd As Long) As Long
Dim rc As RECT, ret
ret = GetWindowRect(hwnd, rc)

Dim itr&

Dim pt As POINTAPI
ret = GetCursorPos(pt)
For itr& = rc.Left To rc.Right Step 8
lockDesktop Unlocked
    SetCursorPos itr&, rc.Top + 10
    Pause 0.06
Next
Pause 0.6
SetCursorPos pt.X, pt.Y
lockDesktop Locked

End Function

Public Function findWindowByPid(pid As Long)
myhw = -1
EnumWindows AddressOf EnumWindowsProc, pid
findWindowByPid = myhw
End Function

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lparam As Long) As Boolean
Dim pid As Long

GetWindowThreadProcessId hwnd, pid
Dim bufg As String
bufg = Space(255)
GetClassName hwnd, bufg, 255
Debug.Print bufg
If pid = lparam Then
myhw = hwnd
EnumWindowsProc = False
End If

EnumWindowsProc = True

End Function
