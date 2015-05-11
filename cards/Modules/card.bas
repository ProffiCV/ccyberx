Attribute VB_Name = "modGeral"
Option Explicit
Public logged As Boolean
Public Const HWND_BOTTOM = 0
Public Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const FIXED = SWP_NOSIZE Or SWP_NOMOVE
'
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Const KB = 1024&
Public Const BT = 0&
Public Const MB = KB * KB
Public Const GB = MB * KB
Public gcmd As String

Private Const WM_SETREDRAW = &HB
Private Const LBS_NOREDRAW = &H4&
Private Const GWL_STYLE = (-16)

Public Declare Sub CopyMemory Lib "kernel32" Alias _
"RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long


Private Declare Function GetDefaultPrinter Lib "Winspool.drv" Alias "GetDefaultPrinterA" (pszBuffer As Any, pcchBuffer As Long) As Long


Private laststyle As Long
Private oldheight As Long
Private oldview As Long

Public Function setRedraw(ByVal hwndi As ListBox, Optional glock = 0)
If glock = 0 Then
laststyle = GetWindowLong(hwndi.hwnd, GWL_STYLE)
oldheight = hwndi.Height
oldview = hwndi.Appearance
    SetWindowLong hwndi.hwnd, GWL_STYLE, LBS_NOREDRAW
Else
    SetWindowLong hwndi.hwnd, GWL_STYLE, laststyle
    SendMessage hwndi.hwnd, WM_SETREDRAW, True, 0
    hwndi.Appearance = oldview
    hwndi.Height = oldheight
 
End If

End Function
Public Function busy()
    frmSup.Enabled = False
    Screen.MousePointer = 11
End Function

Public Function free()
    frmSup.Enabled = True
    Screen.MousePointer = 0
End Function

''''''''''''''GERAL
'translate bytes to multiples
Public Function trasnBytes(ByVal bytes As Long)
Dim tmps$


Select Case bytes
Case BT To KB - 1
   tmps$ = Format(bytes, "0.00") & " Bytes"
Case KB To MB - 1
   tmps$ = Format(bytes \ KB, "0.00") & " KB"
Case MB To GB - 1
   tmps$ = Format(bytes \ MB, "0.00") & " MB"
Case Else
   tmps$ = Format(bytes \ GB, "0.00") & " GB"
End Select

trasnBytes = tmps$
End Function

Public Function Pause(ByVal ms As Double)
Dim init As Double
init = Timer
Do
DoEvents
Loop Until Timer - init >= (ms)
End Function

Public Function getPrinter() As String
Dim buffer$, lb&
buffer = Space(512)
lb = 512
GetDefaultPrinter ByVal buffer, lb
If lb <> 0 Then
getPrinter = Left(buffer, lb - 1)
Else
getPrinter = "None"
End If

End Function

Public Function topMost(where As Long, frm As Form)
SetWindowPos frm.hwnd, where, 0&, 0&, 0&, 0&, FIXED
End Function

