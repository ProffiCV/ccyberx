Attribute VB_Name = "mdGeral"
Option Explicit

Public goend As Boolean
Private Const AW_CENTER = &H10
Private Const AW_HOR_POSITIVE = &H1
Private Const AW_HOR_NEGATIVE = &H2
Private Const AW_HIDE = &H10000
Private Const AW_VER_NEGATIVE = &H8
Private Const AW_VER_POSITIVE = &H4
Private Const AW_BLEND = &H80000
Private Const AW_ACTIVATE = &H20000

Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwtime As Long, ByVal dwFlags As Long) As Long


Public Const VK_LCONTROL = &HA2
Public Const VK_LMENU = &HA4
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'''''''''''''''''''''''''''''''''''''''''
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function TransparentBlt Lib "msimg32" ( _
  ByVal hdcDest As Long, _
  ByVal nXOriginDest As Long, _
  ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, _
  ByVal hHeightDest As Long, _
  ByVal hdcSrc As Long, _
  ByVal nXOriginSrc As Long, _
  ByVal nYOriginSrc As Long, _
  ByVal nWidthSrc As Long, _
  ByVal nHeightSrc As Long, _
  ByVal crTransparent As Long) As Long


'''''''''''''''''''''''''''''''''''''''''
Public myPrcNet As String

Public noti As New systray

Public debugMode As Boolean
Public gInstalled As Boolean
Public deskState As Long

Public Type DESKTOOP
    progman As Long         'Progman
    systray As Long         'ShellTray_Wnd
    start As Long           'DV2ControlHost
End Type

Public DeskTop As DESKTOOP
 
Public Enum LOCKD 'n bloqueia ambiente de trabalho
Locked = 0
Unlocked = 1
End Enum


Public lockit As Boolean

Public sniffer As New sniffer   'sniffer

Public lastDataSent As String
Public server$
Public port&

Private Const TClass = "Shell_TrayWnd"
Public Const KB = 1024&
Public Const bt = 0&
Public Const MB = KB * KB
Public Const GB = MB * KB

Private Type DIMENSION
    width As Long
    height As Long
End Type

Public TBAR As DIMENSION

Public Declare Sub CopyMemNorm Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Public Const HWND_BOTTOM = 0
Public Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1



Private Const FIXED = SWP_NOSIZE Or SWP_NOMOVE

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)

Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000
'''STYLE

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Const SPIF_UPDATEINIFILE = 1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = 2

' Crie esta parte em um botão


Public cCyberXV2FLG As Long 'startup options for clients computers
Public rolante$

'Public Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type

Public Type BUSPRICE
    netp As Long
    windows As Long
    offset As Double
End Type
Public prc As BUSPRICE

Public Function getConfig(ByRef dta$) As Boolean

Pause 0.06
If Len(dta$) = 16 Then
     cCyberXV2FLG = CLng(Val(Right(dta$, 4)))
'     MsgBox cCyberXV2FLG & "," & dta$
    With prc
        .netp = CLng(Val(Mid(dta$, 1, 4)))
        .windows = CLng(Val(Mid(dta$, 5, 4)))
        .offset = CLng(Val(Mid(dta$, 9, 4)))
    
        myPrcNet = "1 hour " & Format(.windows * 100, "0$00") & " + " & _
        Format(.offset / 2, "0.00") & " MB included. Extra " & Format(.netp * 100, "0$00") & _
        " for each MB!"
        
        .offset = .offset / 2
    End With
    
'    Debug.Print "Net Price " & prc.netp
'    Debug.Print "Windows " & prc.windows
'    Debug.Print "Offset " & prc.offset
    tell "Admin " & _
    prc.netp & ".N." & _
    prc.windows & ".W." & _
    prc.offset & ".O."
    
    SaveSetting App.EXEName, "startup", "task", "" & cCyberXV2FLG
   
    getConfig = True
  Else
    getConfig = False
End If

End Function

''''''''''''''''''''''''
Public Function tell(what As String, Optional ByVal tout As Long = 2)
If what = "" Then Exit Function

If frmlogc.Visible = True Then
If frmlogc.Status <> what Then
    frmlogc.Status = what
    'showTips what, 1
End If
Else
If LoggedInFlag = True And uTime.tm.se > 4 Then showTips what, CLng(tout)

End If

End Function


Public Function test()
Shell App.Path & "\" & App.EXEName & ".exe"
End Function

Public Function DisableMenus()
    With frmAg
        .cmdLogin.Enabled = True
        .cmdLogoff.Enabled = False
        .cmdChangecard.Enabled = False
        .cmdshowBalance.Enabled = False
        .cmdViewCardInfo.Enabled = False
    End With
            
End Function

Public Function EnableMenus()
    With frmAg
        .cmdLogin.Enabled = False
        .cmdLogoff.Enabled = True
        .cmdChangecard.Enabled = True
        .cmdshowBalance.Enabled = True
        .cmdViewCardInfo.Enabled = True
    End With
            
End Function

'translate bytes to multiples
Public Function trasnBytes(ByVal bytes As Long) As String
Dim tmps$


Select Case bytes
Case bt To KB - 1
   tmps$ = Format(bytes, "0.00") & " Bs"
Case KB To MB - 1
   tmps$ = Format(bytes / KB, "0.00") & " KB"
Case MB To GB - 1
   tmps$ = Format(bytes / MB, "0.00") & " MB"
Case Else
   tmps$ = Format(bytes / GB, "0.00") & " GB"
End Select

trasnBytes = tmps$
End Function


Public Function detectTaskBarDimensions()
Dim hisHwnd As Long, trc As RECT, ret&
hisHwnd = FindWindow(TClass, vbNullString)
If hisHwnd <> 0 Then
    ret& = GetWindowRect(hisHwnd, trc)
    If ret& Then
        TBAR.height = trc.Bottom - trc.Top
        TBAR.width = trc.Right - trc.Left
    End If
    
End If

End Function

'posicionar janela...
Public Function topMost(where As Long, frm As Form)
SetWindowPos frm.hwnd, where, 0&, 0&, 0&, 0&, FIXED
End Function


Public Function teste(ByVal codedpass As String)
Dim kpass As String, ich As Integer, ch As Integer
Dim ck(3) As Long
ck(0) = 90
ck(1) = 73
ck(2) = 89
ck(3) = 79
'roberto
If Len(codedpass) = 0 Then Exit Function

For ich = 0 To Len(codedpass) - 1
    kpass = kpass & Chr(Asc(Mid(codedpass, ich + 1, 1)) - ck(ch))
    ch = ch + 1
    If ch = 4 Then ch = 0
Next

'Debug.Print "This Password?? """ & codedpass & """ Decode is  [" & kpass & "]"
End Function


Public Function lockDesktop(ByVal how As LOCKD)
'change wall paper..
deskState = how

If debugMode = True Then
'    Debug.Print "lockDesktop on Debug Mode " & how, deskState
    Exit Function
Exit Function
End If

'Debug.Print "Want; lok"
With DeskTop
    .progman = FindWindow("Progman", vbNullString)
    .start = FindWindow("DV2ControlHost", vbNullString)
    .systray = FindWindow("Shell_TrayWnd", vbNullString)

EnableWindow .progman, how
EnableWindow .systray, how
EnableWindow .start, how

End With

fim:
End Function

Public Function lockInputs(ByVal how As Boolean)

If debugMode = True Then
'    Debug.Print "lockDesktop on Debug Mode " & how, cCyberXV2FLG And &H4, cCyberXV2FLG
    Exit Function
Exit Function
End If



If (cCyberXV2FLG And &H4) = 0 Then GoTo fim

If how = True Then
  SetCursorPos frmlogc.Left, frmlogc.Top

Else

End If
BlockInput how

fim:

End Function

Public Function HideCaption(ByVal hwnd As Long)
Dim oldlong As Long
 oldlong = GetWindowLong(hwnd, GWL_STYLE)
 
 oldlong = (oldlong Or WS_BORDER) - WS_CAPTION
oldlong = SetWindowLong(hwnd, GWL_STYLE, oldlong)
End Function


Public Function changeWallPaper(Optional ByVal reset As Boolean = False)
Dim ret&
Dim fName As String
fName = App.Path & "\back3.bmp"
If debugMode = True Then
'    Debug.Print "changeWallPaper on Debug Mode " & reset, cCyberXV2FLG And &H4, cCyberXV2FLG
    Exit Function
Exit Function
End If

If Dir(fName) <> "" Then
If reset = True Then fName = ""
Call SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, fName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End If



End Function


Public Sub clientOK()
lockDesktop Locked
tell "Enter Code and press [ENTER]..."
 frmlogc.Shape3.BorderColor = vbYellow
frmlogc.Shape2.BorderColor = frmlogc.Shape3.BorderColor
End Sub

Public Function showTips(ByVal msg As String, Optional timeout As Long = 2)
frmttip.lbmsg = msg
 AnimateWindow frmttip.hwnd, 150, AW_BLEND

 frmttip.Refresh
 Pause timeout
 AnimateWindow frmttip.hwnd, 800, AW_BLEND Or AW_HIDE

 Unload frmttip
End Function
