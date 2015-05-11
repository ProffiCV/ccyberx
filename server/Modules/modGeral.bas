Attribute VB_Name = "modg"
Option Explicit

Public Enum FLGSHUT  'used to control windows shutdown event on client computer
    EWX_FORCE = &H4
    EWX_LOGOFF = &H0 Or EWX_FORCE
    EWX_SHUTDOWN = &H1 Or EWX_FORCE
    EWX_REBOOT = &H2 Or EWX_FORCE
    EWX_POWEROFF = &H8 Or EWX_FORCE
    EWX_FORCEIFHUNG = &H10
End Enum


Public operation As Long 'loading or unloading...?

Public Const HWND_BOTTOM = 0
Public Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const FIXED = SWP_NOSIZE Or SWP_NOMOVE
'
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)

Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000
'''STYLE

'ººººººººººººººº

Public cCyberXV2FLG As Long 'startup options for clients computers

Public targetP As Long
Public gcmd As String 'permission to continue
Public myInt As Long
Private tlab(13) As Label

Public Const KB = 1024&
Public Const BT = 0&
Public Const MB = KB * KB
Public Const GB = MB * KB

Public Declare Sub CopyMemory Lib "kernel32" Alias _
"RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public tmOutInterval As Long 'tempo de espere antes de fechar....
Public Const NL = vbNewLine
Private hMutex As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



''''''''''Animation
Public Const AW_CENTER = &H10
Public Const AW_HOR_POSITIVE = &H1
Public Const AW_HOR_NEGATIVE = &H2
Public Const AW_HIDE = &H10000
Public Const AW_VER_NEGATIVE = &H8
Public Const AW_VER_POSITIVE = &H4
Public Const AW_BLEND = &H80000
Public Const AW_ACTIVATE = &H20000

Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwtime As Long, ByVal dwFlags As Long) As Long


'''''''''''DETECT other instances
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Any, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Public Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long

Public myMutex As Long

Public Type PRICES
    Pnet As Long
    Pwindows As Long
    offSet As Long
End Type

Public prcSetup As PRICES

'Dados do PC e do Cliente
Public Type PCINFO
    Index As Long
    dispInfo As Integer
    pcName As String
    clientID As String
    code As String
    '''''''''''''''''''''''''''''''''
    state As String
    '''''''''''''''''''''''''''''''''
    login As String
    logoff As String
    ''''''''''''''''''''''
    pcuTime As String
    pcuPrice As String
    ''''''''''''''''''''''''
    netNow As String
    netPrice As String
    netTotal As String
    ''''''''''''''''
    balTotal As String
    balUsed As String
    balRemain As String
End Type
Public pci(8) As PCINFO

''dados da utilização.. temp
Public Type TDET
    huso As String * 12  '00:00:00 + 0000 'price time win
    totu As String * 6   '000$00
    netp As String * 6  '000$00
    id As String * 3        '[000]
    code As String * 19     '[0000-0000-0000-0000]
    date As String * 6      '[000000]
    life As String * 2      '[00]
    flag As String * 1      '[X]R reached , F false T used , N not in use, invalid details
    tbal As String * 4      '[0000]
    tusd As String * 4      '[0000]
    netc As String * 12  '1024 KB
    netn As String * 12
    TLogoff As String * 1   '1 or 0
End Type
Public tmde(8) As TDET

Public Function setPrices(net As Long, windows As Long, offdown As Long, tout As Long)
If tout <> 0 Then tmOutInterval = tout

If net = 0 And windows = 0 Then
    'get clients config
    cCyberXV2FLG = GetSetting(App.EXEName, "Config", "CC", 0)
    
    prcSetup.Pwindows = 100 + 10 * CLng("&h" & GetSetting(App.EXEName, "Config", "P0", 0))
    prcSetup.Pnet = 10 + 5 * CLng("&h" & GetSetting(App.EXEName, "Config", "P1", 0))
    prcSetup.offSet = CLng("&h" & GetSetting(App.EXEName, "Config", "P3", 0))
    tmOutInterval = CLng("&h" & GetSetting(App.EXEName, "Config", "P2", 0))
    
    tmOutInterval = (2 * tmOutInterval + 2) * 60
    frmsv2.tmResetCon.Interval = 1000
    frmsv2.tmResetCon.Enabled = True
    
Else

prcSetup.Pnet = 10 + 5 * net
prcSetup.Pwindows = 100 + 10 * windows
prcSetup.offSet = Val(IIf(offdown = -1, 0, offdown))

    tmOutInterval = (2 * tout + 2) * 60
    frmsv2.tmResetCon.Interval = 1000
    frmsv2.tmResetCon.Enabled = True
End If
'----------------------------

Call frmsv2.displayPrices
End Function

Public Function getPrices() As PRICES
getPrices = prcSetup
'----------------------------
End Function


                          

Public Function tell(ByRef what$)
'mostra mensagens para o utilizador
    If frmsv2.msg = what$ Then Exit Function
    
    If InStr(1, what$, "Updating") = 0 Then
    frmsv2.msg.ForeColor = vbWhite
    Else
    frmsv2.msg.ForeColor = vbRed
    End If
    
    frmsv2.msg = what$
    frmsv2.tmTime.Enabled = False
    frmsv2.tmTime.Enabled = True
    
End Function



Public Function DetectOtherInstances(ByVal who As String) As Boolean
Dim mylong&
   
 If FindWindow(vbNullString, "sServerXV2.") = 0 Then
 DetectOtherInstances = False
   mylong& = frmsv2.hwnd
   SaveSetting "EdsonSoft", "sServerXV2", "hwnd", "" & mylong
Else
mylong& = CLng(GetSetting("EdsonSoft", "sServerXV2", "hwnd"))
frmStart.lbpg = "sServerXV2 is running..."
Pause 2
DetectOtherInstances = True

End If

  
End Function

'translate bytes to multiples
Public Function trasnBytes(ByVal bytes As Long)
Dim tmps$


Select Case bytes
Case BT To KB - 1
   tmps$ = Format(bytes, "0.00") & " Bytes"
Case KB To MB - 1
   tmps$ = Format(bytes / KB, "0.00") & " KB"
Case MB To GB - 1
   tmps$ = Format(bytes / MB, "0.00") & " MB"
Case Else
   tmps$ = Format(bytes / GB, "0.00") & " GB"
End Select

trasnBytes = tmps$
End Function

'no Auto sleep
Public Function resetSensor()
myInt = 0
End Function
'adiciona dias de vida ao cartão...
Public Function addLife() As Long
    Dim dia As String
    dia = GetSetting(App.EXEName, "config", "day", -1)
    If dia <> Format(Day(date), "dd") Then
        dia = Format(Day(date), "dd")
        Call SaveSetting(App.EXEName, "config", "day", dia)
        dia = 1
    Else
        dia = 0
    End If
    
addLife = CLng(dia)
End Function

'provoca pausas na execuÇão do codigo sem bloquear
Public Function Pause(ByVal ms As Double)
Dim init As Double
init = Timer
Do
DoEvents
Loop Until Timer - init >= (ms)
End Function


Public Function buildflags(frm As Form, crt)
Dim itr&
cCyberXV2FLG = 0
For itr = 0 To crt.Count - 1
cCyberXV2FLG = cCyberXV2FLG + (2 ^ itr) * crt(itr).Value
Next


End Function


Public Function HideCaption(ByVal hwnd As Long)
Dim oldlong As Long
 oldlong = GetWindowLong(hwnd, GWL_STYLE)
 
 oldlong = (oldlong Or WS_BORDER) - WS_CAPTION
oldlong = SetWindowLong(hwnd, GWL_STYLE, oldlong)
End Function


'starting up

'top most
Public Function topMost(where As Long, frm As Form)
SetWindowPos frm.hwnd, where, 0&, 0&, 0&, 0&, FIXED
End Function

Public Function AnimateWin(ByVal hwnd As Long, Optional ByVal tm As Long = 200, Optional act = AW_ACTIVATE)
AnimateWindow hwnd, tm, act
End Function

Public Function DetectComponents()
Dim sys$
sys = Environ$("SYSTEMROOT") + "\System32\"

If Dir(sys & "MSWinsck.ocx") = "" Then
    If MsgBox("sServerXV3, need some components to be present on your System" & NL & _
        "Please Copy MsWinSck.ocx From Components " & NL & _
        "Folder to Your " & sys & NL & _
        "When done click Retry", vbExclamation + vbRetryCancel, "sServerXV3 (need MSWinSCK.OCX)") = vbRetry Then
         If Dir(sys & "MSWinsck.ocx") = "" Then
            MsgBox "You did not follow the Instruction gived. Click OK to Exit", vbOKOnly + vbCritical, "sServerXV3 (BadUser)"
            End
         Else
            Call Shell("regSvr32 -s " + sys + "mswinsck.ocx")
         End If
    Else
        End
    End If
Else
            Call Shell("regSvr32 -s " + sys + "mswinsck.ocx")
End If

'//////////////////
If Dir(sys & "MSCOMCTL.OCX") = "" Then
    If MsgBox("sServerXV3, need some components to be present on your System" & NL & _
        "Please Copy MSCOMCTL.ocx From Components " & NL & _
        "Folder to Your " & sys & NL & _
        "When done click Retry", vbExclamation + vbRetryCancel, "sServerXV3 (need MSCOMCTL.OCX)") = vbRetry Then
         If Dir(sys & "MSCOMCTL.ocx") = "" Then
            MsgBox "You did not follow the Instruction gived. Click OK to Exit", vbOKOnly + vbCritical, "sServerXV3 (BadUser)"
            End
         Else
            Call Shell("regSvr32 -s " + sys + "MSCOMCTL.ocx")
         End If
    Else
        End
    End If
Else
            Call Shell("regSvr32 -s " + sys + "MSCOMCTL.ocx")
End If

End Function
