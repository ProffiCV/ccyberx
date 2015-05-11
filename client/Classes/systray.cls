VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "systray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_MESSAGE = &H1

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_LBUTTONUP = &H202

Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hicon As Long
        szTip As String * 64
End Type

Dim lpData As NOTIFYICONDATA

'-------------
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Function setupSystrayIcon(ByVal frm As Form, hicon As Long)
Dim msg As Long
With lpData
    .cbSize = Len(lpData)
    .hicon = hicon
    .hwnd = frm.hwnd
    .uID = 0
    .uCallbackMessage = WM_LBUTTONUP
    .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    .szTip = "cCyberXV3 Client Copyright© " & Chr(10) & "2003-2008 Edson Martins" & Chr(0)
End With

Debug.Print Shell_NotifyIcon(NIM_ADD, lpData)
End Function


Public Function removeSystrayIcon()
Shell_NotifyIcon NIM_DELETE, lpData
End Function

Public Function steelThere() As Long
steelThere = Shell_NotifyIcon(NIM_MODIFY, lpData)
End Function
