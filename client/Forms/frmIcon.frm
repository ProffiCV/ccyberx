VERSION 5.00
Begin VB.Form frmIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon"
   ClientHeight    =   330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   1545
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1
Private Const WM_SETICON = &H80
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Dim icons&, iconb&
Private Const MK_LBUTTON = &H1
Private doubleClick As Long
Private Sub Form_Load()
Dim lpzicon$
lpzicon$ = Environ("Systemroot") & "\System32\Shell32.dll"

If Dir(lpzicon$) <> "" Then
    If ExtractIconEx(lpzicon$, 39, iconb&, icons&, 1) <> 0 Then
        SendMessage Me.hwnd, WM_SETICON, ICON_BIG, ByVal iconb&
        SendMessage Me.hwnd, WM_SETICON, ICON_SMALL, ByVal icons&
    End If
End If
noti.setupSystrayIcon Me, iconb&
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X And MK_LBUTTON Then
  'If LoggedInFlag = True Then
     frmAg.cmdAbout.Caption = "&About " & App.EXEName
     frmMenu.show
 ' End If
  
Else

End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
noti.removeSystrayIcon
noti.removeSystrayIcon
Set noti = Nothing
End Sub
