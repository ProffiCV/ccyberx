VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTip 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   255
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   255
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmanim 
      Interval        =   40
      Left            =   8190
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   210
      TabIndex        =   2
      Top             =   450
      Width           =   225
   End
   Begin VB.Timer tmretry 
      Interval        =   3000
      Left            =   2760
      Top             =   0
   End
   Begin VB.CheckBox chkOntop 
      Caption         =   "&Visible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   30
      TabIndex        =   1
      ToolTipText     =   "Check (Alwais Visible)"
      Top             =   0
      Width           =   945
   End
   Begin MSComctlLib.StatusBar status 
      Height          =   285
      Left            =   1050
      TabIndex        =   0
      Top             =   -30
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   8096
            Key             =   "det"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8096
            Key             =   "from"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clr As Long

Private snifferinstalled As Long
Private sFrom As SockAddr, sStart As SockAddr
Private Buff() As Byte
Private ret&

Dim flDir As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim cpos As POINTAPI

Private Sub chkOntop_Click()
frmTip.tmanim.Enabled = Not CBool(frmTip.chkOntop.Value)

If frmTip.chkOntop.Value = 1 Then
    topMost HWND_TOPMOST, Me
    Else
    topMost HWND_TOPMOST, Me
End If

On Error Resume Next

End Sub

Private Sub Form_Load()
frmTip.Top = 8
frmTip.Left = (Screen.Width - frmTip.Width) / 2
flDir = -1
fromsite = "Site: 000.000.000.000"
topMost HWND_TOPMOST, Me
Load frmre
sniffer.gethwnd
 If sniffer.myWin <> 0 Then
    If sniffer.installSniffer = 0 Then
        'ok
    End If
 Else
    
 End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
sniffer.uninstallSniffer
Set sniffer = Nothing
Unload frmre
End
End Sub

Private Sub status_PanelClick(ByVal Panel As MSComctlLib.Panel)
Unload Me
End Sub

Private Sub tmanim_Timer()
DoEvents
frmTip.tmanim.Enabled = False
GetCursorPos cpos
If cpos.Y <= 8 Then
    flDir = 1
ElseIf cpos.Y > (frmTip.Height \ Screen.TwipsPerPixelY) Then
    flDir = -1
End If

Debug.Print cpos.Y, flDir
'

Debug.Print "TOP " & frmTip.Top
If frmTip.Top >= 8 And flDir = 1 Then
    frmTip.Top = 8

ElseIf frmTip.Top <= -(frmTip.Height - 10) And flDir = -1 Then
    frmTip.Top = -frmTip.Height - 10
Else
frmTip.Top = frmTip.Top + 120 * flDir

End If

frmTip.tmanim.Enabled = True
End Sub



