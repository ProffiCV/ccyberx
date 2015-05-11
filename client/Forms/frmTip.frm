VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   255
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   255
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   315
      Left            =   420
      TabIndex        =   2
      Top             =   510
      Width           =   465
   End
   Begin VB.Timer tmanim 
      Interval        =   80
      Left            =   7470
      Top             =   0
   End
   Begin VB.Timer tmdetails 
      Interval        =   500
      Left            =   10500
      Top             =   -30
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
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   9340
            Key             =   "det"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10328
            MinWidth        =   3528
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
Dim flDir As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim cpos As POINTAPI

Public Sub chkOntop_Click()
Me.tmanim.Enabled = Not CBool(Me.chkOntop.Value)

If Me.chkOntop.Value = 1 Then
    topMost HWND_TOPMOST, Me
    Else
    topMost HWND_TOPMOST, Me
End If

On Error Resume Next
'If frmTip.Visible = True Then Me.Command1.SetFocus
End Sub



Private Sub Form_Load()
Me.Top = 8
Me.Left = (Screen.width - Me.width) / 2
flDir = -1
fromSite = "...Developed by: Edson Martins (microbodix@hotmail.com) ©2003-2008"
topMost -1&, Me

 
End Sub

Private Sub tmanim_Timer()
DoEvents
Me.tmanim.Enabled = False
GetCursorPos cpos
If cpos.Y <= 8 Then
    flDir = 1
ElseIf cpos.Y > (Me.height \ Screen.TwipsPerPixelY) Then
    flDir = -1
End If

'


If Me.Top >= 8 And flDir = 1 Then
    Me.Top = 8

ElseIf Me.Top <= -(Me.height - 10) And flDir = -1 Then
    
    Me.Top = -Me.height - 10
    Me.chkOntop.Value = 0
Else
Me.Top = Me.Top + 120 * flDir

End If

Me.tmanim.Enabled = True
End Sub

Private Sub tmdetails_Timer()
DoEvents

prepareBarDetails

If Me.status.Panels("det").Text <> strToshow Then _
Me.status.Panels("det").Text = strToshow

If Me.status.Panels("from").Text <> fromSite Then _
Me.status.Panels("from").Text = fromSite

End Sub



