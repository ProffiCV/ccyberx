VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   1665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMenu.frx":0000
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   111
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmdo 
      Interval        =   200
      Left            =   1260
      Top             =   570
   End
   Begin VB.Timer tmKidding 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1110
      Top             =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00660000&
      Caption         =   "(4 Sec) What to do?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   60
      TabIndex        =   6
      Top             =   1650
      Width           =   1590
   End
   Begin VB.Label cmdMenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   60
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1290
      Width           =   1575
   End
   Begin VB.Label cmdMenu 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   60
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1050
      Width           =   1575
   End
   Begin VB.Label cmdMenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   60
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   810
      Width           =   1575
   End
   Begin VB.Label cmdMenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Card"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   60
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label cmdMenu 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Logoff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   90
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   300
      Width           =   1575
   End
   Begin VB.Label cmdMenu 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   60
      MouseIcon       =   "frmMenu.frx":A5A2
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myIndex As Integer
Dim mt As Double
Private Sub cmdMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
myIndex = Index
cmdMenu(myIndex).FontBold = True
Me.Visible = False
Select Case Index
 Case 0 'login
 Case 1
 frmAg.cmdLogoff_Click
 Case 2
 frmAg.cmdChangecard_Click
 Case 3 'administrator
 Case 4
 frmAg.cmdAbout_Click
 Case 5
 Me.tmKidding.Enabled = False
 
End Select

Unload Me

End Sub

Private Sub cmdMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = myIndex Then Exit Sub
If myIndex <> -1 Then
cmdMenu(myIndex).FontBold = False
cmdMenu(myIndex).BackStyle = 0
End If

cmdMenu(Index).FontBold = True
cmdMenu(Index).BackStyle = 1

myIndex = Index

End Sub


Private Sub Form_Load()
detectTaskBarDimensions
topMost HWND_TOPMOST, Me
myIndex = -1
Dim Index As Integer
 cmdMenu(0).BackStyle = 1
 cmdMenu(Index).BackColor = &H660000
 cmdMenu(0).BackStyle = 0
 cmdMenu(0).Left = 0
 cmdMenu(0).Caption = "   " & cmdMenu(0).Caption

For Index = 1 To Me.cmdMenu.count - 1
 cmdMenu(Index).MouseIcon = cmdMenu(0).MouseIcon
 cmdMenu(Index).BackStyle = 1
 cmdMenu(Index).BackColor = &H660000
 cmdMenu(Index).BackStyle = 0
 cmdMenu(Index).Left = 0
 cmdMenu(Index).Caption = "   " & cmdMenu(Index).Caption
Next
cmdMenu(1).Enabled = LoggedInFlag
cmdMenu(2).Enabled = LoggedInFlag
cmdMenu(4).Caption = "   Show About"
Me.tmKidding.Enabled = True
Me.Top = Screen.height - TBAR.height * Screen.TwipsPerPixelY - Me.height
Me.Left = Screen.width - Me.width - 20
mt = Time

End Sub

Private Sub tmdo_Timer()
DoEvents
Me.Label1 = "(" & 4 - CLng(Format((Time - mt), "s")) & " Sec) What to do?"
Me.Label1.ForeColor = IIf(Me.Label1.ForeColor = vbWhite, vbRed, vbWhite)
End Sub

Private Sub tmKidding_Timer()
tmdo.Enabled = False
Me.tmKidding.Enabled = False
Unload Me
End Sub
