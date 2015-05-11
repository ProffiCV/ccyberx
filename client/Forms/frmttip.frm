VERSION 5.00
Begin VB.Form frmttip 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmttip.frx":0000
   ScaleHeight     =   1155
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbmsg 
      BackStyle       =   0  'Transparent
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
      Height          =   855
      Left            =   30
      TabIndex        =   0
      Top             =   270
      Width           =   3015
   End
End
Attribute VB_Name = "frmttip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
topMost HWND_TOPMOST, Me
detectTaskBarDimensions

Me.Top = Screen.height - TBAR.height * Screen.TwipsPerPixelY - Me.height
Me.Left = Screen.width - Me.width - 20

End Sub

Private Sub lbmsg_Click()
Unload Me
End Sub
