VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmReduced.frx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5100
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReduced.frx":474B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReduced.frx":489C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4890
      TabIndex        =   0
      Top             =   30
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.lbn.ForeColor <> vbBlack Then Me.lbn.ForeColor = vbBlack
End Sub

Private Sub lbn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lbn.Top = Me.lbn.Top + 10
Me.lbn.Left = Me.lbn.Left + 10
End Sub

Private Sub lbn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.lbn.ForeColor <> vbWhite Then Me.lbn.ForeColor = vbWhite
End Sub

Private Sub lbn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

frmsv2.WindowState = vbNormal

End Sub
