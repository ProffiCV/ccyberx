VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHist 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8685
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2250
      TabIndex        =   3
      Top             =   4740
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refrescar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   4740
      Width           =   1605
   End
   Begin VB.TextBox txtdet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   4290
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   60
      Width           =   4335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3450
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOldCards.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOldCards.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvr 
      Height          =   4605
      Left            =   600
      TabIndex        =   0
      Top             =   60
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8123
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   6930
      Index           =   1
      Left            =   0
      Picture         =   "frmOldCards.frx":08A4
      Top             =   -1770
      Width           =   555
   End
End
Attribute VB_Name = "frmHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TV_FIRST = &H1100
Private Const TVM_SETBKCOLOR = (TV_FIRST + 29)
Private Const TVM_SETTEXTCOLOR = (TV_FIRST + 30)
'TVM_SETBKCOLOR
'    wParam = 0;
'    lParam = (LPARAM)(COLORREF) clrBk
    
Private Sub Command1_Click()
On Error GoTo fim
Me.tvr.Nodes.Clear
Dim a, ind&
Debug.Print getPcs.Count
For Each a In getPcs
   With Me.tvr
   ind& = ind + 1
    .Nodes.Add , , "pc" & ind&, a, 1, 1
   
    getDetails Me.tvr, "pc" & ind&, CStr(a)
   End With

Next

If Me.tvr.Nodes.Count <> 0 Then Me.tvr.Nodes(1).Selected = True

Exit Sub
fim:
tell "Error " & Err.Description
End Sub


Private Sub Command2_Click()
         topMost HWND_TOPMOST, Me
        Call AnimateWin(Me.hwnd, 800, AW_BLEND Or AW_HIDE)
        Pause 4
        Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = frmsv2.Icon
'SendMessage Me.tvr.hwnd, TVM_SETBKCOLOR, 0&, ByVal RGB(255, 230, 255)
Me.Caption = "Use Details " & App.EXEName
HideCaption Me.hwnd
Me.Height = 5235
 topMost HWND_TOPMOST, Me
End Sub

Private Sub tvr_Click()
If tvr.Nodes.Count = 0 Then Exit Sub
If (tvr.SelectedItem.Tag) <> "" Then
    Me.txtdet.Text = tvr.SelectedItem.Tag
End If

End Sub

