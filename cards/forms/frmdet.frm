VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cards Details PacketXV2"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmdet.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   10845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   60
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   7200
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstdet 
      Height          =   6315
      Left            =   30
      TabIndex        =   2
      Top             =   570
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   11139
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16761024
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   4290
      Top             =   0
   End
   Begin VB.CommandButton cmdConUn 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7740
      TabIndex        =   1
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright© 2007 Edson Martins"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   8280
      TabIndex        =   5
      Top             =   7230
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SuperVisorXV3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   9360
      TabIndex        =   4
      Top             =   7080
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   6915
      Index           =   1
      Left            =   10320
      Picture         =   "frmdet.frx":FA752
      Top             =   -3690
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   6915
      Index           =   0
      Left            =   10320
      Picture         =   "frmdet.frx":108E14
      Top             =   540
      Width           =   630
   End
   Begin VB.Label lbp 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   7140
      Width           =   7755
   End
End
Attribute VB_Name = "frmdet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConUn_Click()
'cdlg
If MsgBox("The assistant will print only Selected Cards" & vbCrLf & _
"Do you want to continue?", vbYesNo + vbQuestion, App.EXEName) = vbYes Then
    Dim itr&
    
    If Me.lstdet.ListItems.Count = 0 Then
    MsgBox "No Cards Found. Click Ok to Close this MessageBox", vbInformation, App.EXEName
       Exit Sub
    End If
    
     Me.dlg.ShowPrinter
     If Me.dlg.CancelError = False Then
        frmp.Show vbModal
     End If
     
    
End If

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'"001 1111-2222-3333-4444 221006 00 F 0200 0000 000000000000"
With Me.lstdet
    .ColumnHeaders.Add , , "ID", 0.06 * .Width
    .ColumnHeaders.Add , , "Code", 0.25 * .Width
    .ColumnHeaders.Add , , "Created", 0.12 * .Width
    .ColumnHeaders.Add , , "Days", 0.055 * .Width
    .ColumnHeaders.Add , , "U", 0.03 * .Width
    .ColumnHeaders.Add , , "Price", 0.1 * .Width
    .ColumnHeaders.Add , , "Used", 0.1 * .Width
    .ColumnHeaders.Add , , "Remain", 0.1 * .Width
    .ColumnHeaders.Add , , "Charge"
.LabelEdit = lvwManual
.FullRowSelect = True
.AllowColumnReorder = True
.View = lvwReport
End With

End Sub

Private Sub lstdet_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.lstdet.SortOrder = 1 - Me.lstdet.SortOrder
Me.lstdet.SortKey = ColumnHeader.index - 1
Me.lstdet.Sorted = True
End Sub

Private Sub Timer1_Timer()
DoEvents
Dim tps$
Me.lbp.Alignment = 1 - Me.lbp.Alignment
tps$ = getPrinter
If Me.lbp <> tps$ Then
Me.lbp = "Current Printer [" & tps$ & "]"
End If

End Sub
