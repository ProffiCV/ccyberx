VERSION 5.00
Begin VB.Form frmCfg 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "cCyberXV2 Client Software"
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
      Height          =   1515
      Left            =   600
      TabIndex        =   13
      Top             =   2520
      Width           =   6015
      Begin VB.CheckBox chCt 
         Appearance      =   0  'Flat
         Caption         =   "&Use Own Wallpaper ( When Loaded)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Lock Mouse and keyboard"
         Top             =   1200
         Width           =   3405
      End
      Begin VB.CheckBox chCt 
         Appearance      =   0  'Flat
         Caption         =   "&Lock Desktop (When Free)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Lock Mouse and keyboard"
         Top             =   720
         Width           =   2865
      End
      Begin VB.CheckBox chCt 
         Appearance      =   0  'Flat
         Caption         =   "&Lock Mouse/Keyb (When Free)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Lock Mouse and keyboard"
         Top             =   960
         Width           =   2835
      End
      Begin VB.CheckBox chCt 
         Appearance      =   0  'Flat
         Caption         =   "&Apply Policies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Load and Aply Policies"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Public Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   210
         Width           =   1155
      End
   End
   Begin VB.ComboBox cboCfg 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3780
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   330
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4860
      TabIndex        =   3
      Top             =   90
      Width           =   1755
      Begin VB.CommandButton cmdCfg 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   1020
         Width           =   1605
      End
      Begin VB.CommandButton cmdCfg 
         Caption         =   "&Test Mode"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   660
         Width           =   1605
      End
      Begin VB.CommandButton cmdCfg 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   300
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Prices (ESC CVE's) [$] and Offset Download"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   600
      TabIndex        =   0
      Top             =   90
      Width           =   4245
      Begin VB.ComboBox cboCfg 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1710
         Width           =   975
      End
      Begin VB.ComboBox cboCfg 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1110
         Width           =   975
      End
      Begin VB.ComboBox cboCfg 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   660
         Width           =   975
      End
      Begin VB.Line Line3 
         X1              =   4110
         X2              =   150
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Line Line2 
         X1              =   150
         X2              =   4200
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   4170
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "For New Cards Get Download Value After (x) MB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   1680
         Width           =   2970
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(min) Before Entering Sleep Mode "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Top             =   1170
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Internet Streams (1 Mega Byte):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   720
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Machine (Per Hour):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   1740
      End
   End
   Begin VB.Image Image1 
      Height          =   6930
      Left            =   0
      Picture         =   "frmCfg.frx":0000
      Top             =   -2820
      Width           =   555
   End
End
Attribute VB_Name = "frmCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCfg_Click(Index As Integer)
    Select Case Index
        Case 0
        Me.Visible = False
            If MsgBox("Do you really want to save configurations?", vbQuestion + vbYesNo, App.EXEName) = vbYes Then
                SaveSetting App.EXEName, "Config", "P0", Hex(Me.cboCfg(0).ListIndex)
                SaveSetting App.EXEName, "Config", "P1", Hex(Me.cboCfg(1).ListIndex)
                SaveSetting App.EXEName, "Config", "P2", Hex(Me.cboCfg(2).ListIndex)
                SaveSetting App.EXEName, "Config", "P3", Hex(Me.cboCfg(3).ListIndex)
                
                buildflags Me, Me.chCt
                SaveSetting App.EXEName, "Config", "CC", "" & cCyberXV2FLG
            
              setPrices Me.cboCfg(1).ListIndex, Me.cboCfg(0).ListIndex, Me.cboCfg(3).ListIndex, Me.cboCfg(2).ListIndex
             
            End If
            Me.Visible = True
        Case 1
        topMost HWND_TOPMOST, Me
        Call AnimateWin(Me.hwnd, 800, AW_BLEND Or AW_HIDE)
        Pause 4
        Unload Me
        Case 2
         
            
    End Select

End Sub

Private Sub Form_Load()
topMost HWND_TOPMOST, Me
Dim ival As Double, ind&
Dim trs$
For ival = 100 To 1000 Step 10
    Me.cboCfg(0).AddItem Format(ival, "#,##0.00")
Next

For ival = 10 To 100 Step 5
    Me.cboCfg(1).AddItem Format(ival, "#,##0.00")
Next

For ival = 2 To 4 Step 2
    Me.cboCfg(2).AddItem Format(ival, "#,##0.00")
Next
For ival = 0 To 99 Step 0.5
DoEvents
    Me.cboCfg(3).AddItem Format(ival, "#,##0.00")
    
Next

'Hex(Me.cboCfg(0).ListIndex )

trs$ = GetSetting(App.EXEName, "Config", "P0", "-1")
If trs$ <> "-1" Then
ival = CLng("&h" & trs$)
    Else
    ival = -1
End If

Me.cboCfg(0).ListIndex = ival

trs$ = GetSetting(App.EXEName, "Config", "P1", "-1")
If trs$ <> "-1" Then
ival = CLng("&h" & trs$)
    Else
    ival = -1
End If

Me.cboCfg(1).ListIndex = ival

trs$ = GetSetting(App.EXEName, "Config", "P2", "-1")
If trs$ <> "-1" Then
ival = CLng("&h" & trs$)
    Else
    ival = -1
End If

Me.cboCfg(2).ListIndex = ival

trs$ = GetSetting(App.EXEName, "Config", "P3", "-1")
If trs$ <> "FFFF" And trs$ <> "-1" Then
ival = CLng("&h" & trs$)
    Else
    ival = -1
End If

Me.cboCfg(3).ListIndex = ival
  
  
  For ind = 0 To Me.chCt.Count - 1
    Me.chCt(ind).Value = 0 - CInt(CBool(cCyberXV2FLG And (2 ^ ind)))
  Next
  
  
End Sub

