VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCfg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cCyberXV2 Configurations"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSell 
      Caption         =   "Sell Downloads"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4440
      TabIndex        =   24
      ToolTipText     =   "Mark this to enable selling Downloads"
      Top             =   3690
      Width           =   3075
   End
   Begin VB.CheckBox chAuto 
      Caption         =   "Execute With Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1590
      TabIndex        =   23
      Top             =   3690
      Width           =   2505
   End
   Begin VB.CheckBox chc 
      Caption         =   "TestNode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   22
      Top             =   3690
      Width           =   1425
   End
   Begin VB.CommandButton cmdCfg 
      Caption         =   "&Reset Setting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   5190
      TabIndex        =   21
      Top             =   3270
      Width           =   1455
   End
   Begin VB.CommandButton cmdCfg 
      Caption         =   "&View cCyberXV2 Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   2760
      TabIndex        =   20
      Top             =   3270
      Width           =   2445
   End
   Begin VB.CommandButton cmdCfg 
      Caption         =   "&Build Desktop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   1380
      TabIndex        =   19
      Top             =   3270
      Width           =   1395
   End
   Begin VB.CommandButton cmdCfg 
      Caption         =   "&Edit Policies"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   30
      TabIndex        =   18
      Top             =   3270
      Width           =   1365
   End
   Begin VB.Frame Frame3 
      Caption         =   "Click [Check] to know if the selected Network Card  would work Properly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   13
      Top             =   4140
      Width           =   6735
      Begin VB.CommandButton cmdchk 
         Caption         =   "&Check"
         Height          =   315
         Left            =   5670
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label llog 
         Caption         =   "See activity here..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5265
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6765
      Begin VB.CommandButton cmdCfg 
         Caption         =   "&Save Config"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   5100
         TabIndex        =   17
         Top             =   2760
         Width           =   1605
      End
      Begin VB.CommandButton cmdCfg 
         Caption         =   "&DetectDevices"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   3420
         TabIndex        =   16
         Top             =   2760
         Width           =   1605
      End
      Begin MSComctlLib.ListView lstv 
         Height          =   765
         Left            =   60
         TabIndex        =   12
         Top             =   1950
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   1349
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   4080
         TabIndex        =   4
         Top             =   480
         Width           =   2535
         Begin VB.CheckBox chkDetectType 
            Appearance      =   0  'Flat
            Caption         =   "Protect"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   930
            Width           =   1935
         End
         Begin VB.OptionButton optCard 
            Caption         =   " Wireless"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   8
            Top             =   510
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.OptionButton optCard 
            Caption         =   " Wireless"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   7
            Top             =   510
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.OptionButton optCard 
            Caption         =   " Wired"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton optCard 
            Caption         =   " Wired"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   240
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Shape Shape2 
            Height          =   645
            Left            =   1260
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "automatically detected "
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   210
            TabIndex        =   9
            Top             =   0
            Width           =   1905
         End
         Begin VB.Shape Shape1 
            Height          =   645
            Left            =   120
            Top             =   210
            Visible         =   0   'False
            Width           =   1125
         End
      End
      Begin VB.ListBox lstport 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         ItemData        =   "frmCfg.frx":0000
         Left            =   2910
         List            =   "frmCfg.frx":0002
         TabIndex        =   2
         Top             =   450
         Width           =   1125
      End
      Begin VB.TextBox tHost 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   90
         MaxLength       =   16
         TabIndex        =   1
         Top             =   450
         Width           =   2715
      End
      Begin VB.Shape Shape3 
         Height          =   1215
         Left            =   4050
         Top             =   450
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detected Online Network Devices:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   10
         Top             =   1770
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server Name or IP Address:               Work as:               NetWork Card Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   210
         Width           =   6000
      End
   End
End
Attribute VB_Name = "frmCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rHost$
Private rPort&
Private cardt&
Private valDetect&

Private Sub Check1_Click()

End Sub





Public Sub chc_Click()
frmm.Timer2.Enabled = Me.chc.Value = 1

End Sub

Public Sub chkDetectType_Click()
Dim ind%
valDetect& = Me.chkDetectType.Value
For ind% = 0 To Me.optCard.count - 1
    Me.optCard(ind).Enabled = Not CBool(Me.chkDetectType.Value)
Next

End Sub

Public Sub cmdCfg_Click(Index As Integer)

Select Case Index
    Case 0
        If MsgBox("Do you want to update configurations? Connections Will be Reseted...", vbYesNo + vbQuestion, App.EXEName) = vbYes Then
            SaveSetting App.EXEName, "Remote", "Server", Me.tHost.Text
            SaveSetting App.EXEName, "Remote", "Port", CLng(Me.lstport.ListIndex)
            SaveSetting App.EXEName, "Remote", "Autod", valDetect&
            SaveSetting App.EXEName, "Remote", "CardT", cardt
            
            Dim autorun As Boolean
            autorun = False
            If Me.chAuto.Value = 1 Then autorun = True
            Call enableAutoRun("cCyberEdson", App.Path & "\" & App.EXEName & ".exe", autorun)
            SaveSetting App.EXEName, "Remote", "autorun", CStr(autorun)
            SaveSetting App.EXEName, "Remote", "SellDownload", CStr(Me.chkSell.Value)

            frmCli.tCli.Close
            Unload frmCli
            
                      
            tell "Reseting Self..."
            Pause 1.5
                      
            
            getMerlin
            
        End If

    Case 1
    Me.lstv.ListItems.clear
        detectNetWorkDevices Me.lstv
    Case 2
    If Dir(App.Path & "\ccyberXV2.ini") = "" Then
    burnPolicieFile
    End If
    
    Shell "Notepad " & App.Path & "\ccyberXV2.ini", vbNormalFocus
    Case 3
    createFolders
    MsgBox "Create shortcuts that will be used by clients inside the folder that will be Opened...", vbInformation, App.EXEName
    Shell "Explorer " & App.Path & "\myDesk", vbNormalFocus
    Case 4
    Shell "Explorer " & App.Path, vbNormalFocus
    Case 5
    If MsgBox("Do you really want to reset settings (Value " & cCyberXV2FLG & ")", vbQuestion + vbYesNo, App.EXEName) = vbYes Then
    If GetSetting(App.EXEName, "startup", "task", 0) <> 0 Then
    DeleteSetting App.EXEName, "startup", "task"
    End If
    
    End If
    
    
End Select

End Sub

Private Sub Form_Load()
Dim it&

For it& = 0 To 7
    Me.lstport.AddItem "Client " & it& + 1
    Me.lstport.ItemData(it) = 4001 + it
Next

rHost$ = GetSetting(App.EXEName, "Remote", "Server", "")
rPort& = CLng(GetSetting(App.EXEName, "Remote", "Port", -1))
cardt = CLng(GetSetting(App.EXEName, "Remote", "Cardt", 0))
valDetect = CLng(GetSetting(App.EXEName, "Remote", "Autod", 0))
Dim strret$
strret$ = GetSetting(App.EXEName, "Remote", "autorun", "false")
Me.chAuto.Value = CInt(IIf(UCase(strret$) = "TRUE", 1, 0))
Me.chkSell.Value = CInt(GetSetting(App.EXEName, "Remote", "SellDownload", 0))
Me.tHost.Text = rHost
Me.lstport.ListIndex = rPort&

Me.optCard(cardt).Value = True
Me.chkDetectType.Value = valDetect&

'enable or disable cards type selector
'chkDetectType_Click

With Me.lstv
    .ColumnHeaders.Add , , "Device", 0.6 * .width
    .ColumnHeaders.Add , , "MAC Address"
    .ColumnHeaders.Add , , "Manufacturer"
End With

Me.chkDetectType.Value = 1
chkDetectType_Click


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
Me.chc.Value = 0
chc_Click
cmdCfg_Click 0
If goend = False Then
frmlogc.Visible = True
End If

'
End If

End Sub

Private Sub lstv_Click()
If Me.lstv.ListItems.count = 0 Then Exit Sub
''Me.llog = "Current IP Address " & Me.lstv.SelectedItem.Tag & IIf(frmCli.tCli.LocalIP = Me.lstv.SelectedItem.Tag, " ONLINE.", " OFFLINE.")
''Me.cmdchk.Visible = InStr(Me.llog, "ONLINE.") <> 0
End Sub

Private Sub optCard_Click(Index As Integer)
Me.chkDetectType.Value = 0
cardt = Index
End Sub
