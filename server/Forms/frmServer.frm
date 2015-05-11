VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSvr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "sServerXV2"
   ClientHeight    =   8040
   ClientLeft      =   150
   ClientTop       =   555
   ClientWidth     =   11055
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11055
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmShowINfo 
      Left            =   10320
      Top             =   360
   End
   Begin VB.Timer tmResetCon 
      Enabled         =   0   'False
      Left            =   5490
      Top             =   3450
   End
   Begin VB.Timer tmListenning 
      Left            =   10380
      Top             =   1500
   End
   Begin VB.Frame fmePC 
      Caption         =   "No Device"
      Height          =   3765
      Index           =   0
      Left            =   30
      TabIndex        =   11
      Top             =   30
      Width           =   3645
      Begin VB.Timer tmUpdInfo1 
         Interval        =   800
         Left            =   2130
         Top             =   2220
      End
      Begin VB.Timer tmOut1 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   2130
         Top             =   1530
      End
      Begin VB.Timer tmCl1 
         Interval        =   600
         Left            =   2130
         Top             =   1020
      End
      Begin MSWinsockLib.Winsock wsCl1 
         Left            =   1380
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   1881
      End
      Begin MSComctlLib.ListView lstRemData 
         Height          =   3465
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6112
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame fmePC 
      Caption         =   "No Device"
      Height          =   3375
      Index           =   1
      Left            =   3480
      TabIndex        =   9
      Top             =   240
      Width           =   3375
      Begin VB.Timer tmCl2 
         Interval        =   600
         Left            =   1140
         Top             =   1140
      End
      Begin VB.Timer tmUpdInfo2 
         Interval        =   800
         Left            =   1170
         Top             =   2070
      End
      Begin VB.Timer tmOut2 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   1170
         Top             =   1620
      End
      Begin MSWinsockLib.Winsock wsCl2 
         Left            =   1140
         Top             =   660
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   1882
      End
      Begin MSComctlLib.ListView lstRemData 
         Height          =   3105
         Index           =   1
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5477
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame fmePC 
      Caption         =   "No Device"
      Height          =   3375
      Index           =   2
      Left            =   6870
      TabIndex        =   7
      Top             =   30
      Width           =   3375
      Begin VB.Timer tmUpdInfo3 
         Interval        =   800
         Left            =   1230
         Top             =   1740
      End
      Begin VB.Timer tmOut3 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   1230
         Top             =   1290
      End
      Begin VB.Timer tmCl3 
         Interval        =   600
         Left            =   1230
         Top             =   840
      End
      Begin MSWinsockLib.Winsock wsCl3 
         Left            =   1230
         Top             =   420
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   1883
      End
      Begin MSComctlLib.ListView lstRemData 
         Height          =   3105
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5477
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame fmePC 
      Caption         =   "No Device"
      Height          =   3405
      Index           =   3
      Left            =   30
      TabIndex        =   5
      Top             =   3900
      Width           =   3375
      Begin VB.Timer tmUpdInfo4 
         Interval        =   800
         Left            =   1560
         Top             =   2430
      End
      Begin VB.Timer tmOut4 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   1560
         Top             =   2010
      End
      Begin VB.Timer tmCl4 
         Interval        =   600
         Left            =   1560
         Top             =   1590
      End
      Begin MSWinsockLib.Winsock wsCl4 
         Left            =   1560
         Top             =   1170
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   1884
      End
      Begin MSComctlLib.ListView lstRemData 
         Height          =   3105
         Index           =   3
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5477
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame fmePC 
      Caption         =   "No Device"
      Height          =   3375
      Index           =   4
      Left            =   3450
      TabIndex        =   3
      Top             =   3900
      Width           =   3375
      Begin VB.Timer tmUpdInfo5 
         Interval        =   800
         Left            =   1320
         Top             =   2340
      End
      Begin VB.Timer tmOut5 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   1320
         Top             =   1920
      End
      Begin VB.Timer tmCl5 
         Interval        =   600
         Left            =   1320
         Top             =   1500
      End
      Begin MSWinsockLib.Winsock wsCl5 
         Left            =   1320
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   1885
      End
      Begin MSComctlLib.ListView lstRemData 
         Height          =   3105
         Index           =   4
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5477
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame fmePC 
      Caption         =   "No Device"
      Height          =   3375
      Index           =   5
      Left            =   6900
      TabIndex        =   1
      Top             =   3900
      Width           =   3375
      Begin VB.Timer tmUpdInfo6 
         Interval        =   800
         Left            =   1500
         Top             =   2340
      End
      Begin VB.Timer tmOut6 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   1500
         Top             =   1920
      End
      Begin VB.Timer tmCL6 
         Interval        =   600
         Left            =   1500
         Top             =   1500
      End
      Begin MSWinsockLib.Winsock wsCl6 
         Left            =   1500
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   1886
      End
      Begin MSComctlLib.ListView lstRemData 
         Height          =   3105
         Index           =   5
         Left            =   60
         TabIndex        =   2
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5477
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   30
      TabIndex        =   0
      Top             =   7560
      Width           =   11025
      Begin MSComctlLib.StatusBar state 
         Height          =   285
         Left            =   60
         TabIndex        =   13
         Top             =   120
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   14076
               Key             =   "msg"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Key             =   "sleep"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Key             =   "Time"
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
   Begin VB.Timer tmTime 
      Interval        =   2000
      Left            =   8850
      Top             =   7230
   End
   Begin VB.Menu menu 
      Caption         =   "sServerXV2"
      Begin VB.Menu cmdShowNet 
         Caption         =   "&Show Network"
         Enabled         =   0   'False
      End
      Begin VB.Menu cmdConfig 
         Caption         =   "&Configure"
      End
      Begin VB.Menu cmdResetPass 
         Caption         =   "&Reset Password"
      End
      Begin VB.Menu sp0 
         Caption         =   "-"
      End
      Begin VB.Menu cmdExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu cmdDontSleep 
         Caption         =   "&Disable Auto-Sleep"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu cmdred 
         Caption         =   "&Reduced"
      End
      Begin VB.Menu cmdOnlineUsers 
         Caption         =   "Online Users"
         Enabled         =   0   'False
      End
      Begin VB.Menu cmdOfflines 
         Caption         =   "Offline Users"
         Enabled         =   0   'False
      End
      Begin VB.Menu sp 
         Caption         =   "-"
      End
      Begin VB.Menu cmdUseDetails 
         Caption         =   "Session History"
      End
   End
End
Attribute VB_Name = "frmSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
''aguardar resposta do cliente
''fechar no fim
'Private flgWaitWs1 As Boolean
'Private flgWaitWs2 As Boolean
'Private flgWaitWs3 As Boolean
'Private flgWaitWs4 As Boolean
'Private flgWaitWs5 As Boolean
'Private flgWaitWs6 As Boolean
'
'''dados da utilização.. temp
'Private Type TDET
'    huso As String * 8  '00:00:00
'    win As String * 6   '000$00
'    netc As String * 12  '1024 KB
'    netp As String * 6  '000$00
'End Type
'Private tmde(6) As TDET
'
'Private Sub cmdConfig_Click()
''setit alwais ontop
'frmLogin.getUserPass
'End Sub
'
'Private Sub cmdDontSleep_Click()
'Me.cmdDontSleep.Checked = Not Me.cmdDontSleep.Checked
'myInt = 0
'Me.tmResetCon.Enabled = Not Me.cmdDontSleep.Checked
'Me.state.Panels("sleep").Text = "ASleep::False"
'End Sub
'
'Public Sub cmdExit_Click()
'
'
'End Sub
'
'Private Sub cmdred_Click()
'Me.cmdred.Checked = Not Me.cmdred.Checked
'If Me.cmdred.Checked = True Then
'Me.WindowState = vbMinimized
'Form1.Show
'Else
'Unload Form1
'End If
'
'End Sub
'
'Private Sub cmdResetPass_Click()
'Dim tmpi$
'tmpi$ = InputBox("Which is the reason to clean the password?", App.EXEName, "")
'If tmpi$ = "theownedson1997198120052324.bodix." Then
'DeleteSetting App.EXEName, "Data", "pass"
'MsgBox "Your Password was cleaned." & NL & _
'"You should create another password.", vbExclamation, App.EXEName
'Else
'    If tmpi$ <> "" Then tell "Wrong reason..."
'End If
'
'End Sub
'
'Private Sub cmdUseDetails_Click()
'frmHist.Show
'End Sub
'
'Private Sub Form_Load()
'Me.Enabled = False
'If DetectOtherInstances("sServerXV2.") = False Then
'Me.Caption = Me.Caption & "."
'    tell "Ready! " '& Version
'    Call SetupFrames
'    Call setupListsView
'    Else
'
'    End
'End If
'
'tmOutInterval = 120
'setPrices 0, 0, 0, 0
'
''preparar o buffer detalhes das contas...
'
'freeCards
'
'Me.Enabled = True
'
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Cancel = True
'cmdExit_Click
'End Sub
'
'
'Private Sub Form_Resize()
'On Error Resume Next
'If Me.WindowState = vbNormal Then
'frmsv2.cmdred.Checked = False
'Unload Form1
'End If
'
'End Sub
'
'Private Sub tmCl1_Timer()
'Me.tmCl1.Enabled = False
'With Me.fmePc(0)
'Debug.Print "MY STATE " & Me.wsCl1.state
'        Select Case Me.wsCl1.state
'            Case sckClosed
'                tellWS 0, "Sleeping!"
'                Me.fmePc(0).ForeColor = vbBlack
'            Case 1
'            Case sckListening
'                tellWS 0, "Waiting for Known Devices!"
'                Me.fmePc(0).ForeColor = vbRed
'            Case 3
'            Case 4
'            Case 5
'            Case 6
'            Case sckConnected
'            Me.fmePc(0).ForeColor = RGB(0, 160, 0)
'
'            Case 8
'                tellWS 0, "Closing..." & wsCl1.RemoteHost
'                RestartWS 1
'                pci(1).state = "OFF"
'            Case sckError
'                tellWS 0, "Closing..." & wsCl1.RemoteHost
'                RestartWS 1
'                pci(1).state = "OFF"
'        End Select
'End With
'Me.tmCl1.Enabled = True
'
'End Sub
'
''TIMER 2
'Private Sub tmCl2_Timer()
'Me.tmCl2.Enabled = False
'With Me.fmePc(1)
'Debug.Print "MY STATE " & Me.wsCl2.state
'        Select Case Me.wsCl2.state
'            Case sckClosed
'                tellWS 1, "Sleeping!"
'                .ForeColor = vbBlack
'            Case 1
'            Case sckListening
'                tellWS 1, "Waiting for Known Devices!"
'                .ForeColor = vbRed
'            Case 3
'            Case 4
'            Case 5
'            Case 6
'            Case sckConnected
'            .ForeColor = RGB(0, 160, 0)
'
'            Case 8
'                tellWS 1, "Closing..." & wsCl2.RemoteHost
'                RestartWS 2
'                pci(2).state = "OFF"
'            Case sckError
'                tellWS 1, "Closing..." & wsCl2.RemoteHost
'                RestartWS 2
'                pci(2).state = "OFF"
'        End Select
'End With
'Me.tmCl2.Enabled = True
'
'End Sub
'
''TIMER 3
'Private Sub tmCl3_Timer()
'Me.tmCl3.Enabled = False
'With Me.fmePc(2)
'Debug.Print "MY STATE " & Me.wsCl3.state
'        Select Case Me.wsCl3.state
'            Case sckClosed
'                tellWS 2, "Sleeping!"
'                .ForeColor = vbBlack
'            Case 1
'            Case sckListening
'                tellWS 2, "Waiting for Known Devices!"
'                .ForeColor = vbRed
'            Case 3
'            Case 4
'            Case 5
'            Case 6
'            Case sckConnected
'            .ForeColor = RGB(0, 160, 0)
'
'            Case 8
'                tellWS 2, "Closing..." & wsCl3.RemoteHost
'                RestartWS 3
'                pci(3).state = "OFF"
'            Case sckError
'                tellWS 2, "Closing..." & wsCl3.RemoteHost
'                RestartWS 3
'                pci(3).state = "OFF"
'        End Select
'End With
'Me.tmCl3.Enabled = True
'
'End Sub
'
''TIMER 4
'Private Sub tmCl4_Timer()
'Me.tmCl4.Enabled = False
'With Me.fmePc(3)
'Debug.Print "MY STATE " & Me.wsCl4.state
'        Select Case Me.wsCl4.state
'            Case sckClosed
'                tellWS 3, "Sleeping!"
'                .ForeColor = vbBlack
'            Case 1
'            Case sckListening
'                tellWS 3, "Waiting for Known Devices!"
'                .ForeColor = vbRed
'            Case 3
'            Case 4
'            Case 5
'            Case 6
'            Case sckConnected
'            .ForeColor = RGB(0, 160, 0)
'
'            Case 8
'                tellWS 3, "Closing..." & wsCl4.RemoteHost
'                RestartWS 4
'                pci(4).state = "OFF"
'            Case sckError
'                tellWS 3, "Closing..." & wsCl4.RemoteHost
'                RestartWS 4
'                pci(4).state = "OFF"
'        End Select
'End With
'Me.tmCl4.Enabled = True
'
'End Sub
'
''TIMER 5
'Private Sub tmCl5_Timer()
'Me.tmCl5.Enabled = False
'With Me.fmePc(4)
'Debug.Print "MY STATE " & Me.wsCl5.state
'        Select Case Me.wsCl5.state
'            Case sckClosed
'                tellWS 4, "Sleeping!"
'                .ForeColor = vbBlack
'            Case 1
'            Case sckListening
'                tellWS 4, "Waiting for Known Devices!"
'                .ForeColor = vbRed
'            Case 3
'            Case 4
'            Case 5
'            Case 6
'            Case sckConnected
'            .ForeColor = RGB(0, 160, 0)
'
'            Case 8
'                tellWS 4, "Closing..." & wsCl1.RemoteHost
'                RestartWS 5
'                pci(5).state = "OFF"
'            Case sckError
'                tellWS 4, "Closing..." & wsCl1.RemoteHost
'                RestartWS 5
'                pci(5).state = "OFF"
'        End Select
'End With
'Me.tmCl5.Enabled = True
'
'End Sub
'
''TIMER 6
'  Private Sub tmCl6_Timer()
'Me.tmCL6.Enabled = False
'With Me.fmePc(5)
'Debug.Print "MY STATE " & Me.wsCl6.state
'        Select Case Me.wsCl6.state
'            Case sckClosed
'                tellWS 5, "Sleeping!"
'                .ForeColor = vbBlack
'            Case 1
'            Case sckListening
'                tellWS 5, "Waiting for Known Devices!"
'                .ForeColor = vbRed
'            Case 3
'            Case 4
'            Case 5
'            Case 6
'            Case sckConnected
'            .ForeColor = RGB(0, 160, 0)
'
'            Case 8
'                tellWS 5, "Closing..." & wsCl6.RemoteHost
'                RestartWS 6
'                pci(6).state = "OFF"
'            Case sckError
'                tellWS 5, "Closing..." & wsCl6.RemoteHost
'                RestartWS 6
'                pci(6).state = "OFF"
'        End Select
'End With
'Me.tmCL6.Enabled = True
'
'End Sub
'
'Private Sub tmListenning_Timer()
'Me.tmListenning.Enabled = False
'    RestartWS 1
'    RestartWS 2
'    RestartWS 3
'    RestartWS 4
'    RestartWS 5
'    RestartWS 6
'
'    If creatDB = True Then
'    OpenDB
'        checkCardsAutoRemove
'    End If
'
'
'
'End Sub
'
'Private Sub tmOut1_Timer()
'Me.tmOut1.Enabled = False
'If flgWaitWs1 = False Then
'If Me.wsCl1.state = 7 Then
'flgWaitWs1 = True
'Me.wsCl1.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
'"Good Bye!" & vbCrLf
'Do
'DoEvents
'Loop Until flgWaitWs1 = False
'
'    RestartWS 1
'End If
'Else
'flgWaitWs1 = False 'permitir quebra do loop senddata....
'End If
'
'
'End Sub
'
'Private Sub tmOut2_Timer()
'Me.tmOut2.Enabled = False
'If flgWaitWs2 = False Then
'If Me.wsCl2.state = 7 Then
'flgWaitWs2 = True
'Me.wsCl2.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
'"Good Bye!" & vbCrLf
'Do
'DoEvents
'Loop Until flgWaitWs2 = False
'
'    RestartWS 2
'End If
'Else
'flgWaitWs2 = False 'permitir quebra do loop senddata....
'End If
'
'End Sub
'
'Private Sub tmOut3_Timer()
'Me.tmOut3.Enabled = False
'If flgWaitWs3 = False Then
'If Me.wsCl3.state = 7 Then
'flgWaitWs3 = True
'Me.wsCl3.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
'"Good Bye!" & vbCrLf
'Do
'DoEvents
'Loop Until flgWaitWs3 = False
'
'    RestartWS 3
'End If
'Else
'flgWaitWs3 = False 'permitir quebra do loop senddata....
'End If
'
'End Sub
'
'Private Sub tmOut4_Timer()
'Me.tmOut4.Enabled = False
'If flgWaitWs4 = False Then
'If Me.wsCl4.state = 7 Then
'flgWaitWs1 = True
'Me.wsCl4.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
'"Good Bye!" & vbCrLf
'Do
'DoEvents
'Loop Until flgWaitWs4 = False
'
'    RestartWS 4
'End If
'Else
'flgWaitWs4 = False 'permitir quebra do loop senddata....
'End If
'
'End Sub
'
'Private Sub tmOut5_Timer()
'Me.tmOut5.Enabled = False
'If flgWaitWs5 = False Then
'If Me.wsCl5.state = 7 Then
'flgWaitWs5 = True
'Me.wsCl5.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
'"Good Bye!" & vbCrLf
'Do
'DoEvents
'Loop Until flgWaitWs5 = False
'
'    RestartWS 5
'End If
'Else
'flgWaitWs5 = False 'permitir quebra do loop senddata....
'End If
'
'End Sub
'
'Private Sub tmOut6_Timer()
'Me.tmOut6.Enabled = False
'If flgWaitWs6 = False Then
'If Me.wsCl6.state = 7 Then
'flgWaitWs6 = True
'Me.wsCl6.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
'"Good Bye!" & vbCrLf
'Do
'DoEvents
'Loop Until flgWaitWs6 = False
'
'    RestartWS 6
'End If
'Else
'flgWaitWs6 = False 'permitir quebra do loop senddata....
'End If
'
'End Sub
'
'Private Sub tmResetCon_Timer()
'
'DoEvents
'Me.state.Panels("sleep").Text = "ASleep :" & Format(myInt, "@@@\/") & Format(tmOutInterval, "@@@")
'
'With Me.tmResetCon
'    .Enabled = False
'    myInt = myInt + 1
'    If myInt >= tmOutInterval Then
'    myInt = 0
'    tell "Reseting connections"
'       If pci(1).state = "FREE" Then RestartWS 1
'       If pci(2).state = "FREE" Then RestartWS 2
'       If pci(3).state = "FREE" Then RestartWS 3
'       If pci(4).state = "FREE" Then RestartWS 4
'       If pci(5).state = "FREE" Then RestartWS 5
'       If pci(6).state = "FREE" Then RestartWS 6
'
'    End If
'
'    .Enabled = True
'End With
'
'End Sub
'
'Private Sub tmShowINfo_Timer()
'Me.tmShowINfo.Enabled = False
'Dim tot&, usd&, pcu&, netu&
'
'Dim ind As Integer
'
'For ind = 1 To 6
'If ind > getLicense(4863) Then Exit For
'DoEvents
'    With mCards(ind)
'        If .flag <> "0" Then
'
'
'        tot& = CLng(Val(.tbal))
'        usd& = CLng(Val(.tusd))
'        tellWS ind - 1, pci(ind).pcName
'        Me.lstRemData(ind - 1).ListItems(1).SubItems(1) = pci(ind).clientID
'        'login time
'        Me.lstRemData(ind - 1).ListItems(2).SubItems(1) = pci(ind).login
'        'logoff time
'
'            Me.lstRemData(ind - 1).ListItems(3).SubItems(1) = pci(ind).logoff
'            'preço windows
'        Me.lstRemData(ind - 1).ListItems(4).SubItems(1) = Format(prcSetup.Pwindows * 100, "0$00") & " Per Hour"
'
'        'tempo de uso windows
''        pcu& = CLng(Val(replacepci(1).pcuPrice))
''        netu& = CLng(Val(pci(1).netPrice))
''
'        Me.lstRemData(ind - 1).ListItems(5).SubItems(1) = pci(ind).pcuTime
'        'preço windows
'        Me.lstRemData(ind - 1).ListItems(6).SubItems(1) = pci(ind).pcuPrice
'        'preço internet offset
'        Me.lstRemData(ind - 1).ListItems(7).SubItems(1) = Format(prcSetup.Pnet * 100, "0$00") & " Per MB [" & Format(prcSetup.offSet \ 2, "-0.00") & "]"
'          'download net
'        Me.lstRemData(ind - 1).ListItems(8).SubItems(1) = pci(ind).netCharge
'        'preço net
'        Me.lstRemData(ind - 1).ListItems(9).SubItems(1) = pci(ind).netPrice
'        'used code
'            Me.lstRemData(ind - 1).ListItems(10).SubItems(1) = .code
'        'total balance
'            Me.lstRemData(ind - 1).ListItems(11).SubItems(1) = Format(tot& * 100, "0$00")
'        'used balance
'        Dim tsr$
'
'        pcu& = Val(Format(Replace$(pci(ind).pcuPrice, "$", ","), "0"))
'        netu& = Val(Format(Replace$(pci(ind).netPrice, "$", ","), "0"))
'        usd& = usd& + pcu& + netu&
'
'            Me.lstRemData(ind - 1).ListItems(12).SubItems(1) = Format(usd& * 100, "0$00")
'        'remain balance
'            Me.lstRemData(ind - 1).ListItems(13).SubItems(1) = Format((tot& - usd&) * 100, "0$00")
'        'bytes downb
'        Me.lstRemData(ind - 1).ListItems(14).SubItems(1) = trasnBytes(CLng(Val(.bytes)))
'        End If
'    End With
'
'
'Next
'
'
'
'
'
'Me.tmShowINfo.Enabled = True
'
'End Sub
'
'Private Sub tmTime_Timer()
'DoEvents
'Me.state.Panels("Time") = Format(Now, "dd/mmm/yyyy hh:mm ")
'Me.state.Panels("msg") = getDataBaseDetails(3245) & " Prices: Net - " & _
'Replace(Format(getPrices.Pnet, "#,##0.00"), ",", "$") & " PC per Hour - " & _
'Replace(Format(getPrices.Pwindows, "#,##0.00"), ",", "$")
'End Sub
'
'
'Private Sub SetupFrames()
''organiza os frames
'Dim it%
'For it = 0 To Me.fmePc.Count - 1
'    With Me.fmePc(it)
'        .Font.Bold = True
'        tellWS it, "Starting..."
'    End With
'Next
'End Sub
'Private Sub setupListsView()
''organiza os listviews
'Dim it%
'For it = 0 To Me.lstRemData.Count - 1
'    With Me.lstRemData(it)
'    .BackColor = vbBlack
'    .ForeColor = vbYellow
'    .HideSelection = True
''    .HideColumnHeaders = True
'    Me.lstRemData(it).ColumnHeaders.Add , , "", 0.42 * .Width
'    Me.lstRemData(it).ColumnHeaders.Add , , "", 0.54 * .Width
'    .View = lvwReport
'    .Font.Name = "Arial"
'    .Font.Size = 8
'    .GridLines = False
'    .LabelEdit = lvwManual
'
'        .ListItems.Add , , "Client ID"
'        .ListItems(1).Bold = True
'        .ListItems(1).ForeColor = vbWhite
'        .ListItems.Add , , "...Login"
'        .ListItems.Add , , "...Logoff"
'        .ListItems.Add , , "Windows"
'        .ListItems(4).Bold = True
'        .ListItems(4).ForeColor = vbWhite
'        .ListItems.Add , , "...Used Time"
'        .ListItems.Add , , "...Price"
'        .ListItems.Add , , "Internet"
'        .ListItems(7).Bold = True
'        .ListItems(7).ForeColor = vbWhite
'        .ListItems.Add , , "...Charge"
'        .ListItems.Add , , "...Price"
'        .ListItems.Add , , "Used Code"
'        .ListItems(10).Bold = True
'        .ListItems(10).ForeColor = vbWhite
'        .ListItems.Add , , "...Balance"
'        .ListItems.Add , , "...Used"
'        .ListItems.Add , , "...Remain"
'        .ListItems.Add , , "...Total Charge"
'
'        If it > 0 Then
'            Me.lstRemData(it).Width = Me.lstRemData(0).Width
'            Me.lstRemData(it).Height = Me.lstRemData(0).Height
'
'            Me.fmePc(it).Width = Me.fmePc(0).Width
'            Me.fmePc(it).Height = Me.fmePc(0).Height
'
'            Me.fmePc(3).Top = Me.fmePc(0).Top + Me.fmePc(0).Height + 40
'            Select Case it
'                Case 1, 2
'                    Me.fmePc(it).Top = Me.fmePc(0).Top
'                    Me.fmePc(it).Left = Me.fmePc(it - 1).Left + Me.fmePc(it - 1).Width + 40
'                Case 4, 5
'                    Me.fmePc(it).Top = Me.fmePc(3).Top
'                    Me.fmePc(it).Left = Me.fmePc(it - 1).Left + Me.fmePc(it - 1).Width + 40
'            End Select
'
'
'        End If
'
'    End With
'
'    If it + 1 > getLicense(4863) Then
'        Me.fmePc(it).Enabled = False
'
'        Select Case it
'            Case 0
'                Me.tmCl1.Enabled = False
'                Me.wsCl1.Close
'            Case 1
'                Me.tmCl2.Enabled = False
'                Me.wsCl2.Close
'            Case 2
'                Me.tmCl3.Enabled = False
'                Me.wsCl3.Close
'            Case 3
'                Me.tmCl4.Enabled = False
'                Me.wsCl4.Close
'            Case 4
'                Me.tmCl5.Enabled = False
'                Me.wsCl5.Close
'            Case 5
'                Me.tmCL6.Enabled = False
'                Me.wsCl6.Close
'
'
'        End Select
'
'        Me.lstRemData(it).BackColor = vbButtonFace
'        tellWS it, "No License Found."
'
'    End If
'
'Next
'End Sub
'
'Private Sub tmUpdInfo1_Timer()
'With Me.lstRemData(0)
'
'        Select Case pci(1).state
'            Case Is = "ON"
'                .ListItems(1).SubItems(1) = pci(1).clientID
'            Case Is = "OFF"
'                'ClearList Me.lstRemData(0)
'            Case Is = "SAVE"
'                'User is going out...
'
'        End Select
'
'
'End With
'
'End Sub
'
'Private Sub tmUpdInfo2_Timer()
'With Me.lstRemData(1)
'
'        Select Case pci(2).state
'            Case Is = "ON"
'                .ListItems(1).SubItems(1) = pci(2).clientID
'            Case Is = "OFF"
'                'ClearList Me.lstRemData(1)
'            Case Is = "SAVE"
'                'User is going out...
'
'        End Select
'
'
'End With
'End Sub
'
'Private Sub tmUpdInfo3_Timer()
'With Me.lstRemData(2)
'
'        Select Case pci(3).state
'            Case Is = "ON"
'                .ListItems(1).SubItems(1) = pci(3).clientID
'            Case Is = "OFF"
'                'ClearList Me.lstRemData(2)
'            Case Is = "SAVE"
'                'User is going out...
'
'        End Select
'
'
'End With
'End Sub
'
'Private Sub tmUpdInfo4_Timer()
'With Me.lstRemData(3)
'
'        Select Case pci(4).state
'            Case Is = "ON"
'                .ListItems(1).SubItems(1) = pci(4).clientID
'            Case Is = "OFF"
'                'ClearList Me.lstRemData(3)
'            Case Is = "SAVE"
'                'User is going out...
'
'        End Select
'
'
'End With
'End Sub
'
'Private Sub tmUpdInfo5_Timer()
'With Me.lstRemData(4)
'
'        Select Case pci(5).state
'            Case Is = "ON"
'                .ListItems(1).SubItems(1) = pci(5).clientID
'            Case Is = "OFF"
'                'ClearList Me.lstRemData(4)
'            Case Is = "SAVE"
'                'User is going out...
'
'        End Select
'
'
'End With
'End Sub
'
'Private Sub tmUpdInfo6_Timer()
'With Me.lstRemData(5)
'
'        Select Case pci(6).state
'            Case Is = "ON"
'                .ListItems(1).SubItems(1) = pci(6).clientID
'            Case Is = "OFF"
'                'ClearList Me.lstRemData(5)
'            Case Is = "SAVE"
'                'User is going out...
'
'        End Select
'
'
'End With
'End Sub
'
'
''CONNECTIONS REQUESTED
'Private Sub wsCl1_ConnectionRequest(ByVal requestID As Long)
'tell "1#, Connection Requested From " & wsCl1.RemoteHostIP
'tellWS 0, "Waiting Permission [" & wsCl1.RemoteHostIP & "]"
'    With wsCl1
'        If .state <> 7 Then
'            .Close
'            .Accept requestID
'            If Me.wsCl1.state = 7 Then
'                Me.wsCl1.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
'                "Welcome. The Permission Timer Was Started..." & vbCrLf
'                'request permission
'                Me.tmOut1.Enabled = False
'                Me.tmOut1.Enabled = True
'            End If
'
'        End If
'    End With
'
'End Sub
'
''DATA ARRIVAL
'Private Sub wsCl1_DataArrival(ByVal bytesTotal As Long)
'Dim pdta As String
'With wsCl1
'    If .state = sckConnected Then
'        .GetData pdta
'        InterpretData1 pdta
'    End If
'End With
'
'End Sub
'Private Sub wsCl2_DataArrival(ByVal bytesTotal As Long)
'Dim pdta As String
'With wsCl2
'    If .state = sckConnected Then
'        .GetData pdta
'        InterpretData2 pdta
'    End If
'End With
'
'End Sub
'Private Sub wsCl3_DataArrival(ByVal bytesTotal As Long)
'Dim pdta As String
'With wsCl3
'    If .state = sckConnected Then
'        .GetData pdta
'        InterpretData3 pdta
'    End If
'End With
'
'End Sub
'Private Sub wsCl4_DataArrival(ByVal bytesTotal As Long)
'Dim pdta As String
'With wsCl4
'    If .state = sckConnected Then
'        .GetData pdta
'        InterpretData4 pdta
'    End If
'End With
'
'End Sub
'Private Sub wsCl5_DataArrival(ByVal bytesTotal As Long)
'Dim pdta As String
'With wsCl5
'    If .state = sckConnected Then
'        .GetData pdta
'        InterpretData5 pdta
'    End If
'End With
'
'End Sub
'Private Sub wsCl6_DataArrival(ByVal bytesTotal As Long)
'Dim pdta As String
'With wsCl6
'    If .state = sckConnected Then
'        .GetData pdta
'        InterpretData6 pdta
'    End If
'End With
'
'End Sub
'
'Private Sub wsCl2_ConnectionRequest(ByVal requestID As Long)
'tell "2#, Connection Requested From " & wsCl2.RemoteHostIP
'tellWS 1, "Waiting Permission [" & wsCl2.RemoteHostIP & "]"
'    With wsCl2
'        If .state <> 7 Then
'            .Close
'            .Accept requestID
'            If Me.wsCl2.state = 7 Then
'                Me.wsCl2.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
'                "Welcome. The Permission Timer Was Started..." & vbCrLf
'                Me.tmOut2.Enabled = False
'                Me.tmOut2.Enabled = True
'            End If
'
'        End If
'    End With
'
'End Sub
'
'Private Sub wscl3_ConnectionRequest(ByVal requestID As Long)
'tell "3#, Connection Requested From " & wsCl3.RemoteHostIP
'tellWS 2, "Waiting Permission [" & wsCl3.RemoteHostIP & "]"
'    With wsCl3
'        If .state <> 7 Then
'            .Close
'            .Accept requestID
'            If Me.wsCl3.state = 7 Then
'                Me.wsCl3.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
'                "Welcome. The Permission Timer Was Started..." & vbCrLf
'                Me.tmOut3.Enabled = False
'                Me.tmOut3.Enabled = True
'            End If
'
'        End If
'    End With
'
'End Sub
'
'Private Sub wscl4_ConnectionRequest(ByVal requestID As Long)
'tell "4#, Connection Requested From " & wsCl4.RemoteHostIP
'tellWS 3, "Waiting Permission [" & wsCl4.RemoteHostIP & "]"
'    With wsCl4
'        If .state <> 7 Then
'            .Close
'            .Accept requestID
'            If Me.wsCl4.state = 7 Then
'                Me.wsCl4.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
'                "Welcome. The Permission Timer Was Started..." & vbCrLf
'                Me.tmOut4.Enabled = False
'                Me.tmOut4.Enabled = True
'            End If
'
'        End If
'    End With
'
'End Sub
'
'Private Sub wscl5_ConnectionRequest(ByVal requestID As Long)
'tell "5#, Connection Requested From " & wsCl5.RemoteHostIP
'tellWS 4, "Waiting Permission [" & wsCl5.RemoteHostIP & "]"
'    With wsCl5
'        If .state <> 7 Then
'            .Close
'            .Accept requestID
'            If Me.wsCl5.state = 7 Then
'                Me.wsCl5.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
'                "Welcome. The Permission Timer Was Started..." & vbCrLf
'                Me.tmOut5.Enabled = False
'                Me.tmOut5.Enabled = True
'            End If
'
'        End If
'    End With
'
'End Sub
'
'Private Sub wscl6_ConnectionRequest(ByVal requestID As Long)
'tell "6#, Connection Requested From " & wsCl6.RemoteHostIP
'tellWS 5, "Waiting Permission [" & wsCl6.RemoteHostIP & "]"
'    With wsCl6
'        If .state <> 7 Then
'            .Close
'            .Accept requestID
'            If Me.wsCl6.state = 7 Then
'                Me.wsCl6.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
'                "Welcome. The Permission Timer Was Started..." & vbCrLf
'                Me.tmOut6.Enabled = False
'                Me.tmOut6.Enabled = True
'            End If
'
'        End If
'    End With
'
'End Sub
'
''PERMISSSION TIMER....
'Private Sub wsCl1_SendComplete()
'flgWaitWs1 = False
'End Sub
'Private Sub wsCl2_SendComplete()
'flgWaitWs2 = False
'End Sub
'Private Sub wsCl3_SendComplete()
'flgWaitWs3 = False
'End Sub
'Private Sub wsCl4_SendComplete()
'flgWaitWs4 = False
'End Sub
'Private Sub wsCl5_SendComplete()
'flgWaitWs5 = False
'End Sub
'Private Sub wsCl6_SendComplete()
'flgWaitWs6 = False
'End Sub
'
''SENDPROGRESS
''Private Sub wsCl1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
''tellWS 0, bytesSent & " Bytes Sent to " & wsCl1.RemoteHost
''End Sub
''Private Sub wsCl2_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
''tellWS 1, bytesSent & " Bytes Sent to " & wsCl2.RemoteHost
''End Sub
''Private Sub wsCl3_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
''tellWS 2, bytesSent & " Bytes Sent to " & wsCl3.RemoteHost
''End Sub
''Private Sub wsCl4_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
''tellWS 3, bytesSent & " Bytes Sent to " & wsCl4.RemoteHost
''End Sub
''Private Sub wsCl5_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
''tellWS 4, bytesSent & " Bytes Sent to " & wsCl5.RemoteHost
''End Sub
''Private Sub wsCl6_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
''tellWS 5, bytesSent & " Bytes Sent to " & wsCl6.RemoteHost
''End Sub
'
'
''SUBS AND FUNCTIONS FOR ALL
''mostra titulo do frame...
'Private Sub tellWS(fme As Integer, arg As String)
'Me.fmePc(fme).Caption = arg$
'End Sub
'
'Private Sub RestartWS(Index As Integer)
'
'Dim ws As Object
'For Each ws In frmsv2
'DoEvents
'Debug.Print ws.Name
'If ws.Name = "wsCl" & Index Then
'    ws.Close
'    ws.Listen
'Exit For
'End If
'
'Next
'
'End Sub
'
''INTERPRET DATA
'Private Sub InterpretData1(dta As String)
'Dim td$
'With pci(1)
'dta$ = Replace(dta, vbCrLf, "")
'Debug.Print "Recv " & dta
'    Select Case UCase(Left(dta, 4))
'        Case Is = "NEWM"
'            Me.tmOut1.Enabled = False
'            flgWaitWs1 = False
'            dta = Replace(dta, vbCrLf, "")
'                .pcName = Right(dta, Len(dta) - 4)
'                .state = "ON"
'            tellWS 0, .pcName & " Online..."
'
'            'send configurations here
'
'          Case Is = "KILL"
'            .state = "OFF"
'            If wsCl1.state = 7 Then
'            wsCl1.SendData "OKOUT" & vbCrLf
'            RestartWS 1
'            End If
'            'from here
'          Case "CFG?"
'            If wsCl1.state = 7 Then
'                wsCl1.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
'                Format(CLng(getPrices.Pwindows), "0000") & _
'                Format(CLng(getPrices.offSet), "0000") & vbCrLf
'            End If
'          Case "CODE"
'          'ver se está a ser usado por algum cliente...
'          Dim tc$
'          tc$ = Right(dta, Len(dta) - 4)
'          tellWS 0, "Validating " & tc$
'            If isCodeInUse(tc$, 1) = True Then
'                If wsCl1.state = 7 Then
'                    wsCl1.SendData "BUSY"
'                End If
'
'            Else
'                    'validar o codigo
'               If wsCl1.state = 7 Then
'                   'wsCl1.SendData
'                   wsCl1.SendData valCod(tc$, 1)
'               End If
'
'
'            End If
'          Case "OUTM"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                  addUpdateCard mCards(1).code, td$
'
'                  Dim dtl As DETAILS
'
'
'
'                  pci(1).logoff = Format(Now, "dd/mm/yy hh:mm:ss")
'                  pci(1).state = "FREE  "
'
'                  dtl.pc = pci(1).pcName
'                  dtl.din = pci(1).login
'                  dtl.dout = pci(1).logoff
'                  dtl.Nick = pci(1).clientID
'                  dtl.tmv = pci(1).pcuTime
'                  With mCards(1)
'                      dtl.data = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
'                  End With
'                  addDetails dtl
'            End If
'          Case "SAVE"
'          td$ = Right(dta$, Len(dta$) - 4)
'          If Len(td) > 51 Then td = Left(td, 51)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(1).code, td$
'            End If
'          Case "DTLS"
'          td$ = Right(dta$, Len(dta$) - 4)
'          If Len(td$) > 31 Then td = Left(td, 32)
'            If Len(td$) = 32 Then
'                CopyMemory tmde(1), ByVal td$, 32
'
''                huso As String * 8  '00:00:00
''                win As String * 6   '000$00
''                netc As String * 12  '1024 KB
''                netp As String * 6  '000$00
'    'Cybero
'                    pci(1).pcuTime = tmde(1).huso
'                    pci(1).pcuPrice = Format(CDbl(Trim$(tmde(1).win) * 100), "0$00")
'                    pci(1).netCharge = trasnBytes(CLng(Trim$(tmde(1).netc)))
'                    pci(1).netPrice = IIf(tmde(1).netp <> "Cybero", Format(CDbl("0" & Val(tmde(1).netp)) * 100, "0$00"), "Cyber Offset")
'            End If
'           Case "USER"
'           pci(1).clientID = Trim$(Right(dta$, Len(dta$) - 4))
'           pci(1).login = Format(Now, "dd/mm/yy hh:mm:ss")
'           pci(1).logoff = ""
'           pci(1).state = "BUSY"
'           Case "TYPE"
'
'    End Select
'End With
'
'End Sub
'
'Private Sub InterpretData2(dta As String)
'Dim td$
'With pci(2)
'dta$ = Replace(dta, vbCrLf, "")
'Debug.Print "Recv " & dta
'    Select Case UCase(Left(dta, 4))
'        Case Is = "NEWM"
'            Me.tmOut2.Enabled = False
'            flgWaitWs2 = False
'            dta = Replace(dta, vbCrLf, "")
'                .pcName = Right(dta, Len(dta) - 4)
'                .state = "ON"
'            tellWS 1, .pcName & " Online..."
'
'            'send configurations here
'
'          Case Is = "KILL"
'            .state = "OFF"
'            If wsCl2.state = 7 Then
'            wsCl2.SendData "OKOUT" & vbCrLf
'            RestartWS 2
'            End If
'            'from here
'          Case "CFG?"
'            If wsCl2.state = 7 Then
'                wsCl2.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
'                Format(CLng(getPrices.Pwindows), "0000") & _
'                Format(CLng(getPrices.offSet), "0000") & vbCrLf
'            End If
'          Case "CODE"
'          'ver se está a ser usado por algum cliente...
'          Dim tc$
'          tc$ = Right(dta, Len(dta) - 4)
'          tellWS 1, "Validating " & tc$
'            If isCodeInUse(tc$, 1) = True Then
'                If wsCl2.state = 7 Then
'                    wsCl2.SendData "BUSY"
'                End If
'
'            Else
'                    'validar o codigo
'               If wsCl2.state = 7 Then
'                   'wsCl1.SendData
'                   wsCl2.SendData valCod(tc$, 2)
'               End If
'
'
'            End If
'          Case "OUTM"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(2).code, td$
'            End If
'          pci(2).logoff = Format(Now, "dd/mm/yy,hh:mm:ss")
'          pci(2).state = "FREE  "
'          Case "SAVE"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(2).code, td$
'            End If
'          Case "DTLS"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 28 Then
'                CopyMemory tmde(2), ByVal td$, 28
'
''                huso As String * 8  '00:00:00
''                win As String * 6   '000$00
''                netc As String * 8  '1024 KB
''                netp As String * 6  '000$00
'    'Cybero
'                    pci(2).pcuTime = tmde(2).huso
'                    pci(2).pcuPrice = tmde(2).win
'                    pci(2).netCharge = tmde(2).netc
'                    pci(2).netPrice = tmde(2).netp
'            End If
'           Case "USER"
'           pci(2).clientID = Trim$(Right(dta$, Len(dta$) - 4))
'           pci(2).login = Format(Now, "dd/mm/yy hh:mm:ss")
'           pci(2).logoff = ""
'           pci(2).state = "BUSY"
'           Case "TYPE"
'
'    End Select
'End With
'End Sub
'
'Private Sub InterpretData3(dta As String)
'Dim td$
'With pci(3)
'dta$ = Replace(dta, vbCrLf, "")
'Debug.Print "Recv " & dta
'    Select Case UCase(Left(dta, 4))
'        Case Is = "NEWM"
'            Me.tmOut3.Enabled = False
'            flgWaitWs3 = False
'            dta = Replace(dta, vbCrLf, "")
'                .pcName = Right(dta, Len(dta) - 4)
'                .state = "ON"
'            tellWS 2, .pcName & " Online..."
'
'            'send configurations here
'
'          Case Is = "KILL"
'            .state = "OFF"
'            If wsCl3.state = 7 Then
'            wsCl3.SendData "OKOUT" & vbCrLf
'            RestartWS 3
'            End If
'            'from here
'          Case "CFG?"
'            If wsCl3.state = 7 Then
'                wsCl3.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
'                Format(CLng(getPrices.Pwindows), "0000") & _
'                Format(CLng(getPrices.offSet), "0000") & vbCrLf
'            End If
'          Case "CODE"
'          'ver se está a ser usado por algum cliente...
'          Dim tc$
'          tc$ = Right(dta, Len(dta) - 4)
'          tellWS 2, "Validating " & tc$
'            If isCodeInUse(tc$, 3) = True Then
'                If wsCl3.state = 7 Then
'                    wsCl3.SendData "BUSY"
'                End If
'
'            Else
'                    'validar o codigo
'               If wsCl3.state = 7 Then
'                   'wsCl1.SendData
'                   wsCl3.SendData valCod(tc$, 3)
'               End If
'
'
'            End If
'          Case "OUTM"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(3).code, td$
'            End If
'          pci(3).logoff = Format(Now, "dd/mm/yy,hh:mm:ss")
'          pci(3).state = "FREE  "
'          Case "SAVE"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(3).code, td$
'            End If
'          Case "DTLS"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 28 Then
'                CopyMemory tmde(3), ByVal td$, 28
'
''                huso As String * 8  '00:00:00
''                win As String * 6   '000$00
''                netc As String * 8  '1024 KB
''                netp As String * 6  '000$00
'    'Cybero
'                    pci(3).pcuTime = tmde(3).huso
'                    pci(3).pcuPrice = tmde(3).win
'                    pci(3).netCharge = tmde(3).netc
'                    pci(3).netPrice = tmde(3).netp
'            End If
'           Case "USER"
'           pci(3).clientID = Trim$(Right(dta$, Len(dta$) - 4))
'           pci(3).login = Format(Now, "dd/mm/yy hh:mm:ss")
'           pci(3).logoff = ""
'           pci(3).state = "BUSY"
'           Case "TYPE"
'
'    End Select
'End With
'
'End Sub
'
'Private Sub InterpretData4(dta As String)
'Dim td$
'With pci(4)
'dta$ = Replace(dta, vbCrLf, "")
'Debug.Print "Recv " & dta
'    Select Case UCase(Left(dta, 4))
'        Case Is = "NEWM"
'            Me.tmOut4.Enabled = False
'            flgWaitWs4 = False
'            dta = Replace(dta, vbCrLf, "")
'                .pcName = Right(dta, Len(dta) - 4)
'                .state = "ON"
'            tellWS 3, .pcName & " Online..."
'
'            'send configurations here
'
'          Case Is = "KILL"
'            .state = "OFF"
'            If wsCl4.state = 7 Then
'            wsCl4.SendData "OKOUT" & vbCrLf
'            RestartWS 4
'            End If
'            'from here
'          Case "CFG?"
'            If wsCl4.state = 7 Then
'                wsCl4.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
'                Format(CLng(getPrices.Pwindows), "0000") & _
'                Format(CLng(getPrices.offSet), "0000") & vbCrLf
'            End If
'          Case "CODE"
'          'ver se está a ser usado por algum cliente...
'          Dim tc$
'          tc$ = Right(dta, Len(dta) - 4)
'          tellWS 3, "Validating " & tc$
'            If isCodeInUse(tc$, 4) = True Then
'                If wsCl4.state = 7 Then
'                    wsCl4.SendData "BUSY"
'                End If
'
'            Else
'                    'validar o codigo
'               If wsCl4.state = 7 Then
'                   'wsCl1.SendData
'                   wsCl4.SendData valCod(tc$, 4)
'               End If
'
'
'            End If
'          Case "OUTM"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(4).code, td$
'            End If
'          pci(4).logoff = Format(Now, "dd/mm/yy,hh:mm:ss")
'          pci(4).state = "FREE  "
'          Case "SAVE"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(4).code, td$
'            End If
'          Case "DTLS"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 28 Then
'                CopyMemory tmde(4), ByVal td$, 28
'
''                huso As String * 8  '00:00:00
''                win As String * 6   '000$00
''                netc As String * 8  '1024 KB
''                netp As String * 6  '000$00
'    'Cybero
'                    pci(4).pcuTime = tmde(4).huso
'                    pci(4).pcuPrice = tmde(4).win
'                    pci(1).netCharge = tmde(4).netc
'                    pci(4).netPrice = tmde(4).netp
'            End If
'           Case "USER"
'           pci(4).clientID = Trim$(Right(dta$, Len(dta$) - 4))
'           pci(4).login = Format(Now, "dd/mm/yy hh:mm:ss")
'           pci(4).logoff = ""
'           pci(4).state = "BUSY"
'           Case "TYPE"
'
'    End Select
'End With
'End Sub
'
'Private Sub InterpretData5(dta As String)
'Dim td$
'With pci(5)
'dta$ = Replace(dta, vbCrLf, "")
'Debug.Print "Recv " & dta
'    Select Case UCase(Left(dta, 4))
'        Case Is = "NEWM"
'            Me.tmOut5.Enabled = False
'            flgWaitWs5 = False
'            dta = Replace(dta, vbCrLf, "")
'                .pcName = Right(dta, Len(dta) - 4)
'                .state = "ON"
'            tellWS 4, .pcName & " Online..."
'
'            'send configurations here
'
'          Case Is = "KILL"
'            .state = "OFF"
'            If wsCl5.state = 7 Then
'            wsCl5.SendData "OKOUT" & vbCrLf
'            RestartWS 5
'            End If
'            'from here
'          Case "CFG?"
'            If wsCl5.state = 7 Then
'                wsCl5.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
'                Format(CLng(getPrices.Pwindows), "0000") & _
'                Format(CLng(getPrices.offSet), "0000") & vbCrLf
'            End If
'          Case "CODE"
'          'ver se está a ser usado por algum cliente...
'          Dim tc$
'          tc$ = Right(dta, Len(dta) - 4)
'          tellWS 4, "Validating " & tc$
'            If isCodeInUse(tc$, 5) = True Then
'                If wsCl5.state = 7 Then
'                    wsCl5.SendData "BUSY"
'                End If
'
'            Else
'                    'validar o codigo
'               If wsCl5.state = 7 Then
'                   'wsCl1.SendData
'                   wsCl5.SendData valCod(tc$, 1)
'               End If
'
'
'            End If
'          Case "OUTM"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(5).code, td$
'            End If
'          pci(5).logoff = Format(Now, "dd/mm/yy,hh:mm:ss")
'          pci(5).state = "FREE  "
'          Case "SAVE"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(5).code, td$
'            End If
'          Case "DTLS"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 28 Then
'                CopyMemory tmde(5), ByVal td$, 28
'
''                huso As String * 8  '00:00:00
''                win As String * 6   '000$00
''                netc As String * 8  '1024 KB
''                netp As String * 6  '000$00
'    'Cybero
'                    pci(5).pcuTime = tmde(5).huso
'                    pci(5).pcuPrice = tmde(5).win
'                    pci(5).netCharge = tmde(5).netc
'                    pci(5).netPrice = tmde(5).netp
'            End If
'           Case "USER"
'           pci(5).clientID = Trim$(Right(dta$, Len(dta$) - 4))
'           pci(5).login = Format(Now, "dd/mm/yy hh:mm:ss")
'           pci(5).logoff = ""
'           pci(5).state = "BUSY"
'           Case "TYPE"
'
'    End Select
'End With
'End Sub
'
'Private Sub InterpretData6(dta As String)
'
'Dim td$
'With pci(6)
'dta$ = Replace(dta, vbCrLf, "")
'Debug.Print "Recv " & dta
'    Select Case UCase(Left(dta, 4))
'        Case Is = "NEWM"
'            Me.tmOut6.Enabled = False
'            flgWaitWs6 = False
'            dta = Replace(dta, vbCrLf, "")
'                .pcName = Right(dta, Len(dta) - 4)
'                .state = "ON"
'            tellWS 5, .pcName & " Online..."
'
'            'send configurations here
'
'          Case Is = "KILL"
'            .state = "OFF"
'            If wsCl6.state = 7 Then
'            wsCl6.SendData "OKOUT" & vbCrLf
'            RestartWS 6
'            End If
'            'from here
'          Case "CFG?"
'            If wsCl6.state = 7 Then
'                wsCl6.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
'                Format(CLng(getPrices.Pwindows), "0000") & _
'                Format(CLng(getPrices.offSet), "0000") & vbCrLf
'            End If
'          Case "CODE"
'          'ver se está a ser usado por algum cliente...
'          Dim tc$
'          tc$ = Right(dta, Len(dta) - 4)
'          tellWS 0, "Validating " & tc$
'            If isCodeInUse(tc$, 6) = True Then
'                If wsCl6.state = 7 Then
'                    wsCl6.SendData "BUSY"
'                End If
'
'            Else
'                    'validar o codigo
'               If wsCl6.state = 7 Then
'                   'wsCl1.SendData
'                   wsCl6.SendData valCod(tc$, 6)
'               End If
'
'
'            End If
'          Case "OUTM"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(6).code, td$
'            End If
'          pci(6).logoff = Format(Now, "dd/mm/yy,hh:mm:ss")
'          pci(6).state = "FREE  "
'          Case "SAVE"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'                addUpdateCard mCards(6).code, td$
'            End If
'          Case "DTLS"
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 28 Then
'                CopyMemory tmde(6), ByVal td$, 28
'
'
'            End If
'           Case "USER"
'           pci(6).clientID = Trim$(Right(dta$, Len(dta$) - 4))
'           pci(6).login = Format(Now, "dd/mm/yy hh:mm:ss")
'           pci(6).logoff = ""
'           pci(6).state = "BUSY"
'           Case "TYPE"
'
'    End Select
'End With
'
'End Sub
'
''limpa os dados da lista
'Private Sub ClearList(lst As ListView)
'Dim it%
'    For it% = 1 To lst.ListItems.Count
'    DoEvents
'        lst.ListItems(it%).SubItems(1) = ""
'    Next
'
'End Sub
'
''ver se cliente esta a usar o mesmo codigo
'Private Function isCodeInUse(ByVal cod$, Index&) As Boolean
'Dim it&
'For it& = 1 To 6
'    If Index <> it Then
'        If mCards(it).code = cod$ Then
'                isCodeInUse = True
'                Exit Function
'        End If
'    End If
'
'Next
'
'isCodeInUse = False
'End Function
'
'
'
'
Private Sub tmShowINfo_Timer()

End Sub
