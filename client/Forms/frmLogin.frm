VERSION 5.00
Begin VB.Form frmlogc 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "####-####-####-####"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2070
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   5130
      MaxLength       =   19
      TabIndex        =   5
      Top             =   3270
      Width           =   2265
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   660
      Top             =   4320
   End
   Begin VB.Timer tmgetOrNot 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   5220
      Top             =   660
   End
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   1
      Left            =   -1740
      MaxLength       =   16
      TabIndex        =   2
      Text            =   "Anonymous"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Notice!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   180
      Index           =   0
      Left            =   5010
      MouseIcon       =   "frmLogin.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "Click to read some informations"
      Top             =   3900
      Width           =   480
   End
   Begin VB.Label lbip 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
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
      Left            =   11340
      TabIndex        =   11
      Top             =   7650
      UseMnemonic     =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by microbodix@hotmail.com"
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
      Height          =   210
      Left            =   5250
      TabIndex        =   10
      Top             =   3630
      Width           =   2010
   End
   Begin VB.Label status 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2940
      TabIndex        =   9
      Top             =   7500
      Width           =   1095
   End
   Begin VB.Label updDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Light version march 2008-rev 2009"
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
      Height          =   210
      Left            =   4995
      TabIndex        =   8
      Top             =   3030
      Width           =   2535
   End
   Begin VB.Label cmdMenu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   300
      Left            =   5040
      TabIndex        =   7
      Top             =   3270
      Width           =   2445
   End
   Begin VB.Label lbtime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Left            =   12330
      TabIndex        =   6
      Top             =   30
      UseMnemonic     =   0   'False
      Width           =   210
   End
   Begin VB.Shape sol 
      BorderColor     =   &H80000001&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   570
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   1470
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label cmdUser 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " LOGIN "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   0
      Left            =   2220
      MouseIcon       =   "frmLogin.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4950
      Shape           =   2  'Oval
      Top             =   1500
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   4920
      TabIndex        =   3
      Top             =   -570
      Width           =   45
   End
   Begin VB.Label cmdUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " LOGOFF "
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
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   1
      Left            =   3300
      TabIndex        =   0
      Top             =   870
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label cmdUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " CANCEL"
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
      Height          =   210
      Index           =   2
      Left            =   3540
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmlogc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private glIndex As Long
Private cont As Long
Dim centx#, centy#, raio#
Const PI = 3.1415

Private Type POINT
xx As Long
yy As Long
End Type
Private pt As POINT

'
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Sub SetTransparency()
Const G_E = &HFFEC
Const W_E = &H80000
Const LW_KEY = &H1
SetWindowLong Me.hwnd, G_E, GetWindowLong(Me.hwnd, G_E) Or W_E
SetLayeredWindowAttributes Me.hwnd, RGB(255, 0, 0), 0, LW_KEY
'SetLayeredWindowAttributes Me.hwnd, RGB(254, 0, 0), 0, LW_KEY

End Sub

Private Sub Form_Initialize()
Call SetTransparency
End Sub



Private Sub cmdMenu_DblClick()
frmAg.showMENU
End Sub

Public Sub cmdUser_Click(Index As Integer)
Select Case Index
Case 0
Me.tmgetOrNot.Enabled = False
Me.txtData(0).Enabled = False
Me.txtData(1).Enabled = False
Me.cmdUser(0).Enabled = False

actCodeMem = Me.txtData(0).Text
actUser = Me.txtData(1).Text
Me.txtData(0).Text = ""

tell "Setting up..."
Pause 2

If CInt(GetSetting(App.EXEName, "Remote", "SellDownload", 1)) = 1 Then
    If sniffer.installSniffer(-1) = 0 Then
        frmCli.tmGetNet.Enabled = True
    Else
        tell "IP#A. ErrCode(" & Hex$(Err.LastDllError) & ")"
        Pause 2
        Me.tmgetOrNot.Enabled = True
        Me.txtData(0).Enabled = True
        Me.txtData(1).Enabled = True
        Me.cmdUser(0).Enabled = True

    Exit Sub
    End If
End If

If Loggin(actCodeMem) = True Then
       changeWallPaper
       installcCyberDesktop
    Pause 0.4
    Unload Me
    Else
    Me.txtData(0).Text = actCodeMem
    Me.txtData(0).SelStart = 0
    Me.txtData(0).SelLength = Len(Me.txtData(0).Text)
    Me.txtData(0).Enabled = True
    Me.txtData(1).Enabled = True
    Me.tmgetOrNot.Enabled = True
End If
UpdateMyDate

Case 1

Me.cmdUser(1).Enabled = False
MerlinDologgoff actCodeMem
Case 2
'Me.cmdUser(0).Enabled = False
'Me.cmdUser(1).Enabled = False
'Me.cmdUser(2).Enabled = False
UpdateMyDate
Me.txtData(0).Text = ""
Me.txtData(1).Text = "Anonymous"

'Unload Me
'killKnowProcesses GetCurrentProcessId
'correctMerlinState

End Select

End Sub

Private Sub cmdUser_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.cmdUser(Index).ForeColor <> vbWhite Then Me.cmdUser(Index).ForeColor = vbWhite
'
' Select Case Index
'    Case 0
'        tell "After you specify a code, click here to Get in."
'    Case 1
'        tell "if you are already logged, click here to end your Session."
'    Case 2
'        tell "Click here to hide this Window."
'
' End Select
'
End Sub

Private Sub Form_Load()
'Me.Picture = Me.imgoff.Picture
topMost HWND_TOPMOST, Me
Me.tmgetOrNot.Enabled = True
Me.Status.Font.Name = "Arial"
Me.Status.Font.Bold = True
tell "Enter Code..."


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pt.xx = X
pt.yy = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'limpar realce
 If Button = 0 Then
    If Me.cmdUser(0).ForeColor <> &HE0E0E0 Then Me.cmdUser(0).ForeColor = &HE0E0E0
    If Me.cmdUser(1).ForeColor <> &HE0E0E0 Then Me.cmdUser(1).ForeColor = &HE0E0E0
    If Me.cmdUser(2).ForeColor <> &HE0E0E0 Then Me.cmdUser(2).ForeColor = &HE0E0E0
 ElseIf Button = 1 Then
    Me.Left = Me.Left + (X - pt.xx)
    Me.Top = Me.Top + (Y - pt.yy)
 End If
 
 
End Sub



Private Sub setMyPos(obj As Variant)
obj.Top = centx
centx = centx + 40
obj.Left = centy - (obj.width / 2)
End Sub

Private Sub Form_Resize()
centx = 1
centy = 1
End Sub

Private Sub Label3_Click(Index As Integer)
If Index = 0 Then showTips rolante & ". Tx and " & beGoodBoy, 10
End Sub



Private Sub Timer1_Timer()
DoEvents
Label2.Visible = False
sol.Left = sol.Left + (Math.Rnd * 100) * (centx) * Screen.TwipsPerPixelX + (Me.ScaleWidth / Me.ScaleHeight)
sol.Top = sol.Top + (Math.Rnd * 100) * (centy) * Screen.TwipsPerPixelY
Label2.Visible = True
    
    If sol.Top > Me.height - 1.4 * Label2.height Then
    centy = -centy
    sol.Top = Me.height - 1.4 * Label2.height
    End If
    
    If sol.Top < 100 Then
    centy = -centy
    sol.Top = 100
    End If
    
    
    If sol.Left > (Me.ScaleWidth - Label2.width) - 100 Then
    centx = -centx
    sol.Left = (Me.ScaleWidth - Label2.width) - 100
    End If
    
    If sol.Left < 500 Then
    centx = -centx
    sol.Left = 500
    End If
    
    Label2.Left = sol.Left - Label2.width / 2
    Label2.Top = sol.Top
End Sub

Private Sub tmgetOrNot_Timer()
Static hkill As Long, id As Integer, ang#

DoEvents
Me.lbtime = Format(Now, "mm\/dd yyyy hh:mm:ss ") & Format(Right(CStr(Timer) * 100, 2), "00")

If myPrcNet <> Me.Label3(1).Caption Then
'Me.Label3(1).Caption = myPrcNet
rolante = "After login press CTRL + ALT any time to logoff, use ALTGR to Enter @. " & _
myPrcNet
'Me.Label3(0).Caption = rolante
End If

If Me.lbip <> sniffer.localIP Then Me.lbip = sniffer.localIP

    'If (hkill Mod 6) = 0 Then
    
    'End If
sol.Visible = False

If hkill >= 200 Then
hkill = 0
If debugMode = False Then
UpdateMyDate
killVisibleProcesses
End If

End If
hkill = hkill + 1
End Sub

Private Sub tmTime_Timer()

'tell "Local Date: " & UCase(Format(Now, "dd/mm/yyyy, hh:mm:ss")) & " [4 sec] Login " & Format((120 - cont) / 120, "0%")
End Sub


Private Sub txtData_Change(Index As Integer)
cont = 0
Select Case Index

    Case 1
    If Me.txtData(1).Text = "" Then
        Me.txtData(1) = "Anonymous"
        Me.txtData(1).SelStart = 1
        Me.txtData(1).SelLength = Len(Me.txtData(1).Text)
    End If
    
End Select

End Sub

Private Sub txtData_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
frmCli.Talk "TYPE"

If KeyCode <> 8 And KeyCode <> 13 Then
    Select Case Index
    Case 0
       
        With Me.txtData(Index)
        Me.cmdUser(0).Enabled = False
            Select Case Len(.Text)
                Case 4, 9, 14
                    If Right(.Text, 1) <> "-" Then .Text = .Text & "-"
                    .SelStart = Len(.Text)
                Case 19
                    Me.cmdUser(0).Enabled = True
                    tell "ENTER to Login"
            End Select
        End With
        Case 1
        If Len(Me.txtData(1).Text) = 0 Then
            Me.txtData(1).Text = "Anonymous"
            Me.txtData(1).SelStart = 0
            Me.txtData(1).SelLength = Len(Me.txtData(1).Text)
        End If
    Case 13
    
    
    End Select
    ElseIf KeyCode = 13 Then
    If Len(Me.txtData(0).Text) = 19 Then
    If Me.txtData(1).Text = "" Then Me.txtData(1).Text = "Anonymous"
    Me.cmdUser_Click 0
    End If
End If

End Sub

Private Sub txtData_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Select Case Index
'    Case 0
'        tell "Enter your code here, without minus signal."
'    Case 1
'        tell "Our Staff may use it to contact you, during your Session!"
'End Select
'0429-9783-7810-3000
End Sub

Private Sub updDate_DblClick()
If GetAsyncKeyState(VK_LCONTROL) <> 0 Then
 tryUpdate
End If

End Sub
