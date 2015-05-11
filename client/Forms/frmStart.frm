VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStart.frx":0000
   ScaleHeight     =   885
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "copyright© 2003-2008 edson martins v.2008"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   2940
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   60
      Left            =   3450
      Top             =   780
      Width           =   1755
   End
   Begin VB.Label lbpg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5145
   End
   Begin VB.Image pgb 
      Height          =   105
      Left            =   3450
      Picture         =   "frmStart.frx":136CE
      Top             =   780
      Width           =   1755
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cliccount As Integer
Private Sub Form_Load()

If goend = True Then Unload Me
Dim arrarg As Variant

debugMode = False



If LCase(Command$) = "-r" Then
clearPolicies
Unload frmIcon
End
End If




Load frmIcon
If UCase(Environ$("COMPUTERNAME")) = "EDSONPORTABLE" Then
debugMode = MsgBox("Warning!!! Run in Test Mode?", vbQuestion + vbYesNo, "How to run here?") = vbYes ' true test mode False real mode
End If

If LCase(Command$) = "-t" Then
debugMode = True
End If

 topMost HWND_TOPMOST, Me
 

Me.Icon = frmCli.Icon
Me.Caption = IIf(0 = 0, "Loading...", "Unloading...")
Me.lbpg.Caption = Me.Caption
HideCaption Me.hwnd

Me.height = 975
Me.pgb.width = 0
Do
DoEvents
Me.Visible = True
Loop Until Me.Visible = True

Do
DoEvents
Pause 0.0002
Me.pgb.height = 60
Me.pgb.width = Me.pgb.width + 4

If Me.pgb.width > 1740 Then
    'AnimateWin frmsv2.hWnd, 600, AW_BLEND + AW_HIDE
    Me.lbpg = "Preparing for User Mode..."
    changeWallPaper
End If

'Debug.Print Me.pgb.width
If Me.pgb.width = 391 Then
  ' If operation = 0 Then If DetectOtherInstances("ed") = True Then End
    Me.lbpg = "Reading IP Address"
Call sniffer.installSniffer(0) 'just to get IP address
 createFolders
End If

If Me.pgb.width = 403 Then
Me.lbpg = "Reading user Security Identifier..."
userSID = getUserSID()
Me.lbpg = userSID & "..."
If userSID <> getUserSID() Then
End
End If

End If



If Me.pgb.width = 467 Then
Me.lbpg = "Applying policy..."
If cCyberXV2FLG = 0 Then
cCyberXV2FLG = 15
  
Call AplyPolicies
cCyberXV2FLG = 0
Else
Call AplyPolicies
End If

End If

If Me.pgb.width > 591 And Me.pgb.width < 1000 Then
Me.lbpg = "Protecting Desktop..."

lockDesktop Locked
End If

If Me.pgb.width = 1019 Then
Me.lbpg = "Preparing Client Desktop's..."
 installcCyberDesktop
End If

'Me.lbpg = Format(Me.pgb.Width / 1755, "0%")
Loop Until Me.pgb.width >= 1755

Me.lbpg = IIf(0 = 0, "Staring Login Window...", "Ending...")
Pause 1.5
Unload Me
killVisibleProcesses
Call Main

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y = 465 Then
clearPolicies
End
End If

End Sub

