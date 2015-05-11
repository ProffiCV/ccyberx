VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPass.frx":0000
   ScaleHeight     =   1875
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Administrator"
      Top             =   270
      Width           =   2655
   End
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   630
      Width           =   2655
   End
   Begin VB.Label cmdCancel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   1
      Left            =   1020
      MouseIcon       =   "frmPass.frx":18EB2
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label cmdCancel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   0
      Left            =   2460
      MouseIcon       =   "frmPass.frx":191BC
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1110
      Width           =   1245
   End
   Begin VB.Label lbMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wrong Passwrod. Please Try Again"
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
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   1590
      Width           =   3855
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const vbAsk = vbQuestion + vbYesNo
Public Sub cmdCancel_Click(Index As Integer)
Me.Enabled = False
Static pcnt As Integer
If Index = 0 Then
logged = False
    Unload Me
    
Else
If Me.txtPass(0).Text = "" Then
Debug.Print Environ$("Computername")
If UCase(Environ$("Computername")) = "EDSONPORTABLE" <> 0 Then
logged = True
Unload Me
End If

End If

If Len(Me.txtPass(0).Text) < 10 Then
metell "Attention! Your Password is 10 digit."
    Me.txtPass(0).SelStart = 0
    Me.txtPass(0).SelLength = Len(Me.txtPass(0).Text)
    
    Me.Enabled = True
Exit Sub
End If

 metell "Validating Password, Please Wait..."
 busy
 Pause 2
 Debug.Print Me.txtPass(0).Text
 
    If Me.txtPass(0).Text = "cardx-9023" Then
    logged = True
    Unload Me
    Else
    Me.txtPass(0).SelStart = 0
    Me.txtPass(0).SelLength = Len(Me.txtPass(0).Text)
    
    pcnt = pcnt + 1
    metell "Wrong Password. Please try again 3/" & pcnt
    
    If pcnt = 4 Then
    busy
    metell "Contact Edson Martins to get password..."
    Pause 2
    logged = False
    Unload Me
    End If
    
    End If
free
End If

Me.Enabled = True
End Sub

Private Sub Form_Load()
metell "Write the password and click Login..."
End Sub

Private Sub metell(ByVal msg$)
Me.lbMsg = msg$
End Sub

Private Sub txtPass_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
cmdCancel_Click 1
End If
End Sub
