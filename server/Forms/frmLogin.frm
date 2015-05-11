VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3585
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUserName 
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
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LoginSucceeded As Boolean
Private CreatePass As Boolean
Private tmpPass As String
Private SkypePass As Boolean
Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me
End Sub

Public Sub cmdOK_Click()
    'check for correct password
    If CreatePass = True Then
        If tmpPass = "" Then
        tmpPass = Me.txtPassword.Text
        MsgBox "Type the password again to confirm", vbExclamation, App.EXEName
        
        Me.txtPassword.Text = ""
        Me.Show , frmsv2
        Me.txtPassword.SetFocus
        Exit Sub
        Else
         If Me.txtPassword.Text = tmpPass Then
            'confirmou palavra passe
            SaveSetting App.EXEName, "Data", "pass", tmpPass
            tmpPass = ""
            MsgBox "You has just created the password. " & NL & _
            "Click Configure again and use the new created password.", vbExclamation, App.EXEName
            Unload Me
         Else
            'erro nao coinscide
            MsgBox "The confirmation password should be" & NL & _
            "same to the another typed previously.", vbExclamation, App.EXEName
            Me.Show , frmsv2
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
         End If
         
        
        End If
        
    Else
    
    If SkypePass = True Then
    'show configuration
    Unload Me
    Else
    'confirm password
        If txtPassword = getDefaultPass Or txtPassword = getUserPass(True) Then
            'place code to here to pass the
            'success to the calling sub
            'setting a global var is the easiest
            LoginSucceeded = True
            Unload Me
            frmCfg.Show , frmsv2
            
        Else
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
        End If
    End If
    
    End If
    
End Sub

Public Function getUserPass(Optional check As Boolean = False) As String
'read saved password
    getUserPass = GetSetting(App.EXEName, "Data", "pass", "")
    'defalut password found
    If txtPassword = getDefaultPass Then Exit Function
    
    If getUserPass = "" Then
    getUserPass = getDefaultPass
    If MsgBox("Your password is Blank. " & vbNewLine & _
    "Do you want to create it now?", vbQuestion + vbYesNo, App.EXEName) = vbYes Then
        txtUserName = "Admin"
        txtPassword = ""
        MsgBox "Write a password and click OK to continue", , App.EXEName
        CreatePass = True
        Me.Show , frmsv2
        Me.txtPassword.SetFocus
     Else
            Unload Me
            frmCfg.Show
            
    End If
    Else
    If check = True Then
    Exit Function
    End If
    
        CreatePass = False
        SkypePass = False
        Me.Show , frmsv2
        Me.txtUserName.Text = "Admin"
        Me.txtPassword.SetFocus
    End If
    
End Function

Private Function getDefaultPass()
getDefaultPass = "theownedson19971981"
End Function

Private Sub Form_Load()
SkypePass = False

End Sub

Private Sub txtPassword_Change()
Me.cmdOK.Enabled = Len(Me.txtPassword.Text) > 5
End Sub

Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdOK_Click
End Sub
