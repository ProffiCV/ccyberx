VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmChg 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2715
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChgCrd.frx":0000
   ScaleHeight     =   2715
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   0
      TabIndex        =   2
      Top             =   1710
      Width           =   4785
      Begin VB.Label cmdUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " CONTINUE"
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
         Index           =   0
         Left            =   1260
         TabIndex        =   4
         Top             =   210
         Width           =   840
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
         Left            =   2850
         TabIndex        =   3
         Top             =   210
         Width           =   705
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   990
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   2745
      End
   End
   Begin VB.Timer tmgetOrNot 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4140
      Top             =   390
   End
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "####-####-####-####"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2070
         SubFormatType   =   0
      EndProperty
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
      Index           =   0
      Left            =   660
      MaxLength       =   19
      TabIndex        =   1
      Top             =   1320
      Width           =   4035
   End
   Begin MSComctlLib.StatusBar status 
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   2460
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   8361
         EndProperty
      EndProperty
   End
   Begin VB.Label lba 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Enter New Code Here"
      Enabled         =   0   'False
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
      Height          =   210
      Index           =   1
      Left            =   2130
      TabIndex        =   5
      Top             =   1110
      Width           =   1620
   End
End
Attribute VB_Name = "frmChg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private glIndex As Long
Private cont As Long

Public Sub cmdUser_Click(Index As Integer)
Select Case Index
Case 0
Me.tmgetOrNot.Enabled = False
Me.txtData(0).Enabled = False
Me.cmdUser(0).Enabled = False
actCodeMem = Me.txtData(0).Text
Me.txtData(0).Text = ""


isCodeOK (actCodeMem)
Unload Me
Case 1
Case 2
Me.cmdUser(0).Enabled = False
Me.cmdUser(2).Enabled = False
Unload Me

End Select

End Sub

Private Sub cmdUser_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.cmdUser(Index).ForeColor <> vbRed Then Me.cmdUser(Index).ForeColor = vbRed
    
 Select Case Index
    Case 0
        'tell "New Code please..."
    Case 1
    Case 2
        'tell "Click Cancel..."
 
 End Select
 
End Sub

Private Sub Form_Load()

Me.status.Font.Name = "Arial"
Me.status.Font.Bold = True
If frmlogc.Visible = False Then
tell "New Code..."
Me.tmgetOrNot.Enabled = True
Else

End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'limpar realce
    If Me.cmdUser(0).ForeColor <> &HC00000 Then Me.cmdUser(0).ForeColor = &HC00000
    If Me.cmdUser(2).ForeColor <> &HC00000 Then Me.cmdUser(2).ForeColor = &HC00000

End Sub

Private Sub tmgetOrNot_Timer()
If Me.Visible = True Then
 DoEvents
 cont = cont + 1

 If cont > 20 Then
 cont = 0
 Me.tmgetOrNot.Enabled = False
 Unload Me
 tell "Timeout..."
 Pause 1
 End If
End If

End Sub


Private Sub txtData_Change(Index As Integer)
cont = 0
Select Case Index
    Case 0
    If Me.txtData(0).Text = myCard.code Then
    Me.txtData(0).Text = ""
    Exit Sub
    End If
    

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

If KeyCode <> 8 Then
    Select Case Index
    Case 0
        'tell "Keep writiing the code, without (-)"
        With Me.txtData(Index)
        Me.cmdUser(0).Enabled = False
            Select Case Len(.Text)
                Case 4, 9, 14
                    If Right(.Text, 1) <> "-" Then .Text = .Text & "-"
                    .SelStart = Len(.Text)
                Case 19
                    Me.cmdUser(0).Enabled = True
                    tell "Continue to Go!"
            End Select
        End With
        Case 1
        If Len(Me.txtData(1).Text) = 0 Then
            Me.txtData(1).Text = "Anonymous"
            Me.txtData(1).SelStart = 0
            Me.txtData(1).SelLength = Len(Me.txtData(1).Text)
        End If
        
    End Select
ElseIf KeyCode = 13 Then
If Me.txtData(0).Text <> "" And Me.txtData(0).Text <> myCard.code Then
cmdUser_Click 0
End If

End If

End Sub

Private Sub txtData_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        'tell "Enter your code, without minus signal."
    Case 1
        'tell "Our Staff may use it to contact you, during your Session!"
End Select

End Sub

