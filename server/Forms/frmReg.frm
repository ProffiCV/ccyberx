VERSION 5.00
Begin VB.Form frmReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "frmReg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOrigi 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   300
      Width           =   5805
   End
   Begin VB.TextBox txtOrigi 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   5805
   End
   Begin VB.CommandButton cmdreg 
      Caption         =   "&Continue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   4470
      TabIndex        =   4
      Top             =   1200
      Width           =   1305
   End
   Begin VB.CommandButton cmdreg 
      Caption         =   "&Register"
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
      Height          =   345
      Index           =   0
      Left            =   3180
      TabIndex        =   3
      Top             =   1200
      Width           =   1305
   End
   Begin VB.CommandButton cmdreg 
      Caption         =   "&Generate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   1890
      TabIndex        =   5
      Top             =   2520
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Your Registration Code (Private Key)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   3195
   End
   Begin VB.Label Label1 
      Caption         =   "Public Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   2055
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdreg_Click(Index As Integer)
'Open App.Path & "\DEBUGKEY.TXT" For Output As #1
Select Case Index
    Case 0
    
        'Print #1, "Internal Key" & encodeString(Me.txtOrigi(0).Text)
        'Print #1, "your key " & Me.txtOrigi(1).Text
        'Print #1, "needed " & Mid(encodeString(getMacAddr(0)), 4, 5) & ", " & Mid(Me.txtOrigi(0).Text, 1, 5)
   
    
    
       If encodeString(Me.txtOrigi(0).Text) = Me.txtOrigi(1).Text And Me.txtOrigi(0).Text <> "" Then
        If Mid(encodeString(getMacAddr(0)), 4, 5) = Mid(Me.txtOrigi(0).Text, 1, 5) Then

            regServer
            MsgBox "Software Already Registered. Congratulations!", vbExclamation, App.EXEName
            Unload Me
            frmStart.Show vbModal
            Else
            MsgBox "Wrong public key. Correct it and try again!", vbInformation, App.EXEName

        End If
        
            Else
            MsgBox "Wrong public key. Correct it and try again!", vbInformation, App.EXEName
        End If
        
    Case 1
    Unload Me
    Case 2
'    MousePointer = 11
'    Dim tmp$
'    tmp$ = getMacAddr(0)
'    InputBox "This is your public key! Please Save it.", "Public Key", tmp$
'    MousePointer = 0
End Select

 Close #1
End Sub

Private Sub Form_Load()
Me.txtOrigi(0).Text = Mid(encodeString(getMacAddr(0)), 4, 5) & doCryptString(getMacAddr(0))

End Sub

Private Sub txtOrigi_Change(Index As Integer)
If Index = 1 Then
Debug.Print (Len(Me.txtOrigi(1).Text) >= 10)
Me.cmdreg(0).Enabled = (Len(Me.txtOrigi(1).Text) >= 10)
End If

End Sub
