VERSION 5.00
Begin VB.Form frmm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   180
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   180
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   40
      Left            =   2430
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   570
      Top             =   210
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "frmm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oldlo As Boolean
Private loopcount As Integer
Private Sub Form_Load()
Me.Top = 40
Me.Left = 120
lockit = True
Timer2.Enabled = True
End Sub

'version 2

Private Sub Timer2_Timer()

If LoggedInFlag = True Then Exit Sub

        'Debug.Print "working rfd"
        DoEvents
        Dim wndi&, pid&, myid&, th&, wndt$, pext&
        Timer2.Enabled = False
        Static iKnew As Long
        
       ' doevents
    
          

        If LoggedInFlag = False Then
            loopcount = loopcount + 1
                If loopcount >= 5 Then
                loopcount = 0
                    myid& = GetCurrentProcessId()
                    'tell "Reading Processes Memory..."
                killKnowProcesses myid&, False
                End If
        End If


Timer2.Enabled = True
End Sub

