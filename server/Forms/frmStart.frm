VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStart.frx":0000
   ScaleHeight     =   1095
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   6930
      Left            =   5220
      Picture         =   "frmStart.frx":136CE
      Top             =   -3480
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   60
      Left            =   60
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label lbpg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   30
   End
   Begin VB.Image pgb 
      Height          =   105
      Left            =   60
      Picture         =   "frmStart.frx":20130
      Top             =   960
      Width           =   1755
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ct&
Private Sub Form_Load()


If Command$ = "getmakeyedson2390" Then Clipboard.SetText getMacAddr
 topMost HWND_TOPMOST, Me
 
 frmsv2.Enabled = False
Me.Icon = frmsv2.Icon
Me.Caption = IIf(operation = 0, "Loading...", "Unloading...")
Me.lbpg.Caption = Me.Caption
HideCaption Me.hwnd

Me.Height = 1155
Me.pgb.Width = 0
Do
DoEvents
Me.Visible = True
Loop Until Me.Visible = True

Do
DoEvents
'Pause 0.0002
Me.pgb.Height = 60
Me.pgb.Width = Me.pgb.Width + 4

If operation = 1 Then
If Me.pgb.Width > 1400 Then
    AnimateWin frmsv2.hwnd, 600, AW_BLEND + AW_HIDE
    
End If
Else
End If

Debug.Print Me.pgb.Width
If Me.pgb.Width = 1003 Then
   
   If operation = 0 Then
   If DetectOtherInstances("ed") = True Then
   End
   Else
   buildLicense
   End If
   End If
   
   
  
End If
Debug.Print Me.pgb.Width
'Me.lbpg = Format(Me.pgb.Width / 1755, "0%")
Loop Until Me.pgb.Width >= 1755

If operation = 0 Then Unload frmsv2
If operation = 0 Then frmsv2.myLoad
If operation = 0 Then frmsv2.Enabled = False
Me.lbpg = IIf(operation = 0, "Starting...", "Ending...")
Pause 0.6
If operation = 0 Then frmsv2.Enabled = True
If operation = 0 Then Unload Me

If operation = 1 Then
Unload frmsv2
End 'stop here
End If

If operation = 0 Then operation = 1  'preparing to end
End Sub

Private Sub Image1_Click()
ct = ct + 1
If ct >= 12 Then
    'regServer
    Me.lbpg = "Bad Boy!"
End If



End Sub
