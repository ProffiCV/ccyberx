VERSION 5.00
Begin VB.Form frmAg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   510
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   1770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   1770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuadm 
      Caption         =   "Admin"
      Begin VB.Menu cmdSetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu cmdPolicys 
         Caption         =   "&Clients Policy"
      End
      Begin VB.Menu cmdDesktop 
         Caption         =   "&Clients Destop"
      End
      Begin VB.Menu Sa 
         Caption         =   "-"
      End
      Begin VB.Menu cmdUnload 
         Caption         =   "&Unload"
      End
      Begin VB.Menu cmdCancel 
         Caption         =   "&Cancel"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User"
      Begin VB.Menu cmdLogin 
         Caption         =   "&Login"
      End
      Begin VB.Menu cmdLogoff 
         Caption         =   "L&ogoff"
      End
      Begin VB.Menu Sp1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu cmdChangecard 
         Caption         =   "C&hange Card"
      End
      Begin VB.Menu sp3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu cmdshowBalance 
         Caption         =   "View Balance"
         Visible         =   0   'False
      End
      Begin VB.Menu cmdViewCardInfo 
         Caption         =   "&View Card Info"
         Visible         =   0   'False
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu cmdAdmin 
         Caption         =   "&Administrator"
      End
      Begin VB.Menu cmdAbout 
         Caption         =   "&About cCyberXV3"
      End
      Begin VB.Menu cmdCancel1 
         Caption         =   "&Cancel"
      End
   End
End
Attribute VB_Name = "frmAg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub showMENU()

If frmCli.tCli.State = sckConnected Then
If Ready = False Then Exit Sub
    frmMenu.show
Else
    PopupMenu Me.mnuadm, 4
End If
End Sub

Public Sub cmdAbout_Click()
showTips " " & App.EXEName & " Ver:" & App.Major & "." & _
App.Minor & "." & App.Revision & vbNewLine & _
" Writed and Compiled By Edson Martins. " & vbNewLine & _
" Client Public Computer" & vbNewLine & _
" For ADSL Powered Cybers.", 6

End Sub

Public Sub cmdAdmin_Click()
cmdUnload_Click
End Sub

Public Sub cmdChangecard_Click()
frmChg.show
frmChg.tmgetOrNot.Enabled = True
End Sub

Private Sub cmdLogin_Click()
'If frmCli.tCli.State = 7 Then
'
'    If Not bodix Is Nothing Then
'    frmlogc.show
'    bodix.MoveTo (frmlogc.Left) / TPx - bodix.width, (frmlogc.Top) / TPy - 0.02 * bodix.height, 0
'    MerlinTell "Enter your code and a nickname " & _
'    "(Optional), Then click LOGIN.", "reading"
'    frmlogc.Visible = False
'    frmlogc.show vbModal
'    End If
'
'End If

End Sub

Public Sub cmdLogoff_Click()
Dim msn&
MerlinDologgoff actCodeMem
LoggedInFlag = False
logOffMSN
changeWallPaper
killKnowProcesses GetCurrentProcessId
installcCyberDesktop
lockDesktop Locked
correctMerlinState
End Sub

Private Sub cmdSetup_Click()
Dim msg$


tell "Admin please..."


msg = InputBox("Write down the setup sentence. " & _
"If you inform an invalid sentence, nothing will happen.", App.EXEName, _
"", frmlogc.Left - 280, frmlogc.Top + frmlogc.height + 20)

If msg = "letconfig23.bodix." Then
frmlogc.Visible = False
frmCfg.Visible = True

Else
tell "Wrong sentence..."
End If

End Sub

Private Sub cmdshowBalance_Click()
merlinShowdetails
End Sub

Public Sub cmdUnload_Click()
If LoggedInFlag = True Then GoTo fim
Dim tms As String
tell "Config key..."


tms = InputBox("Write the admin sentence to unload the Client", _
"Admin Sentences", "", frmlogc.Left - 280, frmlogc.Top + frmlogc.height + 20)


If tms = "theownedson19971981200523241" Then
lockDesktop Unlocked
ElseIf tms = "theownedson19971981200523240" Then
lockDesktop Locked
ElseIf tms = "default_policie_file" Then
   Call burnPolicieFile
ElseIf tms = "letsgodown.bodix." Then
releaseMerlin
Else
    If tms <> "" Then
    frmCli.Talk "FUCKClient are trying to Close cCyberXV2." & vbNewLine & _
    "Used Sentence: " & tms
    End If

correctMerlinState

End If

fim:
End Sub


Private Sub Form_Load()
Ready = True
End Sub
