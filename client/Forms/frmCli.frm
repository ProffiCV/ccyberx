VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShowldNotSeeMe"
   ClientHeight    =   705
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5175
   Icon            =   "frmCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmGetNet 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3270
      Top             =   120
   End
   Begin VB.Timer tmWin 
      Interval        =   1000
      Left            =   2790
      Top             =   120
   End
   Begin VB.Timer tmgetWinTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   120
   End
   Begin VB.Timer tmconstate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1740
      Top             =   120
   End
   Begin VB.Timer tmreco 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer mget 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   660
      Top             =   120
   End
   Begin MSWinsockLib.Winsock tCli 
      Left            =   120
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmCtrl 
      Interval        =   400
      Left            =   4320
      Top             =   120
   End
End
Attribute VB_Name = "frmCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdGetPkt_Click()

End Sub

Private Sub Form_Load()
Load frmCfg

server = frmCfg.tHost.Text
If frmCfg.lstport.ListIndex <> -1 Then
port = IIf(frmCfg.lstport.ListIndex <> -1, frmCfg.lstport.ItemData(frmCfg.lstport.ListIndex), 0)
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Talk "KILL" & vbCrLf
Me.tCli.Close

Unload frmCfg
End Sub

Private Sub mget_Timer()
If frmChg.Visible = True Then Exit Sub
Me.mget.Enabled = False
If Me.tCli.State = 7 Then
tmconstate.Enabled = True
Exit Sub
End If

DoEvents
Me.tCli.Close
Me.tCli.RemotePort = port&
Me.tCli.RemoteHost = server$
Me.tCli.Connect

tmreco.Enabled = True 'check if connected
tell "Connecting to " & server & "... 0x" & Hex(port&)
'Me.mget.Enabled = True
End Sub

Private Sub tCli_DataArrival(ByVal bytesTotal As Long)
Dim dta$
If Me.tCli.State = 7 Then
    Me.tCli.GetData dta$
    
        InterPretData dta$
        'showTips dta$, 3
    
End If

End Sub

Private Sub tmconstate_Timer()
DoEvents

    Select Case Me.tCli.State
        Case 1, 2, 3, 4, 5, 6
        frmlogc.Shape2.BorderColor = frmlogc.Shape3.BorderColor = vbRed
        Case 7
                If frmlogc.txtData(0).Enabled = False Then
                     frmlogc.Shape3.BorderColor = vbYellow
                     frmlogc.Shape2.BorderColor = frmlogc.Shape3.BorderColor
                    frmlogc.txtData(0).Enabled = True
                    frmlogc.txtData(1).Enabled = True
                End If
                
            If LoggedInFlag = False Then
                lockDesktop Locked
            Else
                lockDesktop Unlocked
            End If
        Case 8, 9, 0
        frmlogc.Shape2.BorderColor = frmlogc.Shape3.BorderColor = vbRed
        tmconstate.Enabled = False
        frmlogc.txtData(0).Enabled = False
        frmlogc.txtData(1).Enabled = False
        correctMerlinState
        
    End Select

End Sub

Private Sub tmCtrl_Timer()

 If GetAsyncKeyState(VK_LCONTROL) <> 0 And GetAsyncKeyState(VK_LMENU) <> 0 Then
   If LoggedInFlag = True Then
        frmAg.cmdLogoff_Click
    Else
     showTips " Warning!!! you are not Signed in!", 3
   End If
   
 End If
 
End Sub

Private Sub tmGetNet_Timer()
    Call sniffer.receivePacket
End Sub

Private Sub tmgetWinTime_Timer()

Me.tmgetWinTime.Enabled = False
If LoggedInFlag = True Then
   With uTime
   
    'update data here
    With myCard
       .bytes = Format(netCharge, "000000000000")
       .flag = "T"
       If CLng(Val(.bytes)) > prc.offset * MB Then
            .flag = "R" 'marcar que já recebeu os dados que oferecemos ex 10 MB
        Else
        oNetp = CDbl(netCharge) 'lembrar o ultimo acumulado (downloads)
        
       End If
       
       If .flag = "R" Then 'atingido? vender o resto ultrapassado...
          netp = Format(Round(((netCharge - oNetp) / MB) * prc.netp), "0")
       Else
       
       
       End If
       'calcular o preço de uso windows (computador) total usado...
       .tusd = Format(Round(prc.windows * (uTime.tm.tots) / 3600), "0000")
'       Debug.Print .tusd
    End With
    '
   
    aprice = CDbl(Format(Round((prc.windows * .tm.tots) / 3600), "0000"))
    .displ = Format(.tm.ho, "00:") & Format(.tm.mi, "00:") & Format(.tm.se, "00") & Format(aprice, "0000")

    'include net price too
        
    myCard.tusd = Format(wprice + aprice + netp, "0000") 'preço de venda tudo
    'downloads + windows
    
    'evitar valores negativos
    If CDbl(myCard.tusd) < total Then 'sem saldo? auto logoff cliente...
        
        'ver se queria mudar de conta sem sair...
        If changingCard = False Then 'evitar perturbação nos dados da rede me server
            
            If (.tm.tots Mod 14) = 0 Then 'enviar dados a cada 14? segundos...
                merlinShowdetails 'mostrar detalhes de uso para o cliente logado...
                
            End If
        End If
        
     Else
       myCard.tusd = Format(total, "0000") 'calcular total com 4 digitos , sem casas decimais..,
    End If
  
    'alertar utilizador se pretende trocar cartão
   
    remain = total - CLng(Val(myCard.tusd)) 'fazer com que remain seja 0
     
    Debug.Print remain, uTime.tm.tots, CStr((uTime.tm.tots / 30))
        If remain <= 20 And InStr(CStr((uTime.tm.tots / 30)), ",") = 0 Then
            If goChange = False Then 'alertar utilizador para guardar os seus dados...
                tell "Balance " & Format(remain * 100, "0$00") & ", Please, SAVE your DATA Or change to a new CARD. All Documents will be closed if your Balance reach 0$00"
            Else 'mudar de conta caso solicitado...
                tell "Balance " & Format(remain * 100, "0$00") & ", Ready to change to new CARD." & vbCrLf & "7891-9134-4230-3000"
            End If
        End If
        
    
    
    
    If remain <= 0 Then
        'frmCli.Talk "SAVE" & cardData
        frmTip.chkOntop.ForeColor = vbBlack
        If goChange = True Then
            'changeToNewCardNow
            tell "Changing to new Card. Please Wait..."
            Pause 0.8
            Logoff myCard.code 'auto logoff sem sair... mudar conta
            Pause 2
            
            MerlinDologin actCodeMem 'iniciar nova conta...
        Else
            'closeAllUserProcess
            tell "Performing AutoLogOff. Out of Box..."
            
            performUserLogOff
            Pause 0.8
            killKnowProcesses GetCurrentProcessId
        End If
    End If
   End With
End If

Me.tmgetWinTime.Enabled = True
End Sub

Private Sub tmreco_Timer()

'Debug.Print "Checking connection..."
Me.tmreco.Enabled = False
If Me.tCli.State = 7 Then
    Me.mget.Enabled = False
'    Debug.Print "Connected"
    
    Me.tmconstate.Enabled = True
    tmreco.Enabled = False
    'tell "Connection was established with " & server
    Ready = False
    logOffMSN
    Pause 0.8
    Exit Sub


Else
Me.tmconstate.Enabled = False
Me.mget.Enabled = True
End If
End Sub

Private Function InterPretData(dta As String)

dta$ = Replace(dta, vbCrLf, "")

    Select Case UCase$(Left(dta, 4))
        Case Is = "GIVE"
            Talk "NEWM" & Environ$("COMPUTERNAME")
'            Pause 0.8
'            Talk "CFG?" & vbCrLf
        Case Is = "HEIS"
            If getConfig(Right(dta, Len(dta) - 4)) = False Then
                Talk "CFG?" & vbCrLf
             Else
             correctMerlinState
            End If
        Case Is = "WRGC"
        
        If changingCode = True Then
         tell "Wrong Code..."
         changingCard = False
            changingCode = False
        Exit Function
        End If
        
        
        loginReceived = True
        tell "Wrong Code..."
            frmlogc.txtData(0).Text = ""
             Pause 2
           
            correctMerlinState
        Case Is = "NOCR"
        If changingCode = True Then
        tell "No Balance..."
          frmlogc.txtData(0).Text = ""
        changingCard = False
        changingCode = False
        Exit Function
        End If
        
            loginReceived = True
            tell "No Balance..."
            Pause 2
            correctMerlinState
        Case Is = "MSGR"
            If changingCode = True Then
            tell "Data Error..."
              frmlogc.txtData(0).Text = ""
            changingCard = False
            changingCode = False
            Exit Function
            End If
        
        loginReceived = True
            tell "Data Error..."
            Pause 2
            correctMerlinState
        Case Is = "FIXE"
            If changingCode = True Then
            changingCard = False
            changingCode = False
            goChange = True
            tell "Changed..."
            frmTip.chkOntop.ForeColor = RGB(0, 140, 0)
            Exit Function
            End If
        
            loginReceived = True
            cardData = Right(dta$, Len(dta$) - 4)
            If Len(cardData) <> 51 Then
                 tell "Data Error" & Len(dta)
                 correctMerlinState
                Else
                              PerformUserLogin cardData
                
            End If
        Case "BUSY"
        tell "Code in use..."
        Pause 1.8
        correctMerlinState
        Case "LOK0" 'matar
        lockit = True
        If debugMode = False Then
        lockDesktop Locked
        lockInputs True
        End If
        
       
        tell "Terminal Locked..."
    
        Case "BCOD"
        Dim gmsg$
        gmsg$ = Right(dta$, Len(dta) - 4)
        tell gmsg$, 3
        Case "LOK1" 'desmatar
        lockit = False
        'resent...
        
        
        If LoggedInFlag = False Then
           lockDesktop Unlocked
        lockInputs False
            clientOK
          
        End If
        
        Case "SHUT"
        
            Dim flg&
            flg = CLng(Val("0" & Mid(dta$, 5, 4)))
            
            If LoggedInFlag = True Then
            frmAg.cmdLogoff_Click
            End If
            Pause 2
            releaseMerlin True, True
           Pause 1
            ShutdownSystem flg
          
        Case "CODE"
        If CLng("0" & Right(dta$, 1)) = 0 Then
            If LoggedInFlag = True Then
            MerlinDologgoff actCodeMem
            
            logOffMSN
            Else
            releaseMerlin
            End If
            '3714-3547-9924-5800
        End If
        
    End Select
    
End Function

Public Function Talk(ByRef what$)

    If Me.tCli.State = 7 Then
    
        Me.tCli.SendData what$ & vbCrLf
        lastDataSent = what$
    End If

End Function

'local timer control for computer usage...
Private Sub tmWin_Timer()
DoEvents
If LoggedInFlag = True Then
'SendDetails
With uTime
.tm.tots = .tm.tots + 1
   
   .tm.se = .tm.se + 1 'contar segundos
   If .tm.se = 60 Then 'incrementar minuto
   .tm.se = 0
    .tm.mi = .tm.mi + 1
        If .tm.mi = 60 Then 'incrementar hora
        .tm.mi = 0
         .tm.ho = .tm.ho + 1
        End If
   End If
   
 If (.tm.se Mod 10) = 0 Then
   If noti.steelThere() = 0 Then
    Unload frmIcon
    Load frmIcon
   End If
   
 End If
 
 If (.tm.se Mod 14) = 0 Then
 SendDetails
 End If
 
 End With
 End If
End Sub

