Attribute VB_Name = "modMerlin"
Option Explicit

Public startupConfig As Long
Public Ready As Boolean
Public strToshow As String
Public goChange As Boolean 'mudar conta
Public changingCard As Boolean
'hora interna
Public netCharge As Double
Public site As String
Public netNow As Double
Public wprice As Long
Public oNetp As Double
Public total As Long
Public remain As Long
Public netp As Double
Public aprice As Double 'preço actual
Public fromSite As String
Private Type HORA
    se As Long
    mi As Long
    ho As Long
    tots As Long
End Type

Private Type MYTIMER
    displ As String
    tm As HORA
End Type

Public uTime As MYTIMER

'Public dta$ 'dados do servidor...

Public TPx As Long
Public TPy   As Long

Private Type SCRWH
wx As Long 'horizontal
hy As Long 'vertical
End Type
Public video As SCRWH

Public LoggedInFlag As Boolean 'indica se o utilizador esta a utiliza o computador...
Public loginReceived As Boolean 'se recebeu dados do login
Public actCodeMem As String 'codigo usado para entrar...

Public actData As String 'dados do uso servidor...
Public cardData As String 'dados card
Public actUser As String

Public toSend As String 'dados para envoar para servidor

Public Type CARDI
    'mask cards
    id As String * 3        '[000]
    code As String * 19     '[0000-0000-0000-0000]
    date As String * 6      '[000000]
    life As String * 2      '[00]
    flag As String * 1      '[X]R reached , F false T used , N not in use, invalid details
    tbal As String * 4      '[0000]
    tusd As String * 4      '[0000]
    bytes As String * 12    '[000000000000] 'total download
End Type

Public myCard As CARDI

Public Function getMerlin() As Boolean
Load frmCli
frmlogc.show
frmlogc.SetTransparency
If server$ = "" Then
  tell "Warning E#SVR..."
  Else
  tell "Who is [" & server$ & "]?"
  frmCli.mget.Enabled = True
  Ready = False
End If

getMerlin = True

End Function
Public Function releaseMerlin(Optional gEnd As Boolean = True, Optional exitwin As Boolean = False)
Dim obj, tmr
tell "Killing objects..."
goend = gEnd
For Each obj In Forms

    For Each tmr In obj
    
     If UCase(TypeName(tmr)) = "TIMER" Then
     Debug.Print obj.Name & vbTab & tmr.Name
        tmr.Enabled = False
        tmr.Interval = 0
     End If
     
    Next
Next
tell "Closing Con..."
frmCli.tCli.Close
tell "Unloading Objects..."

For Each obj In Forms
If obj.Name <> "frmlogc" Then
tell "Unloading " & obj.Name
Unload obj
End If

Next

clearPolicies
lockDesktop Unlocked
changeWallPaper True
If gEnd = True Then
tell "Ending..."
Unload frmIcon
Unload frmlogc
If Restart = True Then
 Shell App.Path & "\ccxv3updater.exe", vbHide
End If

If exitwin = False Then End
'exit without proccess Server Request (logoff restar or shutdown)
End If

End Function

Public Function MerlinDologin(ByVal cod As String) As Boolean
Dim wtimeout As Long
If goChange = False Then LoggedInFlag = False
loginReceived = False
If goChange = True Then tell "Performing auto login..."
frmCli.Talk "CODE" & cod
Do
wtimeout = wtimeout + 1
Pause 1, True
If loginReceived = True Then Exit Do
Loop Until wtimeout = 20

If wtimeout = 20 Then
    If LoggedInFlag = False Then
    tell "Loging: Timeout..."
    actCodeMem = ""
    actUser = ""
    actData = ""
    wprice = 0
    oNetp = 0
    netCharge = 0
    netNow = 0
    Pause 1
    correctMerlinState
    End If
  Else
  tell ""
End If

MerlinDologin = LoggedInFlag
End Function

Public Function MerlinDologgoff(ByVal cod As String) As Boolean
tell "Logging off..."
If LoggedInFlag = True Then
    loginReceived = False
    Dim wtimeout As Long
    
    performUserLogOff
    
    If LoggedInFlag = False Then
       correctMerlinState
       
    Else
    
    End If
    

End If

End Function

'inicia a execução...
Sub Main()
burnPolicieFile
UpdateMyDate
cCyberXV2FLG = GetSetting(App.EXEName, "startup", "task", 0)

getMerlin
UpdateMyDate
Load frmm 'vigiar processos
UpdateMyDate
lockDesktop Locked
UpdateMyDate
End Sub

'diz bom dia etc para o cliente
Public Function beGoodBoy()
Dim retstr As String
Dim hr As Long
hr = CLng(Format(Time, "hh"))
Select Case hr
    Case 0 To 11
        retstr$ = "Good morning! "
    Case 12 To 18
        retstr$ = "Good afternoom! "
    Case 19 To 23
        retstr$ = "Good night! "
End Select

beGoodBoy = retstr$
End Function

Public Function correctMerlinState()
If frmCli.tCli.State = 7 Then
    If LoggedInFlag = True Then
        'goto windows foot again
      
    Else
       installcCyberDesktop
       Unload frmTip
       Unload frmm
       If frmlogc.Visible = False Then frmlogc.Visible = True
       frmlogc.SetTransparency
       Pause 0.5
       Load frmm
       Pause 0.2
       frmm.Timer2.Enabled = True
       DisableMenus
       frmAg.cmdLogin.Enabled = True
       Ready = True
       clientOK
    End If
    
Else
    tell server & " is Down..."
    If LoggedInFlag = True Then
    frmCli.tmWin.Enabled = False
    tell "Server is Down. You have about 20 seconds to Operate. (FREE) " & vbNewLine & _
    "(Save data, Eject Removable Devices etc. Please do it!", 24
    frmAg.cmdLogoff_Click
    frmCli.mget.Enabled = True
    Else
    killKnowProcesses
    installcCyberDesktop
    lockDesktop Locked
    Unload frmCli
    Pause 0.6
    getMerlin
    End If
    
End If


End Function

Public Function merlinShowdetails()
Dim mytf$
If frmTip.Top <> 8 Then
mytf$ = " To see Details, just move the mouse pointer to the Top Screen(Y=0)..." & vbCrLf & " " & frmTip.status.Panels(1).Text
Load frmTip
frmTip.Visible = True
tell mytf$, 5
End If

End Function

Public Function prepareBarDetails()
Dim p1$, p2$, p3$
If uTime.displ = "" Then uTime.displ = "?"
strToshow = "W{" & Left(uTime.displ, 8) & "," & Format(aprice * 100, "0$00") & "} "
strToshow = strToshow & " N{" & trasnBytes(CLng(Val(netNow))) & "," & _
Format(netp * 100, "0$00") & "} "
strToshow = strToshow & "C{" & Format(total * 100, "0$00") & "," & _
Format(Val(myCard.tusd) * 100, "0$00") & "," & _
Format(remain * 100, "0$00") & "}"


p1$ = "WIN{ " & Left(uTime.displ, 8) & " , " & Format(aprice * 100, "0$00") & "}"
p1$ = IIf(Len(p1$) >= 32, p1$, p1$ & String(32 - Len(p1$), "-"))
p2$ = " NET{ " & trasnBytes(CLng(Val(netNow))) & " , " & Format(netp * 100, "0$00") & "}"
p2$ = IIf(Len(p2$) >= 32, p2$, p2$ & String(32 - Len(p2$), "-"))

p3$ = " CRD{ " & Format(total * 100, "0$00") & " , " & Format(remain * 100, "0$00") & "}"
p3$ = IIf(Len(p3$) >= 32, p3$, p3$ & String(32 - Len(p3$), "-"))

prepareBarDetails = p1$ & p2$ & p3$

End Function

Public Function PerformUserLogin(dta$)
LoggedInFlag = True
EnableMenus

'reset counters
With uTime
    .tm.tots = 0
    .tm.mi = 0
    .tm.ho = 0
    .tm.se = 0
    .displ = ""
    wprice = 0
    oNetp = 0
End With
frmCli.tmgetWinTime.Enabled = True

CopyMemNorm myCard, ByVal dta$, 51
wprice = CLng(Val(myCard.tusd))
netCharge = CLng(Val(myCard.bytes))
oNetp = CDbl(netCharge) 'dados acumulados no ultimo uso
total = CLng(Val(myCard.tbal)) 'saldo total da conta
remain = total - wprice 'saldo restante calculado aqui...
frmCli.Talk "USER" & actUser 'iniciar processo de login enviar dados para servidor.

'Debug.Print "Stop Processo Lookup"
'Debug.Print "Positioning Merlin into basket..."


reduceMerlin 24

Load frmTip
frmTip.Top = 8
topMost HWND_TOPMOST, frmTip
frmTip.chkOntop.Value = 1
frmTip.Visible = True

'frmTip.tmanim.Enabled = True
tell beGoodBoy & _
" Welcome..."
Pause 1, True
'Debug.Print "UNBlock Desktop"
lockDesktop Unlocked
'Debug.Print "Show TIPS on top..."

'Debug.Print "StartCrapuLoop"


End Function

Public Function reduceMerlin(ByVal perc As Long)
detectTaskBarDimensions
End Function

Public Function performUserLogOff()
DisableMenus
frmCli.tmgetWinTime.Enabled = False
frmCli.tmGetNet.Enabled = False

'Debug.Print "StopCrapuLoop"
frmCli.tmgetWinTime.Enabled = False
SendDetails True

If CInt(GetSetting(App.EXEName, "Remote", "SellDownload", 1)) = 1 Then
sniffer.gStopget = True
sniffer.uninstallSniffer
End If



'Debug.Print "Block Desktop"
lockDesktop Locked

killKnowProcesses GetCurrentProcessId
'Debug.Print "Run Process Lookup"
'Debug.Print "Positioning Merlin Normal"
'Debug.Print "Hide TIPS on top..."
LoggedInFlag = False
netCharge = 0
netNow = 0
netp = 0
uTime.tm.ho = 0
uTime.tm.mi = 0
uTime.tm.se = 0
uTime.tm.tots = 0
uTime.displ = "00:00:00"
Unload frmTip

reduceMerlin 0
correctMerlinState

End Function


Public Function SendDetails(Optional outm As Boolean = False)
'usedtimewindows,netcharge,
Static useds$

If Len(myCard) = 51 Then

If myCard.tusd & myCard.bytes <> useds$ Or (myCard.tbal <= CLng(myCard.tusd)) Or outm = True Then
useds$ = myCard.tusd & myCard.bytes
'Debug.Print cardData
With myCard
cardData = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes & Format(netNow, "000000000000")

actData = uTime.displ & Format(CLng(.tusd), "000000") & IIf(netp <> 0, Format(netp, "000000"), "Cybero")
Debug.Print Format(IIf(netp <> 0, Format(netp, "000000"), "Cybero"))
End With

'frmCli.Talk "SAVE" & cardData
'Pause 2
Debug.Print Len(actData & cardData & IIf(outm = True, "1", "0"))
If Len(actData & cardData & IIf(outm = True, "1", "0")) = 88 Then
frmCli.Talk "DTLS" & actData & cardData & IIf(outm = True, "1", "0")
End If

'Pause 0.8
End If

End If

End Function

