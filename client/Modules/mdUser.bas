Attribute VB_Name = "mdUser"
Option Explicit
Public userLogged As Boolean
Public changingCode As Boolean
'inicia a sessão do utilizador....
Public Function Loggin(ByVal code As String) As Boolean
If LoggedInFlag = False Then
    tell "Validating..."
    Pause 0.8, True
    Loggin = MerlinDologin(actCodeMem)
    If Loggin = True Then Unload frmm
    If Loggin = False Then tell "Try Again..."
    frmlogc.cmdUser(0).Enabled = True
    'frmCli.tCli.State
Else
    tell "Welcome..."
End If


End Function

'termina a sessão do utilizador...
Public Function Logoff(ByVal code As String) As Boolean
If goChange = False Then
tell "Logging off..."
MerlinDologgoff code
Else
SendDetails True
netCharge = 0
uTime.tm.ho = 0
uTime.tm.mi = 0
uTime.tm.se = 0
uTime.tm.tots = 0
netp = 0
aprice = 0
wprice = 0
oNetp = 0
End If

End Function

'provoca pausas na execuÇão do codigo sem bloquear
Public Function Pause(ByVal ms As Double, Optional helpwin As Boolean = False)
Dim init As Double
init = Timer

Do
If helpwin = True Then DoEvents
Debug.Print "Pausing " & ms
Loop Until Timer - init >= (ms)
End Function

Public Function isCodeOK(ByVal code As String)
If frmCli.tCli.State = 7 Then
changingCard = True
changingCode = True
    frmCli.tCli.SendData "CHEK" & code
End If

End Function



