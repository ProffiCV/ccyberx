Attribute VB_Name = "modPacket"
Option Explicit
Public eSock As Long 'socket....
Public Type IPHOST
dn As String * 255
ip As String * 14
End Type

Public myMask$

Dim sFrom As SockAddr
Dim sStart As SockAddr
Dim ret As Long
Public Function installSniffer(ByVal ghwnd As Long, ByVal rcTo&) As Long
Dim Installed As Long, hostLng&, maskLng&, Mask&
Installed = -1 'someting is wrong...

Dim ret&, bin As Long, bout As Long, pkl As Long
Dim host As String
Dim lpwsa As WsaData
Dim sckadin As SockAddr
Dim lsa As Long
'iniciando o serviço

lpwsa.wVersion = &H202
ret& = WSAStartup(&H202, lpwsa)

If ret = 0 Then '''''''''''''''''''''''''''''''''''LEVEL 1
host = Environ("ComputerName")
host = Trim$(hostAdress(GetAddressLong(host)).ip)
hostLng = GetAddressLong(host)

myMask = Left(host, InStrRev(host$, ".") - 1)
    eSock = socket(AF_INET, SOCK_RAW, IPPROTO_IP)
    If eSock <> -1 Then '''''''''''''''''''''''''''LEVEL 2
        
        sckadin.sin_family = AF_INET
        sckadin.sin_port = htons(0&)
    
        lsa = LenB(sckadin)
        sckadin.sin_addr = hostLng&

        ret& = setsockopt(eSock, SOL_SOCKET, SO_RCVTIMEO, rcTo, 4)
        
        If ret& = 0 Then '''''''''''''''''''''''''''LEVEL 3
            ret& = bind(eSock, sckadin, lsa)
                
            If ret& = 0 Then '''''''''''''''''''''''LEVEL 4
            bin = 1
                ret& = WSAIoctl(eSock, SIO_RCVALL, bin, Len(bin), _
                bout, Len(bout), pkl, ByVal 0&, ByVal 0&)
                
                 If ret = 0 Then '''''''''''''''''''''LEVEL 5
                    ret = WSAAsyncSelect(eSock, ghwnd, WM_LBUTTONDOWN, ByVal FD_READ)
                        If ret = 0 Then '''''''''''''''''''''LEVEL 6
                            Installed = 0 'we do it
                        End If '''''''''''''''LEVEL 6
                 End If '''''''''''''''LEVEL 5
            End If '''''''''''''''LEVEL 4
        End If '''''''''''''''LEVEL 3
    End If '''''''''''''''LEVEL 2
End If '''''''''''''''LEVEL 1
            

installSniffer = Installed 'what happened..
End Function

Public Function uninstallSniffer() As Long
Debug.Print "Closing Socket " & eSock; closesocket(eSock&)
uninstallSniffer = WSACleanup
End Function







Public Function GetAddressLong(ByVal strHostname As String) As Long
  On Error Resume Next
    
  Dim lpHostent   As Long
  Dim udtHostent  As HostEnt
  Dim AddrList    As Long
  Dim retIP       As Long
    
    retIP = inet_addr(strHostname)
    
    If retIP = -1 Then
        lpHostent = gethostbyname(strHostname)
        
        If lpHostent <> 0 Then
            CopyMemSock udtHostent, ByVal lpHostent, LenB(udtHostent)
            CopyMemSock AddrList, ByVal udtHostent.h_addr_list, 4
            CopyMemSock retIP, ByVal AddrList, udtHostent.h_length
        Else
            retIP = -1
        End If
    End If
    
    GetAddressLong = retIP

End Function
    
'''''''''''''''''''''''''''''


Public Function hostAdress(ByVal ptrStr As Long) As IPHOST
Dim hp(3) As Byte
CopyMemSock hp(0), ByVal VarPtr(ptrStr), 4

hostAdress.dn = getDn(ptrStr)
hostAdress.ip = hp(0) & "." & hp(1) & "." & hp(2) & "." & hp(3)
End Function

Private Function getDn(ByVal ptrAdr As Long) As String
Dim phEntr As Long, hEnt As HostEnt, hName As String
hName = Space$(256)
phEntr = gethostbyaddr(ptrAdr, 4, AF_INET)

If phEntr <> 0 Then
CopyMemSock hEnt, ByVal phEntr, LenB(hEnt)

CopyMemSock ByVal hName, ByVal hEnt.h_name, 256
'Debug.Print hName
getDn = Left(hName, InStr(1, hName, Chr(0)) - 1)
Else
getDn = "Unknown.."
End If

End Function


Public Function peekRawData()
DoEvents
     ReDim buff(1023) As Byte
     sFrom = sStart
        ret& = recvfrom(eSock, ByVal VarPtr(buff(0)), 1024, 0&, ByVal VarPtr(sFrom), LenB(sFrom))
        If ret <> -1 Then
        netNow = ret&
          Dim tho$
          tho$ = Trim$(hostAdress(sFrom.sin_addr).ip)
          If InStr(1, tho$, myMask) = 0 Then
          netCharge = netCharge + ret&
          fromSite = "LPck From:" & tho$
          'site = Trim$(hostAdress(sFrom.sin_addr).dn)
          'strToshow = "Total " & accNetByte
          peekRawData
          Else
         
          End If
          
        End If

End Function


