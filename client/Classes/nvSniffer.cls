VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sniffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim so As Long

Public localIP As String
Private myNetID As String
Public bNow As Long 'bytes received per packet
Public gTotal As Double 'total bytes received...
Public oldb As Double
Public gStopget As Boolean
Const BL = 1500
Private Const WM_MBUTTONUP = &H208
Private Const FD_READ As Long = &H1
Private Const WM_LBUTTONDOWN As Long = &H201 ' wsmg parameter
Private Const FD_WRITE           As Long = &H2
Private Const WM_LBUTTONUP = &H202

Dim buffer(BL) As Byte

Public toshowStr As String

Private Const AF_INET = 2
Private Const SOCK_RAW = 3
Private Const SIO_RCVALL = &H98000001
Private Const SO_RCVTIMEO = &H1006
Private Const SO_RCVBUF = &H1002
Private Const SOL_SOCKET = &HFFFF&
Private Const SO_NONBLOCK = &H8004667E
Private tProt  As String

'-----------------------------------------------------------------------------
Private Type HostEnt
    h_name                      As Long
    h_aliases                   As Long
    h_addrtype                  As Integer
    h_length                    As Integer
    h_addr_list                 As Long
End Type
'-----------------------------------------------------------------------------
Private Type SOCKADDR
  sin_family As Integer    'short
  sin_port   As Integer    'u_short
  sin_addr   As Long       'struct IN_ADDR
  sin_zero   As String * 7 'char[8]
End Type
'-----------------------------------------------------------------------------
Private Type WSADATA 'Requires Windows Sockets 2.0
  wVersion       As Integer      'WORD                        // Version of the Windows Sockets specification that the Ws2_32.dll expects the caller to use.
  wHighVersion   As Integer      'WORD                        // Highest version of the Windows Sockets specification that this .dll can support (also encoded as above). Normally this is the same as wVersion.
  szDescription  As String * 256 'char [WSADESCRIPTION_LEN+1] // Null-terminated ASCII string into which the Ws2_32.dll copies a description of the Windows Sockets implementation. The text (up to 256 characters in length) can contain any characters except control and formatting characters: the most likely use that an application can put this to is to display it (possibly truncated) in a status message.
  szSystemStatus As String * 128 'char [WSASYS_STATUS_LEN+1]  // Null-terminated ASCII string into which the WSs2_32.dll copies relevant status or configuration information. The Ws2_32.dll should use this parameter only if the information might be useful to the user or support staff: it should not be considered as an extension of the szDescription parameter.
  iMaxSockets    As Integer      'unsigned short              // Retained for backward compatibility, but should be ignored for Windows Sockets version 2 and later, as no single value can be appropriate for all underlying service providers.
  iMaxUdpDg      As Integer      'unsigned short              // Ignored for Windows Sockets version 2 and onward. iMaxUdpDg is retained for compatibility with Windows Sockets specification 1.1, but should not be used when developing new applications. For the actual maximum message size specific to a particular Windows Sockets service provider and socket type, applications should use getsockopt to retrieve the value of option SO_MAX_MSG_SIZE after a socket has been created.
  lpVendorInfo   As Long         'char far*                   // Ignored for Windows Sockets version 2 and onward. It is retained for compatibility with Windows Sockets specification 1.1. Applications needing to access vendor-specific configuration information should use getsockopt to retrieve the value of option PVD_CONFIG. The definition of this value (if utilized) is beyond the scope of this specification.
End Type
'-----------------------------------------------------------------------------------
Private Type IPHeader
 m_ver           As Byte             '1IP version number
 m_len           As Byte             '2header length in 32bit words (4 bits each)
 m_tos           As Byte             '3Type Of Service ID (1 octet)
 m_totallength   As Integer          '5Size of Datagram (header + data) in octets
 m_id            As Long             '9IP-ID (16 bits)
 m_offset        As Integer          '11fragmentation flags (3bit) and fragmet offset (13 bits)
 m_ttl           As Byte             '12datagram Time To Live (in network hops)
 m_protocol      As Byte             '13Transport protocol type (byte)
 m_checksum      As Long             'Header Checksum (16 bits)
 m_srcaddr       As String             'Source IP Address (32 bits)
 m_destaddr      As String             'Destination IP Address (32 bits)
 m_srcip         As String
 m_destip        As String
End Type

Const TCP = 6&
Const UDP = 17&
Const ICMP = 1&


''''''''''''''APIS
Private Declare Function WSAIoctl Lib "ws2_32.dll" (ByVal hSocket As Long, ByVal dwIoControlCode As Long, In_Buffer As Any, ByVal In_BufferLen As Long, Out_Buffer As Any, ByVal Out_BufferLen As Long, lpcbBytesReturned As Long, lpOverlapped As Long, lpCompletionRoutine As Long) As Long
Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare Function socket Lib "ws2_32.dll" (ByVal AddressFamily As Long, ByVal SocketType As Long, ByVal Protocol As Long) As Long
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As Any) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal IpAddress As String) As Long
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal InAddr As Long) As Long
Private Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Private Declare Function bind Lib "ws2_32.dll" (ByVal hSocket As Long, ByRef Name As SOCKADDR, ByRef NameLen As Long) As Long
Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Private Declare Function setsockopt Lib "ws2_32.dll" (ByVal hSocket As Long, ByVal Level As Long, ByVal OptionName As Long, ByRef OptionValue As Any, ByVal OptionLength As Long) As Long
Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal hSocket As Long, ByVal OptionName As Long, ByRef OptionValue As Any) As Long
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long, ByVal char As Long)
Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal hSocket As Long, _
ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

'------------------
Private Declare Sub CopyMemSock Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'---------------------------------------------------
Private iph As IPHeader
Private btotal#, brec&


Public Function installSniffer(ByVal hwnd As Long) As Long
Dim ret&, bret&, tp&
Dim wsa As WSADATA, lca As SOCKADDR

ret = WSAStartup(&H202, wsa)


Call getIPmyNetID
If hwnd = 0 Then
Call uninstallSniffer
Exit Function
End If

Debug.Print localIP
If localIP = "255.255.255.255" Then GoTo fim


Dim bin As Long
bin = 1
If ret = 0 Then
    lca.sin_addr = GetAddressLong(Environ("COMPUTERNAME"))
    lca.sin_family = AF_INET
    lca.sin_port = 0&
    so = socket(AF_INET, SOCK_RAW, 0)
    ret = setsockopt(so, SOL_SOCKET, SO_RCVBUF, 1&, 4)
    ret = ioctlsocket(so, SO_NONBLOCK, True)
    ret = setsockopt(so, SOL_SOCKET, SO_RCVTIMEO, 0&, 4)
    
    If so <> 0 And ret = 0 Then
        ret = bind(so, lca, LenB(lca))
        If ret = 0 Then
        For tp = 0 To 6
            ret = WSAIoctl(so, SIO_RCVALL + tp, ByVal VarPtr(bin), 4&, ByVal 0&, ByVal 0&, ByVal VarPtr(bret&), ByVal 0&, ByVal 0&)
            gStopget = False
            If ret = 0 Then Exit For
        Next
        
            If ret& = 0 Then
                 installSniffer = 0
                 Else
                installSniffer = -1
            End If
        End If
    Else
    uninstallSniffer
    End If
End If
Exit Function
fim:
installSniffer = -1
Debug.Print "No IP address"
End Function
'---------------------------------------------------------------------
Public Function uninstallSniffer()
Debug.Print "Stoped..."
Call closesocket(so)
Call WSACleanup
End Function

'----------------------------LOCAL FUNCTION
Private Function GetAddressLong(ByVal strHostname As String) As Long
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
'--------------------------------------------------------
Private Function hostAdress(ByVal ptrStr As Long) As String
Dim hp(3) As Byte
Call CopyMemSock(ByVal VarPtr(hp(0)), ByVal VarPtr(ptrStr), 4)

hostAdress = hp(0) & "." & hp(1) & "." & hp(2) & "." & hp(3)
End Function

Public Function receivePacket()
  Static tmi As Double
  Dim ret&, proto&, dof&
  
    If gStopget = True Then
      Exit Function
    End If '
    
    FillMemory buffer(0), BL, 0&
    
    ret = recv(so, buffer(0), BL, 0&)
    
    If ret > LenB(iph) Then
    
      proto& = buffer(9)
      bNow = 0
      Select Case proto
        Case TCP
        tProt = "TCP"
        'dof& = 20 + (buffer(32) \ &H10) * 4
        'ReDim strBuffer((ret - dof&)) As Byte
        
        bNow = ret& + 40&
            'Debug.Print "Protocol TCP Data Off " & dof & " " & bNow
        Case UDP
        tProt = "UDP"
        'bNow = 2 * (ret - 28)
            'Debug.Print "Protocol UDP " & bNow
        Case ICMP
        tProt = "ICMP"
            bNow = ret - 28
            Debug.Print "Protocol ICMP " & bNow
       End Select
            
            Call CollectBytesTobeSelled
        Else
        
    End If
    
        

End Function

Private Function CollectBytesTobeSelled()
Call ReadSourceDest

        If InStr(1, iph.m_srcaddr, myNetID) <> 0 Then '//Trafico interno
               'Debug.Print bNow; "TRF.Int. Descartado  " & bNow, iph.m_srcaddr, iph.m_destaddr
               bNow = 0 '//descartar
               'fromSite = tProt & ": Internal Pkt. from: " & iph.m_srcaddr & " to: " & iph.m_destaddr
            Else
               'trafico Externo...
             If InStr(1, iph.m_destaddr, localIP) = 0 Then '//externo mas not for me
               'Debug.Print bNow; "TRF.Ext. Descartado  " & bNow, iph.m_srcaddr, iph.m_destaddr
               bNow = 0 '//descartar
               'fromSite = tProt & ": External Pkt. from: " & iph.m_srcaddr & " to: " & iph.m_destaddr
             End If
             
        End If
            
        
        If bNow > 0 Then Call GetShowNetDet
        
      

End Function

Private Function ReadSourceDest()
    iph.m_srcaddr = buffer(12) & "." & buffer(13) & "." & buffer(14) & "." & buffer(15)
    iph.m_destaddr = buffer(16) & "." & buffer(17) & "." & buffer(18) & "." & buffer(19)
End Function


Private Function GetShowNetDet()
netNow = netNow + 1.25 * bNow
netCharge = netCharge + (1.25 * bNow)


fromSite = "Net: Charge on this Section - " & trasnBytes(netNow) & ", Card Charge Total:" & trasnBytes(netCharge)
        
        'update to show activity
'If frmTip.status.Panels("from").Text <> fromSite Then _
'frmTip.status.Panels("from").Text = fromSite
        
End Function
Private Sub getIPmyNetID()
localIP = hostAdress(GetAddressLong(Environ$("Computername")))
myNetID = Left(localIP, InStrRev(localIP, "."))

End Sub

Public Function trasnBytes(ByVal bytes As Long) As String
Dim tmps$
Const bt = 0#
Const KB = 1024#
Const MB = KB * KB
Const GB = MB * KB

Select Case bytes
Case bt To KB - 1
   tmps$ = Format(bytes, "0.00") & " Bs"
Case KB To MB - 1
   tmps$ = Format(bytes / KB, "0.00") & " KB"
Case MB To GB - 1
   tmps$ = Format(bytes / MB, "0.00") & " MB"
Case Else
   tmps$ = Format(bytes / GB, "0.00") & " GB"
End Select

trasnBytes = tmps$
End Function
