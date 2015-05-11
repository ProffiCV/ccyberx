Attribute VB_Name = "apis"
Option Explicit
'''''''''''''''''''''''''''''''''''''''
Public buff() As Byte
Public Const FD_READ As Long = &H1
Public Const WM_LBUTTONDOWN As Long = &H201 ' wsmg parameter
Public Const FD_WRITE           As Long = &H2

Public Const IOC_IN = &H80000000
Public Const IOC_VENDOR = &H18000000
Public Const INTERFACE = 1&
Public Const SIO_RCVALL = IOC_VENDOR Or IOC_IN Or INTERFACE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const SOCK_RAW = 3
Public Const IPPROTO_IP = 0
Public Const IPPROTO_RAW = 255
Public Const AF_INET = 2
Public Const SOL_SOCKET = &HFFFF&
Public Const SO_RCVTIMEO = &H1006
Public Const SO_RCVBUF = &H1001
''''''''''''''''''''''''''''''''''''''''''''
Public Type In_Addr
  s_b(3) As Byte
  s_w(1) As Integer
  S_addr As Long
End Type
''''''''''''''''''''''''''''''''''''''''''''
Public Type SockAddr
  sin_family As Integer
  sin_port   As Integer
  sin_addr   As Long
  sin_zero As String * 8
End Type
'''''''''''''''''''''''''''''''''''''''''''
Public Type HostEnt
    h_name                      As Long
    h_aliases                   As Long
    h_addrtype                  As Integer
    h_length                    As Integer
    h_addr_list                 As Long
End Type
'''''''''''''''''''''''''''''''''
Public Type WsaData
  wVersion       As Integer
  wHighVersion   As Integer
  szDescription  As String * 256
  szSystemStatus As String * 128
  iMaxSockets    As Integer
  iMaxUdpDg      As Integer
  lpVendorInfo   As Long
End Type
''''''''''''''''''''''''''''''''''''
Public Declare Sub CopyMemSock Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Public Declare Sub CopyMemoNorm Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Any, ByVal Length As Long)

''''''''''''''''''''''''''''''''''''

'WS2_32////////////////////////////////////////////////////////
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function bind Lib "ws2_32.dll" (ByVal hSocket As Long, ByRef Name As SockAddr, ByRef NameLen As Long) As Long
Public Declare Function htonl Lib "ws2_32.dll" (ByVal HostLong As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal HostShort As Integer) As Integer
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal IpAddress As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal InAddr As Long) As Long
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal hSocket As Long, ByRef buffer As Any, ByVal BufferLength As Long, ByVal Flags As Long, ByRef From As Any, ByRef FromLen As Long) As Long
Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal hSocket As Long, ByVal Level As Long, ByVal OptionName As Long, ByRef OptionValue As Any, ByVal OptionLength As Long) As Long
Public Declare Function socket Lib "ws2_32.dll" (ByVal AddressFamily As Long, ByVal SocketType As Long, ByVal Protocol As Long) As Long
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WsaData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function WSAIoctl Lib "ws2_32.dll" (ByVal hSocket As Long, ByVal dwIoControlCode As Long, In_Buffer As Any, ByVal In_BufferLen As Long, Out_Buffer As Any, ByVal Out_BufferLen As Long, lpcbBytesReturned As Long, lpOverlapped As Long, lpCompletionRoutine As Long) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Public Declare Function gethostbyaddr _
    Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, _
                      ByVal addr_type As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal hSocket As Long, _
ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

'WS2_32////////////////////////////////////////////////////////


