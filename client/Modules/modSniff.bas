Attribute VB_Name = "modSniff"
Option Explicit

Public Type WSABUF 'Requires Windows Sockets 2.0
  len As Long 'u_long      // The length of the buffer
  buf As String 'char FAR *  // The pointer to the buffer
End Type

Public Type SOCKADDR 'Requires Windows Sockets 2.0 (This structure is used with TCP/IP)
  sin_family As Integer    'short
  sin_port   As Integer    'u_short
  sin_addr   As Long       'struct IN_ADDR
  sin_zero   As String * 7 'char[8]
End Type

Private Declare Function WSARecvFrom Lib "WS2_32.DLL" _
(ByVal hSocket As Long, ByRef lpBuffers As WSABUF, _
ByVal dwBufferCount As Long, ByRef lpNumberOfBytesRecvd As Long, _
ByRef lpFlags As Long, ByRef lpFrom As SOCKADDR, _
ByRef lpFromlen As Long, ByRef lpOverlapped As Any, _
ByVal lpCompletionRoutine As Long) As Long



Public Function rcvF(ByVal sock As Long)
Dim buffe As WSABUF
Dim Froma As SOCKADDR
Dim lFrm As Long
lFrm = Len(Froma)
Dim brec As Long
buffe.buf = Space(521)
buffe.len = 512
Debug.Print WSARecvFrom(sock, buffe, 1, brec, 0, Froma, lFrm, 0&, 0&)
Debug.Print Froma.sin_addr

Debug.Print Err.LastDllError
End Function
