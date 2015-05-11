Attribute VB_Name = "procMonitor"
Option Explicit
Public mydate As Double
Private pcnt& 'program count


Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const PROCESS_TERMINATE = &H1
Public Const SYNCHRONIZE = &H100000

Public Const SC_CLOSE = &HF060&
Public Const SC_MINIMIZE = &HF020&

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Const WM_SYSCOMMAND = &H112

Private fucR As Boolean

Public Function freeVigilant()
fucR = True
fucR = True
End Function



Public Function CloseThis(ByRef hwnds, ByVal pid&) As Long
Dim objClassName As String, lenReaded As Long, vret As Long

If debugMode = True Then
'    Debug.Print "CloseThis on Debug Mode " & hwnds, pid&
    Exit Function
Exit Function
End If

'Exit Function
'Debug.Print "Permission to KILL"
'GoTo fim
Dim myid&, pext&, gpro&
If pid <> -1 And pid <> GetCurrentProcessId() Then
lenReaded = 256
objClassName = Space$(lenReaded)

vret = GetClassName(hwnds, ByVal objClassName, lenReaded)

If vret <> 0 Then
    objClassName = Left(objClassName, vret)
    
    Select Case UCase(objClassName)
     Case "EXPLOREWCLASS", "CABINETWCLASS"
      CloseThis = SendMessage(hwnds, &H10, 0&, ByVal 0&)
      CloseThis = SendMessage(hwnds, WM_SYSCOMMAND, SC_CLOSE, 0&)
     Case Else
      gpro& = OpenProcess(PROCESS_TERMINATE, 0&, pid&)
        If gpro& Then
            myid = GetExitCodeProcess(gpro&, pext&)
            TerminateProcess gpro&, pext&
        End If
    End Select
    
    
End If


End If

fim:
End Function


Public Function killKnowProcesses(Optional ByVal protectPID As Long, Optional show As Boolean)
If debugMode = True Then Exit Function

Dim a As Object, b As Object, c, wndo As Long
Dim piddate As Double

Dim hmp&, hmk&
'Debug.Print "Working. Please wait..."
'Set a = GetObject("Winmgmts:{ImpersonationLevel=Impersonate}\root\cimv2").Instancesof("Win32_Process")

Call killVisibleProcesses
Set a = GetObject("Winmgmts:{ImpersonationLevel=Impersonate}\root\cimv2") '.Instancesof("Win32_Process")

Set c = a.ExecQuery("SELECT * FROM Win32_Process")
If Not c Is Nothing Then
    For Each b In c
    DoEvents
    hmp& = hmp& + 1
    If show = True Then Debug.Print "" & b.Name, b.processid, "" & b.ExecutablePath
                       
        If CLng(b.processid) <> protectPID Then
            wndo = findWindowByPid(CLng(b.processid))
                If wndo <> -1 Then
                Debug.Print b.Creationdate
                Debug.Print Format(Left("" & b.Creationdate, InStr(1, "" & b.Creationdate, ".")))
                If "" & b.Creationdate = "" Then
                piddate = 0
                Else
                piddate = CDbl("" & Format(Left("" & b.Creationdate, InStr(1, "" & b.Creationdate, "."))))
                End If
                    If piddate > mydate Then
                    If InStr("" & b.Name, "explorer.exe") <> 0 Then
                      showTips " Warning Microsoft Windows Explorer Restarted...", 1
                      UpdateMyDate
                      
                    Else
                    hmk = hmk + 1
                    hardKill wndo
                    Pause 0.08
                    End If
                    
                     showTips UCase(" " & b.Name) & vbCrLf & "" & b.ExecutablePath & vbCrLf & " should not run without login.", 1
                    
                    If show = True Then tell "Bad PID:" & b.processid & " [hWnd:" & wndo & "] Killed? " & CBool((ShowWindowAsync(wndo, 1) = 0))
                       
                    End If
                    
                End If
        Else
'        Debug.Print "Protected PID, Ignored ... " & protectPID
        End If
          
            
    Next
    
End If

'Debug.Print "Processes found " & hmp& & " Killed " & hmk
'refrescar a data...

Set b = Nothing
Set a = Nothing

End Function
Public Function hardKill(ByVal hwnd As Long)
Dim myid&, pext&, gpro&, pid&
myid = GetWindowThreadProcessId(hwnd, pid&)
gpro& = OpenProcess(PROCESS_TERMINATE, 0&, pid&)
myid = GetExitCodeProcess(gpro&, pext&)
TerminateProcess gpro&, pext&

End Function

Public Function UpdateMyDate()
mydate = CDbl(Format(Now, "yyyymmddhhmmss"))
End Function

Public Function killVisibleProcesses()
Dim myid, pid&
pcnt& = 0

If debugMode = True Then Exit Function
If LoggedInFlag = True Then Exit Function

myid = GetCurrentProcessId
EnumWindows AddressOf EnumWindowsProc, myid
End Function



Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lparam As Long) As Boolean
Dim pid As Long, wndt$, myid&

GetWindowThreadProcessId hwnd, pid

If pid <> lparam Then

    If IsWindowVisible(hwnd) = 1 And GetParent(hwnd) = 0 Then
    
        wndt$ = Space$(255)
        myid = GetWindowText(hwnd, wndt$, 256)
        If myid <> 0 Then
            wndt$ = Left(wndt$, InStr(1, wndt$, Chr(0)) - 1)
            If UCase(wndt$) <> "PROGRAM MANAGER" And InStr(1, wndt$, "XV3") = 0 Then
             pcnt& = pcnt& + 1
             tell "Killing! " & Format(Now, "hh:mm:ss") & ", " & wndt$
             CloseThis hwnd&, pid
             tell "Killed! " & Format(Now, "hh:mm:ss") & ", " & wndt$
        End If
        
        End If
        
    End If
    
EnumWindowsProc = False
End If

EnumWindowsProc = True

End Function

