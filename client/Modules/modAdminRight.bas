Attribute VB_Name = "modRunAs"

'Writed and compiled by edson martins and as an integrant part of
'ccyberx packet, this piece of code let ccyberx run under a nom previleged
'account and it will be capable to do everything that it was written to do
'Edson Martinsmindeo 2007, December 4th

Option Explicit

Private Const CREATE_DEFAULT_ERROR_MODE = &H4000000

Private Const LOGON_WITH_PROFILE = &H1
Private Const LOGON_NETCREDENTIALS_ONLY = &H2

Private Const LOGON32_LOGON_INTERACTIVE = 2
Private Const LOGON32_PROVIDER_DEFAULT = 0
   
Private Type STARTUPINFO
    cb As Long
    lpReserved As Long ' !!! must be Long for Unicode string
    lpDesktop As Long  ' !!! must be Long for Unicode string
    lpTitle As Long    ' !!! must be Long for Unicode string
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

  
Private Declare Function LogonUser Lib "advapi32.dll" Alias _
        "LogonUserA" _
        (ByVal lpszUsername As String, _
        ByVal lpszDomain As String, _
        ByVal lpszPassword As String, _
        ByVal dwLogonType As Long, _
        ByVal dwLogonProvider As Long, _
        phToken As Long) As Long

' CreateProcessWithLogonW API is available only on Windows 2000 and later.
Private Declare Function CreateProcessWithLogonW Lib "advapi32.dll" _
        (ByVal lpUsername As String, _
        ByVal lpDomain As String, _
        ByVal lpPassword As String, _
        ByVal dwLogonFlags As Long, _
        ByVal lpApplicationName As Long, _
        ByVal lpCommandLine As String, _
        ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, _
        ByVal lpCurrentDirectory As String, _
        ByRef lpStartupInfo As STARTUPINFO, _
        ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
      
Private Declare Function CloseHandle Lib "kernel32.dll" _
        (ByVal hObject As Long) As Long
                             
Private Declare Function SetErrorMode Lib "kernel32.dll" _
        (ByVal uMode As Long) As Long
        
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
                             
' Version Checking APIs
''''''''''''''''''''''''''''''''''''''
Public Type ARGS
    cmd As String
    arg As Variant
End Type

'''''''''''''''''''''''''''''''''''''
Private Declare Function GetVersionExA Lib "kernel32.dll" _
    (lpVersionInformation As OSVERSIONINFO) As Integer

Private Const VER_PLATFORM_WIN32_NT = &H2

'********************************************************************

'                   RunAsUser for Windows 2000 and Later
'********************************************************************
Public Function W2KRunAsUser(ByVal UserName As String, _
        ByVal Password As String, _
        ByVal DomainName As String, _
        garg As ARGS, _
        ByVal CurrentDirectory As String) As Long

    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    
    Dim wUser As String
    Dim wDomain As String
    Dim wPassword As String
    Dim wAppn As String
    Dim wPar As String
    Dim wCurrentDir As String
    
    Dim Result As Long
    
    si.cb = Len(si)
        
    wUser = StrConv(UserName + Chr$(0), vbUnicode)
    wDomain = StrConv(DomainName + Chr$(0), vbUnicode)
    wPassword = StrConv(Password + Chr$(0), vbUnicode)
    'garg.cmd = CurrentDirectory & Dir(CurrentDirectory & garg.cmd + "*")
    wAppn = StrConv(garg.cmd + Chr$(0), vbUnicode)
    'wPar = StrConv(garg.arg + Chr$(0), vbUnicode)
    wCurrentDir = StrConv(CurrentDirectory, vbUnicode)
    
    Result = CreateProcessWithLogonW(wUser, wDomain, wPassword, _
          LOGON_WITH_PROFILE, 0, wAppn, _
          CREATE_DEFAULT_ERROR_MODE, 0&, wCurrentDir, si, pi)
    ' CreateProcessWithLogonW() does not
    If Result <> 0 Then
        CloseHandle pi.hThread
        CloseHandle pi.hProcess
        W2KRunAsUser = 0
    Else
        W2KRunAsUser = Err.LastDllError
        MsgBox "CreateProcessWithLogonW() failed with error " & Err.LastDllError, vbExclamation
    End If

End Function


Public Function RunAsUser(ByVal UserName As String, _
                ByVal Password As String, _
                ByVal DomainName As String, _
                cmdarg As ARGS) As Long

    Dim w2kOrAbove As Boolean
    Dim osinfo As OSVERSIONINFO
    Dim Result As Long
    Dim uErrorMode As Long
    Dim CurrentDirectory As String
    
    CurrentDirectory = getWdir(cmdarg.cmd)
    
    Result = W2KRunAsUser(UserName, Password, DomainName, _
             cmdarg, CurrentDirectory)
   
    RunAsUser = Result
End Function

Private Function getWdir(iimg As String) As String
Dim it&
it = InStrRev(iimg, "\", , vbTextCompare)
If it = 0 Then
 getWdir = Environ$("Systemroot") & "\System32\"
 Else
 getWdir = Left(iimg, it)
End If

End Function



