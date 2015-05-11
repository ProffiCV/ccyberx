Attribute VB_Name = "modSysGuard"
Option Explicit
''''''''''''''''
Private Const ANYSIZE_ARRAY = 1
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"

''''''''''''''
Public Enum FLGSHUT
    EWX_FORCE = &H4
    EWX_LOGOFF = &H0 Or EWX_FORCE
    EWX_SHUTDOWN = &H1 Or EWX_FORCE
    EWX_REBOOT = &H2 Or EWX_FORCE
    EWX_POWEROFF = &H8 Or EWX_FORCE
    EWX_FORCEIFHUNG = &H10
End Enum

Private Const ERROR_SUCCESS = 0&
''''''''''''''''
Private Const SE_PRIVILEGE_ENABLED = &H2

Private Type luid
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
        luid As luid
        Attributes As Integer
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Type LUIDD
     count As Integer
     luid As Long
     attr As Long
End Type


'''''''''''''''''''''''APIS
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, _
ByVal dwReserved As Long) As Long

Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" _
Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, _
lpLuid As Any) As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" _
(ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
(ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
NewState As Any, ByVal BufferLength As Long, _
PreviousState As Any, ReturnLength As Long) As Long


'''''REMOTE
Private Declare Function WTSOpenServer Lib "Wtsapi32.dll" _
Alias "WTSOpenServerA" (ByVal server As String) As Long

Public Declare Function WTSOpenCloseServer Lib "Wtsapi32.dll" _
(ByVal server As String) As Long

'''''''''''''''''


Public Function ShutdownSystem(what As FLGSHUT, Optional sysn As String = vbNullString) As Boolean
    Dim hToken As Long, cp&
    Dim oldl As Long
    Dim tkp As TOKEN_PRIVILEGES
    If sysn = vbNullString Then sysn = Environ$("COMPUTERNAME")
'    MsgBox "Go Open Pro token"
 
    cp& = GetCurrentProcess
    If (OpenProcessToken(cp&, _
        TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) <> 0) Then
'      MsgBox "Go Look previlege  "
        tkp.PrivilegeCount = 1
        If LookupPrivilegeValue(sysn, SE_SHUTDOWN_NAME, tkp.Privileges(0).luid) <> 0 Then
          tkp.PrivilegeCount = 1
          tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
'            MsgBox "Go Ajust Previlege" & hToken
           
            If AdjustTokenPrivileges(hToken, False, tkp, 0&, ByVal 0&, oldl) <> 0 Then
'            Debug.Print Err.LastDllError
                   If Err.LastDllError = ERROR_SUCCESS Then
'                    MsgBox "Go shut down Previlege " & tkp.Privileges(0).Attributes & " " & tkp.PrivilegeCount
                        If (ExitWindowsEx(what, 0&) <> 0) Then
'                            Debug.Print "Go Shut down"
                             'ShutdownSystem = True
                             End
                         Else
'                         MsgBox Err.LastDllError & oldl
                        End If
                    Else
                    
                    End If
                    
                    Else
'                    Debug.Print Err.LastDllError
'                    MsgBox Err.LastDllError & oldl
            End If
        End If
    End If

'ShutdownSystem = False
End
End Function

