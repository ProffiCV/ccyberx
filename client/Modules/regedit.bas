Attribute VB_Name = "regedit"
Option Explicit
'COntantes

'9049-1606-9727-9800

Private Type POLICIE_ENTRY 'entrada da politica
pAlias As String    'ex NoDriveTypeAutoRun
pKey As String
pValue As Long    'Dowrd:2
End Type

Private bpol() As POLICIE_ENTRY

'''''''''''''''''''''
Private Const REG_DWORD = 4
Private Const REG_SZ = 1

Private Const HKEY_USERS = &H80000003
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000

''''''''''''''''''WRITE
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4



Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or _
                           KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

''''''''''''''''''READ
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10

Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
                          KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or _
                          KEY_NOTIFY) And (Not SYNCHRONIZE))

'''''''''''''''
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
(ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RefreshPolicyEx Lib "userenv" (ByVal bMachine As Boolean, _
ByVal dwOptions As Long) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
(ByVal hKey As Long, ByVal lpValueName As String) As Long

'Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
'lpRect As RECT, ByVal bErase As Long) As Long

Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Public Declare Function WaitForSingleObjectEx Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long


Public userSID As String
'find User SID username
'regOpenKey
'regCreatekey
'regRemoveKey
'regCloseKey


Public Function getUserSID(Optional ByVal usern$ = "") As String
Dim slen As Long
Dim a, b
Set a = GetObject("Winmgmts:{ImpersonationLevel=Impersonate}\Root\cimv2").Instancesof("Win32_UserAccount")

If usern$ = "" Then usern$ = Environ$("Username")
For Each b In a
DoEvents
If StrComp("" & b.Name, usern$, vbTextCompare) = 0 Then
usern = b.SID
Exit For
End If

Next

If Len(usern) < 30 Then usern = ""

getUserSID = usern
End Function

 
Public Function AplyPolicies()
Dim fref As Long
Dim polFile$, id$, oldid$, vbt&, fName$, fVal&, pcouunt As Long
Dim tPol, pKey As String   'Explorer, Sstem etc.

If debugMode = True Then
'Debug.Print "AplyPolicies on Debug Mode "
Exit Function
End If

renameItens

Const asig$ = "EDSON REGEDIT POLICIES FOR CYBER USER"
polFile = App.Path & "\ccyberXV2.ini"
If Dir(polFile) <> "" Then

fref = FreeFile
Open polFile For Input As fref

Do Until EOF(fref)
Line Input #fref, id$
    
        If id = asig$ Then
        oldid = asig$
        ElseIf oldid = asig$ Then
        'normal read...
       
        vbt& = InStr(id, vbTab) - 1 'ver se é válido...
        If (vbt&) <> -1 Then
            fName$ = Left(id$, vbt&)
            fVal& = Val("0" & Right(id, Len(id) - vbt))
            'DisableTaskMgr 2
            pcouunt = pcouunt + 1
            ReDim Preserve bpol(pcouunt)
            tPol = Split(id, vbTab, 2)
            
           If tPol(0) <> "" And tPol(1) <> "" Then
           tPol = Array(tPol(0), tPol(1))
            With bpol(pcouunt - 1)
                .pKey = pKey
                .pAlias = tPol(0)
                .pValue = CLng("0" & tPol(1))
            End With
            
           End If
           
        Else
        'check if is a Key Indicator
        If InStr(id, "[") <> 0 Then
        pKey = Replace(id, "[", "")
        pKey = Replace(pKey, "]", "")
        
        End If
        
        End If
        
        
        'end
        End If
        


Loop
Close fref

'if
If UBound(bpol) <> 0 Then
   Call ApplyPolicies
End If


End If

fim:
End Function

Public Sub showDesktopIcons()
tell "Updating Desktop..."
Dim subkey$, rootkey$, handle&, res&
Const mypc$ = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
Const net$ = "{871C5380-42A0-1069-A2EA-08002B30309D}"
rootkey$ = getUserSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel\"
handle& = getSubKeyHandle(rootkey$)
If handle <> 0 Then
    'my pc
    res& = 0
    res& = RegSetValueEx(handle, mypc$, 0&, REG_DWORD, res, Len(res))
    'Internet icon
    res& = 0
    res& = RegSetValueEx(handle, net$, 0&, REG_DWORD, res, Len(res))
    
    res = RegCloseKey(handle&)
End If
End Sub

'rename INternet Explorer
Public Sub renameItens(Optional reset As Boolean = False)
tell IIf(reset = True, "Restoring", "Seting " & "itens names.")
showDesktopIcons
Dim subkey$, rootkey$, handle&, res&, valsz$
Const mypc$ = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
Const net$ = "{871C5380-42A0-1069-A2EA-08002B30309D}"
rootkey$ = getUserSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\"

'internet explorer
subkey$ = rootkey & "{871C5380-42A0-1069-A2EA-08002B30309D}"
handle& = getSubKeyHandle(subkey$)
If handle <> 0 Then
    res& = 0
    valsz$ = IIf(reset = False, "cCyberXV2 Internet", "Internet Explorer")
    res& = RegSetValueEx(handle, vbNullString, 0&, 1, ByVal valsz$, Len(valsz$))
    res = RegCloseKey(handle&)
End If

'my computer
subkey$ = rootkey & "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
handle& = getSubKeyHandle(subkey$)
If handle <> 0 Then
    res& = 0
    valsz$ = IIf(reset = False, "cCyberXV2 Devices", IIf(getLanguage = "PT", "O meu Computador", "My Computer"))
    res& = RegSetValueEx(handle, vbNullString, 0&, 1, ByVal valsz$, Len(valsz$))
    res = RegCloseKey(handle&)
End If
End Sub

Private Function aplSystem()
Dim subkey As String, rootkey, handle&, ret&
rootkey = userSID & "\Software\Microsoft\Windows\CurrentVersion\Policies\"
'Explorer
subkey = rootkey & "System"
handle& = getSubKeyHandle(subkey)
If handle <> 0 Then
    applySystemPoliciesIfAny handle&
    ret = RegCloseKey(handle&)
End If

End Function

Private Function aplNomEnum()
Dim subkey As String, rootkey, handle&, ret&
rootkey = userSID & "\Software\Microsoft\Windows\CurrentVersion\Policies\"

subkey = rootkey & "NonEnum"
handle& = getSubKeyHandle(subkey)
If handle <> 0 Then

    applyNoEnumPoliciesIfAny handle&
    ret = RegCloseKey(handle&)
End If

End Function
Private Function aplExplorer()
Dim subkey As String, rootkey, handle&, ret&

rootkey = userSID & "\Software\Microsoft\Windows\CurrentVersion\Policies\"

subkey = rootkey & "Explorer"
handle& = getSubKeyHandle(subkey)
If handle <> 0 Then
    applyExplorerPoliciesIfAny handle&
    ret = RegCloseKey(handle&)
End If

'System...
End Function

Private Function aplInstaller()
Dim subkey As String, rootkey, handle&, ret&

rootkey = userSID & "\Software\Policies\Microsoft\Windows\"

subkey = rootkey & "Installer"
handle& = getSubKeyHandle(subkey)
If handle <> 0 Then
    applyInstallerPoliciesIfAny handle&
    ret = RegCloseKey(handle&)
End If

'System...
End Function

Private Function ApplyPolicies()
Dim subkey As String, rootkey, handle&, ret&
If goend = False Then frmStart.lbpg = "Applying user policies..."
Call aplSystem
Call aplSystem

Call aplNomEnum
Call aplNomEnum

Call aplInstaller
Call aplInstaller

Call aplExplorer
Call aplExplorer
'System...
If goend = False Then frmStart.lbpg = "Waiting Policiy updater to finish..."
Call waitExplorer

End Function

Public Function clearPolicies()
renameItens True
tell "Reseting usr Settings..."
Dim subkey As String, rootkey$, handle&, ret&
rootkey = userSID & "\Software\Microsoft\Windows\CurrentVersion\Policies"
'Explorer
Debug.Print rootkey
handle& = getSubKeyHandle(rootkey)

If handle <> 0 Then
    subkey = "Explorer"
    ret& = RegDeleteKey(handle&, subkey)
    
    subkey = "System"
    ret& = RegDeleteKey(handle&, subkey)
    subkey = "NonEnum"
    ret& = RegDeleteKey(handle&, subkey)
    
    ret = RegCloseKey(handle&)
End If


rootkey = userSID & "\Software\Policies\Microsoft\Windows"
handle& = getSubKeyHandle(rootkey)

If handle <> 0 Then
    subkey = "Installer"
    ret& = RegDeleteKey(handle&, subkey)
    ret = RegCloseKey(handle&)
End If


Dim buffer$

Call waitExplorer(0)
End Function

Private Function getSubKeyHandle(ByVal subkey As String) As Long
Dim ret&, handle&
Static lop As Long
handle = -1
ret = RegOpenKeyEx(HKEY_USERS, subkey, 0&, KEY_WRITE, handle)

If ret = 2 Then
ret = RegCreateKey(HKEY_USERS, subkey, handle)
ret& = RegCloseKey(handle)
lop = lop + 1
If lop < 2 Then
getSubKeyHandle subkey
Else
lop = 0
'break loop
End If

End If

If ret = 0 Then
Else
End If

getSubKeyHandle = handle&
End Function

'Explorer
Private Function applyExplorerPoliciesIfAny(ByVal handle As Long) As Boolean
Dim ret&, limit&, res&
limit& = UBound(bpol)
If limit <> 0 Then
    For ret = 0 To limit
        If bpol(ret).pKey = "Explorer" Then
'            Debug.Print "Explorer " & bpol(ret).pAlias, bpol(ret).pValue
            res& = RegSetValueEx(handle, bpol(ret).pAlias, 0&, REG_DWORD, bpol(ret).pValue, Len(bpol(ret).pValue))
        End If
        
    Next
End If

End Function

'Installer
Private Function applyInstallerPoliciesIfAny(ByVal handle As Long) As Boolean
Dim ret&, limit&, res&
limit& = UBound(bpol)
If limit <> 0 Then
    For ret = 0 To limit
    Debug.Print bpol(ret).pKey
        If bpol(ret).pKey = "Installer" Then
'            Debug.Print "Explorer " & bpol(ret).pAlias, bpol(ret).pValue
            res& = RegSetValueEx(handle, bpol(ret).pAlias, 0&, REG_DWORD, bpol(ret).pValue, Len(bpol(ret).pValue))
        End If
        
    Next
End If

End Function

'System
Private Function applySystemPoliciesIfAny(ByVal handle As Long) As Boolean
Dim ret&, limit&, res&
limit& = UBound(bpol)
If limit <> 0 Then
    For ret = 0 To limit
        If bpol(ret).pKey = "System" Then
'            Debug.Print "System " & bpol(ret).pAlias, bpol(ret).pValue
           res& = RegSetValueEx(handle, bpol(ret).pAlias, 0&, REG_DWORD, bpol(ret).pValue, Len(bpol(ret).pValue))

        End If
        
    Next
End If

End Function

'NoEnum

Private Function applyNoEnumPoliciesIfAny(ByVal handle As Long) As Boolean
Dim ret&, limit&, res&
limit& = UBound(bpol)
If limit <> 0 Then
    For ret = 0 To limit
        If bpol(ret).pKey = "NonEnum" Then
'            Debug.Print "NonEnum " & bpol(ret).pAlias, bpol(ret).pValue
            res& = RegSetValueEx(handle, bpol(ret).pAlias, 0&, REG_DWORD, bpol(ret).pValue, Len(bpol(ret).pValue))
        End If
        
    Next
End If

End Function

Public Function burnPolicieFile()
Dim polFile$
Dim ffile As Long
polFile = App.Path & "\ccyberXV2.ini"

If Dir(polFile) <> "" Then Exit Function
ffile = FreeFile
Open polFile For Output As #ffile
Print #ffile, "EDSON REGEDIT POLICIES FOR CYBER USER"
Print #ffile, "[System]"
Print #ffile, "DisableTaskMgr" & vbTab & 2
Print #ffile, "DisableLockWorkStation" & vbTab & 2
Print #ffile, "DisableChangePassword" & vbTab & 2
Print #ffile, "DisableStatusMessages" & vbTab & 2
Print #ffile, "DisableRegistryTools" & vbTab & 2
Print #ffile, "VerboseStatus" & vbTab & 0
Print #ffile, "WallpaperStyle" & vbTab & 2
Print #ffile, ""
Print #ffile, "[Explorer]"
Print #ffile, "NoClose" & vbTab & 2
Print #ffile, "NoDrives" & vbTab & 4
Print #ffile, "NoViewOnDrives" & vbTab & 4
Print #ffile, "NoDriveTypeAutoRun" & vbTab & 255
Print #ffile, "NoStartMenuSubFolders" & vbTab & 2
Print #ffile, "NoWindowsUpdate" & vbTab & 2
Print #ffile, "NoCommonGroups" & vbTab & 2
Print #ffile, "NoSMMyDocs" & vbTab & 2
Print #ffile, "NoRecentDocsMenu" & vbTab & 2
Print #ffile, "NoSetFolders" & vbTab & 2
Print #ffile, "NoNetworkConnections" & vbTab & 2
Print #ffile, "NoFavoritesMenu" & vbTab & 2
Print #ffile, "NoFind" & vbTab & 2
Print #ffile, "NoSMHelp" & vbTab & 2
Print #ffile, "NoRun" & vbTab & 2
Print #ffile, "NoSMMyPictures" & vbTab & 2
Print #ffile, "NoStartMenuMyMusic" & vbTab & 2
Print #ffile, "NoStartMenuNetworkPlaces" & vbTab & 2
Print #ffile, "ForceStartMenuLogOff" & vbTab & 2
Print #ffile, "StartMenuLogOff" & vbTab & 2
Print #ffile, "NoClose" & vbTab & 2
Print #ffile, "NoChangeStartMenu" & vbTab & 2
Print #ffile, "NoSetTaskbar" & vbTab & 2
Print #ffile, "NoTrayContextMenu" & vbTab & 2
Print #ffile, "NoRecentDocsHistory" & vbTab & 2
Print #ffile, "ClearRecentDocsOnExit" & vbTab & 2
Print #ffile, "NoInstrumentation" & vbTab & 2
Print #ffile, "NoResolveSearch" & vbTab & 2
Print #ffile, "NoResolveTrack" & vbTab & 2
Print #ffile, "GreyMSIAds" & vbTab & 2
Print #ffile, "NoAutoTrayNotify" & vbTab & 2
Print #ffile, "LockTaskbar" & vbTab & 2
Print #ffile, "NoSMBalloonTip" & vbTab & 2
Print #ffile, "NoStartMenuPinnedList" & vbTab & 2
Print #ffile, "NoStartMenuMFUprogramsList" & vbTab & 2
Print #ffile, "NoStartMenuMorePrograms" & vbTab & 2
Print #ffile, "NoToolbarsOnTaskbar" & vbTab & 2
Print #ffile, "NoSMConfigurePrograms" & vbTab & 2
Print #ffile, "NoPropertiesMyDocuments" & vbTab & 2
Print #ffile, "NoPropertiesMyComputer" & vbTab & 2
Print #ffile, "NoPropertiesRecycleBin" & vbTab & 2
Print #ffile, "NoNetHood" & vbTab & 2
Print #ffile, "NoInternetIcon" & vbTab & 0
Print #ffile, "NoRecentDocsNetHood" & vbTab & 2
Print #ffile, "DisablePersonalDirChange" & vbTab & 2
Print #ffile, "NoCloseDragDropBands" & vbTab & 2
Print #ffile, "NoMovingBands" & vbTab & 2
Print #ffile, "NoSaveSettings" & vbTab & 2
Print #ffile, "NoDesktopCleanupWizard" & vbTab & 2
Print #ffile, "ForceActiveDesktopOn" & vbTab & 0
Print #ffile, "NoActiveDesktop" & vbTab & 2
Print #ffile, "NoControlPanel" & vbTab & 2
Print #ffile, "NoAutoUpdate" & vbTab & 2
Print #ffile, "NoFolderOptions" & vbTab & 2
Print #ffile, "NoNetConnectDisconnect" & vbTab & 2
Print #ffile, "NoManageMyComputerVerb" & vbTab & 2
Print #ffile, "NoComputersNearMe" & vbTab & 2
Print #ffile, "NoSharedDocuments" & vbTab & 2
Print #ffile, "NoWinKeys" & vbTab & 2
Print #ffile, ""
Print #ffile, "[NonEnum]"
Print #ffile, "{450D8FBA-AD25-11D0-98A8-0800361B1103}" & vbTab & 1
Print #ffile, "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" & vbTab & 0
Print #ffile, ""
Print #ffile, "[Installer]"
Print #ffile, "DisableMedia" & vbTab & 2
Close #ffile

If debugMode = True Then Shell "Notepad " & polFile, vbNormalFocus
End Function

''''''''''''''''''
Public Function waitExplorer(Optional ini As Long = 1)
On Error Resume Next
Dim handle&
'Debug.Print "Killing Explorer"
'find and kill explorer
If ini = 1 Then
handle = Shell("gpupdate /force", vbHide)
handle = OpenProcess(SYNCHRONIZE, False, handle)
Do
Loop Until WaitForSingleObjectEx(handle, 1000, True) = 0

Unload frmIcon
Load frmIcon
End If

Call CloseHandle(handle)
If goend = False Then frmStart.lbpg = "Reseting Explorer..."
handle = FindWindow("Progman", vbNullString)
Call hardKill(handle)
If goend = False Then frmStart.lbpg = "Waiting, Explorer to boot up..."

Do
Loop Until FindWindow("Shell_TrayWnd", vbNullString) <> 0
If goend = False Then frmStart.lbpg = "Almost done..."

If ini = 1 Then
Unload frmIcon
Load frmIcon
End If


'Debug.Print "Explorer reloaded"
End Function


Public Function getLanguage() As String
Const LANG_PORTUGUESE = 16
Const LANG_ENGLISH = 9
Dim userl As Integer, prlang&, sublang&
userl = GetUserDefaultLangID

prlang& = Hex(userl And 1023)
Select Case prlang&
    Case LANG_PORTUGUESE
        getLanguage = "PT"
    Case Else
        getLanguage = "EN"
End Select

End Function

Public Function enableAutoRun(ByVal lpzRegSZ$, ByVal lpzRegVal$, ByVal autorun As Boolean) As Boolean
'adiciona entrada lpzPath$ no regedit...
Dim gpath$, ghandle As Long, ret&
gpath$ = userSID & "\Software\Microsoft\Windows\CurrentVersion\Run"

ghandle = getSubKeyHandle(gpath$)

If ghandle <> -1 Then
    If autorun = False Then 'delete
        ret& = RegDeleteValue(ghandle, lpzRegSZ$)
    Else 'adicionar
        ret& = RegSetValueEx(ghandle&, lpzRegSZ$, 0&, REG_SZ, ByVal lpzRegVal$, Len(lpzRegVal$))
    End If
End If


Call RegCloseKey(ghandle)

enableAutoRun = CBool(ret& = 0&)
End Function
