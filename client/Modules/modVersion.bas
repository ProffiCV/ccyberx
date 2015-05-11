Attribute VB_Name = "modVersion"
Option Explicit

Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Const DRIVE_REMOVABLE = 2
Public Restart As Boolean
Private lptHandle As Long
Private lpVLen As Long
Private unicodVersion As String

Private Type VERSION
    ver As Long
End Type
Private vr As VERSION
Public Function QueryNewVersion(site) As String
Dim ret As Long
Dim tmp As String

lpVLen = GetFileVersionInfoSize(site, lptHandle)
If lpVLen <> 0 Then
    unicodVersion = Space(lpVLen + 20)
    ret = GetFileVersionInfo(site, 0, lpVLen, ByVal unicodVersion)
    If ret <> 0 Then 'readed
    tmp = Space(400)
    ret = 400
        Debug.Print VerQueryValue(ByVal unicodVersion, "FileVersion", ByVal VarPtr(tmp), ByVal ret)
    End If
    
End If

End Function

Public Function tryUpdate()
Dim it&, root$
showTips "Searching for base module please wait..."
For it = 68 To 74
If GetDriveType(Chr(it) & ":") = DRIVE_REMOVABLE Then
  root$ = Chr(it)
  If Dir(root$ & ":\updccxv3\ccxv3.upd") <> "" Then
   doUpdate root$ & ":\updccxv3"
   Exit Function
  End If
  
End If
Next
showTips "Base module not found..."
End Function

Function doUpdate(ByVal site As String)
Dim bt
If Dir(site & "\ccxv3.upd") <> "" Then
frmCli.mget.Enabled = False
  Open site & "\ccxv3.upd" For Binary Access Read As #1
    Open App.Path & "\ccxv3.upd" For Output As #2
     Do
       bt = Input(50, 1)
       Print #2, bt;
       DoEvents
       frmlogc.Label3(0).Caption = "LET USBS!!!"
       tell "Updating Image " & Format((Loc(1) / LOF(1)) * 100, "0") & "%"
     Loop Until Loc(1) >= LOF(1)
    Close #2
  Close #1
  showTips "Updated Success. Restarting cCyberXV3...", 3
  Pause 3
  Restart = True
  releaseMerlin True
Else
End If

End Function
