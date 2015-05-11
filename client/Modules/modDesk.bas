Attribute VB_Name = "modDesk"
Option Explicit
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Any) As Long

'<job id="vbs">
'<script language="VBScript">
'   Set WshShell = WScript.CreateObject("WScript.Shell")
'   strDesktop = WshShell.SpecialFolders("Desktop")
'   Set oShellLink = WshShell.CreateShortcut(strDesktop & "\Shortcut Script.lnk")
'   oShellLink.TargetPath = WScript.ScriptFullName
'   oShellLink.WindowStyle = 1
'   oShellLink.HotKey = "CTRL+SHIFT+F"
'   oShellLink.IconLocation = "notepad.exe, 0"
'   oShellLink.Description = "Shortcut Script"
'   oShellLink.WorkingDirectory = strDesktop
'   oShellLink.Save
'
'   Set oUrlLink = WshShell.CreateShortcut(strDesktop & "\Microsoft Web Site.url")
'   oUrlLink.TargetPath = "http://www.microsoft.com"
'   oUrlLink.Save
'</script>
'</job>

Private Sub saveItens(ByVal what$, ByVal wher As Integer)
Dim itens$
If UCase(Environ$("Computername")) <> "EDSONPORTABLE" Then
itens$ = Dir(what$ & "\*")
If itens$ = "" Then Exit Sub
Do

Select Case wher
    Case 0 'desktop
        
        Call FileCopy(what$ & "\" & itens$, App.Path & "\oDesktop\" & itens$)
        Kill what$ & "\" & itens$
    Case 1 'allusers
         Call FileCopy(what$ & "\" & itens$, App.Path & "\oAllUser\" & itens$)
        Kill what$ & "\" & itens$
End Select

itens$ = Dir
  
Loop Until itens = ""
End If

End Sub

'limpar desktop
Public Function clearDesktopItens()
If UCase(Environ$("Computername")) <> "EDSONPORTABLE" Then
Dim WshShell, strDesktop$, itens$
Set WshShell = CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
saveItens strDesktop, 0
strDesktop = WshShell.SpecialFolders("AllUsersDesktop")
saveItens strDesktop, 1

Set WshShell = Nothing
End If

End Function

'usar do ccyberx
Public Function cCyberDesktop(Optional clear As Boolean = False)
If UCase(Environ$("Computername")) <> "EDSONPORTABLE" Then

Dim WshShell, strDesktop$, itens$
Set WshShell = CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
itens$ = Dir(App.Path & "\mydesk\*")
If itens$ = "" Then GoTo fim
Do

    If clear = True Then
    Open strDesktop & "\" & itens$ For Output As #4 'avoiding error
    Print #4, "Testing..."
    Close #4
      Kill strDesktop & "\" & itens$
    
    Else
            Call FileCopy(App.Path & "\myDesk\" & itens$, strDesktop & "\" & itens$)
    End If
    
itens$ = Dir
Loop Until itens = ""

fim:
Set WshShell = Nothing
End If

End Function

Public Function installcCyberDesktop()
If UCase(Environ$("Computername")) <> "EDSONPORTABLE" Then

If debugMode = True Then
'    Debug.Print "installcCyberDesktop on Debug Mode "
    Exit Function
Exit Function
End If



       cCyberDesktop True  'tirar o que e do ccyberx
       clearDesktopItens 'tirar resto
       cCyberDesktop 'usar ccyber
End If
       changeWallPaper

End Function

Public Function createFolders()
Debug.Print CreateDirectory(App.Path & "\oDesktop", ByVal 0)
Debug.Print CreateDirectory(App.Path & "\oAllUser", ByVal 0)
Debug.Print CreateDirectory(App.Path & "\myDesk", ByVal 0)
End Function
