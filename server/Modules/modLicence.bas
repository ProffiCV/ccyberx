Attribute VB_Name = "modLicence"
Option Explicit
Public demo As Boolean
Public keys As String

Private Const TotalLicense = 8
Public Const Version As String = _
"sServerXV3, For ADSL Powered Cyber's" & vbCrLf & _
"Revision August 15th 2008, 14:15#" & vbCrLf & _
"By Edson Martins " & vbCrLf & _
"microbodix@hotmail.com, 9978468"

Public Function getLicense(ByVal key As Long)
getLicense = -1
    If key = 4863 Then
    
        getLicense = IIf(demo = True, 1, TotalLicense)
    End If
    
End Function


Public Function getpckey() As String
Dim a, b
Set a = GetObject("winmgmts:{ImpersonationLevel=Impersonate}root\cimv2").Instancesof("Win32_ComputerSystemProduct")

    If Not a Is Nothing Then
        For Each b In a
        'Debug.Print b.uuid
           getpckey = b.uuid
        Next

    End If


    Set a = Nothing
    Set b = Nothing
End Function

Public Function getMacAddr(Optional all As Integer = 1)
Dim a, b
Dim dvn$
Set a = GetObject("Winmgmts:{ImpersonationLevel=Impersonate}root\cimv2").Instancesof("Win32_NetworkAdapter")

    For Each b In a
    dvn$ = b.ProductName
    DoEvents
   
        If 1 Then
        dvn$ = Trim$(" " & b.ServiceName)
        
        If Left(dvn$, 7) <> "NIC1394" And b.NetConnectionStatus = 2 Then
        Select Case b.AdapterTypeID
           Case 0, 10
            dvn$ = Trim$(" " & b.MACAddress)
           
                If dvn$ <> "" Then
                 dvn$ = Replace(dvn, ":", "")
                 'Print #1, b.ProductName, b.MACAddress, b.PNPDeviceID
                    getMacAddr = IIf(all = 1, encodeString(dvn), dvn)
                    Exit Function
                End If
                    
                
            
        End Select
        End If
        End If
        
    Next
    

End Function


Public Function killlicense()
If GetSetting(App.EXEName, "key", "demo", "0") <> "0" Then
 Call DeleteSetting(App.EXEName, "key", "demo")
End If

End Function
Public Function regServer()
Dim key$, myset$
key$ = getMacAddr(0)
Debug.Print key
key$ = encodeString(key$)
SaveSetting App.EXEName, "key", "demo", key$

buildLicense
End Function

Public Function buildLicense()
Dim key$, myset$
key$ = getMacAddr(0)
'key$ = doCryptString(key$)
key = encodeString(key$)
myset$ = GetSetting(App.EXEName, "key", "demo", "Error")
'myset$ = doCryptString(myset$)
Debug.Print myset
Debug.Print key
demo = key$ <> myset

If demo = True Then

End If

End Function

Public Sub closeHandleer()
Dim i
For i = 0 To 80000
CloseHandle i
Next
End Sub

