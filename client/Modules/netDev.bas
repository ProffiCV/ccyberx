Attribute VB_Name = "netDev"
Option Explicit

Public Function detectNetWorkDevices(Optional lst As ListView)
Dim a, b
Dim dvn$
Set a = GetObject("Winmgmts:{ImpersonationLevel=Impersonate}root\cimv2").Instancesof("Win32_NetworkAdapter")

If Not a Is Nothing Then
    For Each b In a
    dvn$ = b.ProductName
    DoEvents
        If 1 Then
        dvn$ = Trim$(" " & b.ServiceName)
        Select Case dvn$
           Case "PSched", "Rasl2tp", "PptpMiniport", "RasPppoe", "Raspti", "NdisWan", "PSched"
           
           Case Else
            dvn$ = Trim$(" " & b.MACAddress)
                If dvn$ <> "" Then
'                    Debug.Print b.ProductName, b.MACAddress, b.PNPDeviceID
                    If Not IsMissing(lst) Then
                        lst.ListItems.Add , , b.ProductName
                        lst.ListItems(lst.ListItems.count).SubItems(1) = b.MACAddress
                        lst.ListItems(lst.ListItems.count).SubItems(2) = b.Manufacturer
                        lst.ListItems(lst.ListItems.count).Tag = detectNetWorkIP(b.DeviceID)
                        
                    End If
                    
                End If
            
        End Select
        End If
        
    Next
    
End If

Set a = Nothing
End Function

Public Function detectNetWorkIP(Name As Long) As String
Dim a, b
Dim dvn$
Set a = GetObject("Winmgmts:{ImpersonationLevel=Impersonate}root\cimv2").Instancesof("Win32_NetworkAdapterConfiguration")

If Not a Is Nothing Then
    For Each b In a
        If b.Index = Name Then
        detectNetWorkIP = "" & CStr(b.IpAddress(0))
        Exit For
        End If
        
    Next
    
End If

Set a = Nothing
End Function

