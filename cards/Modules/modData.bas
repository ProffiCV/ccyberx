Attribute VB_Name = "modData"
Option Explicit

Dim totalC As Long, Used As Long, News As Long

Public Function getDataBaseDetails(ByVal key As Long)
If key <> 3245 Then
    getDataBaseDetails = "Fuck you"
Else
    getDataBaseDetails = "Total Cards " & totalC & ", Used " & Used & ", New " & News
End If

End Function


'6640-1368-7827-1000
'0000-1111-2222-3333
