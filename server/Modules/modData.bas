Attribute VB_Name = "modData"
Option Explicit
Public totalC As Long, Used As Long, News As Long

Public Function getDataBaseDetails(ByVal key As Long)
If key <> 3245 Then
    getDataBaseDetails = "Fuck you"
Else
        If dbOpenned = True Then
            testRSt
            totalC = 0
            rst.Open "SELECT * FROM Cards", con, adOpenStatic, adLockOptimistic
            If rst.EOF = False Then
                totalC = rst.RecordCount
            End If
            
            testRSt
            Used = 0
            rst.Open "SELECT * FROM Cards WHERE data like '%T%'", con, adOpenStatic, adLockOptimistic
            If rst.EOF = False Then
                Used = rst.RecordCount
            End If
            
            testRSt
            rst.Open "SELECT * FROM Cards WHERE data like '%R%'", con, adOpenStatic, adLockOptimistic
            If rst.EOF = False Then
                Used = Used + rst.RecordCount
            End If
            
            testRSt
            News = 0
            rst.Open "SELECT * FROM Cards WHERE data like '%F%'", con, adOpenStatic, adLockOptimistic
            If rst.EOF = False Then
                News = rst.RecordCount
            End If
        End If
   
    getDataBaseDetails = "Total Cards " & totalC & ", Used " & Used & ", New " & News & " OK:" & CBool((Used + News) = totalC)
End If

End Function
