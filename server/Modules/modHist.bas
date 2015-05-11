Attribute VB_Name = "modHist"
Option Explicit

Public Type DETAILS
    pc As String * 24
    data As String * 52
    din As String * 18
    dout As String * 18
    tmv As String * 8
End Type

Private myFile As String

Public Function addDetails(dtl As DETAILS)
myFile = App.Path & "\" & Trim$(dtl.pc) & ".txt"
Dim fln As Long
fln = FreeFile
Dim recn As Long

Open myFile For Random Access Write As fln Len = Len(dtl)
recn = FileLen(myFile) \ Len(dtl) + 1
    Put fln, recn, dtl
Close fln


End Function

Public Function getDetails(nod As TreeView, key As String, id As String)
Dim tgl As DETAILS, tot As Long
Dim fnam As String, fln As Long
Dim mca As CARDI, mk$
fln = FreeFile
fnam = App.Path & "\" & id & ".txt"
    If Dir(fnam) <> "" Then
    Open fnam For Random Access Read As fln Len = Len(tgl)
            tot = FileLen(fnam) \ Len(tgl)
            
            If tot = 0 Then Exit Function
            Do
                DoEvents
                Get fln, tot, tgl
                nod.Refresh
                
                'Debug.print Len(Trim$(tgl.data))
             
                If Len(Trim$(tgl.data)) = 51 Then
                CopyMemory mca, ByVal tgl.data, 51
                mk$ = key & (nod.Nodes(key).Children + 1)
                nod.Nodes.Add key, tvwChild, mk$, Trim$(mca.code), 2, 2
                nod.Nodes(mk$).Tag = _
                "User Data" & vbNewLine & _
                " Login" & vbTab & Trim(tgl.din) & vbNewLine & _
                " LogOff" & vbTab & Trim$(tgl.dout) & vbNewLine & _
                " Timer" & vbTab & Trim$(tgl.tmv) & vbNewLine & _
                String(40, "·") & vbNewLine & _
                "Card Data" & vbNewLine & _
                " Code" & vbTab & mca.code & vbNewLine & _
                " ID" & vbTab & mca.id & vbNewLine & _
                " Date" & vbTab & Format(mca.date, "00-00-00") & vbNewLine & _
                " Life" & vbTab & mca.life & " day" & IIf(CLng(mca.life) > 1, "s", "") & vbNewLine & _
                " Flag" & vbTab & mca.flag & vbNewLine & _
                String(40, "·") & vbNewLine & _
                "Balance" & vbNewLine & _
                " Total" & vbTab & Format(CLng(mca.tbal) * 100, "0$00") & vbNewLine & _
                " Used" & vbTab & Format(CLng(mca.tusd) * 100, "0$00") & vbNewLine & _
                " Remain" & vbTab & Format((CLng(mca.tbal) - CLng(mca.tusd)) * 100, "0$00") & vbNewLine & _
                String(40, "·") & vbNewLine & _
                "Internet" & vbNewLine & _
                "Charge" & vbTab & trasnBytes(CLng(mca.bytes))

                                     
                End If
                tot = tot - 1
            Loop Until tot <= 0
            
    End If
    
End Function

Public Function getPcs()
Dim fln As String
Dim cd As New Collection
fln = Dir(App.Path & "\*.txt")

Do
If fln <> "" Then
cd.Add Left(fln, Len(fln) - 4)
Else
tell "No History Found..."
Exit Do

End If

DoEvents
fln = Dir

Loop Until fln = ""
Set getPcs = cd
End Function

