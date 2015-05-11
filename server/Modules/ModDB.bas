Attribute VB_Name = "ModDB"
Option Explicit

Public Type CARDI
    'mask cards
    id As String * 3        '[000]
    code As String * 19     '[0000-0000-0000-0000]
    date As String * 6      '[000000]
    life As String * 2      '[00]
    flag As String * 1      '[X]R reached , F false T used , N not in use, invalid details
    tbal As String * 4      '[0000]
    tusd As String * 4      '[0000]
    bytes As String * 12    '[000000000000]
End Type

Private myCard As CARDI
Public mCards(8) As CARDI 'each card in use...

Public con As Connection
Public rst As Recordset
Public cmd As Command

Public dbName As String
Public dbOpenned As Boolean
Public Function OpenDB() As Boolean
Set con = New Connection
Set cmd = New Command
Set rst = New Recordset

OpenDB = False

If Not con Is Nothing Then
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                           "Data Source=" & dbName & ";" & _
                           "Jet OLEDB:Database Password=microbodix23pass1981;"
    con.Open
    
    dbOpenned = con.state = adStateOpen
    OpenDB = dbOpenned
End If

End Function

Public Function closeDb() As Long

If Not con Is Nothing And con <> "" Then con.Close
If Not rst Is Nothing And rst.state Then rst.Close

Set con = New Connection
Set cmd = New Command
Set rst = New Recordset
closeDb = True
End Function



Public Function creatDB(Optional kil As Boolean = False) As Boolean
dbName = App.Path & "\sServerXV2.mdb"
Dim ws As Workspace
Dim db As Database
Dim fl As TableDef


If Dir(dbName) <> "" Then
        If kil = True Then
        Call closeDb
         Kill dbName
         Else
         tell "Data Base Checked... (OK)"
         OpenDB
         creatDB = True
         Exit Function
        End If
    
End If

''''''criar o esqueleto da base de dados...
creatDB = False
Set ws = Workspaces(0)
If Not ws Is Nothing Then
Set db = ws.CreateDatabase(dbName, dbLangGeneral, dbEncrypt Or dbVersion40)

    If Not db Is Nothing Then
            Set fl = db.CreateTableDef("Cards")
                If Not fl Is Nothing Then
                    With fl
                        .Fields.Append .CreateField("code", dbText, 19)
                        .Fields.Append .CreateField("Data", dbText)
                        
                        fl.Fields.Refresh
                        
                    End With
                    db.TableDefs.Append fl
                    db.TableDefs.Refresh
                    tell "Data Base Created Successfully..."
                    db.NewPassword "", "microbodix23pass1981"
                    creatDB = True
                Else
                MsgBox "Can not Create Table Entry Cards", vbCritical, App.EXEName
                End If
                
            
        Else
    MsgBox "Can not create Data Base for Cards Storage", vbCritical, App.EXEName
    End If
    
Else
MsgBox "Can not handle Workspaces at index=0", vbCritical, App.EXEName

End If


Set con = Nothing
Set cmd = Nothing
Set rst = Nothing

End Function


'adiciona ou actualiza dados na base de dados
Public Function addUpdateCard(ByVal cod$, ByVal data$) As Boolean
Dim res As Long
res = 0
'--------------------------------------------------
If dbOpenned = True Then
    If Len(data) <> 51 Then
        addUpdateCard = False
        Exit Function
    End If
    
    CopyMemory myCard, ByVal data$, 51
    interpretCard myCard
        If cod$ <> myCard.code Then
            tell "Error Code less than Data.Code..."
                addUpdateCard = False
             Exit Function
          Else
            cmd.CommandText = "UPDATE Cards SET data = '" & data$ & "' WHERE code = '" & cod$ & "'"
            
            cmd.ActiveConnection = con
            cmd.Execute res, , adCmdText
        
            If res = 0 Then 'não existe criar...
                cmd.CommandText = "INSERT INTO Cards(code,Data) VALUES ('" & cod$ & "','" & data$ & "')"
                cmd.Execute res, , adCmdText
            End If
        
        End If
End If

'result...
addUpdateCard = (res <> 0)
End Function

'ler dados da base de dados
Public Function getCard(ByVal cod$) As String
Dim tdados$
If dbOpenned = True Then
testRSt

rst.Open "SELECT data  FROM Cards WHERE code = '" & cod$ & "'", con, adOpenStatic, adLockOptimistic

If rst.EOF = False Then
tdados$ = String(Len(myCard), "0")
 tdados$ = Trim$(rst!data)
 CopyMemory myCard, ByVal tdados, Len(tdados)
Else

End If

End If

getCard = Trim$(Left(tdados, 51))
End Function

'normaliza a base de dados para nova consulta
Public Function testRSt()
If Not rst Is Nothing Then
    If rst.state <> adRecOK Then rst.Close
End If

End Function

Public Function interpretCard(crd As CARDI) As String
With crd
Debug.Print .id
Debug.Print .code
Debug.Print .date
Debug.Print .life
Debug.Print .flag
Debug.Print .tbal
Debug.Print .tusd
Debug.Print .bytes
End With

End Function


Public Function checkCardsAutoRemove()
Dim gc As CARDI, tmpc$
'Debug.print con
Dim dia&
dia = addLife()

If dia <> 0 Then
    testRSt
    rst.Open "SELECT * from Cards", con, adOpenStatic, adLockOptimistic
    If rst.EOF = False Then
    rst.MoveFirst
    
    Do
    tmpc$ = rst("data")
    CopyMemory gc, ByVal tmpc$, Len(tmpc$)
    'Debug.print gc.life
    If CLng(gc.life) < 99 Then 'not more than 99 days...
    gc.life = Format(CLng(gc.life) + 1, "00")
    End If
    
    tmpc$ = ""
    With gc
       tmpc$ = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
       rst("data") = tmpc$
       rst.Update
    End With
    
        DoEvents
        rst.MoveNext
        
    Loop Until rst.EOF = True
    
    
    End If
    
End If

End Function

''''''receive and ask for more cards from manager
Public Function readAndSendData()

testRSt
rst.Open "SELECT Data From Cards", con, adOpenStatic, adLockOptimistic
If rst.EOF = False Then
rst.MoveFirst
    Do
    DoEvents
        tell "Attention: Sending data to Admin. " & Format(rst.AbsolutePosition / rst.RecordCount, "0% ") & rst.AbsolutePosition & "/" & rst.RecordCount
        frmTr.Talk "TAK" & Left$(rst("data"), 51)
        rst.MoveNext
        If rst.EOF = False Then
        gcmd = ""
            Do
                DoEvents
                'Debug.print "Waiting gcmd " & gcmd <> "" & " " & rst.EOF
                If frmTr.sup.state <> 7 Then Exit Function
                tell "Attention: Sending data to Admin. " & Format(rst.AbsolutePosition / rst.RecordCount, "0% ") & rst.AbsolutePosition & "/" & rst.RecordCount
       
            Loop Until gcmd = "NEX"
        End If
        
    Loop Until rst.EOF
Else
frmTr.Talk "NOC"
End If


End Function
'recebe e guarda dados na base de dados
Public Function getandSave(ByVal data As String) As Boolean
Dim gcr As CARDI
 CopyMemory gcr, ByVal data, Len(gcr)
 tell "Card Manager is Updating cards..."
getandSave = addUpdateCard(gcr.code, data)
End Function

'informa numero de cartao na base de dados
Public Function countc(Optional op& = -1) As Long
testRSt
If dbOpenned = False Or con = "" Then OpenDB
If Not con Is Nothing Then
rst.Open "Select code from cards", con, adOpenStatic, adLockOptimistic
    If op = -1 Then
        If rst.EOF = False Then
            countc = rst.RecordCount
        Else
            countc = 0
        End If
    End If

countc = rst.RecordCount
Else
countc = 0
End If
End Function
''''''''''''

Public Function valCod(ByVal cod$, ByVal Index As Integer) As String
  Dim tmp$, rTmp$
  Dim Total As Long, Remain As Long
  
  tmp = getCard(cod$)
  If Len(tmp) = 51 Then
    If pci(Index).state = "BUSY" Then
            rTmp$ = "BUSY"
        Else
            CopyMemory mCards(Index), ByVal tmp$, 51
            Total = CLng(Val(mCards(Index).tbal))
            Remain = CLng(Val(mCards(Index).tusd))
            Total = Total - Remain
            If Total > 0 Then
                pci(Index).state = "BUSY"
                With pci(Index)
                    .login = Format(Now, "dd/mm/yyhh:mm:ss")
                    .balTotal = Format(CLng(Val(mCards(Index).tbal)) * 100, "0$00")
                    .balUsed = Format(CLng(Val(mCards(Index).tusd) * 100), "0$00")
                    .balRemain = Format((CLng(Val(mCards(Index).tbal)) - CLng(Val(mCards(Index).tusd))) * 100, "0$00")
                    .netTotal = trasnBytes(CLng(mCards(Index).bytes))
                    
                    .pcuTime = "00:00:00"
                    .pcuPrice = "0$00"
                    .netNow = "0 Bytes"
                    .netPrice = "0$00"
                    .dispInfo = 1
                    .code = cod$
                    .clientID = "Client::" & Chr(64 + Index)
                    .logoff = "Using..."
                End With
            rTmp$ = "FIXE" & tmp
            Else
            rTmp$ = "NOCR"
        End If
    End If
  Else
    If Len(tmp) = 0 Then
        rTmp$ = "WRGC"
    Else
        rTmp$ = "MSGR"
    End If
  End If
''''''''''''''
'''''''''''''''RESULT
valCod = rTmp$
End Function


'limpa detalhes da conta....
Public Function freeCards(Optional Index As Integer = 0)
Dim tbuff As String
tbuff = String(63, "0")
If Index <> 0 Then   'This card index
        CopyMemory mCards(Index), ByVal tbuff, 51
  Else 'All cards
    For Index = 1 To 8
        CopyMemory mCards(Index), ByVal tbuff, 51
    Next
End If
End Function


