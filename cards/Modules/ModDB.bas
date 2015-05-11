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

Public Type CARDO
    'mask cards
    id As String * 3        '[000]
    code As String * 19     '[0000-0000-0000-0000]
    tbal As String * 4      '[0000] 'changes
    life As String * 2      '[00] 'changes
    flag As String * 1      '[X]R reached , F false T used , N not in use, invalid details
    date As String * 6      '[000000]
    tusd As String * 4      '[0000]
    trem As String * 4      '[0000] 'cahnged
    bytes As String * 12    '[000000000000]
End Type

'076 1210-3612-7764-9159 0200 01 T 261106 0078 0122 000001842805
'ID    Code...........................
Private mycard As CARDI
Private oldcard As CARDO

Private con As Connection
Private rst As Recordset
Private cmd As Command

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
    
    dbOpenned = con.State = adStateOpen
    OpenDB = dbOpenned
End If

End Function

Public Function closeDb() As Long

If Not con Is Nothing And con <> "" Then con.Close
If Not rst Is Nothing And rst.State Then rst.Close

Set con = New Connection
Set cmd = New Command
Set rst = New Recordset
closeDb = True
End Function



Public Function creatDB(Optional kil As Boolean = False) As Boolean
dbName = App.Path & "\sSuper.mdb"
Dim ws As Workspace
Dim db As Database
Dim fl As TableDef


If Dir(dbName) <> "" Then
        If kil = True Then
        Debug.Print closeDb
         Kill dbName
         Else
         frmSup.tell "Data Base Checked... (OK)"
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
                    frmSup.tell "Data Base Created Successfully..."
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


'adiciona dados na base de dados
Public Function addUpdateCard(ByVal cod$, ByVal data$) As Boolean
If dbOpenned = True Then
Dim res As Long
If Len(data) <> 51 Then
addUpdateCard = False
Exit Function
End If


CopyMemory mycard, ByVal data$, 51

    If cod$ <> mycard.code Then
    frmSup.tell "Error Code less than Data.Code..."
    addUpdateCard = False
    Exit Function
        Else
    cmd.CommandText = "UPDATE Cards SET data = '" & data$ & _
    "' WHERE code = '" & cod$ & "'"
    '"0011111-2222-3333-444422100600F02000000000000000000"
    cmd.ActiveConnection = con
    
    cmd.Execute res, , adCmdText
    
    If res = 0 Then 'não existe criar...
    cmd.CommandText = "INSERT INTO Cards(code,Data) VALUES (" & _
    "'" & cod$ & "','" & data$ & "'" & _
    ")"
    
        cmd.Execute res, , adCmdText
        If res <> 0 Then
            addUpdateCard = True
        Else
            addUpdateCard = False
        End If
     Else
     addUpdateCard = True
    End If
    
    End If
Else
frmSup.tell "Operation permitted only if data Base is Openned"
End If

End Function

'ler dados do cartao
Public Function getCard(code As ListBox, data As ListBox) As String
Dim tdados$
If dbOpenned = True Then
testRSt
code.Clear
data.Clear
rst.Open "SELECT * FROM Cards  ORDER By code", con, adOpenStatic, adLockOptimistic

If rst.EOF = False Then
frmSup.pgb.Value = 0
frmSup.pgb.Min = 0
frmSup.pgb.Max = rst.RecordCount
Do

    DoEvents
    tdados$ = String(Len(mycard), "0")
    
     CopyMemory mycard, ByVal tdados, Len(tdados)
         
        
        tdados$ = Trim$(rst!data)
        tdados$ = Left(tdados$, 51)
     CopyMemory mycard, ByVal tdados, Len(tdados)
     code.AddItem mycard.code
     data.AddItem Left(tdados$, 51)
     rst.MoveNext
     frmSup.pgb.Value = frmSup.lscard.ListCount
     frmSup.tell Format(frmSup.lscard.ListCount / rst.RecordCount, "0%") & ", " & frmSup.lscard.ListCount & " card" & IIf(frmSup.lscard.ListCount > 1, "s", "") & " loaded! "
 Loop Until rst.EOF
Else
frmSup.tell "No Data Found in Data Base..."
End If
Else
frmSup.tell "Data Base not Openned..."
End If

getCard = Trim$(Left(tdados, 51))
End Function

'normaliza a base de dados para nova consulta
Public Function testRSt()
If Not rst Is Nothing Then
    If rst.State <> adRecOK Then rst.Close
End If

End Function

Public Function removecard(ByVal code As String) As Boolean
Dim res&
res& = 0
If dbOpenned = True Then
cmd.CommandText = "DELETE * From Cards where code ='" & code & "'"
cmd.ActiveConnection = con
cmd.Execute res&, , adCmdText
    removecard = res <> 0
End If

End Function


Public Function importDataFromOldContentor(ByVal fname, Optional user$, Optional pass$)
On Error GoTo fim
Dim con1 As New Connection
                            con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                           "Data Source=" & fname & ";" & _
                           "Jet OLEDB:Database Password=microbodix23pass1981;"
    frmSup.tell "Connecting to DB " & fname & "..."
    con1.Open
    
    Pause 0.8
    frmSup.tell IIf(con1.State = adStateOpen, "Connected ", "Error Opening ") & fname & "..."
    If con.State = adStateOpen Then
    testRSt
      rst.Open "SELECT * FROM CARDS ", con1, adOpenStatic, adLockOptimistic
      
      If rst.EOF = False Then
        rst.MoveFirst
        frmSup.pgb.Min = 0
        frmSup.pgb.Max = rst.RecordCount
        frmSup.lscard.Clear
       
       Dim tmd$
        Do
            DoEvents
            Debug.Print Len(rst!coddata)
           ' Debug.Print rst!codnum, rst!coddata, Len(rst!coddata)
           tmd$ = String(55, "0")
           CopyMemory oldcard, ByVal tmd$, 55
            tmd$ = Left(rst!coddata, 55)
            
            CopyMemory oldcard, ByVal tmd$, 55
            
            With oldcard
            tmd$ = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
            End With
            If Len(tmd$) = 51 Then
            If Trim$(rst!codnum) = oldcard.code Then
                frmSup.pgb.Value = frmSup.pgb.Value + 1
                 frmSup.lbp.Caption = "Read Progress (Good Cards)" & Format(frmSup.pgb.Value / frmSup.pgb.Max, "0%")
                frmSup.lscard.AddItem rst!codnum
                frmSup.lcData.AddItem tmd$
                
            End If
            
            End If
            
                rst.MoveNext
        Loop Until rst.EOF
        Pause 0.6
        If (frmSup.pgb.Max - frmSup.pgb.Value) = 0 Then
        frmSup.lbp.Caption = "Process completed successful (" & frmSup.lscard.ListCount & ")"
        Else
        frmSup.lbp.Caption = "Data Error Detected... Total " & (frmSup.pgb.Max - frmSup.pgb.Value)
        End If
        
      End If
      
    End If
    
    con1.Close
   
    
    

Set con1 = Nothing

Exit Function
fim:
MsgBox "Error " & Err.Number & " Is this an OLD ccyber Data Base?", vbQuestion, App.EXEName

End Function

Public Function getCardCount() As Long
testRSt
If con = "" Then OpenDB
rst.Open "Select * from cards", con, adOpenStatic

getCardCount = rst.RecordCount

End Function

