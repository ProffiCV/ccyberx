VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTr 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "sServerXV2"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1545
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   1545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmn 
      Interval        =   4000
      Left            =   810
      Top             =   180
   End
   Begin MSWinsockLib.Winsock sup 
      Left            =   180
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
End
Attribute VB_Name = "frmTr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Talk(ByVal what$)
    If Me.sup.state = 7 Then
        sup.SendData what$
    
    End If
    
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption & "."
End Sub

Private Sub sup_ConnectionRequest(ByVal requestID As Long)
If sup.state <> 0 Then sup.Close

sup.Accept requestID
tell "Data Base admin is now Connected..."
End Sub

Private Sub sup_DataArrival(ByVal bytesTotal As Long)
Dim dta$
dbName = App.Path & "\sServerXV2.mdb"
frmsv2.tmTime.Enabled = False
sup.GetData dta$

Select Case Left(dta$, 3)
    Case "DEL"
        If InStr(1, dta, "19812005") <> 0 Then
        closeDb
        If Dir(dbName) <> "" Then
        
        If countc(0) <> 0 Then
        closeDb
        Pause 0.3
        FileCopy dbName, App.Path & "\Killed" & Format(Now, "ddmm_hhmmss") & ".mdb"
        
        End If
        
              creatDB True 'refazer a base de dados
              Talk "MSGData Base Recreated at " & Environ$("COMPUTERNAME") & _
              ", " & Format(Now, "hh:mm:ss")
              OpenDB
        End If
        Else
        Talk "Troubles..."
        End If
        
    Case "GET"
        readAndSendData
    Case "NEX"
    'permission to send new...
    gcmd = "NEX"
    Case "SAV"
    If getandSave(Right(dta, Len(dta) - 3)) = True Then
    
        Talk "NEX" 'pedir nova conta...
    End If
    Case "REP"
    testRSt
    Talk "CNT" & countc()
    Case "DTL"
    Talk "MSG" & getDataBaseDetails(3245) & ", " & Format(Now, "dd-mm-yy hh:mm:ss")
End Select
frmsv2.tmTime.Enabled = True

End Sub


Private Sub tmn_Timer()
'Debug.print "Waiting for connection"
Select Case sup.state
    Case 0, 8, 9
    sup.Close
    sup.Listen
End Select

End Sub
