Attribute VB_Name = "Module1"
Option Explicit

Public Enum FLGSHUT
    EWX_FORCE = &H4
    EWX_LOGOFF = &H0 Or EWX_FORCE
    EWX_SHUTDOWN = &H1 Or EWX_FORCE
    EWX_REBOOT = &H2 Or EWX_FORCE
    EWX_POWEROFF = &H8 Or EWX_FORCE
    EWX_FORCEIFHUNG = &H10
End Enum


'Case "SAVE"
'          td$ = Right(dta$, Len(dta$) - 4)
'          If Len(td) > 50 Then td = Left(td, 51)
'            If Len(td$) = 51 Then
'            tell "Updating card " & mCards(1).code & " Source:" & pci(1).pcName
'                addUpdateCard mCards(1).code, td$
'            End If
'Case "OUTM" 'ver
'          td$ = Right(dta$, Len(dta$) - 4)
'            If Len(td$) = 51 Then
'            tell "Logoff Process Detected by " & pci(1).clientID & " Src:" & pci(1).pcName
'
'                  addUpdateCard mCards(1).code, td$
'
'                  Dim dtl As DETAILS
'
'
'
'                  .logoff = Format(Now, "dd/mm/yy hh:mm:ss")
'                  .state = "FREE"
'
'                  dtl.pc = pci(1).pcName
'                  dtl.din = pci(1).login
'                  dtl.dout = pci(1).logoff
'                  dtl.Nick = pci(1).clientID
'                  dtl.tmv = pci(1).pcuTime
'
'                  CopyMemory mCards(1), ByVal td$, 51
'
'                  With mCards(1)
'                      dtl.data = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
'                  End With
'
'                  addDetails dtl
'
'            End If
