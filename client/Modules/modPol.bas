Attribute VB_Name = "modPol"
'Aplica as politicas de segurança para o computadors
Option Explicit

    Private Type PGROUP_POLICY_OBJECT
    
    End Type
    
    Private Type ASYNCCOMPLETIONHANDLE
    
    End Type

    Private Declare Function ProcessGroupPolicy Lib "edson" ( _
    ByVal dwFlag As Long, _
    ByVal hToken As Long, _
    ByVal hKeyRoot As Long, _
    pDeletedGPOList As PGROUP_POLICY_OBJECT, _
    pChangedGPOList As PGROUP_POLICY_OBJECT, _
    pHandle As ASYNCCOMPLETIONHANDLE, _
    pbAbort As Boolean, _
    pStatusCallback As PFNSTATUSMESSAGECALLBACK) As Long


    Public Function applyPolicy()
    
    End Function
    
