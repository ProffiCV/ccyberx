Attribute VB_Name = "modCryptoAPIEdson"
'this portion of code was developed by Edson Martins
'at Archote instalations on 9 october 2006 12:34

'use:
'-------------------CryptMsgCalculateEncodedLength
'-------------------CryptMsgOpenToEncode
'-------------------CryptMsgUpdate
'-------------------CryptMsgGetParam
'-------------------CryptMsgClose
'-------------------CryptMsgOpenToDecode
'-------------------
'-------------------
Private Const defPassword = "edson"
Private hMsg As Long
Private cbContent As Long
Private pbContent As String 'pointer to cbcontent
Private cbEncodedBlob As Long
Private pbEncodedBlob As String 'pointer for cbEncodedBlob

'''''''''''''DEcoding variables
Private Const cbData = 4 'size of long
Private cbDecoded As Long
Private pbDecoded As String 'pointer for


''''''''''''Constants
'Encoding type
Private Const X509_ASN_ENCODING = &H1
Private Const PKCS_7_ASN_ENCODING = &H10000

'Message type
Private Const CMSG_DATA = 1
Private Const CMSG_CONTENT_PARAM = 2
Private Const CMSG_BARE_CONTENT_PARAM = 3

''''''''''APIS
Private Declare Function CryptMsgCalculateEncodedLength Lib "crypt32.dll" ( _
  ByVal dwMsgEncodingType As Long, _
  ByVal dwFlags As Long, _
  ByVal dwMsgType As Long, _
  pvMsgEncodeInfo As Long, _
  pszInnerContentObjID As String, _
  cbData As Long) As Long

Private Declare Function CryptMsgOpenToEncode Lib "crypt32.dll" ( _
  ByVal dwMsgEncodingType As Long, _
  ByVal dwFlags As Long, _
  ByVal dwMsgType As Long, _
  pvMsgEncodeInfo As Long, _
  pszInnerContentObjID As String, _
  pStreamInfo As Long) As Long

Private Declare Function CryptMsgOpenToDecode Lib "crypt32.dll" ( _
  ByVal dwMsgEncodingType As Long, _
  ByVal dwFlags As Long, _
  ByVal dwMsgType As Long, _
  hCryptProv As Long, _
  pRecipientInfo As Long, _
  pStreamInfo As Long) As Long
  
Private Declare Function CryptMsgUpdate Lib "crypt32.dll" ( _
  ByVal hCryptMsg As Long, _
  pbData As String, _
  ByVal cbData As Long, _
  ByVal fFinal As Boolean) As Long

Private Declare Function CryptMsgGetParam Lib "crypt32.dll" ( _
  ByVal hCryptMsg As Long, _
  ByVal dwParamType As Long, _
  ByVal dwIndex As Long, _
  pvData As String, _
  pcbData As Long) As Long

Private Declare Function CryptMsgClose Lib "crypt32.dll" ( _
  ByVal hCryptMsg As Long) As Long
  
  



Option Explicit


Public Function doCryptString(ByVal strv As String) As String
pbContent = strv  'get pointer for input string strv
cbContent = Len(strv) + 1 'get input string len and terminator

''Debug.print "The original message is " & strv

'
cbEncodedBlob = CryptMsgCalculateEncodedLength( _
             PKCS_7_ASN_ENCODING, _
             0, _
             CMSG_DATA, _
             ByVal 0, _
             ByVal 0, _
             ByVal cbContent)
             
If cbEncodedBlob <> 0 Then
    ''Debug.print "The length of the data for EncodedBlob has been calculated " & cbEncodedBlob
    'allocating memory for encrypted string
    pbEncodedBlob = Space$(cbEncodedBlob)
    
    hMsg = CryptMsgOpenToEncode( _
          PKCS_7_ASN_ENCODING, _
            0, _
           CMSG_DATA, _
          ByVal 0, _
          ByVal 0, _
          ByVal 0)
    
    If hMsg Then
        ''Debug.print "The message to be encoded has been op " & hMsg
                    
        If (CryptMsgUpdate( _
              hMsg, _
              StrPtr(pbContent), _
              cbContent, _
            True)) Then
            
            If CryptMsgGetParam(hMsg, 3, 0, ByVal pbEncodedBlob, cbEncodedBlob) Then
                ''Debug.print "The encoded data is " & pbEncodedBlob
                
                
                Dim i&, lim&, tmp1$, tmp$
                tmp1$ = Trim$(pbEncodedBlob)
                lim& = Len(tmp1$)
                tmp$ = ""
                For i = 1 To lim&
                    tmp$ = tmp$ & Hex(Asc(Mid(tmp1$, i, 1)))
                Next
                
                doCryptString = tmp$
                
            Else
                ''Debug.print "Error CryptMsgGetParam"
            End If
            

            ''Debug.print "Content has been added to the encoded message."
           Else
            ''Debug.print "Error CryptMsgUpdate"
        End If
            


        'Debug.print "Close CryptMsg " & CryptMsgClose(hMsg)
    Else
        ''Debug.print "Error CryptMsgOpenToEncode"

    End If
    

Else
    ''Debug.print "Error CryptMsgCalculateEncodedLength"
End If



End Function



Public Function doDeCryptString(ByVal strv As String) As String
pbContent = Space(1024)  'get pointer for input string strv
cbContent = 1024 'get input string len and terminator

''Debug.print "The original message is " & strv

'
hMsg = CryptMsgOpenToDecode( _
   PKCS_7_ASN_ENCODING, _
   0, _
   0, _
   ByVal 0, _
   ByVal 0, _
   ByVal 0)
    
    If hMsg Then
        ''Debug.print "The message to be encoded has been op " & hMsg
    
        If CryptMsgGetParam(hMsg, CMSG_DATA, 0, VarPtr(pbContent), cbContent) Then
                ''Debug.print "The encoded data is " & pbEncodedBlob
                'Debug.print "the encoded lenght is " & cbContent
            Else
                ''Debug.print "Error CryptMsgGetParam"
        End If
        
        If (CryptMsgUpdate( _
              hMsg, _
              pbContent, _
              cbContent, _
             True)) Then
            
            If CryptMsgGetParam(hMsg, CMSG_BARE_CONTENT_PARAM, 0, ByVal pbEncodedBlob, cbEncodedBlob) Then
                ''Debug.print "The encoded data is " & pbEncodedBlob
                doDeCryptString = Trim$(pbEncodedBlob)
            Else
                ''Debug.print "Error CryptMsgGetParam"
            End If
            

            ''Debug.print "Content has been added to the encoded message."
           Else
            ''Debug.print "Error CryptMsgUpdate"
        End If
            


        'Debug.print "Close CryptMsg " & CryptMsgClose(hMsg)
    Else
        ''Debug.print "Error CryptMsgOpenToEncode"

    End If
    




End Function

Public Function getLicen() As String
getLicen = doCryptString(Mid(getpckey(), 8))
'if getLicen ="
End Function


Public Function encodeString(ByVal strs$) As String
Dim i&, lim&, ch$
ch$ = ""
lim& = getSum(strs$)
For i& = 5 To Len(strs)
ch = ch & Format(Hex(CLng(CLng(i& ^ 2 + Asc(Mid(strs, i, 1))) + Format(lim + Asc(Mid(strs, i, 1)), "000"))), "00")
Next
encodeString = Mid(ch, 5, 10)
End Function

Private Function getSum(vl As String) As Long
Dim i, ret&
For i = 1 To Len(vl)
ret& = ret& + Asc(Mid(vl, i, 1))
Next
getSum = ret
End Function
'VENDIDO
'Santo Antão
'Cyber Interfone        : 15.000
'Cyber Maky             : 15.000
'Mindelo
'Cyber Gigatel          : 15.000
'Cyber Impena           : 15.000
'Cyber Jimms            : 15.000
'Cyber Furnalha*        : 15.000
'Cyber Compucv          : 15.000
'Cyber Ribeira Julião   : 15.000
'Praia
'Cyber Apostolica       : 15.000
'Cyber Church           : 15.000
'Cyber Other            : 15.000



