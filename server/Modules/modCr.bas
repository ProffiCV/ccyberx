Attribute VB_Name = "modCr"
Option Explicit

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

'Debug.Print "The original message is " & strv

'
cbEncodedBlob = CryptMsgCalculateEncodedLength( _
             PKCS_7_ASN_ENCODING, _
             0, _
             CMSG_DATA, _
             ByVal 0, _
             ByVal 0, _
             ByVal cbContent)
             
If cbEncodedBlob <> 0 Then
    'Debug.Print "The length of the data for EncodedBlob has been calculated " & cbEncodedBlob
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
        'Debug.Print "The message to be encoded has been op " & hMsg
                    
        If (CryptMsgUpdate( _
              hMsg, _
              StrPtr(pbContent), _
              cbContent, _
            True)) Then
            
            If CryptMsgGetParam(hMsg, CMSG_BARE_CONTENT_PARAM, 0, ByVal pbEncodedBlob, cbEncodedBlob) Then
                'Debug.Print "The encoded data is " & pbEncodedBlob
                doCryptString = Trim$(pbEncodedBlob)
            Else
                'Debug.Print "Error CryptMsgGetParam"
            End If
            

            'Debug.Print "Content has been added to the encoded message."
           Else
            'Debug.Print "Error CryptMsgUpdate"
        End If
            


        Debug.Print "Close CryptMsg " & CryptMsgClose(hMsg)
    Else
        'Debug.Print "Error CryptMsgOpenToEncode"

    End If
    

Else
    'Debug.Print "Error CryptMsgCalculateEncodedLength"
End If



End Function



Public Function doDeCryptString(ByVal strv As String) As String
pbContent = strv  'get pointer for input string strv
cbContent = Len(strv) + 1 'get input string len and terminator

'Debug.Print "The original message is " & strv

'
hMsg = CryptMsgOpenToDecode( _
   PKCS_7_ASN_ENCODING, _
   0, _
   0, _
   ByVal 0, _
   ByVal 0, _
   ByVal 0)
    
    If hMsg Then
        'Debug.Print "The message to be encoded has been op " & hMsg
    
        If CryptMsgGetParam(hMsg, CMSG_CONTENT_PARAM, 0, ByVal 0&, cbContent) Then
                'Debug.Print "The encoded data is " & pbEncodedBlob
                Debug.Print "the encoded lenght is " & cbContent
            Else
                'Debug.Print "Error CryptMsgGetParam"
        End If
        
        If (CryptMsgUpdate( _
              hMsg, _
              StrPtr(pbContent), _
              cbContent, _
            True)) Then
            
            If CryptMsgGetParam(hMsg, CMSG_BARE_CONTENT_PARAM, 0, ByVal pbEncodedBlob, cbEncodedBlob) Then
                'Debug.Print "The encoded data is " & pbEncodedBlob
                doDeCryptString = Trim$(pbEncodedBlob)
            Else
                'Debug.Print "Error CryptMsgGetParam"
            End If
            

            'Debug.Print "Content has been added to the encoded message."
           Else
            'Debug.Print "Error CryptMsgUpdate"
        End If
            


        Debug.Print "Close CryptMsg " & CryptMsgClose(hMsg)
    Else
        'Debug.Print "Error CryptMsgOpenToEncode"

    End If
    




End Function




