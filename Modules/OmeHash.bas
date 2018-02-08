Attribute VB_Name = "OmeHash"
'**************************************************************************
'   Module Name: Hash
'**************************************************************************
' Version: 1.0.0.0
' Created on: 17 April 2016
' Updated on: 08 Feb 2018
' Created by: Aeric Poon Yip Hoon
' Description: Create MD5 Hashing to store in database
'**************************************************************************
' Modified based on:
' http://khoiriyyah.blogspot.com/2012/06/vb6-hash-class-md5-sha-1-sha-256-sha.html
' Credit to: KARIM WAFI
' Posted Date: TUESDAY, OCTOBER 1, 2013
' Description: VB6 Hash Class - MD5 , SHA-1, SHA-256, SHA-384, SHA-512
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
                            (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
                             ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
                            (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" _
                            (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, _
                             ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
                            (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" _
                            (ByVal hHash As Long, pbData As Any, ByVal cbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
                            (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, ByRef pcbData As Long, _
                             ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL   As Long = 1
Private Const PROV_RSA_AES    As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Private Const HP_HASHVAL      As Long = 2
Private Const HP_HASHSIZE     As Long = 4

Private Const ALG_TYPE_ANY    As Long = 0
Private Const ALG_CLASS_HASH  As Long = 32768

'Private Const ALG_SID_MD2     As Long = 1
'Private Const ALG_SID_MD4     As Long = 2
Private Const ALG_SID_MD5     As Long = 3
Private Const ALG_SID_SHA     As Long = 4
Private Const ALG_SID_SHA_256 As Long = 12
'Private Const ALG_SID_SHA_384 As Long = 13
'Private Const ALG_SID_SHA_512 As Long = 14

'Private Const CALG_MD2        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
'Private Const CALG_MD4        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
Private Const CALG_MD5        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
Private Const CALG_SHA        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
Private Const CALG_SHA_256    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)
'Private Const CALG_SHA_384    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_384)
'Private Const CALG_SHA_512    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512)

' Create Hash
Private Function CreateHash(abytData() As Byte, ByVal lngAlgID As Long) As String
    Dim hProv As Long, hHash As Long
    Dim abytHash(0 To 63) As Byte
    Dim lngLength As Long
    Dim lngResult As Long
    Dim strHash As String
    Dim i As Long
    strHash = ""
    If CryptAcquireContext(hProv, vbNullString, vbNullString, _
                           IIf(lngAlgID >= CALG_SHA_256, PROV_RSA_AES, PROV_RSA_FULL), _
                           CRYPT_VERIFYCONTEXT) <> 0& Then
        If CryptCreateHash(hProv, lngAlgID, 0&, 0&, hHash) <> 0& Then
            lngLength = UBound(abytData()) - LBound(abytData()) + 1
            If lngLength > 0 Then lngResult = CryptHashData(hHash, abytData(LBound(abytData())), lngLength, 0&) _
                             Else lngResult = CryptHashData(hHash, ByVal 0&, 0&, 0&)
            If lngResult <> 0& Then
                lngLength = UBound(abytHash()) - LBound(abytHash()) + 1
                If CryptGetHashParam(hHash, HP_HASHVAL, abytHash(LBound(abytHash())), lngLength, 0&) <> 0& Then
                    For i = 0 To lngLength - 1
                        strHash = strHash & Right$("0" & Hex$(abytHash(LBound(abytHash()) + i)), 2)
                    Next
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0&
    End If
    CreateHash = LCase$(strHash)
End Function

' Create Hash From String(Shift_JIS)
Private Function CreateHashString(ByVal strData As String, ByVal lngAlgID As Long) As String
    CreateHashString = CreateHash(StrConv(strData, vbFromUnicode), lngAlgID)
End Function

' MD5
Public Function MD5(ByVal strData As String) As String
    MD5 = CreateHashString(strData, CALG_MD5)
End Function

' SHA-1
Public Function SHA1(ByVal strData As String) As String
    SHA1 = CreateHashString(strData, CALG_SHA)
End Function



