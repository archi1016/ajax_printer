Attribute VB_Name = "func_crypt"
Option Explicit

Dim hProv As Long

Public Function HashStringToMD5Bin(ByVal srcStr As String, OutBinary() As Byte) As Boolean
    Dim hHash As Long
    Dim bStr() As Byte
    Dim bStrLen As Long
    Dim OutBinarySize As Long
    
    HashStringToMD5Bin = False
    If HashStartup() Then
        If HashObjectCreate(CALG_MD5, hHash) <> 0 Then
            bStr = StrConv(srcStr, vbFromUnicode)
            bStrLen = UBound(bStr) + 1
            If bStrLen > 0 Then
                If CryptHashData(hHash, VarPtr(bStr(0)), bStrLen, 0) <> 0 Then
                    OutBinarySize = 16
                    If CryptGetHashParam(hHash, HP_HASHVAL, VarPtr(OutBinary(0)), OutBinarySize, 0) <> 0 Then
                        HashStringToMD5Bin = True
                    End If
                End If
            End If
            Call HashObjectDestroy(hHash)
        End If
        Call HashCleanup
    End If
    
    Erase bStr
End Function

Public Function HashStringToSHAbin(ByVal srcStr As String, OutBinary() As Byte) As Boolean
    Dim hHash As Long
    Dim bStr() As Byte
    Dim bStrLen As Long
    Dim OutBinarySize As Long
    
    HashStringToSHAbin = False
    If HashStartup() Then
        If HashObjectCreate(CALG_SHA, hHash) <> 0 Then
            bStr = StrConv(srcStr, vbFromUnicode)
            bStrLen = UBound(bStr) + 1
            If bStrLen > 0 Then
                If CryptHashData(hHash, VarPtr(bStr(0)), bStrLen, 0) <> 0 Then
                    OutBinarySize = 20
                    If CryptGetHashParam(hHash, HP_HASHVAL, VarPtr(OutBinary(0)), OutBinarySize, 0) <> 0 Then
                        HashStringToSHAbin = True
                    End If
                End If
            End If
            Call HashObjectDestroy(hHash)
        End If
        Call HashCleanup
    End If
    
    Erase bStr
End Function

Public Function HashStringToMD5(ByVal srcStr As String, OuStr As String) As Boolean
    Dim hHash As Long
    Dim bStr() As Byte
    Dim bStrLen As Long
    Dim hStr(15) As Byte
    Dim hStrLen As Long
    
    HashStringToMD5 = False
    If HashStartup() Then
        If HashObjectCreate(CALG_MD5, hHash) <> 0 Then
            bStr = StrConv(srcStr, vbFromUnicode)
            bStrLen = UBound(bStr) + 1
            If bStrLen > 0 Then
                If CryptHashData(hHash, VarPtr(bStr(0)), bStrLen, 0) <> 0 Then
                    hStrLen = 16
                    If CryptGetHashParam(hHash, HP_HASHVAL, VarPtr(hStr(0)), hStrLen, 0) <> 0 Then
                        OuStr = vbNullString
                        For bStrLen = 0 To 15
                            OuStr = OuStr + Right$("0" + Hex$(hStr(bStrLen)), 2)
                        Next
                        OuStr = LCase$(OuStr)
                        HashStringToMD5 = True
                    End If
                End If
            End If
            Call HashObjectDestroy(hHash)
        End If
        Call HashCleanup
    End If
    
    Erase bStr
End Function

Public Function HashStringToSHA(ByVal srcStr As String, OuStr As String) As Boolean
    Dim hHash As Long
    Dim bStr() As Byte
    Dim bStrLen As Long
    Dim hStr(19) As Byte
    Dim hStrLen As Long
    
    HashStringToSHA = False
    If HashStartup() Then
        If HashObjectCreate(CALG_SHA, hHash) <> 0 Then
            bStr = StrConv(srcStr, vbFromUnicode)
            bStrLen = UBound(bStr) + 1
            If bStrLen > 0 Then
                If CryptHashData(hHash, VarPtr(bStr(0)), bStrLen, 0) <> 0 Then
                    hStrLen = 20
                    If CryptGetHashParam(hHash, HP_HASHVAL, VarPtr(hStr(0)), hStrLen, 0) <> 0 Then
                        OuStr = vbNullString
                        For bStrLen = 0 To 19
                            OuStr = OuStr + Right$("0" + Hex$(hStr(bStrLen)), 2)
                        Next
                        OuStr = LCase$(OuStr)
                        HashStringToSHA = True
                    End If
                End If
            End If
            Call HashObjectDestroy(hHash)
        End If
        Call HashCleanup
    End If
    
    Erase bStr
End Function

Public Function HashStartup() As Boolean
    HashStartup = (CryptAcquireContextW(hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0)
End Function

Public Sub HashCleanup()
    CryptReleaseContext hProv, 0
End Sub

Public Function HashObjectCreate(ByVal Al As Long, hHash As Long) As Boolean
    HashObjectCreate = (CryptCreateHash(hProv, Al, 0, 0, hHash) <> 0)
End Function

Public Sub HashObjectDestroy(ByVal hHash As Long)
    CryptDestroyHash hHash
End Sub

Public Function DecodeBase64String(ByVal srcStr As String, RetBuf() As Byte, RetSize As Long) As Boolean
    DecodeBase64String = False
    
    RetSize = 0
    If srcStr <> "" Then
        If CryptStringToBinaryW(StrPtr(srcStr), 0, CRYPT_STRING_BASE64, 0, RetSize, 0, 0) <> 0 Then
            ReDim RetBuf(RetSize - 1)
            If CryptStringToBinaryW(StrPtr(srcStr), 0, CRYPT_STRING_BASE64, VarPtr(RetBuf(0)), RetSize, 0, 0) <> 0 Then
                DecodeBase64String = True
            End If
        End If
    End If
End Function
