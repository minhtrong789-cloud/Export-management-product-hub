Attribute VB_Name = "Extension_Cryptography"
Option Compare Database
Option Explicit
'@Folder "Extensions"
'====================================================================
' MODULE    : Extension_Cryptography
' PURPOSE   : Cryptographic SHA-256 hashing using Windows CNG (bcrypt.dll).
' PLATFORM  : 32-bit & 64-bit compatible (PtrSafe & LongPtr).
' DEPENDENCY: bcrypt.dll (Windows native system library).
'====================================================================

' 1. Open Cryptographic Algorithm Provider (e.g. SHA256, MD5) and retrieve handle
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" _
                                                            (ByRef phAlgorithm As LongPtr, _
                                                             ByVal pszAlgId As LongPtr, _
                                                             ByVal pszImplementation As LongPtr, _
                                                             ByVal dwFlags As Long) As Long

' 2. Compute one-shot hash from input buffer into output digest buffer
Private Declare PtrSafe Function BCryptHash Lib "bcrypt.dll" _
                                            (ByVal hAlgorithm As LongPtr, _
                                             ByVal pbSecret As LongPtr, _
                                             ByVal cbSecret As Long, _
                                             pbInput As Byte, _
                                             ByVal cbInput As Long, _
                                             pbOutput As Byte, _
                                             ByVal cbOutput As Long) As Long

' 3. Close Algorithm Provider handle and release system resources
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" _
                                                                (ByVal hAlgorithm As LongPtr, _
                                                                 ByVal dwFlags As Long) As Long

'--------------------------------------------------------------------
' FUNCTION  : SHA256_CNG
' PURPOSE   : Computes a 64-character SHA-256 hex digest of a string.
' INPUT     : inputString - Text to hash.
' OUTPUT    : 64-character lowercase hex string (or "" on failure).
'--------------------------------------------------------------------
Public Function SHA256_CNG(ByVal inputString As String) As String
    On Error GoTo ErrorHandler

    Dim dataBytes() As Byte
    Dim hAlgorithm As LongPtr
    Dim hashBytes(0 To 31) As Byte

    ' 1. Convert input text to byte array
    dataBytes = StrConv(inputString, vbFromUnicode)

    ' 2. Open SHA-256 algorithm provider
    If BCryptOpenAlgorithmProvider(hAlgorithm, StrPtr("SHA256"), 0, 0) <> 0 Then
        Err.Raise vbObjectError + 1, "SHA256_CNG", "Failed to open algorithm provider."
    End If

    ' 3. Compute 32-byte SHA-256 digest
    If BCryptHash(hAlgorithm, 0, 0, dataBytes(0), UBound(dataBytes) + 1, hashBytes(0), 32) <> 0 Then
        Err.Raise vbObjectError + 2, "SHA256_CNG", "Hash computation failed."
    End If

    ' 4. Convert 32 bytes to 64-character hex string
    Dim strHex As String, i As Long
    For i = 0 To 31
        strHex = strHex & LCase(Right("0" & Hex(hashBytes(i)), 2))
    Next i

    SHA256_CNG = strHex

Cleanup:
    If hAlgorithm <> 0 Then BCryptCloseAlgorithmProvider hAlgorithm, 0
    Exit Function

ErrorHandler:
    SHA256_CNG = ""
    Resume Cleanup
End Function
