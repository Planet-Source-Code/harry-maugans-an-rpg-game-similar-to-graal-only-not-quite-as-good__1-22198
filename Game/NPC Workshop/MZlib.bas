Attribute VB_Name = "MZlib"
'Encryption function
Public Function Encrypt(ByVal Plain As String)
    Dim Letter As String
    For i = 1 To Len(Plain)
        Letter = Mid$(Plain, i, 1)
        Mid$(Plain, i, 1) = Chr(Asc(Letter) + 1)
    Next i
    Encrypt = Plain
End Function


'Here's the Decryption function:
Public Function Decrypt(ByVal Encrypted As String)
Dim Letter As String
    For i = 1 To Len(Encrypted)
        Letter = Mid$(Encrypted, i, 1)
        Mid$(Encrypted, i, 1) = Chr(Asc(Letter) - 1)
    Next i
    Decrypt = Encrypted
End Function

