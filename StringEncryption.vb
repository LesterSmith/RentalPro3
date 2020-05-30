Imports System.Text
Imports System.Security.Cryptography
Imports System.IO

Public Class StringEncryption
    Private Const Key As String = "SecureAutomatedEndToEndP"
    Private Const IV As String = "SAEPNO01"

    Public Shared Function EncryptString(ByVal Value As String) As String
        Dim ky As Byte() = (New ASCIIEncoding).GetBytes(Key)
        Dim InitVect As Byte() = (New ASCIIEncoding).GetBytes(IV)
        Dim InputBytes As Byte() = (New ASCIIEncoding).GetBytes(Value)
        Dim Des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
        Dim crypTrans As ICryptoTransform = Des.CreateEncryptor(ky, InitVect)
        Dim encryptStream As MemoryStream = New MemoryStream
        Dim cryptStream As CryptoStream = New CryptoStream(encryptStream, crypTrans, CryptoStreamMode.Write)

        If Value = String.Empty Then Return String.Empty

        cryptStream.Write(InputBytes, 0, InputBytes.Length)
        cryptStream.FlushFinalBlock()
        encryptStream.Position = 0

        Dim result As Byte() = encryptStream.ToArray

        encryptStream.Read(result, 0, Convert.ToInt32(encryptStream.Length))
        cryptStream.Close()
        Return Convert.ToBase64String(result)
    End Function

    Public Shared Function DecryptString(ByVal Value As String) As String
        Dim ky As Byte() = (New ASCIIEncoding).GetBytes(Key)
        Dim InitVect As Byte() = (New ASCIIEncoding).GetBytes(IV)
        Dim InputBytes As Byte() = Convert.FromBase64String(Value)
        Dim Des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
        Dim crypTrans As ICryptoTransform = Des.CreateDecryptor(ky, InitVect)
        Dim decryptStream As MemoryStream = New MemoryStream
        Dim cryptStream As CryptoStream = New CryptoStream(decryptStream, crypTrans, CryptoStreamMode.Write)

        If Value = String.Empty Then Return String.Empty

        cryptStream.Write(InputBytes, 0, InputBytes.Length)
        cryptStream.FlushFinalBlock()
        decryptStream.Position = 0

        Dim result(Convert.ToInt32(decryptStream.Length - 1)) As Byte

        decryptStream.Read(result, 0, Convert.ToInt32(decryptStream.Length))
        cryptStream.Close()
        Return (New ASCIIEncoding).GetString(result)
    End Function
End Class
