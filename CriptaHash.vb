Imports System.Text
Imports System.Security.Cryptography

Public Class CriptaHash
    Public Function HashPassword(password As String) As String
        Using sha256 As SHA256 = sha256.Create()
            Dim bytes As Byte() = Encoding.UTF8.GetBytes(password)
            Dim hash As Byte() = sha256.ComputeHash(bytes)
            Return Convert.ToBase64String(hash)
        End Using
    End Function
    Public Function VerificaPassword(password As String, hashedPassword As String) As Boolean
        Dim hashedInput As String = HashPassword(password)
        Return hashedInput.Equals(hashedPassword)
    End Function

End Class
