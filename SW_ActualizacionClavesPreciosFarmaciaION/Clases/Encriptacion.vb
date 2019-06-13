Imports System.Text
Imports System.Security.Cryptography

Namespace CodeShare.Cryptography
Public Class Encriptacion

    Public Shared Function GenerateSHA256String(ByVal inputString) As String
      Dim sha256 As SHA256 = SHA256Managed.Create()
      Dim bytes As Byte() = Encoding.UTF8.GetBytes(inputString)
      Dim hash As Byte() = sha256.ComputeHash(bytes)
      Dim stringBuilder As New StringBuilder()

            For i As Integer = 0 To hash.Length - 1
                'El argumento X2 señala que la cadena se regrese en hexadecimal
                stringBuilder.Append(hash(i).ToString("X2"))
            Next

            Return stringBuilder.ToString()
    End Function

    Public Shared Function GenerateSHA512String(ByVal inputString) As String
      Dim sha512 As SHA512 = SHA512Managed.Create()
      Dim bytes As Byte() = Encoding.UTF8.GetBytes(inputString)
      Dim hash As Byte() = sha512.ComputeHash(bytes)
      Dim stringBuilder As New StringBuilder()

            For i As Integer = 0 To hash.Length - 1
                'El argumento X2 señala que la cadena se regrese en hexadecimal
                stringBuilder.Append(hash(i).ToString("X2"))
            Next

            Return stringBuilder.ToString()
   End Function
 End Class
End Namespace

