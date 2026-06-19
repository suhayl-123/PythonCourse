Function SHA256Hash(ByVal text As String) As String
    Dim enc As Object, utf8 As Object
    Dim bytes() As Byte, hashBytes() As Byte
    Dim i As Long, result As String
    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.SHA256Managed")
    bytes = utf8.GetBytes_4(text)
    hashBytes = enc.ComputeHash_2(bytes)
    For i = LBound(hashBytes) To UBound(hashBytes)
        result = result & Right("0" & Hex(hashBytes(i)), 2)
    Next i
    SHA256Hash = LCase(result)
End Function
