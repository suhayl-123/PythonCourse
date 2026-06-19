Function SHA256Hash(ByVal text As String) As String
    Dim K() As Long, H() As Long, msg() As Byte
    Dim i As Long, j As Long, n As Long, nBlocks As Long
    Dim w(0 To 63) As Long
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As Long, h0 As Long
    Dim t1 As Long, t2 As Long, s0 As Long, s1 As Long, ch As Long, maj As Long

    ReDim K(0 To 63)
    Dim kHex As Variant
    kHex = Array(&H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
        &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
        &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
        &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
        &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
        &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
        &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
        &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2)
    For i = 0 To 63: K(i) = kHex(i): Next i

    ReDim H(0 To 7)
    H(0) = &H6A09E667: H(1) = &HBB67AE85: H(2) = &H3C6EF372: H(3) = &HA54FF53A
    H(4) = &H510E527F: H(5) = &H9B05688C: H(6) = &H1F83D9AB: H(7) = &H5BE0CD19

    Dim raw() As Byte, ln As Long
    raw = StrConv(text, vbFromUnicode)
    On Error Resume Next
    ln = UBound(raw) - LBound(raw) + 1
    If Err.Number <> 0 Then ln = 0
    On Error GoTo 0

    Dim bitLen As Double
    bitLen = ln * 8
    Dim total As Long
    total = ln + 1
    Do While (total Mod 64) <> 56
        total = total + 1
    Loop
    total = total + 8
    ReDim msg(0 To total - 1)
    For i = 0 To ln - 1: msg(i) = raw(i): Next i
    msg(ln) = &H80
    For i = 0 To 7
        msg(total - 1 - i) = Int((bitLen / (2 ^ (8 * i)))) And &HFF
    Next i

    nBlocks = total \ 64
    For n = 0 To nBlocks - 1
        For i = 0 To 15
            w(i) = (CLng(msg(n * 64 + i * 4)) And &HFF)
            w(i) = LShift(w(i), 8) Or (CLng(msg(n * 64 + i * 4 + 1)) And &HFF)
            w(i) = LShift(w(i), 8) Or (CLng(msg(n * 64 + i * 4 + 2)) And &HFF)
            w(i) = LShift(w(i), 8) Or (CLng(msg(n * 64 + i * 4 + 3)) And &HFF)
        Next i
        For i = 16 To 63
            s0 = RotR(w(i - 15), 7) Xor RotR(w(i - 15), 18) Xor RShift(w(i - 15), 3)
            s1 = RotR(w(i - 2), 17) Xor RotR(w(i - 2), 19) Xor RShift(w(i - 2), 10)
            w(i) = AddU(AddU(AddU(w(i - 16), s0), w(i - 7)), s1)
        Next i

        a = H(0): b = H(1): c = H(2): d = H(3)
        e = H(4): f = H(5): g = H(6): h0 = H(7)

        For i = 0 To 63
            s1 = RotR(e, 6) Xor RotR(e, 11) Xor RotR(e, 25)
            ch = (e And f) Xor ((Not e) And g)
            t1 = AddU(AddU(AddU(AddU(h0, s1), ch), K(i)), w(i))
            s0 = RotR(a, 2) Xor RotR(a, 13) Xor RotR(a, 22)
            maj = (a And b) Xor (a And c) Xor (b And c)
            t2 = AddU(s0, maj)
            h0 = g: g = f: f = e
            e = AddU(d, t1)
            d = c: c = b: b = a
            a = AddU(t1, t2)
        Next i

        H(0) = AddU(H(0), a): H(1) = AddU(H(1), b): H(2) = AddU(H(2), c): H(3) = AddU(H(3), d)
        H(4) = AddU(H(4), e): H(5) = AddU(H(5), f): H(6) = AddU(H(6), g): H(7) = AddU(H(7), h0)
    Next n

    Dim res As String
    For i = 0 To 7
        res = res & Right$("00000000" & LCase$(Hex$(H(i))), 8)
    Next i
    SHA256Hash = res
End Function

Private Function LShift(ByVal v As Long, ByVal n As Long) As Long
    Dim i As Long
    For i = 1 To n
        If (v And &H40000000) <> 0 Then
            v = (v And &H3FFFFFFF) * 2 Or &H80000000
        Else
            v = (v And &H3FFFFFFF) * 2
        End If
    Next i
    LShift = v
End Function

Private Function RShift(ByVal v As Long, ByVal n As Long) As Long
    Dim i As Long
    For i = 1 To n
        If (v And &H80000000) <> 0 Then
            v = (v And &H7FFFFFFF) \ 2 Or &H40000000
        Else
            v = v \ 2
        End If
    Next i
    RShift = v
End Function

Private Function RotR(ByVal v As Long, ByVal n As Long) As Long
    RotR = RShift(v, n) Or LShift(v, 32 - n)
End Function

Private Function AddU(ByVal x As Long, ByVal y As Long) As Long
    Dim lo As Long
    lo = (x And &HFFFF&) + (y And &HFFFF&)
    Dim hi As Long
    hi = ((x \ &H10000) And &HFFFF&) + ((y \ &H10000) And &HFFFF&) + (lo \ &H10000)
    AddU = (lo And &HFFFF&) Or LShift(hi And &HFFFF&, 16)
End Function
