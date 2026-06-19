'====================================================================
'  PESCA - Suivi Intervention   |   Module1   (mailto / New Outlook)
'
'  Opens a draft automatically in the default mail client (New Outlook)
'  via mailto:. The TSC reviews and sends manually.
'   - Intro changes by migration status (OK = succes / NOK = echec)
'   - REF SOLA is on its OWN line
'   - Plain-text body (no table) so New Outlook keeps it
'====================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

'=========================  CONFIG  =================================
Private Const MAIL_TO As String = "irfaan.jeewooth@orange.com;tsc-delivery-gdo.pesca@orange.com"
Private Const MAIL_CC As String = "lajay.coonjul@orange.com;kailash.chooramun@orange.com;azhar.domah@orange.com;cinii.bihari@orange.com"
Private Const MGR_PWD_HASH As String = "a66058ae4f8c01e35eba6fa08d6cb06f9e661ec2426400b7f44ed7031e46d856"


'========================
' EMPLOYEE MODE  (runs on open)
'========================
Sub EmployeeMode()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Launch Form" Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    ThisWorkbook.Sheets("Launch Form").Activate
    Application.ScreenUpdating = True
End Sub

'========================
' MANAGER MODE
'========================
Sub ManagerMode()
    Dim pwd As String, ws As Worksheet
    pwd = InputBox("Enter manager password:", "Manager Access")
    If pwd = "" Then Exit Sub
    If pwd = MGR_PWD_HASH Then
        Application.ScreenUpdating = False
        For Each ws In ThisWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next ws
        Application.ScreenUpdating = True
        MsgBox "Manager access granted.", vbInformation
    Else
        MsgBox "Incorrect password.", vbExclamation
    End If
End Sub

Sub LockBackToEmployee()
    Call EmployeeMode
    MsgBox "Workbook locked to employee view.", vbInformation
End Sub

'========================
' OPEN THE FORM  (button target)
'========================
Sub LancerFormulaire()
    Suivi_Migration.Show
End Sub

'====================================================================
'  BUILD + OPEN THE EMAIL DRAFT VIA MAILTO
'  Argument order MUST match the form's btnSave_Click call:
'    date, tsc, site, fiche, duration, migration, tskc, refSola,
'    motif, justification, supervision, resupervision, pvcat, cloture
'====================================================================
Sub SendMigrationMailEx( _
        ByVal migrationDate As String, ByVal tscName As String, _
        ByVal siteName As String, ByVal fiche As String, _
        ByVal migrationTime As String, ByVal migrationStatus As String, _
        ByVal sollicitationTSKC As String, ByVal refSola As String, _
        ByVal motifSollicitation As String, ByVal justificationNOK As String, _
        ByVal supervision As String, ByVal resupervision As String, _
        ByVal PVcat As String, ByVal clotureGamme As String)

    Dim subjectText As String, intro As String, bodyText As String, url As String
    Dim isOK As Boolean

    isOK = (UCase$(Trim$(migrationStatus)) = "OK")
    subjectText = "[Advisory] [Migration du site " & siteName & "]"

    If isOK Then
        intro = "La migration du site " & siteName & " a ete realisee avec succes :"
    Else
        intro = "La migration du site " & siteName & " est en echec :"
    End If

    bodyText = "Bonjour," & vbCrLf & vbCrLf & _
        intro & vbCrLf & vbCrLf & _
        "Date : " & migrationDate & vbCrLf & _
        "Nom du TSC : " & tscName & vbCrLf & _
        "Sitename : " & siteName & vbCrLf & _
        "Fiche : " & fiche & vbCrLf & _
        "Temps migration (mins) : " & migrationTime & vbCrLf & _
        "Migration (OK/NOK) : " & migrationStatus & vbCrLf & _
        "Sollicitation TSKC : " & NA(sollicitationTSKC) & vbCrLf & _
        "Motif Sollicitation : " & NA(motifSollicitation) & vbCrLf & _
        "REF SOLA : " & NA(refSola) & vbCrLf & _
        "Justification (si migration NOK) : " & NA(justificationNOK) & vbCrLf & _
        "Supervision : " & NA(supervision) & vbCrLf & _
        "Resupervision (si migration NOK) : " & NA(resupervision) & vbCrLf & _
        "PV CAT : " & NA(PVcat) & vbCrLf & _
        "Cloture Gamme : " & NA(clotureGamme) & vbCrLf & vbCrLf & _
        "Merci."

    url = "mailto:" & UrlEncode(MAIL_TO) & _
          "?cc=" & UrlEncode(MAIL_CC) & _
          "&subject=" & UrlEncode(subjectText) & _
          "&body=" & UrlEncode(bodyText)

    ShellExecute 0, "open", url, vbNullString, vbNullString, 1
End Sub

Private Function NA(ByVal s As String) As String
    If Trim$(s) = "" Then NA = "N/A" Else NA = s
End Function

Private Function UrlEncode(ByVal s As String) As String
    Dim i As Long, c As String, code As Integer, out As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        code = Asc(c)
        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122
                out = out & c
            Case 45, 46, 95, 126
                out = out & c
            Case 13
                ' skip CR
            Case 10
                out = out & "%0D%0A"
            Case Else
                out = out & "%" & Right$("0" & Hex$(code), 2)
        End Select
    Next i
    UrlEncode = out
End Function

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
