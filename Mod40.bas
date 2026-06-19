'====================================================================
'  PESCA - Suivi Intervention   |   Module1   (FINAL - fixed order)
'
'  Builds an HTML table email (bordered, like the approved mockup).
'   - Intro changes by migration status (OK = succes / NOK = echec)
'   - REF SOLA is on its own row, under Sollicitation TSKC
'   - Uses each TSC's OWN Outlook signature (via .Display)
'
'  Delivery:
'   1) Classic Outlook (COM) available -> opens a real HTML draft with
'      the TSC's signature appended automatically.
'   2) Otherwise (New-Outlook-only) -> copies the HTML table to the
'      clipboard; the TSC pastes it (Ctrl+V) into a new mail.
'
'  NOTE: All Declare statements MUST sit at the very top of the module,
'  before any Sub/Function. That is a VBA rule.
'====================================================================
Option Explicit

'=========================  WINDOWS API  ============================
#If VBA7 Then
    Private Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As LongPtr, ByVal lpString2 As String) As LongPtr
#Else
    Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
#End If
Private Const GMEM_MOVEABLE As Long = &H2

'=========================  CONFIG  =================================
Private Const MAIL_TO As String = "nathalie.calvary@orange.com;irfaan.jeewooth@orange.com;tsc-delivery-gdo.pesca@orange.com"
Private Const MAIL_CC As String = "lajay.coonjul@orange.com;kailash.chooramun@orange.com;azhar.domah@orange.com;cinii.bihari@orange.com"
Private Const MGR_PWD As String = "Manager123"   ' change this

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
    If pwd = MGR_PWD Then
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
'  BUILD + OPEN THE EMAIL  (HTML table)
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
        ByVal pvCAT As String, ByVal clotureGamme As String)

    Dim subjectText As String, intro As String, html As String
    Dim isOK As Boolean

    isOK = (UCase$(Trim$(migrationStatus)) = "OK")
    subjectText = "[Advisory] [Migration du site " & siteName & "]"

    If isOK Then
        intro = "La migration du site <b>" & HtmlEsc(siteName) & _
                "</b> a &eacute;t&eacute; r&eacute;alis&eacute;e avec succ&egrave;s :"
    Else
        intro = "La migration du site <b>" & HtmlEsc(siteName) & _
                "</b> est en &eacute;chec :"
    End If

    html = "<div style='font-family:Calibri,Arial,sans-serif;font-size:11pt;'>" & _
           "<p>Bonjour,</p>" & _
           "<p>" & intro & "</p>" & _
           "<table border='1' cellspacing='0' cellpadding='5' " & _
           "style='border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;'>"

    html = html & TR("Date", migrationDate)
    html = html & TR("Nom du TSC", tscName)
    html = html & TR("Sitename", siteName)
    html = html & TR("Fiche", fiche)
    html = html & TR("Temps migration (mins)", migrationTime)
    html = html & TR("Migration (OK/NOK)", migrationStatus)
    html = html & TR("Sollicitation TSKC", BlankToNA(sollicitationTSKC))
    html = html & TR("Motif Sollicitation (et ref SOLA)", BlankToNA(motifSollicitation))
    html = html & TR("REF SOLA", BlankToNA(refSola))
    html = html & TR("Justification (si migration NOK)", BlankToNA(justificationNOK))
    html = html & TR("Supervision", BlankToNA(supervision))
    html = html & TR("Resupervision (si migration NOK)", BlankToNA(resupervision))
    html = html & TR("PV CAT", BlankToNA(pvCAT))
    html = html & TR("Cloture Gamme", BlankToNA(clotureGamme))

    html = html & "</table><p>Merci.</p></div>"

    If Not TryOutlookHTML(subjectText, html) Then
        Call CopyHtmlToClipboard(html)
        MsgBox "Classic Outlook not available on this PC." & vbCrLf & vbCrLf & _
               "The formatted table has been COPIED to your clipboard." & vbCrLf & _
               "1) Open a new email in Outlook" & vbCrLf & _
               "2) Press Ctrl+V in the body (the table keeps its formatting)" & vbCrLf & _
               "3) Add recipients and this subject:" & vbCrLf & vbCrLf & _
               subjectText, vbInformation, "Paste into a new mail"
    End If
End Sub

' Open a classic-Outlook HTML draft. Returns False if COM unavailable.
Private Function TryOutlookHTML(ByVal subj As String, ByVal htmlBody As String) As Boolean
    Dim OutApp As Object, OutMail As Object
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")
    If OutApp Is Nothing Then
        TryOutlookHTML = False
        Exit Function
    End If
    Set OutMail = OutApp.CreateItem(0)   ' 0 = olMailItem
    If OutMail Is Nothing Then
        TryOutlookHTML = False
        Exit Function
    End If
    With OutMail
        .To = MAIL_TO
        .CC = MAIL_CC
        .Subject = subj
        .GetInspector.Activate                 ' loads the user's signature
        .HTMLBody = htmlBody & .HTMLBody        ' table ABOVE existing signature
        .Display                               ' open for review; TSC sends manually
    End With
    On Error GoTo 0
    TryOutlookHTML = True
End Function

' One table row: <tr><td>label</td><td>value</td></tr>
Private Function TR(ByVal label As String, ByVal value As String) As String
    TR = "<tr><td style='background:#f2f2f2;'>" & HtmlEsc(label) & _
         "</td><td>" & HtmlEsc(value) & "</td></tr>"
End Function

Private Function BlankToNA(ByVal s As String) As String
    If Trim$(s) = "" Then BlankToNA = "N/A" Else BlankToNA = s
End Function

Private Function HtmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlEsc = s
End Function

' Build CF_HTML payload (with byte offsets) and put it on the clipboard.
Private Sub CopyHtmlToClipboard(ByVal htmlFragment As String)
    Dim full As String
    Dim hStart As Long, hEnd As Long, fStart As Long, fEnd As Long

    Const HDR As String = "Version:0.9" & vbCrLf & _
        "StartHTML:aaaaaaaaaa" & vbCrLf & "EndHTML:bbbbbbbbbb" & vbCrLf & _
        "StartFragment:cccccccccc" & vbCrLf & "EndFragment:dddddddddd" & vbCrLf

    full = HDR & "<html><body><!--StartFragment-->" & htmlFragment & "<!--EndFragment--></body></html>"

    hStart = InStr(full, "<html>") - 1
    hEnd = Len(full)
    fStart = InStr(full, "<!--StartFragment-->") - 1 + Len("<!--StartFragment-->")
    fEnd = InStr(full, "<!--EndFragment-->") - 1

    full = Replace(full, "aaaaaaaaaa", Format(hStart, "0000000000"))
    full = Replace(full, "bbbbbbbbbb", Format(hEnd, "0000000000"))
    full = Replace(full, "cccccccccc", Format(fStart, "0000000000"))
    full = Replace(full, "dddddddddd", Format(fEnd, "0000000000"))

    PutHtmlOnClipboard full
End Sub

Private Sub PutHtmlOnClipboard(ByVal payload As String)
    Dim cfHtml As Long, hMem As LongPtr, lpMem As LongPtr
    Dim n As Long

    cfHtml = RegisterClipboardFormat("HTML Format")
    n = LenB(StrConv(payload, vbFromUnicode)) + 1   ' +1 for null terminator

    hMem = GlobalAlloc(GMEM_MOVEABLE, n)
    lpMem = GlobalLock(hMem)
    lstrcpy lpMem, payload
    GlobalUnlock hMem

    If OpenClipboard(0) <> 0 Then
        EmptyClipboard
        SetClipboardData cfHtml, hMem
        CloseClipboard
    End If
End Sub
