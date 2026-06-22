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
        ByVal pvCAT As String, ByVal clotureGamme As String)

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
        "PV CAT : " & NA(pvCAT) & vbCrLf & _
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
