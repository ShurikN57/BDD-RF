Attribute VB_Name = "BoutonAnnuler"
Option Explicit
' =============================================
'               Bouton Annuler
' =============================================
' Logique conservée :
' - SauvegarderEtatCandidat / SauvegarderEtat / ValiderEtatCandidatCommeUndo
' - AnnulerDerniereAction reste dédié au rejet de saisie / rollback interne
'
' Nouveau :
' - Gestion correcte des zones multiples (filtre actif / multi-blocs)
' - AnnulerUnique = annule 1 action utilisateur
' - RefaireUnique = remet 1 action utilisateur
' - Historique limité à 10 niveaux
' =============================================

' =============================================
'0.Constante
' =============================================
Private Const MAX_HISTORY As Long = 10
Private Const MAX_CELLS_UNDO As Long = 50000

' ===== Sauvegarde candidate =====
Public CandidateSheet As String
Public CandidateMainAddresses As Variant
Public CandidateMainDatas As Variant
Public CandidateLinkedAddresses As Variant
Public CandidateLinkedDatas As Variant

' ===== Sauvegarde officielle =====
Public BackupSheet As String
Public BackupMainAddresses As Variant
Public BackupMainDatas As Variant
Public BackupLinkedAddresses As Variant
Public BackupLinkedDatas As Variant

' ===== Historique utilisateur =====
Private Type UndoEntry
    sheetName As String
    mainAddresses As Variant
    mainDatas As Variant
    linkedAddresses As Variant
    linkedDatas As Variant
End Type

Private UndoStack(1 To MAX_HISTORY) As UndoEntry
Private UndoTop As Long

Private RedoStack(1 To MAX_HISTORY) As UndoEntry
Private RedoTop As Long

' =============================================
'1.TableauRenseigne
' =============================================
Private Function TableauRenseigne(ByVal v As Variant) As Boolean

    On Error GoTo Fin

    If IsEmpty(v) Then Exit Function

    If IsArray(v) Then
        TableauRenseigne = (UBound(v) >= LBound(v))
    Else
        TableauRenseigne = True
    End If

Fin:

End Function

' =============================================
' 2.CaptureEtat
' =============================================
Private Sub CaptureEtat(ByVal rng As Range, _
                        ByRef outSheet As String, _
                        ByRef outMainAddresses As Variant, _
                        ByRef outMainDatas As Variant, _
                        ByRef outLinkedAddresses As Variant, _
                        ByRef outLinkedDatas As Variant)

    Dim ws As Worksheet
    Dim area As Range
    Dim zoneConf As Range
    Dim i As Long
    Dim firstRow As Long
    Dim lastRow As Long

    outSheet = ""
    outMainAddresses = Empty
    outMainDatas = Empty
    outLinkedAddresses = Empty
    outLinkedDatas = Empty

    If rng Is Nothing Then Exit Sub

    If rng.Cells.CountLarge > MAX_CELLS_UNDO Then Exit Sub

    Set ws = rng.Worksheet
    outSheet = ws.Name

    ReDim outMainAddresses(1 To rng.Areas.Count)
    ReDim outMainDatas(1 To rng.Areas.Count)

    i = 0
    For Each area In rng.Areas
        i = i + 1
        outMainAddresses(i) = area.Address
        outMainDatas(i) = area.Value
    Next area

    Set zoneConf = Intersect(rng, ws.Columns(COL_CONF))
    If zoneConf Is Nothing Then Exit Sub

    ReDim outLinkedAddresses(1 To zoneConf.Areas.Count)
    ReDim outLinkedDatas(1 To zoneConf.Areas.Count)

    i = 0
    For Each area In zoneConf.Areas
        i = i + 1
        firstRow = area.Row
        lastRow = area.Row + area.Rows.Count - 1

        outLinkedAddresses(i) = ws.Range(COL_DATE & firstRow & ":" & COL_NOM & lastRow).Address
        outLinkedDatas(i) = ws.Range(COL_DATE & firstRow & ":" & COL_NOM & lastRow).Value
    Next area

End Sub

' =============================================
' 3.MakeEntryFromCurrentRanges
' =============================================
Private Sub MakeEntryFromCurrentRanges(ByVal sheetName As String, _
                                       ByVal mainAddresses As Variant, _
                                       ByVal linkedAddresses As Variant, _
                                       ByRef entry As UndoEntry)

    Dim ws As Worksheet
    Dim i As Long

    Call ClearEntry(entry)

    If sheetName = "" Then Exit Sub
    If Not TableauRenseigne(mainAddresses) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(sheetName)

    entry.sheetName = sheetName
    entry.mainAddresses = mainAddresses
    ReDim entry.mainDatas(LBound(mainAddresses) To UBound(mainAddresses))

    For i = LBound(mainAddresses) To UBound(mainAddresses)
        entry.mainDatas(i) = ws.Range(CStr(mainAddresses(i))).Value
    Next i

    If TableauRenseigne(linkedAddresses) Then
        entry.linkedAddresses = linkedAddresses
        ReDim entry.linkedDatas(LBound(linkedAddresses) To UBound(linkedAddresses))

        For i = LBound(linkedAddresses) To UBound(linkedAddresses)
            entry.linkedDatas(i) = ws.Range(CStr(linkedAddresses(i))).Value
        Next i
    End If

End Sub

' =============================================
' 4.RestoreEntry
' =============================================
Private Sub RestoreEntry(ByRef entry As UndoEntry)

    Dim ws As Worksheet
    Dim i As Long
    Dim adresseLignes As String

    If entry.sheetName = "" Then Exit Sub
    If Not TableauRenseigne(entry.mainAddresses) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(entry.sheetName)

    If TableauRenseigne(entry.linkedAddresses) Then
        For i = LBound(entry.linkedAddresses) To UBound(entry.linkedAddresses)
            ws.Range(CStr(entry.linkedAddresses(i))).Value = entry.linkedDatas(i)
        Next i
    End If

    For i = LBound(entry.mainAddresses) To UBound(entry.mainAddresses)
        ws.Range(CStr(entry.mainAddresses(i))).Value = entry.mainDatas(i)
    Next i

    If ws.Name = SHEET_MAIN Then
        adresseLignes = ConstruireAdresseLignesImpactees(ws, entry.mainAddresses, entry.linkedAddresses)

        If adresseLignes <> "" Then
            Application.Run "'" & ThisWorkbook.Name & "'!" & ws.CodeName & ".RafraichirCouleursConformiteSurLignes", adresseLignes
            Application.Run "'" & ThisWorkbook.Name & "'!" & ws.CodeName & ".RafraichirCouleursValidationSurLignes", adresseLignes
        End If
    End If

End Sub

' =============================================
' 5. PushUndoEntry
' =============================================
Private Sub PushUndoEntry(ByRef entry As UndoEntry)

    Dim i As Long

    If entry.sheetName = "" Then Exit Sub
    If Not TableauRenseigne(entry.mainAddresses) Then Exit Sub

    If UndoTop >= MAX_HISTORY Then
        For i = 1 To MAX_HISTORY - 1
            UndoStack(i) = UndoStack(i + 1)
        Next i
        UndoStack(MAX_HISTORY) = entry
        UndoTop = MAX_HISTORY
    Else
        UndoTop = UndoTop + 1
        UndoStack(UndoTop) = entry
    End If

End Sub

' =============================================
' 6. PushRedoEntry
' =============================================
Private Sub PushRedoEntry(ByRef entry As UndoEntry)

    Dim i As Long

    If entry.sheetName = "" Then Exit Sub
    If Not TableauRenseigne(entry.mainAddresses) Then Exit Sub

    If RedoTop >= MAX_HISTORY Then
        For i = 1 To MAX_HISTORY - 1
            RedoStack(i) = RedoStack(i + 1)
        Next i
        RedoStack(MAX_HISTORY) = entry
        RedoTop = MAX_HISTORY
    Else
        RedoTop = RedoTop + 1
        RedoStack(RedoTop) = entry
    End If

End Sub

' =============================================
' 7. ClearRedoStack
' =============================================
Private Sub ClearRedoStack()

    Dim i As Long

    For i = 1 To MAX_HISTORY
        Call ClearEntry(RedoStack(i))
    Next i
    RedoTop = 0

End Sub

' =============================================
' 7-bis. TableauxAdressesEgaux
' =============================================
Private Function TableauxAdressesEgaux(ByVal a As Variant, ByVal b As Variant) As Boolean

    Dim i As Long

    If Not TableauRenseigne(a) And Not TableauRenseigne(b) Then
        TableauxAdressesEgaux = True
        Exit Function
    End If

    If TableauRenseigne(a) <> TableauRenseigne(b) Then Exit Function
    If LBound(a) <> LBound(b) Then Exit Function
    If UBound(a) <> UBound(b) Then Exit Function

    For i = LBound(a) To UBound(a)
        If CStr(a(i)) <> CStr(b(i)) Then Exit Function
    Next i

    TableauxAdressesEgaux = True

End Function

' =============================================
' 7-ter. RetirerDernierUndoSiCorrespondAuBackup
' =============================================
Private Sub RetirerDernierUndoSiCorrespondAuBackup()

    If UndoTop = 0 Then Exit Sub
    If UndoStack(UndoTop).sheetName <> BackupSheet Then Exit Sub

    If Not TableauxAdressesEgaux(UndoStack(UndoTop).mainAddresses, BackupMainAddresses) Then Exit Sub
    If Not TableauxAdressesEgaux(UndoStack(UndoTop).linkedAddresses, BackupLinkedAddresses) Then Exit Sub

    Call ClearEntry(UndoStack(UndoTop))
    UndoTop = UndoTop - 1

End Sub

' =============================================
' 8. ClearEntry
' =============================================
Private Sub ClearEntry(ByRef entry As UndoEntry)

    entry.sheetName = ""
    entry.mainAddresses = Empty
    entry.mainDatas = Empty
    entry.linkedAddresses = Empty
    entry.linkedDatas = Empty

End Sub

' =============================================
' 9. LibelleEntry
' =============================================
Private Function LibelleEntry(ByRef entry As UndoEntry) As String

    If entry.sheetName = "" Then
        LibelleEntry = "(vide)"
    ElseIf TableauRenseigne(entry.mainAddresses) Then
        LibelleEntry = entry.sheetName & " | " & CStr(entry.mainAddresses(LBound(entry.mainAddresses)))
    Else
        LibelleEntry = entry.sheetName & " | (sans adresse)"
    End If

End Function

' =============================================
' 10. SauvegarderEtatCandidat
' =============================================
Public Sub SauvegarderEtatCandidat(ByVal rng As Range)

    CaptureEtat rng, _
        CandidateSheet, CandidateMainAddresses, CandidateMainDatas, _
        CandidateLinkedAddresses, CandidateLinkedDatas

End Sub

' =============================================
' 11. SauvegarderEtat
' =============================================
Public Sub SauvegarderEtat(ByVal rng As Range)

    Dim entry As UndoEntry

    CaptureEtat rng, _
        BackupSheet, BackupMainAddresses, BackupMainDatas, _
        BackupLinkedAddresses, BackupLinkedDatas

    If BackupSheet = "" Or Not TableauRenseigne(BackupMainAddresses) Then Exit Sub

    entry.sheetName = BackupSheet
    entry.mainAddresses = BackupMainAddresses
    entry.mainDatas = BackupMainDatas
    entry.linkedAddresses = BackupLinkedAddresses
    entry.linkedDatas = BackupLinkedDatas

    Call PushUndoEntry(entry)
    Call ClearRedoStack

End Sub

' =============================================
' 12. ValiderEtatCandidatCommeUndo
' =============================================
Public Sub ValiderEtatCandidatCommeUndo()

    Dim entry As UndoEntry

    If CandidateSheet = "" Or Not TableauRenseigne(CandidateMainAddresses) Then Exit Sub

    BackupSheet = CandidateSheet
    BackupMainAddresses = CandidateMainAddresses
    BackupMainDatas = CandidateMainDatas
    BackupLinkedAddresses = CandidateLinkedAddresses
    BackupLinkedDatas = CandidateLinkedDatas

    entry.sheetName = CandidateSheet
    entry.mainAddresses = CandidateMainAddresses
    entry.mainDatas = CandidateMainDatas
    entry.linkedAddresses = CandidateLinkedAddresses
    entry.linkedDatas = CandidateLinkedDatas

    Call PushUndoEntry(entry)
    Call ClearRedoStack

End Sub

' =============================================
' 13. AnnulerDerniereAction
' =============================================
Public Sub AnnulerDerniereAction()

    Dim ws As Worksheet
    Dim i As Long
    Dim adresseLignes As String

    On Error GoTo Fin

    If BackupSheet = "" Or Not TableauRenseigne(BackupMainAddresses) Then
        MsgBox "Aucune action à annuler.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(BackupSheet)

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If TableauRenseigne(BackupLinkedAddresses) Then
        For i = LBound(BackupLinkedAddresses) To UBound(BackupLinkedAddresses)
            ws.Range(CStr(BackupLinkedAddresses(i))).Value = BackupLinkedDatas(i)
        Next i
    End If

    For i = LBound(BackupMainAddresses) To UBound(BackupMainAddresses)
        ws.Range(CStr(BackupMainAddresses(i))).Value = BackupMainDatas(i)
    Next i

    If ws.Name = SHEET_MAIN Then
        adresseLignes = ConstruireAdresseLignesImpactees(ws, BackupMainAddresses, BackupLinkedAddresses)

        If adresseLignes <> "" Then
            Application.Run "'" & ThisWorkbook.Name & "'!" & ws.CodeName & ".RafraichirCouleursConformiteSurLignes", adresseLignes
            Application.Run "'" & ThisWorkbook.Name & "'!" & ws.CodeName & ".RafraichirCouleursValidationSurLignes", adresseLignes
        End If
    End If

    Call RetirerDernierUndoSiCorrespondAuBackup

Fin:
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'annulation : " & Err.description, vbExclamation
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

' =============================================
' 14. AnnulerUnique
' =============================================
' Bouton "Annuler"
Public Sub AnnulerUnique()

    Dim entryUndo As UndoEntry
    Dim entryRedo As UndoEntry

    On Error GoTo Fin

    If UndoTop = 0 Then
        MsgBox "Aucune action à annuler.", vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    entryUndo = UndoStack(UndoTop)

    Call MakeEntryFromCurrentRanges(entryUndo.sheetName, entryUndo.mainAddresses, entryUndo.linkedAddresses, entryRedo)

    Call RestoreEntry(entryUndo)

    Call PushRedoEntry(entryRedo)

    Call ClearEntry(UndoStack(UndoTop))
    UndoTop = UndoTop - 1

Fin:
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'annulation : " & Err.description, vbExclamation
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

' =============================================
' 15. RefaireUnique
' =============================================
Public Sub RefaireUnique()

    Dim entryRedo As UndoEntry
    Dim entryUndo As UndoEntry

    On Error GoTo Fin

    If RedoTop = 0 Then
        MsgBox "Aucune action à remettre.", vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    entryRedo = RedoStack(RedoTop)

    Call MakeEntryFromCurrentRanges(entryRedo.sheetName, entryRedo.mainAddresses, entryRedo.linkedAddresses, entryUndo)

    Call RestoreEntry(entryRedo)

    Call PushUndoEntry(entryUndo)

    Call ClearEntry(RedoStack(RedoTop))
    RedoTop = RedoTop - 1

Fin:
    If Err.Number <> 0 Then
        MsgBox "Erreur lors du rétablissement : " & Err.description, vbExclamation
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
' =============================================
' 16. DebugAfficherHistorique
' =============================================
Public Sub DebugAfficherHistorique()

    Dim msg As String
    Dim i As Long

    msg = "UNDO (" & UndoTop & "/" & MAX_HISTORY & ")" & vbCrLf
    If UndoTop = 0 Then
        msg = msg & "  (vide)"
    Else
        For i = 1 To UndoTop
            msg = msg & "  [" & i & "] " & LibelleEntry(UndoStack(i)) & vbCrLf
        Next i
    End If

    msg = msg & vbCrLf & "REDO (" & RedoTop & "/" & MAX_HISTORY & ")" & vbCrLf
    If RedoTop = 0 Then
        msg = msg & "  (vide)"
    Else
        For i = 1 To RedoTop
            msg = msg & "  [" & i & "] " & LibelleEntry(RedoStack(i)) & vbCrLf
        Next i
    End If

    MsgBox msg, vbInformation, "Historique"

End Sub

' =============================================
' 17. DebugReinitialiserHistorique
' =============================================
Public Sub DebugReinitialiserHistorique()

    Dim i As Long

    For i = 1 To MAX_HISTORY
        Call ClearEntry(UndoStack(i))
        Call ClearEntry(RedoStack(i))
    Next i

    UndoTop = 0
    RedoTop = 0

    BackupSheet = ""
    BackupMainAddresses = Empty
    BackupMainDatas = Empty
    BackupLinkedAddresses = Empty
    BackupLinkedDatas = Empty

    CandidateSheet = ""
    CandidateMainAddresses = Empty
    CandidateMainDatas = Empty
    CandidateLinkedAddresses = Empty
    CandidateLinkedDatas = Empty

    MsgBox "Historique réinitialisé.", vbInformation

End Sub

' =============================================
' 18. InvaliderEtatCandidat
' =============================================
Public Sub InvaliderEtatCandidat()

    CandidateSheet = ""
    CandidateMainAddresses = Empty
    CandidateMainDatas = Empty
    CandidateLinkedAddresses = Empty
    CandidateLinkedDatas = Empty

End Sub

' =============================================
' 19. ConstruireAdresseLignesImpactees
' =============================================
Private Function ConstruireAdresseLignesImpactees(ByVal ws As Worksheet, _
                                                  ByVal mainAddresses As Variant, _
                                                  ByVal linkedAddresses As Variant) As String

    Dim rngRows As Range
    Dim rngPart As Range
    Dim i As Long
    Dim firstRow As Long
    Dim lastRow As Long

    On Error GoTo Fin

    If TableauRenseigne(mainAddresses) Then
        For i = LBound(mainAddresses) To UBound(mainAddresses)
            firstRow = ws.Range(CStr(mainAddresses(i))).Row
            lastRow = firstRow + ws.Range(CStr(mainAddresses(i))).Rows.Count - 1

            Set rngPart = ws.Rows(CStr(firstRow) & ":" & CStr(lastRow))

            If rngRows Is Nothing Then
                Set rngRows = rngPart
            Else
                Set rngRows = Union(rngRows, rngPart)
            End If
        Next i
    End If

    If TableauRenseigne(linkedAddresses) Then
        For i = LBound(linkedAddresses) To UBound(linkedAddresses)
            firstRow = ws.Range(CStr(linkedAddresses(i))).Row
            lastRow = firstRow + ws.Range(CStr(linkedAddresses(i))).Rows.Count - 1

            Set rngPart = ws.Rows(CStr(firstRow) & ":" & CStr(lastRow))

            If rngRows Is Nothing Then
                Set rngRows = rngPart
            Else
                Set rngRows = Union(rngRows, rngPart)
            End If
        Next i
    End If

    If Not rngRows Is Nothing Then
        ConstruireAdresseLignesImpactees = rngRows.Address
    Else
        ConstruireAdresseLignesImpactees = ""
    End If

Fin:

End Function

