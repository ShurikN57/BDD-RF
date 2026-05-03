Attribute VB_Name = "AuditPerf"
Option Explicit

Private Const SHEET_AUDIT As String = "AUDIT_PERF_"

' =============================================
' 1. AuditPerf
' =============================================
Public Sub AuditPerf()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsAudit As Worksheet
    Dim nextRow As Long
    Dim structureProtegee As Boolean

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevDisplayAlerts As Boolean

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    prevDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    structureProtegee = False
    On Error Resume Next
    structureProtegee = wb.ProtectStructure
    If structureProtegee Then wb.Unprotect Password:=MDP_DEV
    On Error GoTo ErrHandler

    Set wsAudit = ObtenirFeuilleAudit(wb)

    NettoyerFeuilleAudit wsAudit
    PreparerFeuilleAudit wsAudit

    nextRow = 2

    For Each ws In wb.Worksheets
        If ws.Name <> SHEET_AUDIT Then
            AnalyserFeuilleRapide ws, wsAudit, nextRow
            nextRow = nextRow + 1
        End If
    Next ws

    EcrireSyntheseClasseur wb, wsAudit
    MettreEnFormeAudit wsAudit

SortiePropre:
    On Error Resume Next
    If structureProtegee Then wb.Protect Password:=MDP_DEV, Structure:=True
    Application.DisplayAlerts = prevDisplayAlerts
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de l'audit rapide : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

' =============================================
' 2. AnalyserFeuilleRapide
' =============================================
Private Sub AnalyserFeuilleRapide(ByVal ws As Worksheet, ByVal wsAudit As Worksheet, ByVal outRow As Long)

    Dim ur As Range
    Dim nbRows As Long
    Dim nbCols As Long
    Dim nbCells As Double
    Dim usedAddr As String

    Dim nbFormules As Double
    Dim nbHyperlinks As Long
    Dim nbShapes As Long
    Dim nbOLE As Long
    Dim nbComments As Long
    Dim nbTables As Long
    Dim nbPivot As Long
    Dim nbNamesLocal As Long
    Dim nbMFC_Areas As Long
    Dim nbValidationAreas As Long
    Dim score As Long
    Dim diagnostic As String

    Set ur = SafeUsedRange(ws)

    If ur Is Nothing Then
        usedAddr = ""
        nbRows = 0
        nbCols = 0
        nbCells = 0
    Else
        usedAddr = ur.Address(False, False)
        nbRows = ur.Rows.Count
        nbCols = ur.Columns.Count
        nbCells = CDbl(nbRows) * CDbl(nbCols)
    End If

    nbFormules = CountFormulasFast(ws)
    nbHyperlinks = ws.Hyperlinks.Count
    nbShapes = ws.Shapes.Count
    nbOLE = CountOLESafe(ws)
    nbComments = CountCommentsFast(ws)
    nbTables = ws.ListObjects.Count
    nbPivot = ws.PivotTables.Count
    nbNamesLocal = ws.Names.Count
    nbMFC_Areas = CountFormatConditionAreasFast(ws)
    nbValidationAreas = CountValidationAreasFast(ws)

    score = EvaluerScoreRapide(nbCells, nbFormules, nbShapes, nbHyperlinks, nbOLE, nbComments, nbTables, nbPivot, nbMFC_Areas, nbValidationAreas)
    diagnostic = DiagnosticRapide(nbCells, nbFormules, nbShapes, nbHyperlinks, nbOLE, nbComments, nbTables, nbPivot, nbMFC_Areas, nbValidationAreas)

    wsAudit.Cells(outRow, 1).Value = ws.Name
    wsAudit.Cells(outRow, 2).Value = usedAddr
    wsAudit.Cells(outRow, 3).Value = nbRows
    wsAudit.Cells(outRow, 4).Value = nbCols
    wsAudit.Cells(outRow, 5).Value = nbCells
    wsAudit.Cells(outRow, 6).Value = nbFormules
    wsAudit.Cells(outRow, 7).Value = nbMFC_Areas
    wsAudit.Cells(outRow, 8).Value = nbValidationAreas
    wsAudit.Cells(outRow, 9).Value = nbShapes
    wsAudit.Cells(outRow, 10).Value = nbHyperlinks
    wsAudit.Cells(outRow, 11).Value = nbOLE
    wsAudit.Cells(outRow, 12).Value = nbComments
    wsAudit.Cells(outRow, 13).Value = nbTables
    wsAudit.Cells(outRow, 14).Value = nbPivot
    wsAudit.Cells(outRow, 15).Value = nbNamesLocal
    wsAudit.Cells(outRow, 16).Value = score
    wsAudit.Cells(outRow, 17).Value = diagnostic

End Sub

' =============================================
' 3. PreparerFeuilleAudit
' =============================================
Private Sub PreparerFeuilleAudit(ByVal wsAudit As Worksheet)

    ' Ne touche jamais à la ligne 1 :
    ' - pas de réécriture des titres
    ' - pas de modification des retours à la ligne
    ' - pas de modification de la largeur des colonnes
    '
    ' Les titres doivent être préparés une seule fois manuellement
    ' dans la feuille AUDIT_PERF_.

End Sub

' =============================================
' 4. ObtenirFeuilleAudit
' =============================================
Private Function ObtenirFeuilleAudit(ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_AUDIT)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_AUDIT
    End If

    Set ObtenirFeuilleAudit = ws

End Function

' =============================================
' 4-bis. NettoyerFeuilleAudit
' =============================================
Private Sub NettoyerFeuilleAudit(ByVal wsAudit As Worksheet)

    Dim lastRow As Long

    On Error GoTo Fin

    lastRow = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row

    ' On vide uniquement les anciennes données, pas la ligne 1.
    ' La mise en forme, les largeurs et les retours à la ligne restent inchangés.
    If lastRow >= 2 Then
        wsAudit.Range("A2:Q" & lastRow).ClearContents
    End If

    ' Synthèse classeur : on vide seulement les valeurs S2:T4.
    ' S1 reste intact.
    wsAudit.Range("S2:T4").ClearContents

Fin:
    Err.Clear

End Sub

' =============================================
' 5. SafeUsedRange
' =============================================
Private Function SafeUsedRange(ByVal ws As Worksheet) As Range

    On Error Resume Next
    Set SafeUsedRange = ws.UsedRange
    On Error GoTo 0

End Function

' =============================================
' 6. CountFormulasFast
' =============================================
Private Function CountFormulasFast(ByVal ws As Worksheet) As Double

    Dim rng As Range

    On Error Resume Next
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If rng Is Nothing Then
        CountFormulasFast = 0
    Else
        CountFormulasFast = rng.CountLarge
    End If

End Function

' =============================================
' 7. CountValidationAreasFast
' =============================================
Private Function CountValidationAreasFast(ByVal ws As Worksheet) As Long

    Dim rng As Range

    On Error Resume Next
    Set rng = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    On Error GoTo 0

    If rng Is Nothing Then
        CountValidationAreasFast = 0
    Else
        CountValidationAreasFast = rng.Areas.Count
    End If

End Function

' =============================================
' 8. CountCommentsFast
' =============================================
Private Function CountCommentsFast(ByVal ws As Worksheet) As Long

    Dim n As Long

    On Error Resume Next
    n = ws.Comments.Count
    If Err.Number <> 0 Then
        Err.Clear
        n = 0
    End If
    On Error GoTo 0

    CountCommentsFast = n

End Function

' =============================================
' 9. CountOLESafe
' =============================================
Private Function CountOLESafe(ByVal ws As Worksheet) As Long

    On Error Resume Next
    CountOLESafe = ws.OLEObjects.Count
    If Err.Number <> 0 Then
        Err.Clear
        CountOLESafe = 0
    End If
    On Error GoTo 0

End Function

' =============================================
' 10. CountFormatConditionAreasFast
' =============================================
Private Function CountFormatConditionAreasFast(ByVal ws As Worksheet) As Long

    Dim fc As FormatCondition
    Dim total As Long

    On Error GoTo Fin

    For Each fc In ws.Cells.FormatConditions
        total = total + 1
    Next fc

Fin:
    CountFormatConditionAreasFast = total

End Function

' =============================================
' 11. EvaluerScoreRapide
' =============================================
Private Function EvaluerScoreRapide(ByVal nbCells As Double, _
                                    ByVal nbFormules As Double, _
                                    ByVal nbShapes As Long, _
                                    ByVal nbHyperlinks As Long, _
                                    ByVal nbOLE As Long, _
                                    ByVal nbComments As Long, _
                                    ByVal nbTables As Long, _
                                    ByVal nbPivot As Long, _
                                    ByVal nbMFC_Areas As Long, _
                                    ByVal nbValidationAreas As Long) As Long

    Dim score As Long

    If nbCells > 100000 Then score = score + 2
    If nbCells > 500000 Then score = score + 3
    If nbCells > 2000000 Then score = score + 4

    If nbFormules > 1000 Then score = score + 2
    If nbFormules > 10000 Then score = score + 3

    If nbMFC_Areas > 20 Then score = score + 2
    If nbMFC_Areas > 100 Then score = score + 3

    If nbValidationAreas > 20 Then score = score + 1
    If nbValidationAreas > 100 Then score = score + 2

    If nbShapes > 20 Then score = score + 1
    If nbShapes > 100 Then score = score + 2

    If nbHyperlinks > 500 Then score = score + 1
    If nbOLE > 0 Then score = score + 2
    If nbComments > 100 Then score = score + 1
    If nbTables > 5 Then score = score + 1
    If nbPivot > 0 Then score = score + 1

    EvaluerScoreRapide = score

End Function

' =============================================
' 12. DiagnosticRapide
' =============================================
Private Function DiagnosticRapide(ByVal nbCells As Double, _
                                  ByVal nbFormules As Double, _
                                  ByVal nbShapes As Long, _
                                  ByVal nbHyperlinks As Long, _
                                  ByVal nbOLE As Long, _
                                  ByVal nbComments As Long, _
                                  ByVal nbTables As Long, _
                                  ByVal nbPivot As Long, _
                                  ByVal nbMFC_Areas As Long, _
                                  ByVal nbValidationAreas As Long) As String

    Dim msg As String

    If nbCells > 500000 Then msg = msg & "UsedRange large; "
    If nbCells > 2000000 Then msg = msg & "UsedRange très large; "
    If nbFormules > 10000 Then msg = msg & "beaucoup de formules; "
    If nbMFC_Areas > 20 Then msg = msg & "plusieurs MFC; "
    If nbValidationAreas > 20 Then msg = msg & "plusieurs validations; "
    If nbShapes > 100 Then msg = msg & "beaucoup de formes; "
    If nbHyperlinks > 1000 Then msg = msg & "beaucoup d'hyperliens; "
    If nbOLE > 0 Then msg = msg & "objets OLE présents; "
    If nbComments > 100 Then msg = msg & "beaucoup de commentaires; "
    If nbTables > 5 Then msg = msg & "plusieurs tableaux; "
    If nbPivot > 0 Then msg = msg & "TCD présents; "

    If msg = "" Then
        DiagnosticRapide = "RAS majeur"
    Else
        DiagnosticRapide = Left$(msg, Len(msg) - 2)
    End If

End Function

' =============================================
' 13. EcrireSyntheseClasseur
' =============================================
Private Sub EcrireSyntheseClasseur(ByVal wb As Workbook, ByVal wsAudit As Worksheet)

    Dim nbNoms As Long
    Dim nbLiensExternes As Long

    On Error Resume Next
    nbNoms = wb.Names.Count
    On Error GoTo 0

    nbLiensExternes = CompterLiensExternes(wb)

    ' Ne touche pas à S1.
    ' S1 doit être préparé une seule fois manuellement.
    wsAudit.Range("S2").Value = "Nb feuilles analysées"
    wsAudit.Range("T2").Value = wb.Worksheets.Count - 1

    wsAudit.Range("S3").Value = "Nb noms définis classeur"
    wsAudit.Range("T3").Value = nbNoms

    wsAudit.Range("S4").Value = "Nb liens externes"
    wsAudit.Range("T4").Value = nbLiensExternes

End Sub

' =============================================
' 14. CompterLiensExternes
' =============================================
Private Function CompterLiensExternes(ByVal wb As Workbook) As Long

    Dim arr As Variant

    On Error Resume Next
    arr = wb.LinkSources(xlExcelLinks)
    On Error GoTo 0

    If IsEmpty(arr) Then
        CompterLiensExternes = 0
    Else
        CompterLiensExternes = UBound(arr) - LBound(arr) + 1
    End If

End Function

' =============================================
' 15. MettreEnFormeAudit
' =============================================
Private Sub MettreEnFormeAudit(ByVal wsAudit As Worksheet)

    ' Ne touche pas aux largeurs de colonnes.
    ' Ne touche pas à la ligne 1.
    ' Ne fait pas d'AutoFit.
    ' Ne modifie pas les retours à la ligne.

    wsAudit.Activate
    wsAudit.Range("A1").Select

End Sub
