Attribute VB_Name = "zRFImportBDD"
Option Explicit

' =============================================
' Synchronisation BDD-RF agents -> BDD-RF perso
' =============================================

Private Const NOM_CLASSEUR_SOURCE_DEFAULT  As String = "BDD-RF-24-04"
Private Const NOM_ONGLET_SOURCE_DEFAULT    As String = "GMC"
Private Const NOM_CLASSEUR_CIBLE_DEFAULT   As String = "BDD-RF"
Private Const NOM_ONGLET_CIBLE_DEFAULT     As String = "GMC"

Private Const ONGLET_ID_ABSENTS   As String = "AG_ID_absents_BDD-RF"
Private Const ONGLET_ID_DOUBLONS  As String = "AG_ID_doublons"
Private Const ONGLET_ECARTS       As String = "AG_Ecarts_E_AD"

Private Const NOM_FORME_ACTUALISATION As String = "Actualisation"

Private Const FERMER_APRES_SYNCHRO As Boolean = True

Private Const COL_DEBUT_SYNCHRO As String = "E"   ' E:AD à synchroniser

' =============================================
' 0-a. Sécurité / mot de passe
' =============================================
Private Function MotDePasseValideImportBDD() As Boolean

    Dim MDP As String

    MDP = InputBox("Mot de passe développeur :", "Synchronisation BDD-RF")

    If MDP <> MDP_DEV Then
        MsgBox "Mot de passe incorrect.", vbCritical
        MotDePasseValideImportBDD = False
    Else
        MotDePasseValideImportBDD = True
    End If

End Function

' =============================================
' 0-b. Déprotection
' =============================================
Private Sub DeprotegerFeuilleSiPossible(ByVal ws As Worksheet)

    On Error GoTo ErrHandler
    ws.Unprotect Password:=MDP_DEV
    Exit Sub

ErrHandler:
    Debug.Print "[zRFImportBDD] Déprotection impossible (" & ws.Name & ") : " & Err.Number & " - " & Err.description
    Err.Clear

End Sub

' =============================================
' 0-c. Protection
' =============================================
Private Sub ProtegerFeuilleSiPossible(ByVal ws As Worksheet)

    On Error GoTo ErrHandler
    ws.Protect Password:=MDP_DEV, UserInterfaceOnly:=True, _
               AllowFiltering:=True, AllowSorting:=True
    Exit Sub

ErrHandler:
    Debug.Print "[zRFImportBDD] Protection impossible (" & ws.Name & ") : " & Err.Number & " - " & Err.description
    Err.Clear

End Sub

' =============================================
' 1. SynchroniserDonneesAgents
' =============================================
Public Sub SynchroniserDonneesAgents()

    Dim wbSource As Workbook
    Dim wbCible  As Workbook
    Dim wsSource As Worksheet
    Dim wsCible  As Worksheet

    Dim lastRowSource As Long
    Dim lastRowCible  As Long

    Dim dictCible As Object
    Dim dictCount As Object

    Dim i As Long
    Dim idVal As String
    Dim confSource As String

    Dim arrSourceID As Variant
    Dim arrSourceEAD As Variant
    Dim arrCibleID As Variant
    Dim arrCibleEAD As Variant

    Dim wsAbs As Worksheet
    Dim wsDoublons As Worksheet
    Dim wsEcarts As Worksheet

    Dim rowAbs As Long
    Dim rowDoublons As Long
    Dim rowEcarts As Long

    Dim nbMaj As Long
    Dim nbAbs As Long
    Dim nbDoublons As Long
    Dim nbEcarts As Long
    Dim nbIgnorees As Long

    Dim ligneCible As Long
    Dim eadSource As Variant
    Dim eadCible As Variant

    Dim rngConfImpactee As Range

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim etatApplicationSauve As Boolean
    Dim cibleDeprotegee As Boolean

    Dim nomClasseurSrc As String
    Dim nomOngletSrc As String
    Dim nomClasseurCible As String
    Dim nomOngletCible As String

    On Error GoTo ErrHandler

    If Not MotDePasseValideImportBDD() Then Exit Sub

    Dim frm As UF_ImportBDD
    Set frm = New UF_ImportBDD

    frm.InitialiserImportBDD NOM_CLASSEUR_SOURCE_DEFAULT, NOM_ONGLET_SOURCE_DEFAULT
    frm.Show vbModal

    Dim bConfirmed As Boolean
    bConfirmed = frm.Confirmed
    nomClasseurSrc = frm.nomClasseurSrc
    nomOngletSrc = frm.nomOngletSrc
    nomClasseurCible = frm.nomClasseurCible
    nomOngletCible = frm.nomOngletCible

    Unload frm
    Set frm = Nothing

    If Not bConfirmed Then Exit Sub

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    etatApplicationSauve = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wbSource = GetWorkbookByBaseName(nomClasseurSrc)
    If wbSource Is Nothing Then
        MsgBox "Classeur source introuvable : " & nomClasseurSrc, vbExclamation
        GoTo SortiePropre
    End If

    Set wbCible = GetWorkbookByBaseName(nomClasseurCible)
    If wbCible Is Nothing Then
        MsgBox "Classeur cible introuvable : " & nomClasseurCible, vbExclamation
        GoTo SortiePropre
    End If

    Set wsSource = GetWorksheetSafe(wbSource, nomOngletSrc)
    If wsSource Is Nothing Then
        MsgBox "Onglet source introuvable : " & nomOngletSrc, vbExclamation
        GoTo SortiePropre
    End If

    Set wsCible = GetWorksheetSafe(wbCible, nomOngletCible)
    If wsCible Is Nothing Then
        MsgBox "Onglet cible introuvable : " & nomOngletCible, vbExclamation
        GoTo SortiePropre
    End If

    DeprotegerFeuilleSiPossible wsCible
    cibleDeprotegee = (Not wsCible.ProtectContents)

    Set wsAbs = PreparerOngletRapport(wbCible, ONGLET_ID_ABSENTS)
    Set wsDoublons = PreparerOngletRapport(wbCible, ONGLET_ID_DOUBLONS)
    Set wsEcarts = PreparerOngletRapport(wbCible, ONGLET_ECARTS)

    InitialiserRapportAbsents wsAbs
    InitialiserRapportDoublons wsDoublons
    InitialiserRapportEcarts wsEcarts

    rowAbs = 2
    rowDoublons = 2
    rowEcarts = 2

    lastRowSource = wsSource.Cells(wsSource.Rows.Count, COL_RF_CONCAT).End(xlUp).Row
    If lastRowSource < ROW_START Then
        MsgBox "Aucune donnée source à traiter.", vbInformation
        GoTo SortiePropre
    End If

    lastRowCible = wsCible.Cells(wsCible.Rows.Count, COL_RF_CONCAT).End(xlUp).Row
    If lastRowCible < ROW_START Then
        MsgBox "Aucune donnée cible à comparer.", vbExclamation
        GoTo SortiePropre
    End If

    arrSourceID = wsSource.Range(COL_RF_CONCAT & ROW_START & ":" & COL_RF_CONCAT & lastRowSource).Value2
    arrCibleID = wsCible.Range(COL_RF_CONCAT & ROW_START & ":" & COL_RF_CONCAT & lastRowCible).Value2

    arrSourceEAD = wsSource.Range(COL_DEBUT_SYNCHRO & ROW_START & ":" & COL_OBS & lastRowSource).Value2
    arrCibleEAD = wsCible.Range(COL_DEBUT_SYNCHRO & ROW_START & ":" & COL_OBS & lastRowCible).Value2

    Set dictCible = CreateObject("Scripting.Dictionary")
    Set dictCount = CreateObject("Scripting.Dictionary")
    dictCible.CompareMode = vbTextCompare
    dictCount.CompareMode = vbTextCompare

    ConstruireIndexCible arrCibleID, dictCible, dictCount

    For i = 1 To UBound(arrSourceID, 1)

        idVal = Trim$(CStr(arrSourceID(i, 1)))
        confSource = Trim$(CStr(wsSource.Cells(i + ROW_START - 1, COL_CONF).Value))

        If idVal <> "" And confSource <> "" Then

            eadSource = ExtraireEADDepuisArray(arrSourceEAD, i)

            If Not dictCount.Exists(idVal) Then

                EcrireRapportAbsent wsAbs, rowAbs, idVal, i + ROW_START - 1, eadSource
                rowAbs = rowAbs + 1
                nbAbs = nbAbs + 1

            ElseIf CLng(dictCount(idVal)) > 1 Then

                EcrireRapportDoublon wsDoublons, rowDoublons, idVal, i + ROW_START - 1, CStr(dictCible(idVal)), eadSource
                rowDoublons = rowDoublons + 1
                nbDoublons = nbDoublons + 1

            Else

                ligneCible = CLng(dictCible(idVal))
                eadCible = ExtraireEADDepuisArray(arrCibleEAD, ligneCible - ROW_START + 1)

                If EADEstVide(eadCible) Then
                    wsCible.Range(COL_DEBUT_SYNCHRO & ligneCible & ":" & COL_OBS & ligneCible).Value = eadSource
                    AjouterCelluleConformiteImpactee wsCible, ligneCible, rngConfImpactee
                    nbMaj = nbMaj + 1

                ElseIf EADEgaux(eadSource, eadCible) Then
                    nbIgnorees = nbIgnorees + 1

                Else
                    EcrireRapportEcart wsEcarts, rowEcarts, idVal, i + ROW_START - 1, ligneCible, eadSource, eadCible
                    rowEcarts = rowEcarts + 1
                    nbEcarts = nbEcarts + 1
                End If

            End If
        End If
    Next i

    If Not rngConfImpactee Is Nothing Then
        Application.Run "'" & wbCible.Name & "'!" & wsCible.CodeName & ".RafraichirCouleursConformiteSurLignes", rngConfImpactee.EntireRow.Address
        Application.Run "'" & wbCible.Name & "'!" & wsCible.CodeName & ".RafraichirCouleursValidationSurLignes", rngConfImpactee.EntireRow.Address
    End If

    AjusterRapports wsAbs
    AjusterRapports wsDoublons
    AjusterRapports wsEcarts

    MettreAJourTexteActualisation wbCible, nomOngletCible, NOM_FORME_ACTUALISATION, nomClasseurSrc
    NettoyerContexteApresSynchronisation wbCible, wsCible
    EnregistrerJournalSynchro wbCible, nomClasseurSrc, nbMaj, nbAbs, nbDoublons, nbEcarts, nbIgnorees

    If FERMER_APRES_SYNCHRO Then
        MsgBox "Synchronisation terminée." & vbCrLf & vbCrLf & _
               "Mises à jour : " & nbMaj & vbCrLf & _
               "ID absents : " & nbAbs & vbCrLf & _
               "ID doublons : " & nbDoublons & vbCrLf & _
               "Écarts valeurs : " & nbEcarts & vbCrLf & _
               "Déjà identiques : " & nbIgnorees & vbCrLf & vbCrLf & _
               "Sauvegarde du fichier en cours." & vbCrLf & _
               "Veuillez rouvrir BDD-RF.", vbInformation

        FinaliserEtFermerApresSynchronisation wbSource, wbCible, wsCible, cibleDeprotegee, _
                                              prevCalculation, prevEnableEvents, prevScreenUpdating
        Exit Sub
    End If

    MsgBox "Synchronisation terminée." & vbCrLf & vbCrLf & _
           "Mises à jour : " & nbMaj & vbCrLf & _
           "ID absents : " & nbAbs & vbCrLf & _
           "ID doublons : " & nbDoublons & vbCrLf & _
           "Écarts valeurs : " & nbEcarts & vbCrLf & _
           "Déjà identiques : " & nbIgnorees, vbInformation

SortiePropre:
    If cibleDeprotegee Then
        ProtegerFeuilleSiPossible wsCible
    End If

    If etatApplicationSauve Then
        Application.Calculation = prevCalculation
        Application.EnableEvents = prevEnableEvents
        Application.ScreenUpdating = prevScreenUpdating
    End If
    Exit Sub

ErrHandler:
    MsgBox "Erreur SynchroniserDonneesAgents RF : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

' =============================================
' 2. FinaliserEtFermerApresSynchronisation
' =============================================
Private Sub FinaliserEtFermerApresSynchronisation(ByVal wbSource As Workbook, _
                                                  ByVal wbCible As Workbook, _
                                                  ByVal wsCible As Worksheet, _
                                                  ByVal cibleDeprotegee As Boolean, _
                                                  ByVal prevCalculation As XlCalculation, _
                                                  ByVal prevEnableEvents As Boolean, _
                                                  ByVal prevScreenUpdating As Boolean)

    On Error GoTo ErrHandler

    If cibleDeprotegee Then
        ProtegerFeuilleSiPossible wsCible
    End If

    Application.CutCopyMode = False
    Application.Calculation = prevCalculation
    Application.ScreenUpdating = prevScreenUpdating

    DesactiverCollageValeursRecherche
    Application.OnKey "^l"
    Application.OnKey "%{F11}"

    Application.EnableEvents = False
    wbCible.Save
    Application.EnableEvents = prevEnableEvents

    If Not wbSource Is Nothing Then
        wbSource.Close SaveChanges:=False
    End If

    wbCible.Close SaveChanges:=False

    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.Calculation = prevCalculation
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents

    If cibleDeprotegee Then
        ProtegerFeuilleSiPossible wsCible
    End If

    Debug.Print "[zRFImportBDD] FinaliserEtFermerApresSynchronisation : " & Err.Number & " - " & Err.description
    MsgBox "Erreur lors de la finalisation / fermeture après synchronisation : " & Err.description, vbExclamation
    On Error GoTo 0

End Sub

' =============================================
' 3. ConstruireIndexCible
' =============================================
Private Sub ConstruireIndexCible(ByVal arrCibleID As Variant, ByVal dictCible As Object, ByVal dictCount As Object)

    Dim i As Long
    Dim idVal As String

    For i = 1 To UBound(arrCibleID, 1)
        idVal = Trim$(CStr(arrCibleID(i, 1)))

        If idVal <> "" Then
            If Not dictCount.Exists(idVal) Then
                dictCount.Add idVal, 1
                dictCible.Add idVal, i + ROW_START - 1
            Else
                dictCount(idVal) = CLng(dictCount(idVal)) + 1
                dictCible(idVal) = CStr(dictCible(idVal)) & "," & CStr(i + ROW_START - 1)
            End If
        End If
    Next i

End Sub

' =============================================
' 4. ExtraireEADDepuisArray
' =============================================
Private Function ExtraireEADDepuisArray(ByVal arr As Variant, ByVal indexLigne As Long) As Variant

    Dim nbCols As Long
    Dim j As Long
    Dim t() As Variant

    nbCols = UBound(arr, 2)
    ReDim t(1 To 1, 1 To nbCols)

    For j = 1 To nbCols
        t(1, j) = arr(indexLigne, j)
    Next j

    ExtraireEADDepuisArray = t

End Function

' =============================================
' 5. EADEstVide
' =============================================
Private Function EADEstVide(ByVal ead As Variant) As Boolean

    Dim j As Long

    For j = 1 To UBound(ead, 2)
        If Trim$(CStr(ead(1, j))) <> "" Then
            EADEstVide = False
            Exit Function
        End If
    Next j

    EADEstVide = True

End Function

' =============================================
' 6. EADEgaux
' =============================================
Private Function EADEgaux(ByVal a As Variant, ByVal b As Variant) As Boolean

    Dim j As Long

    If UBound(a, 2) <> UBound(b, 2) Then
        EADEgaux = False
        Exit Function
    End If

    For j = 1 To UBound(a, 2)
        If NormaliserValeur(a(1, j)) <> NormaliserValeur(b(1, j)) Then
            EADEgaux = False
            Exit Function
        End If
    Next j

    EADEgaux = True

End Function

' =============================================
' 7. NormaliserValeur
' =============================================
Private Function NormaliserValeur(ByVal v As Variant) As String

    If IsError(v) Then
        NormaliserValeur = "#ERREUR#"
    ElseIf IsDate(v) Then
        NormaliserValeur = Format$(Int(CDate(v)), "dd/mm/yyyy")
    Else
        NormaliserValeur = Trim$(CStr(v))
    End If

End Function

' =============================================
' 8. AjouterCelluleConformiteImpactee
' =============================================
Private Sub AjouterCelluleConformiteImpactee(ByVal ws As Worksheet, ByVal lig As Long, ByRef rngConfImpactee As Range)

    Dim rngCell As Range

    If lig < ROW_START Then Exit Sub

    Set rngCell = ws.Range(COL_CONF & lig)

    If rngConfImpactee Is Nothing Then
        Set rngConfImpactee = rngCell
    Else
        Set rngConfImpactee = Union(rngConfImpactee, rngCell)
    End If

End Sub

' =============================================
' 9. MettreAJourTexteActualisation
' =============================================
Private Sub MettreAJourTexteActualisation(ByVal wb As Workbook, ByVal nomOnglet As String, ByVal nomForme As String, ByVal nomSource As String)

    Dim ws As Worksheet
    Dim titre As String

    On Error GoTo Fin

    titre = "Dernière actualisation :"

    Set ws = wb.Worksheets(nomOnglet)

    With ws.Shapes(nomForme).TextFrame
        .Characters.Text = titre & " " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf & _
                           "Source : " & nomSource
        .Characters(1, Len(titre)).Font.Color = RGB(229, 158, 221)
    End With

Fin:

End Sub

' =============================================
' 10. PreparerOngletRapport
' =============================================
Private Function PreparerOngletRapport(ByVal wb As Workbook, ByVal nomOnglet As String) As Worksheet

    Dim ws As Worksheet

    Set ws = GetWorksheetSafe(wb, nomOnglet)

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nomOnglet
    Else
        ws.Cells.Clear
    End If

    Set PreparerOngletRapport = ws

End Function

' =============================================
' 11. InitialiserRapportAbsents
' =============================================
Private Sub InitialiserRapportAbsents(ByVal ws As Worksheet)

    ws.Range("A1:G1").Value = Array("ID", "Ligne source", "Date source", "Nom source", "Conformité source", "Observation source", "Motif")

End Sub

' =============================================
' 12. InitialiserRapportDoublons
' =============================================
Private Sub InitialiserRapportDoublons(ByVal ws As Worksheet)

    ws.Range("A1:H1").Value = Array("ID", "Ligne source", "Lignes cible", "Date source", "Nom source", "Conformité source", "Observation source", "Motif")

End Sub

' =============================================
' 13. InitialiserRapportEcarts
' =============================================
Private Sub InitialiserRapportEcarts(ByVal ws As Worksheet)

    ws.Range("A1:K1").Value = Array("ID", "Ligne source", "Ligne cible", "Date source", "Nom source", "Conformité source", "Observation source", "Date cible", "Nom cible", "Conformité cible", "Observation cible")

End Sub

' =============================================
' 14. EcrireRapportAbsent
' =============================================
Private Sub EcrireRapportAbsent(ByVal ws As Worksheet, ByVal r As Long, ByVal idVal As String, ByVal ligneSource As Long, ByVal eadSource As Variant)

    ws.Cells(r, 1).Value = idVal
    ws.Cells(r, 2).Value = ligneSource
    ws.Cells(r, 3).Value = GetValeurEAD(eadSource, COL_DATE)
    ws.Cells(r, 4).Value = GetValeurEAD(eadSource, COL_NOM)
    ws.Cells(r, 5).Value = GetValeurEAD(eadSource, COL_CONF)
    ws.Cells(r, 6).Value = GetValeurEAD(eadSource, COL_OBS)
    ws.Cells(r, 7).Value = "ID absent du fichier cible"

End Sub

' =============================================
' 15. EcrireRapportDoublon
' =============================================
Private Sub EcrireRapportDoublon(ByVal ws As Worksheet, ByVal r As Long, ByVal idVal As String, ByVal ligneSource As Long, ByVal lignesCible As String, ByVal eadSource As Variant)

    ws.Cells(r, 1).Value = idVal
    ws.Cells(r, 2).Value = ligneSource
    ws.Cells(r, 3).Value = lignesCible
    ws.Cells(r, 4).Value = GetValeurEAD(eadSource, COL_DATE)
    ws.Cells(r, 5).Value = GetValeurEAD(eadSource, COL_NOM)
    ws.Cells(r, 6).Value = GetValeurEAD(eadSource, COL_CONF)
    ws.Cells(r, 7).Value = GetValeurEAD(eadSource, COL_OBS)
    ws.Cells(r, 8).Value = "ID présent plusieurs fois dans la cible"

End Sub

' =============================================
' 16. EcrireRapportEcart
' =============================================
Private Sub EcrireRapportEcart(ByVal ws As Worksheet, ByVal r As Long, ByVal idVal As String, ByVal ligneSource As Long, ByVal ligneCible As Long, ByVal eadSource As Variant, ByVal eadCible As Variant)

    ws.Cells(r, 1).Value = idVal
    ws.Cells(r, 2).Value = ligneSource
    ws.Cells(r, 3).Value = ligneCible

    ws.Cells(r, 4).Value = GetValeurEAD(eadSource, COL_DATE)
    ws.Cells(r, 5).Value = GetValeurEAD(eadSource, COL_NOM)
    ws.Cells(r, 6).Value = GetValeurEAD(eadSource, COL_CONF)
    ws.Cells(r, 7).Value = GetValeurEAD(eadSource, COL_OBS)

    ws.Cells(r, 8).Value = GetValeurEAD(eadCible, COL_DATE)
    ws.Cells(r, 9).Value = GetValeurEAD(eadCible, COL_NOM)
    ws.Cells(r, 10).Value = GetValeurEAD(eadCible, COL_CONF)
    ws.Cells(r, 11).Value = GetValeurEAD(eadCible, COL_OBS)

End Sub

' =============================================
' 17. ColNum
' Convertit une lettre de colonne en numéro
' Exemple : "A" = 1, "AA" = 27
' =============================================
Private Function ColNum(ByVal colLettre As String) As Long

    ColNum = ThisWorkbook.Worksheets(SHEET_MAIN).Range(colLettre & "1").Column

End Function

' =============================================
' 18. GetValeurEAD
' =============================================
Private Function GetValeurEAD(ByVal ead As Variant, ByVal colLettre As String) As Variant

    Dim idxRelatif As Long

    idxRelatif = ColNum(colLettre) - ColNum(COL_DEBUT_SYNCHRO) + 1

    If idxRelatif >= 1 And idxRelatif <= UBound(ead, 2) Then
        GetValeurEAD = ead(1, idxRelatif)
    Else
        GetValeurEAD = ""
    End If

End Function

' =============================================
' 19. AjusterRapports
' =============================================
Private Sub AjusterRapports(ByVal ws As Worksheet)

    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit

End Sub

' =============================================
' 20. GetWorkbookByBaseName
' =============================================
Private Function GetWorkbookByBaseName(ByVal baseName As String) As Workbook

    Dim wb As Workbook
    Dim nomSansExtension As String

    For Each wb In Application.Workbooks
        nomSansExtension = wb.Name

        If InStrRev(nomSansExtension, ".") > 0 Then
            nomSansExtension = Left$(nomSansExtension, InStrRev(nomSansExtension, ".") - 1)
        End If

        If StrComp(nomSansExtension, baseName, vbTextCompare) = 0 Then
            Set GetWorkbookByBaseName = wb
            Exit Function
        End If
    Next wb

End Function

' =============================================
' 21. GetWorksheetSafe
' =============================================
Private Function GetWorksheetSafe(ByVal wb As Workbook, ByVal nomOnglet As String) As Worksheet

    On Error Resume Next
    Set GetWorksheetSafe = wb.Worksheets(nomOnglet)
    On Error GoTo 0

End Function

' =============================================
' 22. NettoyerContexteApresSynchronisation
' =============================================
Private Sub NettoyerContexteApresSynchronisation(ByVal wbCible As Workbook, ByVal wsCible As Worksheet)

    On Error GoTo Fin

    Application.CutCopyMode = False

    wbCible.Activate
    wsCible.Activate
    wsCible.Range("A1").Select

    Application.CutCopyMode = False

Fin:
    Err.Clear

End Sub

' =============================================
' 23. EnregistrerJournalSynchro
' =============================================
Private Sub EnregistrerJournalSynchro(ByVal wb As Workbook, _
                                      ByVal nomSource As String, _
                                      ByVal nbMaj As Long, _
                                      ByVal nbAbs As Long, _
                                      ByVal nbDoublons As Long, _
                                      ByVal nbEcarts As Long, _
                                      ByVal nbIgnorees As Long)

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error GoTo Fin

    Set ws = GetOrCreateSheetSynchro(wb)

    If Trim$(CStr(ws.Range("A1").Value)) = "" Then
        ws.Range("A1:H1").Value = Array("Date", "Heure", "Source", "Mises à jour", "ID absents", "ID doublons", "Écarts valeurs", "Déjà identiques")
        ws.Rows(1).Font.Bold = True
    End If

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Date
    ws.Cells(nextRow, 1).NumberFormat = "dd/mm/yyyy"

    ws.Cells(nextRow, 2).Value = Time
    ws.Cells(nextRow, 2).NumberFormat = "hh:mm:ss"

    ws.Cells(nextRow, 3).Value = nomSource
    ws.Cells(nextRow, 4).Value = nbMaj
    ws.Cells(nextRow, 5).Value = nbAbs
    ws.Cells(nextRow, 6).Value = nbDoublons
    ws.Cells(nextRow, 7).Value = nbEcarts
    ws.Cells(nextRow, 8).Value = nbIgnorees

    ws.Columns("A:H").AutoFit

Fin:

End Sub

' =============================================
' 24. GetOrCreateSheetSynchro
' =============================================
Private Function GetOrCreateSheetSynchro(ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet

    Set ws = GetWorksheetSafe(wb, "Synchro")

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Synchro"
    End If

    Set GetOrCreateSheetSynchro = ws

End Function

