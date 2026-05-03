Attribute VB_Name = "zRFGitSync"
Option Explicit

' =============================================
' Synchronisation BDD-RF <-> Dossier
' - src        = version UTF-8 pour Codex / GitHub
' - vba_import = version native destinée au réimport Excel
'
' Version stabilisée :
' - Export complet vers src + vba_import
' - Import depuis vba_import uniquement
' - zRFConstance importé en premier
' - .bas / .cls / .frm importés en natif avec remplacement transactionnel
' - nettoyage des composants temporaires tmpOld_... avant import
' - GMC.txt / ThisWorkbook.txt remplacés en place
' - zRFGitSync et UF_GitSync ne sont jamais réimportés automatiquement
' - Journal disque : %TEMP%\BDD_RF_IMPORT_DEBUG.txt
' =============================================

Private Const DOSSIER_REPO_PC1 As String = "C:\Users\FMF00CDN\Desktop\BDD-RF"
Private Const DOSSIER_REPO_PC2 As String = "C:\BDD-RF"

Private Const NOM_THISWORKBOOK As String = "ThisWorkbook"
Private Const NOM_FEUILLE_CIBLE As String = "GMC"
Private Const NOM_MODULE_UTILITAIRE As String = "zRFGitSync"
Private Const NOM_USERFORM_GITSYNC As String = "UF_GitSync"

Private Const MODULES_PRIORITAIRES As String = "zRFConstance"

Private Const TYPE_STD_MODULE As Long = 1
Private Const TYPE_CLASS_MODULE As Long = 2
Private Const TYPE_USERFORM As Long = 3
Private Const TYPE_DOCUMENT As Long = 100

Private Type ComposantAImporter
    cheminFichier As String
    nomComp As String
    typeComp As Long
    estPrioritaire As Boolean
End Type

' =============================================
' 0. Wrappers simples
' =============================================
Public Sub ExporterVersDossier()
    ExporterProjetVersDossier
End Sub

Public Sub ReimporterDansExcel()
    ImporterProjetVersExcel
End Sub

Public Sub LancerVerification()
    VerifierPreRequisGitSync
End Sub

' =============================================
' DemanderCheminRepo
' =============================================
Private Function DemanderCheminRepo(Optional ByVal modeAction As String = "") As String

    Dim chemin As String
    Dim cleRegistre As String
    Dim frm As UF_GitSync

    cleRegistre = "BDD-RF"

    Set frm = New UF_GitSync

    frm.InitialiserGitSync cleRegistre, DOSSIER_REPO_PC1, DOSSIER_REPO_PC2, modeAction
    frm.Show vbModal

    If frm.Confirmed Then
        chemin = frm.CheminRepo
    End If

    Unload frm
    Set frm = Nothing

    DemanderCheminRepo = chemin

End Function

' =============================================
' ExporterProjetVersDossier
' =============================================
Public Sub ExporterProjetVersDossier()

    Dim vbProj As Object
    Dim vbComp As Object
    Dim CheminRepo As String
    Dim cheminCodex As String
    Dim cheminImport As String
    Dim cheminTxt As String

    If Not ProjetVBAccessible() Then Exit Sub

    CheminRepo = DemanderCheminRepo("EXPORT")
    If Len(CheminRepo) = 0 Then Exit Sub

    cheminCodex = CheminRepo & "\src"
    cheminImport = CheminRepo & "\vba_import"
    cheminTxt = CheminRepo & "\src_txt"

    CreerDossierSiAbsent CheminRepo
    CreerDossierSiAbsent cheminCodex
    CreerDossierSiAbsent cheminImport
    CreerDossierSiAbsent cheminTxt

    Set vbProj = ThisWorkbook.VBProject

    For Each vbComp In vbProj.VBComponents

        If DoitEtreIgnoreExport(vbComp.Name) Then GoTo SuiteComposant

        Select Case vbComp.Type

            Case TYPE_STD_MODULE
                ExporterCodeLisibleUTF8 vbComp, cheminCodex & "\" & vbComp.Name & ".bas"
                ExporterCodeLisibleUTF8 vbComp, cheminTxt & "\" & vbComp.Name & ".bas.txt"
                ExporterComposantNatif vbComp, cheminImport & "\" & vbComp.Name & ".bas"

            Case TYPE_CLASS_MODULE
                ExporterCodeLisibleUTF8 vbComp, cheminCodex & "\" & vbComp.Name & ".cls"
                ExporterCodeLisibleUTF8 vbComp, cheminTxt & "\" & vbComp.Name & ".cls.txt"
                ExporterComposantNatif vbComp, cheminImport & "\" & vbComp.Name & ".cls"

            Case TYPE_USERFORM
                ExporterUserFormComplet vbComp, cheminCodex, cheminImport, cheminTxt

            Case TYPE_DOCUMENT
                If vbComp.Name = NOM_THISWORKBOOK Or vbComp.Name = NOM_FEUILLE_CIBLE Then
                    ExporterCodeLisibleUTF8 vbComp, cheminCodex & "\" & vbComp.Name & ".txt"
                    ExporterCodeLisibleUTF8 vbComp, cheminTxt & "\" & vbComp.Name & ".txt"
                    ExporterDocumentModuleBrut vbComp, cheminImport & "\" & vbComp.Name & ".txt"
                End If

        End Select

SuiteComposant:
    Next vbComp

    MettreAJourShapeGitSync "DernierExport"

    MsgBox "Export terminé :" & vbCrLf & _
           "- GitHub / Codex : " & cheminCodex & vbCrLf & _
           "- Réimport Excel : " & cheminImport & vbCrLf & _
           "- Lecture ChatGPT : " & cheminTxt, vbInformation

End Sub


' =============================================
' ImporterProjetVersExcel
' =============================================
Public Sub ImporterProjetVersExcel()

    On Error GoTo ErrHandler

    Dim CheminRepo As String
    Dim cheminImport As String
    Dim liste() As ComposantAImporter
    Dim nbListe As Long
    Dim nbImportes As Long
    Dim nbDocs As Long
    Dim nbErreurs As Long
    Dim journalErreurs As String
    Dim msg As String
    Dim i As Long
    Dim etatApplicationSauve As Boolean
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean

    ViderJournalImportCrash
    JournaliserImportCrash "DEBUT ImporterProjetVersExcel"

    If Not ProjetVBAccessible() Then
        JournaliserImportCrash "STOP", "ProjetVBAccessible = False"
        Exit Sub
    End If

    JournaliserImportCrash "AVANT NettoyerComposantsTemporairesGitSync"
    NettoyerComposantsTemporairesGitSync
    JournaliserImportCrash "APRES NettoyerComposantsTemporairesGitSync"

    JournaliserImportCrash "AVANT DemanderCheminRepo"
    CheminRepo = DemanderCheminRepo("IMPORT")
    JournaliserImportCrash "APRES DemanderCheminRepo", CheminRepo

    If Len(CheminRepo) = 0 Then
        JournaliserImportCrash "STOP", "CheminRepo vide"
        Exit Sub
    End If

    cheminImport = CheminRepo & "\vba_import"
    JournaliserImportCrash "cheminImport", cheminImport

    If Not DossierExiste(cheminImport) Then
        JournaliserImportCrash "STOP", "Dossier introuvable : " & cheminImport
        MsgBox "Dossier introuvable : " & cheminImport, vbExclamation
        Exit Sub
    End If

    JournaliserImportCrash "AVANT VerifierAvantImport"

    If Not VerifierAvantImport(cheminImport) Then
        JournaliserImportCrash "STOP", "VerifierAvantImport = False"
        MsgBox "Import annulé.", vbInformation
        Exit Sub
    End If

    JournaliserImportCrash "APRES VerifierAvantImport", "Confirmation OK"

    nbListe = 0

    JournaliserImportCrash "AVANT CollecterModulesPrioritaires"
    CollecterModulesPrioritaires cheminImport, liste, nbListe
    JournaliserImportCrash "APRES CollecterModulesPrioritaires", "nbListe=" & nbListe

    JournaliserImportCrash "AVANT Collecter *.bas"
    CollecterFichiersParExtension cheminImport, "*.bas", TYPE_STD_MODULE, liste, nbListe
    JournaliserImportCrash "APRES Collecter *.bas", "nbListe=" & nbListe

    JournaliserImportCrash "AVANT Collecter *.cls"
    CollecterFichiersParExtension cheminImport, "*.cls", TYPE_CLASS_MODULE, liste, nbListe
    JournaliserImportCrash "APRES Collecter *.cls", "nbListe=" & nbListe

    JournaliserImportCrash "AVANT Collecter *.frm"
    CollecterFichiersParExtension cheminImport, "*.frm", TYPE_USERFORM, liste, nbListe
    JournaliserImportCrash "APRES Collecter *.frm", "nbListe=" & nbListe

    If nbListe = 0 Then
        JournaliserImportCrash "STOP", "Aucun composant à importer"
        MsgBox "Aucun composant à importer.", vbInformation
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    etatApplicationSauve = True

    JournaliserImportCrash "AVANT Désactivation Application"

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    JournaliserImportCrash "APRES Désactivation Application"

    For i = 1 To nbListe
        If liste(i).estPrioritaire Then

            JournaliserImportCrash "AVANT IMPORT PRIORITAIRE", _
                i & "/" & nbListe & " | " & liste(i).nomComp & " | type=" & liste(i).typeComp & " | " & liste(i).cheminFichier

            If ImporterOuRemplacerComposant(liste(i).cheminFichier, liste(i).nomComp, liste(i).typeComp) Then
                nbImportes = nbImportes + 1
                JournaliserImportCrash "APRES IMPORT PRIORITAIRE OK", liste(i).nomComp
            Else
                nbErreurs = nbErreurs + 1
                journalErreurs = journalErreurs & "  - " & liste(i).nomComp & " (prioritaire non importé)" & vbCrLf
                JournaliserImportCrash "APRES IMPORT PRIORITAIRE ERREUR", liste(i).nomComp
            End If

        End If
    Next i

    For i = 1 To nbListe
        If Not liste(i).estPrioritaire Then

            JournaliserImportCrash "AVANT IMPORT", _
                i & "/" & nbListe & " | " & liste(i).nomComp & " | type=" & liste(i).typeComp & " | " & liste(i).cheminFichier

            If ImporterOuRemplacerComposant(liste(i).cheminFichier, liste(i).nomComp, liste(i).typeComp) Then
                nbImportes = nbImportes + 1
                JournaliserImportCrash "APRES IMPORT OK", liste(i).nomComp
            Else
                nbErreurs = nbErreurs + 1
                journalErreurs = journalErreurs & "  - " & liste(i).nomComp & " (non importé)" & vbCrLf
                JournaliserImportCrash "APRES IMPORT ERREUR", liste(i).nomComp
            End If

        End If
    Next i

    JournaliserImportCrash "AVANT Remplacer GMC.txt"
    If RemplacerCodeDocumentDepuisFichier(cheminImport & "\" & NOM_FEUILLE_CIBLE & ".txt", NOM_FEUILLE_CIBLE) Then
        nbDocs = nbDocs + 1
        JournaliserImportCrash "APRES Remplacer GMC.txt OK"
    Else
        nbErreurs = nbErreurs + 1
        journalErreurs = journalErreurs & "  - " & NOM_FEUILLE_CIBLE & ".txt (document non remplacé)" & vbCrLf
        JournaliserImportCrash "APRES Remplacer GMC.txt ERREUR"
    End If

    JournaliserImportCrash "AVANT Remplacer ThisWorkbook.txt"
    If RemplacerCodeDocumentDepuisFichier(cheminImport & "\" & NOM_THISWORKBOOK & ".txt", NOM_THISWORKBOOK) Then
        nbDocs = nbDocs + 1
        JournaliserImportCrash "APRES Remplacer ThisWorkbook.txt OK"
    Else
        nbErreurs = nbErreurs + 1
        journalErreurs = journalErreurs & "  - " & NOM_THISWORKBOOK & ".txt (document non remplacé)" & vbCrLf
        JournaliserImportCrash "APRES Remplacer ThisWorkbook.txt ERREUR"
    End If

    msg = "Import terminé." & vbCrLf & vbCrLf & _
          "Composants importés/remplacés : " & nbImportes & vbCrLf & _
          "Modules document remplacés : " & nbDocs & vbCrLf & _
          "Total fichiers traités : " & (nbImportes + nbDocs) & vbCrLf & _
          "Erreurs : " & nbErreurs

    If nbErreurs > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Détail :" & vbCrLf & journalErreurs
    End If

    JournaliserImportCrash "AVANT MettreAJourShapeGitSync"
    MettreAJourShapeGitSync "DernierImport"
    JournaliserImportCrash "APRES MettreAJourShapeGitSync"

    JournaliserImportCrash "FIN ImporterProjetVersExcel"

    MsgBox msg, IIf(nbErreurs > 0, vbExclamation, vbInformation)

SortiePropre:
    If etatApplicationSauve Then
        Application.DisplayAlerts = oldAlerts
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If

    JournaliserImportCrash "SORTIE PROPRE"
    Exit Sub

ErrHandler:
    JournaliserImportCrash "ERREUR VBA", Err.Number & " - " & Err.description
    MsgBox "Erreur import critique : " & Err.Number & " - " & Err.description, vbCritical
    Resume SortiePropre

End Sub

' =============================================
' NettoyerComposantsTemporairesGitSync
' =============================================
Private Sub NettoyerComposantsTemporairesGitSync()

    On Error GoTo ErrHandler

    Dim vbProj As Object
    Dim vbComp As Object
    Dim i As Long
    Dim nomComp As String

    Set vbProj = ThisWorkbook.VBProject

    DechargerUserFormsCharges

    For i = vbProj.VBComponents.Count To 1 Step -1

        Set vbComp = vbProj.VBComponents(i)
        nomComp = vbComp.Name

        If Left$(nomComp, 7) = "tmpOld_" Then
            JournaliserImportCrash "SUPPRESSION COMPOSANT TEMPORAIRE", nomComp
            vbProj.VBComponents.Remove vbComp
        End If

    Next i

    Exit Sub

ErrHandler:
    JournaliserImportCrash "ERREUR NettoyerComposantsTemporairesGitSync", _
                           Err.Number & " - " & Err.description

End Sub

' =============================================
' VerifierAvantImport
' =============================================
Private Function VerifierAvantImport(ByVal cheminImport As String) As Boolean

    Dim listeTrouves As String
    Dim listeManquants As String
    Dim nbTrouves As Long
    Dim nbManquants As Long
    Dim msg As String
    Dim reponse As VbMsgBoxResult

    VerifierFichiersDryRun cheminImport, MODULES_PRIORITAIRES, ".bas", _
                           listeTrouves, listeManquants, nbTrouves, nbManquants

    VerifierExtensionDryRun cheminImport, "*.bas", _
                            listeTrouves, listeManquants, nbTrouves, nbManquants

    VerifierExtensionDryRun cheminImport, "*.cls", _
                            listeTrouves, listeManquants, nbTrouves, nbManquants

    VerifierExtensionDryRun cheminImport, "*.frm", _
                            listeTrouves, listeManquants, nbTrouves, nbManquants

    VerifierDocumentDryRun cheminImport, NOM_FEUILLE_CIBLE, _
                           listeTrouves, listeManquants, nbTrouves, nbManquants

    VerifierDocumentDryRun cheminImport, NOM_THISWORKBOOK, _
                           listeTrouves, listeManquants, nbTrouves, nbManquants

    msg = "=== Vérification avant import ===" & vbCrLf & vbCrLf

    msg = msg & "TROUVÉS (" & nbTrouves & ") :" & vbCrLf
    If nbTrouves > 0 Then
        msg = msg & listeTrouves
    Else
        msg = msg & "  (aucun)" & vbCrLf
    End If

    msg = msg & vbCrLf & "MANQUANTS (" & nbManquants & ") :" & vbCrLf
    If nbManquants > 0 Then
        msg = msg & listeManquants
    Else
        msg = msg & "  (aucun)" & vbCrLf
    End If

    msg = msg & vbCrLf & "Lancer l'import ?"

    If nbManquants > 0 Then
        msg = msg & vbCrLf & vbCrLf & "ATTENTION : les fichiers manquants ne seront pas importés."
    End If

    reponse = MsgBox(msg, vbYesNo + IIf(nbManquants > 0, vbExclamation, vbQuestion), _
                     "Confirmation import")

    VerifierAvantImport = (reponse = vbYes)

End Function

' =============================================
' VerifierFichiersDryRun
' =============================================
Private Sub VerifierFichiersDryRun(ByVal cheminImport As String, _
                                   ByVal listeNoms As String, _
                                   ByVal extension As String, _
                                   ByRef listeTrouves As String, _
                                   ByRef listeManquants As String, _
                                   ByRef nbTrouves As Long, _
                                   ByRef nbManquants As Long)

    Dim noms() As String
    Dim i As Long
    Dim nomComp As String
    Dim chemin As String

    noms = Split(listeNoms, ",")

    For i = LBound(noms) To UBound(noms)
        nomComp = Trim$(noms(i))
        If Len(nomComp) = 0 Then GoTo SuiteNom

        chemin = cheminImport & "\" & nomComp & extension

        If FichierExiste(chemin) Then
            nbTrouves = nbTrouves + 1
            listeTrouves = listeTrouves & "  [PRIO] " & nomComp & extension & vbCrLf
        Else
            nbManquants = nbManquants + 1
            listeManquants = listeManquants & "  [PRIO] " & nomComp & extension & vbCrLf
        End If

SuiteNom:
    Next i

End Sub

' =============================================
' VerifierExtensionDryRun
' =============================================
Private Sub VerifierExtensionDryRun(ByVal cheminImport As String, _
                                    ByVal filtre As String, _
                                    ByRef listeTrouves As String, _
                                    ByRef listeManquants As String, _
                                    ByRef nbTrouves As Long, _
                                    ByRef nbManquants As Long)

    Dim fichier As String
    Dim nomComp As String
    Dim cheminFichier As String

    fichier = Dir(cheminImport & "\" & filtre)

    Do While Len(fichier) > 0
        nomComp = nomSansExtension(fichier)
        cheminFichier = cheminImport & "\" & fichier

        If Not DoitEtreIgnoreImport(nomComp) _
           And Not FichierImportDangereux(cheminFichier) _
           And Not EstModulePrioritaire(nomComp) Then

            nbTrouves = nbTrouves + 1
            listeTrouves = listeTrouves & "  " & fichier & vbCrLf

        End If

        fichier = Dir
    Loop

End Sub

' =============================================
' VerifierDocumentDryRun
' =============================================
Private Sub VerifierDocumentDryRun(ByVal cheminImport As String, _
                                   ByVal nomComp As String, _
                                   ByRef listeTrouves As String, _
                                   ByRef listeManquants As String, _
                                   ByRef nbTrouves As Long, _
                                   ByRef nbManquants As Long)

    Dim chemin As String

    chemin = cheminImport & "\" & nomComp & ".txt"

    If FichierExiste(chemin) Then
        nbTrouves = nbTrouves + 1
        listeTrouves = listeTrouves & "  [DOC] " & nomComp & ".txt" & vbCrLf
    Else
        nbManquants = nbManquants + 1
        listeManquants = listeManquants & "  [DOC] " & nomComp & ".txt" & vbCrLf
    End If

End Sub

' =============================================
' CollecterModulesPrioritaires
' =============================================
Private Sub CollecterModulesPrioritaires(ByVal cheminImport As String, _
                                         ByRef liste() As ComposantAImporter, _
                                         ByRef nbListe As Long)

    Dim listePrio() As String
    Dim i As Long
    Dim nomComp As String
    Dim cheminFichier As String

    listePrio = Split(MODULES_PRIORITAIRES, ",")

    For i = LBound(listePrio) To UBound(listePrio)

        nomComp = Trim$(listePrio(i))
        If Len(nomComp) = 0 Then GoTo SuitePrio

        cheminFichier = cheminImport & "\" & nomComp & ".bas"
        If Not FichierExiste(cheminFichier) Then GoTo SuitePrio

        nbListe = nbListe + 1
        ReDim Preserve liste(1 To nbListe)
        liste(nbListe).cheminFichier = cheminFichier
        liste(nbListe).nomComp = nomComp
        liste(nbListe).typeComp = TYPE_STD_MODULE
        liste(nbListe).estPrioritaire = True

SuitePrio:
    Next i

End Sub

' =============================================
' CollecterFichiersParExtension
' =============================================
Private Sub CollecterFichiersParExtension(ByVal cheminImport As String, _
                                          ByVal filtre As String, _
                                          ByVal typeComp As Long, _
                                          ByRef liste() As ComposantAImporter, _
                                          ByRef nbListe As Long)

    Dim fichiers() As String
    Dim nbFichiers As Long
    Dim fichier As String
    Dim nomComp As String
    Dim cheminFichier As String
    Dim i As Long

    nbFichiers = 0
    fichier = Dir(cheminImport & "\" & filtre)

    Do While Len(fichier) > 0
        nbFichiers = nbFichiers + 1
        ReDim Preserve fichiers(1 To nbFichiers)
        fichiers(nbFichiers) = fichier
        fichier = Dir
    Loop

    If nbFichiers = 0 Then Exit Sub

    For i = 1 To nbFichiers

        nomComp = nomSansExtension(fichiers(i))
        cheminFichier = cheminImport & "\" & fichiers(i)

        If DoitEtreIgnoreImport(nomComp) Then GoTo SuiteFichier
        If FichierImportDangereux(cheminFichier) Then GoTo SuiteFichier
        If EstModulePrioritaire(nomComp) Then GoTo SuiteFichier

        nbListe = nbListe + 1
        ReDim Preserve liste(1 To nbListe)
        liste(nbListe).cheminFichier = cheminFichier
        liste(nbListe).nomComp = nomComp
        liste(nbListe).typeComp = typeComp
        liste(nbListe).estPrioritaire = False

SuiteFichier:
    Next i

End Sub

' =============================================
' ImporterOuRemplacerComposant
' =============================================
Private Function ImporterOuRemplacerComposant(ByVal cheminFichier As String, _
                                              ByVal nomComp As String, _
                                              ByVal typeComp As Long) As Boolean

    On Error GoTo ErrHandler

    JournaliserImportCrash "ENTREE ImporterOuRemplacerComposant", nomComp & " | " & cheminFichier

    If Not FichierExiste(cheminFichier) Then
        JournaliserImportCrash "FICHIER INTROUVABLE", cheminFichier
        JournaliserIO "ImporterOuRemplacerComposant", cheminFichier, 0, "Fichier introuvable"
        ImporterOuRemplacerComposant = False
        Exit Function
    End If

    If FichierImportDangereux(cheminFichier) Then
        JournaliserImportCrash "FICHIER DANGEREUX IGNORE", nomComp & " | " & cheminFichier
        JournaliserIO "ImporterOuRemplacerComposant", cheminFichier, 0, "Fichier ignoré car dangereux pour l'import en cours"
        ImporterOuRemplacerComposant = True
        Exit Function
    End If

    Select Case typeComp

        Case TYPE_STD_MODULE, TYPE_CLASS_MODULE
            ImporterOuRemplacerComposant = ImporterOuRemplacerComposantNatif(cheminFichier, nomComp, typeComp)

        Case TYPE_USERFORM
            ImporterOuRemplacerComposant = ImporterOuMettreAJourUserForm(cheminFichier, nomComp)

        Case Else
            JournaliserImportCrash "TYPE NON PRIS EN CHARGE", nomComp & " | type=" & typeComp
            ImporterOuRemplacerComposant = False

    End Select

    Exit Function

ErrHandler:
    JournaliserImportCrash "ERREUR ImporterOuRemplacerComposant", nomComp & " | " & Err.Number & " - " & Err.description
    JournaliserIO "ImporterOuRemplacerComposant", cheminFichier, Err.Number, Err.description
    ImporterOuRemplacerComposant = False

End Function

' =============================================
' ImporterOuRemplacerComposantNatif
'
' Réservé aux .bas / .cls.
' Les .frm sont traités par ImporterOuMettreAJourUserForm.
' =============================================
Private Function ImporterOuRemplacerComposantNatif(ByVal cheminFichier As String, _
                                                   ByVal nomComp As String, _
                                                   ByVal typeComp As Long) As Boolean

    On Error GoTo ErrHandler

    Dim vbCompExistant As Object
    Dim vbCompImporte As Object
    Dim nomTemp As String
    Dim ancienRenomme As Boolean
    Dim nomImporte As String

    JournaliserImportCrash "ENTREE ImporterOuRemplacerComposantNatif", _
                           nomComp & " | type=" & typeComp & " | " & cheminFichier

    If typeComp = TYPE_USERFORM Then
        JournaliserImportCrash "NATIF REFUSE USERFORM", nomComp
        ImporterOuRemplacerComposantNatif = False
        Exit Function
    End If

    If Not FichierExiste(cheminFichier) Then
        JournaliserImportCrash "NATIF FICHIER INTROUVABLE", cheminFichier
        ImporterOuRemplacerComposantNatif = False
        Exit Function
    End If

    Set vbCompExistant = ObtenirComposantVBA(nomComp)

    If Not vbCompExistant Is Nothing Then

        JournaliserImportCrash "NATIF EXISTANT", nomComp & " | Type=" & vbCompExistant.Type

        If vbCompExistant.Type <> typeComp Then
            JournaliserImportCrash "NATIF TYPE INCOHERENT", _
                                   nomComp & " | Type réel=" & vbCompExistant.Type & " | Type attendu=" & typeComp
            ImporterOuRemplacerComposantNatif = False
            Exit Function
        End If

        nomTemp = NomTemporaireLibre(nomComp)

        JournaliserImportCrash "NATIF AVANT RENOMMAGE ANCIEN", nomComp & " -> " & nomTemp

        If Not RenommerComposantVBA(vbCompExistant, nomTemp) Then
            JournaliserImportCrash "NATIF RENOMMAGE ANCIEN ECHEC", nomComp & " -> " & nomTemp
            ImporterOuRemplacerComposantNatif = False
            Exit Function
        End If

        ancienRenomme = True
        JournaliserImportCrash "NATIF APRES RENOMMAGE ANCIEN", nomTemp

    End If

    JournaliserImportCrash "NATIF AVANT IMPORT", cheminFichier
    Set vbCompImporte = ThisWorkbook.VBProject.VBComponents.Import(cheminFichier)
    JournaliserImportCrash "NATIF APRES IMPORT", cheminFichier

    If vbCompImporte Is Nothing Then
        JournaliserImportCrash "NATIF IMPORT RETOUR NOTHING", nomComp
        GoTo RestaurerAncien
    End If

    If vbCompImporte.Type <> typeComp Then
        JournaliserImportCrash "NATIF TYPE IMPORTE INCOHERENT", _
                               nomComp & " | Type importé=" & vbCompImporte.Type & " | Type attendu=" & typeComp
        GoTo RestaurerAncien
    End If

    nomImporte = vbCompImporte.Name

    JournaliserImportCrash "NATIF NOM IMPORTE", "Attendu=" & nomComp & " | Obtenu=" & nomImporte

    If StrComp(nomImporte, nomComp, vbTextCompare) <> 0 Then

        JournaliserImportCrash "NATIF AVANT RENOMMAGE IMPORTE", nomImporte & " -> " & nomComp

        If Not RenommerComposantVBA(vbCompImporte, nomComp) Then
            JournaliserImportCrash "NATIF RENOMMAGE IMPORTE ECHEC", nomImporte & " -> " & nomComp
            GoTo RestaurerAncien
        End If

        JournaliserImportCrash "NATIF APRES RENOMMAGE IMPORTE", nomComp

    End If

    If ancienRenomme Then
        JournaliserImportCrash "NATIF AVANT REMOVE ANCIEN TEMP", nomTemp
        ThisWorkbook.VBProject.VBComponents.Remove vbCompExistant
        JournaliserImportCrash "NATIF APRES REMOVE ANCIEN TEMP", nomTemp
    End If

    ImporterOuRemplacerComposantNatif = True
    JournaliserImportCrash "SORTIE OK ImporterOuRemplacerComposantNatif", nomComp
    Exit Function

RestaurerAncien:
    On Error Resume Next

    If Not vbCompImporte Is Nothing Then
        JournaliserImportCrash "NATIF SUPPRESSION IMPORTE ECHEC", vbCompImporte.Name
        ThisWorkbook.VBProject.VBComponents.Remove vbCompImporte
    End If

    If ancienRenomme Then
        JournaliserImportCrash "NATIF RESTAURATION ANCIEN", nomTemp & " -> " & nomComp
        Call RenommerComposantVBA(vbCompExistant, nomComp)
    End If

    On Error GoTo 0

    ImporterOuRemplacerComposantNatif = False
    Exit Function

ErrHandler:
    JournaliserImportCrash "ERREUR ImporterOuRemplacerComposantNatif", _
                           nomComp & " | " & Err.Number & " - " & Err.description
    JournaliserIO "ImporterOuRemplacerComposantNatif", cheminFichier, Err.Number, Err.description

    On Error Resume Next

    If Not vbCompImporte Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbCompImporte
    End If

    If ancienRenomme Then
        Call RenommerComposantVBA(vbCompExistant, nomComp)
    End If

    On Error GoTo 0

    ImporterOuRemplacerComposantNatif = False

End Function

' =============================================
' ImporterOuMettreAJourUserForm
'
' Si le UserForm existe déjà :
'   - on garde le design existant
'   - on remplace uniquement le code VBA interne du UserForm
'
' Si le UserForm n'existe pas :
'   - import natif du .frm
'
' Cela évite l'erreur au 2e import :
' "Le nom UF_xxx est déjà utilisé"
' =============================================
Private Function ImporterOuMettreAJourUserForm(ByVal cheminFichier As String, _
                                               ByVal nomComp As String) As Boolean

    On Error GoTo ErrHandler

    Dim vbCompExistant As Object
    Dim vbCompImporte As Object
    Dim contenuFrm As String
    Dim codeVBA As String
    Dim ancienCode As String
    Dim nbLignesAncien As Long

    JournaliserImportCrash "ENTREE ImporterOuMettreAJourUserForm", nomComp & " | " & cheminFichier

    If Not FichierExiste(cheminFichier) Then
        JournaliserImportCrash "USERFORM FICHIER INTROUVABLE", cheminFichier
        ImporterOuMettreAJourUserForm = False
        Exit Function
    End If

    DechargerUserFormsCharges

    Set vbCompExistant = ObtenirComposantVBA(nomComp)

    ' Cas 1 : le UserForm n'existe pas encore -> import natif
    If vbCompExistant Is Nothing Then

        JournaliserImportCrash "USERFORM ABSENT - IMPORT NATIF", nomComp

        Set vbCompImporte = ThisWorkbook.VBProject.VBComponents.Import(cheminFichier)

        If vbCompImporte Is Nothing Then
            JournaliserImportCrash "USERFORM IMPORT NATIF RETOUR NOTHING", nomComp
            ImporterOuMettreAJourUserForm = False
            Exit Function
        End If

        If vbCompImporte.Type <> TYPE_USERFORM Then
            JournaliserImportCrash "USERFORM IMPORT TYPE INCOHERENT", nomComp & " | Type=" & vbCompImporte.Type
            ThisWorkbook.VBProject.VBComponents.Remove vbCompImporte
            ImporterOuMettreAJourUserForm = False
            Exit Function
        End If

        If StrComp(vbCompImporte.Name, nomComp, vbTextCompare) <> 0 Then
            Call RenommerComposantVBA(vbCompImporte, nomComp)
        End If

        ImporterOuMettreAJourUserForm = True
        JournaliserImportCrash "SORTIE OK UserForm import natif", nomComp
        Exit Function

    End If

    ' Cas 2 : le UserForm existe déjà -> remplacement du code uniquement
    If vbCompExistant.Type <> TYPE_USERFORM Then
        JournaliserImportCrash "USERFORM EXISTANT TYPE INCOHERENT", nomComp & " | Type=" & vbCompExistant.Type
        ImporterOuMettreAJourUserForm = False
        Exit Function
    End If

    JournaliserImportCrash "USERFORM EXISTANT - MAJ CODE", nomComp

    contenuFrm = LireFichierVBAImport(cheminFichier)
    codeVBA = ExtraireCodeVBADepuisFrm(contenuFrm)

    With vbCompExistant.CodeModule

        nbLignesAncien = .CountOfLines

        If nbLignesAncien > 0 Then
            ancienCode = .Lines(1, nbLignesAncien)
        Else
            ancienCode = ""
        End If

        JournaliserImportCrash "USERFORM AVANT DeleteLines", nomComp & " | CountOfLines=" & .CountOfLines
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        JournaliserImportCrash "USERFORM APRES DeleteLines", nomComp

        If Len(codeVBA) > 0 Then
            JournaliserImportCrash "USERFORM AVANT AddFromString CODE", nomComp & " | Len=" & Len(codeVBA)
            .AddFromString codeVBA
            JournaliserImportCrash "USERFORM APRES AddFromString CODE", nomComp
        End If

    End With

    ImporterOuMettreAJourUserForm = True
    JournaliserImportCrash "SORTIE OK UserForm MAJ code", nomComp
    Exit Function

ErrHandler:
    JournaliserImportCrash "ERREUR ImporterOuMettreAJourUserForm", _
                           nomComp & " | " & Err.Number & " - " & Err.description
    JournaliserIO "ImporterOuMettreAJourUserForm", cheminFichier, Err.Number, Err.description

    On Error Resume Next

    If Not vbCompExistant Is Nothing Then
        With vbCompExistant.CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            If Len(ancienCode) > 0 Then .AddFromString ancienCode
        End With
    End If

    On Error GoTo 0

    ImporterOuMettreAJourUserForm = False

End Function

' =============================================
' ExtraireCodeVBADepuisFrm
'
' Extrait uniquement le code VBA d'un fichier .frm.
' Ignore :
' - VERSION
' - Begin / End du design
' - propriétés visuelles
' - Attribute ...
' =============================================
Private Function ExtraireCodeVBADepuisFrm(ByVal contenuFrm As String) As String

    Dim lignes() As String
    Dim i As Long
    Dim ligne As String
    Dim ligneTrim As String
    Dim ligneUpper As String
    Dim vuAttribut As Boolean
    Dim dansCode As Boolean
    Dim resultat As String

    contenuFrm = NettoyerBOMTexte(contenuFrm)
    contenuFrm = Replace(contenuFrm, vbCrLf, vbLf)
    contenuFrm = Replace(contenuFrm, vbCr, vbLf)

    lignes = Split(contenuFrm, vbLf)

    For i = LBound(lignes) To UBound(lignes)

        ligne = lignes(i)
        ligneTrim = Trim$(ligne)
        ligneUpper = UCase$(ligneTrim)

        If Left$(ligneUpper, 9) = "ATTRIBUTE" Then
            vuAttribut = True
            GoTo LigneSuivante
        End If

        If vuAttribut And Not dansCode Then
            If Len(ligneTrim) = 0 Then GoTo LigneSuivante
            dansCode = True
        End If

        If dansCode Then
            If resultat = "" Then
                resultat = ligne
            Else
                resultat = resultat & vbCrLf & ligne
            End If
        End If

LigneSuivante:
    Next i

    ExtraireCodeVBADepuisFrm = resultat

End Function


' =============================================
' RenommerComposantVBA
' =============================================
Private Function RenommerComposantVBA(ByVal vbComp As Object, ByVal nouveauNom As String) As Boolean

    On Error GoTo ErrHandler

    If vbComp Is Nothing Then Exit Function

    vbComp.Name = nouveauNom

    If vbComp.Type = TYPE_USERFORM Then
        vbComp.Properties("Name").Value = nouveauNom
    End If

    DoEvents

    RenommerComposantVBA = (StrComp(vbComp.Name, nouveauNom, vbTextCompare) = 0)

    If vbComp.Type = TYPE_USERFORM Then
        RenommerComposantVBA = RenommerComposantVBA And _
            (StrComp(CStr(vbComp.Properties("Name").Value), nouveauNom, vbTextCompare) = 0)
    End If

    Exit Function

ErrHandler:
    JournaliserImportCrash "ERREUR RenommerComposantVBA", _
                           nouveauNom & " | " & Err.Number & " - " & Err.description
    RenommerComposantVBA = False

End Function

' =============================================
' NomTemporaireLibre
' =============================================
Private Function NomTemporaireLibre(ByVal nomComp As String) As String

    Dim baseNom As String
    Dim nomTest As String
    Dim i As Long

    baseNom = "tmpOld_" & nomComp

    If Len(baseNom) > 25 Then
        baseNom = Left$(baseNom, 25)
    End If

    i = 1

    Do
        nomTest = baseNom & "_" & CStr(i)

        If ObtenirComposantVBA(nomTest) Is Nothing Then
            NomTemporaireLibre = nomTest
            Exit Function
        End If

        i = i + 1
    Loop

End Function

' =============================================
' DechargerUserFormsCharges
' =============================================
Private Sub DechargerUserFormsCharges()

    Dim i As Long

    On Error Resume Next

    For i = UserForms.Count - 1 To 0 Step -1
        Unload UserForms(i)
    Next i

    On Error GoTo 0

End Sub

' =============================================
' ExporterCodeLisibleUTF8
' =============================================
Private Sub ExporterCodeLisibleUTF8(ByVal vbComp As Object, ByVal cheminFinal As String)
    EcrireTexteUTF8 cheminFinal, LireCodeAvecHeader(vbComp)
End Sub

' =============================================
' ExporterDocumentModuleBrut
' =============================================
Private Sub ExporterDocumentModuleBrut(ByVal vbComp As Object, ByVal cheminFinal As String)
    EcrireTexteUTF8 cheminFinal, LireCodeBrut(vbComp)
End Sub

' =============================================
' ExporterComposantNatif
' =============================================
Private Sub ExporterComposantNatif(ByVal vbComp As Object, ByVal cheminFinal As String)
    SupprimerFichierSiExiste cheminFinal, "ExporterComposantNatif"
    vbComp.Export cheminFinal
End Sub

' =============================================
' ExporterUserFormComplet
' =============================================
Private Sub ExporterUserFormComplet(ByVal vbComp As Object, _
                                    ByVal cheminCodex As String, _
                                    ByVal cheminImport As String, _
                                    ByVal cheminTxt As String)

    Dim cheminImportFrm As String
    Dim cheminImportFrx As String
    Dim cheminTempFrm As String
    Dim cheminTempFrx As String
    Dim cheminCodexFrm As String
    Dim cheminCodexFrx As String
    Dim cheminTxtFrm As String
    Dim contenuFrm As String

    cheminImportFrm = cheminImport & "\" & vbComp.Name & ".frm"
    cheminImportFrx = cheminImport & "\" & vbComp.Name & ".frx"

    SupprimerFichierSiExiste cheminImportFrm, "ExporterUserFormComplet"
    SupprimerFichierSiExiste cheminImportFrx, "ExporterUserFormComplet"

    vbComp.Export cheminImportFrm

    cheminTempFrm = Environ$("TEMP") & "\" & vbComp.Name & "_codex.frm"
    cheminTempFrx = Environ$("TEMP") & "\" & vbComp.Name & "_codex.frx"

    cheminCodexFrm = cheminCodex & "\" & vbComp.Name & ".frm"
    cheminCodexFrx = cheminCodex & "\" & vbComp.Name & ".frx"
    cheminTxtFrm = cheminTxt & "\" & vbComp.Name & ".frm.txt"

    SupprimerFichierSiExiste cheminTempFrm, "ExporterUserFormComplet"
    SupprimerFichierSiExiste cheminTempFrx, "ExporterUserFormComplet"
    SupprimerFichierSiExiste cheminCodexFrm, "ExporterUserFormComplet"
    SupprimerFichierSiExiste cheminCodexFrx, "ExporterUserFormComplet"
    SupprimerFichierSiExiste cheminTxtFrm, "ExporterUserFormComplet"

    vbComp.Export cheminTempFrm

    contenuFrm = LireFichierTexteSysteme(cheminTempFrm)

    EcrireTexteUTF8 cheminCodexFrm, contenuFrm

    EcrireTexteUTF8 cheminTxtFrm, contenuFrm

    ' Le .frx reste uniquement en binaire, pas de copie .txt
    If FichierExiste(cheminTempFrx) Then
        CopierFichierBinaire cheminTempFrx, cheminCodexFrx
    End If

    SupprimerFichierSiExiste cheminTempFrm, "ExporterUserFormComplet"
    SupprimerFichierSiExiste cheminTempFrx, "ExporterUserFormComplet"

End Sub

' =============================================
' RemplacerCodeDocumentDepuisFichier
' =============================================
Private Function RemplacerCodeDocumentDepuisFichier(ByVal chemin As String, ByVal nomComp As String) As Boolean

    On Error GoTo ErrHandler

    Dim vbComp As Object
    Dim contenu As String
    Dim ancienCode As String
    Dim nbLignesAncien As Long

    JournaliserImportCrash "ENTREE RemplacerCodeDocumentDepuisFichier", nomComp & " | " & chemin

    If Not FichierExiste(chemin) Then
        JournaliserImportCrash "DOC FICHIER INTROUVABLE", nomComp & " | " & chemin
        RemplacerCodeDocumentDepuisFichier = False
        Exit Function
    End If

    Set vbComp = ObtenirComposantVBA(nomComp)

    If vbComp Is Nothing Then
        JournaliserImportCrash "DOC COMPOSANT INTROUVABLE", nomComp
        JournaliserIO "RemplacerCodeDocumentDepuisFichier", nomComp, 0, "Composant document introuvable"
        RemplacerCodeDocumentDepuisFichier = False
        Exit Function
    End If

    If vbComp.Type <> TYPE_DOCUMENT Then
        JournaliserImportCrash "DOC TYPE INVALIDE", nomComp & " | Type=" & vbComp.Type
        JournaliserIO "RemplacerCodeDocumentDepuisFichier", nomComp, 0, "Le composant n'est pas un module document"
        RemplacerCodeDocumentDepuisFichier = False
        Exit Function
    End If

    JournaliserImportCrash "DOC AVANT LireFichierVBAImport", nomComp
    contenu = LireFichierVBAImport(chemin)
    JournaliserImportCrash "DOC APRES LireFichierVBAImport", nomComp & " | Len=" & Len(contenu)

    JournaliserImportCrash "DOC AVANT NettoyerEnteteExport", nomComp
    contenu = NettoyerEnteteExport(contenu)
    JournaliserImportCrash "DOC APRES NettoyerEnteteExport", nomComp & " | Len=" & Len(contenu)

    With vbComp.CodeModule

        nbLignesAncien = .CountOfLines

        If nbLignesAncien > 0 Then
            ancienCode = .Lines(1, nbLignesAncien)
        End If

        JournaliserImportCrash "DOC AVANT DeleteLines", nomComp & " | CountOfLines=" & .CountOfLines
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        JournaliserImportCrash "DOC APRES DeleteLines", nomComp

        JournaliserImportCrash "DOC AVANT AddFromString", nomComp & " | Len=" & Len(contenu)
        If Len(contenu) > 0 Then .AddFromString contenu
        JournaliserImportCrash "DOC APRES AddFromString", nomComp

    End With

    RemplacerCodeDocumentDepuisFichier = True
    JournaliserImportCrash "SORTIE OK RemplacerCodeDocumentDepuisFichier", nomComp
    Exit Function

ErrHandler:
    JournaliserImportCrash "ERREUR RemplacerCodeDocumentDepuisFichier", _
                           nomComp & " | " & Err.Number & " - " & Err.description

    On Error Resume Next

    If Not vbComp Is Nothing Then
        With vbComp.CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            If Len(ancienCode) > 0 Then .AddFromString ancienCode
        End With
    End If

    On Error GoTo 0

    RemplacerCodeDocumentDepuisFichier = False

End Function

' =============================================
' LireCodeAvecHeader
' =============================================
Private Function LireCodeAvecHeader(ByVal vbComp As Object) As String
    LireCodeAvecHeader = "Attribute VB_Name = """ & vbComp.Name & """" & vbCrLf & LireCodeBrut(vbComp)
End Function

' =============================================
' LireCodeBrut
' =============================================
Private Function LireCodeBrut(ByVal vbComp As Object) As String

    If vbComp.CodeModule.CountOfLines > 0 Then
        LireCodeBrut = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
    Else
        LireCodeBrut = ""
    End If

End Function

' =============================================
' NettoyerEnteteExport
' =============================================
Private Function NettoyerEnteteExport(ByVal contenu As String) As String

    Dim lignes() As String
    Dim i As Long
    Dim resultat As String
    Dim ligne As String
    Dim ligneTrimmed As String
    Dim ligneUpper As String
    Dim dansEntete As Boolean
    Dim attributsCommences As Boolean

    If Len(contenu) = 0 Then Exit Function

    contenu = NettoyerBOMTexte(contenu)
    lignes = Split(Replace(contenu, vbCrLf, vbLf), vbLf)

    dansEntete = True
    attributsCommences = False

    For i = LBound(lignes) To UBound(lignes)

        ligne = lignes(i)
        ligneTrimmed = Trim$(ligne)
        ligneUpper = UCase$(ligneTrimmed)

        If dansEntete Then

            If Left$(ligneUpper, 9) = "ATTRIBUTE" Then
                attributsCommences = True
                GoTo LigneSuivante
            End If

            If attributsCommences Then
                dansEntete = False
                GoTo AjouterLigne
            End If

            If LigneSembleEtreDebutCode(ligneTrimmed) Then
                dansEntete = False
                GoTo AjouterLigne
            End If

            GoTo LigneSuivante

        End If

AjouterLigne:
        If resultat = "" Then
            resultat = ligne
        Else
            resultat = resultat & vbCrLf & ligne
        End If

LigneSuivante:
    Next i

    NettoyerEnteteExport = resultat

End Function

' =============================================
' NettoyerBOMTexte
' =============================================
Private Function NettoyerBOMTexte(ByVal contenu As String) As String

    contenu = Replace(contenu, ChrW$(&HFEFF), "")
    contenu = Replace(contenu, "ï»¿", "")
    contenu = Replace(contenu, Chr$(239) & Chr$(187) & Chr$(191), "")

    NettoyerBOMTexte = contenu

End Function

' =============================================
' LigneSembleEtreDebutCode
' =============================================
Private Function LigneSembleEtreDebutCode(ByVal ligneTrimmed As String) As Boolean

    Dim s As String

    If Len(ligneTrimmed) = 0 Then Exit Function

    s = UCase$(ligneTrimmed)

    LigneSembleEtreDebutCode = _
        Left$(s, 6) = "OPTION" Or _
        Left$(s, 6) = "PUBLIC" Or _
        Left$(s, 7) = "PRIVATE" Or _
        Left$(s, 3) = "DIM" Or _
        Left$(s, 5) = "CONST" Or _
        Left$(s, 3) = "SUB" Or _
        Left$(s, 8) = "FUNCTION" Or _
        Left$(s, 8) = "PROPERTY" Or _
        Left$(s, 4) = "TYPE" Or _
        Left$(s, 4) = "ENUM" Or _
        Left$(s, 7) = "DECLARE" Or _
        Left$(s, 1) = "'"

End Function

' =============================================
' ProjetVBAccessible
' =============================================
Private Function ProjetVBAccessible() As Boolean

    On Error GoTo ErrHandler

    Dim n As Long
    n = ThisWorkbook.VBProject.VBComponents.Count

    ProjetVBAccessible = True
    Exit Function

ErrHandler:
    ProjetVBAccessible = False
    MsgBox "Accès refusé au projet VBA." & vbCrLf & _
           "Active l'option :" & vbCrLf & _
           "Fichier > Options > Centre de gestion de la confidentialité > " & _
           "Paramètres des macros > Accès approuvé au modèle d'objet du projet VBA.", vbExclamation

End Function

' =============================================
' ObtenirComposantVBA
' =============================================
Private Function ObtenirComposantVBA(ByVal nomComp As String) As Object

    On Error Resume Next
    Set ObtenirComposantVBA = ThisWorkbook.VBProject.VBComponents(nomComp)
    On Error GoTo 0

End Function

' =============================================
' DoitEtreIgnoreExport
' =============================================
Private Function DoitEtreIgnoreExport(ByVal nomComp As String) As Boolean
    DoitEtreIgnoreExport = False
End Function

' =============================================
' DoitEtreIgnoreImport
' =============================================
Private Function DoitEtreIgnoreImport(ByVal nomComp As String) As Boolean

    Select Case UCase$(Trim$(nomComp))

        Case UCase$(NOM_MODULE_UTILITAIRE)
            DoitEtreIgnoreImport = True

        Case UCase$(NOM_USERFORM_GITSYNC)
            DoitEtreIgnoreImport = True

        Case Else
            DoitEtreIgnoreImport = False

    End Select

End Function

' =============================================
' FichierImportDangereux
' =============================================
Private Function FichierImportDangereux(ByVal cheminFichier As String) As Boolean

    Dim contenu As String
    Dim contenuUpper As String

    On Error GoTo Fin

    If Not FichierExiste(cheminFichier) Then Exit Function

    contenu = LireFichierVBAImport(cheminFichier)
    contenuUpper = UCase$(contenu)

    If InStr(1, contenuUpper, "SYNCHRONISATION BDD-RF <-> DOSSIER", vbTextCompare) > 0 Then
        FichierImportDangereux = True
        Exit Function
    End If

    If InStr(1, contenuUpper, "PUBLIC SUB EXPORTERVERSDOSSIER", vbTextCompare) > 0 Then
        FichierImportDangereux = True
        Exit Function
    End If

    If InStr(1, contenuUpper, "PUBLIC SUB REIMPORTERDANSEXCEL", vbTextCompare) > 0 Then
        FichierImportDangereux = True
        Exit Function
    End If

    If InStr(1, contenuUpper, "PRIVATE FUNCTION DEMANDERCHEMINREPO", vbTextCompare) > 0 Then
        FichierImportDangereux = True
        Exit Function
    End If

    If InStr(1, contenuUpper, "UF_GITSYNC", vbTextCompare) > 0 _
       And InStr(1, contenuUpper, "INITIALISERGITSYNC", vbTextCompare) > 0 Then
        FichierImportDangereux = True
        Exit Function
    End If

Fin:

End Function

' =============================================
' EstModulePrioritaire
' =============================================
Private Function EstModulePrioritaire(ByVal nomComp As String) As Boolean

    Dim listePrio() As String
    Dim i As Long

    listePrio = Split(MODULES_PRIORITAIRES, ",")

    For i = LBound(listePrio) To UBound(listePrio)
        If StrComp(Trim$(listePrio(i)), nomComp, vbTextCompare) = 0 Then
            EstModulePrioritaire = True
            Exit Function
        End If
    Next i

    EstModulePrioritaire = False

End Function

' =============================================
' EcrireTexteUTF8
' =============================================
Private Sub EcrireTexteUTF8(ByVal chemin As String, ByVal contenu As String)

    Dim stm As Object

    SupprimerFichierSiExiste chemin, "EcrireTexteUTF8"

    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText contenu
        .SaveToFile chemin, 2
        .Close
    End With

    Set stm = Nothing

End Sub

' =============================================
' LireFichierTexteUTF8
' =============================================
Private Function LireFichierTexteUTF8(ByVal chemin As String) As String

    Dim stm As Object

    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile chemin
        LireFichierTexteUTF8 = .ReadText
        .Close
    End With

    Set stm = Nothing

End Function

' =============================================
' LireFichierTexteSysteme
' =============================================
Private Function LireFichierTexteSysteme(ByVal chemin As String) As String

    Dim fso As Object
    Dim ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(chemin, 1, False, -2)

    LireFichierTexteSysteme = ts.ReadAll

    ts.Close
    Set ts = Nothing
    Set fso = Nothing

End Function

' =============================================
' LireFichierVBAImport
' =============================================
Private Function LireFichierVBAImport(ByVal chemin As String) As String

    If FichierEstUTF8(chemin) Then
        LireFichierVBAImport = LireFichierTexteUTF8(chemin)
    Else
        LireFichierVBAImport = LireFichierTexteSysteme(chemin)
    End If

    LireFichierVBAImport = NettoyerBOMTexte(LireFichierVBAImport)

End Function

' =============================================
' FichierEstUTF8
' =============================================
Private Function FichierEstUTF8(ByVal chemin As String) As Boolean

    Dim bytes() As Byte
    Dim f As Integer
    Dim taille As Long
    Dim i As Long
    Dim b As Long
    Dim nbSuite As Long

    On Error GoTo Fin

    If Not FichierExiste(chemin) Then Exit Function

    taille = FileLen(chemin)
    If taille = 0 Then Exit Function

    ReDim bytes(1 To taille)

    f = FreeFile
    Open chemin For Binary Access Read As #f
    Get #f, , bytes
    Close #f

    If taille >= 3 Then
        If bytes(1) = &HEF And bytes(2) = &HBB And bytes(3) = &HBF Then
            FichierEstUTF8 = True
            Exit Function
        End If
    End If

    i = 1

    Do While i <= taille

        b = bytes(i)

        If b < &H80 Then

            i = i + 1

        ElseIf b >= &HC2 And b <= &HDF Then

            nbSuite = 1
            If Not SuitesUTF8Valides(bytes, i, taille, nbSuite) Then Exit Function
            i = i + nbSuite + 1

        ElseIf b >= &HE0 And b <= &HEF Then

            nbSuite = 2
            If Not SuitesUTF8Valides(bytes, i, taille, nbSuite) Then Exit Function
            i = i + nbSuite + 1

        ElseIf b >= &HF0 And b <= &HF4 Then

            nbSuite = 3
            If Not SuitesUTF8Valides(bytes, i, taille, nbSuite) Then Exit Function
            i = i + nbSuite + 1

        Else

            Exit Function

        End If

    Loop

    FichierEstUTF8 = True
    Exit Function

Fin:
    On Error Resume Next
    If f <> 0 Then Close #f
    FichierEstUTF8 = False

End Function

' =============================================
' SuitesUTF8Valides
' =============================================
Private Function SuitesUTF8Valides(ByRef bytes() As Byte, _
                                   ByVal posDepart As Long, _
                                   ByVal taille As Long, _
                                   ByVal nbSuite As Long) As Boolean

    Dim j As Long
    Dim b As Long

    If posDepart + nbSuite > taille Then Exit Function

    For j = 1 To nbSuite

        b = bytes(posDepart + j)

        If b < &H80 Or b > &HBF Then
            SuitesUTF8Valides = False
            Exit Function
        End If

    Next j

    SuitesUTF8Valides = True

End Function

' =============================================
' CopierFichierBinaire
' =============================================
Private Sub CopierFichierBinaire(ByVal cheminSource As String, ByVal cheminCible As String)

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next

    If fso.FileExists(cheminCible) Then fso.DeleteFile cheminCible, True

    If Err.Number <> 0 Then
        JournaliserIO "CopierFichierBinaire", cheminCible, Err.Number, Err.description
        Err.Clear
    End If

    On Error GoTo 0

    fso.CopyFile cheminSource, cheminCible, True

    Set fso = Nothing

End Sub

' =============================================
' CreerDossierSiAbsent
' =============================================
Private Sub CreerDossierSiAbsent(ByVal chemin As String)

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(chemin) Then
        fso.CreateFolder chemin
    End If

    Set fso = Nothing

End Sub

' =============================================
' FichierExiste
' =============================================
Private Function FichierExiste(ByVal chemin As String) As Boolean

    Dim fso As Object

    If Len(Trim$(chemin)) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    FichierExiste = fso.FileExists(chemin)
    Set fso = Nothing

End Function

' =============================================
' DossierExiste
' =============================================
Private Function DossierExiste(ByVal chemin As String) As Boolean

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    DossierExiste = fso.FolderExists(chemin)
    Set fso = Nothing

End Function

' =============================================
' NomSansExtension
' =============================================
Private Function nomSansExtension(ByVal nomFichier As String) As String

    Dim pos As Long

    pos = InStrRev(nomFichier, ".")

    If pos > 0 Then
        nomSansExtension = Left$(nomFichier, pos - 1)
    Else
        nomSansExtension = nomFichier
    End If

End Function

' =============================================
' SupprimerFichierSiExiste
' =============================================
Private Sub SupprimerFichierSiExiste(ByVal chemin As String, ByVal contexte As String)

    If Not FichierExiste(chemin) Then Exit Sub

    On Error Resume Next

    Kill chemin

    If Err.Number <> 0 Then
        JournaliserIO contexte, chemin, Err.Number, Err.description
        Err.Clear
    End If

    On Error GoTo 0

End Sub

' =============================================
' JournaliserIO
' =============================================
Private Sub JournaliserIO(ByVal contexte As String, _
                          ByVal chemin As String, _
                          ByVal numero As Long, _
                          ByVal description As String)

    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [zRFGitSync] " & contexte & _
                " | " & chemin & " | Err " & CStr(numero) & " - " & description

End Sub

' =============================================
' Journal disque spécial crash import
' =============================================
Private Sub JournaliserImportCrash(ByVal etape As String, Optional ByVal detail As String = "")

    Dim cheminLog As String
    Dim numFichier As Integer

    On Error Resume Next

    cheminLog = Environ$("TEMP") & "\BDD_RF_IMPORT_DEBUG.txt"
    numFichier = FreeFile

    Open cheminLog For Append As #numFichier
    Print #numFichier, Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
                       " | " & etape & _
                       IIf(Len(detail) > 0, " | " & detail, "")
    Close #numFichier

    On Error GoTo 0

End Sub

' =============================================
' ViderJournalImportCrash
' =============================================
Private Sub ViderJournalImportCrash()

    Dim cheminLog As String

    On Error Resume Next

    cheminLog = Environ$("TEMP") & "\BDD_RF_IMPORT_DEBUG.txt"

    If FichierExiste(cheminLog) Then
        Kill cheminLog
    End If

    On Error GoTo 0

End Sub

' =============================================
' VerifierPreRequisGitSync
' =============================================
Public Sub VerifierPreRequisGitSync()

    Dim CheminRepo As String
    Dim cheminCodex As String
    Dim cheminImport As String
    Dim cheminTxt As String
    Dim message As String

    CheminRepo = DemanderCheminRepo("CHECK")
    If Len(CheminRepo) = 0 Then Exit Sub

    cheminCodex = CheminRepo & "\src"
    cheminImport = CheminRepo & "\vba_import"
    cheminTxt = CheminRepo & "\src_txt"

    CreerDossierSiAbsent CheminRepo
    CreerDossierSiAbsent cheminCodex
    CreerDossierSiAbsent cheminImport
    CreerDossierSiAbsent cheminTxt

    message = "=== Pré-check GitSync BDD-RF ===" & vbCrLf & vbCrLf

    If ProjetVBAccessible() Then
        message = message & "[OK] Accès VBProject (" & _
                  ThisWorkbook.VBProject.VBComponents.Count & " composants)" & vbCrLf
    Else
        message = message & "[KO] Accès VBProject" & vbCrLf
    End If

    message = message & IIf(TesterEcritureDossier(CheminRepo), "[OK]", "[KO]") & _
              " Dossier repo : " & CheminRepo & vbCrLf

    message = message & IIf(TesterEcritureDossier(cheminCodex), "[OK]", "[KO]") & _
              " Dossier src : " & cheminCodex & vbCrLf

    message = message & IIf(TesterEcritureDossier(cheminImport), "[OK]", "[KO]") & _
              " Dossier vba_import : " & cheminImport & vbCrLf

    message = message & IIf(TesterEcritureDossier(cheminTxt), "[OK]", "[KO]") & _
              " Dossier src_txt : " & cheminTxt & vbCrLf & vbCrLf

    message = message & "Fichiers dans vba_import :" & vbCrLf & _
              "  .bas : " & CompterFichiers(cheminImport, "*.bas") & vbCrLf & _
              "  .cls : " & CompterFichiers(cheminImport, "*.cls") & vbCrLf & _
              "  .frm : " & CompterFichiers(cheminImport, "*.frm") & vbCrLf & vbCrLf

    message = message & "Fichiers dans src_txt :" & vbCrLf & _
              "  .bas.txt : " & CompterFichiers(cheminTxt, "*.bas.txt") & vbCrLf & _
              "  .cls.txt : " & CompterFichiers(cheminTxt, "*.cls.txt") & vbCrLf & _
              "  .frm.txt : " & CompterFichiers(cheminTxt, "*.frm.txt") & vbCrLf & _
              "  .txt : " & CompterFichiers(cheminTxt, "*.txt") & vbCrLf & vbCrLf

    message = message & "Module prioritaire : " & MODULES_PRIORITAIRES

    MsgBox message, vbInformation

End Sub

' =============================================
' CompterFichiers
' =============================================
Private Function CompterFichiers(ByVal dossier As String, ByVal filtre As String) As Long

    Dim n As Long
    Dim f As String

    f = Dir(dossier & "\" & filtre)

    Do While Len(f) > 0
        n = n + 1
        f = Dir
    Loop

    CompterFichiers = n

End Function
' =============================================
' MettreAJourShapeGitSync
' =============================================
Private Sub MettreAJourShapeGitSync(ByVal nomForme As String)

    Dim ws As Worksheet
    Dim shp As Shape
    Dim titre As String
    Dim texteComplet As String

    On Error GoTo Fin

    titre = "Modules"

    texteComplet = titre & vbCrLf & _
                   "Dernier " & IIf(nomForme = "DernierExport", "Export", "Import") & " :" & vbCrLf & _
                   Format(Now, "dd/mm/yyyy à hh:mm")

    Set shp = TrouverShapeDansClasseur(ThisWorkbook, nomForme)

    If shp Is Nothing Then
        MsgBox "Export terminé, mais la forme '" & nomForme & "' est introuvable dans le classeur.", vbExclamation
        Exit Sub
    End If

    With shp.TextFrame
        .Characters.Text = texteComplet

        ' Tout le texte en blanc
        .Characters.Font.Color = RGB(255, 255, 255)

        ' Seulement "Modules" en rose
        .Characters(1, Len(titre)).Font.Color = RGB(229, 158, 221)
    End With

Fin:
    Err.Clear

End Sub

' =============================================
' TrouverShapeDansClasseur
' =============================================
Private Function TrouverShapeDansClasseur(ByVal wb As Workbook, ByVal nomForme As String) As Shape

    Dim ws As Worksheet
    Dim shp As Shape

    For Each ws In wb.Worksheets
        On Error Resume Next
        Set shp = ws.Shapes(nomForme)
        On Error GoTo 0

        If Not shp Is Nothing Then
            Set TrouverShapeDansClasseur = shp
            Exit Function
        End If
    Next ws

End Function

' =============================================
' TesterEcritureDossier
' =============================================
Private Function TesterEcritureDossier(ByVal dossier As String) As Boolean

    Dim cheminTest As String

    cheminTest = dossier & "\__zRFGitSync_write_test__.tmp"

    On Error GoTo ErrHandler

    EcrireTexteUTF8 cheminTest, "ok"
    TesterEcritureDossier = FichierExiste(cheminTest)
    SupprimerFichierSiExiste cheminTest, "TesterEcritureDossier"
    Exit Function

ErrHandler:
    JournaliserIO "TesterEcritureDossier", dossier, Err.Number, Err.description
    Err.Clear
    TesterEcritureDossier = False

End Function



