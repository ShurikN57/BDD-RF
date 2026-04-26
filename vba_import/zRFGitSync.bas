Attribute VB_Name = "zRFGitSync"
Option Explicit

' =============================================
' Synchronisation BDD-DOC <-> GitHub / Codex
' - src        = version UTF-8 lisible pour GitHub / Codex
' - vba_import = version native destinée au réimport Excel
'
' Version sécurisée :
' - Export complet vers src + vba_import
' - Import depuis vba_import uniquement
' - zRFConstance importé/remplacé en premier
' - Les modules existants ne sont plus supprimés/recréés :
'   leur code est remplacé en place.
' - Cela évite les doublons zRFConstance1, BoutonSauvegarde1, etc.
' - Les composants absents sont importés normalement.
' - zRFGitSync est exporté mais jamais réimporté par-dessus lui-même.
' =============================================

Private Const DOSSIER_REPO_PC1 As String = "C:\Users\FMF00CDN\Desktop\BDD-RF-GitHub"
Private Const DOSSIER_REPO_PC2 As String = "C:\Users\micho\OneDrive\Desktop\RF"

Private Const NOM_THISWORKBOOK As String = "zRFThisWorkbook"
Private Const NOM_FEUILLE_CIBLE As String = "GMC"
Private Const NOM_MODULE_UTILITAIRE As String = "zRFGitSync"

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
Public Sub ExporterVersGitHub()
    ExporterProjetVersGitHubEtImportExcel
End Sub

Public Sub ReimporterDansExcel()
    ImporterProjetDepuisGitHub
End Sub

' =============================================
' 1. Export complet
' =============================================
Public Sub ExporterProjetVersGitHubEtImportExcel()

    Dim vbProj As Object
    Dim vbComp As Object
    Dim cheminRepo As String
    Dim cheminCodex As String
    Dim cheminImport As String

    If Not ProjetVBAccessible() Then Exit Sub

    cheminRepo = DossierRepo()
    If Len(cheminRepo) = 0 Then Exit Sub

    cheminCodex = DossierCodex()
    cheminImport = DossierImport()

    CreerDossierSiAbsent cheminRepo
    CreerDossierSiAbsent cheminCodex
    CreerDossierSiAbsent cheminImport

    Set vbProj = ThisWorkbook.VBProject

    For Each vbComp In vbProj.VBComponents

        If DoitEtreIgnoreExport(vbComp.Name) Then GoTo SuiteComposant

        Select Case vbComp.Type

            Case TYPE_STD_MODULE
                ExporterCodeLisibleUTF8 vbComp, cheminCodex & "\" & vbComp.Name & ".bas"
                ExporterComposantNatif vbComp, cheminImport & "\" & vbComp.Name & ".bas"

            Case TYPE_CLASS_MODULE
                ExporterCodeLisibleUTF8 vbComp, cheminCodex & "\" & vbComp.Name & ".cls"
                ExporterComposantNatif vbComp, cheminImport & "\" & vbComp.Name & ".cls"

            Case TYPE_USERFORM
                ExporterUserFormComplet vbComp, cheminCodex, cheminImport

            Case TYPE_DOCUMENT
                If vbComp.Name = NOM_THISWORKBOOK Or vbComp.Name = NOM_FEUILLE_CIBLE Then
                    ExporterCodeLisibleUTF8 vbComp, cheminCodex & "\" & vbComp.Name & ".txt"
                    ExporterDocumentModuleBrut vbComp, cheminImport & "\" & vbComp.Name & ".txt"
                End If

        End Select

SuiteComposant:
    Next vbComp

    MsgBox "Export terminé :" & vbCrLf & _
           "- GitHub / Codex : " & cheminCodex & vbCrLf & _
           "- Réimport Excel : " & cheminImport, vbInformation

End Sub

' =============================================
' 2. Import depuis vba_import
' =============================================
Public Sub ImporterProjetDepuisGitHub()

    On Error GoTo ErrHandler

    Dim cheminRepo As String
    Dim cheminImport As String
    Dim liste() As ComposantAImporter
    Dim nbListe As Long
    Dim nbImportes As Long
    Dim nbErreurs As Long
    Dim journalErreurs As String
    Dim msg As String
    Dim i As Long

    If Not ProjetVBAccessible() Then Exit Sub

    cheminRepo = DossierRepo()
    If Len(cheminRepo) = 0 Then Exit Sub

    cheminImport = DossierImport()

    If Not DossierExiste(cheminImport) Then
        MsgBox "Dossier introuvable : " & cheminImport, vbExclamation
        Exit Sub
    End If

    If Not VerifierAvantImport(cheminImport) Then
        MsgBox "Import annulé.", vbInformation
        Exit Sub
    End If

    nbListe = 0
    CollecterModulesPrioritaires cheminImport, liste, nbListe
    CollecterFichiersParExtension cheminImport, "*.bas", TYPE_STD_MODULE, liste, nbListe
    CollecterFichiersParExtension cheminImport, "*.cls", TYPE_CLASS_MODULE, liste, nbListe
    CollecterFichiersParExtension cheminImport, "*.frm", TYPE_USERFORM, liste, nbListe

    If nbListe = 0 Then
        MsgBox "Aucun composant à importer.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Import / remplacement des modules prioritaires
    For i = 1 To nbListe
        If liste(i).estPrioritaire Then
            If ImporterOuRemplacerComposant(liste(i).cheminFichier, liste(i).nomComp, liste(i).typeComp) Then
                nbImportes = nbImportes + 1
            Else
                nbErreurs = nbErreurs + 1
                journalErreurs = journalErreurs & "  - " & liste(i).nomComp & " (prioritaire non importé)" & vbCrLf
            End If
        End If
    Next i

    ' Import / remplacement des autres composants
    For i = 1 To nbListe
        If Not liste(i).estPrioritaire Then
            If ImporterOuRemplacerComposant(liste(i).cheminFichier, liste(i).nomComp, liste(i).typeComp) Then
                nbImportes = nbImportes + 1
            Else
                nbErreurs = nbErreurs + 1
                journalErreurs = journalErreurs & "  - " & liste(i).nomComp & " (non importé)" & vbCrLf
            End If
        End If
    Next i

    ' Modules document : Base / ThisWorkbook
    RemplacerCodeDocumentDepuisFichier cheminImport & "\" & NOM_FEUILLE_CIBLE & ".txt", NOM_FEUILLE_CIBLE
    RemplacerCodeDocumentDepuisFichier cheminImport & "\" & NOM_THISWORKBOOK & ".txt", NOM_THISWORKBOOK

    msg = "Import terminé." & vbCrLf & vbCrLf & _
          "Composants importés/remplacés : " & nbImportes & vbCrLf & _
          "Erreurs : " & nbErreurs

    If nbErreurs > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Détail :" & vbCrLf & journalErreurs
    End If

    MsgBox msg, IIf(nbErreurs > 0, vbExclamation, vbInformation)

SortiePropre:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Erreur import critique : " & Err.Number & " - " & Err.description, vbCritical
    Resume SortiePropre

End Sub

' =============================================
' 2-alpha. Vérification avant import
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

Private Sub VerifierExtensionDryRun(ByVal cheminImport As String, _
                                    ByVal filtre As String, _
                                    ByRef listeTrouves As String, _
                                    ByRef listeManquants As String, _
                                    ByRef nbTrouves As Long, _
                                    ByRef nbManquants As Long)

    Dim fichier As String
    Dim nomComp As String

    fichier = Dir(cheminImport & "\" & filtre)

    Do While Len(fichier) > 0
        nomComp = nomSansExtension(fichier)

        If Not DoitEtreIgnoreImport(nomComp) And Not EstModulePrioritaire(nomComp) Then
            nbTrouves = nbTrouves + 1
            listeTrouves = listeTrouves & "  " & fichier & vbCrLf
        End If

        fichier = Dir
    Loop

End Sub

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
' 2-bis. Collecte des composants à importer
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

Private Sub CollecterFichiersParExtension(ByVal cheminImport As String, _
                                          ByVal filtre As String, _
                                          ByVal typeComp As Long, _
                                          ByRef liste() As ComposantAImporter, _
                                          ByRef nbListe As Long)

    Dim fichiers() As String
    Dim nbFichiers As Long
    Dim fichier As String
    Dim nomComp As String
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

        If DoitEtreIgnoreImport(nomComp) Then GoTo SuiteFichier
        If EstModulePrioritaire(nomComp) Then GoTo SuiteFichier

        nbListe = nbListe + 1
        ReDim Preserve liste(1 To nbListe)
        liste(nbListe).cheminFichier = cheminImport & "\" & fichiers(i)
        liste(nbListe).nomComp = nomComp
        liste(nbListe).typeComp = typeComp
        liste(nbListe).estPrioritaire = False

SuiteFichier:
    Next i

End Sub

' =============================================
' 2-ter. Import ou remplacement sécurisé
' =============================================
Private Function ImporterOuRemplacerComposant(ByVal cheminFichier As String, _
                                              ByVal nomComp As String, _
                                              ByVal typeComp As Long) As Boolean

    On Error GoTo ErrHandler

    Dim vbCompExistant As Object
    Dim vbCompImporte As Object
    Dim nomImporte As String
    Dim contenu As String

    If Not FichierExiste(cheminFichier) Then
        JournaliserIO "ImporterOuRemplacerComposant", cheminFichier, 0, "Fichier introuvable"
        ImporterOuRemplacerComposant = False
        Exit Function
    End If

    Set vbCompExistant = ObtenirComposantVBA(nomComp)

    If Not vbCompExistant Is Nothing Then

        If vbCompExistant.Type <> typeComp Then
            JournaliserIO "ImporterOuRemplacerComposant", nomComp, 0, _
                          "Type incohérent. Attendu=" & typeComp & " / Réel=" & vbCompExistant.Type
            ImporterOuRemplacerComposant = False
            Exit Function
        End If

        If vbCompExistant.Type = TYPE_DOCUMENT Then
            JournaliserIO "ImporterOuRemplacerComposant", nomComp, 0, "Module document non traité ici"
            ImporterOuRemplacerComposant = False
            Exit Function
        End If

        contenu = LireFichierTexteSysteme(cheminFichier)
        contenu = NettoyerEnteteExport(contenu)

        With vbCompExistant.CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            If Len(contenu) > 0 Then .AddFromString contenu
        End With

        ImporterOuRemplacerComposant = True
        Exit Function

    End If

    Set vbCompImporte = ThisWorkbook.VBProject.VBComponents.Import(cheminFichier)

    If vbCompImporte Is Nothing Then
        JournaliserIO "ImporterOuRemplacerComposant", nomComp, 0, "Import sans composant retourné"
        ImporterOuRemplacerComposant = False
        Exit Function
    End If

    nomImporte = vbCompImporte.Name

    If StrComp(nomImporte, nomComp, vbTextCompare) <> 0 Then

        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Remove vbCompImporte
        On Error GoTo ErrHandler

        JournaliserIO "ImporterOuRemplacerComposant", cheminFichier, 0, _
                      "Nom incohérent. Attendu=" & nomComp & " / Obtenu=" & nomImporte

        ImporterOuRemplacerComposant = False
        Exit Function

    End If

    ImporterOuRemplacerComposant = True
    Exit Function

ErrHandler:
    JournaliserIO "ImporterOuRemplacerComposant", cheminFichier, Err.Number, Err.description
    ImporterOuRemplacerComposant = False

End Function

' =============================================
' 3. Exports unitaires
' =============================================
Private Sub ExporterCodeLisibleUTF8(ByVal vbComp As Object, ByVal cheminFinal As String)

    EcrireTexteUTF8 cheminFinal, LireCodeAvecHeader(vbComp)

End Sub

Private Sub ExporterDocumentModuleBrut(ByVal vbComp As Object, ByVal cheminFinal As String)

    EcrireTexteUTF8 cheminFinal, LireCodeBrut(vbComp)

End Sub

Private Sub ExporterComposantNatif(ByVal vbComp As Object, ByVal cheminFinal As String)

    SupprimerFichierSiExiste cheminFinal, "ExporterComposantNatif"
    vbComp.Export cheminFinal

End Sub

Private Sub ExporterUserFormComplet(ByVal vbComp As Object, _
                                    ByVal cheminCodex As String, _
                                    ByVal cheminImport As String)

    Dim cheminImportFrm As String
    Dim cheminImportFrx As String
    Dim cheminTempFrm As String
    Dim cheminTempFrx As String
    Dim cheminCodexFrm As String
    Dim cheminCodexFrx As String
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

    SupprimerFichierSiExiste cheminTempFrm, "ExporterUserFormComplet"
    SupprimerFichierSiExiste cheminTempFrx, "ExporterUserFormComplet"
    SupprimerFichierSiExiste cheminCodexFrm, "ExporterUserFormComplet"
    SupprimerFichierSiExiste cheminCodexFrx, "ExporterUserFormComplet"

    vbComp.Export cheminTempFrm

    contenuFrm = LireFichierTexteSysteme(cheminTempFrm)
    EcrireTexteUTF8 cheminCodexFrm, contenuFrm

    If FichierExiste(cheminTempFrx) Then
        CopierFichierBinaire cheminTempFrx, cheminCodexFrx
    End If

    SupprimerFichierSiExiste cheminTempFrm, "ExporterUserFormComplet"
    SupprimerFichierSiExiste cheminTempFrx, "ExporterUserFormComplet"

End Sub

' =============================================
' 4. Document modules
' =============================================
Private Sub RemplacerCodeDocumentDepuisFichier(ByVal chemin As String, ByVal nomComp As String)

    Dim vbComp As Object
    Dim contenu As String

    If Not FichierExiste(chemin) Then Exit Sub

    Set vbComp = ObtenirComposantVBA(nomComp)

    If vbComp Is Nothing Then
        JournaliserIO "RemplacerCodeDocumentDepuisFichier", nomComp, 0, "Composant document introuvable"
        Exit Sub
    End If

    If vbComp.Type <> TYPE_DOCUMENT Then
        JournaliserIO "RemplacerCodeDocumentDepuisFichier", nomComp, 0, "Le composant n'est pas un module document"
        Exit Sub
    End If

    contenu = LireFichierTexteUTF8(chemin)
    contenu = NettoyerEnteteExport(contenu)

    With vbComp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(contenu) > 0 Then .AddFromString contenu
    End With

End Sub

' =============================================
' 5. Helpers code
' =============================================
Private Function LireCodeAvecHeader(ByVal vbComp As Object) As String

    LireCodeAvecHeader = "Attribute VB_Name = """ & vbComp.Name & """" & vbCrLf & LireCodeBrut(vbComp)

End Function

Private Function LireCodeBrut(ByVal vbComp As Object) As String

    If vbComp.CodeModule.CountOfLines > 0 Then
        LireCodeBrut = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
    Else
        LireCodeBrut = ""
    End If

End Function

Private Function NettoyerEnteteExport(ByVal contenu As String) As String

    Dim lignes() As String
    Dim i As Long
    Dim resultat As String
    Dim ligne As String
    Dim ligneTrimmed As String
    Dim dansEntete As Boolean
    Dim attributsCommences As Boolean

    If Len(contenu) = 0 Then Exit Function

    contenu = Replace(contenu, ChrW$(&HFEFF), "")
    lignes = Split(Replace(contenu, vbCrLf, vbLf), vbLf)

    dansEntete = True
    attributsCommences = False

    For i = LBound(lignes) To UBound(lignes)

        ligne = lignes(i)
        ligneTrimmed = Trim$(ligne)

        If dansEntete Then

            If Left$(ligneTrimmed, 9) = "Attribute" Then
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

Private Function LigneSembleEtreDebutCode(ByVal ligneTrimmed As String) As Boolean

    If Len(ligneTrimmed) = 0 Then Exit Function

    LigneSembleEtreDebutCode = _
        Left$(ligneTrimmed, 6) = "Option" Or _
        Left$(ligneTrimmed, 6) = "Public" Or _
        Left$(ligneTrimmed, 7) = "Private" Or _
        Left$(ligneTrimmed, 3) = "Dim" Or _
        Left$(ligneTrimmed, 5) = "Const" Or _
        Left$(ligneTrimmed, 3) = "Sub" Or _
        Left$(ligneTrimmed, 8) = "Function" Or _
        Left$(ligneTrimmed, 8) = "Property" Or _
        Left$(ligneTrimmed, 4) = "Type" Or _
        Left$(ligneTrimmed, 4) = "Enum" Or _
        Left$(ligneTrimmed, 7) = "Declare" Or _
        Left$(ligneTrimmed, 1) = "'"

End Function

' =============================================
' 6. Helpers VBProject
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

Private Function ObtenirComposantVBA(ByVal nomComp As String) As Object

    On Error Resume Next
    Set ObtenirComposantVBA = ThisWorkbook.VBProject.VBComponents(nomComp)
    On Error GoTo 0

End Function

Private Function ComposantExiste(ByVal nomComp As String) As Boolean

    ComposantExiste = Not (ObtenirComposantVBA(nomComp) Is Nothing)

End Function

Private Function DoitEtreIgnoreExport(ByVal nomComp As String) As Boolean

    DoitEtreIgnoreExport = False

End Function

Private Function DoitEtreIgnoreImport(ByVal nomComp As String) As Boolean

    DoitEtreIgnoreImport = (StrComp(nomComp, NOM_MODULE_UTILITAIRE, vbTextCompare) = 0)

End Function

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
' 7. Helpers fichiers
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

Private Sub CreerDossierSiAbsent(ByVal chemin As String)

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(chemin) Then
        fso.CreateFolder chemin
    End If

    Set fso = Nothing

End Sub

Private Function FichierExiste(ByVal chemin As String) As Boolean

    FichierExiste = (Len(Dir(chemin)) > 0)

End Function

Private Function DossierExiste(ByVal chemin As String) As Boolean

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    DossierExiste = fso.FolderExists(chemin)
    Set fso = Nothing

End Function

Private Function nomSansExtension(ByVal nomFichier As String) As String

    Dim pos As Long

    pos = InStrRev(nomFichier, ".")

    If pos > 0 Then
        nomSansExtension = Left$(nomFichier, pos - 1)
    Else
        nomSansExtension = nomFichier
    End If

End Function

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

Private Sub JournaliserIO(ByVal contexte As String, _
                          ByVal chemin As String, _
                          ByVal numero As Long, _
                          ByVal description As String)

    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [zRFGitSync] " & contexte & _
                " | " & chemin & " | Err " & CStr(numero) & " - " & description

End Sub

' =============================================
' 8. Chemins repo
' =============================================
Private Function DossierRepo() As String

    If DossierExiste(DOSSIER_REPO_PC1) Then
        DossierRepo = NormaliserCheminSansSlash(DOSSIER_REPO_PC1)
        Exit Function
    End If

    If DossierExiste(DOSSIER_REPO_PC2) Then
        DossierRepo = NormaliserCheminSansSlash(DOSSIER_REPO_PC2)
        Exit Function
    End If

    MsgBox "Aucun dossier repo valide trouvé." & vbCrLf & vbCrLf & _
           "PC1 : " & DOSSIER_REPO_PC1 & vbCrLf & _
           "PC2 : " & DOSSIER_REPO_PC2, vbExclamation

    DossierRepo = ""

End Function

Private Function DossierCodex() As String

    Dim cheminRepo As String

    cheminRepo = DossierRepo()
    If Len(cheminRepo) = 0 Then Exit Function

    DossierCodex = cheminRepo & "\src"

End Function

Private Function DossierImport() As String

    Dim cheminRepo As String

    cheminRepo = DossierRepo()
    If Len(cheminRepo) = 0 Then Exit Function

    DossierImport = cheminRepo & "\vba_import"

End Function

Private Function NormaliserCheminSansSlash(ByVal chemin As String) As String

    Dim resultat As String

    resultat = Replace(Trim$(chemin), "/", "\")

    Do While Len(resultat) > 0 And Right$(resultat, 1) = "\"
        resultat = Left$(resultat, Len(resultat) - 1)
    Loop

    NormaliserCheminSansSlash = resultat

End Function

' =============================================
' 9. Self-check
' =============================================
Public Sub VerifierPreRequisGitSync()

    Dim cheminRepo As String
    Dim cheminCodex As String
    Dim cheminImport As String
    Dim message As String

    cheminRepo = DossierRepo()
    If Len(cheminRepo) = 0 Then Exit Sub

    cheminCodex = DossierCodex()
    cheminImport = DossierImport()

    CreerDossierSiAbsent cheminRepo
    CreerDossierSiAbsent cheminCodex
    CreerDossierSiAbsent cheminImport

    message = "=== Pré-check GitSync BDD-RF ===" & vbCrLf & vbCrLf

    If ProjetVBAccessible() Then
        message = message & "[OK] Accès VBProject (" & _
                  ThisWorkbook.VBProject.VBComponents.Count & " composants)" & vbCrLf
    Else
        message = message & "[KO] Accès VBProject" & vbCrLf
    End If

    message = message & IIf(TesterEcritureDossier(cheminRepo), "[OK]", "[KO]") & _
              " Dossier repo : " & cheminRepo & vbCrLf
    message = message & IIf(TesterEcritureDossier(cheminCodex), "[OK]", "[KO]") & _
              " Dossier src : " & cheminCodex & vbCrLf
    message = message & IIf(TesterEcritureDossier(cheminImport), "[OK]", "[KO]") & _
              " Dossier vba_import : " & cheminImport & vbCrLf & vbCrLf

    message = message & "Fichiers dans vba_import :" & vbCrLf & _
              "  .bas : " & CompterFichiers(cheminImport, "*.bas") & vbCrLf & _
              "  .cls : " & CompterFichiers(cheminImport, "*.cls") & vbCrLf & _
              "  .frm : " & CompterFichiers(cheminImport, "*.frm") & vbCrLf & vbCrLf

    message = message & "Module prioritaire : " & MODULES_PRIORITAIRES

    MsgBox message, vbInformation

End Sub

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



