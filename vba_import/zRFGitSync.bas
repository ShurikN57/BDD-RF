Attribute VB_Name = "zRFGitSync"
Option Explicit

' =============================================
' Synchronisation BDD-RF <-> GitHub / Codex
' - src        = version UTF-8 lisible pour GitHub / Codex
' - vba_import = version native réimportable dans Excel
' =============================================

Private Const DOSSIER_REPO_PC1 As String = "C:\Users\FMF00CDN\Desktop\BDD-RF-GitHub"
Private Const DOSSIER_REPO_PC2 As String = "D:\BDD-RF-GitHub"   ' <-- A ADAPTER

Private Const NOM_THISWORKBOOK As String = "zRFThisWorkbook"
Private Const NOM_FEUILLE_CIBLE As String = "GMC"
Private Const NOM_MODULE_UTILITAIRE As String = "zRFGitSync"

Private Const TYPE_STD_MODULE As Long = 1
Private Const TYPE_CLASS_MODULE As Long = 2
Private Const TYPE_USERFORM As Long = 3
Private Const TYPE_DOCUMENT As Long = 100

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
                    ExporterDocumentModule vbComp, cheminCodex & "\" & vbComp.Name & ".txt"
                    ExporterDocumentModule vbComp, cheminImport & "\" & vbComp.Name & ".txt"
                End If

        End Select

SuiteComposant:
    Next vbComp

    MsgBox "Export terminé :" & vbCrLf & _
           "- GitHub / Codex : " & cheminCodex & vbCrLf & _
           "- Réimport Excel : " & cheminImport, vbInformation

End Sub

' =============================================
' 2. Import depuis le dépôt local
' =============================================
Public Sub ImporterProjetDepuisGitHub()

    On Error GoTo ErrHandler

    Dim cheminRepo As String
    Dim cheminImport As String

    If Not ProjetVBAccessible() Then Exit Sub

    cheminRepo = DossierRepo()
    If Len(cheminRepo) = 0 Then Exit Sub

    cheminImport = DossierImport()

    If Dir(cheminImport, vbDirectory) = "" Then
        MsgBox "Dossier introuvable : " & cheminImport, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ImporterModulesStandards cheminImport
    ImporterClasses cheminImport
    ImporterUserForms cheminImport

    RemplacerCodeDocumentDepuisFichier cheminImport & "\" & NOM_FEUILLE_CIBLE & ".txt", NOM_FEUILLE_CIBLE
    RemplacerCodeDocumentDepuisFichier cheminImport & "\" & NOM_THISWORKBOOK & ".txt", NOM_THISWORKBOOK

SortiePropre:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Erreur import : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

' =============================================
' 3. Exports unitaires
' =============================================
Private Sub ExporterCodeLisibleUTF8(ByVal vbComp As Object, ByVal cheminFinal As String)

    EcrireTexteUTF8 cheminFinal, LireCodeComponent(vbComp, vbComp.Name)

End Sub

Private Sub ExporterDocumentModule(ByVal vbComp As Object, ByVal cheminFinal As String)

    EcrireTexteUTF8 cheminFinal, LireCodeComponent(vbComp, vbComp.Name)

End Sub

Private Sub ExporterComposantNatif(ByVal vbComp As Object, ByVal cheminFinal As String)

    SupprimerFichierSiExiste cheminFinal, "ExporterComposantNatif"
    vbComp.Export cheminFinal

End Sub

Private Sub ExporterUserFormComplet(ByVal vbComp As Object, ByVal cheminCodex As String, ByVal cheminImport As String)

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
' 4. Imports unitaires
' =============================================
Private Sub ImporterModulesStandards(ByVal cheminImport As String)

    Dim fichier As String
    Dim nomComp As String

    fichier = Dir(cheminImport & "\*.bas")

    Do While Len(fichier) > 0
        nomComp = NomSansExtension(fichier)

        If Not DoitEtreIgnoreImport(nomComp) Then
            SupprimerComposantSiExiste nomComp, TYPE_STD_MODULE
            ThisWorkbook.VBProject.VBComponents.Import cheminImport & "\" & fichier
        End If

        fichier = Dir
    Loop

End Sub

Private Sub ImporterClasses(ByVal cheminImport As String)

    Dim fichier As String
    Dim nomComp As String

    fichier = Dir(cheminImport & "\*.cls")

    Do While Len(fichier) > 0
        nomComp = NomSansExtension(fichier)

        If Not DoitEtreIgnoreImport(nomComp) Then
            SupprimerComposantSiExiste nomComp, TYPE_CLASS_MODULE
            ThisWorkbook.VBProject.VBComponents.Import cheminImport & "\" & fichier
        End If

        fichier = Dir
    Loop

End Sub

Private Sub ImporterUserForms(ByVal cheminImport As String)

    Dim fichier As String
    Dim nomComp As String

    fichier = Dir(cheminImport & "\*.frm")

    Do While Len(fichier) > 0
        nomComp = NomSansExtension(fichier)

        If Not DoitEtreIgnoreImport(nomComp) Then
            SupprimerComposantSiExiste nomComp, TYPE_USERFORM
            ThisWorkbook.VBProject.VBComponents.Import cheminImport & "\" & fichier
        End If

        fichier = Dir
    Loop

End Sub

Private Sub RemplacerCodeDocumentDepuisFichier(ByVal chemin As String, ByVal nomComp As String)

    Dim vbComp As Object
    Dim contenu As String

    If Not FichierExiste(chemin) Then Exit Sub

    Set vbComp = ThisWorkbook.VBProject.VBComponents(nomComp)
    contenu = LireFichierTexteUTF8(chemin)
    contenu = NettoyerEnteteExport(contenu)

    With vbComp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(contenu) > 0 Then .AddFromString contenu
    End With

End Sub

' =============================================
' 5. Helpers code / fichiers
' =============================================
Private Function LireCodeComponent(ByVal vbComp As Object, ByVal nomComp As String) As String

    Dim codeTexte As String

    If vbComp.CodeModule.CountOfLines > 0 Then
        codeTexte = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
    Else
        codeTexte = ""
    End If

    LireCodeComponent = "Attribute VB_Name = """ & nomComp & """" & vbCrLf & codeTexte

End Function

Private Function NettoyerEnteteExport(ByVal contenu As String) As String

    Dim lignes() As String
    Dim i As Long
    Dim resultat As String
    Dim ligne As String

    If Len(contenu) = 0 Then Exit Function

    lignes = Split(Replace(contenu, vbCrLf, vbLf), vbLf)

    For i = LBound(lignes) To UBound(lignes)
        ligne = lignes(i)

        If Left$(Trim$(ligne), 17) <> "Attribute VB_Name" Then
            If resultat = "" Then
                resultat = ligne
            Else
                resultat = resultat & vbCrLf & ligne
            End If
        End If
    Next i

    NettoyerEnteteExport = resultat

End Function

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
    If Not fso.FolderExists(chemin) Then fso.CreateFolder chemin
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

Private Function NomSansExtension(ByVal nomFichier As String) As String

    Dim pos As Long

    pos = InStrRev(nomFichier, ".")
    If pos > 0 Then
        NomSansExtension = Left$(nomFichier, pos - 1)
    Else
        NomSansExtension = nomFichier
    End If

End Function

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

Private Sub SupprimerFichierSiExiste(ByVal chemin As String, ByVal contexte As String)

    If Len(Dir(chemin)) = 0 Then Exit Sub

    On Error Resume Next
    Kill chemin
    If Err.Number <> 0 Then
        JournaliserIO contexte, chemin, Err.Number, Err.description
        Err.Clear
    End If
    On Error GoTo 0

End Sub

Private Sub JournaliserIO(ByVal contexte As String, ByVal chemin As String, ByVal numero As Long, ByVal description As String)

    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [zRFGitSync] " & contexte & _
                " | " & chemin & " | Err " & CStr(numero) & " - " & description

End Sub

' =============================================
' 6. Self-check
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

    message = "Pré-check GitSync" & vbCrLf & vbCrLf

    If ProjetVBAccessible() Then
        message = message & "[OK] Accès VBProject" & vbCrLf
    Else
        message = message & "[KO] Accès VBProject" & vbCrLf
    End If

    CreerDossierSiAbsent cheminRepo
    CreerDossierSiAbsent cheminCodex
    CreerDossierSiAbsent cheminImport

    message = message & IIf(TesterEcritureDossier(cheminRepo), "[OK]", "[KO]") & " Dossier repo: " & cheminRepo & vbCrLf
    message = message & IIf(TesterEcritureDossier(cheminCodex), "[OK]", "[KO]") & " Dossier src: " & cheminCodex & vbCrLf
    message = message & IIf(TesterEcritureDossier(cheminImport), "[OK]", "[KO]") & " Dossier import: " & cheminImport & vbCrLf & vbCrLf
    message = message & "Chemins testés :" & vbCrLf & _
              "PC1 : " & DOSSIER_REPO_PC1 & vbCrLf & _
              "PC2 : " & DOSSIER_REPO_PC2

    MsgBox message, vbInformation

End Sub

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

' =============================================
' 7. Helpers VBProject
' =============================================
Private Function ProjetVBAccessible() As Boolean

    On Error GoTo ErrHandler

    Dim n As Long
    n = ThisWorkbook.VBProject.VBComponents.Count
    ProjetVBAccessible = True
    Exit Function

ErrHandler:
    MsgBox "Accès refusé au projet VBA." & vbCrLf & _
           "Active l'option :" & vbCrLf & _
           "Fichier > Options > Centre de gestion de la confidentialité > Paramètres des macros > " & _
           "Accès approuvé au modèle d'objet du projet VBA.", vbExclamation

End Function

Private Function DoitEtreIgnoreExport(ByVal nomComp As String) As Boolean

    DoitEtreIgnoreExport = False

End Function

Private Function DoitEtreIgnoreImport(ByVal nomComp As String) As Boolean

    Select Case LCase$(nomComp)
        Case LCase$(NOM_MODULE_UTILITAIRE)
            DoitEtreIgnoreImport = True
    End Select

End Function

Private Sub SupprimerComposantSiExiste(ByVal nomComp As String, ByVal typeAttendu As Long)

    Dim vbComp As Object

    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(nomComp)
    On Error GoTo 0

    If vbComp Is Nothing Then Exit Sub
    If vbComp.Type = TYPE_DOCUMENT Then Exit Sub
    If vbComp.Type <> typeAttendu Then Exit Sub
    If DoitEtreIgnoreImport(vbComp.Name) Then Exit Sub

    ThisWorkbook.VBProject.VBComponents.Remove vbComp

End Sub

