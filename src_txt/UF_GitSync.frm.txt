VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_GitSync 
   Caption         =   "Dossier Modules"
   ClientHeight    =   2100
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   5760
   OleObjectBlob   =   "UF_GitSync_codex.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_GitSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' UF_GitSync
' Sélection du dossier racine repo avant export ou import
'
' Champ txtChemin :
'   = dossier racine du repo
'
' Label lblSousDossiers :
'   = affiche le vrai dossier utilisé selon le mode :
'       EXPORT -> \src
'       IMPORT -> \vba_import
'       CHECK  -> repo racine
' ============================================================

Public Confirmed  As Boolean
Public CheminRepo As String
Public AppKey     As String
Public modeAction As String

' Chemins par défaut injectés par l'appelant
Public DefaultPC1 As String
Public DefaultPC2 As String

Private Const APPKEY_FALLBACK As String = "GitSync"

' ============================================================
' Initialize
' ============================================================
Private Sub UserForm_Initialize()

    Confirmed = False
    CheminRepo = ""
    AppKey = ""
    modeAction = ""

End Sub

' ============================================================
' Activate
' ============================================================
Private Sub UserForm_Activate()

    CentrerUserFormSurMoniteurExcel Me, 0.33

End Sub

' ============================================================
' InitialiserGitSync
' ============================================================
Public Sub InitialiserGitSync(ByVal pAppKey As String, _
                              ByVal pDefaultPC1 As String, _
                              ByVal pDefaultPC2 As String, _
                              Optional ByVal pModeAction As String = "")

    Dim dernierChemin As String
    Dim cheminInitial As String

    Confirmed = False
    CheminRepo = ""

    AppKey = Trim$(pAppKey)
    modeAction = UCase$(Trim$(pModeAction))

    If Len(AppKey) = 0 Then AppKey = APPKEY_FALLBACK

    DefaultPC1 = NormaliserRepoUF(pDefaultPC1)
    DefaultPC2 = NormaliserRepoUF(pDefaultPC2)

    dernierChemin = NormaliserRepoUF(GetSetting(AppKey, "GitSync", "DernierRepo", ""))

    If Len(dernierChemin) > 0 And DossierExisteUF(dernierChemin) Then
        cheminInitial = dernierChemin
    ElseIf DossierExisteUF(DefaultPC1) Then
        cheminInitial = DefaultPC1
    ElseIf DossierExisteUF(DefaultPC2) Then
        cheminInitial = DefaultPC2
    Else
        cheminInitial = ""
    End If

    txtChemin.Text = cheminInitial
    MajLabelSousDossiers

End Sub

' ============================================================
' Parcourir
' ============================================================
Private Sub cmdParcourir_Click()

    Dim dlg As FileDialog
    Dim cheminInitial As String
    Dim cheminChoisi As String

    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)

    cheminInitial = NormaliserRepoUF(txtChemin.Text)

    If DossierExisteUF(cheminInitial) Then
        dlg.InitialFileName = cheminInitial & "\"
    End If

    dlg.Title = "Sélectionner le dossier racine du repo"

    If dlg.Show = -1 Then
        cheminChoisi = NormaliserRepoUF(dlg.SelectedItems(1))
        txtChemin.Text = cheminChoisi
        MajLabelSousDossiers
    End If

    Set dlg = Nothing

End Sub

' ============================================================
' Changement manuel du chemin
' ============================================================
Private Sub txtChemin_Change()

    MajLabelSousDossiers

End Sub

' ============================================================
' Mise à jour du label selon le mode
' ============================================================
Private Sub MajLabelSousDossiers()

    Dim chemin As String

    chemin = NormaliserRepoUF(txtChemin.Text)

    If Len(chemin) = 0 Then
        lblSousDossiers.Caption = "Dossier repo : (chemin non défini)"
        Exit Sub
    End If

    Select Case modeAction

        Case "EXPORT"
            lblSousDossiers.Caption = _
                "Export Codex / GitHub vers :" & vbCrLf & _
                "  " & chemin & "\src"

        Case "IMPORT"
            lblSousDossiers.Caption = _
                "Réimport Excel depuis :" & vbCrLf & _
                "  " & chemin & "\vba_import"

        Case "CHECK"
            lblSousDossiers.Caption = _
                "Vérification repo :" & vbCrLf & _
                "  " & chemin

        Case Else
            lblSousDossiers.Caption = _
                "Dossier repo :" & vbCrLf & _
                "  " & chemin

    End Select

End Sub

' ============================================================
' Valider
' ============================================================
Private Sub cmdValider_Click()

    Dim chemin As String

    chemin = NormaliserRepoUF(txtChemin.Text)

    If Len(chemin) = 0 Then
        MsgBox "Veuillez saisir ou sélectionner le dossier racine du repo.", vbExclamation
        txtChemin.SetFocus
        Exit Sub
    End If

    If Not DossierExisteUF(chemin) Then
        MsgBox "Dossier introuvable :" & vbCrLf & chemin, vbExclamation
        txtChemin.SetFocus
        Exit Sub
    End If

    txtChemin.Text = chemin

    CheminRepo = chemin
    Confirmed = True

    If Len(AppKey) = 0 Then AppKey = APPKEY_FALLBACK

    SaveSetting AppKey, "GitSync", "DernierRepo", chemin

    Me.Hide

End Sub

' ============================================================
' Annuler
' ============================================================
Private Sub cmdAnnuler_Click()

    Confirmed = False
    CheminRepo = ""
    Me.Hide

End Sub

' ============================================================
' Fermeture avec la croix
' ============================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Confirmed = False
        CheminRepo = ""
        Me.Hide
        Cancel = True
    End If

End Sub

' ============================================================
' DossierExisteUF
' ============================================================
Private Function DossierExisteUF(ByVal chemin As String) As Boolean

    Dim fso As Object

    chemin = NormaliserUF(chemin)

    If Len(chemin) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    DossierExisteUF = fso.FolderExists(chemin)
    Set fso = Nothing

End Function

' ============================================================
' NormaliserUF
' ============================================================
Private Function NormaliserUF(ByVal chemin As String) As String

    Dim s As String

    s = Replace(Trim$(chemin), "/", "\")

    Do While Len(s) > 0 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop

    NormaliserUF = s

End Function

' ============================================================
' NormaliserRepoUF
'
' Si l'utilisateur ou le registre renvoie :
'   C:\...\BDD-DOC\src
' ou
'   C:\...\BDD-DOC\vba_import
'
' alors on revient automatiquement à :
'   C:\...\BDD-DOC
' ============================================================
Private Function NormaliserRepoUF(ByVal chemin As String) As String

    Dim s As String
    Dim sLower As String

    s = NormaliserUF(chemin)
    sLower = LCase$(s)

    If Len(sLower) >= 4 Then
        If Right$(sLower, 4) = "\src" Then
            s = Left$(s, Len(s) - 4)
        End If
    End If

    sLower = LCase$(s)

    If Len(sLower) >= 11 Then
        If Right$(sLower, 11) = "\vba_import" Then
            s = Left$(s, Len(s) - 11)
        End If
    End If

    NormaliserRepoUF = NormaliserUF(s)

End Function



