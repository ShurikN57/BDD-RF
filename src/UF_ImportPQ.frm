VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ImportPQ 
   Caption         =   "Import PowerQuerry ->Base"
   ClientHeight    =   4515
   ClientLeft      =   75
   ClientTop       =   270
   ClientWidth     =   5535
   OleObjectBlob   =   "UF_ImportPQ_codex.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_ImportPQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' UF_ImportPQ
' Sélection classeur source / onglet source
' Classeur cible et onglet cible = ThisWorkbook / SHEET_MAIN (fixes)
' ============================================================

Public Confirmed   As Boolean
Public nomClasseur As String
Public nomOnglet   As String

' Valeurs par défaut injectées par l'appelant
Public DefaultClasseur As String
Public DefaultOnglet   As String

' ============================================================
' Initialize
' ============================================================
Private Sub UserForm_Initialize()

    Confirmed = False
    nomClasseur = ""
    nomOnglet = ""

End Sub

' ============================================================
' Activate
' ============================================================
Private Sub UserForm_Activate()
    CentrerUserFormSurMoniteurExcel Me, 0.33
End Sub

' ============================================================
' InitialiserImportPQ
' Appelé depuis zDocImportPowerQuery via New UF_ImportPQ
' ============================================================
Public Sub InitialiserImportPQ(ByVal pDefaultClasseur As String, _
                                ByVal pDefaultOnglet As String)

    DefaultClasseur = pDefaultClasseur
    DefaultOnglet = pDefaultOnglet

    ' --- Peuplement ComboBox classeur source uniquement ---
    PeuplerClasseurs

    ' --- Valeur par défaut classeur ---
    If DefaultClasseur <> "" Then
        cbClasseur.Text = DefaultClasseur
    End If

    ' --- Peuplement onglets si classeur reconnu ---
    PeuplerOnglets cbClasseur.Text

    ' --- Valeur par défaut onglet ---
    If DefaultOnglet <> "" Then
        cbOnglet.Text = DefaultOnglet
    End If

    ' --- Cible forcée : classeur courant / feuille principale ---
    cbClasseurCible.Value = nomSansExtension(ThisWorkbook.Name)
    cbOngletCible.Value = SHEET_MAIN

    cbClasseurCible.Enabled = False
    cbOngletCible.Enabled = False

End Sub

' ============================================================
' PeuplerClasseurs
' Liste tous les classeurs Excel ouverts sauf ThisWorkbook
' ============================================================
Private Sub PeuplerClasseurs()

    Dim wb As Workbook
    Dim nomSansExt As String

    cbClasseur.Clear

    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            nomSansExt = wb.Name
            If InStrRev(nomSansExt, ".") > 0 Then
                nomSansExt = Left$(nomSansExt, InStrRev(nomSansExt, ".") - 1)
            End If
            cbClasseur.AddItem nomSansExt
        End If
    Next wb

End Sub

' ============================================================
' PeuplerOnglets
' Peuple cbOnglet avec les feuilles du classeur sélectionné
' ============================================================
Private Sub PeuplerOnglets(ByVal nomClasseurSansExt As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nomSansExt As String

    cbOnglet.Clear

    If Trim$(nomClasseurSansExt) = "" Then Exit Sub

    For Each wb In Application.Workbooks
        nomSansExt = wb.Name
        If InStrRev(nomSansExt, ".") > 0 Then
            nomSansExt = Left$(nomSansExt, InStrRev(nomSansExt, ".") - 1)
        End If

        If StrComp(nomSansExt, Trim$(nomClasseurSansExt), vbTextCompare) = 0 Then
            For Each ws In wb.Worksheets
                cbOnglet.AddItem ws.Name
            Next ws
            Exit For
        End If
    Next wb

End Sub

' ============================================================
' Changement de classeur -> repeuplement onglets
' ============================================================
Private Sub cbClasseur_Change()

    PeuplerOnglets cbClasseur.Text

    If DefaultOnglet <> "" Then
        Dim i As Long
        For i = 0 To cbOnglet.ListCount - 1
            If StrComp(cbOnglet.List(i), DefaultOnglet, vbTextCompare) = 0 Then
                cbOnglet.Text = DefaultOnglet
                Exit For
            End If
        Next i
    End If

End Sub

' ============================================================
' Valider
' ============================================================
Private Sub cmdValider_Click()

    If Trim$(cbClasseur.Text) = "" Then
        MsgBox "Veuillez saisir ou sélectionner le classeur source.", vbExclamation
        cbClasseur.SetFocus
        Exit Sub
    End If

    If Trim$(cbOnglet.Text) = "" Then
        MsgBox "Veuillez saisir ou sélectionner l'onglet source.", vbExclamation
        cbOnglet.SetFocus
        Exit Sub
    End If

    nomClasseur = Trim$(cbClasseur.Text)
    nomOnglet = Trim$(cbOnglet.Text)
    Confirmed = True

    Me.Hide

End Sub

' ============================================================
' Annuler
' ============================================================
Private Sub cmdAnnuler_Click()
    Confirmed = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Confirmed = False
        Me.Hide
        Cancel = True
    End If
End Sub

' ============================================================
' NomSansExtension
' ============================================================
Private Function nomSansExtension(ByVal nomFichier As String) As String

    Dim pos As Long

    pos = InStrRev(nomFichier, ".")

    If pos > 0 Then
        nomSansExtension = Left$(nomFichier, pos - 1)
    Else
        nomSansExtension = nomFichier
    End If

End Function










