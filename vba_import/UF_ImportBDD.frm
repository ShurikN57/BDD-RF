VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ImportBDD 
   Caption         =   "Import Agents- > BDD-DOC"
   ClientHeight    =   4584
   ClientLeft      =   15
   ClientTop       =   90
   ClientWidth     =   5595
   OleObjectBlob   =   "UF_ImportBDD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_ImportBDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' UF_ImportBDD
' Sélection classeur source / onglet source
' Classeur cible et onglet cible = ThisWorkbook / SHEET_MAIN (fixes)
' ============================================================

Public Confirmed        As Boolean
Public nomClasseurSrc   As String
Public nomOngletSrc     As String
Public nomClasseurCible As String
Public nomOngletCible   As String

' Valeurs par défaut injectées par l'appelant
Public DefaultClasseurSrc   As String
Public DefaultOngletSrc     As String
Public DefaultClasseurCible As String
Public DefaultOngletCible   As String

' ============================================================
' Initialize
' ============================================================
Private Sub UserForm_Initialize()

    Confirmed = False
    nomClasseurSrc = ""
    nomOngletSrc = ""
    nomClasseurCible = ""
    nomOngletCible = ""

End Sub

' ============================================================
' Activate
' ============================================================
Private Sub UserForm_Activate()
    CentrerUserFormSurMoniteurExcel Me, 0.33
End Sub

' ============================================================
' InitialiserImportBDD
' Appelé depuis zDocImportBDD via New UF_ImportBDD
' ============================================================
Public Sub InitialiserImportBDD(ByVal pDefaultClasseurSrc As String, _
                                 ByVal pDefaultOngletSrc As String)

    DefaultClasseurSrc = pDefaultClasseurSrc
    DefaultOngletSrc = pDefaultOngletSrc

    ' --- Peuplement ComboBox classeur source uniquement ---
    PeuplerClasseurs cbClasseurSrc

    ' --- Valeurs par défaut source ---
    If DefaultClasseurSrc <> "" Then
        cbClasseurSrc.Value = DefaultClasseurSrc
        PeuplerOnglets cbClasseurSrc.Value, cbOngletSrc
    End If

    If DefaultOngletSrc <> "" Then
        cbOngletSrc.Value = DefaultOngletSrc
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
Private Sub PeuplerClasseurs(ByRef cb As MSForms.ComboBox)

    Dim wb As Workbook
    Dim nomSansExt As String

    cb.Clear

    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            nomSansExt = wb.Name
            If InStrRev(nomSansExt, ".") > 0 Then
                nomSansExt = Left$(nomSansExt, InStrRev(nomSansExt, ".") - 1)
            End If
            cb.AddItem nomSansExt
        End If
    Next wb

End Sub

' ============================================================
' PeuplerOnglets
' Peuple cbOnglet avec les feuilles du classeur sélectionné
' ============================================================
Private Sub PeuplerOnglets(ByVal nomClasseurSansExt As String, ByRef cb As MSForms.ComboBox)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nomSansExt As String

    cb.Clear

    If Trim$(nomClasseurSansExt) = "" Then Exit Sub

    For Each wb In Application.Workbooks
        nomSansExt = wb.Name
        If InStrRev(nomSansExt, ".") > 0 Then
            nomSansExt = Left$(nomSansExt, InStrRev(nomSansExt, ".") - 1)
        End If

        If StrComp(nomSansExt, Trim$(nomClasseurSansExt), vbTextCompare) = 0 Then
            For Each ws In wb.Worksheets
                cb.AddItem ws.Name
            Next ws
            Exit For
        End If
    Next wb

End Sub

' ============================================================
' Changement classeur source -> repeuplement onglets source
' ============================================================
Private Sub cbClasseurSrc_Change()

    PeuplerOnglets cbClasseurSrc.Text, cbOngletSrc

    If DefaultOngletSrc <> "" Then
        Dim i As Long
        For i = 0 To cbOngletSrc.ListCount - 1
            If StrComp(cbOngletSrc.List(i), DefaultOngletSrc, vbTextCompare) = 0 Then
                cbOngletSrc.Text = DefaultOngletSrc
                Exit For
            End If
        Next i
    End If

End Sub

' ============================================================
' Valider
' ============================================================
Private Sub cmdValider_Click()

    If Trim$(cbClasseurSrc.Text) = "" Then
        MsgBox "Veuillez saisir ou sélectionner le classeur source.", vbExclamation
        cbClasseurSrc.SetFocus
        Exit Sub
    End If

    If Trim$(cbOngletSrc.Text) = "" Then
        MsgBox "Veuillez saisir ou sélectionner l'onglet source.", vbExclamation
        cbOngletSrc.SetFocus
        Exit Sub
    End If

    nomClasseurSrc = Trim$(cbClasseurSrc.Text)
    nomOngletSrc = Trim$(cbOngletSrc.Text)
    nomClasseurCible = nomSansExtension(ThisWorkbook.Name)
    nomOngletCible = SHEET_MAIN
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








