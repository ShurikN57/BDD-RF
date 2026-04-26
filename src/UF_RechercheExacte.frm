VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_RechercheExacte 
   Caption         =   "Recherche exacte RF principal"
   ClientHeight    =   1335
   ClientLeft      =   390
   ClientTop       =   1545
   ClientWidth     =   4560
   OleObjectBlob   =   "UF_RechercheExacte_codex.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_RechercheExacte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub txtRecherche_Change()

End Sub

' ============================================================
' UF_RechercheExacte
' ============================================================

Private Sub UserForm_Initialize()

    CentrerUserFormSurMoniteurExcel Me, 0.5

End Sub

Private Sub cmdRechercher_Click()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim valeurRecherchee As String
    Dim lastRow As Long
    Dim firstDataCol As Long
    Dim lastDataCol As Long
    Dim filterRange As Range
    Dim fieldIndex As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    valeurRecherchee = Trim$(Me.txtRecherche.Value)

    If valeurRecherchee = "" Then
        MsgBox "Veuillez saisir une valeur.", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, COL_FIRST).End(xlUp).Row

    If lastRow < ROW_START Then
        MsgBox "Aucune donnée dans la feuille.", vbExclamation
        Exit Sub
    End If

    firstDataCol = ws.Range(COL_FIRST & "1").Column
    lastDataCol = ws.Range(COL_LAST & "1").Column

    Set filterRange = ws.Range(ws.Cells(ROW_HEADER, firstDataCol), ws.Cells(lastRow, lastDataCol))
    fieldIndex = ws.Range(COL_RECHERCHE_EXACTE & "1").Column - firstDataCol + 1

    Application.ScreenUpdating = False

    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo ErrHandler

    If ws.AutoFilterMode = False Then
        filterRange.AutoFilter
    End If

    filterRange.AutoFilter Field:=fieldIndex, Criteria1:="=" & valeurRecherchee

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la recherche : " & Err.description, vbCritical

End Sub

Private Sub cmdReset_Click()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim firstDataCol As Long
    Dim lastDataCol As Long
    Dim filterRange As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    lastRow = ws.Cells(ws.Rows.Count, COL_FIRST).End(xlUp).Row

    firstDataCol = ws.Range(COL_FIRST & "1").Column
    lastDataCol = ws.Range(COL_LAST & "1").Column

    Application.ScreenUpdating = False

    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo ErrHandler

    If ws.AutoFilterMode = False And lastRow >= ROW_START Then
        Set filterRange = ws.Range(ws.Cells(ROW_HEADER, firstDataCol), ws.Cells(lastRow, lastDataCol))
        filterRange.AutoFilter
    End If

    Me.txtRecherche.Value = ""

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la réinitialisation : " & Err.description, vbCritical

End Sub

Private Sub cmdFermer_Click()
    Unload Me
End Sub



