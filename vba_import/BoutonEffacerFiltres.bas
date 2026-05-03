Attribute VB_Name = "BoutonEffacerFiltres"
Option Explicit

' =============================================
'           Bouton Effacer Filtres
' =============================================
Public Sub EffacerFiltres()

    Dim ws As Worksheet
    Dim tbl As ListObject

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Set tbl = ws.ListObjects(1)

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    If Not tbl.AutoFilter Is Nothing Then
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
    End If
    On Error GoTo ErrHandler

    ws.Range(PLAGE_RECHERCHE).ClearContents
    InitialiserPlaceholdersFeuillePrincipale
    NettoyerBordureSelectionApresFiltre ws

SortiePropre:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de l'effacement des filtres : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

' =============================================
' Nettoyage bordure après filtre
' =============================================
Private Sub NettoyerBordureSelectionApresFiltre(ByVal ws As Worksheet)

    Dim rngLigne As Range
    Dim lig As Long

    On Error GoTo Fin

    If ws Is Nothing Then Exit Sub
    If Not ws Is ActiveSheet Then Exit Sub
    If ActiveCell Is Nothing Then Exit Sub

    lig = ActiveCell.Row
    If lig < ROW_START Then Exit Sub

    Set rngLigne = ws.Range(ws.Cells(lig, 1), ws.Cells(lig, NB_COL_UI))

    With rngLigne
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = COLOR_BORDURE_BLEUE
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = COLOR_BORDURE_BLEUE
        End With
    End With

Fin:

End Sub


