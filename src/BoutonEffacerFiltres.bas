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

SortiePropre:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de l'effacement des filtres : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub