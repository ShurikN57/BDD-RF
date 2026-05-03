Attribute VB_Name = "BoutonAppliquerFiltres"
Option Explicit

' =============================================
'          Bouton Appliquer Filtres
' =============================================
Public Sub AppliquerFiltres()

    Dim ws As Worksheet
    Dim wsTitres As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim val As String
    Dim arrBarre As Variant
    Dim arrTitres As Variant

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Set wsTitres = ThisWorkbook.Worksheets(SHEET_TITRES)
    Set tbl = ws.ListObjects(1)

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo ErrHandler

    arrBarre = ws.Range(PLAGE_RECHERCHE).Value
    arrTitres = wsTitres.Range(COL_FIRST & ROW_TITRES & ":" & COL_LAST_RECHERCHE & ROW_TITRES).Value

    For i = 1 To NB_COL_RECHERCHE
        val = Trim$(CStr(arrBarre(1, i)))
        If val = Trim$(CStr(arrTitres(1, i))) Then val = ""
        If val <> "" Then
            tbl.Range.AutoFilter Field:=i, Criteria1:=val
        End If
    Next i

    NettoyerBordureSelectionApresFiltre ws

SortiePropre:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de l'application des filtres : " & Err.description, vbExclamation
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


