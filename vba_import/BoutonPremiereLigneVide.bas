Attribute VB_Name = "BoutonPremiereLigneVide"
Option Explicit

' =============================================
'         Bouton Premiere Ligne Vide
' =============================================
Public Sub AllerPremiereLigneVide()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim arr As Variant
    Dim i As Long

    On Error GoTo Fin

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    lastRow = ws.Cells(ws.Rows.Count, COL_PREMIERE_LIGNE_VIDE).End(xlUp).Row

    If lastRow < ROW_START Then
        ws.Activate
        Application.GoTo ws.Range(COL_PREMIERE_LIGNE_VIDE & ROW_START), False
        Exit Sub
    End If

    arr = ws.Range(COL_PREMIERE_LIGNE_VIDE & ROW_START & ":" & COL_PREMIERE_LIGNE_VIDE & lastRow).Value

    For i = 1 To UBound(arr, 1)
        If Trim$(CStr(arr(i, 1))) = "" Then
            ws.Activate
            Application.GoTo ws.Cells(i + ROW_START - 1, COL_PREMIERE_LIGNE_VIDE), False
            Exit Sub
        End If
    Next i

    ws.Activate
    Application.GoTo ws.Cells(lastRow + 1, COL_PREMIERE_LIGNE_VIDE), False

Fin:
End Sub

