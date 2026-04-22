Attribute VB_Name = "Utilitaires"

Option Explicit

' =============================================
'              Utilitaires communs
' =============================================

Public Function DerniereLigneUtileMain() As Long

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    lastRow = ws.Cells(ws.Rows.Count, COL_FIRST).End(xlUp).Row

    If lastRow < ROW_START Then lastRow = ROW_START

    DerniereLigneUtileMain = lastRow

End Function
