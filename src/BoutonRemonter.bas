Attribute VB_Name = "BoutonRemonter"
Option Explicit

' =============================================
'              Bouton Remonter
' =============================================
Public Sub AllerEnHaut()

    Dim ws As Worksheet

    On Error GoTo Fin

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)

    Application.ScreenUpdating = False

    ws.Activate
    Application.GoTo ws.Range(COL_FIRST & ROW_RECHERCHE), True

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

Fin:
    Application.ScreenUpdating = True

End Sub



