Attribute VB_Name = "BoutonAfficher"
Option Explicit

' =============================================
'       Afficher / Masquer Colonnes
' =============================================
Public Sub AfficherColonnesSecondaires()

    Dim ws As Worksheet

    On Error GoTo Fin

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)

    Application.ScreenUpdating = False
    ws.Columns(COLONNES_MASQUEES).Hidden = Not ws.Columns(COLONNES_MASQUEES).Hidden

Fin:
    Application.ScreenUpdating = True

End Sub

