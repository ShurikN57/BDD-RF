Attribute VB_Name = "BoutonRFPrincipal"

Option Explicit

' =============================================
'         Bouton Recherche RF Principal
' =============================================
Public Sub OuvrirRechercheRF()

    On Error GoTo Fin

    UF_RechercheExacte.Show

Fin:
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'ouverture de la recherche : " & Err.description, vbExclamation
    End If

End Sub

