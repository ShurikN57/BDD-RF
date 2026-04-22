Attribute VB_Name = "BoutonSauvegarde"
Option Explicit

' =============================================
'            Bouton Sauvegarde
' =============================================
Public SauvegardeAutorisee As Boolean

Public Sub SauvegarderClasseur()

    On Error GoTo ErrHandler

    SauvegardeAutorisee = False
    ThisWorkbook.Save

    If SauvegardeAutorisee Then
        MsgBox "Classeur sauvegardé, merci pour votre contribution à ce projet.", vbInformation
    End If

    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de la sauvegarde : " & Err.description, vbExclamation

End Sub

