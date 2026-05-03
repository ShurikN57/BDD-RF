Attribute VB_Name = "AllignerBoutons"
Option Explicit

' =============================================
' 1. Alligner Boutons
' =============================================
Sub AlignerBoutons()

    ' =============================================
    ' SEUL PARAMÈTRE À MODIFIER
    Dim margeInterieure As Double: margeInterieure = 6  ' espace bord ? boutons (haut, bas, gauche, droite)
    ' =============================================
    ' Taille des boutons calculée automatiquement
    ' pour remplir Grand1 de façon homogène

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim grandRect As Shape
    Set grandRect = ws.Shapes("Grand1")

    ' Calcul automatique largeur et hauteur des boutons
    Dim largeurBouton As Double
    Dim hauteurBouton As Double
    Dim espaceH As Double  ' espace horizontal entre les 2 colonnes
    Dim espaceV As Double  ' espace vertical entre les 2 lignes

    ' 2 colonnes ? largeur dispo partagée en 2
    largeurBouton = (grandRect.Width - (2 * margeInterieure) - margeInterieure) / 2
    espaceH = margeInterieure

    ' 2 lignes ? hauteur dispo partagée en 2
    hauteurBouton = (grandRect.Height - (2 * margeInterieure) - margeInterieure) / 2
    espaceV = margeInterieure

    ' Positions
    Dim left1 As Double: left1 = grandRect.Left + margeInterieure
    Dim left2 As Double: left2 = left1 + largeurBouton + espaceH
    Dim top1  As Double: top1 = grandRect.Top + margeInterieure
    Dim top2  As Double: top2 = top1 + hauteurBouton + espaceV

    With ws.Shapes("Appliquer")
        .Width = largeurBouton: .Height = hauteurBouton
        .Left = left1: .Top = top1
    End With

    With ws.Shapes("RF")
        .Width = largeurBouton: .Height = hauteurBouton
        .Left = left2: .Top = top1
    End With
    

    With ws.Shapes("Effacer")
        .Width = largeurBouton: .Height = hauteurBouton
        .Left = left1: .Top = top2
    End With

    With ws.Shapes("Annuler")
        .Width = largeurBouton: .Height = hauteurBouton
        .Left = left2: .Top = top2
    End With

    MsgBox "Boutons alignés !", vbInformation

End Sub
 
    ' =============================================
    ' 2. AlignerBoutons2
    ' =============================================
Sub AlignerBoutons2()

  
    Dim margeInterieure As Double: margeInterieure = 6
    ' =============================================

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim grandRect As Shape
    Set grandRect = ws.Shapes("Grand2")

    ' Calcul automatique
    Dim largeurBouton As Double
    Dim hauteurBouton As Double

    largeurBouton = (grandRect.Width - (2 * margeInterieure) - margeInterieure) / 2
    hauteurBouton = (grandRect.Height - (2 * margeInterieure) - margeInterieure) / 2

    ' Positions
    Dim left1 As Double: left1 = grandRect.Left + margeInterieure
    Dim left2 As Double: left2 = left1 + largeurBouton + margeInterieure
    Dim top1  As Double: top1 = grandRect.Top + margeInterieure
    Dim top2  As Double: top2 = top1 + hauteurBouton + margeInterieure

    With ws.Shapes("Remonter")
        .Width = largeurBouton: .Height = hauteurBouton
        .Left = left1: .Top = top1
    End With

    With ws.Shapes("Zoom")
        .Width = largeurBouton: .Height = hauteurBouton
        .Left = left2: .Top = top1
    End With

    With ws.Shapes("1ere Ligne")
        .Width = largeurBouton: .Height = hauteurBouton
        .Left = left1: .Top = top2
    End With

    With ws.Shapes("Afficher")
        .Width = largeurBouton: .Height = hauteurBouton
        .Left = left2: .Top = top2
    End With

    MsgBox "Boutons Grand2 alignés !", vbInformation

End Sub

    ' =============================================
    ' 3. AlignerBoutons3
    ' =============================================
Sub AlignerBoutons3()

    Dim margeInterieure As Double: margeInterieure = 5
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim grandRect As Shape
    Set grandRect = ws.Shapes("Grand3")

    Dim largeurBouton As Double
    Dim hauteurBouton As Double
    Dim largeurSauveg  As Double

    largeurBouton = (grandRect.Width - (2 * margeInterieure) - margeInterieure) / 2
    hauteurBouton = (grandRect.Height - (2 * margeInterieure) - margeInterieure) / 2
    largeurSauveg = grandRect.Width - (2 * margeInterieure)

    Dim top1 As Double: top1 = grandRect.Top + margeInterieure
    Dim top2 As Double: top2 = top1 + hauteurBouton + margeInterieure

    ' ON ? bord gauche de Grand3
    Dim leftON  As Double: leftON = grandRect.Left + margeInterieure
    ' OFF ? bord DROIT de Grand3 (indépendant de ON)
    Dim leftOFF As Double: leftOFF = grandRect.Left + grandRect.Width - margeInterieure - largeurBouton

    ws.Shapes("ON").Width = largeurBouton
    ws.Shapes("ON").Height = hauteurBouton
    ws.Shapes("ON").Left = leftON
    ws.Shapes("ON").Top = top1

    ws.Shapes("OFF").Width = largeurBouton
    ws.Shapes("OFF").Height = hauteurBouton
    ws.Shapes("OFF").Left = leftOFF
    ws.Shapes("OFF").Top = top1

    ws.Shapes("SAUVEGARDER").Width = largeurSauveg
    ws.Shapes("SAUVEGARDER").Height = hauteurBouton
    ws.Shapes("SAUVEGARDER").Left = leftON
    ws.Shapes("SAUVEGARDER").Top = top2

    MsgBox "Boutons Grand3 alignés !", vbInformation

End Sub