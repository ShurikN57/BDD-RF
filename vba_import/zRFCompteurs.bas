Attribute VB_Name = "zRFCompteurs"
Option Explicit

' =============================================
' Remplace les 3 formules lourdes de MENU DEROULANT J1:L1
'
' J1 = nombre de valeurs uniques visibles en AE
' K1 = nombre de valeurs uniques visibles en AF
' L1 = concatenation J1 & " | " & K1
' =============================================

Private Const COL_COMPTEUR_1 As String = "AE"
Private Const COL_COMPTEUR_2 As String = "AF"

Private Const CELL_NB_AE As String = "J1"
Private Const CELL_NB_AF As String = "K1"
Private Const CELL_SYNTHESE As String = "L1"

Private m_nbAE_Total As Long
Private m_nbAF_Total As Long
Private m_Initialise As Boolean

' =============================================
' Calcul complet ouverture + filtre
' =============================================
Public Sub MettreAJourCompteurs()

    Dim ws As Worksheet
    Dim wsMenu As Worksheet
    Dim lastRow As Long
    Dim dictAE As Object
    Dim dictAF As Object
    Dim nbAE As Long
    Dim nbAF As Long
    Dim bFiltre As Boolean
    Dim rngVisibleRows As Range
    Dim area As Range
    Dim firstLig As Long
    Dim lastLigArea As Long
    Dim arrAE As Variant
    Dim arrAF As Variant

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Set wsMenu = ThisWorkbook.Worksheets(SHEET_MENU_DEROULANT)

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    lastRow = DerniereLigneCompteursRF(ws)

    If lastRow < ROW_START Then
        EcrireCompteursRF wsMenu, 0, 0
        GoTo SortiePropre
    End If

    Set dictAE = CreateObject("Scripting.Dictionary")
    Set dictAF = CreateObject("Scripting.Dictionary")

    dictAE.CompareMode = vbTextCompare
    dictAF.CompareMode = vbTextCompare

    bFiltre = ws.FilterMode

    If Not bFiltre Then

        arrAE = ws.Range(COL_COMPTEUR_1 & ROW_START & ":" & COL_COMPTEUR_1 & lastRow).Value2
        arrAF = ws.Range(COL_COMPTEUR_2 & ROW_START & ":" & COL_COMPTEUR_2 & lastRow).Value2

        CompterUniques2Colonnes arrAE, arrAF, dictAE, dictAF

    Else

        On Error Resume Next
        Set rngVisibleRows = ws.Range(COL_FIRST & ROW_START & ":" & COL_FIRST & lastRow).SpecialCells(xlCellTypeVisible)
        On Error GoTo ErrHandler

        If Not rngVisibleRows Is Nothing Then
            For Each area In rngVisibleRows.Areas

                firstLig = area.Row
                lastLigArea = area.Row + area.Rows.Count - 1

                arrAE = ws.Range(COL_COMPTEUR_1 & firstLig & ":" & COL_COMPTEUR_1 & lastLigArea).Value2
                arrAF = ws.Range(COL_COMPTEUR_2 & firstLig & ":" & COL_COMPTEUR_2 & lastLigArea).Value2

                CompterUniques2Colonnes arrAE, arrAF, dictAE, dictAF

            Next area
        End If

    End If

    nbAE = dictAE.Count
    nbAF = dictAF.Count

    If Not bFiltre Then
        m_nbAE_Total = nbAE
        m_nbAF_Total = nbAF
        m_Initialise = True
    End If

    EcrireCompteursRF wsMenu, nbAE, nbAF

SortiePropre:
    Set dictAE = Nothing
    Set dictAF = Nothing
    Set rngVisibleRows = Nothing

    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de la mise à jour des compteurs RF : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

' =============================================
' Derniere ligne utile pour les compteurs RF
' =============================================
Private Function DerniereLigneCompteursRF(ByVal ws As Worksheet) As Long

    Dim lastA As Long
    Dim lastAE As Long
    Dim lastAF As Long
    Dim lastRow As Long

    lastA = ws.Cells(ws.Rows.Count, COL_FIRST).End(xlUp).Row
    lastAE = ws.Cells(ws.Rows.Count, COL_COMPTEUR_1).End(xlUp).Row
    lastAF = ws.Cells(ws.Rows.Count, COL_COMPTEUR_2).End(xlUp).Row

    lastRow = lastA
    If lastAE > lastRow Then lastRow = lastAE
    If lastAF > lastRow Then lastRow = lastAF

    If lastRow < ROW_START Then lastRow = ROW_START

    DerniereLigneCompteursRF = lastRow

End Function

' =============================================
' Comptage simultane des 2 colonnes
' =============================================
Private Sub CompterUniques2Colonnes(ByVal arrAE As Variant, ByVal arrAF As Variant, _
                                    ByVal dictAE As Object, ByVal dictAF As Object)

    Dim i As Long
    Dim valAE As String
    Dim valAF As String

    If Not IsArray(arrAE) Then

        valAE = Trim$(CStr(arrAE))
        valAF = Trim$(CStr(arrAF))

        If valAE <> "" Then
            If Not dictAE.Exists(valAE) Then dictAE.Add valAE, 1
        End If

        If valAF <> "" Then
            If Not dictAF.Exists(valAF) Then dictAF.Add valAF, 1
        End If

        Exit Sub

    End If

    For i = 1 To UBound(arrAE, 1)

        valAE = Trim$(CStr(arrAE(i, 1)))
        valAF = Trim$(CStr(arrAF(i, 1)))

        If valAE <> "" Then
            If Not dictAE.Exists(valAE) Then dictAE.Add valAE, 1
        End If

        If valAF <> "" Then
            If Not dictAF.Exists(valAF) Then dictAF.Add valAF, 1
        End If

    Next i

End Sub

' =============================================
' Restauration instantanee apres effacement filtre
' =============================================
Public Sub RestaurerCompteursInitiaux()

    Dim wsMenu As Worksheet

    On Error GoTo ErrHandler

    If Not m_Initialise Then
        MettreAJourCompteurs
        Exit Sub
    End If

    Set wsMenu = ThisWorkbook.Worksheets(SHEET_MENU_DEROULANT)

    EcrireCompteursRF wsMenu, m_nbAE_Total, m_nbAF_Total
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de la restauration des compteurs RF : " & Err.description, vbExclamation

End Sub

' =============================================
' Ecriture dans MENU DEROULANT
' =============================================
Private Sub EcrireCompteursRF(ByVal wsMenu As Worksheet, ByVal nbAE As Long, ByVal nbAF As Long)

    wsMenu.Range(CELL_NB_AE).Value = nbAE & " RF principal"
    wsMenu.Range(CELL_NB_AF).Value = nbAF & " RF associés"
    wsMenu.Range(CELL_SYNTHESE).Value = nbAE & " RF principal | " & nbAF & " RF associés"

End Sub

' =============================================
' Wrappers BDD-RF pour les boutons de la feuille
'
' Affectations conseillees :
' - Bouton "Appliquer" -> AppliquerFiltresRF
' - Bouton "Effacer"   -> EffacerFiltresRF
' =============================================
Public Sub AppliquerFiltresRF()

    AppliquerFiltres
    MettreAJourCompteurs

End Sub

Public Sub EffacerFiltresRF()

    EffacerFiltres
    RestaurerCompteursInitiaux

End Sub

Public Sub InitialiserPlaceholdersFeuillePrincipale()
    GMC.InitialiserPlaceholders
End Sub

