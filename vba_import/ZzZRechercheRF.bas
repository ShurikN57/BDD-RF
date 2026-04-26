Attribute VB_Name = "ZzZRechercheRF"
Option Explicit
' =============================================
' RechercheRF (utile pour reconstruire AE seulement)
' AE = concaténation de A:F
' =============================================
Public Sub RebuildCleAE()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim arrIn As Variant
    Dim arrOut As Variant
    Dim i As Long
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim succes As Boolean

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    succes = False

    If lastRow < 8 Then
        succes = True
        GoTo SortiePropre
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1 seule lecture matricielle A:F
    arrIn = ws.Range("A8:F" & lastRow).Value

    ReDim arrOut(1 To UBound(arrIn, 1), 1 To 1)

    For i = 1 To UBound(arrIn, 1)
        arrOut(i, 1) = CStr(arrIn(i, 1)) & _
                       CStr(arrIn(i, 2)) & _
                       CStr(arrIn(i, 3)) & _
                       CStr(arrIn(i, 4)) & _
                       CStr(arrIn(i, 5)) & _
                       CStr(arrIn(i, 6))
    Next i

    ' 1 seule écriture matricielle AE
    ws.Range("AE8").Resize(UBound(arrOut, 1), 1).Value = arrOut

    succes = True

SortiePropre:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating

    If succes Then
        If lastRow >= 8 Then
            MsgBox "Colonne AE reconstruite (" & UBound(arrOut, 1) & " lignes).", vbInformation
        Else
            MsgBox "Aucune ligne à reconstruire dans la colonne AE.", vbInformation
        End If
    End If
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de la reconstruction de la clé AE : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

' =============================================
' RechercheRF (utile pour reconstruire AF seulement)
' AF = concaténation de N:S
' =============================================
Public Sub RebuildCleAF()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim arrIn As Variant
    Dim arrOut As Variant
    Dim i As Long
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim succes As Boolean

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    succes = False

    If lastRow < 8 Then
        succes = True
        GoTo SortiePropre
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1 seule lecture matricielle N:S
    arrIn = ws.Range("N8:S" & lastRow).Value

    ReDim arrOut(1 To UBound(arrIn, 1), 1 To 1)

    For i = 1 To UBound(arrIn, 1)
        arrOut(i, 1) = CStr(arrIn(i, 1)) & _
                       CStr(arrIn(i, 2)) & _
                       CStr(arrIn(i, 3)) & _
                       CStr(arrIn(i, 4)) & _
                       CStr(arrIn(i, 5)) & _
                       CStr(arrIn(i, 6))
    Next i

    ' 1 seule écriture matricielle AF
    ws.Range("AF8").Resize(UBound(arrOut, 1), 1).Value = arrOut

    succes = True

SortiePropre:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating

    If succes Then
        If lastRow >= 8 Then
            MsgBox "Colonne AF reconstruite (" & UBound(arrOut, 1) & " lignes).", vbInformation
        Else
            MsgBox "Aucune ligne à reconstruire dans la colonne AF.", vbInformation
        End If
    End If
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de la reconstruction de la clé AF : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

