Attribute VB_Name = "BoutonZoom"
Option Explicit

Public Sub ZoomAutoTableau()

    Dim tbl As ListObject
    Dim rngVisible As Range

    On Error GoTo Fin

    If ActiveWindow Is Nothing Then GoTo Fin
    If ActiveSheet Is Nothing Then GoTo Fin
    If ActiveSheet.ListObjects.Count = 0 Then GoTo Fin

    Application.ScreenUpdating = False

    Set tbl = ActiveSheet.ListObjects(1)

    On Error Resume Next
    Set rngVisible = tbl.HeaderRowRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo Fin

    If rngVisible Is Nothing Then GoTo Fin

    rngVisible.Select
    ActiveWindow.Zoom = True

    tbl.HeaderRowRange.Cells(1, 1).Select

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = rngVisible.Cells(1, 1).Column

Fin:
    Application.ScreenUpdating = True

End Sub
