Attribute VB_Name = "BoutonZoom"
Option Explicit
' =============================================
'                BoutonZoom
' =============================================

#If VBA7 Then
    Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal dwFlags As Long) As LongPtr
    Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As LongPtr, lpmi As MONITORINFO) As Long
#Else
    Private Declare Function MonitorFromWindow Lib "user32" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
    Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, lpmi As MONITORINFO) As Long
#End If

Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const MONITORINFOF_PRIMARY As Long = 1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Public Sub ZoomPrincipalSecondaire()

    Dim tbl As ListObject
    Dim i As Long
    Dim premiereColVisible As Long

#If VBA7 Then
    Dim hMon As LongPtr
#Else
    Dim hMon As Long
#End If

    Dim mi As MONITORINFO
    Dim estPrincipal As Boolean

    On Error GoTo Fin

    If ActiveWindow Is Nothing Then GoTo Fin
    If ActiveSheet Is Nothing Then GoTo Fin
    If ActiveSheet.ListObjects.Count = 0 Then GoTo Fin

    Application.ScreenUpdating = False

    Set tbl = ActiveSheet.ListObjects(1)

    For i = 1 To tbl.ListColumns.Count
        If Not tbl.ListColumns(i).Range.EntireColumn.Hidden Then
            premiereColVisible = tbl.ListColumns(i).Range.Column
            Exit For
        End If
    Next i

    If premiereColVisible > 0 Then
        ActiveWindow.ScrollColumn = premiereColVisible
    End If

    hMon = MonitorFromWindow(Application.hwnd, MONITOR_DEFAULTTONEAREST)
    If hMon = 0 Then GoTo Fin

    mi.cbSize = Len(mi)
    If GetMonitorInfo(hMon, mi) = 0 Then GoTo Fin

    estPrincipal = ((mi.dwFlags And MONITORINFOF_PRIMARY) <> 0)

    If estPrincipal Then
        ActiveWindow.Zoom = ZOOM_ECRAN_PRINCIPAL
    Else
        ActiveWindow.Zoom = ZOOM_ECRAN_SECONDAIRE
    End If

    If premiereColVisible > 0 Then
        ActiveWindow.ScrollColumn = premiereColVisible
    End If

Fin:
    Application.ScreenUpdating = True

End Sub
