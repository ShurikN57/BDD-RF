Attribute VB_Name = "CentrerUserForm"
Option Explicit

' =============================================
'            Centrer UserForm
' =============================================
#If VBA7 Then
    Public Declare PtrSafe Function MonitorFromWindow Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal dwFlags As Long) As LongPtr
    Public Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" ( _
        ByVal hMonitor As LongPtr, lpmi As MONITORINFO) As Long
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
#Else
    Public Declare Function MonitorFromWindow Lib "user32" ( _
        ByVal hwnd As Long, ByVal dwFlags As Long) As Long
    Public Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" ( _
        ByVal hMonitor As Long, lpmi As MONITORINFO) As Long
    Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
#End If

Public Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const LOGPIXELSX As Long = 88

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Public Sub CentrerUserFormSurMoniteurExcel(ByVal frm As Object, _
                                           Optional ByVal ratioTop As Double = 0.5)
#If VBA7 Then
    Dim hMon As LongPtr
    Dim hDC As LongPtr
#Else
    Dim hMon As Long
    Dim hDC As Long
#End If

    Dim mi As MONITORINFO
    Dim dpiX As Long
    Dim ptPerPx As Double
    Dim zoneLeftPt As Double
    Dim zoneTopPt As Double
    Dim zoneWidthPt As Double
    Dim zoneHeightPt As Double

    On Error GoTo Fallback

    frm.StartUpPosition = 0

    ' --- DPI réel de l'écran ---
    hDC = GetDC(Application.hwnd)
    If hDC <> 0 Then
        dpiX = GetDeviceCaps(hDC, LOGPIXELSX)
        ReleaseDC Application.hwnd, hDC
    End If
    If dpiX = 0 Then dpiX = 96

    ptPerPx = 72# / CDbl(dpiX)

    ' --- Moniteur courant ---
    hMon = MonitorFromWindow(Application.hwnd, MONITOR_DEFAULTTONEAREST)
    If hMon = 0 Then GoTo Fallback

    mi.cbSize = LenB(mi)
    If GetMonitorInfo(hMon, mi) = 0 Then GoTo Fallback

    ' --- Conversion zone de travail en points ---
    zoneLeftPt = mi.rcWork.Left * ptPerPx
    zoneTopPt = mi.rcWork.Top * ptPerPx
    zoneWidthPt = (mi.rcWork.Right - mi.rcWork.Left) * ptPerPx
    zoneHeightPt = (mi.rcWork.Bottom - mi.rcWork.Top) * ptPerPx

    frm.Left = zoneLeftPt + (zoneWidthPt - frm.Width) / 2
    frm.Top = zoneTopPt + (zoneHeightPt - frm.Height) * ratioTop

    Exit Sub

Fallback:
    frm.StartUpPosition = 0
    frm.Left = Application.Left + (Application.Width - frm.Width) / 2
    frm.Top = Application.Top + (Application.Height - frm.Height) * ratioTop

End Sub
