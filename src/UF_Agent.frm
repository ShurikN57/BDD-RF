VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Agent 
   Caption         =   "Sélection de l'agent"
   ClientHeight    =   3225
   ClientLeft      =   330
   ClientTop       =   1305
   ClientWidth     =   4290
   OleObjectBlob   =   "UF_Agent_codex.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_Agent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' UserForm UF_Agent
' ============================================================

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Dim i As Long
    Dim val As String

    Set ws = ThisWorkbook.Worksheets(SHEET_MENU_DEROULANT)

    For i = ROW_AGENTS_START To ROW_AGENTS_END
        val = Trim$(CStr(ws.Cells(i, COL_AGENTS).Value))
        If val <> "" Then cbAgent.AddItem val
    Next i

    cbAgent.ListRows = 8

    CentrerUserFormSurMoniteurExcel Me, 0.5

End Sub

Private Sub cmdValider_Click()

    If cbAgent.Value = "" Then
        MsgBox "Veuillez sélectionner un nom.", vbExclamation
        Exit Sub
    End If

    With ThisWorkbook.Worksheets(SHEET_MAIN)
        .Range(CELL_NOM_SESSION).Value = cbAgent.Value
        .Range(CELL_DATE_SESSION).Value = Date
    End With

    Unload Me

End Sub

Private Sub cmdAnnuler_Click()

    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Application.DisplayAlerts = False
        ThisWorkbook.Close
        Application.DisplayAlerts = True
    End If

End Sub

