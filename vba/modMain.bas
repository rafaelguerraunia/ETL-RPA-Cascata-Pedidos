Attribute VB_Name = "modMain"
Option Explicit

' Ponto de entrada - ligar este Sub a um botão na aba "Início".

Public Sub RunETLCascata()
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    On Error GoTo Finalizar

    modCascade.ExecutarCascata

Finalizar:
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Erro inesperado: " & Err.Description, vbCritical, "ETL Cascata"
    End If
End Sub
