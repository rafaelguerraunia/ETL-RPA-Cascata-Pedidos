Attribute VB_Name = "modLog"
Option Explicit

' Log estruturado: grava em arquivo texto (append, mesmo padrão do RPA BASE.vbs)
' e espelha na aba Log do workbook. Níveis: INFO / WARN / ERROR / FATAL.

Private mCaminhoLog As String
Private mProximaLinhaLog As Long

Public Sub InicializarLog(ByVal caminhoArquivoLog As String)
    mCaminhoLog = caminhoArquivoLog

    Dim ws As Worksheet
    Set ws = modUtils.ObterOuCriarAba(ABA_LOG)
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Nivel"
        ws.Cells(1, 3).Value = "Hop"
        ws.Cells(1, 4).Value = "Mensagem"
        ws.Rows(1).Font.Bold = True
    End If
    mProximaLinhaLog = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If mProximaLinhaLog < 2 Then mProximaLinhaLog = 2
End Sub

Public Sub Log(ByVal nivel As String, ByVal msg As String, Optional ByVal hop As String = "")
    Dim linha As String
    linha = Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & nivel & vbTab & hop & vbTab & msg

    ' arquivo texto (append)
    On Error Resume Next
    Dim iFile As Integer
    iFile = FreeFile
    Open mCaminhoLog For Append As #iFile
    Print #iFile, linha
    Close #iFile
    On Error GoTo 0

    ' aba Log
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ABA_LOG)
    If Not ws Is Nothing Then
        ws.Cells(mProximaLinhaLog, 1).Value = Now
        ws.Cells(mProximaLinhaLog, 2).Value = nivel
        ws.Cells(mProximaLinhaLog, 3).Value = hop
        ws.Cells(mProximaLinhaLog, 4).Value = msg
        mProximaLinhaLog = mProximaLinhaLog + 1
    End If
    On Error GoTo 0

    Debug.Print linha
End Sub

Public Sub LogInfo(ByVal msg As String, Optional ByVal hop As String = "")
    Log LOG_INFO, msg, hop
End Sub

Public Sub LogWarn(ByVal msg As String, Optional ByVal hop As String = "")
    Log LOG_WARN, msg, hop
End Sub

Public Sub LogErro(ByVal msg As String, Optional ByVal hop As String = "")
    Log LOG_ERROR, msg, hop
End Sub

' Erro fatal: loga e aborta o run (usado só para condições que impedem qualquer progresso).
Public Sub Falhar(ByVal msg As String)
    Log LOG_FATAL, msg
    MsgBox msg, vbCritical, "ETL Cascata - Falha"
End Sub

Public Sub Informar(ByVal msg As String)
    MsgBox msg, vbInformation, "ETL Cascata"
End Sub
