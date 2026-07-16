Attribute VB_Name = "modExcelWrite"
Option Explicit

' Upsert por chave nas abas de cascata (histórico preservado entre execuções)
' e leitura de abas já gravadas (usado pelo modJoinBuilder).

' Grava dadosBrutos (array 2D com cabeçalho na linha 1, como veio de modSQVIEngine.RunSQVI)
' na aba de destino da spec, mesclando com o que já existe por chave (spec.ChaveUpsert).
' Reescreve a aba inteira de uma vez (mais rápido que atualizar linha a linha e garante
' que linhas de execuções anteriores não tocadas neste run permaneçam intactas).
Public Sub UpsertRows(ByVal spec As clsSQVISpec, ByVal dadosBrutos As Variant)
    If Not modUtils.EhArrayValido(dadosBrutos) Then
        modLog.LogInfo "Nada para gravar em " & spec.AbaDestino & " (sem dados).", spec.AbaDestino
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = modUtils.ObterOuCriarAba(spec.AbaDestino)

    Dim numColunas As Long
    numColunas = spec.FELD.Count

    If ws.Cells(1, 1).Value = "" Then
        Dim colIdx As Long
        For colIdx = 1 To numColunas
            ws.Cells(1, colIdx).Value = spec.FELD(colIdx)
        Next colIdx
        ws.Rows(1).Font.Bold = True
    End If

    Dim posicoesChave() As Long
    ReDim posicoesChave(1 To spec.ChaveUpsert.Count)
    Dim k As Long
    For k = 1 To spec.ChaveUpsert.Count
        posicoesChave(k) = modUtils.ColunaDoCampo(spec, spec.ChaveUpsert(k))
    Next k

    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    Dim ordemChaves As New Collection

    ' -------- carrega o que já existe na aba --------
    Dim ultimaLinhaExistente As Long
    ultimaLinhaExistente = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If ultimaLinhaExistente >= 2 Then
        Dim existente As Variant
        existente = ws.Range(ws.Cells(2, 1), ws.Cells(ultimaLinhaExistente, numColunas)).Value2

        If modUtils.EhArrayValido(existente) Then
            Dim i As Long, colIdx2 As Long
            For i = LBound(existente, 1) To UBound(existente, 1)
                Dim linhaVals() As Variant
                ReDim linhaVals(1 To numColunas)
                For colIdx2 = 1 To numColunas
                    linhaVals(colIdx2) = existente(i, colIdx2)
                Next colIdx2

                Dim chaveTxt As String
                chaveTxt = MontarChave(linhaVals, posicoesChave)
                If Len(chaveTxt) > 0 Then
                    If Not mapa.Exists(chaveTxt) Then ordemChaves.Add chaveTxt
                    mapa(chaveTxt) = linhaVals
                End If
            Next i
        End If
    End If

    ' -------- aplica os dados novos (sobrescreve por chave, ou adiciona) --------
    Dim r As Long, colIdx3 As Long
    For r = LBound(dadosBrutos, 1) + 1 To UBound(dadosBrutos, 1)   ' pula cabeçalho (linha 1 do export SAP)
        Dim novaLinha() As Variant
        ReDim novaLinha(1 To numColunas)
        For colIdx3 = 1 To numColunas
            If colIdx3 <= UBound(dadosBrutos, 2) Then
                novaLinha(colIdx3) = dadosBrutos(r, colIdx3)
            Else
                novaLinha(colIdx3) = ""
            End If
        Next colIdx3

        Dim chaveNova As String
        chaveNova = MontarChave(novaLinha, posicoesChave)
        If Len(chaveNova) > 0 Then
            If Not mapa.Exists(chaveNova) Then ordemChaves.Add chaveNova
            mapa(chaveNova) = novaLinha
        End If
    Next r

    ' -------- reescreve a aba inteira de uma vez --------
    Dim totalLinhas As Long
    totalLinhas = ordemChaves.Count

    If totalLinhas > 0 Then
        Dim saida() As Variant
        ReDim saida(1 To totalLinhas, 1 To numColunas)
        Dim idx As Long, colIdxOut As Long
        idx = 1
        Dim chv As Variant, vals As Variant
        For Each chv In ordemChaves
            vals = mapa(CStr(chv))
            For colIdxOut = 1 To numColunas
                saida(idx, colIdxOut) = vals(colIdxOut)
            Next colIdxOut
            idx = idx + 1
        Next chv

        ws.Range(ws.Cells(2, 1), ws.Cells(1 + totalLinhas, numColunas)).Value2 = saida
    End If

    modLog.LogInfo "Aba " & spec.AbaDestino & ": " & totalLinhas & " linha(s) após upsert.", spec.AbaDestino
End Sub

Private Function MontarChave(ByVal linhaVals As Variant, ByVal posicoesChave() As Long) As String
    Dim resultado As String
    Dim k As Long
    For k = LBound(posicoesChave) To UBound(posicoesChave)
        If posicoesChave(k) = 0 Then
            MontarChave = ""
            Exit Function
        End If
        If Len(resultado) > 0 Then resultado = resultado & "|"
        resultado = resultado & Trim$(CStr(linhaVals(posicoesChave(k))))
    Next k
    MontarChave = resultado
End Function

' Lê uma aba inteira (cabeçalho na linha 1 + dados) como array 2D, para uso do modJoinBuilder.
' Retorna Empty se a aba não existir ou estiver vazia.
Public Function LerAbaComoArray(ByVal nomeAba As String) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomeAba)
    On Error GoTo 0
    If ws Is Nothing Then
        LerAbaComoArray = Empty
        Exit Function
    End If

    Dim ultimaLinha As Long, ultimaColuna As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If ultimaLinha < 1 Or ultimaColuna < 1 Then
        LerAbaComoArray = Empty
        Exit Function
    End If

    LerAbaComoArray = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaLinha, ultimaColuna)).Value2
End Function
