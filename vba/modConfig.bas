Attribute VB_Name = "modConfig"
Option Explicit

' Lê Config_SQVI e Config_Geral (abas do workbook) para estruturas em memória.
' Config_SQVI é a fonte única de verdade sobre as 7 SQVIs da cascata - evita
' 7 sub-rotinas quase duplicadas.

' Colunas de Config_SQVI: A=SQVIName B=Titulo C=Tabela D=FELD E=SELE
'                          F=CamposObrigatoriosFELD G=AbaDestino H=ChaveUpsert I=PodeCriarAutomaticamente
Public Function LoadSQVISpecs() As Object   ' Dictionary(SQVIName -> clsSQVISpec)
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ABA_CONFIG_SQVI)

    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim spec As clsSQVISpec
    For i = 2 To ultimaLinha
        If Len(Trim$(ws.Cells(i, 1).Value)) > 0 Then
            ' IMPORTANTE: criar uma instância NOVA a cada linha. Usar
            ' "Dim spec As New clsSQVISpec" aqui seria um bug clássico do VBA:
            ' o "As New" instancia o objeto UMA única vez por chamada da função,
            ' então todas as linhas acabariam apontando para o MESMO objeto (o da
            ' última linha lida), fazendo specs("-RGVS-VBAP") devolver o VBRK.
            Set spec = New clsSQVISpec
            spec.SQVIName = Trim$(ws.Cells(i, 1).Value)
            spec.Titulo = Trim$(ws.Cells(i, 2).Value)
            spec.Tabela = Trim$(ws.Cells(i, 3).Value)
            Set spec.FELD = modUtils.SplitTrim(CStr(ws.Cells(i, 4).Value))
            Set spec.SELE = modUtils.SplitTrim(CStr(ws.Cells(i, 5).Value))
            Set spec.CamposObrigatoriosFELD = modUtils.SplitTrim(CStr(ws.Cells(i, 6).Value))
            spec.AbaDestino = Trim$(ws.Cells(i, 7).Value)
            Set spec.ChaveUpsert = modUtils.SplitTrim(CStr(ws.Cells(i, 8).Value))
            Dim textoPodeCriar As String
            textoPodeCriar = UCase$(Trim$(CStr(ws.Cells(i, 9).Value)))
            spec.PodeCriarAutomaticamente = (textoPodeCriar = "VERDADEIRO" Or textoPodeCriar = "TRUE" Or textoPodeCriar = "1")

            resultado.Add spec.SQVIName, spec
        End If
    Next i

    Set LoadSQVISpecs = resultado
End Function

' Colunas de Config_Geral: A=Parametro B=Valor
Public Function LoadGlobalSettings() As Object   ' Dictionary(Parametro -> Valor texto)
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ABA_CONFIG_GERAL)

    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To ultimaLinha
        Dim chave As String
        chave = Trim$(ws.Cells(i, 1).Value)
        If Len(chave) > 0 Then
            resultado(chave) = Trim$(CStr(ws.Cells(i, 2).Value))
        End If
    Next i

    Set LoadGlobalSettings = resultado
End Function

Public Function ObterPlantas(ByVal settings As Object) As Collection
    Set ObterPlantas = modUtils.SplitTrim(CStr(settings(PARM_PLANTAS)))
End Function

Public Function ObterJanelaDias(ByVal settings As Object) As Long
    ObterJanelaDias = CLng(settings(PARM_JANELA_DIAS))
End Function

Public Function ObterMaxLinhasSeguranca(ByVal settings As Object) As Long
    If settings.Exists(PARM_MAX_LINHAS_SEGURANCA) Then
        ObterMaxLinhasSeguranca = CLng(settings(PARM_MAX_LINHAS_SEGURANCA))
    Else
        ObterMaxLinhasSeguranca = 20000
    End If
End Function

Public Function ObterCaminhoDownload(ByVal settings As Object) As String
    ObterCaminhoDownload = settings(PARM_CAMINHO_DOWNLOAD)
End Function

Public Function ObterCaminhoLog(ByVal settings As Object) As String
    ObterCaminhoLog = settings(PARM_CAMINHO_CONSOLIDADA) & "\" & settings(PARM_ARQUIVO_LOG)
End Function
