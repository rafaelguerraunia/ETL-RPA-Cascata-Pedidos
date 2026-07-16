Attribute VB_Name = "modSQVIEngine"
Option Explicit

' Motor genérico de SQVI: verifica/cria, aplica filtros na tela de seleção,
' executa, exporta para Excel e lê de volta. Parametrizado por clsSQVISpec -
' generalização direta de CriarSQVIFaltante/AddCampo/ExecutarSQVI_RESB do RPA BASE.vbs.

' ===================== Verificar / criar =====================

' Garante que a SQVI existe e está pronta para ser executada. Não altera SQVIs
' já existentes (só cria as que estão marcadas como PodeCriarAutomaticamente e
' ainda não existem). Loga (WARN) campos de seleção esperados que não aparecem
' na tela, mas não impede o restante do fluxo - a falha real de escopo é
' verificada na hora de aplicar cada filtro em RunSQVI.
Public Function EnsureSQVIRunnable(ByVal spec As clsSQVISpec) As Boolean
    modSAPConnection.NavegarTransacao "/nsqvi"
    modSAPConnection.Definir "wnd[0]/usr/ctxtRS38R-QNUM", spec.SQVIName
    modSAPConnection.Pressionar "wnd[0]/usr/btnP1"
    modSAPConnection.AguardarSAP

    If modSAPConnection.StatusBarIndicaNaoExiste() Then
        If Not spec.PodeCriarAutomaticamente Then
            modLog.LogErro spec.SQVIName & " não existe no SAP e não está marcada para criação automática.", spec.SQVIName
            EnsureSQVIRunnable = False
            Exit Function
        End If

        CreateSQVI spec

        modSAPConnection.NavegarTransacao "/nsqvi"
        modSAPConnection.Definir "wnd[0]/usr/ctxtRS38R-QNUM", spec.SQVIName
        modSAPConnection.Pressionar "wnd[0]/usr/btnP1"
        modSAPConnection.AguardarSAP

        If modSAPConnection.StatusBarIndicaNaoExiste() Then
            modLog.LogErro "Falha ao criar " & spec.SQVIName & " - ainda não encontrada após a tentativa de criação.", spec.SQVIName
            EnsureSQVIRunnable = False
            Exit Function
        End If
    End If

    On Error Resume Next
    modSAPConnection.Definir "wnd[0]/usr/txtMAX_SEL", ""   ' remove limite de linhas de seleção, se existir
    On Error GoTo 0

    Dim i As Long
    Dim faltando As New Collection
    For i = 1 To spec.SELE.Count
        If Not CampoSeleExisteNaTela(i) Then faltando.Add spec.SELE(i)
    Next i

    If faltando.Count > 0 Then
        modLog.LogWarn spec.SQVIName & " sem os seguintes campos de seleção na tela (posição esperada entre parênteses): " & _
            DescreverFaltantes(faltando, spec) & " -> adicione em SQVI (modo Alterar), NESSA ORDEM, ao final da lista já existente.", _
            spec.SQVIName
    End If

    EnsureSQVIRunnable = True
End Function

Private Function DescreverFaltantes(ByVal faltando As Collection, ByVal spec As clsSQVISpec) As String
    Dim item As Variant, resultado As String
    For Each item In faltando
        If Len(resultado) > 0 Then resultado = resultado & ", "
        resultado = resultado & CStr(item) & " (posição " & spec.PosicaoSELE(CStr(item)) & ")"
    Next item
    DescreverFaltantes = resultado
End Function

Private Function CampoSeleExisteNaTela(ByVal posicao As Long) As Boolean
    Dim sufixo As String
    sufixo = "SP$0000" & CStr(posicao)
    CampoSeleExisteNaTela = modSAPConnection.ElementExists("wnd[0]/usr/ctxt" & sufixo & "-LOW") _
        Or modSAPConnection.ElementExists("wnd[0]/usr/btn%_" & sufixo & "_%_APP_%-VALU_PUSH")
End Function

' ===================== Criação de SQVI (só quando ainda não existe) =====================

Public Sub CreateSQVI(ByVal spec As clsSQVISpec)
    modLog.LogInfo "Criando SQVI " & spec.SQVIName & " (tabela " & spec.Tabela & ")...", spec.SQVIName

    modSAPConnection.Pressionar "wnd[0]/usr/btnP7"    ' Criar
    modSAPConnection.Definir "wnd[1]/usr/txtRS38R-HDTITLE", spec.Titulo
    modSAPConnection.Definir "wnd[1]/usr/subSUBSOURCE:SAPMS38R:3110/ctxtRS38Q-DDNAME", spec.Tabela
    modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[0]"
    modSAPConnection.AguardarSAP

    modSAPConnection.Pressionar "wnd[0]/tbar[1]/btn[17]"   ' gerar lista de campos (abre aba FELD)

    Dim campo As Variant
    For Each campo In spec.FELD
        AddCampo "FELD", CStr(campo)
    Next campo

    modSAPConnection.Selecionar "wnd[0]/usr/tabsTAB100/tabpSELE"
    modSAPConnection.Pressionar "wnd[0]/tbar[1]/btn[17]"

    For Each campo In spec.SELE
        AddCampo "SELE", CStr(campo)
    Next campo

    modSAPConnection.Pressionar "wnd[0]/tbar[0]/btn[11]"   ' gerar/salvar a query
    modSAPConnection.Pressionar "wnd[0]/tbar[0]/btn[3]"    ' voltar

    On Error Resume Next
    modSAPConnection.Pressionar "wnd[1]/usr/btnBUTTON_1"   ' confirmar popup de geração, se aparecer
    On Error GoTo 0

    modLog.LogInfo "SQVI " & spec.SQVIName & " criada: " & spec.FELD.Count & " campos de saída, " & _
        spec.SELE.Count & " campos de seleção.", spec.SQVIName
End Sub

Private Sub AddCampo(ByVal aba As String, ByVal campo As String)
    Dim subtela As String
    If aba = "FELD" Then
        subtela = "3120"
    Else
        subtela = "3140"
    End If
    Dim base As String
    base = "wnd[0]/usr/tabsTAB100/tabp" & aba & "/ssubSUBSOURCE:SAPMS38R:" & subtela

    modSAPConnection.Pressionar base & "/btnRSEA"
    modSAPConnection.Definir "wnd[1]/usr/txtTC_SORTPOOLENTRY-NAME", campo
    modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[71]"
    modSAPConnection.Pressionar base & "/btnMVLE"
End Sub

' ===================== Executar =====================

' Roda a SQVI com os filtros informados, exporta para Excel, lê os dados de volta.
' Se algum filtro não puder ser aplicado (campo ausente na tela ou lista de chaves
' vazia), a execução é ABORTADA (resultado.Success = False) em vez de rodar sem o
' escopo correto - evita puxar a tabela inteira por engano.
Public Function RunSQVI(ByVal spec As clsSQVISpec, ByVal filtros As Collection, _
                         ByVal nomeArquivoExport As String, ByVal caminhoDownload As String) As clsRunResult
    Dim resultado As New clsRunResult

    modUtils.DeletarSeExistir caminhoDownload & "\" & nomeArquivoExport

    modSAPConnection.NavegarTransacao "/nsqvi"
    modSAPConnection.Definir "wnd[0]/usr/ctxtRS38R-QNUM", spec.SQVIName
    modSAPConnection.Pressionar "wnd[0]/usr/btnP1"
    modSAPConnection.AguardarSAP

    If modSAPConnection.StatusBarIndicaNaoExiste() Then
        resultado.MensagemErro = spec.SQVIName & " não encontrada ao tentar executar (era esperado que já existisse a essa altura)."
        modLog.LogErro resultado.MensagemErro, spec.SQVIName
        Set RunSQVI = resultado
        Exit Function
    End If

    On Error Resume Next
    modSAPConnection.Definir "wnd[0]/usr/txtMAX_SEL", ""
    On Error GoTo 0

    Dim todosAplicados As Boolean
    todosAplicados = True

    Dim f As clsFiltro
    For Each f In filtros
        Dim posicao As Long
        posicao = spec.PosicaoSELE(f.CampoLogico)
        If posicao = 0 Then
            modLog.LogErro spec.SQVIName & ": campo de filtro '" & f.CampoLogico & "' não está definido no SELE da spec.", spec.SQVIName
            resultado.CamposSELEFaltando.Add f.CampoLogico
            todosAplicados = False
        ElseIf Not AplicarFiltro(posicao, f) Then
            modLog.LogErro spec.SQVIName & ": não foi possível aplicar o filtro '" & f.CampoLogico & "' (posição " & posicao & ").", spec.SQVIName
            resultado.CamposSELEFaltando.Add f.CampoLogico
            todosAplicados = False
        End If
    Next f

    If Not todosAplicados Then
        resultado.Success = False
        resultado.ModoDegradado = True
        resultado.MensagemErro = "Um ou mais filtros não puderam ser aplicados em " & spec.SQVIName & _
            " - execução abortada para não trazer dados fora do escopo pretendido."
        Set RunSQVI = resultado
        Exit Function
    End If

    modSAPConnection.Pressionar "wnd[0]/tbar[1]/btn[8]"   ' executar
    modSAPConnection.AguardarSAP

    ExportarPlanilhaAtual nomeArquivoExport, caminhoDownload

    If Not modUtils.AguardarArquivo(caminhoDownload & "\" & nomeArquivoExport, 90) Then
        resultado.MensagemErro = "Timeout esperando o arquivo " & nomeArquivoExport & " ser gerado."
        modLog.LogErro resultado.MensagemErro, spec.SQVIName
        Set RunSQVI = resultado
        Exit Function
    End If

    modUtils.FecharPlanilhaSeAberta nomeArquivoExport

    Dim wbTemp As Workbook
    On Error Resume Next
    Set wbTemp = Nothing
    Set wbTemp = Application.Workbooks.Open(caminhoDownload & "\" & nomeArquivoExport, 0, True)
    On Error GoTo 0

    If wbTemp Is Nothing Then
        resultado.MensagemErro = "Não foi possível abrir o arquivo exportado " & nomeArquivoExport
        modLog.LogErro resultado.MensagemErro, spec.SQVIName
        Set RunSQVI = resultado
        Exit Function
    End If

    Dim bruto As Variant
    bruto = wbTemp.Sheets(1).UsedRange.Value2
    wbTemp.Close False

    resultado.Dados = bruto

    If modUtils.EhArrayValido(bruto) Then
        resultado.NumLinhas = UBound(bruto, 1) - LBound(bruto, 1)   ' descontando a linha de cabeçalho
        If resultado.NumLinhas < 0 Then resultado.NumLinhas = 0

        Dim numColunasEsperadas As Long, numColunasReais As Long
        numColunasEsperadas = spec.FELD.Count
        numColunasReais = UBound(bruto, 2) - LBound(bruto, 2) + 1
        If numColunasReais < numColunasEsperadas Then
            Dim j As Long
            For j = numColunasReais + 1 To numColunasEsperadas
                resultado.CamposFELDFaltando.Add spec.FELD(j)
            Next j
            modLog.LogWarn spec.SQVIName & " retornou " & numColunasReais & " colunas, esperava " & numColunasEsperadas & _
                " - possivelmente faltam estes campos no FELD: " & modUtils.JoinColecao(resultado.CamposFELDFaltando), spec.SQVIName
        End If
    Else
        resultado.NumLinhas = 0
    End If

    modLog.LogInfo spec.SQVIName & ": " & resultado.NumLinhas & " linha(s) retornada(s).", spec.SQVIName

    resultado.Success = True
    Set RunSQVI = resultado
End Function

Private Sub ExportarPlanilhaAtual(ByVal nomeArquivo As String, ByVal caminhoDownload As String)
    ' Menu "Lista > Exportar > Planilha": itens de menu (GuiMenu) usam .Select,
    ' NUNCA .press - chamar .press aqui gera "Object doesn't support this
    ' property or method" (erro 438). Por isso usamos Selecionar (.Select) no
    ' menu e Pressionar (.press) só nos botões da janela seguinte.
    modSAPConnection.Selecionar "wnd[0]/mbar/menu[0]/menu[3]/menu[0]"   ' Lista > Exportar > Planilha
    modSAPConnection.Definir "wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME", nomeArquivo
    modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[20]"
    modSAPConnection.Definir "wnd[1]/usr/ctxtDY_PATH", caminhoDownload
    modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[0]"
    modSAPConnection.AguardarSAP
End Sub

' ===================== Aplicação de filtros na tela de seleção =====================
' Reproduz exatamente os padrões de clique confirmados em RPA BASE.vbs para cada
' modo de filtro (multi-inclusão digitada, colar via clipboard, excluir 1 flag,
' intervalo de datas). Retorna False se o elemento de tela esperado não existir.

Private Function AplicarFiltro(ByVal posicao As Long, ByVal f As clsFiltro) As Boolean
    Dim sufixo As String
    sufixo = "SP$0000" & CStr(posicao)

    Select Case f.Modo

        Case FILTRO_INTERVALO_DATA
            Dim idLow As String, idHigh As String
            idLow = "wnd[0]/usr/ctxt" & sufixo & "-LOW"
            idHigh = "wnd[0]/usr/ctxt" & sufixo & "-HIGH"
            If Not modSAPConnection.ElementExists(idLow) Then
                AplicarFiltro = False
                Exit Function
            End If
            modSAPConnection.Definir idLow, f.ValorInicial
            modSAPConnection.Definir idHigh, f.ValorFinal
            AplicarFiltro = True

        Case FILTRO_INCLUIR_MULTIPLO
            Dim idPush As String
            idPush = "wnd[0]/usr/btn%_" & sufixo & "_%_APP_%-VALU_PUSH"
            If Not modSAPConnection.ElementExists(idPush) Then
                AplicarFiltro = False
                Exit Function
            End If
            modSAPConnection.Pressionar idPush
            modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[16]"

            Dim i As Long, v As Variant
            i = 0
            For Each v In f.Valores
                modSAPConnection.Definir _
                    "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]", _
                    CStr(v)
                i = i + 1
            Next v

            modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[8]"
            AplicarFiltro = True

        Case FILTRO_EXCLUIR_FLAG
            Dim idPush2 As String
            idPush2 = "wnd[0]/usr/btn%_" & sufixo & "_%_APP_%-VALU_PUSH"
            If Not modSAPConnection.ElementExists(idPush2) Then
                AplicarFiltro = False
                Exit Function
            End If
            If f.Valores.Count = 0 Then
                AplicarFiltro = False
                Exit Function
            End If
            modSAPConnection.Pressionar idPush2
            modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[16]"
            modSAPConnection.Selecionar "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV"
            modSAPConnection.Definir _
                "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]", _
                CStr(f.Valores(1))
            modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[16]"
            modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[8]"
            AplicarFiltro = True

        Case FILTRO_COLAR_MULTIPLO
            Dim idPush3 As String
            idPush3 = "wnd[0]/usr/btn%_" & sufixo & "_%_APP_%-VALU_PUSH"
            If Not modSAPConnection.ElementExists(idPush3) Then
                AplicarFiltro = False
                Exit Function
            End If
            If f.Valores.Count = 0 Then
                ' Lista de chaves vazia: NUNCA prosseguir sem filtro (traria a tabela inteira).
                ' O chamador (modCascade) deveria ter detectado isso antes e pulado o hop -
                ' isto aqui é uma segunda trava de segurança.
                AplicarFiltro = False
                Exit Function
            End If
            modSAPConnection.Pressionar idPush3
            modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[16]"
            modUtils.CopiarClipboard JoinColecaoQuebraLinha(f.Valores)
            modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[24]"
            modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[16]"
            modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[8]"
            AplicarFiltro = True

        Case Else
            modLog.LogErro "Modo de filtro desconhecido: " & f.Modo
            AplicarFiltro = False
    End Select
End Function

Private Function JoinColecaoQuebraLinha(ByVal c As Collection) As String
    Dim resultado As String
    Dim item As Variant
    For Each item In c
        If Len(resultado) > 0 Then resultado = resultado & vbCrLf
        resultado = resultado & CStr(item)
    Next item
    JoinColecaoQuebraLinha = resultado
End Function
