Attribute VB_Name = "modCascade"
Option Explicit

' Orquestrador da cascata: Pedido (VBAP) -> VBFA -> Entrega (LIPS) -> ramifica em
' {Embarque (VTTP->VTTK), Faturamento (VBFA->VBRP->VBRK)}. Cada hop filtra pelas
' chaves derivadas do hop anterior (nunca UNION ALL). Falha em um hop é logada e
' PULA só aquele ramo - não aborta o run inteiro.

Private mSpecs As Object
Private mSettings As Object
Private mCaminhoDownload As String
Private mPlantas As Collection
Private mJanelaDias As Long
Private mVbtypEntrega As String
Private mVbtypFatura As String

Private mPedidosProcessados As Long
Private mPedidosSemEntrega As Long
Private mHopsPulados As Long

Public Sub ExecutarCascata()
    On Error GoTo ErroFatal

    Set mSpecs = modConfig.LoadSQVISpecs()
    Set mSettings = modConfig.LoadGlobalSettings()
    modLog.InicializarLog modConfig.ObterCaminhoLog(mSettings)

    mCaminhoDownload = modConfig.ObterCaminhoDownload(mSettings)
    Set mPlantas = modConfig.ObterPlantas(mSettings)
    mJanelaDias = modConfig.ObterJanelaDias(mSettings)
    mVbtypEntrega = mSettings(PARM_VBTYP_ENTREGA)
    mVbtypFatura = mSettings(PARM_VBTYP_FATURA)

    mPedidosProcessados = 0
    mPedidosSemEntrega = 0
    mHopsPulados = 0

    modLog.LogInfo "===== Início da execução ====="

    If Not modSAPConnection.ConectarSAP() Then
        modLog.Falhar "Não foi possível conectar a uma sessão SAP GUI já aberta. Abra e faça login no SAP antes de rodar."
        Exit Sub
    End If

    ' ---------- Passo 0: Pedido (VBAP) ----------
    Dim specPedido As clsSQVISpec
    Set specPedido = mSpecs(SQVI_PEDIDO)

    If Not modSQVIEngine.EnsureSQVIRunnable(specPedido) Then
        modLog.Falhar "Não foi possível preparar a SQVI de Pedido (" & SQVI_PEDIDO & "). Veja a aba Log para detalhes."
        Exit Sub
    End If

    Dim filtrosPedido As New Collection
    filtrosPedido.Add NovoFiltroIncluirMultiplo("WERKS", mPlantas)
    filtrosPedido.Add NovoFiltroExcluirFlag("LOEKZ", "X")
    filtrosPedido.Add NovoFiltroIntervaloData("ERDAT", modUtils.DataSAP(Date - mJanelaDias), modUtils.DataSAP(Date))

    Dim resPedido As clsRunResult
    Set resPedido = modSQVIEngine.RunSQVI(specPedido, filtrosPedido, "PEDIDO.xlsx", mCaminhoDownload)

    If Not resPedido.Success Then
        modLog.Falhar "Não foi possível extrair os Pedidos: " & resPedido.MensagemErro
        Exit Sub
    End If

    modExcelWrite.UpsertRows specPedido, resPedido.Dados
    mPedidosProcessados = resPedido.NumLinhas

    If resPedido.NumLinhas = 0 Then
        modLog.LogInfo "Nenhum pedido encontrado no filtro (plantas/data/não-deletado)."
        modJoinBuilder.ConstruirSaidasFinais mSpecs, mSettings
        modLog.Informar "ETL Cascata finalizado. Nenhum pedido encontrado no período filtrado."
        Exit Sub
    End If

    Dim colVBELN_Pedido As Long, colPOSNR_Pedido As Long
    colVBELN_Pedido = modUtils.ColunaDoCampo(specPedido, "VBELN")
    colPOSNR_Pedido = modUtils.ColunaDoCampo(specPedido, "POSNR")

    Dim pedidoVBELNList As Collection
    Set pedidoVBELNList = modUtils.ValoresDistintosPorColuna(resPedido.Dados, 2, colVBELN_Pedido)
    Dim pedidoParesValidos As Object
    Set pedidoParesValidos = modUtils.ParesDistintosDict(resPedido.Dados, 2, colVBELN_Pedido, colPOSNR_Pedido)

    ' ---------- Hop 1: Pedido -> VBFA -> Entrega (LIPS) ----------
    Dim specVBFA As clsSQVISpec
    Set specVBFA = mSpecs(SQVI_VBFA)
    modSQVIEngine.EnsureSQVIRunnable specVBFA

    Dim colVBELV As Long, colPOSNV As Long, colVBELNsub As Long, colVBTYPN As Long
    colVBELV = modUtils.ColunaDoCampo(specVBFA, "VBELV")
    colPOSNV = modUtils.ColunaDoCampo(specVBFA, "POSNV")
    colVBELNsub = modUtils.ColunaDoCampo(specVBFA, "VBELN")
    colVBTYPN = modUtils.ColunaDoCampo(specVBFA, "VBTYP_N")

    Dim entregaVBELNList As New Collection

    If colVBELV = 0 Or colPOSNV = 0 Then
        modLog.LogErro SQVI_VBFA & " não tem VBELV/POSNV cadastrados no FELD (Config_SQVI) - hop Pedido->Entrega pulado. " & _
            "Adicione VBELV e POSNV ao FELD e ao SELE de " & SQVI_VBFA & " (modo Alterar no SAP) e rode de novo.", SQVI_VBFA
        mHopsPulados = mHopsPulados + 1
    Else
        Dim filtrosVBFA1 As New Collection
        filtrosVBFA1.Add NovoFiltroColarMultiplo("VBELV", pedidoVBELNList)

        Dim resVBFA1 As clsRunResult
        Set resVBFA1 = modSQVIEngine.RunSQVI(specVBFA, filtrosVBFA1, "VBFA_HOP1.xlsx", mCaminhoDownload)

        If resVBFA1.Success Then
            If resVBFA1.NumLinhas > 0 Then
                modExcelWrite.UpsertRows specVBFA, resVBFA1.Dados
                Set entregaVBELNList = DocumentosSubsequentesVBFA(resVBFA1.Dados, colVBELV, colPOSNV, colVBELNsub, _
                    colVBTYPN, pedidoParesValidos, mVbtypEntrega)
            End If
        Else
            modLog.LogErro "Hop Pedido->Entrega (VBFA) falhou: " & resVBFA1.MensagemErro, SQVI_VBFA
            mHopsPulados = mHopsPulados + 1
        End If
    End If

    ' ---------- Entrega (LIPS) ----------
    Dim specLIPS As clsSQVISpec
    Set specLIPS = mSpecs(SQVI_LIPS)

    If entregaVBELNList.Count = 0 Then
        modLog.LogInfo "Nenhuma entrega encontrada para os pedidos filtrados neste run.", SQVI_LIPS
        mPedidosSemEntrega = mPedidosProcessados
    Else
        modSQVIEngine.EnsureSQVIRunnable specLIPS

        Dim filtrosLIPS As New Collection
        filtrosLIPS.Add NovoFiltroColarMultiplo("VBELN", entregaVBELNList)

        Dim resLIPS As clsRunResult
        Set resLIPS = modSQVIEngine.RunSQVI(specLIPS, filtrosLIPS, "LIPS.xlsx", mCaminhoDownload)

        If resLIPS.Success Then
            If resLIPS.NumLinhas > 0 Then modExcelWrite.UpsertRows specLIPS, resLIPS.Dados
        Else
            modLog.LogErro "Falha ao extrair LIPS: " & resLIPS.MensagemErro, SQVI_LIPS
            mHopsPulados = mHopsPulados + 1
        End If

        ' ---------- Ramo A: LIPS -> VTTP -> VTTK ----------
        RamoEmbarque entregaVBELNList

        ' ---------- Ramo B: LIPS -> VBFA(hop2) -> VBRP -> VBRK ----------
        RamoFaturamento entregaVBELNList
    End If

    ' ---------- Abas finais (cruzamento a partir do acumulado completo) ----------
    modJoinBuilder.ConstruirSaidasFinais mSpecs, mSettings

    Dim resumo As String
    resumo = "ETL Cascata finalizado." & vbCrLf & _
        "Pedidos processados: " & mPedidosProcessados & vbCrLf & _
        "Pedidos sem entrega neste run: " & mPedidosSemEntrega & vbCrLf & _
        "Hops pulados por gap/erro: " & mHopsPulados & vbCrLf & _
        "Log: " & modConfig.ObterCaminhoLog(mSettings)
    modLog.LogInfo resumo
    modLog.Informar resumo
    Exit Sub

ErroFatal:
    Dim msgErro As String
    msgErro = Err.Description
    modLog.LogErro "Erro inesperado interrompeu a cascata: " & msgErro & _
        " - tentando ainda assim reconstruir as abas finais com o que já foi acumulado."
    On Error Resume Next
    modJoinBuilder.ConstruirSaidasFinais mSpecs, mSettings
    On Error GoTo 0
    modLog.Falhar "ETL Cascata interrompida por erro inesperado: " & msgErro & vbCrLf & _
        "Pedidos processados antes da interrupção: " & mPedidosProcessados & vbCrLf & "Verifique a aba Log."
End Sub

' ===================== Ramos =====================

Private Sub RamoEmbarque(ByVal entregaVBELNList As Collection)
    On Error GoTo TratarErro
    If entregaVBELNList.Count = 0 Then Exit Sub

    Dim specVTTP As clsSQVISpec
    Set specVTTP = mSpecs(SQVI_VTTP)
    modSQVIEngine.EnsureSQVIRunnable specVTTP

    Dim colTKNUM As Long
    colTKNUM = modUtils.ColunaDoCampo(specVTTP, "TKNUM")
    If colTKNUM = 0 Then
        modLog.LogErro SQVI_VTTP & " não tem TKNUM cadastrado no FELD (Config_SQVI) - ramo de Embarque pulado. " & _
            "Adicione TKNUM ao FELD de " & SQVI_VTTP & " (modo Alterar no SAP) e rode de novo.", SQVI_VTTP
        mHopsPulados = mHopsPulados + 1
        Exit Sub
    End If

    Dim filtrosVTTP As New Collection
    filtrosVTTP.Add NovoFiltroColarMultiplo("VBELN", entregaVBELNList)

    Dim resVTTP As clsRunResult
    Set resVTTP = modSQVIEngine.RunSQVI(specVTTP, filtrosVTTP, "VTTP.xlsx", mCaminhoDownload)

    If Not resVTTP.Success Then
        modLog.LogErro "Falha ao extrair VTTP: " & resVTTP.MensagemErro, SQVI_VTTP
        mHopsPulados = mHopsPulados + 1
        Exit Sub
    End If
    If resVTTP.NumLinhas = 0 Then
        modLog.LogInfo "Nenhum embarque encontrado para as entregas deste run.", SQVI_VTTP
        Exit Sub
    End If

    modExcelWrite.UpsertRows specVTTP, resVTTP.Dados

    Dim tknumList As Collection
    Set tknumList = modUtils.ValoresDistintosPorColuna(resVTTP.Dados, 2, colTKNUM)
    If tknumList.Count = 0 Then Exit Sub

    Dim specVTTK As clsSQVISpec
    Set specVTTK = mSpecs(SQVI_VTTK)
    modSQVIEngine.EnsureSQVIRunnable specVTTK

    Dim filtrosVTTK As New Collection
    filtrosVTTK.Add NovoFiltroColarMultiplo("TKNUM", tknumList)

    Dim resVTTK As clsRunResult
    Set resVTTK = modSQVIEngine.RunSQVI(specVTTK, filtrosVTTK, "VTTK.xlsx", mCaminhoDownload)

    If resVTTK.Success Then
        If resVTTK.NumLinhas > 0 Then modExcelWrite.UpsertRows specVTTK, resVTTK.Dados
    Else
        modLog.LogErro "Falha ao extrair VTTK: " & resVTTK.MensagemErro, SQVI_VTTK
        mHopsPulados = mHopsPulados + 1
    End If
    Exit Sub

TratarErro:
    modLog.LogErro "Erro inesperado no ramo de Embarque (VTTP/VTTK): " & Err.Description & _
        " - ramo pulado, restante da cascata continua.", SQVI_VTTP
    mHopsPulados = mHopsPulados + 1
End Sub

Private Sub RamoFaturamento(ByVal entregaVBELNList As Collection)
    On Error GoTo TratarErro
    If entregaVBELNList.Count = 0 Then Exit Sub

    Dim specVBFA As clsSQVISpec
    Set specVBFA = mSpecs(SQVI_VBFA)

    Dim colVBELV As Long, colPOSNV As Long, colVBELNsub As Long, colVBTYPN As Long
    colVBELV = modUtils.ColunaDoCampo(specVBFA, "VBELV")
    colPOSNV = modUtils.ColunaDoCampo(specVBFA, "POSNV")
    colVBELNsub = modUtils.ColunaDoCampo(specVBFA, "VBELN")
    colVBTYPN = modUtils.ColunaDoCampo(specVBFA, "VBTYP_N")

    If colVBELV = 0 Or colPOSNV = 0 Then
        modLog.LogErro SQVI_VBFA & " não tem VBELV/POSNV cadastrados no FELD (Config_SQVI) - ramo de Faturamento pulado.", SQVI_VBFA
        mHopsPulados = mHopsPulados + 1
        Exit Sub
    End If

    ' chaves válidas de Entrega (para o filtro exato client-side dentro de DocumentosSubsequentesVBFA)
    Dim dadosLIPSAtual As Variant
    dadosLIPSAtual = modExcelWrite.LerAbaComoArray(ABA_LIPS)
    Dim specLIPS As clsSQVISpec
    Set specLIPS = mSpecs(SQVI_LIPS)
    Dim colVBELN_LIPS As Long, colPOSNR_LIPS As Long
    colVBELN_LIPS = modUtils.ColunaDoCampo(specLIPS, "VBELN")
    colPOSNR_LIPS = modUtils.ColunaDoCampo(specLIPS, "POSNR")

    Dim entregaParesValidos As Object
    If modUtils.EhArrayValido(dadosLIPSAtual) Then
        Set entregaParesValidos = modUtils.ParesDistintosDict(dadosLIPSAtual, 2, colVBELN_LIPS, colPOSNR_LIPS)
    Else
        Set entregaParesValidos = CreateObject("Scripting.Dictionary")
    End If

    Dim filtrosVBFA2 As New Collection
    filtrosVBFA2.Add NovoFiltroColarMultiplo("VBELV", entregaVBELNList)

    Dim resVBFA2 As clsRunResult
    Set resVBFA2 = modSQVIEngine.RunSQVI(specVBFA, filtrosVBFA2, "VBFA_HOP2.xlsx", mCaminhoDownload)

    If Not resVBFA2.Success Then
        modLog.LogErro "Hop Entrega->Fatura (VBFA) falhou: " & resVBFA2.MensagemErro, SQVI_VBFA
        mHopsPulados = mHopsPulados + 1
        Exit Sub
    End If
    If resVBFA2.NumLinhas = 0 Then
        modLog.LogInfo "Nenhum documento de faturamento encontrado para as entregas deste run.", SQVI_VBFA
        Exit Sub
    End If

    modExcelWrite.UpsertRows specVBFA, resVBFA2.Dados

    Dim faturaVBELNList As Collection
    Set faturaVBELNList = DocumentosSubsequentesVBFA(resVBFA2.Dados, colVBELV, colPOSNV, colVBELNsub, _
        colVBTYPN, entregaParesValidos, mVbtypFatura)

    If faturaVBELNList.Count = 0 Then Exit Sub

    Dim specVBRP As clsSQVISpec
    Set specVBRP = mSpecs(SQVI_VBRP)
    modSQVIEngine.EnsureSQVIRunnable specVBRP

    Dim filtrosVBRP As New Collection
    filtrosVBRP.Add NovoFiltroColarMultiplo("VBELN", faturaVBELNList)

    Dim resVBRP As clsRunResult
    Set resVBRP = modSQVIEngine.RunSQVI(specVBRP, filtrosVBRP, "VBRP.xlsx", mCaminhoDownload)

    If Not resVBRP.Success Then
        modLog.LogErro "Falha ao extrair VBRP: " & resVBRP.MensagemErro, SQVI_VBRP
        mHopsPulados = mHopsPulados + 1
        Exit Sub
    End If
    If resVBRP.NumLinhas = 0 Then Exit Sub

    modExcelWrite.UpsertRows specVBRP, resVBRP.Dados

    Dim colVBELN_VBRP As Long
    colVBELN_VBRP = modUtils.ColunaDoCampo(specVBRP, "VBELN")
    Dim vbrkList As Collection
    Set vbrkList = modUtils.ValoresDistintosPorColuna(resVBRP.Dados, 2, colVBELN_VBRP)
    If vbrkList.Count = 0 Then Exit Sub

    Dim specVBRK As clsSQVISpec
    Set specVBRK = mSpecs(SQVI_VBRK)
    modSQVIEngine.EnsureSQVIRunnable specVBRK

    Dim filtrosVBRK As New Collection
    filtrosVBRK.Add NovoFiltroColarMultiplo("VBELN", vbrkList)

    Dim resVBRK As clsRunResult
    Set resVBRK = modSQVIEngine.RunSQVI(specVBRK, filtrosVBRK, "VBRK.xlsx", mCaminhoDownload)

    If resVBRK.Success Then
        If resVBRK.NumLinhas > 0 Then modExcelWrite.UpsertRows specVBRK, resVBRK.Dados
    Else
        modLog.LogErro "Falha ao extrair VBRK: " & resVBRK.MensagemErro, SQVI_VBRK
        mHopsPulados = mHopsPulados + 1
    End If
    Exit Sub

TratarErro:
    modLog.LogErro "Erro inesperado no ramo de Faturamento (VBFA/VBRP/VBRK): " & Err.Description & _
        " - ramo pulado, restante da cascata continua.", SQVI_VBFA
    mHopsPulados = mHopsPulados + 1
End Sub

' ===================== Derivação de chaves via VBFA =====================

' A partir de linhas de VBFA (filtradas no SAP só por VBELV), aplica um filtro exato
' client-side: só aceita pares (VBELV,POSNV) que batem com paresValidos (chaves reais
' do nível anterior) E cuja categoria do documento subsequente (VBTYP_N) bate com a
' esperada (ex: "J" para entrega, "M" para fatura) - evita vazar itens de outro
' pedido/planta que por acaso compartilhem o mesmo número de documento, e evita
' misturar tipos de documento subsequente que não interessam a este hop.
Private Function DocumentosSubsequentesVBFA(ByVal dados As Variant, ByVal colVBELV As Long, ByVal colPOSNV As Long, _
        ByVal colVBELNsub As Long, ByVal colVBTYPN As Long, ByVal paresValidos As Object, ByVal vbtypEsperado As String) As Collection
    Dim resultado As New Collection
    Dim vistos As Object
    Set vistos = CreateObject("Scripting.Dictionary")

    If Not modUtils.EhArrayValido(dados) Then
        Set DocumentosSubsequentesVBFA = resultado
        Exit Function
    End If

    Dim i As Long
    For i = LBound(dados, 1) + 1 To UBound(dados, 1)
        Dim chave As String
        chave = Trim$(CStr(dados(i, colVBELV))) & "|" & Trim$(CStr(dados(i, colPOSNV)))

        Dim tipoOk As Boolean
        If colVBTYPN = 0 Or Len(vbtypEsperado) = 0 Then
            tipoOk = True
        Else
            tipoOk = (UCase$(Trim$(CStr(dados(i, colVBTYPN)))) = UCase$(vbtypEsperado))
        End If

        If tipoOk And paresValidos.Exists(chave) Then
            Dim v As String
            v = Trim$(CStr(dados(i, colVBELNsub)))
            If Len(v) > 0 Then
                If Not vistos.Exists(v) Then
                    vistos.Add v, True
                    resultado.Add v
                End If
            End If
        End If
    Next i

    Set DocumentosSubsequentesVBFA = resultado
End Function

' ===================== Construtores de filtro =====================

Private Function NovoFiltroIncluirMultiplo(ByVal campoLogico As String, ByVal valores As Collection) As clsFiltro
    Dim f As New clsFiltro
    f.CampoLogico = campoLogico
    f.Modo = FILTRO_INCLUIR_MULTIPLO
    Set f.Valores = valores
    Set NovoFiltroIncluirMultiplo = f
End Function

Private Function NovoFiltroColarMultiplo(ByVal campoLogico As String, ByVal valores As Collection) As clsFiltro
    Dim f As New clsFiltro
    f.CampoLogico = campoLogico
    f.Modo = FILTRO_COLAR_MULTIPLO
    Set f.Valores = valores
    Set NovoFiltroColarMultiplo = f
End Function

Private Function NovoFiltroExcluirFlag(ByVal campoLogico As String, ByVal valorExcluir As String) As clsFiltro
    Dim f As New clsFiltro
    f.CampoLogico = campoLogico
    f.Modo = FILTRO_EXCLUIR_FLAG
    f.Valores.Add valorExcluir
    Set NovoFiltroExcluirFlag = f
End Function

Private Function NovoFiltroIntervaloData(ByVal campoLogico As String, ByVal dataIni As String, ByVal dataFim As String) As clsFiltro
    Dim f As New clsFiltro
    f.CampoLogico = campoLogico
    f.Modo = FILTRO_INTERVALO_DATA
    f.ValorInicial = dataIni
    f.ValorFinal = dataFim
    Set NovoFiltroIntervaloData = f
End Function
