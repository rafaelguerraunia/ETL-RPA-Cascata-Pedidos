Attribute VB_Name = "modJoinBuilder"
Option Explicit

' Constrói as duas abas finais (Saída Enxuta / Saída Completa) cruzando TODAS as
' abas de cascata já acumuladas (não só os dados desta execução), via índices
' Dictionary(chave -> linha(s)) para evitar O(n²). Pedidos sem entrega/fatura/
' embarque ainda aparecem, com as colunas seguintes em branco.

Private mSpecPedido As clsSQVISpec, mSpecLIPS As clsSQVISpec, mSpecVTTP As clsSQVISpec
Private mSpecVTTK As clsSQVISpec, mSpecVBFA As clsSQVISpec, mSpecVBRP As clsSQVISpec, mSpecVBRK As clsSQVISpec

Private mDadosPedido As Variant, mDadosLIPS As Variant, mDadosVTTP As Variant, mDadosVTTK As Variant
Private mDadosVBFA As Variant, mDadosVBRP As Variant, mDadosVBRK As Variant

Private mColPedVBELN As Long, mColPedPOSNR As Long
Private mColVBFA_VBELV As Long, mColVBFA_POSNV As Long, mColVBFA_VBELN As Long, mColVBFA_VBTYPN As Long
Private mColLIPS_VBELN As Long, mColLIPS_POSNR As Long
Private mColVTTP_VBELN As Long, mColVTTP_TKNUM As Long
Private mColVTTK_TKNUM As Long
Private mColVBRP_VBELN As Long
Private mColVBRK_VBELN As Long

Private mIdxVBFAporPrecedente As Object   ' "VBELV|POSNV" -> Collection(linha)
Private mIdxLIPSporVBELN As Object        ' "VBELN" -> Collection(linha)
Private mIdxVTTPporVBELN As Object        ' "VBELN" -> Collection(linha)
Private mIdxVTTKporTKNUM As Object        ' "TKNUM" -> linha (1:1)
Private mIdxVBRPporVBELN As Object        ' "VBELN" -> Collection(linha)
Private mIdxVBRKporVBELN As Object        ' "VBELN" -> linha (1:1)

Private mVbtypEntrega As String
Private mVbtypFatura As String

' ===================== Entrada pública =====================

Public Sub ConstruirSaidasFinais(ByVal specs As Object, ByVal settings As Object)
    CarregarEstado specs, settings

    If Not modUtils.EhArrayValido(mDadosPedido) Then
        modLog.LogInfo "ConstruirSaidasFinais: aba PEDIDO vazia, nada a cruzar."
        Exit Sub
    End If

    Dim resolvidas As New Collection

    Dim iPed As Long
    For iPed = LBound(mDadosPedido, 1) + 1 To UBound(mDadosPedido, 1)
        Dim pedidoRow As Variant
        pedidoRow = ExtrairLinha(mDadosPedido, iPed)

        Dim chavePed As String
        chavePed = Trim$(CStr(mDadosPedido(iPed, mColPedVBELN))) & "|" & Trim$(CStr(mDadosPedido(iPed, mColPedPOSNR)))

        Dim algumaEntrega As Boolean
        algumaEntrega = False

        If mIdxVBFAporPrecedente.Exists(chavePed) Then
            Dim linhasVBFA1 As Collection
            Set linhasVBFA1 = mIdxVBFAporPrecedente(chavePed)

            Dim jV1 As Variant
            For Each jV1 In linhasVBFA1
                Dim iVBFA1 As Long
                iVBFA1 = CLng(jV1)

                If mColVBFA_VBTYPN = 0 Or Trim$(CStr(mDadosVBFA(iVBFA1, mColVBFA_VBTYPN))) = mVbtypEntrega Then
                    Dim vbelnEntrega As String
                    vbelnEntrega = Trim$(CStr(mDadosVBFA(iVBFA1, mColVBFA_VBELN)))

                    If mIdxLIPSporVBELN.Exists(vbelnEntrega) Then
                        Dim linhasLIPS As Collection
                        Set linhasLIPS = mIdxLIPSporVBELN(vbelnEntrega)

                        Dim jL As Variant
                        For Each jL In linhasLIPS
                            algumaEntrega = True
                            ProcessarEntrega resolvidas, pedidoRow, iVBFA1, CLng(jL)
                        Next jL
                    End If
                End If
            Next jV1
        End If

        If Not algumaEntrega Then
            EmitirLinha resolvidas, pedidoRow, Empty, Empty, Empty, Empty, Empty, Empty, Empty
        End If
    Next iPed

    EscreverSaidaCompleta resolvidas
    EscreverSaidaEnxuta resolvidas

    modLog.LogInfo "Abas finais reconstruídas a partir do acumulado: " & resolvidas.Count & " linha(s)."
End Sub

' ===================== Carga de estado (specs, dados, índices) =====================

Private Sub CarregarEstado(ByVal specs As Object, ByVal settings As Object)
    Set mSpecPedido = specs(SQVI_PEDIDO)
    Set mSpecLIPS = specs(SQVI_LIPS)
    Set mSpecVTTP = specs(SQVI_VTTP)
    Set mSpecVTTK = specs(SQVI_VTTK)
    Set mSpecVBFA = specs(SQVI_VBFA)
    Set mSpecVBRP = specs(SQVI_VBRP)
    Set mSpecVBRK = specs(SQVI_VBRK)

    mVbtypEntrega = settings(PARM_VBTYP_ENTREGA)
    mVbtypFatura = settings(PARM_VBTYP_FATURA)

    mDadosPedido = modExcelWrite.LerAbaComoArray(ABA_PEDIDO)
    mDadosLIPS = modExcelWrite.LerAbaComoArray(ABA_LIPS)
    mDadosVTTP = modExcelWrite.LerAbaComoArray(ABA_VTTP)
    mDadosVTTK = modExcelWrite.LerAbaComoArray(ABA_VTTK)
    mDadosVBFA = modExcelWrite.LerAbaComoArray(ABA_VBFA)
    mDadosVBRP = modExcelWrite.LerAbaComoArray(ABA_VBRP)
    mDadosVBRK = modExcelWrite.LerAbaComoArray(ABA_VBRK)

    mColPedVBELN = modUtils.ColunaDoCampo(mSpecPedido, "VBELN")
    mColPedPOSNR = modUtils.ColunaDoCampo(mSpecPedido, "POSNR")

    mColVBFA_VBELV = modUtils.ColunaDoCampo(mSpecVBFA, "VBELV")
    mColVBFA_POSNV = modUtils.ColunaDoCampo(mSpecVBFA, "POSNV")
    mColVBFA_VBELN = modUtils.ColunaDoCampo(mSpecVBFA, "VBELN")
    mColVBFA_VBTYPN = modUtils.ColunaDoCampo(mSpecVBFA, "VBTYP_N")

    mColLIPS_VBELN = modUtils.ColunaDoCampo(mSpecLIPS, "VBELN")
    mColLIPS_POSNR = modUtils.ColunaDoCampo(mSpecLIPS, "POSNR")

    mColVTTP_VBELN = modUtils.ColunaDoCampo(mSpecVTTP, "VBELN")
    mColVTTP_TKNUM = modUtils.ColunaDoCampo(mSpecVTTP, "TKNUM")

    mColVTTK_TKNUM = modUtils.ColunaDoCampo(mSpecVTTK, "TKNUM")

    mColVBRP_VBELN = modUtils.ColunaDoCampo(mSpecVBRP, "VBELN")
    mColVBRK_VBELN = modUtils.ColunaDoCampo(mSpecVBRK, "VBELN")

    If mColVBFA_VBELV = 0 Or mColVBFA_POSNV = 0 Then
        Set mIdxVBFAporPrecedente = CreateObject("Scripting.Dictionary")
        modLog.LogWarn "VBFA sem VBELV/POSNV no FELD (Config_SQVI) - abas finais não conseguirão ligar Pedido/Entrega."
    Else
        Set mIdxVBFAporPrecedente = ConstruirIndicePorPar(mDadosVBFA, mColVBFA_VBELV, mColVBFA_POSNV)
    End If

    Set mIdxLIPSporVBELN = ConstruirIndice1aN(mDadosLIPS, mColLIPS_VBELN)
    Set mIdxVTTPporVBELN = ConstruirIndice1aN(mDadosVTTP, mColVTTP_VBELN)
    Set mIdxVTTKporTKNUM = ConstruirIndice1a1(mDadosVTTK, mColVTTK_TKNUM)
    Set mIdxVBRPporVBELN = ConstruirIndice1aN(mDadosVBRP, mColVBRP_VBELN)
    Set mIdxVBRKporVBELN = ConstruirIndice1a1(mDadosVBRK, mColVBRK_VBELN)
End Sub

' ===================== Resolução do grafo por pedido =====================

Private Sub ProcessarEntrega(ByVal resultado As Collection, ByVal pedidoRow As Variant, _
        ByVal vbfa1Idx As Long, ByVal lipsIdx As Long)

    Dim lipsRow As Variant
    lipsRow = ExtrairLinha(mDadosLIPS, lipsIdx)
    Dim vbfa1Row As Variant
    vbfa1Row = ExtrairLinha(mDadosVBFA, vbfa1Idx)

    Dim vbelnEntrega As String
    vbelnEntrega = Trim$(CStr(mDadosLIPS(lipsIdx, mColLIPS_VBELN)))

    Dim linhasVTTP As Collection
    If mIdxVTTPporVBELN.Exists(vbelnEntrega) Then
        Set linhasVTTP = mIdxVTTPporVBELN(vbelnEntrega)
    Else
        Set linhasVTTP = New Collection
    End If

    Dim chaveLIPS As String
    chaveLIPS = vbelnEntrega & "|" & Trim$(CStr(mDadosLIPS(lipsIdx, mColLIPS_POSNR)))

    Dim combosFatura As Collection
    Set combosFatura = ResolverFaturamento(chaveLIPS)

    If linhasVTTP.Count = 0 And combosFatura.Count = 0 Then
        EmitirLinha resultado, pedidoRow, vbfa1Row, lipsRow, Empty, Empty, Empty, Empty, Empty
        Exit Sub
    End If

    If linhasVTTP.Count = 0 Then
        Dim comboF As Variant
        For Each comboF In combosFatura
            EmitirLinha resultado, pedidoRow, vbfa1Row, lipsRow, Empty, Empty, comboF(0), comboF(1), comboF(2)
        Next comboF
        Exit Sub
    End If

    Dim jT As Variant
    For Each jT In linhasVTTP
        Dim iVTTP As Long
        iVTTP = CLng(jT)
        Dim vttpRow As Variant
        vttpRow = ExtrairLinha(mDadosVTTP, iVTTP)

        Dim tknum As String
        tknum = Trim$(CStr(mDadosVTTP(iVTTP, mColVTTP_TKNUM)))
        Dim vttkRow As Variant
        vttkRow = Empty
        If Len(tknum) > 0 Then
            If mIdxVTTKporTKNUM.Exists(tknum) Then
                vttkRow = ExtrairLinha(mDadosVTTK, CLng(mIdxVTTKporTKNUM(tknum)))
            End If
        End If

        If combosFatura.Count > 0 Then
            Dim comboF2 As Variant
            For Each comboF2 In combosFatura
                EmitirLinha resultado, pedidoRow, vbfa1Row, lipsRow, vttpRow, vttkRow, comboF2(0), comboF2(1), comboF2(2)
            Next comboF2
        Else
            EmitirLinha resultado, pedidoRow, vbfa1Row, lipsRow, vttpRow, vttkRow, Empty, Empty, Empty
        End If
    Next jT
End Sub

' Retorna Collection de arrays [vbfa2Row, vbrpRow, vbrkRow] para uma chave "VBELN|POSNR" de LIPS.
Private Function ResolverFaturamento(ByVal chaveLIPS As String) As Collection
    Dim resultado As New Collection

    If Not mIdxVBFAporPrecedente.Exists(chaveLIPS) Then
        Set ResolverFaturamento = resultado
        Exit Function
    End If

    Dim linhasVBFA2 As Collection
    Set linhasVBFA2 = mIdxVBFAporPrecedente(chaveLIPS)

    Dim jV2 As Variant
    For Each jV2 In linhasVBFA2
        Dim iVBFA2 As Long
        iVBFA2 = CLng(jV2)

        If mColVBFA_VBTYPN = 0 Or Trim$(CStr(mDadosVBFA(iVBFA2, mColVBFA_VBTYPN))) = mVbtypFatura Then
            Dim vbfa2Row As Variant
            vbfa2Row = ExtrairLinha(mDadosVBFA, iVBFA2)

            Dim vbelnFatura As String
            vbelnFatura = Trim$(CStr(mDadosVBFA(iVBFA2, mColVBFA_VBELN)))

            If mIdxVBRPporVBELN.Exists(vbelnFatura) Then
                Dim linhasVBRP As Collection
                Set linhasVBRP = mIdxVBRPporVBELN(vbelnFatura)

                Dim jP As Variant
                For Each jP In linhasVBRP
                    Dim iVBRP As Long
                    iVBRP = CLng(jP)
                    Dim vbrpRow As Variant
                    vbrpRow = ExtrairLinha(mDadosVBRP, iVBRP)

                    Dim vbrkRow As Variant
                    vbrkRow = Empty
                    If mIdxVBRKporVBELN.Exists(vbelnFatura) Then
                        vbrkRow = ExtrairLinha(mDadosVBRK, CLng(mIdxVBRKporVBELN(vbelnFatura)))
                    End If

                    Dim combo(0 To 2) As Variant
                    combo(0) = vbfa2Row
                    combo(1) = vbrpRow
                    combo(2) = vbrkRow
                    resultado.Add combo
                Next jP
            End If
        End If
    Next jV2

    Set ResolverFaturamento = resultado
End Function

Private Sub EmitirLinha(ByVal resultado As Collection, ByVal pedidoRow As Variant, ByVal vbfa1Row As Variant, _
        ByVal lipsRow As Variant, ByVal vttpRow As Variant, ByVal vttkRow As Variant, _
        ByVal vbfa2Row As Variant, ByVal vbrpRow As Variant, ByVal vbrkRow As Variant)
    Dim lr As New clsLinhaResolvida
    lr.Pedido = pedidoRow
    lr.VBFA1 = vbfa1Row
    lr.LIPS = lipsRow
    lr.VTTP = vttpRow
    lr.VTTK = vttkRow
    lr.VBFA2 = vbfa2Row
    lr.VBRP = vbrpRow
    lr.VBRK = vbrkRow
    resultado.Add lr
End Sub

' ===================== Índices =====================

Private Function ConstruirIndice1a1(ByVal dados As Variant, ByVal coluna As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If coluna > 0 And modUtils.EhArrayValido(dados) Then
        Dim i As Long, chave As String
        For i = LBound(dados, 1) + 1 To UBound(dados, 1)
            chave = Trim$(CStr(dados(i, coluna)))
            If Len(chave) > 0 Then d(chave) = i
        Next i
    End If
    Set ConstruirIndice1a1 = d
End Function

Private Function ConstruirIndice1aN(ByVal dados As Variant, ByVal coluna As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If coluna > 0 And modUtils.EhArrayValido(dados) Then
        Dim i As Long, chave As String
        For i = LBound(dados, 1) + 1 To UBound(dados, 1)
            chave = Trim$(CStr(dados(i, coluna)))
            If Len(chave) > 0 Then
                If Not d.Exists(chave) Then Set d(chave) = New Collection
                d(chave).Add i
            End If
        Next i
    End If
    Set ConstruirIndice1aN = d
End Function

Private Function ConstruirIndicePorPar(ByVal dados As Variant, ByVal col1 As Long, ByVal col2 As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If col1 > 0 And col2 > 0 And modUtils.EhArrayValido(dados) Then
        Dim i As Long, chave As String
        For i = LBound(dados, 1) + 1 To UBound(dados, 1)
            chave = Trim$(CStr(dados(i, col1))) & "|" & Trim$(CStr(dados(i, col2)))
            If Not d.Exists(chave) Then Set d(chave) = New Collection
            d(chave).Add i
        Next i
    End If
    Set ConstruirIndicePorPar = d
End Function

Private Function ExtrairLinha(ByVal dados As Variant, ByVal linha As Long) As Variant
    Dim numCols As Long
    numCols = UBound(dados, 2) - LBound(dados, 2) + 1
    Dim resultado() As Variant
    ReDim resultado(1 To numCols)
    Dim c As Long, idx As Long
    idx = 1
    For c = LBound(dados, 2) To UBound(dados, 2)
        resultado(idx) = dados(linha, c)
        idx = idx + 1
    Next c
    ExtrairLinha = resultado
End Function

' ===================== Escrita das abas finais =====================

Private Sub EscreverSaidaCompleta(ByVal resolvidas As Collection)
    Dim ws As Worksheet
    Set ws = modUtils.ObterOuCriarAba(ABA_SAIDA_COMPLETA)
    ws.Cells.Clear

    Dim cabecalho As New Collection
    AdicionarCabecalhosComPrefixo cabecalho, "PEDIDO", mSpecPedido.FELD
    AdicionarCabecalhosComPrefixo cabecalho, "LIPS", mSpecLIPS.FELD
    AdicionarCabecalhosComPrefixo cabecalho, "VTTP", mSpecVTTP.FELD
    AdicionarCabecalhosComPrefixo cabecalho, "VTTK", mSpecVTTK.FELD
    AdicionarCabecalhosComPrefixo cabecalho, "VBFA_PEDIDO_ENTREGA", mSpecVBFA.FELD
    AdicionarCabecalhosComPrefixo cabecalho, "VBFA_ENTREGA_FATURA", mSpecVBFA.FELD
    AdicionarCabecalhosComPrefixo cabecalho, "VBRP", mSpecVBRP.FELD
    AdicionarCabecalhosComPrefixo cabecalho, "VBRK", mSpecVBRK.FELD

    EscreverCabecalho ws, cabecalho

    If resolvidas.Count = 0 Then Exit Sub

    Dim numCols As Long
    numCols = cabecalho.Count

    Dim saida() As Variant
    ReDim saida(1 To resolvidas.Count, 1 To numCols)

    Dim linha As Long
    linha = 1
    Dim lr As clsLinhaResolvida
    For Each lr In resolvidas
        Dim col As Long
        col = 1
        col = PreencherBloco(saida, linha, col, lr.Pedido, mSpecPedido.FELD.Count)
        col = PreencherBloco(saida, linha, col, lr.LIPS, mSpecLIPS.FELD.Count)
        col = PreencherBloco(saida, linha, col, lr.VTTP, mSpecVTTP.FELD.Count)
        col = PreencherBloco(saida, linha, col, lr.VTTK, mSpecVTTK.FELD.Count)
        col = PreencherBloco(saida, linha, col, lr.VBFA1, mSpecVBFA.FELD.Count)
        col = PreencherBloco(saida, linha, col, lr.VBFA2, mSpecVBFA.FELD.Count)
        col = PreencherBloco(saida, linha, col, lr.VBRP, mSpecVBRP.FELD.Count)
        col = PreencherBloco(saida, linha, col, lr.VBRK, mSpecVBRK.FELD.Count)
        linha = linha + 1
    Next lr

    ws.Range(ws.Cells(2, 1), ws.Cells(1 + resolvidas.Count, numCols)).Value2 = saida
End Sub

' Campos curados (aba "Saída Enxuta" do Livro de Regras SAP.xlsx + chaves do Pedido).
Private Sub EscreverSaidaEnxuta(ByVal resolvidas As Collection)
    Dim ws As Worksheet
    Set ws = modUtils.ObterOuCriarAba(ABA_SAIDA_ENXUTA)
    ws.Cells.Clear

    Dim camposPedido As Collection: Set camposPedido = modUtils.SplitTrim("VBELN;POSNR;WERKS;MATNR;ERDAT;KWMENG;NETWR")
    Dim camposLIPS As Collection: Set camposLIPS = modUtils.SplitTrim("POSNR;MATNR;ARKTX;WERKS;LGORT;CHARG;LFIMG;VRKME;VGBEL;VGPOS")
    Dim camposVTTP As Collection: Set camposVTTP = modUtils.SplitTrim("VBELN;TKNUM")
    Dim camposVTTK As Collection: Set camposVTTK = modUtils.SplitTrim("TDLNR;TKNUM;SHTYP;TPLST;ERDAT;ERZET;ROUTE;EXTI2;STTRG;DPTBG;UPTBG;DATBG;UATBG;DPTEN;UPTEN;DATEN;UATEN")
    Dim camposVBFA As Collection: Set camposVBFA = modUtils.SplitTrim("VBELN;POSNN;VBTYP_N")
    Dim camposVBRP As Collection: Set camposVBRP = modUtils.SplitTrim("VBELN;POSNR;FKIMG")
    Dim camposVBRK As Collection: Set camposVBRK = modUtils.SplitTrim("FKART;FKDAT;FKSTO")

    Dim cabecalho As New Collection
    AdicionarCabecalhosComPrefixo cabecalho, "PEDIDO", camposPedido
    AdicionarCabecalhosComPrefixo cabecalho, "LIPS", camposLIPS
    AdicionarCabecalhosComPrefixo cabecalho, "VTTP", camposVTTP
    AdicionarCabecalhosComPrefixo cabecalho, "VTTK", camposVTTK
    AdicionarCabecalhosComPrefixo cabecalho, "VBFA_ENTREGA_FATURA", camposVBFA
    AdicionarCabecalhosComPrefixo cabecalho, "VBRP", camposVBRP
    AdicionarCabecalhosComPrefixo cabecalho, "VBRK", camposVBRK

    EscreverCabecalho ws, cabecalho

    If resolvidas.Count = 0 Then Exit Sub

    Dim numCols As Long
    numCols = cabecalho.Count

    Dim saida() As Variant
    ReDim saida(1 To resolvidas.Count, 1 To numCols)

    Dim linha As Long
    linha = 1
    Dim lr As clsLinhaResolvida
    For Each lr In resolvidas
        Dim col As Long
        col = 1
        col = PreencherBlocoSeletivo(saida, linha, col, lr.Pedido, mSpecPedido, camposPedido)
        col = PreencherBlocoSeletivo(saida, linha, col, lr.LIPS, mSpecLIPS, camposLIPS)
        col = PreencherBlocoSeletivo(saida, linha, col, lr.VTTP, mSpecVTTP, camposVTTP)
        col = PreencherBlocoSeletivo(saida, linha, col, lr.VTTK, mSpecVTTK, camposVTTK)
        col = PreencherBlocoSeletivo(saida, linha, col, lr.VBFA2, mSpecVBFA, camposVBFA)
        col = PreencherBlocoSeletivo(saida, linha, col, lr.VBRP, mSpecVBRP, camposVBRP)
        col = PreencherBlocoSeletivo(saida, linha, col, lr.VBRK, mSpecVBRK, camposVBRK)
        linha = linha + 1
    Next lr

    ws.Range(ws.Cells(2, 1), ws.Cells(1 + resolvidas.Count, numCols)).Value2 = saida
End Sub

Private Function PreencherBloco(ByRef saida() As Variant, ByVal linha As Long, ByVal colInicial As Long, _
        ByVal dadosLinha As Variant, ByVal numColunasBloco As Long) As Long
    Dim c As Long
    If IsEmpty(dadosLinha) Then
        For c = 1 To numColunasBloco
            saida(linha, colInicial + c - 1) = ""
        Next c
    Else
        For c = 1 To numColunasBloco
            saida(linha, colInicial + c - 1) = dadosLinha(c)
        Next c
    End If
    PreencherBloco = colInicial + numColunasBloco
End Function

Private Function PreencherBlocoSeletivo(ByRef saida() As Variant, ByVal linha As Long, ByVal colInicial As Long, _
        ByVal dadosLinha As Variant, ByVal spec As clsSQVISpec, ByVal camposDesejados As Collection) As Long
    Dim c As Long, posOriginal As Long
    c = 0
    Dim campo As Variant
    For Each campo In camposDesejados
        c = c + 1
        If IsEmpty(dadosLinha) Then
            saida(linha, colInicial + c - 1) = ""
        Else
            posOriginal = modUtils.ColunaDoCampo(spec, CStr(campo))
            If posOriginal = 0 Then
                saida(linha, colInicial + c - 1) = ""
            Else
                saida(linha, colInicial + c - 1) = dadosLinha(posOriginal)
            End If
        End If
    Next campo
    PreencherBlocoSeletivo = colInicial + camposDesejados.Count
End Function

Private Sub AdicionarCabecalhosComPrefixo(ByVal cabecalho As Collection, ByVal prefixo As String, ByVal campos As Collection)
    Dim campo As Variant
    For Each campo In campos
        cabecalho.Add prefixo & "." & CStr(campo)
    Next campo
End Sub

Private Sub EscreverCabecalho(ByVal ws As Worksheet, ByVal cabecalho As Collection)
    Dim c As Long
    c = 1
    Dim item As Variant
    For Each item In cabecalho
        ws.Cells(1, c).Value = CStr(item)
        c = c + 1
    Next item
    ws.Rows(1).Font.Bold = True
End Sub
