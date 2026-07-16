Attribute VB_Name = "modDebug"
Option Explicit

' Ferramenta de diagnóstico (não faz parte do fluxo normal do ETL): testa a
' criação de UMA SQVI passo a passo, checando a existência de cada elemento
' de tela do SAP GUI antes de usá-lo, e registrando tudo na aba Log (Hop =
' "DEBUG-<nome da SQVI>") sem parar no primeiro erro - para identificar
' exatamente em qual tela/campo os IDs esperados por modSQVIEngine.CreateSQVI
' não batem com a versão/layout do SAP GUI do usuário.
'
' Uso: Alt+F8 -> DebugCriarSQVIPassoAPasso -> Executar. Pede o nome da SQVI
' (ex: -RGVS-VBAP). Depois, veja a aba Log filtrando a coluna Hop por
' "DEBUG-<nome>".
'
' Repete as MESMAS ações de modSQVIEngine.CreateSQVI (então, se tudo existir,
' a SQVI é criada de verdade) - a diferença é que aqui cada passo é checado
' e logado individualmente, e os passos finais de salvar/gerar a query
' (tbar[0]/btn[11] e btn[3]) são deliberadamente OMITIDOS, para não persistir
' no SAP uma SQVI incompleta caso algum campo tenha falhado no meio do caminho.

Public Sub DebugCriarSQVIPassoAPasso()
    Dim nomeTeste As String
    nomeTeste = InputBox("Nome da SQVI para testar (ex: -RGVS-VBAP):", "Debug criação de SQVI", "-RGVS-VBAP")
    If Len(Trim$(nomeTeste)) = 0 Then Exit Sub
    nomeTeste = Trim$(nomeTeste)

    modLog.InicializarLog modConfig.ObterCaminhoLog(modConfig.LoadGlobalSettings())

    If Not modSAPConnection.ConectarSAP() Then
        modLog.LogErro "Debug: não foi possível conectar a uma sessão SAP GUI já aberta.", "DEBUG"
        MsgBox "SAP não conectado. Abra e faça login antes de rodar o debug.", vbExclamation
        Exit Sub
    End If

    Dim specs As Object
    Set specs = modConfig.LoadSQVISpecs()
    If Not specs.Exists(nomeTeste) Then
        modLog.LogErro "Debug: '" & nomeTeste & "' não encontrada em Config_SQVI.", "DEBUG"
        MsgBox "SQVI '" & nomeTeste & "' não está cadastrada em Config_SQVI.", vbExclamation
        Exit Sub
    End If

    Dim spec As clsSQVISpec
    Set spec = specs(nomeTeste)
    Dim hop As String
    hop = "DEBUG-" & spec.SQVIName

    modLog.LogInfo "===== Debug de criação: " & spec.SQVIName & " (tabela " & spec.Tabela & ") =====", hop

    modSAPConnection.NavegarTransacao "/nsqvi"
    modSAPConnection.Definir "wnd[0]/usr/ctxtRS38R-QNUM", spec.SQVIName
    modSAPConnection.Pressionar "wnd[0]/usr/btnP1"
    modSAPConnection.AguardarSAP
    modLog.LogInfo "Status bar após consultar: " & modSAPConnection.StatusBarTexto(), hop

    If Not modSAPConnection.StatusBarIndicaNaoExiste() Then
        modLog.LogInfo spec.SQVIName & " já existe no SAP - nada a criar. Teste encerrado.", hop
        MsgBox spec.SQVIName & " já existe no SAP.", vbInformation
        Exit Sub
    End If

    If Not spec.PodeCriarAutomaticamente Then
        modLog.LogInfo spec.SQVIName & " não existe, e PodeCriarAutomaticamente = FALSO. Teste encerrado.", hop
        MsgBox spec.SQVIName & " não existe e não está marcada para criação automática.", vbInformation
        Exit Sub
    End If

    ' -------- passo a passo, checando cada elemento antes de usar --------
    If Not TestarElemento("wnd[0]/usr/btnP7", hop, "botão 'Criar'") Then GoTo Fim
    modSAPConnection.Pressionar "wnd[0]/usr/btnP7"
    modSAPConnection.AguardarSAP

    If Not TestarElemento("wnd[1]/usr/txtRS38R-HDTITLE", hop, "título da query (janela de criação)") Then GoTo Fim
    modSAPConnection.Definir "wnd[1]/usr/txtRS38R-HDTITLE", spec.Titulo

    If Not TestarElemento("wnd[1]/usr/subSUBSOURCE:SAPMS38R:3110/ctxtRS38Q-DDNAME", hop, "campo tabela de origem") Then GoTo Fim
    modSAPConnection.Definir "wnd[1]/usr/subSUBSOURCE:SAPMS38R:3110/ctxtRS38Q-DDNAME", spec.Tabela

    If Not TestarElemento("wnd[1]/tbar[0]/btn[0]", hop, "confirmar (Enter) janela de criação") Then GoTo Fim
    modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[0]"
    modSAPConnection.AguardarSAP
    modLog.LogInfo "Status bar após confirmar criação: " & modSAPConnection.StatusBarTexto(), hop

    If Not TestarElemento("wnd[0]/tbar[1]/btn[17]", hop, "botão 'gerar lista de campos' (abre aba FELD)") Then GoTo Fim
    modSAPConnection.Pressionar "wnd[0]/tbar[1]/btn[17]"

    Dim campo As Variant
    For Each campo In spec.FELD
        TestarAddCampo "FELD", CStr(campo), hop
    Next campo

    If Not TestarElemento("wnd[0]/usr/tabsTAB100/tabpSELE", hop, "aba SELE") Then GoTo Fim
    modSAPConnection.Selecionar "wnd[0]/usr/tabsTAB100/tabpSELE"
    modSAPConnection.Pressionar "wnd[0]/tbar[1]/btn[17]"

    For Each campo In spec.SELE
        TestarAddCampo "SELE", CStr(campo), hop
    Next campo

Fim:
    modLog.LogInfo "Passo a passo concluído (debug NÃO salva/gera a query no SAP, de propósito). " & _
        "Revise a aba Log filtrando a coluna Hop por '" & hop & "' para ver o que faltou.", hop
    MsgBox "Debug concluído. Veja a aba Log (Hop = " & hop & ") para o resultado passo a passo.", vbInformation
End Sub

' Retorna True/False e registra o resultado; não interrompe a execução.
Private Function TestarElemento(ByVal idElemento As String, ByVal hop As String, ByVal descricao As String) As Boolean
    If modSAPConnection.ElementExists(idElemento) Then
        modLog.LogInfo "OK - " & descricao & " (" & idElemento & ")", hop
        TestarElemento = True
    Else
        modLog.LogErro "NÃO ENCONTRADO - " & descricao & " (" & idElemento & ") - a tela do seu SAP não tem " & _
            "esse elemento nesse ID; é provável que a versão/layout do SAP GUI seja diferente do esperado " & _
            "pelo código nesse ponto.", hop
        TestarElemento = False
    End If
End Function

Private Sub TestarAddCampo(ByVal aba As String, ByVal campo As String, ByVal hop As String)
    Dim subtela As String
    If aba = "FELD" Then
        subtela = "3120"
    Else
        subtela = "3140"
    End If
    Dim base As String
    base = "wnd[0]/usr/tabsTAB100/tabp" & aba & "/ssubSUBSOURCE:SAPMS38R:" & subtela

    If Not TestarElemento(base & "/btnRSEA", hop, "campo '" & campo & "' (" & aba & ") - botão de busca") Then Exit Sub
    modSAPConnection.Pressionar base & "/btnRSEA"

    If Not TestarElemento("wnd[1]/usr/txtTC_SORTPOOLENTRY-NAME", hop, "campo '" & campo & "' (" & aba & ") - caixa de busca") Then Exit Sub
    modSAPConnection.Definir "wnd[1]/usr/txtTC_SORTPOOLENTRY-NAME", campo

    If Not TestarElemento("wnd[1]/tbar[0]/btn[71]", hop, "campo '" & campo & "' (" & aba & ") - confirmar busca") Then Exit Sub
    modSAPConnection.Pressionar "wnd[1]/tbar[0]/btn[71]"

    If TestarElemento(base & "/btnMVLE", hop, "campo '" & campo & "' (" & aba & ") - mover para a direita") Then
        modSAPConnection.Pressionar base & "/btnMVLE"
    End If
End Sub
