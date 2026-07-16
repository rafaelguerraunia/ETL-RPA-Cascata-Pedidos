Attribute VB_Name = "modUtils"
Option Explicit

' Utilitários gerais: sleep, clipboard, datas, arquivos, listas/dicionários.
' Reaproveita os padrões comprovados em RPA BASE.vbs, adaptados para VBA.

#If VBA7 Then
    Private Declare PtrSafe Sub SleepAPI Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub SleepAPI Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Public Sub Esperar(ByVal ms As Long)
    SleepAPI ms
End Sub

' ===================== Abas =====================

Public Function ObterOuCriarAba(ByVal nome As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nome)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nome
    End If
    Set ObterOuCriarAba = ws
End Function

' ===================== Clipboard (mesmo padrão de CopiarClipboard do RPA BASE) =====================

Public Sub CopiarClipboard(ByVal texto As String)
    Dim fso As Object, sh As Object, ts As Object
    Dim caminhoTemp As String

    If Len(texto) = 0 Then
        modLog.LogWarn "CopiarClipboard chamado com lista vazia; clipboard não alterado."
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sh = CreateObject("WScript.Shell")

    caminhoTemp = fso.GetSpecialFolder(2) & "\etl_clip_" & fso.GetTempName()
    Set ts = fso.CreateTextFile(caminhoTemp, True, False)
    ts.Write texto
    ts.Close

    sh.Run "cmd /c clip < """ & caminhoTemp & """", 0, True

    On Error Resume Next
    fso.DeleteFile caminhoTemp
    On Error GoTo 0
End Sub

' ===================== Datas =====================

' Formato MM/DD/YYYY - mesmo formato usado pelo RPA BASE.vbs para preencher campos de data
' do SAP (depende da configuração regional do usuário SAP; validar se o sistema usar outro formato).
Public Function DataSAP(ByVal d As Date) As String
    DataSAP = Right$("0" & Month(d), 2) & "/" & Right$("0" & Day(d), 2) & "/" & Year(d)
End Function

' ===================== Arquivos (mesmos padrões do RPA BASE) =====================

Public Sub DeletarSeExistir(ByVal caminho As String)
    On Error Resume Next
    If Len(Dir$(caminho)) > 0 Then Kill caminho
    On Error GoTo 0
End Sub

Public Sub MoverESubstituir(ByVal origem As String, ByVal destino As String)
    On Error Resume Next
    If Len(Dir$(origem)) > 0 Then
        If Len(Dir$(destino)) > 0 Then Kill destino
        Name origem As destino
    End If
    On Error GoTo 0
End Sub

Public Function AguardarArquivo(ByVal caminho As String, ByVal timeoutSeg As Long) As Boolean
    Dim limite As Date, limiteExtra As Date
    Dim tam1 As Currency, tam2 As Currency   ' Currency para arquivos grandes sem estourar Long

    limite = DateAdd("s", timeoutSeg, Now)
    Do While Len(Dir$(caminho)) = 0 And Now < limite
        Esperar 300
    Loop
    If Len(Dir$(caminho)) = 0 Then
        AguardarArquivo = False
        Exit Function
    End If

    tam1 = -1
    limiteExtra = DateAdd("s", 15, limite)
    Do
        On Error Resume Next
        tam2 = FileLen(caminho)
        On Error GoTo 0
        If tam2 = tam1 And tam2 > 0 Then Exit Do
        tam1 = tam2
        Esperar 600
    Loop While Now < limiteExtra
    AguardarArquivo = True
End Function

' Fecha qualquer instância aberta (nossa ou externa) de um workbook com esse nome de arquivo.
' Mesmo propósito de FecharPlanilhaSAP no RPA BASE - necessário pois o SAP GUI pode abrir
' o arquivo exportado automaticamente em uma instância de Excel (a nossa ou uma nova).
Public Sub FecharPlanilhaSeAberta(ByVal nomeArquivo As String)
    Dim wb As Workbook
    On Error Resume Next
    For Each wb In Application.Workbooks
        If LCase$(wb.Name) = LCase$(nomeArquivo) Then wb.Close False
    Next wb
    On Error GoTo 0

    Dim appExterno As Object
    Set appExterno = Nothing
    On Error Resume Next
    Set appExterno = GetObject(, "Excel.Application")
    On Error GoTo 0
    If Not appExterno Is Nothing Then
        Dim wb2 As Object
        On Error Resume Next
        For Each wb2 In appExterno.Workbooks
            If LCase$(wb2.Name) = LCase$(nomeArquivo) Then wb2.Close False
        Next wb2
        On Error GoTo 0
    End If
End Sub

' ===================== Listas / texto =====================

' Divide um texto delimitado (ex: célula "VBELN;POSNR;WERKS") em uma Collection,
' removendo espaços e itens vazios.
Public Function SplitTrim(ByVal texto As String, Optional ByVal delim As String = ";") As Collection
    Dim resultado As New Collection
    Dim partes() As String
    Dim i As Long, v As String

    If Len(Trim$(texto)) = 0 Then
        Set SplitTrim = resultado
        Exit Function
    End If

    partes = Split(texto, delim)
    For i = LBound(partes) To UBound(partes)
        v = Trim$(partes(i))
        If Len(v) > 0 Then resultado.Add v
    Next i
    Set SplitTrim = resultado
End Function

Public Function JoinColecao(ByVal c As Collection, Optional ByVal delim As String = "; ") As String
    Dim resultado As String
    Dim item As Variant
    For Each item In c
        If Len(resultado) > 0 Then resultado = resultado & delim
        resultado = resultado & CStr(item)
    Next item
    JoinColecao = resultado
End Function

Public Function ColecaoContem(ByVal c As Collection, ByVal valor As String) As Boolean
    Dim item As Variant
    For Each item In c
        If UCase$(CStr(item)) = UCase$(valor) Then
            ColecaoContem = True
            Exit Function
        End If
    Next item
    ColecaoContem = False
End Function

' ===================== Leitura de dados exportados do SAP (acesso posicional) =====================
'
' IMPORTANTE: o cabeçalho exportado pelo SAP traz o TEXTO DE DESCRIÇÃO do campo
' (ex: "Delivery", "Plant"), não o nome técnico (ex: "VBELN", "WERKS") - por isso
' NÃO fazemos matching por nome de coluna. Em vez disso, usamos a ORDEM em que os
' campos foram cadastrados no FELD da SQVI (clsSQVISpec.FELD), que corresponde à
' ordem das colunas no export - mesmo princípio já usado no RPA BASE.vbs
' (LerMatriz: "Material = coluna A (1)", "coluna G (7)").

' Posição 1-based do campo dentro do FELD da spec (= número da coluna no array exportado).
Public Function ColunaDoCampo(ByVal spec As clsSQVISpec, ByVal nomeCampo As String) As Long
    Dim i As Long
    For i = 1 To spec.FELD.Count
        If UCase$(spec.FELD(i)) = UCase$(nomeCampo) Then
            ColunaDoCampo = i
            Exit Function
        End If
    Next i
    ColunaDoCampo = 0
End Function

' Valores distintos e não-vazios de uma coluna (1-based) de um array 2D, a partir de linhaInicial.
Public Function ValoresDistintosPorColuna(ByVal dados As Variant, ByVal linhaInicial As Long, ByVal coluna As Long) As Collection
    Dim resultado As New Collection
    Dim vistos As Object
    Set vistos = CreateObject("Scripting.Dictionary")
    Dim i As Long, v As String

    If Not EhArrayValido(dados) Then
        Set ValoresDistintosPorColuna = resultado
        Exit Function
    End If

    For i = linhaInicial To UBound(dados, 1)
        v = Trim$(CStr(dados(i, coluna)))
        If Len(v) > 0 Then
            If Not vistos.Exists(v) Then
                vistos.Add v, True
                resultado.Add v
            End If
        End If
    Next i
    Set ValoresDistintosPorColuna = resultado
End Function

' Dictionary "valorCol1|valorCol2" -> True, para checagem exata de pares de chave.
Public Function ParesDistintosDict(ByVal dados As Variant, ByVal linhaInicial As Long, ByVal col1 As Long, ByVal col2 As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long, chave As String

    If Not EhArrayValido(dados) Then
        Set ParesDistintosDict = d
        Exit Function
    End If

    For i = linhaInicial To UBound(dados, 1)
        chave = Trim$(CStr(dados(i, col1))) & "|" & Trim$(CStr(dados(i, col2)))
        If Not d.Exists(chave) Then d.Add chave, True
    Next i
    Set ParesDistintosDict = d
End Function

Public Function EhArrayValido(ByVal v As Variant) As Boolean
    On Error GoTo naoEValido
    EhArrayValido = (UBound(v, 1) >= LBound(v, 1))
    Exit Function
naoEValido:
    EhArrayValido = False
End Function
