VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQVISpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Especificação declarativa de uma SQVI (QuickView) usada na cascata.
' Uma instância por linha da aba Config_SQVI.

Public SQVIName As String          ' ex: "-RGVS-VBAP"
Public Titulo As String            ' título da QuickView, ex: "E2E VBAP (Pedido)"
Public Tabela As String            ' tabela SAP de origem, ex: "VBAP"
Public FELD As Collection          ' campos de saída, na ordem em que devem ser cadastrados
Public SELE As Collection          ' campos de seleção, na ordem em que estão cadastrados na SQVI
                                    ' (a posição 1-based define o sufixo SP$0000N em runtime)
Public CamposObrigatoriosFELD As Collection   ' subconjunto de FELD sem o qual o hop não pode ser usado
Public AbaDestino As String        ' nome da aba no workbook onde os dados são gravados (upsert)
Public ChaveUpsert As Collection   ' nomes de coluna que formam a chave de upsert (ex: VBELN, POSNR)
Public PodeCriarAutomaticamente As Boolean  ' False p/ SQVIs que já existem e não devem ser criadas do zero

Private Sub Class_Initialize()
    Set FELD = New Collection
    Set SELE = New Collection
    Set CamposObrigatoriosFELD = New Collection
    Set ChaveUpsert = New Collection
    PodeCriarAutomaticamente = True
End Sub

' Retorna a posição 1-based do campo dentro de SELE, ou 0 se não encontrado.
Public Function PosicaoSELE(ByVal nomeCampo As String) As Long
    Dim i As Long
    For i = 1 To SELE.Count
        If UCase$(SELE(i)) = UCase$(nomeCampo) Then
            PosicaoSELE = i
            Exit Function
        End If
    Next i
    PosicaoSELE = 0
End Function
