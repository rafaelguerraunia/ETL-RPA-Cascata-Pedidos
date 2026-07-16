VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRunResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Resultado de uma execução de SQVI (modSQVIEngine.RunSQVI).

Public Success As Boolean
Public MensagemErro As String
Public Cabecalho As Variant            ' array 1D com os nomes das colunas (1ª linha do export)
Public Dados As Variant                ' array 2D (linhas x colunas), SEM o cabeçalho
Public NumLinhas As Long
Public CamposFELDFaltando As Collection    ' campos obrigatórios que não vieram no cabeçalho
Public CamposSELEFaltando As Collection    ' campos de filtro que não existiam na tela de seleção
Public ModoDegradado As Boolean            ' True se algum filtro pedido não pôde ser aplicado

Private Sub Class_Initialize()
    Set CamposFELDFaltando = New Collection
    Set CamposSELEFaltando = New Collection
    Success = False
    ModoDegradado = False
    NumLinhas = 0
End Sub
