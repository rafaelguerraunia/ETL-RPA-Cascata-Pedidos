VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Representa um filtro a aplicar na tela de seleção de uma SQVI em runtime.
' O campo é identificado pelo NOME LÓGICO (nome técnico SAP, ex: "WERKS"),
' e modSQVIEngine resolve isso para a posição SP$0000N usando clsSQVISpec.PosicaoSELE.

Public CampoLogico As String       ' ex: "WERKS", "VBELN", "ERDAT"
Public Modo As String              ' um dos: INCLUIR_MULTIPLO, COLAR_MULTIPLO, EXCLUIR_FLAG, INTERVALO_DATA
Public Valores As Collection       ' lista de valores (INCLUIR_MULTIPLO, COLAR_MULTIPLO) ou 1 valor (EXCLUIR_FLAG)
Public ValorInicial As String      ' INTERVALO_DATA: data inicial (formato SAP, ver modUtils.DataSAP)
Public ValorFinal As String        ' INTERVALO_DATA: data final

Private Sub Class_Initialize()
    Set Valores = New Collection
End Sub
