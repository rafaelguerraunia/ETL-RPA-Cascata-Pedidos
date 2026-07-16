VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLinhaResolvida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Uma "linha" resolvida da cascata completa: um caminho Pedido -> Entrega ->
' {Embarque, Faturamento}. Cada campo é um array 1D (uma linha inteira daquela
' tabela) ou Empty se não houve match naquele nível (pedido incompleto).

Public Pedido As Variant
Public LIPS As Variant
Public VTTP As Variant
Public VTTK As Variant
Public VBFA1 As Variant   ' hop Pedido -> Entrega
Public VBFA2 As Variant   ' hop Entrega -> Fatura
Public VBRP As Variant
Public VBRK As Variant
