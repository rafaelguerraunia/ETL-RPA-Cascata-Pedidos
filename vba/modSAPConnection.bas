Attribute VB_Name = "modSAPConnection"
Option Explicit

' Conexão com uma sessão SAP GUI JÁ ABERTA E LOGADA (não faz login).
' Mesmo padrão de ConectarSAP em RPA BASE.vbs.
' Nota: variável chamada "SapApp" (não "Application") para não colidir com o
' objeto Application do próprio Excel/VBA.

Public SapGuiAuto As Object
Public SapApp As Object
Public SapConnection As Object
Public Session As Object

Public Function ConectarSAP() As Boolean
    On Error Resume Next
    Set SapGuiAuto = Nothing
    Set SapGuiAuto = GetObject("SAPGUI")
    If SapGuiAuto Is Nothing Then
        ConectarSAP = False
        Exit Function
    End If

    Set SapApp = SapGuiAuto.GetScriptingEngine
    If SapApp Is Nothing Then
        ConectarSAP = False
        Exit Function
    End If

    Set SapConnection = SapApp.Children(0)
    If SapConnection Is Nothing Then
        ConectarSAP = False
        Exit Function
    End If

    If SapConnection.Children.Count > 0 Then
        Set Session = SapConnection.Children(SapConnection.Children.Count - 1)
    Else
        ConectarSAP = False
        Exit Function
    End If
    On Error GoTo 0

    ConectarSAP = Not (Session Is Nothing)
End Function

Public Sub AguardarSAP()
    Dim tentativas As Long
    On Error Resume Next
    Do While Session.Busy
        modUtils.Esperar 50
        tentativas = tentativas + 1
        If tentativas > 1200 Then Exit Do   ' trava de segurança: no máx ~60s de espera
    Loop
    On Error GoTo 0
End Sub

Public Function ElementExists(ByVal idElemento As String) As Boolean
    Dim el As Object
    On Error Resume Next
    Set el = Nothing
    Set el = Session.findById(idElemento)
    On Error GoTo 0
    ElementExists = Not (el Is Nothing)
End Function

Public Sub NavegarTransacao(ByVal codigoTransacao As String)
    Session.findById("wnd[0]/tbar[0]/okcd").Text = codigoTransacao
    Session.findById("wnd[0]").sendVKey 0
    AguardarSAP
End Sub

Public Function StatusBarTexto() As String
    On Error Resume Next
    StatusBarTexto = LCase$(Session.findById("wnd[0]/sbar").Text)
    On Error GoTo 0
End Function

' Mesma heurística usada em RPA BASE.vbs para decidir se uma SQVI precisa ser criada.
Public Function StatusBarIndicaNaoExiste() As Boolean
    Dim msg As String
    msg = StatusBarTexto()
    StatusBarIndicaNaoExiste = (InStr(msg, "não") > 0 Or InStr(msg, "not") > 0 _
        Or InStr(msg, "exist") > 0 Or InStr(msg, "criad") > 0)
End Function

' Wrappers curtos (deixam modSQVIEngine mais legível).
Public Sub Definir(ByVal idElemento As String, ByVal valor As String)
    Session.findById(idElemento).Text = valor
End Sub

Public Sub Pressionar(ByVal idElemento As String)
    Session.findById(idElemento).press
End Sub

Public Sub Selecionar(ByVal idElemento As String)
    Session.findById(idElemento).Select
End Sub
