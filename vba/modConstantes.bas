Attribute VB_Name = "modConstantes"
Option Explicit

' ===================== Modos de filtro (clsFiltro.Modo) =====================
Public Const FILTRO_INCLUIR_MULTIPLO As String = "INCLUIR_MULTIPLO"   ' lista curta digitada direto (ex: plantas)
Public Const FILTRO_COLAR_MULTIPLO As String = "COLAR_MULTIPLO"       ' lista grande colada via clipboard (ex: chaves)
Public Const FILTRO_EXCLUIR_FLAG As String = "EXCLUIR_FLAG"           ' exclui 1 valor único (ex: LOEKZ <> "X")
Public Const FILTRO_INTERVALO_DATA As String = "INTERVALO_DATA"       ' campo de data direto, LOW/HIGH

' ===================== Níveis de log =====================
Public Const LOG_INFO As String = "INFO"
Public Const LOG_WARN As String = "WARN"
Public Const LOG_ERROR As String = "ERROR"
Public Const LOG_FATAL As String = "FATAL"

' ===================== Nomes de abas =====================
Public Const ABA_CONFIG_SQVI As String = "Config_SQVI"
Public Const ABA_CONFIG_GERAL As String = "Config_Geral"
Public Const ABA_LOG As String = "Log"
Public Const ABA_PEDIDO As String = "PEDIDO"
Public Const ABA_LIPS As String = "LIPS"
Public Const ABA_VTTP As String = "VTTP"
Public Const ABA_VTTK As String = "VTTK"
Public Const ABA_VBFA As String = "VBFA"
Public Const ABA_VBRP As String = "VBRP"
Public Const ABA_VBRK As String = "VBRK"
Public Const ABA_SAIDA_ENXUTA As String = "Saída Enxuta"
Public Const ABA_SAIDA_COMPLETA As String = "Saída Completa"
Public Const ABA_INICIO As String = "Início"

' ===================== Nomes de SQVI =====================
Public Const SQVI_PEDIDO As String = "-RGVS-VBAP"
Public Const SQVI_VBFA As String = "-RGVS-VBFA"
Public Const SQVI_LIPS As String = "-RGVS-LIPS"
Public Const SQVI_VTTP As String = "-RGVS-VTTP"
Public Const SQVI_VTTK As String = "-RGVS-VTTK"
Public Const SQVI_VBRP As String = "-RGVS-VBRP"
Public Const SQVI_VBRK As String = "-RGVS-VBRK"

' ===================== Parâmetros de Config_Geral (nomes de chave) =====================
Public Const PARM_PLANTAS As String = "Plantas"
Public Const PARM_JANELA_DIAS As String = "JanelaDias"
Public Const PARM_CAMINHO_CONSOLIDADA As String = "CaminhoConsolidada"
Public Const PARM_CAMINHO_DOWNLOAD As String = "CaminhoDownload"
Public Const PARM_ARQUIVO_LOG As String = "ArquivoLog"
Public Const PARM_MAX_LINHAS_SEGURANCA As String = "MaxLinhasSeguranca"
Public Const PARM_VBTYP_ENTREGA As String = "VBTYP_Entrega"
Public Const PARM_VBTYP_FATURA As String = "VBTYP_Fatura"
