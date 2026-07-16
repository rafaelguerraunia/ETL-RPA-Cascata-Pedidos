# Build_Workbook.ps1
# Monta o workbook ETL_Cascata_Pedidos.xlsm: cria as abas, popula Config_SQVI/Config_Geral,
# importa os modulos VBA (.bas/.cls) e salva como .xlsm.
#
# Pre-requisito (uma vez, no Excel): Arquivo > Opcoes > Central de Confiabilidade >
# Configuracoes da Central de Confiabilidade > Configuracoes de Macro > marcar
# "Confiar no acesso ao modelo de objeto de projetos do VBA".
#
# Uso:  powershell -File Build_Workbook.ps1

$ErrorActionPreference = "Stop"

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputPath = Join-Path $scriptDir "ETL_Cascata_Pedidos.xlsm"

# Caminhos reais de producao (mesma pasta ja usada pelo RPA BASE.vbs) - gravados em
# Config_Geral para quando o arquivo for movido para o ambiente real do usuario.
$caminhoConsolidada = "O:\Shared drives\SmartHub\Automação\ETL (Extract, Transform, Load)"
$caminhoDownload     = "$caminhoConsolidada\Download"

Write-Host "Removendo build anterior, se existir..."
if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

# ===================== Especificacao das 7 SQVIs =====================
# Colunas: SQVIName, Titulo, Tabela, FELD, SELE, CamposObrigatoriosFELD, AbaDestino, ChaveUpsert, PodeCriarAutomaticamente
$sqviSpecs = @(
    @{
        Nome = "-RGVS-VBAP"; Titulo = "E2E VBAP (Pedido)"; Tabela = "VBAP"
        FELD = "VBELN;POSNR;WERKS;MATNR;LOEKZ;ERDAT;KWMENG;MEINS;NETWR;ABGRU;PSTYV"
        SELE = "WERKS;LOEKZ;ERDAT"
        Obrigatorios = "VBELN;POSNR;WERKS;LOEKZ;ERDAT"
        Aba = "PEDIDO"; Chave = "VBELN;POSNR"; PodeCriar = "VERDADEIRO"
    },
    @{
        Nome = "-RGVS-VBFA"; Titulo = "E2E VBFA"; Tabela = "VBFA"
        FELD = "VBELN;POSNN;VBTYP_N;ERDAT;ERZET;VBELV;POSNV;VBTYP_V"
        SELE = "VBELN;POSNN;VBELV;POSNV"
        Obrigatorios = "VBELN;POSNN;VBELV;POSNV;VBTYP_N"
        Aba = "VBFA"; Chave = "VBELV;POSNV;VBELN;POSNN"; PodeCriar = "FALSO"
    },
    @{
        Nome = "-RGVS-LIPS"; Titulo = "E2E LIPS"; Tabela = "LIPS"
        FELD = "VBELN;POSNR;PSTYV;MATNR;ARKTX;MATKL;WERKS;LGORT;CHARG;LICHA;LFIMG;VRKME;LGMNG;MEINS;NTGEW;BRGEW;GEWEI;VOLUM;VOLEH;VGBEL;VGPOS"
        SELE = "VBELN;POSNR;WERKS;MATNR"
        Obrigatorios = "VBELN;POSNR;VGBEL;VGPOS"
        Aba = "LIPS"; Chave = "VBELN;POSNR"; PodeCriar = "FALSO"
    },
    @{
        Nome = "-RGVS-VTTP"; Titulo = "E2E VTTP"; Tabela = "VTTP"
        FELD = "VBELN;TPRFO;TKNUM"
        SELE = "VBELN"
        Obrigatorios = "VBELN;TKNUM"
        Aba = "VTTP"; Chave = "VBELN;TKNUM"; PodeCriar = "FALSO"
    },
    @{
        Nome = "-RGVS-VTTK"; Titulo = "E2E VTTK"; Tabela = "VTTK"
        FELD = "TDLNR;TKNUM;SHTYP;TPLST;ERDAT;ERZET;ROUTE;EXTI1;EXTI2;STDIS;DTDIS;UZDIS;STTRG;STTBG;DPTBG;UPTBG;DATBG;UATBG;STTEN;DPTEN;UPTEN;DATEN;UATEN;STABF;DTABF;UZABF"
        SELE = "TKNUM;ERDAT;TDLNR;ROUTE"
        Obrigatorios = "TKNUM"
        Aba = "VTTK"; Chave = "TKNUM"; PodeCriar = "FALSO"
    },
    @{
        Nome = "-RGVS-VBRP"; Titulo = "E2E VBRP"; Tabela = "VBRP"
        FELD = "VBELN;POSNR;VGBEL;VGPOS;FKIMG;VRKME;MATNR;ARKTX;NETWR"
        SELE = "VBELN"
        Obrigatorios = "VBELN;POSNR"
        Aba = "VBRP"; Chave = "VBELN;POSNR"; PodeCriar = "FALSO"
    },
    @{
        Nome = "-RGVS-VBRK"; Titulo = "E2E VBRK"; Tabela = "VBRK"
        FELD = "VBELN;FKART;FKDAT;FKSTO;SFAKN;BELNR;GJAHR;BUKRS"
        SELE = "VBELN;FKDAT;XBLNR"
        Obrigatorios = "VBELN"
        Aba = "VBRK"; Chave = "VBELN"; PodeCriar = "FALSO"
    }
)

$globalSettings = @(
    @{ Nome = "Plantas";              Valor = "BR10;BR12;BR14" },
    @{ Nome = "JanelaDias";           Valor = "30" },
    @{ Nome = "CaminhoConsolidada";   Valor = $caminhoConsolidada },
    @{ Nome = "CaminhoDownload";      Valor = $caminhoDownload },
    @{ Nome = "ArquivoLog";           Valor = "ETL_Cascata_Pedidos.log" },
    @{ Nome = "MaxLinhasSeguranca";   Valor = "20000" },
    @{ Nome = "VBTYP_Entrega";        Valor = "J" },
    @{ Nome = "VBTYP_Fatura";         Valor = "M" }
)

# ===================== Excel COM =====================
Write-Host "Abrindo Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()

    # remove folhas extras que o Add() cria (mantem so a primeira, que vira "Inicio")
    while ($wb.Sheets.Count -gt 1) {
        $wb.Sheets.Item($wb.Sheets.Count).Delete()
    }
    $wsInicio = $wb.Sheets.Item(1)
    $wsInicio.Name = "Início"

    function New-Sheet($nome) {
        $novaAba = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
        $novaAba.Name = $nome
        return $novaAba
    }

    Write-Host "Criando abas..."
    $wsConfigSQVI  = New-Sheet "Config_SQVI"
    $wsConfigGeral = New-Sheet "Config_Geral"
    $wsLog         = New-Sheet "Log"
    $wsPedido      = New-Sheet "PEDIDO"
    $wsLIPS        = New-Sheet "LIPS"
    $wsVTTP        = New-Sheet "VTTP"
    $wsVTTK        = New-Sheet "VTTK"
    $wsVBFA        = New-Sheet "VBFA"
    $wsVBRP        = New-Sheet "VBRP"
    $wsVBRK        = New-Sheet "VBRK"
    $wsEnxuta      = New-Sheet "Saída Enxuta"
    $wsCompleta    = New-Sheet "Saída Completa"

    # -------- Config_SQVI --------
    Write-Host "Populando Config_SQVI..."
    $cabecalhoSQVI = @("SQVIName","Titulo","Tabela","FELD","SELE","CamposObrigatoriosFELD","AbaDestino","ChaveUpsert","PodeCriarAutomaticamente")
    for ($c = 0; $c -lt $cabecalhoSQVI.Count; $c++) { $wsConfigSQVI.Cells.Item(1, $c + 1) = $cabecalhoSQVI[$c] }
    $wsConfigSQVI.Rows.Item(1).Font.Bold = $true

    $r = 2
    foreach ($spec in $sqviSpecs) {
        $wsConfigSQVI.Cells.Item($r, 1) = $spec.Nome
        $wsConfigSQVI.Cells.Item($r, 2) = $spec.Titulo
        $wsConfigSQVI.Cells.Item($r, 3) = $spec.Tabela
        $wsConfigSQVI.Cells.Item($r, 4) = $spec.FELD
        $wsConfigSQVI.Cells.Item($r, 5) = $spec.SELE
        $wsConfigSQVI.Cells.Item($r, 6) = $spec.Obrigatorios
        $wsConfigSQVI.Cells.Item($r, 7) = $spec.Aba
        $wsConfigSQVI.Cells.Item($r, 8) = $spec.Chave
        $wsConfigSQVI.Cells.Item($r, 9) = $spec.PodeCriar
        $r++
    }
    $wsConfigSQVI.Columns.Item("A:I").EntireColumn.AutoFit() | Out-Null

    # -------- Config_Geral --------
    Write-Host "Populando Config_Geral..."
    $wsConfigGeral.Cells.Item(1,1) = "Parametro"
    $wsConfigGeral.Cells.Item(1,2) = "Valor"
    $wsConfigGeral.Rows.Item(1).Font.Bold = $true
    $r = 2
    foreach ($p in $globalSettings) {
        $wsConfigGeral.Cells.Item($r, 1) = $p.Nome
        $wsConfigGeral.Cells.Item($r, 2) = $p.Valor
        $r++
    }
    $wsConfigGeral.Columns.Item("A:B").EntireColumn.AutoFit() | Out-Null

    # -------- Log --------
    $cabecalhoLog = @("Timestamp","Nivel","Hop","Mensagem")
    for ($c = 0; $c -lt $cabecalhoLog.Count; $c++) { $wsLog.Cells.Item(1, $c + 1) = $cabecalhoLog[$c] }
    $wsLog.Rows.Item(1).Font.Bold = $true

    # -------- Cabecalhos das abas de cascata (a partir do FELD de cada spec) --------
    function Set-Headers($ws, $feldTexto) {
        $campos = $feldTexto -split ";"
        for ($c = 0; $c -lt $campos.Count; $c++) { $ws.Cells.Item(1, $c + 1) = $campos[$c] }
        $ws.Rows.Item(1).Font.Bold = $true
    }
    Set-Headers $wsPedido ($sqviSpecs | Where-Object { $_.Nome -eq "-RGVS-VBAP" }).FELD
    Set-Headers $wsLIPS   ($sqviSpecs | Where-Object { $_.Nome -eq "-RGVS-LIPS" }).FELD
    Set-Headers $wsVTTP   ($sqviSpecs | Where-Object { $_.Nome -eq "-RGVS-VTTP" }).FELD
    Set-Headers $wsVTTK   ($sqviSpecs | Where-Object { $_.Nome -eq "-RGVS-VTTK" }).FELD
    Set-Headers $wsVBFA   ($sqviSpecs | Where-Object { $_.Nome -eq "-RGVS-VBFA" }).FELD
    Set-Headers $wsVBRP   ($sqviSpecs | Where-Object { $_.Nome -eq "-RGVS-VBRP" }).FELD
    Set-Headers $wsVBRK   ($sqviSpecs | Where-Object { $_.Nome -eq "-RGVS-VBRK" }).FELD

    $wsEnxuta.Cells.Item(1,1)   = "Rode a macro (aba Início) para gerar esta aba automaticamente."
    $wsCompleta.Cells.Item(1,1) = "Rode a macro (aba Início) para gerar esta aba automaticamente."

    # -------- Aba Início --------
    $wsInicio.Cells.Item(2,2)  = "ETL Cascata SAP - Pedido -> Entrega -> Embarque / Faturamento"
    $wsInicio.Cells.Item(2,2).Font.Size = 14
    $wsInicio.Cells.Item(2,2).Font.Bold = $true

    $instrucoes = @(
        "Antes da primeira execução:",
        "1) Confirme os caminhos na aba Config_Geral (CaminhoConsolidada / CaminhoDownload).",
        "2) No SAP, ajuste manualmente (ver README.md):",
        "   - -RGVS-VTTP: adicionar TKNUM ao FELD.",
        "   - -RGVS-VBFA: adicionar VBELV e POSNV ao FELD e ao SELE (nesta ordem, ao final); VBTYP_V só ao FELD.",
        "3) Abra o SAP GUI e faça login antes de rodar a macro.",
        "",
        "Para rodar: clique no botão 'Rodar ETL Cascata' abaixo, ou Alt+F8 -> RunETLCascata."
    )
    $linha = 4
    foreach ($texto in $instrucoes) {
        $wsInicio.Cells.Item($linha, 2) = $texto
        $linha++
    }
    $wsInicio.Columns.Item("B").ColumnWidth = 100

    Write-Host "Importando módulos VBA..."
    $vbProj = $wb.VBProject

    # classes primeiro, depois módulos (ordem não é estritamente necessária, mas fica organizado)
    $arquivosClasse = @("clsSQVISpec.cls","clsFiltro.cls","clsRunResult.cls","clsLinhaResolvida.cls")
    $arquivosModulo = @("modConstantes.bas","modLog.bas","modUtils.bas","modSAPConnection.bas","modConfig.bas", `
                         "modSQVIEngine.bas","modExcelWrite.bas","modCascade.bas","modJoinBuilder.bas","modMain.bas")

    foreach ($f in ($arquivosClasse + $arquivosModulo)) {
        $caminhoCompleto = Join-Path $scriptDir $f
        if (-not (Test-Path $caminhoCompleto)) {
            throw "Arquivo de módulo não encontrado: $caminhoCompleto"
        }
        Write-Host "  Importando $f"
        $vbProj.VBComponents.Import($caminhoCompleto) | Out-Null
    }

    # -------- Botão na aba Início --------
    Write-Host "Adicionando botão de execução..."
    $btn = $wsInicio.Buttons().Add(200, 60, 220, 30)
    $btn.OnAction = "RunETLCascata"
    $btn.Characters.Text = "Rodar ETL Cascata"

    $wsInicio.Activate()

    Write-Host "Salvando como $outputPath ..."
    # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    $wb.SaveAs($outputPath, 52)
    $wb.Close($false)

    Write-Host "OK: workbook criado em $outputPath"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
