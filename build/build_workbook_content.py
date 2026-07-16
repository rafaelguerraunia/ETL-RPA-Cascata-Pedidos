# -*- coding: utf-8 -*-
"""Builds the plain (macro-free) .xlsx spreadsheet content -- sheets,
Config_SQVI / Config_Geral data, headers -- mirroring Build_Workbook.ps1,
via LibreOffice headless/UNO. VBA injection happens in a later step."""

import os
import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect():
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)
    ctx = resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop


CAMINHO_CONSOLIDADA = r"O:\Shared drives\SmartHub\Automação\ETL (Extract, Transform, Load)"
CAMINHO_DOWNLOAD = CAMINHO_CONSOLIDADA + r"\Download"

SQVI_SPECS = [
    dict(Nome="-RGVS-VBAP", Titulo="E2E VBAP (Pedido)", Tabela="VBAP",
         FELD="VBELN;POSNR;WERKS;MATNR;ERDAT;KWMENG;MEINS;NETWR;ABGRU;PSTYV",
         SELE="WERKS;ERDAT",
         Obrigatorios="VBELN;POSNR;WERKS;ERDAT",
         Aba="PEDIDO", Chave="VBELN;POSNR", PodeCriar="VERDADEIRO"),
    dict(Nome="-RGVS-VBFA", Titulo="E2E VBFA", Tabela="VBFA",
         FELD="VBELN;POSNN;VBTYP_N;ERDAT;ERZET;VBELV;POSNV;VBTYP_V",
         SELE="VBELN;POSNN;VBELV;POSNV",
         Obrigatorios="VBELN;POSNN;VBELV;POSNV;VBTYP_N",
         Aba="VBFA", Chave="VBELV;POSNV;VBELN;POSNN", PodeCriar="VERDADEIRO"),
    dict(Nome="-RGVS-LIPS", Titulo="E2E LIPS", Tabela="LIPS",
         FELD="VBELN;POSNR;PSTYV;MATNR;ARKTX;MATKL;WERKS;LGORT;CHARG;LICHA;LFIMG;VRKME;LGMNG;MEINS;NTGEW;BRGEW;GEWEI;VOLUM;VOLEH;VGBEL;VGPOS",
         SELE="VBELN;POSNR;WERKS;MATNR",
         Obrigatorios="VBELN;POSNR;VGBEL;VGPOS",
         Aba="LIPS", Chave="VBELN;POSNR", PodeCriar="VERDADEIRO"),
    dict(Nome="-RGVS-VTTP", Titulo="E2E VTTP", Tabela="VTTP",
         FELD="VBELN;TKNUM",
         SELE="VBELN",
         Obrigatorios="VBELN;TKNUM",
         Aba="VTTP", Chave="VBELN;TKNUM", PodeCriar="VERDADEIRO"),
    dict(Nome="-RGVS-VTTK", Titulo="E2E VTTK", Tabela="VTTK",
         FELD="TDLNR;TKNUM;SHTYP;TPLST;ERDAT;ERZET;ROUTE;EXTI1;EXTI2;STDIS;DTDIS;UZDIS;STTRG;STTBG;DPTBG;UPTBG;DATBG;UATBG;STTEN;DPTEN;UPTEN;DATEN;UATEN;STABF;DTABF;UZABF",
         SELE="TKNUM;ERDAT;TDLNR;ROUTE",
         Obrigatorios="TKNUM",
         Aba="VTTK", Chave="TKNUM", PodeCriar="VERDADEIRO"),
    dict(Nome="-RGVS-VBRP", Titulo="E2E VBRP", Tabela="VBRP",
         FELD="VBELN;POSNR;VGBEL;VGPOS;FKIMG;VRKME;MATNR;ARKTX;NETWR",
         SELE="VBELN",
         Obrigatorios="VBELN;POSNR",
         Aba="VBRP", Chave="VBELN;POSNR", PodeCriar="VERDADEIRO"),
    dict(Nome="-RGVS-VBRK", Titulo="E2E VBRK", Tabela="VBRK",
         FELD="VBELN;FKART;FKDAT;FKSTO;SFAKN;NETWR;WAERK;KUNAG;BUKRS",
         SELE="VBELN;FKDAT;XBLNR",
         Obrigatorios="VBELN",
         Aba="VBRK", Chave="VBELN", PodeCriar="VERDADEIRO"),
]

GLOBAL_SETTINGS = [
    ("Plantas", "BR10;BR12;BR14"),
    ("JanelaDias", "30"),
    ("CaminhoConsolidada", CAMINHO_CONSOLIDADA),
    ("CaminhoDownload", CAMINHO_DOWNLOAD),
    ("ArquivoLog", "ETL_Cascata_Pedidos.log"),
    ("MaxLinhasSeguranca", "20000"),
    ("VBTYP_Entrega", "J"),
    ("VBTYP_Fatura", "M"),
]

SHEET_ORDER = ["Início", "Config_SQVI", "Config_Geral", "Log", "PEDIDO", "LIPS",
               "VTTP", "VTTK", "VBFA", "VBRP", "VBRK", "Saída Enxuta", "Saída Completa"]


def set_cell(sheet, row, col, value):
    cell = sheet.getCellByPosition(col, row)
    cell.setString(value)


def bold_row(sheet, row, ncols):
    cell_range = sheet.getCellRangeByPosition(0, row, max(ncols - 1, 0), row)
    cell_range.CharWeight = com_sun_star_awt_FontWeight_BOLD


def main():
    global com_sun_star_awt_FontWeight_BOLD
    com_sun_star_awt_FontWeight_BOLD = 150.0  # com.sun.star.awt.FontWeight.BOLD

    desktop = connect()
    doc = desktop.loadComponentFromURL(
        "private:factory/scalc", "_blank", 0, (prop("Hidden", True),))

    sheets = doc.getSheets()
    # rename first auto-created sheet, add the rest
    sheets.getByIndex(0).setName(SHEET_ORDER[0])
    for name in SHEET_ORDER[1:]:
        sheets.insertNewByName(name, sheets.getCount())
    # remove any extra default sheets beyond what we created
    all_names = list(sheets.getElementNames())
    for n in all_names:
        if n not in SHEET_ORDER:
            sheets.removeByName(n)

    ws_inicio = sheets.getByName("Início")
    ws_config_sqvi = sheets.getByName("Config_SQVI")
    ws_config_geral = sheets.getByName("Config_Geral")
    ws_log = sheets.getByName("Log")

    # ---- Config_SQVI ----
    header_sqvi = ["SQVIName", "Titulo", "Tabela", "FELD", "SELE",
                   "CamposObrigatoriosFELD", "AbaDestino", "ChaveUpsert", "PodeCriarAutomaticamente"]
    for c, h in enumerate(header_sqvi):
        set_cell(ws_config_sqvi, 0, c, h)
    bold_row(ws_config_sqvi, 0, len(header_sqvi))
    for r, spec in enumerate(SQVI_SPECS, start=1):
        set_cell(ws_config_sqvi, r, 0, spec["Nome"])
        set_cell(ws_config_sqvi, r, 1, spec["Titulo"])
        set_cell(ws_config_sqvi, r, 2, spec["Tabela"])
        set_cell(ws_config_sqvi, r, 3, spec["FELD"])
        set_cell(ws_config_sqvi, r, 4, spec["SELE"])
        set_cell(ws_config_sqvi, r, 5, spec["Obrigatorios"])
        set_cell(ws_config_sqvi, r, 6, spec["Aba"])
        set_cell(ws_config_sqvi, r, 7, spec["Chave"])
        set_cell(ws_config_sqvi, r, 8, spec["PodeCriar"])
    ws_config_sqvi.getColumns().getByIndex(0)  # no-op touch
    for c in range(len(header_sqvi)):
        ws_config_sqvi.getColumns().getByIndex(c).OptimalWidth = True

    # ---- Config_Geral ----
    set_cell(ws_config_geral, 0, 0, "Parametro")
    set_cell(ws_config_geral, 0, 1, "Valor")
    bold_row(ws_config_geral, 0, 2)
    for r, (k, v) in enumerate(GLOBAL_SETTINGS, start=1):
        set_cell(ws_config_geral, r, 0, k)
        set_cell(ws_config_geral, r, 1, v)
    for c in range(2):
        ws_config_geral.getColumns().getByIndex(c).OptimalWidth = True

    # ---- Log ----
    for c, h in enumerate(["Timestamp", "Nivel", "Hop", "Mensagem"]):
        set_cell(ws_log, 0, c, h)
    bold_row(ws_log, 0, 4)

    # ---- cascade sheets: headers from FELD ----
    feld_by_sqvi = {s["Nome"]: s["FELD"] for s in SQVI_SPECS}
    sheet_to_sqvi = {
        "PEDIDO": "-RGVS-VBAP",
        "LIPS": "-RGVS-LIPS",
        "VTTP": "-RGVS-VTTP",
        "VTTK": "-RGVS-VTTK",
        "VBFA": "-RGVS-VBFA",
        "VBRP": "-RGVS-VBRP",
        "VBRK": "-RGVS-VBRK",
    }
    for sheet_name, sqvi_name in sheet_to_sqvi.items():
        ws = sheets.getByName(sheet_name)
        campos = feld_by_sqvi[sqvi_name].split(";")
        for c, campo in enumerate(campos):
            set_cell(ws, 0, c, campo)
        bold_row(ws, 0, len(campos))

    # ---- Saída Enxuta / Saída Completa placeholders ----
    set_cell(sheets.getByName("Saída Enxuta"), 0, 0,
             "Rode a macro (aba Início) para gerar esta aba automaticamente.")
    set_cell(sheets.getByName("Saída Completa"), 0, 0,
             "Rode a macro (aba Início) para gerar esta aba automaticamente.")

    # ---- Início ----
    title_cell = ws_inicio.getCellByPosition(1, 1)
    title_cell.setString("ETL Cascata SAP - Pedido -> Entrega -> Embarque / Faturamento")
    title_cell.CharHeight = 14
    title_cell.CharWeight = com_sun_star_awt_FontWeight_BOLD

    instrucoes = [
        "Antes da primeira execução:",
        "1) Confirme os caminhos na aba Config_Geral (CaminhoConsolidada / CaminhoDownload).",
        "2) No SAP, ajuste manualmente (ver README.md):",
        "   - -RGVS-VTTP: adicionar TKNUM ao FELD.",
        "   - -RGVS-VBFA: adicionar VBELV e POSNV ao FELD e ao SELE (nesta ordem, ao final); VBTYP_V só ao FELD.",
        "3) Abra o SAP GUI e faça login antes de rodar a macro.",
        "",
        "Para rodar: Alt+F8 -> RunETLCascata -> Executar.",
    ]
    for i, texto in enumerate(instrucoes):
        set_cell(ws_inicio, 3 + i, 1, texto)
    ws_inicio.getColumns().getByIndex(1).Width = 100 * 250  # ~100 chars, in 1/100 mm

    sheets.getByName("Início")  # ensure exists
    doc.getCurrentController().setActiveSheet(ws_inicio)

    out_url = "file://" + os.path.join(HERE, "workbook_plain.xlsx")
    doc.storeToURL(out_url, (prop("FilterName", "Calc MS Excel 2007 XML"), prop("Overwrite", True)))
    doc.close(False)
    print("Saved:", out_url)


if __name__ == "__main__":
    main()
