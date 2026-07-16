# ETL Cascata SAP — Pedido → Entrega → Embarque / Faturamento (VBA)

Cascata: Pedido de Venda (VBAP) → VBFA → Entrega (LIPS) → ramifica em
**Embarque** (VTTP→VTTK) e **Faturamento** (VBFA→VBRP→VBRK). Cada passo filtra
pelas chaves do passo anterior (nunca `UNION ALL`). Dados gravados por upsert
(histórico preservado) em abas próprias, mais duas abas finais cruzadas.

Todos os módulos `.bas`/`.cls` já estão prontos nesta pasta. Faltam só duas
coisas manuais: (1) um ajuste pontual em 2 SQVIs no SAP, (2) montar o workbook
uma vez (5–10 min, sem precisar mexer em nenhuma configuração de segurança).

---

## 1) Pré-requisito no SAP (uma vez, antes de rodar)

Transação `SQVI` → **Alterar**:

- **`-RGVS-VTTP`**: aba FELD → adicionar o campo **`TKNUM`** (nada mais).
- **`-RGVS-VBFA`**: aba FELD → adicionar **`VBELV`**, **`POSNV`**, **`VBTYP_V`** (nesta ordem, ao final da lista já existente). Aba SELE → adicionar **`VBELV`**, **`POSNV`** (nesta ordem, ao final — isso é essencial: eles precisam ficar nas posições 3 e 4, depois de VBELN/POSNN que já estão lá).

Sem isso, o programa detecta o problema sozinho, loga um aviso claro e **pula
só o ramo afetado** (não trava o resto) — mas os ramos de Embarque e/ou
Faturamento ficam incompletos até o ajuste ser feito.

## 2) Montar o workbook (uma vez)

1. Abra o Excel, crie uma pasta de trabalho nova e salve como **`ETL_Cascata_Pedidos.xlsm`** (macro-enabled).
2. Crie estas abas (nomes exatos, maiúsculas/minúsculas importam):
   `Início`, `Config_SQVI`, `Config_Geral`, `Log`, `PEDIDO`, `LIPS`, `VTTP`, `VTTK`, `VBFA`, `VBRP`, `VBRK`, `Saída Enxuta`, `Saída Completa`.
3. Na aba **Config_SQVI**, cole o bloco abaixo a partir da célula A1 (selecione A1, Ctrl+V — o Excel separa por colunas automaticamente pois é TAB-delimitado):

```
SQVIName	Titulo	Tabela	FELD	SELE	CamposObrigatoriosFELD	AbaDestino	ChaveUpsert	PodeCriarAutomaticamente
-RGVS-VBAP	E2E VBAP (Pedido)	VBAP	VBELN;POSNR;WERKS;MATNR;LOEKZ;ERDAT;KWMENG;MEINS;NETWR;ABGRU;PSTYV	WERKS;LOEKZ;ERDAT	VBELN;POSNR;WERKS;LOEKZ;ERDAT	PEDIDO	VBELN;POSNR	VERDADEIRO
-RGVS-VBFA	E2E VBFA	VBFA	VBELN;POSNN;VBTYP_N;ERDAT;ERZET;VBELV;POSNV;VBTYP_V	VBELN;POSNN;VBELV;POSNV	VBELN;POSNN;VBELV;POSNV;VBTYP_N	VBFA	VBELV;POSNV;VBELN;POSNN	FALSO
-RGVS-LIPS	E2E LIPS	LIPS	VBELN;POSNR;PSTYV;MATNR;ARKTX;MATKL;WERKS;LGORT;CHARG;LICHA;LFIMG;VRKME;LGMNG;MEINS;NTGEW;BRGEW;GEWEI;VOLUM;VOLEH;VGBEL;VGPOS	VBELN;POSNR;WERKS;MATNR	VBELN;POSNR;VGBEL;VGPOS	LIPS	VBELN;POSNR	FALSO
-RGVS-VTTP	E2E VTTP	VTTP	VBELN;TPRFO;TKNUM	VBELN	VBELN;TKNUM	VTTP	VBELN;TKNUM	FALSO
-RGVS-VTTK	E2E VTTK	VTTK	TDLNR;TKNUM;SHTYP;TPLST;ERDAT;ERZET;ROUTE;EXTI1;EXTI2;STDIS;DTDIS;UZDIS;STTRG;STTBG;DPTBG;UPTBG;DATBG;UATBG;STTEN;DPTEN;UPTEN;DATEN;UATEN;STABF;DTABF;UZABF	TKNUM;ERDAT;TDLNR;ROUTE	TKNUM	VTTK	TKNUM	FALSO
-RGVS-VBRP	E2E VBRP	VBRP	VBELN;POSNR;VGBEL;VGPOS;FKIMG;VRKME;MATNR;ARKTX;NETWR	VBELN	VBELN;POSNR	VBRP	VBELN;POSNR	FALSO
-RGVS-VBRK	E2E VBRK	VBRK	VBELN;FKART;FKDAT;FKSTO;SFAKN;BELNR;GJAHR;BUKRS	VBELN;FKDAT;XBLNR	VBELN	VBRK	VBELN	FALSO
```

4. Na aba **Config_Geral**, cole a partir de A1:

```
Parametro	Valor
Plantas	BR10;BR12;BR14
JanelaDias	30
CaminhoConsolidada	O:\Shared drives\SmartHub\Automação\ETL (Extract, Transform, Load)
CaminhoDownload	O:\Shared drives\SmartHub\Automação\ETL (Extract, Transform, Load)\Download
ArquivoLog	ETL_Cascata_Pedidos.log
MaxLinhasSeguranca	20000
VBTYP_Entrega	J
VBTYP_Fatura	M
```

   Ajuste `CaminhoConsolidada`/`CaminhoDownload` se a pasta real for outra.

5. Na aba **Log**, célula A1: cole o cabeçalho `Timestamp	Nivel	Hop	Mensagem`.
   As demais abas (`PEDIDO`, `LIPS`, `VTTP`, `VTTK`, `VBFA`, `VBRP`, `VBRK`,
   `Saída Enxuta`, `Saída Completa`) **não precisam de nada** — o programa
   escreve os cabeçalhos sozinho na primeira execução.

6. Abra o Editor VBA (**Alt+F11**). Botão direito no nome do projeto (`VBAProject (ETL_Cascata_Pedidos.xlsm)`) → **Importar Arquivo…** e importe, um por um, todos os arquivos desta pasta:

   ```
   clsSQVISpec.cls
   clsFiltro.cls
   clsRunResult.cls
   clsLinhaResolvida.cls
   modConstantes.bas
   modLog.bas
   modUtils.bas
   modSAPConnection.bas
   modConfig.bas
   modSQVIEngine.bas
   modExcelWrite.bas
   modCascade.bas
   modJoinBuilder.bas
   modMain.bas
   ```

   (Isso não precisa de nenhuma configuração especial de segurança — é diferente do acesso programático que tentei automatizar antes.)

7. (Opcional) Na aba **Início**, insira uma forma/botão (Inserir → Formas) e em
   **Atribuir Macro** escolha `RunETLCascata`. Sem isso, basta rodar via
   **Alt+F8 → RunETLCascata → Executar**.

8. Salve (Ctrl+S, mantendo o formato `.xlsm`).

## 3) Rodar

1. Abra o SAP GUI e faça login normalmente.
2. Abra `ETL_Cascata_Pedidos.xlsm` e rode a macro (botão, ou Alt+F8).
3. Acompanhe pela aba `Log` (ou pelo arquivo `.log` na pasta consolidada).
4. Ao final, aparece um resumo (pedidos processados, pedidos sem entrega
   ainda, hops pulados por erro/gap).

## 4) O que cada aba guarda

| Aba | Conteúdo |
|---|---|
| `Config_SQVI` | Especificação das 7 SQVIs — única fonte de verdade, não precisa editar código para ajustar campos |
| `Config_Geral` | Plantas, janela de dias, caminhos, categorias VBTYP |
| `Log` | Histórico de execuções (INFO/WARN/ERROR/FATAL) |
| `PEDIDO` | Pedidos de venda filtrados (BR10/12/14, não deletados, últimos 30 dias) |
| `LIPS`, `VTTP`, `VTTK`, `VBFA`, `VBRP`, `VBRK` | Uma aba por tabela SAP, acumulada por upsert (histórico preservado entre execuções) |
| `Saída Enxuta` | Cruzamento final, campos curados (reconstruída do zero a cada execução, a partir de TODO o histórico acumulado) |
| `Saída Completa` | Cruzamento final, todos os campos de todas as tabelas |

Pedidos sem entrega/embarque/fatura ainda aparecem nas duas abas finais, com
as colunas seguintes em branco (processo em andamento).

## 5) Pontos de atenção (verificar na primeira execução real)

- **`-RGVS-LIPS` tem `LICHA` duplicado na gravação original** — se a SQVI real
  tiver essa coluna repetida, as colunas depois dela no export ficarão
  deslocadas em 1 posição. Confira as primeiras linhas exportadas contra a
  ordem declarada em `Config_SQVI` antes de confiar no upsert em produção.
- **`-RGVS-VBRP`**: a gravação mostra `SELE = VBELN` apenas (o resumo inicial
  citava também POSNR/WERKS/MATNR). Para esta cascata só precisamos filtrar
  por VBELN, então não bloqueia — mas vale conferir em SQVI/modo Exibir.
- **Formato de data**: `modUtils.DataSAP` usa `MM/DD/YYYY` (mesmo formato do
  RPA BASE.vbs). Se o seu usuário SAP tiver outro formato regional, ajuste essa
  função.
- **Categorias VBTYP** (`J`=Entrega, `M`=Fatura) são o padrão SAP mais comum;
  confirme se batem com o seu sistema (ajustável em `Config_Geral`, sem
  precisar mexer em código).
- Todos os textos de campo no cabeçalho exportado pelo SAP são a
  **descrição** do campo, não o nome técnico — por isso o programa lê os
  dados **por posição** (ordem do FELD em `Config_SQVI`), não por nome de
  coluna. Se você reordenar campos manualmente em uma SQVI existente, precisa
  também reordenar a lista FELD correspondente em `Config_SQVI`.
