# Build alternativo (Linux/Mac, sem Excel)

O `Build_Workbook.ps1` na raiz do repositório monta o `.xlsm` via automação COM do
Excel e só roda no Windows com Excel instalado. Estes scripts fazem a mesma coisa
sem depender de Windows/Excel:

- `ovba_compress.py` — codec do formato "Compressed Container" do MS-OVBA
  (compressão usada dentro do `vbaProject.bin`).
- `cfb.py` — escritor do formato de arquivo composto OLE/CFB (o container binário
  do `vbaProject.bin`), com suporte a mini-stream/MiniFAT.
- `build_vba_project.py` — monta o `vbaProject.bin` real a partir dos módulos em
  `vba/*.bas` e `vba/*.cls`, seguindo a especificação MS-OVBA da Microsoft.
- `build_workbook_content.py` — monta as 13 abas e os dados de `Config_SQVI` /
  `Config_Geral` via LibreOffice headless (UNO), espelhando o `Build_Workbook.ps1`.
- `assemble_xlsm.py` — injeta o `vbaProject.bin` no `.xlsx` gerado, transformando-o
  em `.xlsm` macro-habilitado (ajusta `[Content_Types].xml` e os relacionamentos).

## Como rodar

Requer `python3` e um LibreOffice headless escutando via UNO:

```bash
soffice --headless --invisible --accept="socket,host=localhost,port=2002;urp;" &
python3 build/build_vba_project.py       # gera build/vbaProject.bin
python3 build/build_workbook_content.py  # gera build/workbook_plain.xlsx
python3 build/assemble_xlsm.py           # gera ETL_Cascata_Pedidos.xlsm
```

## O que foi validado

- Todos os 14 módulos comprimidos/descomprimidos batem byte a byte com o
  código-fonte original (teste automatizado em `ovba_compress.py`).
- O `vbaProject.bin` gerado foi verificado com `oletools`/`olevba` (biblioteca
  independente de leitura de macros): zero avisos, extração perfeita dos 14
  módulos.
- O LibreOffice abre o `.xlsm` final sem nenhum diálogo de erro/reparo; as 13
  abas e os dados de configuração aparecem corretos.

## Ressalva importante

**Não foi possível testar em um Excel real** (este ambiente não tem Windows/Excel).
Ao tentar confirmar que os módulos aparecem como editáveis dentro do próprio editor
Basic do LibreOffice, eles apareceram como uma biblioteca vazia — o projeto é
reconhecido, mas o LibreOffice não populou os módulos na interface. Não ficou claro
se isso é uma limitação conhecida do suporte a VBA do LibreOffice (que já mostrou
outras lacunas nesta investigação) ou um problema real do binário.

**Ao abrir `ETL_Cascata_Pedidos.xlsm` no Excel de verdade, confira primeiro:**
Alt+F11 → o projeto VBA deve mostrar os 14 módulos (10 em "Modules", 4 em
"Class Modules") com o código-fonte completo. Se aparecer vazio ou o Excel
mostrar um aviso de reparo, **o fallback é o processo manual já descrito no
README principal** (Alt+F11 → Importar Arquivo, um por um, arquivos de `vba/`).
