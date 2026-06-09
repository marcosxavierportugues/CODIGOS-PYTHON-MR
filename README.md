# Lista de Materiais IFMT - Gerador Universal

Este repositório contém o script `gerar_lista_materiais_total.py`, que lê planilhas Excel geradas pelo AltoQi/Eberick e gera uma lista de materiais organizada com:

- leitura de quantidades com `.` ou `,` como separador decimal
- normalização de descrições de concreto como `Concreto - C-30`
- redução de descrições de forma para `Forma`
- remoção de unidades redundantes (`m³`, `m2`, `kg`) de colunas e descrições
- abas nomeadas conforme `TODOS - <ELEMENTO> - <SEGMENTO>` e `RES- <ELEMENTO> - <SEGMENTO>`
- título de resumo como `RESUMO DE MATERIAIS - ...`
- título de totais como `LISTA DE MATERIAIS TOTAL - ...`
- diálogos com opção de cancelar em todas as janelas

## Uso

1. Instale as dependências:

```powershell
pip install pandas openpyxl
```

2. Execute o script:

```powershell
python gerar_lista_materiais_total.py
```

3. Selecione o arquivo TOTAL e, opcionalmente, arquivos por segmento.
4. Salve a planilha gerada.

## Observações

- O arquivo de backup local `gerar_lista_materiais_total_backup_*.py` é mantido fora do controle de versão.
- Outras planilhas Excel geradas não são comitadas pelo repositório.
