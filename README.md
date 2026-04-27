# Automação de Faturamento — BASIS

## Uso rápido

```bash
# Interface gráfica
python main.py

# Linha de comando
python main.py --cli batidas.xlsx Faturamento_posto_03_2026.xlsx saida.xlsx
```

## Estrutura

```
faturamento_app/
├── data/
│   ├── input/     ← colocar batidas.xlsx e Faturamento_posto_*.xlsx aqui
│   └── output/    ← arquivos processados (histórico)
├── src/
│   ├── core/regras.py       ← regras de negócio PAR/ÍMPAR, validações
│   ├── io/leitor.py         ← leitura de Excel
│   ├── services/processador.py  ← orquestração
│   ├── services/escritor.py     ← escrita no Excel
│   └── ui/interface.py      ← interface Tkinter
└── main.py
```

## Fluxo

1. Coloca os arquivos em `data/input/`
2. Roda `python main.py`
3. Clica **PROCESSAR**
4. Clica **Abrir resultado**

## Validações

| Camada | O que verifica | Se falhar |
|--------|---------------|-----------|
| C1 | PAR + ÍMPAR = total dias do mês | Para tudo |
| C2 | disp + indisp = max_grupo | Alerta no log |
