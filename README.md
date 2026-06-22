# FIRST Intelligence - Forecast Estratégico de Estoque, Compras e Locação

## Como usar

1. Mantenha os arquivos fixos em `dados/`:
   - `faturamento.xlsx` com guia `Base`
   - `estoque.xlsx` MATR260
   - `microtech.xlsx` opcional

2. Rode o app:

```bash
streamlit run app.py
```

3. No menu lateral, é possível substituir temporariamente as bases por upload.

## Atualização diária incremental

Use o campo **Adicionar Faturamento Diário** para enviar um ou mais arquivos diários.
O app soma apenas linhas novas, ignorando duplicidades por Nota/Data/Produto/Cliente/Quantidade/Valor.

## Regras principais

- Forecast padrão: horizonte de 30 dias.
- Locação recorrente não entra como consumo de estoque.
- Faturamento de locação alimenta a aba Parque de Locação por Recorrência.
- Grupo oficial vem do estoque/MATR260.
- Produtos com sufixos como `_RV`, `_TC`, `_AT`, `-RV` são agrupados pelo produto-base.
- Aba Microtech usa o arquivo de planejamento do fabricante quando disponível.
