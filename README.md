# FIRST Intelligence

App Streamlit para forecast estratégico de estoque, compras e locação.

## Como usar

1. Coloque os arquivos fixos na pasta `dados`:
   - `faturamento.xlsx`
   - `estoque.xlsx`
   - `contratos.xlsx` opcional
2. Execute:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Lógica principal

- Grupo oficial vem somente do estoque/MATR260.
- Códigos com sufixos como `_RV`, `_TC`, `_AT` são cruzados por `Produto_Base`.
- Locação não entra como consumo de estoque para compra.
- Compras são sugeridas com base em venda/outros.
- Locação vira módulo separado de recorrência e expansão de parque.
