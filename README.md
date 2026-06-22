# FIRST Intelligence

Forecast estratégico de estoque, compras, locação e planejamento Microtech.

## Como usar

1. Suba a pasta inteira no GitHub/Streamlit.
2. Mantenha os arquivos padrão dentro de `/dados`:
   - `faturamento.xlsx`
   - `estoque.xlsx`
   - `microtech.xlsx` opcional
   - `contratos.xlsx` opcional
3. No app, marque **Usar arquivos fixos do Git/dados** para usar os arquivos base.
4. Para atualizar o faturamento sem substituir a base, envie o **Faturamento Diário**. O app soma apenas linhas novas e ignora duplicidades.

## Destaques da versão V8

- Forecast de estoque sem considerar locação recorrente como consumo.
- Grupo oficial vindo do estoque/MATR260.
- Tratamento de sufixos como `_RV`, `_TC`, `_AT`.
- Análise por ARMZ e sugestão de transferência antes da compra.
- Parque de locação por recorrência.
- Planejamento Estratégico Microtech com:
  - oportunidades por SKU;
  - Rolling Forecast;
  - compra x sell-out;
  - crescimento por família;
  - risco de importação.
