# FIRST Intelligence

Forecast estratégico de estoque, compras e locação.

## Como usar no Streamlit Cloud

Mantenha os arquivos padrão na pasta `dados`:

- `dados/faturamento.xlsx`
- `dados/estoque.xlsx`
- `dados/contratos.xlsx` opcional

O app usa esses arquivos automaticamente. Também é possível substituir as bases pela tela lateral com upload temporário.

## Ajustes da versão

- Locação não entra como consumo de estoque.
- Locação alimenta apenas a aba de inteligência de locação.
- Grupo oficial vem do estoque/MATR260.
- Valores monetários exibidos em formato brasileiro.
- Legendas técnicas removidas das telas executivas.
- Cobertura exibida em dias e meses.
- Tratamento de códigos com sufixos como `_RV`, `_TC`, `_AT`.
