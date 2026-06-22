# FIRST Intelligence

App Streamlit para forecast estratégico de estoque, compras, transferências por ARMZ e inteligência de locação.

## Como rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Bases esperadas

1. Relatório de Faturamento 2026 em Excel, com guia `Base`.
2. Relatório MATR260 de estoque em Excel, contendo a coluna `ARMZ`.

## Lógica padrão

- Horizonte: 30 dias.
- Forecast: 70% consumo dos últimos 30 dias + 30% histórico dos últimos 180 dias.
- Estoque de segurança: parametrizável no menu lateral.
- Transferência: prioriza redistribuição entre ARMZ antes de sugerir compra.
