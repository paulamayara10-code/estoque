# -*- coding: utf-8 -*-
"""
FIRST Intelligence | Forecast Estratégico de Estoque, Compras e Locação

V4 - lógica ajustada para a realidade First:
- Arquivos fixos na pasta /dados, com opção de substituição por upload.
- Grupo oficial vindo somente do estoque/MATR260.
- Cruzamento por Produto_Base, tratando sufixos como _RV, _TC, _AT, -RV etc.
- Locação NÃO consome estoque para compra imediata.
- Motor separado para Parque de Locação com recorrência cliente + produto.
"""

from __future__ import annotations

import math
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

APP_NAME = "FIRST Intelligence"
APP_SUBTITLE = "Forecast Estratégico de Estoque, Compras e Locação"
BASE_DIR = Path(__file__).resolve().parent
DADOS_DIR = BASE_DIR / "dados"
DEFAULT_FATURAMENTO = DADOS_DIR / "faturamento.xlsx"
DEFAULT_ESTOQUE = DADOS_DIR / "estoque.xlsx"
DEFAULT_CONTRATOS = DADOS_DIR / "contratos.xlsx"

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# ESTILO
# =========================
st.markdown(
    """
    <style>
        .main {background-color: #f7f9fc;}
        .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
        .first-header {
            background: linear-gradient(135deg, #071b35 0%, #0b3c78 55%, #1167b1 100%);
            padding: 22px 26px; border-radius: 18px; color: white; margin-bottom: 18px;
            box-shadow: 0 10px 25px rgba(7,27,53,.18);
        }
        .first-header h1 {font-size: 34px; margin: 0; font-weight: 800;}
        .first-header p {font-size: 15px; margin: 6px 0 0 0; opacity: .92;}
        .kpi-card {
            background: white; border-radius: 16px; padding: 18px 18px; min-height: 112px;
            box-shadow: 0 6px 18px rgba(11,60,120,.09); border: 1px solid #e9eef7;
        }
        .kpi-label {font-size: 13px; color: #536173; margin-bottom: 8px; font-weight: 600;}
        .kpi-value {font-size: 25px; color: #071b35; font-weight: 800; line-height: 1.1;}
        .kpi-help {font-size: 12px; color: #7b8794; margin-top: 8px;}
        .section-title {font-size: 21px; font-weight: 800; color: #071b35; margin-top: 10px;}
        div[data-testid="stMetricValue"] {font-size: 24px;}
        .small-note {font-size: 12px; color: #6b7280;}
        .logic-box {background:#ffffff; border:1px solid #e9eef7; border-radius:14px; padding:15px; margin: 8px 0;}
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================
# UTILITÁRIOS
# =========================
def norm_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    repl = str.maketrans("ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇÑ", "AAAAAEEEEIIIIOOOOOUUUUCN")
    text = text.translate(repl)
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_code(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper().replace(" ", "")
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def produto_base(value: object) -> str:
    """Normaliza códigos para cruzamento, removendo sufixos operacionais.

    Exemplos:
    010_RV -> 010
    010-TC -> 010
    DI4000BIVBR_AT -> DI4000BIVBR
    """
    text = normalize_code(value)
    if not text:
        return ""
    text = re.sub(r"([_\-\.])(RV|TC|AT|AV|LT|LOC|LOCA|MNT|MANUT|REP|RENT|R)$", "", text)
    # remove sufixo final quando for composto por 1 a 4 letras após separador
    text = re.sub(r"([_\-])[A-Z]{1,4}$", "", text)
    return text


def normalize_col(col: object) -> str:
    text = norm_text(col)
    text = re.sub(r"[^A-Z0-9]+", "_", text).strip("_")
    return text


def find_col(df: pd.DataFrame, candidates: list[str], required: bool = True) -> Optional[str]:
    normalized = {normalize_col(c): c for c in df.columns}
    for cand in candidates:
        key = normalize_col(cand)
        if key in normalized:
            return normalized[key]
    for cand in candidates:
        key = normalize_col(cand)
        for norm, original in normalized.items():
            if key and key in norm:
                return original
    if required:
        raise ValueError(f"Coluna não encontrada. Tentei: {candidates}. Colunas disponíveis: {list(df.columns)}")
    return None


def brl(value: float) -> str:
    try:
        value = float(value)
    except Exception:
        value = 0.0
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_num(value: float, casas: int = 0) -> str:
    try:
        value = float(value)
    except Exception:
        value = 0.0
    if np.isinf(value) or np.isnan(value):
        return "-"
    return f"{value:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def kpi_card(label: str, value: str, help_text: str = ""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-help">{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def to_excel_download(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe_name = name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
            ws = writer.sheets[safe_name]
            for col_idx, col_name in enumerate(df.columns, start=1):
                width = min(max(len(str(col_name)) + 2, 12), 42)
                ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width
    return output.getvalue()


def read_file_bytes(uploaded_file, default_path: Path) -> tuple[Optional[bytes], str]:
    if uploaded_file is not None:
        return uploaded_file.getvalue(), "Upload da sessão"
    if default_path.exists():
        return default_path.read_bytes(), f"Arquivo fixo: {default_path.name}"
    return None, "Não encontrado"

# =========================
# LEITURA DE BASES
# =========================
@st.cache_data(show_spinner=False)
def load_faturamento(file_bytes: bytes, sheet_name: str = "Base") -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]

    col_data = find_col(df, ["DT Emissao", "Data Emissao", "Emissao"])
    col_produto = find_col(df, ["Produto", "Codigo", "Cod Produto"])
    col_desc = find_col(df, ["DESCRIÇÃO", "Descricao", "Descrição"], required=False)
    col_qtd = find_col(df, ["Quantidade", "Qtd", "Qtde"])
    col_valor = find_col(df, ["Vlr.Total", "Valor Total", "Valor Bruto"], required=False)
    col_cliente = find_col(df, ["Nome Cliente", "Cliente", "Razao Social", "Razão Social"], required=False)
    col_grupo = find_col(df, ["GRUPO", "Grupo"], required=False)
    col_linha = find_col(df, ["LINHA DE PRODUTO", "Linha de Produto"], required=False)
    col_nova = find_col(df, ["NOVA", "Classificacao Nova"], required=False)
    col_class = find_col(df, ["CLASSIFICAÇÃO", "Classificacao"], required=False)
    col_finalidade = find_col(df, ["FINALIDADE", "Finalidade"], required=False)
    col_categoria = find_col(df, ["CATEGORIA", "Categoria"], required=False)
    col_nf = find_col(df, ["Nota Fiscal", "NF"], required=False)

    out = pd.DataFrame()
    out["Produto"] = df[col_produto].map(normalize_code)
    out["Produto_Base"] = out["Produto"].map(produto_base)
    out["Descrição"] = df[col_desc].fillna("").astype(str).str.strip() if col_desc else ""
    out["Data"] = pd.to_datetime(df[col_data], errors="coerce", dayfirst=True)
    out["Mês"] = out["Data"].dt.to_period("M").astype(str)
    out["Quantidade"] = pd.to_numeric(df[col_qtd], errors="coerce").fillna(0)
    out["Valor"] = pd.to_numeric(df[col_valor], errors="coerce").fillna(0) if col_valor else 0
    out["Cliente"] = df[col_cliente].fillna("").astype(str).str.strip() if col_cliente else ""
    out["Grupo_Faturamento"] = df[col_grupo].fillna("").astype(str).str.strip() if col_grupo else ""
    out["Linha_Faturamento"] = df[col_linha].fillna("").astype(str).str.strip() if col_linha else ""
    out["Nova"] = df[col_nova].fillna("").astype(str).str.strip() if col_nova else ""
    out["Classificação"] = df[col_class].fillna("").astype(str).str.strip() if col_class else ""
    out["Finalidade"] = df[col_finalidade].fillna("").astype(str).str.strip() if col_finalidade else ""
    out["Categoria"] = df[col_categoria].fillna("").astype(str).str.strip() if col_categoria else ""
    out["Nota Fiscal"] = df[col_nf].fillna("").astype(str).str.strip() if col_nf else ""

    texto_loc = (
        out["Nova"].map(norm_text) + " " +
        out["Classificação"].map(norm_text) + " " +
        out["Finalidade"].map(norm_text) + " " +
        out["Categoria"].map(norm_text)
    )
    out["Tipo Operação"] = np.where(texto_loc.str.contains("LOCACAO|LOCACÃO|LOCAÇÃO", regex=True), "Locação", "Venda/Outros")
    out = out[(out["Produto_Base"] != "") & out["Data"].notna()].copy()
    return out


@st.cache_data(show_spinner=False)
def load_estoque(file_bytes: bytes, sheet_name: Optional[str] = None) -> pd.DataFrame:
    if sheet_name:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)
    else:
        xl = pd.ExcelFile(BytesIO(file_bytes))
        target = None
        for s in xl.sheet_names:
            if "RELACAO" in normalize_col(s) or "POSICAO" in normalize_col(s):
                target = s
                break
        target = target or xl.sheet_names[-1]
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=target)

    df.columns = [str(c).strip() for c in df.columns]
    col_codigo = find_col(df, ["CODIGO", "Código", "Produto"])
    col_desc = find_col(df, ["DESCRICAO", "Descrição", "DESCRIÇÃO"], required=False)
    col_grupo = find_col(df, ["GRUPO", "Grupo"], required=False)
    col_armz = find_col(df, ["ARMZ", "Armazem", "Armazém", "Local"])
    col_saldo = find_col(df, ["SALDO EM ESTOQUE", "Saldo", "Saldo Estoque"])
    col_empenho = find_col(df, ["EMPENHO PARA REQ/PV/RESERVA", "Empenho"], required=False)
    col_disp = find_col(df, ["ESTOQUE DISPONIVEL", "Disponivel", "Disponível"], required=False)
    col_valor = find_col(df, ["VALOR EM ESTOQUE", "Valor Estoque", "Valor"], required=False)

    for col in [col_codigo, col_desc, col_grupo]:
        if col:
            df[col] = df[col].ffill()

    out = pd.DataFrame()
    out["Produto"] = df[col_codigo].map(normalize_code)
    out["Produto_Base"] = out["Produto"].map(produto_base)
    out["Descrição Estoque"] = df[col_desc].fillna("").astype(str).str.strip() if col_desc else ""
    # Grupo oficial vem SOMENTE do estoque.
    out["Grupo"] = df[col_grupo].fillna("").astype(str).str.strip() if col_grupo else ""
    out["ARMZ"] = df[col_armz].fillna("").astype(str).str.strip()
    out["Saldo Estoque"] = pd.to_numeric(df[col_saldo], errors="coerce").fillna(0)
    out["Empenho"] = pd.to_numeric(df[col_empenho], errors="coerce").fillna(0) if col_empenho else 0
    if col_disp:
        out["Estoque Disponível"] = pd.to_numeric(df[col_disp], errors="coerce").fillna(0)
    else:
        out["Estoque Disponível"] = out["Saldo Estoque"] - out["Empenho"]
    out["Valor Estoque"] = pd.to_numeric(df[col_valor], errors="coerce").fillna(0) if col_valor else 0
    out = out[(out["Produto_Base"] != "") & (out["ARMZ"] != "")].copy()
    return out


@st.cache_data(show_spinner=False)
def load_contratos(file_bytes: Optional[bytes]) -> pd.DataFrame:
    if not file_bytes:
        return pd.DataFrame()
    try:
        xl = pd.ExcelFile(BytesIO(file_bytes))
        sheet = "FIRST" if "FIRST" in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]
        col_ct = find_col(df, ["Nº CT", "N CT", "Contrato"], required=False)
        col_cliente = find_col(df, ["RAZÃO SOCIAL", "Razao Social", "Cliente"], required=False)
        col_inicio = find_col(df, ["INICIO", "Início", "Inicio"], required=False)
        col_valor = find_col(df, ["VALOR FATURAMENTO", "Valor Faturamento", "Valor"], required=False)
        col_linha = find_col(df, ["LINHA DE PRODUTO", "Linha"], required=False)
        out = pd.DataFrame()
        out["Contrato"] = df[col_ct].fillna("").astype(str).str.strip() if col_ct else ""
        out["Cliente"] = df[col_cliente].fillna("").astype(str).str.strip() if col_cliente else ""
        out["Início"] = pd.to_datetime(df[col_inicio], errors="coerce", dayfirst=True) if col_inicio else pd.NaT
        out["Valor Faturamento"] = pd.to_numeric(df[col_valor], errors="coerce").fillna(0) if col_valor else 0
        out["Linha Contrato"] = df[col_linha].fillna("").astype(str).str.strip() if col_linha else ""
        out = out[(out["Contrato"] != "") | (out["Cliente"] != "") | (out["Linha Contrato"] != "")].copy()
        return out
    except Exception:
        return pd.DataFrame()

# =========================
# MODELOS
# =========================
def build_locacao_recorrente(faturamento: pd.DataFrame) -> pd.DataFrame:
    loc = faturamento[faturamento["Tipo Operação"].eq("Locação")].copy()
    if loc.empty:
        return pd.DataFrame()

    base = loc.groupby(["Produto_Base", "Cliente", "Mês"], as_index=False).agg(
        Qtd_Mes=("Quantidade", "sum"),
        Receita_Mes=("Valor", "sum"),
        Produto_Exemplo=("Produto", "last"),
        Descrição=("Descrição", "last"),
        Linha_Faturamento=("Linha_Faturamento", "last"),
    )
    resumo = base.groupby(["Produto_Base", "Cliente"], as_index=False).agg(
        Meses_Faturados=("Mês", "nunique"),
        Primeiro_Mês=("Mês", "min"),
        Último_Mês=("Mês", "max"),
        Qtd_Locacao=("Qtd_Mes", "sum"),
        Receita_Locacao=("Receita_Mes", "sum"),
        Produto_Exemplo=("Produto_Exemplo", "last"),
        Descrição=("Descrição", "last"),
        Linha_Faturamento=("Linha_Faturamento", "last"),
    )
    resumo["Score_Recorrência"] = np.select(
        [resumo["Meses_Faturados"] >= 12, resumo["Meses_Faturados"] >= 6, resumo["Meses_Faturados"] >= 3, resumo["Meses_Faturados"] >= 2],
        [100, 80, 60, 40],
        default=20,
    )
    resumo["Locação_Ativa_Provável"] = np.where(resumo["Score_Recorrência"] >= 60, "Sim", "Não/Validar")
    return resumo


def build_forecast(
    estoque: pd.DataFrame,
    faturamento: pd.DataFrame,
    contratos: pd.DataFrame,
    horizonte: int = 30,
    dias_seguranca: int = 15,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    # IMPORTANTE: locação não consome estoque. Consumo de reposição usa apenas Venda/Outros.
    fat_consumo = faturamento[faturamento["Tipo Operação"].eq("Venda/Outros")].copy()
    fat_total = faturamento.copy()
    fat_loc = faturamento[faturamento["Tipo Operação"].eq("Locação")].copy()

    data_ref = fat_total["Data"].max() if not fat_total.empty else pd.Timestamp.today().normalize()
    inicio_30 = data_ref - pd.Timedelta(days=30)
    inicio_180 = data_ref - pd.Timedelta(days=180)
    inicio_90 = data_ref - pd.Timedelta(days=90)

    consumo_30 = fat_consumo[fat_consumo["Data"] > inicio_30].groupby("Produto_Base", as_index=False).agg(
        Qtd_30d=("Quantidade", "sum"), Receita_30d=("Valor", "sum")
    )
    consumo_180 = fat_consumo[fat_consumo["Data"] > inicio_180].groupby("Produto_Base", as_index=False).agg(
        Qtd_180d=("Quantidade", "sum"), Receita_180d=("Valor", "sum")
    )
    ultimo_consumo = fat_consumo.groupby("Produto_Base", as_index=False).agg(
        Última_Venda=("Data", "max"), Receita_Venda_Total=("Valor", "sum"), Qtd_Venda_Total=("Quantidade", "sum")
    )
    cad_fat = fat_total.sort_values("Data").groupby("Produto_Base", as_index=False).agg(
        Descrição_Faturamento=("Descrição", "last"),
        Grupo_Faturamento=("Grupo_Faturamento", "last"),
        Linha_Faturamento=("Linha_Faturamento", "last"),
    )
    receita_total = fat_total.groupby("Produto_Base", as_index=False).agg(
        Receita_Total=("Valor", "sum"), Qtd_Total=("Quantidade", "sum")
    )
    loc_prod = fat_loc.groupby("Produto_Base", as_index=False).agg(
        Receita_Locacao=("Valor", "sum"), Qtd_Locacao=("Quantidade", "sum"), Clientes_Locacao=("Cliente", "nunique")
    ) if not fat_loc.empty else pd.DataFrame(columns=["Produto_Base", "Receita_Locacao", "Qtd_Locacao", "Clientes_Locacao"])

    loc_rec = build_locacao_recorrente(faturamento)
    loc_ativo_prod = loc_rec.groupby("Produto_Base", as_index=False).agg(
        Clientes_Ativos_Prováveis=("Cliente", lambda s: s[loc_rec.loc[s.index, "Locação_Ativa_Provável"].eq("Sim")].nunique()),
        Score_Recorrência_Médio=("Score_Recorrência", "mean"),
        Maior_Score_Recorrência=("Score_Recorrência", "max"),
    ) if not loc_rec.empty else pd.DataFrame(columns=["Produto_Base", "Clientes_Ativos_Prováveis", "Score_Recorrência_Médio", "Maior_Score_Recorrência"])

    est_total = estoque.groupby("Produto_Base", as_index=False).agg(
        Produto_Exemplo=("Produto", "first"),
        Descrição_Estoque=("Descrição Estoque", "first"),
        Grupo=("Grupo", "first"),
        Estoque_Total=("Saldo Estoque", "sum"),
        Estoque_Disponível=("Estoque Disponível", "sum"),
        Valor_Estoque=("Valor Estoque", "sum"),
        Qtde_ARMZ=("ARMZ", "nunique"),
    )

    base = est_total.merge(cad_fat, on="Produto_Base", how="left")
    base = base.merge(consumo_30, on="Produto_Base", how="left")
    base = base.merge(consumo_180, on="Produto_Base", how="left")
    base = base.merge(ultimo_consumo, on="Produto_Base", how="left")
    base = base.merge(receita_total, on="Produto_Base", how="left")
    base = base.merge(loc_prod, on="Produto_Base", how="left")
    base = base.merge(loc_ativo_prod, on="Produto_Base", how="left")

    numeric_cols = [
        "Qtd_30d", "Receita_30d", "Qtd_180d", "Receita_180d", "Receita_Venda_Total", "Qtd_Venda_Total",
        "Receita_Total", "Qtd_Total", "Receita_Locacao", "Qtd_Locacao", "Clientes_Locacao", "Clientes_Ativos_Prováveis",
        "Score_Recorrência_Médio", "Maior_Score_Recorrência",
    ]
    for c in numeric_cols:
        if c in base.columns:
            base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)

    base["Produto"] = base["Produto_Exemplo"].fillna(base["Produto_Base"])
    base["Descrição"] = base["Descrição_Estoque"].replace("", np.nan).fillna(base["Descrição_Faturamento"]).fillna("")
    base["Grupo"] = base["Grupo"].fillna("").astype(str)  # oficial do estoque
    base["Linha"] = base["Linha_Faturamento"].fillna("")

    base["Consumo_Diário_30d"] = base["Qtd_30d"] / 30
    base["Consumo_Diário_180d"] = base["Qtd_180d"] / 180
    base["Consumo_Diário_Forecast"] = np.where(
        base["Qtd_30d"] > 0,
        (base["Consumo_Diário_30d"] * 0.70) + (base["Consumo_Diário_180d"] * 0.30),
        base["Consumo_Diário_180d"],
    )
    base["Forecast_30d"] = base["Consumo_Diário_Forecast"] * horizonte
    base["Estoque_Segurança"] = base["Consumo_Diário_Forecast"] * dias_seguranca
    base["Cobertura_Dias"] = np.where(base["Consumo_Diário_Forecast"] > 0, base["Estoque_Disponível"] / base["Consumo_Diário_Forecast"], np.inf)
    base["Cobertura_Meses"] = base["Cobertura_Dias"] / 30
    base["Necessidade_Bruta"] = base["Forecast_30d"] + base["Estoque_Segurança"] - base["Estoque_Disponível"]
    base["Comprar_Qtd"] = np.ceil(base["Necessidade_Bruta"].clip(lower=0)).astype(int)
    base["Custo_Médio"] = np.where(base["Estoque_Disponível"] > 0, base["Valor_Estoque"] / base["Estoque_Disponível"], 0)
    base["Comprar_R$"] = base["Comprar_Qtd"] * base["Custo_Médio"]
    base["Excesso_Estoque_Qtd"] = (base["Estoque_Disponível"] - (base["Forecast_30d"] + base["Estoque_Segurança"])).clip(lower=0)
    base["Excesso_Estoque_R$"] = base["Excesso_Estoque_Qtd"] * base["Custo_Médio"]
    base["Dias_Sem_Venda"] = np.where(base["Última_Venda"].notna(), (data_ref - base["Última_Venda"]).dt.days, np.nan)

    def status(row):
        if row["Consumo_Diário_Forecast"] <= 0:
            return "⚫ Sem Giro"
        if row["Cobertura_Dias"] <= 15:
            return "🔴 Crítico"
        if row["Cobertura_Dias"] <= 30:
            return "🟠 Atenção"
        if row["Cobertura_Dias"] <= 60:
            return "🟡 Monitorar"
        return "🟢 Saudável"

    base["Status"] = base.apply(status, axis=1)
    base["Ação"] = np.select(
        [base["Comprar_Qtd"] > 0, base["Receita_Locacao"] > 0, base["Status"].eq("⚫ Sem Giro") & (base["Valor_Estoque"] > 0)],
        ["Comprar", "Analisar parque de locação", "Avaliar capital parado"],
        default="Manter",
    )
    base["Índice_Expansão_Locação"] = (
        np.where(base["Maior_Score_Recorrência"] >= 80, 40, np.where(base["Maior_Score_Recorrência"] >= 60, 25, 0))
        + np.where(base["Receita_Locacao"] > base["Receita_Locacao"].quantile(0.80), 30, 0)
        + np.where(base["Estoque_Disponível"] <= base["Clientes_Ativos_Prováveis"], 30, 0)
    ).clip(0, 100)
    base["Score_Oportunidade"] = (
        np.where(base["Status"].eq("🔴 Crítico"), 45, 0)
        + np.where(base["Status"].eq("🟠 Atenção"), 25, 0)
        + np.where(base["Comprar_Qtd"] > 0, 25, 0)
        + np.where(base["Excesso_Estoque_R$"] > base["Excesso_Estoque_R$"].quantile(0.80), 15, 0)
        + np.where(base["Índice_Expansão_Locação"] >= 70, 15, 0)
    ).clip(0, 100)

    # ABCs
    def abc_class(df: pd.DataFrame, metric: str, label: str) -> pd.DataFrame:
        x = df[["Produto_Base", metric]].copy().sort_values(metric, ascending=False)
        total = x[metric].sum()
        x[f"%_{label}"] = np.where(total > 0, x[metric] / total, 0)
        x[f"%_Acum_{label}"] = x[f"%_{label}"].cumsum()
        x[f"ABC_{label}"] = np.select([x[f"%_Acum_{label}"] <= 0.80, x[f"%_Acum_{label}"] <= 0.95], ["A", "B"], default="C")
        return x[["Produto_Base", f"ABC_{label}"]]

    base = base.merge(abc_class(base, "Receita_Total", "Receita"), on="Produto_Base", how="left")
    base = base.merge(abc_class(base, "Qtd_Venda_Total", "Giro"), on="Produto_Base", how="left")
    base = base.merge(abc_class(base, "Valor_Estoque", "Estoque"), on="Produto_Base", how="left")

    # ARMZ
    armz = estoque.groupby(["Produto_Base", "ARMZ"], as_index=False).agg(
        Estoque_ARMZ=("Saldo Estoque", "sum"),
        Disponível_ARMZ=("Estoque Disponível", "sum"),
        Valor_ARMZ=("Valor Estoque", "sum"),
    )
    armz = armz.merge(base[["Produto_Base", "Produto", "Descrição", "Grupo", "Linha", "Consumo_Diário_Forecast", "Forecast_30d", "Estoque_Segurança", "Status"]], on="Produto_Base", how="left")
    armz["Cobertura_ARMZ_Dias"] = np.where(armz["Consumo_Diário_Forecast"].fillna(0) > 0, armz["Disponível_ARMZ"] / armz["Consumo_Diário_Forecast"], np.inf)

    # Transferência: somente quando estoque consolidado cobre, mas um ARMZ está baixo e outro tem sobra.
    transfer_rows = []
    base_no_buy = base[base["Comprar_Qtd"].eq(0) & (base["Consumo_Diário_Forecast"] > 0)]
    for _, prod in base_no_buy.iterrows():
        p = prod["Produto_Base"]
        part = armz[armz["Produto_Base"].eq(p)].copy()
        if len(part) < 2:
            continue
        demanda_ref_armz = max(1, math.ceil(prod["Forecast_30d"] / max(part["ARMZ"].nunique(), 1)))
        deficit = part[part["Disponível_ARMZ"] < demanda_ref_armz].sort_values("Disponível_ARMZ")
        sobra = part[part["Disponível_ARMZ"] > demanda_ref_armz].sort_values("Disponível_ARMZ", ascending=False)
        if deficit.empty or sobra.empty:
            continue
        origem = sobra.iloc[0]
        destino = deficit.iloc[0]
        qtd = int(min(origem["Disponível_ARMZ"] - demanda_ref_armz, demanda_ref_armz - destino["Disponível_ARMZ"]))
        if qtd > 0:
            transfer_rows.append({
                "Produto": prod["Produto"],
                "Produto_Base": p,
                "Descrição": prod["Descrição"],
                "Grupo": prod["Grupo"],
                "Origem_ARMZ": origem["ARMZ"],
                "Destino_ARMZ": destino["ARMZ"],
                "Qtd_Sugerida": qtd,
                "Motivo": "Redistribuir estoque antes de comprar",
            })
    transfer = pd.DataFrame(transfer_rows)

    # Parque de locação por cliente/produto com dados de estoque agregados.
    parque = loc_rec.merge(base[["Produto_Base", "Produto", "Grupo", "Estoque_Disponível", "Valor_Estoque", "Índice_Expansão_Locação"]], on="Produto_Base", how="left") if not loc_rec.empty else pd.DataFrame()
    if not parque.empty:
        parque["Produto"] = parque["Produto"].fillna(parque["Produto_Exemplo"])
        parque = parque.sort_values(["Score_Recorrência", "Receita_Locacao"], ascending=False)

    # Contratos por linha: informativo, não desconta estoque.
    contratos_linha = pd.DataFrame()
    if contratos is not None and not contratos.empty:
        contratos_linha = contratos.groupby("Linha Contrato", as_index=False).agg(
            Contratos=("Contrato", "nunique"),
            Clientes=("Cliente", "nunique"),
            Valor_Faturamento=("Valor Faturamento", "sum"),
        ).sort_values("Valor_Faturamento", ascending=False)

    return base, armz, transfer, parque, contratos_linha, fat_consumo

# =========================
# INTERFACE
# =========================
st.markdown(
    f"""
    <div class="first-header">
        <h1>{APP_NAME}</h1>
        <p>{APP_SUBTITLE}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### 📁 Bases")
    st.caption("O app usa os arquivos fixos da pasta /dados. O upload abaixo substitui apenas nesta sessão.")
    fat_file = st.file_uploader("Substituir Faturamento", type=["xlsx"], help="Usar a guia Base")
    est_file = st.file_uploader("Substituir Estoque MATR260", type=["xlsx"], help="Relatório com coluna ARMZ")
    contratos_file = st.file_uploader("Contratos ativos (opcional)", type=["xlsx"], help="Se não enviar, o app tenta usar /dados/contratos.xlsx")

    st.markdown("### ⚙️ Parâmetros")
    horizonte = st.number_input("Horizonte do forecast (dias)", min_value=7, max_value=180, value=30, step=1)
    dias_seguranca = st.number_input("Estoque de segurança (dias)", min_value=0, max_value=90, value=15, step=1)
    st.markdown("<div class='small-note'>Reposição usa apenas Venda/Outros. Locação vai para Parque de Locação.</div>", unsafe_allow_html=True)

fat_bytes, fat_source = read_file_bytes(fat_file, DEFAULT_FATURAMENTO)
est_bytes, est_source = read_file_bytes(est_file, DEFAULT_ESTOQUE)
contratos_bytes, contratos_source = read_file_bytes(contratos_file, DEFAULT_CONTRATOS)

with st.expander("📌 Origem das bases", expanded=False):
    st.write(f"**Faturamento:** {fat_source}")
    st.write(f"**Estoque:** {est_source}")
    st.write(f"**Contratos:** {contratos_source}")

if not fat_bytes or not est_bytes:
    st.info("Adicione os arquivos fixos em `/dados` ou envie o Faturamento e o MATR260 pela barra lateral.")
    st.stop()

try:
    faturamento = load_faturamento(fat_bytes, "Base")
    estoque = load_estoque(est_bytes)
    contratos = load_contratos(contratos_bytes)
    forecast, armz, transferencias, parque_locacao, contratos_linha, fat_consumo = build_forecast(
        estoque, faturamento, contratos, horizonte=int(horizonte), dias_seguranca=int(dias_seguranca)
    )
except Exception as e:
    st.error("Não consegui processar os arquivos. Verifique se os relatórios estão no layout esperado.")
    st.exception(e)
    st.stop()

# Filtros globais
with st.expander("🔎 Filtros", expanded=True):
    c1, c2, c3, c4, c5 = st.columns(5)
    grupos = sorted([g for g in forecast["Grupo"].dropna().unique() if str(g).strip()])
    linhas = sorted([l for l in forecast["Linha"].dropna().unique() if str(l).strip()])
    statuses = ["🔴 Crítico", "🟠 Atenção", "🟡 Monitorar", "🟢 Saudável", "⚫ Sem Giro"]
    armzs = sorted([a for a in armz["ARMZ"].dropna().unique() if str(a).strip()])
    with c1:
        filtro_grupo = st.multiselect("Grupo oficial do estoque", grupos)
    with c2:
        filtro_linha = st.multiselect("Linha do faturamento", linhas)
    with c3:
        filtro_status = st.multiselect("Status", statuses)
    with c4:
        filtro_armz = st.multiselect("ARMZ", armzs)
    with c5:
        busca = st.text_input("Produto / descrição")

view = forecast.copy()
if filtro_grupo:
    view = view[view["Grupo"].isin(filtro_grupo)]
if filtro_linha:
    view = view[view["Linha"].isin(filtro_linha)]
if filtro_status:
    view = view[view["Status"].isin(filtro_status)]
if busca:
    b = norm_text(busca)
    view = view[view["Produto"].map(norm_text).str.contains(b, na=False) | view["Descrição"].map(norm_text).str.contains(b, na=False) | view["Produto_Base"].map(norm_text).str.contains(b, na=False)]
if filtro_armz:
    produtos_armz = armz[armz["ARMZ"].isin(filtro_armz)]["Produto_Base"].unique()
    view = view[view["Produto_Base"].isin(produtos_armz)]

armz_view = armz[armz["Produto_Base"].isin(view["Produto_Base"])]
if filtro_armz:
    armz_view = armz_view[armz_view["ARMZ"].isin(filtro_armz)]

# KPIs
criticos = int(view["Status"].eq("🔴 Crítico").sum())
sem_giro = int(view["Status"].eq("⚫ Sem Giro").sum())
compras = view[view["Comprar_Qtd"] > 0]
capital_parado = view[(view["Status"].eq("⚫ Sem Giro")) & (view["Valor_Estoque"] > 0)]["Valor_Estoque"].sum()
valor_estoque = view["Valor_Estoque"].sum()
transf_count = len(transferencias[transferencias["Produto_Base"].isin(view["Produto_Base"])]) if not transferencias.empty else 0
cobertura_media = view.replace([np.inf, -np.inf], np.nan)["Cobertura_Dias"].mean()
receita_loc = view["Receita_Locacao"].sum()
excesso = view["Excesso_Estoque_R$"].sum()

k1, k2, k3, k4 = st.columns(4)
with k1:
    kpi_card("Valor Total em Estoque", brl(valor_estoque), "Base MATR260")
with k2:
    kpi_card("Produtos Críticos", fmt_num(criticos), "Venda/Outros com cobertura até 15 dias")
with k3:
    kpi_card("Compras Recomendadas", brl(compras["Comprar_R$"].sum()), f"{len(compras)} produtos")
with k4:
    kpi_card("Capital Parado", brl(capital_parado), f"{sem_giro} produtos sem venda")

k5, k6, k7, k8 = st.columns(4)
with k5:
    kpi_card("Transferências", fmt_num(transf_count), "Sugestões entre ARMZ")
with k6:
    kpi_card("Cobertura Média", f"{fmt_num(cobertura_media, 1)} dias" if not np.isnan(cobertura_media) else "-", "Somente consumo de venda/outros")
with k7:
    kpi_card("Receita Locação", brl(receita_loc), "Não gera compra automática")
with k8:
    kpi_card("Excesso Potencial", brl(excesso), "Acima do forecast + segurança")

aba1, aba2, aba3, aba4, aba5, aba6, aba7, aba8, aba9 = st.tabs([
    "🏠 Radar Executivo",
    "📈 Forecast Estoque",
    "🛒 Compras",
    "🔄 Transferências",
    "📦 Capital Parado",
    "🎯 Curva ABC",
    "🏥 Parque Locação",
    "🏢 ARMZ",
    "📄 Contratos",
])

cols_forecast = [
    "Produto", "Produto_Base", "Descrição", "Grupo", "Linha", "Estoque_Disponível", "Valor_Estoque", "Qtd_30d", "Qtd_180d",
    "Forecast_30d", "Estoque_Segurança", "Cobertura_Dias", "Cobertura_Meses", "Excesso_Estoque_Qtd", "Excesso_Estoque_R$",
    "Comprar_Qtd", "Comprar_R$", "Status", "Ação", "Score_Oportunidade", "ABC_Receita", "ABC_Giro", "ABC_Estoque",
]

num_config = {
    "Valor_Estoque": st.column_config.NumberColumn("Valor Estoque", format="R$ %.2f"),
    "Forecast_30d": st.column_config.NumberColumn("Forecast 30d", format="%.2f"),
    "Estoque_Segurança": st.column_config.NumberColumn("Segurança", format="%.2f"),
    "Cobertura_Dias": st.column_config.NumberColumn("Cobertura Dias", format="%.1f"),
    "Cobertura_Meses": st.column_config.NumberColumn("Cobertura Meses", format="%.1f"),
    "Comprar_R$": st.column_config.NumberColumn("Comprar R$", format="R$ %.2f"),
    "Excesso_Estoque_R$": st.column_config.NumberColumn("Excesso R$", format="R$ %.2f"),
    "Score_Oportunidade": st.column_config.ProgressColumn("Score", min_value=0, max_value=100),
}

with aba1:
    st.markdown("<div class='section-title'>Radar Executivo</div>", unsafe_allow_html=True)
    st.markdown("""
    <div class="logic-box">
    <b>Lógica atual:</b> locação não entra como consumo de estoque. O forecast de compra considera apenas venda/outros.
    A locação aparece no módulo Parque Locação para medir recorrência, receita e possível expansão do parque.
    </div>
    """, unsafe_allow_html=True)
    radar = view.sort_values(["Score_Oportunidade", "Comprar_R$", "Excesso_Estoque_R$"], ascending=False)[cols_forecast].copy()
    st.dataframe(radar, use_container_width=True, hide_index=True, column_config=num_config)
    c1, c2 = st.columns(2)
    with c1:
        status_df = view.groupby("Status", as_index=False).agg(Produtos=("Produto", "count"))
        st.plotly_chart(px.bar(status_df, x="Status", y="Produtos", title="Produtos por Status"), use_container_width=True)
    with c2:
        top_excesso = view.sort_values("Excesso_Estoque_R$", ascending=False).head(10)
        st.plotly_chart(px.bar(top_excesso, x="Produto", y="Excesso_Estoque_R$", title="Top 10 Excesso Potencial"), use_container_width=True)

with aba2:
    st.markdown("<div class='section-title'>Forecast de Estoque</div>", unsafe_allow_html=True)
    st.caption("Forecast = 70% últimos 30 dias + 30% últimos 180 dias, usando somente Venda/Outros. Locação recorrente foi excluída do consumo.")
    st.dataframe(view.sort_values("Cobertura_Dias")[cols_forecast], use_container_width=True, hide_index=True, column_config=num_config)

with aba3:
    st.markdown("<div class='section-title'>Compras Recomendadas</div>", unsafe_allow_html=True)
    compra_view = compras.sort_values(["Comprar_R$", "Cobertura_Dias"], ascending=[False, True])[cols_forecast]
    st.dataframe(compra_view, use_container_width=True, hide_index=True, column_config=num_config)

with aba4:
    st.markdown("<div class='section-title'>Transferências entre ARMZ</div>", unsafe_allow_html=True)
    if transferencias.empty:
        st.success("Nenhuma transferência recomendada com os critérios atuais.")
    else:
        tv = transferencias[transferencias["Produto_Base"].isin(view["Produto_Base"])].copy()
        st.dataframe(tv, use_container_width=True, hide_index=True)

with aba5:
    st.markdown("<div class='section-title'>Capital Parado e Sem Venda</div>", unsafe_allow_html=True)
    parado = view[(view["Status"].eq("⚫ Sem Giro")) | (view["Dias_Sem_Venda"] >= 90)].copy()
    parado["Faixa Sem Venda"] = pd.cut(
        parado["Dias_Sem_Venda"].fillna(9999),
        bins=[-1, 90, 180, 365, 99999],
        labels=["Até 90 dias", "90 a 180 dias", "180 a 365 dias", "> 365 dias / sem histórico"],
    )
    st.dataframe(parado.sort_values("Valor_Estoque", ascending=False)[[
        "Produto", "Produto_Base", "Descrição", "Grupo", "Linha", "Estoque_Disponível", "Valor_Estoque", "Última_Venda", "Dias_Sem_Venda", "Faixa Sem Venda", "Status", "Excesso_Estoque_R$"
    ]], use_container_width=True, hide_index=True)

with aba6:
    st.markdown("<div class='section-title'>Curva ABC</div>", unsafe_allow_html=True)
    abc_view = view.sort_values("Receita_Total", ascending=False)[[
        "Produto", "Descrição", "Grupo", "Receita_Total", "Qtd_Venda_Total", "Valor_Estoque", "ABC_Receita", "ABC_Giro", "ABC_Estoque", "Status"
    ]]
    st.dataframe(abc_view, use_container_width=True, hide_index=True)
    abc_chart = abc_view.groupby("ABC_Estoque", as_index=False).agg(Valor_Estoque=("Valor_Estoque", "sum"), Produtos=("Produto", "count"))
    st.plotly_chart(px.bar(abc_chart, x="ABC_Estoque", y="Valor_Estoque", text="Produtos", title="ABC por Valor em Estoque"), use_container_width=True)

with aba7:
    st.markdown("<div class='section-title'>Parque de Locação por Recorrência</div>", unsafe_allow_html=True)
    if parque_locacao.empty:
        st.info("Não há itens de locação identificados nos filtros atuais.")
    else:
        pl = parque_locacao[parque_locacao["Produto_Base"].isin(view["Produto_Base"])].copy()
        st.caption("Recorrência = mesmo cliente + mesmo produto em meses diferentes. Score >= 60 indica locação ativa provável.")
        st.dataframe(pl[[
            "Produto", "Produto_Base", "Descrição", "Cliente", "Meses_Faturados", "Primeiro_Mês", "Último_Mês", "Qtd_Locacao", "Receita_Locacao",
            "Score_Recorrência", "Locação_Ativa_Provável", "Estoque_Disponível", "Índice_Expansão_Locação", "Grupo"
        ]], use_container_width=True, hide_index=True)
        resumo_parque = pl.groupby(["Produto_Base", "Produto", "Descrição", "Grupo"], as_index=False).agg(
            Clientes=("Cliente", "nunique"),
            Clientes_Ativos=("Locação_Ativa_Provável", lambda s: (s == "Sim").sum()),
            Receita_Locacao=("Receita_Locacao", "sum"),
            Score_Médio=("Score_Recorrência", "mean"),
            Estoque_Disponível=("Estoque_Disponível", "first"),
            Índice_Expansão=("Índice_Expansão_Locação", "max"),
        ).sort_values(["Índice_Expansão", "Receita_Locacao"], ascending=False)
        st.markdown("#### Ranking de expansão do parque")
        st.dataframe(resumo_parque, use_container_width=True, hide_index=True)
        st.plotly_chart(px.bar(resumo_parque.head(15), x="Produto", y="Receita_Locacao", title="Top 15 Receita de Locação Recorrente"), use_container_width=True)

with aba8:
    st.markdown("<div class='section-title'>Análise por ARMZ</div>", unsafe_allow_html=True)
    resumo_armz = armz_view.groupby("ARMZ", as_index=False).agg(
        Valor_Estoque=("Valor_ARMZ", "sum"), Estoque_Disponivel=("Disponível_ARMZ", "sum"), Produtos=("Produto_Base", "nunique")
    )
    st.dataframe(resumo_armz.sort_values("Valor_Estoque", ascending=False), use_container_width=True, hide_index=True)
    st.plotly_chart(px.bar(resumo_armz.sort_values("Valor_Estoque", ascending=False), x="ARMZ", y="Valor_Estoque", title="Valor em Estoque por ARMZ"), use_container_width=True)
    st.dataframe(armz_view.sort_values(["Produto", "ARMZ"]), use_container_width=True, hide_index=True)

with aba9:
    st.markdown("<div class='section-title'>Contratos Ativos</div>", unsafe_allow_html=True)
    if contratos.empty:
        st.info("Base de contratos não carregada. Ela é opcional e não bloqueia o forecast.")
    else:
        st.caption("Contratos entram como informação gerencial por linha. Não descontam estoque porque a base não possui produto + quantidade.")
        st.dataframe(contratos, use_container_width=True, hide_index=True)
        if not contratos_linha.empty:
            st.markdown("#### Pressão contratual por linha")
            st.dataframe(contratos_linha, use_container_width=True, hide_index=True)
            st.plotly_chart(px.bar(contratos_linha.head(15), x="Linha Contrato", y="Valor_Faturamento", text="Contratos", title="Contratos por Linha"), use_container_width=True)

st.divider()
excel_bytes = to_excel_download({
    "Radar Executivo": view[cols_forecast].sort_values("Score_Oportunidade", ascending=False),
    "Compras": compras[cols_forecast].sort_values("Comprar_R$", ascending=False),
    "Transferencias": transferencias if not transferencias.empty else pd.DataFrame(columns=["Produto", "Descrição", "Origem_ARMZ", "Destino_ARMZ", "Qtd_Sugerida", "Motivo"]),
    "Capital Parado": view[(view["Status"].eq("⚫ Sem Giro")) | (view["Dias_Sem_Venda"] >= 90)],
    "ARMZ": armz_view,
    "Parque Locacao": parque_locacao[parque_locacao["Produto_Base"].isin(view["Produto_Base"])] if not parque_locacao.empty else pd.DataFrame(),
    "Contratos": contratos if not contratos.empty else pd.DataFrame(),
})
st.download_button(
    "📥 Baixar análise em Excel",
    data=excel_bytes,
    file_name=f"first_intelligence_forecast_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(f"Última atualização da análise: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
