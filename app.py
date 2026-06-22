# -*- coding: utf-8 -*-
"""
FIRST Intelligence | Forecast Estratégico de Estoque, Compras e Locação
App Streamlit para cruzar MATR260 (estoque) + Relatório de Faturamento (guia Base).
"""

from __future__ import annotations

import math
import re
from datetime import datetime
from io import BytesIO
from typing import Optional

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

APP_NAME = "FIRST Intelligence"
APP_SUBTITLE = "Forecast Estratégico de Estoque, Compras e Locação"

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
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================
# FUNÇÕES UTILITÁRIAS
# =========================
def norm_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_code(value: object) -> str:
    """Normaliza código preservando sufixos como _RV, _AT, _TC."""
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = text.replace(" ", "")
    if text.endswith(".0") and re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def normalize_col(col: object) -> str:
    text = norm_text(col)
    # remove acentos de forma simples
    repl = str.maketrans("ÁÀÂÃÉÊÍÓÔÕÚÇ", "AAAAEEIOOOUC")
    text = text.translate(repl)
    text = re.sub(r"[^A-Z0-9]+", "_", text).strip("_")
    return text


def find_col(df: pd.DataFrame, candidates: list[str], required: bool = True) -> Optional[str]:
    normalized = {normalize_col(c): c for c in df.columns}
    for cand in candidates:
        key = normalize_col(cand)
        if key in normalized:
            return normalized[key]
    # tentativa por contém
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
                width = min(max(len(str(col_name)) + 2, 12), 35)
                ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width
    return output.getvalue()


# =========================
# LEITURA E TRATAMENTO
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
    col_grupo = find_col(df, ["GRUPO", "Grupo"], required=False)
    col_linha = find_col(df, ["LINHA DE PRODUTO", "Linha de Produto"], required=False)
    col_nova = find_col(df, ["NOVA", "Classificacao Nova"], required=False)
    col_class = find_col(df, ["CLASSIFICAÇÃO", "Classificacao"], required=False)
    col_finalidade = find_col(df, ["FINALIDADE", "Finalidade"], required=False)
    col_categoria = find_col(df, ["CATEGORIA", "Categoria"], required=False)

    out = pd.DataFrame()
    out["Produto"] = df[col_produto].map(normalize_code)
    out["Descrição"] = df[col_desc].fillna("").astype(str).str.strip() if col_desc else ""
    out["Data"] = pd.to_datetime(df[col_data], errors="coerce", dayfirst=True)
    out["Quantidade"] = pd.to_numeric(df[col_qtd], errors="coerce").fillna(0)
    out["Valor"] = pd.to_numeric(df[col_valor], errors="coerce").fillna(0) if col_valor else 0
    out["Grupo"] = df[col_grupo].fillna("").astype(str).str.strip() if col_grupo else ""
    out["Linha"] = df[col_linha].fillna("").astype(str).str.strip() if col_linha else ""
    out["Nova"] = df[col_nova].fillna("").astype(str).str.strip() if col_nova else ""
    out["Classificação"] = df[col_class].fillna("").astype(str).str.strip() if col_class else ""
    out["Finalidade"] = df[col_finalidade].fillna("").astype(str).str.strip() if col_finalidade else ""
    out["Categoria"] = df[col_categoria].fillna("").astype(str).str.strip() if col_categoria else ""

    texto_loc = (
        out["Nova"].map(norm_text) + " " +
        out["Classificação"].map(norm_text) + " " +
        out["Finalidade"].map(norm_text) + " " +
        out["Categoria"].map(norm_text)
    )
    out["Tipo Operação"] = np.where(texto_loc.str.contains("LOCACAO|LOCAÇÃO", regex=True), "Locação", "Venda/Outros")
    out = out[(out["Produto"] != "") & out["Data"].notna()].copy()
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

    # O MATR260 pode deixar CODIGO/DESCRICAO em branco nas linhas seguintes do mesmo produto por ARMZ.
    for col in [col_codigo, col_desc, col_grupo]:
        if col:
            df[col] = df[col].ffill()

    out = pd.DataFrame()
    out["Produto"] = df[col_codigo].map(normalize_code)
    out["Descrição Estoque"] = df[col_desc].fillna("").astype(str).str.strip() if col_desc else ""
    out["Grupo Estoque"] = df[col_grupo].fillna("").astype(str).str.strip() if col_grupo else ""
    out["ARMZ"] = df[col_armz].fillna("").astype(str).str.strip()
    out["Saldo Estoque"] = pd.to_numeric(df[col_saldo], errors="coerce").fillna(0)
    out["Empenho"] = pd.to_numeric(df[col_empenho], errors="coerce").fillna(0) if col_empenho else 0
    if col_disp:
        out["Estoque Disponível"] = pd.to_numeric(df[col_disp], errors="coerce").fillna(0)
    else:
        out["Estoque Disponível"] = out["Saldo Estoque"] - out["Empenho"]
    out["Valor Estoque"] = pd.to_numeric(df[col_valor], errors="coerce").fillna(0) if col_valor else 0
    out = out[(out["Produto"] != "") & (out["ARMZ"] != "")].copy()
    return out


def build_forecast(
    estoque: pd.DataFrame,
    faturamento: pd.DataFrame,
    horizonte: int = 30,
    dias_seguranca: int = 15,
    usar_tipo: str = "Todos",
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    fat = faturamento.copy()
    if usar_tipo != "Todos":
        fat = fat[fat["Tipo Operação"].eq(usar_tipo)].copy()

    data_ref = fat["Data"].max() if not fat.empty else pd.Timestamp.today().normalize()
    inicio_30 = data_ref - pd.Timedelta(days=30)
    inicio_180 = data_ref - pd.Timedelta(days=180)
    inicio_90 = data_ref - pd.Timedelta(days=90)

    fat_30 = fat[fat["Data"] > inicio_30]
    fat_180 = fat[fat["Data"] > inicio_180]
    fat_90 = fat[fat["Data"] > inicio_90]

    consumo_30 = fat_30.groupby("Produto", as_index=False).agg(Qtd_30d=("Quantidade", "sum"), Receita_30d=("Valor", "sum"))
    consumo_180 = fat_180.groupby("Produto", as_index=False).agg(Qtd_180d=("Quantidade", "sum"), Receita_180d=("Valor", "sum"))
    ultimo_mov = fat.groupby("Produto", as_index=False).agg(
        Última_Movimentação=("Data", "max"),
        Receita_Total=("Valor", "sum"),
        Qtd_Total=("Quantidade", "sum"),
    )
    cad = fat.sort_values("Data").groupby("Produto", as_index=False).agg(
        Descrição_Faturamento=("Descrição", "last"),
        Grupo=("Grupo", "last"),
        Linha=("Linha", "last"),
    )
    loc = fat[fat["Tipo Operação"].eq("Locação")].groupby("Produto", as_index=False).agg(
        Qtd_Locacao=("Quantidade", "sum"),
        Receita_Locacao=("Valor", "sum"),
    )

    est_total = estoque.groupby("Produto", as_index=False).agg(
        Descrição_Estoque=("Descrição Estoque", "first"),
        Grupo_Estoque=("Grupo Estoque", "first"),
        Estoque_Total=("Saldo Estoque", "sum"),
        Estoque_Disponível=("Estoque Disponível", "sum"),
        Valor_Estoque=("Valor Estoque", "sum"),
        Qtde_ARMZ=("ARMZ", "nunique"),
    )

    base = est_total.merge(cad, on="Produto", how="left")
    base = base.merge(consumo_30, on="Produto", how="left")
    base = base.merge(consumo_180, on="Produto", how="left")
    base = base.merge(ultimo_mov, on="Produto", how="left")
    base = base.merge(loc, on="Produto", how="left")

    for c in ["Qtd_30d", "Receita_30d", "Qtd_180d", "Receita_180d", "Receita_Total", "Qtd_Total", "Qtd_Locacao", "Receita_Locacao"]:
        base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)

    base["Descrição"] = base["Descrição_Faturamento"].fillna(base["Descrição_Estoque"]).fillna("")
    base["Grupo"] = base["Grupo"].replace("", np.nan).fillna(base["Grupo_Estoque"]).fillna("")
    base["Linha"] = base["Linha"].fillna("")

    base["Consumo_Diário_30d"] = base["Qtd_30d"] / 30
    base["Consumo_Diário_180d"] = base["Qtd_180d"] / 180
    # Inteligência: 70% recente + 30% histórico. Se não teve 30d, usa histórico.
    base["Consumo_Diário_Forecast"] = np.where(
        base["Qtd_30d"] > 0,
        (base["Consumo_Diário_30d"] * 0.70) + (base["Consumo_Diário_180d"] * 0.30),
        base["Consumo_Diário_180d"],
    )
    base["Forecast_30d"] = base["Consumo_Diário_Forecast"] * horizonte
    base["Estoque_Segurança"] = base["Consumo_Diário_Forecast"] * dias_seguranca
    base["Cobertura_Dias"] = np.where(
        base["Consumo_Diário_Forecast"] > 0,
        base["Estoque_Disponível"] / base["Consumo_Diário_Forecast"],
        np.inf,
    )
    base["Necessidade_Bruta"] = base["Forecast_30d"] + base["Estoque_Segurança"] - base["Estoque_Disponível"]
    base["Comprar_Qtd"] = np.ceil(base["Necessidade_Bruta"].clip(lower=0)).astype(int)
    base["Custo_Médio"] = np.where(base["Estoque_Disponível"] > 0, base["Valor_Estoque"] / base["Estoque_Disponível"], 0)
    base["Comprar_R$"] = base["Comprar_Qtd"] * base["Custo_Médio"]
    base["Dias_Sem_Giro"] = np.where(
        base["Última_Movimentação"].notna(),
        (data_ref - base["Última_Movimentação"]).dt.days,
        np.nan,
    )

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
        [base["Comprar_Qtd"] > 0, base["Status"].eq("⚫ Sem Giro") & (base["Valor_Estoque"] > 0)],
        ["Comprar", "Avaliar capital parado"],
        default="Manter",
    )
    base["Score_Oportunidade"] = (
        np.where(base["Status"].eq("🔴 Crítico"), 45, 0)
        + np.where(base["Status"].eq("🟠 Atenção"), 25, 0)
        + np.where(base["Comprar_Qtd"] > 0, 25, 0)
        + np.where(base["Receita_Locacao"] > 0, 15, 0)
        + np.where(base["Receita_Total"] > base["Receita_Total"].quantile(0.80), 15, 0)
    ).clip(0, 100)

    # Curva ABC por receita
    abc = base[["Produto", "Descrição", "Receita_Total", "Qtd_Total", "Valor_Estoque"]].copy()
    abc = abc.sort_values("Receita_Total", ascending=False)
    total_receita = abc["Receita_Total"].sum()
    abc["% Receita"] = np.where(total_receita > 0, abc["Receita_Total"] / total_receita, 0)
    abc["% Acumulado"] = abc["% Receita"].cumsum()
    abc["Classe ABC"] = np.select(
        [abc["% Acumulado"] <= 0.80, abc["% Acumulado"] <= 0.95],
        ["A", "B"],
        default="C",
    )
    base = base.merge(abc[["Produto", "Classe ABC"]], on="Produto", how="left")

    # Estoque por ARMZ com distribuição referencial da demanda entre armazéns com saldo.
    armz = estoque.groupby(["Produto", "ARMZ"], as_index=False).agg(
        Estoque_ARMZ=("Saldo Estoque", "sum"),
        Disponível_ARMZ=("Estoque Disponível", "sum"),
        Valor_ARMZ=("Valor Estoque", "sum"),
    )
    armz = armz.merge(base[["Produto", "Descrição", "Grupo", "Linha", "Consumo_Diário_Forecast", "Forecast_30d", "Estoque_Segurança", "Status"]], on="Produto", how="left")
    armz["Cobertura_ARMZ_Dias"] = np.where(
        armz["Consumo_Diário_Forecast"].fillna(0) > 0,
        armz["Disponível_ARMZ"] / armz["Consumo_Diário_Forecast"],
        np.inf,
    )

    # Sugestão de transferência: quando consolidado não precisa comprar, mas existe ARMZ zerado/baixo e outro com sobra.
    transfer_rows = []
    base_no_buy = base[base["Comprar_Qtd"].eq(0) & (base["Consumo_Diário_Forecast"] > 0)]
    for _, prod in base_no_buy.iterrows():
        p = prod["Produto"]
        part = armz[armz["Produto"].eq(p)].copy()
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
                "Produto": p,
                "Descrição": prod["Descrição"],
                "Origem_ARMZ": origem["ARMZ"],
                "Destino_ARMZ": destino["ARMZ"],
                "Qtd_Sugerida": qtd,
                "Motivo": "Redistribuir estoque antes de comprar",
            })
    transfer = pd.DataFrame(transfer_rows)

    # Locação
    locacao = base[base["Receita_Locacao"] > 0].copy()
    locacao["Índice_Pressão_Locação"] = np.where(
        locacao["Estoque_Disponível"] > 0,
        locacao["Qtd_Locacao"] / locacao["Estoque_Disponível"],
        locacao["Qtd_Locacao"],
    )
    locacao = locacao.sort_values(["Índice_Pressão_Locação", "Receita_Locacao"], ascending=False)

    return base, armz, transfer, locacao


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
    fat_file = st.file_uploader("Relatório de Faturamento 2026", type=["xlsx"], help="Usar a guia Base")
    est_file = st.file_uploader("Relatório de Estoque MATR260", type=["xlsx"], help="Relatório com coluna ARMZ")

    st.markdown("### ⚙️ Parâmetros")
    horizonte = st.number_input("Horizonte do forecast (dias)", min_value=7, max_value=180, value=30, step=1)
    dias_seguranca = st.number_input("Estoque de segurança (dias)", min_value=0, max_value=90, value=15, step=1)
    tipo_operacao = st.selectbox("Base de consumo", ["Todos", "Locação", "Venda/Outros"], index=0)
    st.markdown("<div class='small-note'>Forecast padrão: 70% últimos 30 dias + 30% últimos 180 dias.</div>", unsafe_allow_html=True)

if not fat_file or not est_file:
    st.info("Envie o **Relatório de Faturamento** e o **MATR260 de Estoque** para iniciar a análise.")
    st.stop()

try:
    fat_bytes = fat_file.getvalue()
    est_bytes = est_file.getvalue()
    faturamento = load_faturamento(fat_bytes, "Base")
    estoque = load_estoque(est_bytes)
    forecast, armz, transferencias, locacao = build_forecast(
        estoque, faturamento, horizonte=int(horizonte), dias_seguranca=int(dias_seguranca), usar_tipo=tipo_operacao
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
        filtro_grupo = st.multiselect("Grupo", grupos)
    with c2:
        filtro_linha = st.multiselect("Linha", linhas)
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
    view = view[view["Produto"].map(norm_text).str.contains(b, na=False) | view["Descrição"].map(norm_text).str.contains(b, na=False)]
if filtro_armz:
    produtos_armz = armz[armz["ARMZ"].isin(filtro_armz)]["Produto"].unique()
    view = view[view["Produto"].isin(produtos_armz)]

armz_view = armz[armz["Produto"].isin(view["Produto"])]
if filtro_armz:
    armz_view = armz_view[armz_view["ARMZ"].isin(filtro_armz)]

# KPIs
criticos = int(view["Status"].eq("🔴 Crítico").sum())
sem_giro = int(view["Status"].eq("⚫ Sem Giro").sum())
compras = view[view["Comprar_Qtd"] > 0]
capital_parado = view[(view["Status"].eq("⚫ Sem Giro")) & (view["Valor_Estoque"] > 0)]["Valor_Estoque"].sum()
valor_estoque = view["Valor_Estoque"].sum()
transf_count = len(transferencias[transferencias["Produto"].isin(view["Produto"])]) if not transferencias.empty else 0
cobertura_media = view.replace([np.inf, -np.inf], np.nan)["Cobertura_Dias"].mean()

k1, k2, k3, k4 = st.columns(4)
with k1:
    kpi_card("Valor Total em Estoque", brl(valor_estoque), "Base MATR260")
with k2:
    kpi_card("Produtos Críticos", fmt_num(criticos), "Cobertura até 15 dias")
with k3:
    kpi_card("Compras Recomendadas", brl(compras["Comprar_R$"].sum()), f"{len(compras)} produtos")
with k4:
    kpi_card("Capital Parado", brl(capital_parado), f"{sem_giro} produtos sem giro")

k5, k6, k7, k8 = st.columns(4)
with k5:
    kpi_card("Transferências", fmt_num(transf_count), "Sugestões entre ARMZ")
with k6:
    kpi_card("Cobertura Média", f"{fmt_num(cobertura_media, 1)} dias" if not np.isnan(cobertura_media) else "-", "Somente itens com consumo")
with k7:
    kpi_card("Receita Locação", brl(view["Receita_Locacao"].sum()), "Itens com histórico de locação")
with k8:
    kpi_card("Produtos Analisados", fmt_num(len(view)), "Após filtros aplicados")

# Abas
aba1, aba2, aba3, aba4, aba5, aba6, aba7, aba8 = st.tabs([
    "🏠 Radar Executivo",
    "📈 Forecast",
    "🛒 Compras",
    "🔄 Transferências",
    "📦 Capital Parado",
    "🎯 Curva ABC",
    "🏥 Locação",
    "🏢 ARMZ",
])

cols_forecast = [
    "Produto", "Descrição", "Grupo", "Linha", "Estoque_Disponível", "Valor_Estoque", "Qtd_30d", "Qtd_180d",
    "Forecast_30d", "Estoque_Segurança", "Cobertura_Dias", "Comprar_Qtd", "Comprar_R$", "Status", "Ação", "Score_Oportunidade", "Classe ABC"
]

with aba1:
    st.markdown("<div class='section-title'>Radar Executivo</div>", unsafe_allow_html=True)
    radar = view.sort_values(["Score_Oportunidade", "Comprar_R$", "Valor_Estoque"], ascending=False)[cols_forecast].copy()
    st.dataframe(
        radar,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Valor_Estoque": st.column_config.NumberColumn("Valor Estoque", format="R$ %.2f"),
            "Forecast_30d": st.column_config.NumberColumn("Forecast", format="%.2f"),
            "Cobertura_Dias": st.column_config.NumberColumn("Cobertura Dias", format="%.1f"),
            "Comprar_R$": st.column_config.NumberColumn("Comprar R$", format="R$ %.2f"),
            "Score_Oportunidade": st.column_config.ProgressColumn("Score", min_value=0, max_value=100),
        },
    )
    c1, c2 = st.columns(2)
    with c1:
        status_df = view.groupby("Status", as_index=False).agg(Produtos=("Produto", "count"))
        st.plotly_chart(px.bar(status_df, x="Status", y="Produtos", title="Produtos por Status"), use_container_width=True)
    with c2:
        top_compra = compras.sort_values("Comprar_R$", ascending=False).head(10)
        st.plotly_chart(px.bar(top_compra, x="Produto", y="Comprar_R$", title="Top 10 Compras Recomendadas"), use_container_width=True)

with aba2:
    st.markdown("<div class='section-title'>Forecast Inteligente</div>", unsafe_allow_html=True)
    st.caption("Forecast = 70% consumo dos últimos 30 dias + 30% histórico dos últimos 180 dias.")
    st.dataframe(view.sort_values("Cobertura_Dias")[cols_forecast], use_container_width=True, hide_index=True)

with aba3:
    st.markdown("<div class='section-title'>Compras Recomendadas</div>", unsafe_allow_html=True)
    compra_view = compras.sort_values(["Status", "Comprar_R$"], ascending=[True, False])[cols_forecast]
    st.dataframe(compra_view, use_container_width=True, hide_index=True)

with aba4:
    st.markdown("<div class='section-title'>Transferências entre ARMZ</div>", unsafe_allow_html=True)
    if transferencias.empty:
        st.success("Nenhuma transferência recomendada com os critérios atuais.")
    else:
        tv = transferencias[transferencias["Produto"].isin(view["Produto"])].copy()
        st.dataframe(tv, use_container_width=True, hide_index=True)

with aba5:
    st.markdown("<div class='section-title'>Capital Parado e Sem Giro</div>", unsafe_allow_html=True)
    parado = view[(view["Status"].eq("⚫ Sem Giro")) | (view["Dias_Sem_Giro"] >= 90)].copy()
    parado["Faixa Sem Giro"] = pd.cut(
        parado["Dias_Sem_Giro"].fillna(9999),
        bins=[-1, 90, 180, 365, 99999],
        labels=["Até 90 dias", "90 a 180 dias", "180 a 365 dias", "> 365 dias / sem histórico"],
    )
    st.dataframe(parado.sort_values("Valor_Estoque", ascending=False)[[
        "Produto", "Descrição", "Grupo", "Linha", "Estoque_Disponível", "Valor_Estoque", "Última_Movimentação", "Dias_Sem_Giro", "Faixa Sem Giro", "Status"
    ]], use_container_width=True, hide_index=True)

with aba6:
    st.markdown("<div class='section-title'>Curva ABC</div>", unsafe_allow_html=True)
    abc_view = view.sort_values("Receita_Total", ascending=False)[[
        "Produto", "Descrição", "Receita_Total", "Qtd_Total", "Valor_Estoque", "Classe ABC", "Status"
    ]]
    st.dataframe(abc_view, use_container_width=True, hide_index=True)
    abc_chart = abc_view.groupby("Classe ABC", as_index=False).agg(Receita=("Receita_Total", "sum"), Produtos=("Produto", "count"))
    st.plotly_chart(px.bar(abc_chart, x="Classe ABC", y="Receita", text="Produtos", title="Receita por Classe ABC"), use_container_width=True)

with aba7:
    st.markdown("<div class='section-title'>Inteligência de Locação</div>", unsafe_allow_html=True)
    lv = locacao[locacao["Produto"].isin(view["Produto"])].copy()
    if lv.empty:
        st.info("Não há itens de locação identificados nos filtros atuais.")
    else:
        st.dataframe(lv[[
            "Produto", "Descrição", "Grupo", "Linha", "Estoque_Disponível", "Qtd_Locacao", "Receita_Locacao", "Índice_Pressão_Locação", "Status", "Comprar_Qtd"
        ]], use_container_width=True, hide_index=True)
        st.plotly_chart(px.bar(lv.head(15), x="Produto", y="Receita_Locacao", title="Top 15 Receita de Locação"), use_container_width=True)

with aba8:
    st.markdown("<div class='section-title'>Análise por ARMZ</div>", unsafe_allow_html=True)
    resumo_armz = armz_view.groupby("ARMZ", as_index=False).agg(
        Valor_Estoque=("Valor_ARMZ", "sum"),
        Estoque_Disponivel=("Disponível_ARMZ", "sum"),
        Produtos=("Produto", "nunique"),
    )
    st.dataframe(resumo_armz.sort_values("Valor_Estoque", ascending=False), use_container_width=True, hide_index=True)
    st.plotly_chart(px.bar(resumo_armz.sort_values("Valor_Estoque", ascending=False), x="ARMZ", y="Valor_Estoque", title="Valor em Estoque por ARMZ"), use_container_width=True)
    st.dataframe(armz_view.sort_values(["Produto", "ARMZ"]), use_container_width=True, hide_index=True)

# Download consolidado
st.divider()
excel_bytes = to_excel_download({
    "Radar Executivo": view[cols_forecast].sort_values("Score_Oportunidade", ascending=False),
    "Compras": compras[cols_forecast].sort_values("Comprar_R$", ascending=False),
    "Transferencias": transferencias if not transferencias.empty else pd.DataFrame(columns=["Produto", "Descrição", "Origem_ARMZ", "Destino_ARMZ", "Qtd_Sugerida", "Motivo"]),
    "Capital Parado": view[(view["Status"].eq("⚫ Sem Giro")) | (view["Dias_Sem_Giro"] >= 90)],
    "ARMZ": armz_view,
    "Locacao": locacao[locacao["Produto"].isin(view["Produto"])] if not locacao.empty else pd.DataFrame(),
})
st.download_button(
    "📥 Baixar análise em Excel",
    data=excel_bytes,
    file_name=f"first_forecast_estoque_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(f"Última atualização da análise: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
