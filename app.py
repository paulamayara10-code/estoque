# -*- coding: utf-8 -*-
"""
FIRST Intelligence | Forecast Estratégico de Estoque, Compras e Locação
Cruza MATR260 (estoque) + Relatório de Faturamento (guia Base), com atualização diária incremental
 e planejamento estratégico Microtech opcional.
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

st.set_page_config(page_title=APP_NAME, page_icon="📦", layout="wide", initial_sidebar_state="expanded")

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
    """Normaliza código preservando sufixos relevantes e removendo variações simples."""
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = text.replace(" ", "")
    if text.endswith(".0") and re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def produto_base(value: object) -> str:
    """Cria produto-base para unir variações como _RV, _TC, -AT, /RV, REV."""
    text = normalize_code(value)
    text = re.sub(r"([_\-\/\.])(RV|TC|AT|REV|R|V[0-9]+)$", "", text)
    text = re.sub(r"(_RV|_TC|_AT|-RV|-TC|-AT|/RV|/TC|/AT)$", "", text)
    return text


def normalize_col(col: object) -> str:
    text = norm_text(col)
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
    for cand in candidates:
        key = normalize_col(cand)
        for norm, original in normalized.items():
            if key and key in norm:
                return original
    if required:
        raise ValueError(f"Coluna não encontrada. Tentei: {candidates}. Colunas disponíveis: {list(df.columns)}")
    return None


def brl(value: float, short: bool = False) -> str:
    try:
        value = float(value)
    except Exception:
        value = 0.0
    if short:
        abs_v = abs(value)
        if abs_v >= 1_000_000:
            return f"R$ {value/1_000_000:,.2f} MM".replace(",", "X").replace(".", ",").replace("X", ".")
        if abs_v >= 1_000:
            return f"R$ {value/1_000:,.1f} mil".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")



def usd(value: float, short: bool = False) -> str:
    try:
        value = float(value)
    except Exception:
        value = 0.0
    if short:
        abs_v = abs(value)
        if abs_v >= 1_000_000:
            return f"US$ {value/1_000_000:,.2f} MM".replace(",", "X").replace(".", ",").replace("X", ".")
        if abs_v >= 1_000:
            return f"US$ {value/1_000:,.1f} mil".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"US$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

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


def file_bytes_from_path(path: Path) -> Optional[bytes]:
    try:
        if path.exists() and path.suffix.lower() == ".xlsx":
            return path.read_bytes()
    except Exception:
        return None
    return None


def to_excel_download(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe_name = name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
            ws = writer.sheets[safe_name]
            for col_idx, col_name in enumerate(df.columns, start=1):
                width = min(max(len(str(col_name)) + 2, 12), 36)
                ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width
    return output.getvalue()


def money_cols_config(cols: list[str]) -> dict:
    cfg = {}
    for c in cols:
        n = normalize_col(c)
        if any(k in n for k in ["VALOR", "RECEITA", "COMPRAR_R", "CAPITAL", "CUSTO", "USD"]):
            cfg[c] = st.column_config.NumberColumn(c.replace("_", " "), format="R$ %.2f")
        elif any(k in n for k in ["COBERTURA", "FORECAST", "CONSUMO", "INDICE", "SCORE"]):
            cfg[c] = st.column_config.NumberColumn(c.replace("_", " "), format="%.1f")
        elif any(k in n for k in ["QTD", "QUANTIDADE", "ESTOQUE", "SALDO", "DISPONIVEL", "DISPONIVEL"]):
            cfg[c] = st.column_config.NumberColumn(c.replace("_", " "), format="%.0f")
    return cfg


def df_view(df: pd.DataFrame, **kwargs):
    st.dataframe(df, use_container_width=True, hide_index=True, column_config=money_cols_config(list(df.columns)), **kwargs)

# =========================
# LEITURA E TRATAMENTO
# =========================
@st.cache_data(show_spinner=False)
def load_faturamento(file_bytes: bytes, sheet_name: str = "Base", origem: str = "Base") -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, engine="openpyxl")
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
    col_cliente = find_col(df, ["CLIENTE", "Cliente", "Nome Cliente", "Razao Social", "Razão Social"], required=False)
    col_nf = find_col(df, ["Nota Fiscal", "Nota", "NF", "Num NF", "Nro Nota"], required=False)

    out = pd.DataFrame()
    out["Produto_Original"] = df[col_produto].map(normalize_code)
    out["Produto"] = out["Produto_Original"].map(produto_base)
    out["Descrição"] = df[col_desc].fillna("").astype(str).str.strip() if col_desc else ""
    out["Data"] = pd.to_datetime(df[col_data], errors="coerce", dayfirst=True)
    out["Quantidade"] = pd.to_numeric(df[col_qtd], errors="coerce").fillna(0)
    out["Valor"] = pd.to_numeric(df[col_valor], errors="coerce").fillna(0) if col_valor else 0
    out["Grupo_Faturamento"] = df[col_grupo].fillna("").astype(str).str.strip() if col_grupo else ""
    out["Linha"] = df[col_linha].fillna("").astype(str).str.strip() if col_linha else ""
    out["Nova"] = df[col_nova].fillna("").astype(str).str.strip() if col_nova else ""
    out["Classificação"] = df[col_class].fillna("").astype(str).str.strip() if col_class else ""
    out["Finalidade"] = df[col_finalidade].fillna("").astype(str).str.strip() if col_finalidade else ""
    out["Categoria"] = df[col_categoria].fillna("").astype(str).str.strip() if col_categoria else ""
    out["Cliente"] = df[col_cliente].fillna("").astype(str).str.strip() if col_cliente else ""
    out["Nota"] = df[col_nf].fillna("").astype(str).str.strip() if col_nf else ""
    out["Origem"] = origem

    texto_loc = (
        out["Nova"].map(norm_text) + " " + out["Classificação"].map(norm_text) + " " +
        out["Finalidade"].map(norm_text) + " " + out["Categoria"].map(norm_text)
    )
    out["Tipo Operação"] = np.where(texto_loc.str.contains("LOCACAO|LOCAÇÃO", regex=True), "Locação", "Venda/Outros")
    out = out[(out["Produto"] != "") & out["Data"].notna()].copy()
    return out


@st.cache_data(show_spinner=False)
def load_estoque(file_bytes: bytes, sheet_name: Optional[str] = None) -> pd.DataFrame:
    if sheet_name:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, engine="openpyxl")
    else:
        xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
        target = None
        for s in xl.sheet_names:
            if "RELACAO" in normalize_col(s) or "POSICAO" in normalize_col(s):
                target = s
                break
        target = target or xl.sheet_names[-1]
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=target, engine="openpyxl")

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
    out["Produto_Original"] = df[col_codigo].map(normalize_code)
    out["Produto"] = out["Produto_Original"].map(produto_base)
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


def combinar_faturamentos(base: pd.DataFrame, diarios: list[pd.DataFrame]) -> tuple[pd.DataFrame, int, int]:
    if not diarios:
        return base.copy(), 0, 0
    base2 = base.copy()
    daily = pd.concat(diarios, ignore_index=True) if diarios else pd.DataFrame()
    before = len(base2) + len(daily)
    all_df = pd.concat([base2, daily], ignore_index=True)
    # Deduplica priorizando a base original. Quando existir nota, ela pesa na chave; quando não existir, cai para data/produto/cliente/qtd/valor.
    all_df["__nota_key"] = all_df["Nota"].astype(str).str.strip().replace({"nan": "", "None": ""})
    all_df["__valor_key"] = all_df["Valor"].round(2)
    all_df["__qtd_key"] = all_df["Quantidade"].round(4)
    all_df["__data_key"] = all_df["Data"].dt.strftime("%Y-%m-%d")
    key_with_nf = np.where(all_df["__nota_key"].ne(""), all_df["__nota_key"], "SEMNF")
    all_df["__key"] = (
        key_with_nf.astype(str) + "|" + all_df["__data_key"].astype(str) + "|" + all_df["Produto"].astype(str) + "|" +
        all_df["Cliente"].map(norm_text).astype(str) + "|" + all_df["__qtd_key"].astype(str) + "|" + all_df["__valor_key"].astype(str)
    )
    all_df["__prioridade"] = np.where(all_df["Origem"].eq("Base"), 0, 1)
    all_df = all_df.sort_values("__prioridade").drop_duplicates("__key", keep="first")
    adicionadas = int((all_df["Origem"].ne("Base")).sum())
    duplicadas = int(before - len(all_df))
    all_df = all_df.drop(columns=[c for c in all_df.columns if c.startswith("__")])
    return all_df, adicionadas, duplicadas


@st.cache_data(show_spinner=False)
def load_microtech(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    out = {}
    for sheet in xl.sheet_names:
        try:
            df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            out[sheet] = df.dropna(how="all")
        except Exception:
            continue
    return out



def parse_number(value: object) -> float:
    """Converte números vindos de Excel, texto com moeda ou separadores BR/US."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float, np.number)):
        try:
            return float(value)
        except Exception:
            return 0.0
    txt = str(value).strip()
    if not txt:
        return 0.0
    txt = txt.replace("$", "").replace("R$", "").replace("US$", "").replace("USD", "")
    txt = txt.replace(" ", "")
    # Formato BR: 1.234,56
    if "," in txt and "." in txt and txt.rfind(",") > txt.rfind("."):
        txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt and "." not in txt:
        txt = txt.replace(",", ".")
    txt = re.sub(r"[^0-9\.-]", "", txt)
    try:
        return float(txt)
    except Exception:
        return 0.0


def parse_microtech_sales_qty(book: dict[str, pd.DataFrame], sheet_name: str) -> pd.DataFrame:
    df = book.get(sheet_name, pd.DataFrame()).copy()
    if df.empty:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    id_col = df.columns[0]
    out = df.rename(columns={id_col: "SKU" if "SKU" in normalize_col(sheet_name) else "Família"}).copy()
    for c in out.columns[1:]:
        out[c] = out[c].map(parse_number)
    anos = [c for c in out.columns if re.fullmatch(r"20\d{2}", str(c))]
    if anos:
        ano_final = sorted(anos)[-1]
        ano_base = sorted(anos)[-2] if len(anos) >= 2 else None
        out["Qtd Atual"] = out[ano_final]
        out["Ano Atual"] = str(ano_final)
        if ano_base:
            out["Crescimento %"] = np.where(out[ano_base] > 0, (out[ano_final] / out[ano_base]) - 1, np.where(out[ano_final] > 0, 1, 0))
        else:
            out["Crescimento %"] = 0
    return out


def parse_microtech_sales_money(book: dict[str, pd.DataFrame]) -> pd.DataFrame:
    df = book.get("Sales by SKU - $", pd.DataFrame()).copy()
    if df.empty:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    id_col = df.columns[0]
    out = pd.DataFrame({"SKU": df[id_col].map(normalize_code)})
    value_cols = [c for c in df.columns[1:] if "SALES" in normalize_col(c) or re.search(r"20\d{2}", str(c))]
    for c in value_cols:
        year = re.search(r"20\d{2}", str(c))
        name = year.group(0) if year else str(c)
        out[name] = df[c].map(parse_number)
    anos = [c for c in out.columns if re.fullmatch(r"20\d{2}", str(c))]
    if anos:
        out["Receita Atual USD"] = out[sorted(anos)[-1]]
    return out[out["SKU"].ne("")]


def parse_microtech_rolling(book: dict[str, pd.DataFrame]) -> pd.DataFrame:
    # Lê diretamente do workbook raw para preservar as duas linhas de cabeçalho.
    if "Rolling Forecast" not in book:
        return pd.DataFrame()
    try:
        # book já tem header=0; para esta aba é melhor reabrir via ExcelFile no load? Fallback usando estrutura conhecida.
        # Como o header=0 perde a linha superior, tentamos com o dataframe original carregado: as colunas vieram como Unnamed e a primeira linha ainda contém dados.
        # Mais robusto: reconstruir a partir de book com colunas + linhas.
        df0 = book["Rolling Forecast"].copy()
        # Se já veio com colunas estranhas, relê a partir do objeto salvo em st.session não existe. Então tratamos pelo padrão.
        # read_excel padrão deixou linha 0 do arquivo como header; nesse arquivo os dados úteis começam nas linhas com Reference Code.
    except Exception:
        return pd.DataFrame()
    return pd.DataFrame()


def load_microtech_raw(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, engine="openpyxl", header=None)


@st.cache_data(show_spinner=False)
def build_microtech_strategy(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    book_std = {s: pd.read_excel(BytesIO(file_bytes), sheet_name=s, engine="openpyxl").dropna(how="all") for s in xl.sheet_names}

    # 1) Rolling Forecast 2026 por SKU
    rolling = pd.DataFrame()
    if "Rolling Forecast" in xl.sheet_names:
        raw = load_microtech_raw(file_bytes, "Rolling Forecast")
        header_row = None
        for i in range(min(len(raw), 15)):
            row_vals = [normalize_col(v) for v in raw.iloc[i].tolist()]
            if "REFERENCE_CODE" in row_vals:
                header_row = i
                break
        if header_row is not None and header_row > 0:
            months_row = raw.iloc[header_row - 1]
            sub_row = raw.iloc[header_row]
            data = raw.iloc[header_row + 1:].copy()
            code_col = list(sub_row.map(normalize_col)).index("REFERENCE_CODE")
            desc_col = None
            for j, v in enumerate(sub_row.map(normalize_col)):
                if "PRODUCT_DESCRIPTION" in v:
                    desc_col = j
                    break
            qty_cols, usd_cols = [], []
            months = ["JAN", "FEB", "MAR", "ABR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
            for j in range(raw.shape[1]):
                m = norm_text(months_row.iloc[j])
                sub = norm_text(sub_row.iloc[j])
                if m in months and sub == "QTY":
                    qty_cols.append(j)
                if m in months and sub == "USD":
                    usd_cols.append(j)
            rows = []
            for _, r in data.iterrows():
                sku = normalize_code(r.iloc[code_col])
                if not sku or sku in ["TOTAL", "NAN"]:
                    continue
                qtd_total = sum(parse_number(r.iloc[j]) for j in qty_cols)
                usd_total = sum(parse_number(r.iloc[j]) * parse_number(r.iloc[j-1]) if j-1 in qty_cols else 0 for j in usd_cols)
                # O arquivo também possui coluna TOTAL em USD; usa se existir e for maior.
                total_cols = [j for j in range(raw.shape[1]) if norm_text(months_row.iloc[j]) == "TOTAL"]
                total_usd = max([parse_number(r.iloc[j]) for j in total_cols] + [usd_total])
                rows.append({
                    "SKU": sku,
                    "Descrição Microtech": str(r.iloc[desc_col]).strip() if desc_col is not None and not pd.isna(r.iloc[desc_col]) else "",
                    "Forecast 2026 Qtd": qtd_total,
                    "Forecast 2026 USD": total_usd,
                    "Média Mensal Forecast": qtd_total / 12 if qtd_total else 0,
                })
            rolling = pd.DataFrame(rows)
            if not rolling.empty:
                rolling = rolling.groupby(["SKU", "Descrição Microtech"], as_index=False).sum(numeric_only=True)

    # 2) Crescimento por SKU e por Família
    sku_qty = parse_microtech_sales_qty(book_std, "Sales by SKU - QTY") if "Sales by SKU - QTY" in book_std else pd.DataFrame()
    if not sku_qty.empty:
        sku_qty["SKU"] = sku_qty["SKU"].map(normalize_code)
        sku_qty = sku_qty[sku_qty["SKU"].ne("")]
    fam_qty = parse_microtech_sales_qty(book_std, "Sales Hist by FAMILY QTY") if "Sales Hist by FAMILY QTY" in book_std else pd.DataFrame()
    sku_money = parse_microtech_sales_money(book_std)

    # 3) Purchase x Sell Out por família
    pvs = pd.DataFrame()
    if "Purchases vs Sell Out" in xl.sheet_names:
        raw = load_microtech_raw(file_bytes, "Purchases vs Sell Out")
        # Mapeia colunas por ano + trimestre.
        years = []
        last_year = ""
        for j in range(raw.shape[1]):
            y = str(raw.iloc[0, j]).strip()
            if re.fullmatch(r"20\d{2}", y):
                last_year = y
            years.append(last_year)
        qs = [str(raw.iloc[1, j]).strip() for j in range(raw.shape[1])]
        periods = {j: f"{years[j]} {qs[j]}" for j in range(1, raw.shape[1]) if years[j] and re.fullmatch(r"Q[1-4]", qs[j])}
        def collect_section(start_label: str, end_start: int) -> pd.DataFrame:
            rows = []
            for i in range(end_start, len(raw)):
                name = str(raw.iloc[i, 0]).strip()
                if not name or name.upper() in ["NAN", "TOTAL"]:
                    if name.upper() == "TOTAL":
                        continue
                    continue
                if "TOTAL" in name.upper() or name.upper() in ["PURCHASES", "SELL OUT"]:
                    continue
                d = {"Família": name}
                for j, per in periods.items():
                    d[per] = parse_number(raw.iloc[i, j])
                rows.append(d)
            return pd.DataFrame(rows)
        # Blocos conhecidos: compras linhas 2:12, sell-out 16:26
        compras_sec = collect_section("compras", 2).iloc[:10] if len(raw) > 12 else pd.DataFrame()
        sell_sec = collect_section("sell", 16).iloc[:10] if len(raw) > 26 else pd.DataFrame()
        if not compras_sec.empty and not sell_sec.empty:
            compras_long = compras_sec.melt(id_vars="Família", var_name="Período", value_name="Compras")
            sell_long = sell_sec.melt(id_vars="Família", var_name="Período", value_name="Sell Out")
            pvs = compras_long.merge(sell_long, on=["Família", "Período"], how="outer").fillna(0)
            pvs_resumo = pvs.groupby("Família", as_index=False).agg(Compras=("Compras", "sum"), Sell_Out=("Sell Out", "sum"))
            pvs_resumo["Diferença"] = pvs_resumo["Compras"] - pvs_resumo["Sell_Out"]
            pvs_resumo["Cobertura Compra/SellOut"] = np.where(pvs_resumo["Sell_Out"] > 0, pvs_resumo["Compras"] / pvs_resumo["Sell_Out"], np.nan)
            pvs = pvs_resumo.sort_values("Diferença", ascending=False)

    # 4) Master SKU para score Microtech
    master = pd.DataFrame()
    if not sku_qty.empty:
        master = sku_qty[["SKU", "Qtd Atual", "Crescimento %"]].copy()
    if not rolling.empty:
        master = rolling.merge(master, on="SKU", how="outer") if not master.empty else rolling.copy()
    if not sku_money.empty:
        master = master.merge(sku_money[["SKU", "Receita Atual USD"]], on="SKU", how="left") if not master.empty else sku_money.copy()
    if not master.empty:
        for c in ["Forecast 2026 Qtd", "Forecast 2026 USD", "Média Mensal Forecast", "Qtd Atual", "Crescimento %", "Receita Atual USD"]:
            if c not in master.columns:
                master[c] = 0
            master[c] = pd.to_numeric(master[c], errors="coerce").fillna(0)
        # Score: crescimento + forecast + receita atual. Escalas por ranking para funcionar em qualquer arquivo.
        def pct_rank(s):
            return s.rank(pct=True).fillna(0) * 100
        master["Score Microtech"] = (
            pct_rank(master["Forecast 2026 Qtd"]) * 0.35
            + pct_rank(master["Qtd Atual"]) * 0.25
            + pct_rank(master["Receita Atual USD"]) * 0.20
            + np.clip(master["Crescimento %"], -1, 2).add(1).div(3).mul(100) * 0.20
        ).round(1)
        master["Sinal Estratégico"] = np.select(
            [master["Score Microtech"] >= 80, master["Score Microtech"] >= 55, master["Score Microtech"] >= 30],
            ["🚀 Priorizar", "🟡 Monitorar", "🔵 Manter"],
            default="⚪ Baixa prioridade",
        )
        master = master.sort_values("Score Microtech", ascending=False)

    return {"rolling": rolling, "sku_qty": sku_qty, "familia_qty": fam_qty, "sku_money": sku_money, "pvs": pvs, "master": master}

def build_forecast(estoque: pd.DataFrame, faturamento: pd.DataFrame, horizonte: int = 30, dias_seguranca: int = 15):
    fat_all = faturamento.copy()
    fat_consumo = fat_all[~fat_all["Tipo Operação"].eq("Locação")].copy()

    data_ref = fat_all["Data"].max() if not fat_all.empty else pd.Timestamp.today().normalize()
    inicio_30 = data_ref - pd.Timedelta(days=30)
    inicio_180 = data_ref - pd.Timedelta(days=180)

    fat_30 = fat_consumo[fat_consumo["Data"] > inicio_30]
    fat_180 = fat_consumo[fat_consumo["Data"] > inicio_180]

    consumo_30 = fat_30.groupby("Produto", as_index=False).agg(Qtd_30d=("Quantidade", "sum"), Receita_30d=("Valor", "sum"))
    consumo_180 = fat_180.groupby("Produto", as_index=False).agg(Qtd_180d=("Quantidade", "sum"), Receita_180d=("Valor", "sum"))
    ultimo_mov = fat_consumo.groupby("Produto", as_index=False).agg(
        Última_Movimentação=("Data", "max"), Receita_Consumo=("Valor", "sum"), Qtd_Total=("Quantidade", "sum")
    )
    receita_total = fat_all.groupby("Produto", as_index=False).agg(Receita_Total=("Valor", "sum"))
    cad = fat_all.sort_values("Data").groupby("Produto", as_index=False).agg(
        Descrição_Faturamento=("Descrição", "last"), Linha=("Linha", "last"), Grupo_Faturamento=("Grupo_Faturamento", "last")
    )
    loc = fat_all[fat_all["Tipo Operação"].eq("Locação")].groupby("Produto", as_index=False).agg(
        Qtd_Locacao=("Quantidade", "sum"), Receita_Locacao=("Valor", "sum"), Clientes_Locacao=("Cliente", "nunique")
    )

    est_total = estoque.groupby("Produto", as_index=False).agg(
        Descrição_Estoque=("Descrição Estoque", "first"), Grupo_Estoque=("Grupo Estoque", "first"),
        Estoque_Total=("Saldo Estoque", "sum"), Estoque_Disponível=("Estoque Disponível", "sum"),
        Valor_Estoque=("Valor Estoque", "sum"), Qtde_ARMZ=("ARMZ", "nunique")
    )

    base = est_total.merge(cad, on="Produto", how="left").merge(consumo_30, on="Produto", how="left").merge(consumo_180, on="Produto", how="left")
    base = base.merge(ultimo_mov, on="Produto", how="left").merge(receita_total, on="Produto", how="left").merge(loc, on="Produto", how="left")

    for c in ["Qtd_30d", "Receita_30d", "Qtd_180d", "Receita_180d", "Receita_Total", "Receita_Consumo", "Qtd_Total", "Qtd_Locacao", "Receita_Locacao", "Clientes_Locacao"]:
        base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)

    base["Descrição"] = base["Descrição_Faturamento"].fillna(base["Descrição_Estoque"]).fillna("")
    base["Grupo"] = base["Grupo_Estoque"].fillna("").astype(str).str.strip()
    base["Linha"] = base["Linha"].fillna("")

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
    base["Excesso_Estoque"] = (base["Estoque_Disponível"] - (base["Forecast_30d"] + base["Estoque_Segurança"])).clip(lower=0)
    base["Necessidade_Bruta"] = base["Forecast_30d"] + base["Estoque_Segurança"] - base["Estoque_Disponível"]
    base["Comprar_Qtd"] = np.ceil(base["Necessidade_Bruta"].clip(lower=0)).astype(int)
    base["Custo_Médio"] = np.where(base["Estoque_Disponível"] > 0, base["Valor_Estoque"] / base["Estoque_Disponível"], 0)
    base["Comprar_R$"] = base["Comprar_Qtd"] * base["Custo_Médio"]
    base["Excesso_R$"] = base["Excesso_Estoque"] * base["Custo_Médio"]
    base["Dias_Sem_Giro"] = np.where(base["Última_Movimentação"].notna(), (data_ref - base["Última_Movimentação"]).dt.days, np.nan)

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
        [base["Comprar_Qtd"] > 0, base["Status"].eq("⚫ Sem Giro") & (base["Valor_Estoque"] > 0), base["Excesso_Estoque"] > base["Forecast_30d"] * 3],
        ["Comprar", "Avaliar capital parado", "Avaliar excesso"],
        default="Manter",
    )
    base["Score_Oportunidade"] = (
        np.where(base["Status"].eq("🔴 Crítico"), 45, 0) + np.where(base["Status"].eq("🟠 Atenção"), 25, 0) +
        np.where(base["Comprar_Qtd"] > 0, 25, 0) + np.where(base["Receita_Locacao"] > 0, 10, 0) +
        np.where(base["Receita_Total"] > base["Receita_Total"].quantile(0.80), 15, 0)
    ).clip(0, 100)

    abc = base[["Produto", "Descrição", "Receita_Total", "Qtd_Total", "Valor_Estoque"]].copy().sort_values("Receita_Total", ascending=False)
    total_receita = abc["Receita_Total"].sum()
    abc["% Receita"] = np.where(total_receita > 0, abc["Receita_Total"] / total_receita, 0)
    abc["% Acumulado"] = abc["% Receita"].cumsum()
    abc["Classe ABC Receita"] = np.select([abc["% Acumulado"] <= 0.80, abc["% Acumulado"] <= 0.95], ["A", "B"], default="C")
    base = base.merge(abc[["Produto", "Classe ABC Receita"]], on="Produto", how="left")

    armz = estoque.groupby(["Produto", "ARMZ"], as_index=False).agg(
        Estoque_ARMZ=("Saldo Estoque", "sum"), Disponível_ARMZ=("Estoque Disponível", "sum"), Valor_ARMZ=("Valor Estoque", "sum")
    )
    armz = armz.merge(base[["Produto", "Descrição", "Grupo", "Linha", "Consumo_Diário_Forecast", "Forecast_30d", "Estoque_Segurança", "Status"]], on="Produto", how="left")
    armz["Cobertura_ARMZ_Dias"] = np.where(armz["Consumo_Diário_Forecast"].fillna(0) > 0, armz["Disponível_ARMZ"] / armz["Consumo_Diário_Forecast"], np.inf)

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
        origem, destino = sobra.iloc[0], deficit.iloc[0]
        qtd = int(min(origem["Disponível_ARMZ"] - demanda_ref_armz, demanda_ref_armz - destino["Disponível_ARMZ"]))
        if qtd > 0:
            transfer_rows.append({"Produto": p, "Descrição": prod["Descrição"], "Origem_ARMZ": origem["ARMZ"], "Destino_ARMZ": destino["ARMZ"], "Qtd_Sugerida": qtd, "Motivo": "Redistribuir estoque antes de comprar"})
    transfer = pd.DataFrame(transfer_rows)

    locacao = build_locacao_recorrencia(fat_all, base)
    return base, armz, transfer, locacao, data_ref


def max_consecutive_months(periods: list[pd.Period]) -> int:
    if not periods:
        return 0
    ords = sorted(set([p.ordinal for p in periods]))
    best = cur = 1
    for i in range(1, len(ords)):
        if ords[i] == ords[i - 1] + 1:
            cur += 1
            best = max(best, cur)
        else:
            cur = 1
    return best


def build_locacao_recorrencia(fat_all: pd.DataFrame, base: pd.DataFrame) -> pd.DataFrame:
    loc = fat_all[fat_all["Tipo Operação"].eq("Locação")].copy()
    if loc.empty:
        return pd.DataFrame()
    loc["Mes"] = loc["Data"].dt.to_period("M")
    grouped = loc.groupby(["Produto", "Cliente"], as_index=False).agg(
        Receita_Locacao=("Valor", "sum"), Qtd_Locacao=("Quantidade", "sum"), Meses_Faturados=("Mes", "nunique"),
        Primeiro_Faturamento=("Data", "min"), Ultimo_Faturamento=("Data", "max")
    )
    seq = loc.groupby(["Produto", "Cliente"])["Mes"].apply(lambda s: max_consecutive_months(list(s))).reset_index(name="Meses_Consecutivos")
    grouped = grouped.merge(seq, on=["Produto", "Cliente"], how="left")
    grouped["Score_Recorrência"] = np.select(
        [grouped["Meses_Consecutivos"] >= 12, grouped["Meses_Consecutivos"] >= 6, grouped["Meses_Consecutivos"] >= 3, grouped["Meses_Consecutivos"] >= 2, grouped["Meses_Consecutivos"] >= 1],
        [100, 80, 60, 40, 20], default=0
    )
    produto = grouped.groupby("Produto", as_index=False).agg(
        Clientes_Ativos_Estimados=("Cliente", "nunique"), Receita_Locacao=("Receita_Locacao", "sum"), Qtd_Locacao=("Qtd_Locacao", "sum"),
        Score_Recorrência=("Score_Recorrência", "max"), Meses_Consecutivos_Max=("Meses_Consecutivos", "max"),
        Meses_Faturados_Total=("Meses_Faturados", "sum"), Ultimo_Faturamento=("Ultimo_Faturamento", "max")
    )
    produto = produto.merge(base[["Produto", "Descrição", "Grupo", "Linha", "Estoque_Disponível", "Valor_Estoque", "Status", "Comprar_Qtd"]], on="Produto", how="left")
    produto["Índice_Ocupação_Estimado"] = np.where(produto["Estoque_Disponível"].fillna(0) > 0, produto["Clientes_Ativos_Estimados"] / produto["Estoque_Disponível"], np.nan)
    produto["Sinal_Investimento"] = np.select(
        [(produto["Score_Recorrência"] >= 80) & (produto["Índice_Ocupação_Estimado"].fillna(0) >= 0.80), produto["Score_Recorrência"] >= 60],
        ["🔵 Avaliar expansão", "🟢 Locação recorrente"], default="Monitorar"
    )
    return produto.sort_values(["Score_Recorrência", "Receita_Locacao"], ascending=False)

# =========================
# INTERFACE
# =========================
st.markdown(f"""<div class="first-header"><h1>{APP_NAME}</h1><p>{APP_SUBTITLE}</p></div>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 📁 Bases")
    usar_git = st.checkbox("Usar arquivos fixos do Git/dados", value=True)
    fat_file = st.file_uploader("Substituir Faturamento Base", type=["xlsx"], help="Opcional. Se não enviar, usa /dados/faturamento.xlsx")
    diarios_files = st.file_uploader("Adicionar Faturamento Diário", type=["xlsx"], accept_multiple_files=True, help="Somente linhas novas serão somadas; duplicidades serão ignoradas.")
    est_file = st.file_uploader("Substituir Estoque MATR260", type=["xlsx"], help="Opcional. Se não enviar, usa /dados/estoque.xlsx")
    micro_file = st.file_uploader("Planejamento Microtech", type=["xlsx"], help="Opcional. Se não enviar, tenta usar /dados/microtech.xlsx")

    st.markdown("### ⚙️ Parâmetros")
    horizonte = st.number_input("Horizonte do forecast (dias)", min_value=7, max_value=180, value=30, step=1)
    dias_seguranca = st.number_input("Estoque de segurança (dias)", min_value=0, max_value=90, value=15, step=1)
    with st.expander("ℹ️ Metodologia", expanded=False):
    

# Carregar bases fixas ou upload
fat_bytes = fat_file.getvalue() if fat_file else (file_bytes_from_path(DADOS_DIR / "faturamento.xlsx") if usar_git else None)
est_bytes = est_file.getvalue() if est_file else (file_bytes_from_path(DADOS_DIR / "estoque.xlsx") if usar_git else None)
micro_bytes = micro_file.getvalue() if micro_file else (file_bytes_from_path(DADOS_DIR / "microtech.xlsx") if usar_git else None)

if not fat_bytes or not est_bytes:
    st.info("Envie o **Relatório de Faturamento** e o **MATR260 de Estoque**, ou mantenha os arquivos em `/dados` no Git.")
    st.stop()

try:
    faturamento_base = load_faturamento(fat_bytes, "Base", origem="Base")
    diarios = []
    for idx, f in enumerate(diarios_files or [], start=1):
        diarios.append(load_faturamento(f.getvalue(), "Base", origem=f"Diário {idx}"))
    faturamento, linhas_diarias_adicionadas, linhas_duplicadas = combinar_faturamentos(faturamento_base, diarios)
    estoque = load_estoque(est_bytes)
    forecast, armz, transferencias, locacao, data_ref = build_forecast(estoque, faturamento, horizonte=int(horizonte), dias_seguranca=int(dias_seguranca))
except Exception as e:
    st.error("Não consegui processar os arquivos. Verifique se os relatórios estão no layout esperado.")
    st.exception(e)
    st.stop()

# Indicador discreto de atualização
base_info = f"Base: {len(faturamento_base):,} linhas".replace(",", ".")
if diarios_files:
    base_info += f" | Diário adicionado: {linhas_diarias_adicionadas:,} | Duplicidades ignoradas: {linhas_duplicadas:,}".replace(",", ".")
st.caption(f"Data de referência: {data_ref.strftime('%d/%m/%Y')} | {base_info}")

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

criticos = int(view["Status"].eq("🔴 Crítico").sum())
sem_giro = int(view["Status"].eq("⚫ Sem Giro").sum())
compras = view[view["Comprar_Qtd"] > 0]
capital_parado = view[(view["Status"].eq("⚫ Sem Giro")) & (view["Valor_Estoque"] > 0)]["Valor_Estoque"].sum()
valor_estoque = view["Valor_Estoque"].sum()
transf_count = len(transferencias[transferencias["Produto"].isin(view["Produto"])]) if not transferencias.empty else 0
cobertura_media = view.replace([np.inf, -np.inf], np.nan)["Cobertura_Dias"].mean()

k1, k2, k3, k4 = st.columns(4)
with k1: kpi_card("Valor Total em Estoque", brl(valor_estoque, short=True), "MATR260")
with k2: kpi_card("Produtos Críticos", fmt_num(criticos), "Cobertura até 15 dias")
with k3: kpi_card("Compras Recomendadas", brl(compras["Comprar_R$"].sum(), short=True), f"{len(compras)} produtos")
with k4: kpi_card("Capital Parado", brl(capital_parado, short=True), f"{sem_giro} produtos sem giro")

k5, k6, k7, k8 = st.columns(4)
with k5: kpi_card("Transferências", fmt_num(transf_count), "Sugestões entre ARMZ")
with k6: kpi_card("Cobertura Média", f"{fmt_num(cobertura_media, 1)} dias" if not np.isnan(cobertura_media) else "-", "Itens com consumo")
with k7: kpi_card("Receita Locação", brl(view["Receita_Locacao"].sum(), short=True), "Não gera compra automática")
with k8: kpi_card("Produtos Analisados", fmt_num(len(view)), "Após filtros")

aba1, aba2, aba3, aba4, aba5, aba6, aba7, aba8, aba9 = st.tabs([
    "🏠 Radar", "📈 Forecast", "🛒 Compras", "🔄 Transferências", "📦 Capital Parado", "🎯 ABC", "🏥 Locação", "🏢 ARMZ", "📊 Microtech"
])

cols_forecast = ["Produto", "Descrição", "Grupo", "Linha", "Estoque_Disponível", "Valor_Estoque", "Qtd_30d", "Qtd_180d", "Forecast_30d", "Estoque_Segurança", "Cobertura_Dias", "Cobertura_Meses", "Excesso_Estoque", "Excesso_R$", "Comprar_Qtd", "Comprar_R$", "Status", "Ação", "Score_Oportunidade", "Classe ABC Receita"]

with aba1:
    st.markdown("<div class='section-title'>Radar Executivo</div>", unsafe_allow_html=True)
    radar = view.sort_values(["Score_Oportunidade", "Comprar_R$", "Valor_Estoque"], ascending=False)[cols_forecast]
    df_view(radar)
    c1, c2 = st.columns(2)
    with c1:
        status_df = view.groupby("Status", as_index=False).agg(Produtos=("Produto", "count"))
        st.plotly_chart(px.bar(status_df, x="Status", y="Produtos", title="Produtos por Status"), use_container_width=True)
    with c2:
        top_compra = compras.sort_values("Comprar_R$", ascending=False).head(10)
        st.plotly_chart(px.bar(top_compra, x="Produto", y="Comprar_R$", title="Top 10 Compras Recomendadas"), use_container_width=True)

with aba2:
    st.markdown("<div class='section-title'>Forecast Inteligente</div>", unsafe_allow_html=True)
    df_view(view.sort_values("Cobertura_Dias")[cols_forecast])

with aba3:
    st.markdown("<div class='section-title'>Compras Recomendadas</div>", unsafe_allow_html=True)
    df_view(compras.sort_values(["Status", "Comprar_R$"], ascending=[True, False])[cols_forecast])

with aba4:
    st.markdown("<div class='section-title'>Transferências entre ARMZ</div>", unsafe_allow_html=True)
    if transferencias.empty:
        st.success("Nenhuma transferência recomendada com os critérios atuais.")
    else:
        tv = transferencias[transferencias["Produto"].isin(view["Produto"])].copy()
        df_view(tv)

with aba5:
    st.markdown("<div class='section-title'>Capital Parado e Sem Giro</div>", unsafe_allow_html=True)
    parado = view[(view["Status"].eq("⚫ Sem Giro")) | (view["Dias_Sem_Giro"] >= 90)].copy()
    parado["Faixa Sem Giro"] = pd.cut(parado["Dias_Sem_Giro"].fillna(9999), bins=[-1, 90, 180, 365, 99999], labels=["Até 90 dias", "90 a 180 dias", "180 a 365 dias", "> 365 dias / sem histórico"])
    df_view(parado.sort_values("Valor_Estoque", ascending=False)[["Produto", "Descrição", "Grupo", "Linha", "Estoque_Disponível", "Valor_Estoque", "Última_Movimentação", "Dias_Sem_Giro", "Faixa Sem Giro", "Status"]])

with aba6:
    st.markdown("<div class='section-title'>Curva ABC</div>", unsafe_allow_html=True)
    abc_view = view.sort_values("Receita_Total", ascending=False)[["Produto", "Descrição", "Receita_Total", "Qtd_Total", "Valor_Estoque", "Classe ABC Receita", "Status"]]
    df_view(abc_view)
    abc_chart = abc_view.groupby("Classe ABC Receita", as_index=False).agg(Receita=("Receita_Total", "sum"), Produtos=("Produto", "count"))
    st.plotly_chart(px.bar(abc_chart, x="Classe ABC Receita", y="Receita", text="Produtos", title="Receita por Classe ABC"), use_container_width=True)

with aba7:
    st.markdown("<div class='section-title'>Parque de Locação por Recorrência</div>", unsafe_allow_html=True)
    lv = locacao[locacao["Produto"].isin(view["Produto"])] if not locacao.empty else pd.DataFrame()
    if lv.empty:
        st.info("Não há itens de locação identificados nos filtros atuais.")
    else:
        df_view(lv[["Produto", "Descrição", "Grupo", "Linha", "Estoque_Disponível", "Clientes_Ativos_Estimados", "Meses_Consecutivos_Max", "Score_Recorrência", "Receita_Locacao", "Índice_Ocupação_Estimado", "Sinal_Investimento"]])
        st.plotly_chart(px.bar(lv.head(15), x="Produto", y="Receita_Locacao", title="Top 15 Receita de Locação"), use_container_width=True)

with aba8:
    st.markdown("<div class='section-title'>Análise por ARMZ</div>", unsafe_allow_html=True)
    resumo_armz = armz_view.groupby("ARMZ", as_index=False).agg(Valor_Estoque=("Valor_ARMZ", "sum"), Estoque_Disponivel=("Disponível_ARMZ", "sum"), Produtos=("Produto", "nunique"))
    df_view(resumo_armz.sort_values("Valor_Estoque", ascending=False))
    st.plotly_chart(px.bar(resumo_armz.sort_values("Valor_Estoque", ascending=False), x="ARMZ", y="Valor_Estoque", title="Valor em Estoque por ARMZ"), use_container_width=True)
    df_view(armz_view.sort_values(["Produto", "ARMZ"]))

with aba9:
    st.markdown("<div class='section-title'>Planejamento Estratégico Microtech</div>", unsafe_allow_html=True)
    if not micro_bytes:
        st.info("Envie o arquivo de planejamento Microtech ou mantenha `/dados/microtech.xlsx` no Git.")
    else:
        try:
            micro = build_microtech_strategy(micro_bytes)
            master = micro.get("master", pd.DataFrame())
            rolling = micro.get("rolling", pd.DataFrame())
            sku_qty = micro.get("sku_qty", pd.DataFrame())
            familia_qty = micro.get("familia_qty", pd.DataFrame())
            pvs = micro.get("pvs", pd.DataFrame())

            if master.empty and rolling.empty and sku_qty.empty and familia_qty.empty and pvs.empty:
                st.warning("Arquivo Microtech carregado, mas o layout não foi reconhecido para as análises estratégicas.")
            else:
                # Cruza SKU com estoque atual da First quando houver SKU equivalente.
                if not master.empty:
                    master["Produto"] = master["SKU"].map(produto_base)
                    master_join = master.merge(
                        view[["Produto", "Estoque_Disponível", "Valor_Estoque", "Grupo", "Status", "Cobertura_Dias"]],
                        on="Produto", how="left"
                    )
                    master_join["Cobertura Importação Meses"] = np.where(
                        master_join["Média Mensal Forecast"].fillna(0) > 0,
                        master_join["Estoque_Disponível"].fillna(0) / master_join["Média Mensal Forecast"],
                        np.nan,
                    )
                    master_join["Risco Importação"] = np.select(
                        [master_join["Cobertura Importação Meses"].fillna(999) < 3, master_join["Cobertura Importação Meses"].fillna(999) < 6],
                        ["🔴 Risco", "🟡 Atenção"],
                        default="🟢 Confortável",
                    )
                else:
                    master_join = pd.DataFrame()

                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    kpi_card("SKUs Microtech", fmt_num(master["SKU"].nunique() if not master.empty else 0), "Base estratégica")
                with c2:
                    kpi_card("Forecast 2026", fmt_num(master["Forecast 2026 Qtd"].sum() if "Forecast 2026 Qtd" in master else 0), "Quantidade planejada")
                with c3:
                    kpi_card("Forecast USD", usd(master["Forecast 2026 USD"].sum() if "Forecast 2026 USD" in master else 0, short=True), "Planejamento fabricante")
                with c4:
                    risco = int(master_join["Risco Importação"].eq("🔴 Risco").sum()) if not master_join.empty and "Risco Importação" in master_join else 0
                    kpi_card("Risco Importação", fmt_num(risco), "Cobertura < 3 meses")

                mt1, mt2, mt3, mt4 = st.tabs(["🚀 Oportunidades", "📈 Rolling Forecast", "🔁 Compra x Sell-Out", "🏷️ Famílias"])

                with mt1:
                    if master_join.empty:
                        st.info("Sem SKUs suficientes para montar o score estratégico.")
                    else:
                        cols = [
                            "SKU", "Descrição Microtech", "Grupo", "Estoque_Disponível", "Valor_Estoque",
                            "Qtd Atual", "Crescimento %", "Forecast 2026 Qtd", "Média Mensal Forecast",
                            "Cobertura Importação Meses", "Risco Importação", "Receita Atual USD", "Score Microtech", "Sinal Estratégico"
                        ]
                        cols = [c for c in cols if c in master_join.columns]
                        df_view(master_join.sort_values("Score Microtech", ascending=False)[cols].head(300))
                        top = master_join.sort_values("Score Microtech", ascending=False).head(20)
                        st.plotly_chart(px.bar(top, x="SKU", y="Score Microtech", title="Top 20 Oportunidades Microtech"), use_container_width=True)

                with mt2:
                    if rolling.empty:
                        st.info("Rolling Forecast não identificado no arquivo.")
                    else:
                        rview = rolling.sort_values("Forecast 2026 Qtd", ascending=False).head(300)
                        df_view(rview)
                        st.plotly_chart(px.bar(rview.head(20), x="SKU", y="Forecast 2026 Qtd", title="Top 20 SKUs por Forecast 2026"), use_container_width=True)

                with mt3:
                    if pvs.empty:
                        st.info("Compra x Sell-Out não identificado no arquivo.")
                    else:
                        pvs2 = pvs.copy()
                        pvs2["Sinal"] = np.select(
                            [pvs2["Diferença"] > 0, pvs2["Diferença"] < 0],
                            ["Comprou acima do sell-out", "Sell-out acima da compra"],
                            default="Equilibrado",
                        )
                        df_view(pvs2.sort_values("Diferença", ascending=False))
                        st.plotly_chart(px.bar(pvs2, x="Família", y=["Compras", "Sell_Out"], barmode="group", title="Compras Microtech x Sell-Out First"), use_container_width=True)

                with mt4:
                    if familia_qty.empty:
                        st.info("Histórico por família não identificado no arquivo.")
                    else:
                        cols = [c for c in ["Família", "2023", "2024", "2025", "Qtd Atual", "Crescimento %"] if c in familia_qty.columns]
                        df_view(familia_qty.sort_values("Qtd Atual", ascending=False)[cols])
                        st.plotly_chart(px.bar(familia_qty.sort_values("Qtd Atual", ascending=False).head(20), x="Família", y="Qtd Atual", title="Top Famílias Microtech"), use_container_width=True)
        except Exception as e:
            st.error("Não consegui processar a aba Microtech.")
            st.exception(e)

st.divider()
excel_bytes = to_excel_download({
    "Radar Executivo": view[cols_forecast].sort_values("Score_Oportunidade", ascending=False),
    "Compras": compras[cols_forecast].sort_values("Comprar_R$", ascending=False),
    "Transferencias": transferencias if not transferencias.empty else pd.DataFrame(columns=["Produto", "Descrição", "Origem_ARMZ", "Destino_ARMZ", "Qtd_Sugerida", "Motivo"]),
    "Capital Parado": view[(view["Status"].eq("⚫ Sem Giro")) | (view["Dias_Sem_Giro"] >= 90)],
    "ARMZ": armz_view,
    "Locacao": locacao[locacao["Produto"].isin(view["Produto"])] if not locacao.empty else pd.DataFrame(),
})
st.download_button("📥 Baixar análise em Excel", data=excel_bytes, file_name=f"first_forecast_estoque_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption(f"Última atualização da análise: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
