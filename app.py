import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import re

# ============================
# CONFIG APP
# ============================

st.set_page_config(
    page_title="Conversor W4",
    layout="centered"
)

# ============================
# CSS – NATAL 🎄
# ============================

st.markdown("""
<style>

body {
    background-image: url('https://images.unsplash.com/photo-1513670800287-29d3b6b4a3d8');
    background-size: cover;
    background-repeat: no-repeat;
    background-attachment: fixed;
}

.block-container {
    backdrop-filter: blur(6px);
    background: rgba(255, 255, 255, 0.75);
    padding: 2rem;
    border-radius: 12px;
}

h1 {
    text-align: center;
    color: #8B0000 !important;
    font-weight: 900 !important;
    text-shadow: 1px 1px 2px #ffffff;
}

.stButton>button {
    background-color: #b30000;
    color: white;
    border-radius: 10px;
    padding: 0.6rem 1.2rem;
    border: none;
    font-weight: bold;
}

.stButton>button:hover {
    background-color: #660000;
}

</style>
""", unsafe_allow_html=True)

# ============================
# FUNCTIONS
# ============================

def normalize_text(texto):
    texto = str(texto).lower().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))
    texto = re.sub(r'[^a-z0-9]+', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()

def preparar_categorias(df_cat):
    col_desc = "Descrição da categoria financeira"
    df = df_cat.copy()

    def tirar_codigo_inicial(txt):
        txt = str(txt).strip()
        partes = txt.split(" ", 1)
        if len(partes) == 2 and any(ch.isdigit() for ch in partes[0]):
            return partes[1].strip()
        return txt

    df["nome_base"] = df[col_desc].apply(tirar_codigo_inicial).apply(normalize_text)
    return df

def formatar_data_coluna(serie):
    datas = pd.to_datetime(serie, errors="coerce")
    return datas.dt.strftime("%d/%m/%Y")

def converter_valor(valor_str, is_despesa):
    if pd.isna(valor_str):
        return None
    txt = str(valor_str).strip().lstrip("+- ")
    txt = txt.replace(".", "").replace(",", ".")
    try:
        numero = float(txt)
    except:
        return None
    if is_despesa:
        numero = -numero
    return numero

# ============================
# CONVERSÃO PRINCIPAL
# ============================

def converter_w4(df_w4, df_categorias_prep):

    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError("Coluna 'Detalhe Conta / Objeto' não existe no arquivo W4.")

    col_cat = "Detalhe Conta / Objeto"

    # Remover transferências
    mascara_transfer = df_w4[col_cat].astype(str).str.contains(
        "Transferência Entre Disponíveis", case=False, na=False)
    df = df_w4.loc[~mascara_transfer].copy()

    # Categorias
    col_desc_cat = "Descrição da categoria financeira"
    df["nome_base_w4"] = df[col_cat].astype(str).apply(normalize_text)

    df = df.merge(
        df_categorias_prep[["nome_base", col_desc_cat]],
        left_on="nome_base_w4",
        right_on="nome_base",
        how="left"
    )

    df["Categoria_final"] = df[col_desc_cat].where(df[col_desc_cat].notna(), df[col_cat])

    # regra do empréstimo
    if "Processo" in df.columns:
        processo_lower = df["Processo"].astype(str).str.lower()
        mask_emp = processo_lower.str.contains("emprestimo", na=False)
        df.loc[mask_emp, "Categoria_final"] = df.loc[mask_emp, "Processo"]

    # ===============================
    # Regras Fluxo + Processo
    # ===============================
    fluxo = df.get("Fluxo", pd.Series("", index=df.index)).astype(str).str.lower()

    cond_receita_fluxo = fluxo.str.contains("receita", na=False)
    cond_despesa_fluxo = fluxo.str.contains("despesa", na=False)

    fluxo_vazio = fluxo.str.strip().isin(["", "nan", "none"])

    if "Processo" in df.columns:
        processo = df["Processo"].astype(str).str.lower()
        cond_pagamento = fluxo_vazio & processo.str.contains("pagamento", na=False)
        cond_recebimento = fluxo_vazio & processo.str.contains("recebimento", na=False)
    else:
        cond_pagamento = False
        cond_recebimento = False

    detalhe_lower = df[col_cat].astype(str).str.lower()

    cond_desp_palavra = (
        fluxo_vazio &
        ~cond_recebimento &
        (detalhe_lower.str.contains("custo", na=False) |
         detalhe_lower.str.contains("despesa", na=False))
    )

    df["is_despesa"] = (
        cond_despesa_fluxo |
        cond_pagamento |
        cond_desp_palavra
    )

    df.loc[cond_receita_fluxo | cond_recebimento, "is_despesa"] = False

    # valor
    df["Valor_str_final"] = [
        converter_valor(v, d) for v, d in zip(df["Valor total"], df["is_despesa"])
    ]

    # Datas = Data da Tesouraria
    data_tes = formatar_data_coluna(df["Data da Tesouraria"])

    # =====================
    # Montagem final
    # =====================

    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes
    out["Valor"] = df["Valor_str_final"]
    out["Categoria"] = df["Categoria_final"]

    # descrição com ID
    if "Id Item tesouraria" in df.columns:
        out["Descrição"] = df["Id Item tesouraria"].astype(str) + " " + df["Descrição"].astype(str)
    else:
        out["Descrição"] = df["Descrição"]

    out["Cliente/Fornecedor CNPJ/CPF"] = ""
    out["Cliente/Fornecedor"] = ""
    out["Centro de Custo"] = ""
    out["Observações"] = ""

    out = out[
        [
            "Data de Competência",
            "Data de Vencimento",
            "Data de Pagamento",
            "Valor",
            "Categoria",
            "Descrição",
            "Cliente/Fornecedor CNPJ/CPF",
            "Cliente/Fornecedor",
            "Centro de Custo",
            "Observações"
        ]
    ]

    return out

# ============================
# CARREGAR W4
# ============================

def carregar_arquivo_w4(arq):
    if arq.name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(arq)
    else:
        return pd.read_csv(arq, sep=";", encoding="latin1")

# ============================
# CARREGAR CATEGORIAS
# ============================

df_cat_raw = pd.read_excel("categorias_contabeis.xlsx")
df_cat_prep = preparar_categorias(df_cat_raw)

# ============================
# UI
# ============================

st.title("🎄 Conversor W4 🎄")

st.markdown("### Envie o arquivo W4 (CSV ou Excel)")

arq_w4 = st.file_uploader("Selecione o arquivo W4", type=["csv]()
