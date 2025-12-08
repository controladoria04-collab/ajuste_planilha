import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import re

# ============================
# CONFIGURAÇÃO DO APP
# ============================

st.set_page_config(
    page_title="Conversor W4",
    layout="centered"
)

# ============================
# CSS – DECORAÇÃO DE NATAL 🎄
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

label, .stMarkdown, .stButton>button {
    font-size: 18px !important;
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
# FUNÇÕES AUXILIARES
# ============================

def normalize_text(texto):
    texto = str(texto).lower().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))
    texto = re.sub(r'[^a-z0-9]+', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

def preparar_categorias(df_cat):
    col_desc = "Descrição da categoria financeira"
    df = df_cat.copy()

    def tirar_codigo_inicial(texto):
        texto = str(texto).strip()
        partes = texto.split(" ", 1)
        if len(partes) == 2 and any(ch.isdigit() for ch in partes[0]):
            return partes[1].strip()
        return texto

    df["nome_base"] = df[col_desc].apply(tirar_codigo_inicial).apply(normalize_text)
    return df

def formatar_data_coluna(serie):
    datas = pd.to_datetime(serie, errors="coerce")
    return datas.dt.strftime("%d/%m/%Y")

# 🔥 AGORA SALVAMOS VALOR COMO NÚMERO REAL, NÃO TEXTO
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
# FUNÇÃO PRINCIPAL — CONVERSÃO
# ============================

def converter_w4(df_w4, df_categorias_prep):

    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError("Coluna 'Detalhe Conta / Objeto' não existe no arquivo W4.")

    col_cat = "Detalhe Conta / Objeto"

    # Remover transferências
    mascara_transfer = df_w4[col_cat].astype(str).str.contains(
        "Transferência Entre Disponíveis", case=False, na=False)
    df = df_w4.loc[~mascara_transfer].copy()

    # Preparar categorias
    col_desc_cat = "Descrição da categoria financeira"
    df["nome_base_w4"] = df[col_cat].astype(str).apply(normalize_text)

    df = df.merge(
        df_categorias_prep[["nome_base", col_desc_cat]],
        left_on="nome_base_w4",
        right_on="nome_base",
        how="left"
    )

    df["Categoria_final"] = df[col_desc_cat].where(df[col_desc_cat].notna(), df[col_cat])

    # 🔥 REGRA DOS EMPRÉSTIMOS — MANTIDA COMO NO CÓDIGO ORIGINAL
    if "Processo" in df.columns:
        proc_lower = df["Processo"].astype(str).str.lower()
        mask_emp = proc_lower.str.contains("emprestimo", na=False)
        df.loc[mask_emp, "Categoria_final"] = df.loc[mask_emp, "Processo"]

    # =========
