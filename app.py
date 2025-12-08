import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import re

# ============================
# CONFIG DO APP
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
# FUNÇÕES AUXILIARES
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

# ============================
# CONVERTER VALOR
# ============================

def converter_valor(valor_str, is_despesa):
    if pd.isna(valor_str):
        return ""

    # número vindo como float
    if isinstance(valor_str, (int, float)):
        base = f"{valor_str:.2f}".replace(".", ",")
    else:
        base = str(valor_str).strip()

    base_sem_sinal = base.lstrip("+- ").strip()

    if is_despesa:
        return "-" + base_sem_sinal
    else:
        return base_sem_sinal

# ============================
# FUNÇÃO PRINCIPAL DE CONVERSÃO
# ============================

def converter_w4(df_w4, df_categorias_prep):

    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError("Coluna 'Detalhe Conta / Objeto' não existe no W4.")

    col_cat = "Detalhe Conta / Objeto"

    # Remover transferências
    mascara_transf = df_w4[col_cat].astype(str).str.contains(
        "Transferência Entre Disponíveis", case=False, na=False)
    df = df_w4.loc[~mascara_transf].copy()

    # REMOVER EMPRESTIMOS DA BASE (NOVA REGRA)
    if "Processo" in df.columns:
        proc_rem = df["Processo"].astype(str).str.lower()
        df = df.loc[~proc_rem.str.contains("emprestimo", na=False)].copy()

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

    # ===============================
    # REGRAS DE FLUXO E PROCESSO
    # ===============================

    fluxo = df.get("Fluxo", pd.Series("", index=df.index)).astype(str).str.lower()

    cond_fluxo_receita = fluxo.str.contains("receita", na=False)
    cond_fluxo_despesa = fluxo.str.contains("despesa", na=False)

    fluxo_vazio = fluxo.str.strip().isin(["", "nan", "none"])

    # Normalizar processo
    if "Processo" in df.columns:
        proc = df["Processo"].astype(str).str.lower()
        proc = proc.apply(lambda t: unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode("ascii"))

        cond_pagamento = fluxo_vazio & proc.str.contains("pagament", na=False)
        cond_recebimento = fluxo_vazio & proc.str.contains("receb", na=False)

    else:
        cond_pagamento = False
        cond_recebimento = False

    detalhe_lower = df[col_cat].astype(str).str.lower()

    cond_desp_palavra = (
        fluxo_vazio &
        ~cond_recebimento &
        (
            detalhe_lower.str.contains("custo", na=False) |
            detalhe_lower.str.contains("despesa", na=False)
        )
    )

    # NOVA REGRA: se FLUXO contém imobilizado → é despesa
    cond_imobilizado = fluxo.str.contains("imobilizado", na=False)

    # Regra final combinada
    df["is_despesa"] = (
        cond_fluxo_despesa |
        cond_pagamento |
        cond_desp_palavra |
        cond_imobilizado
    )

    # Receita explícita
    df.loc[cond_fluxo_receita | cond_recebimento, "is_despesa"] = False

    # Valor final (formato W4 + sinal)
    df["Valor_str_final"] = [
        converter_valor(v, d)
        for v, d in zip(df["Valor total"], df["is_despesa"])
    ]

    # Datas (todas = Data da Tesouraria)
    data_tes = formatar_data_coluna(df["Data da Tesouraria"])

    # ============================
    # MONTAGEM FINAL
    # ============================

    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes
    out["Valor"] = df["Valor_str_final"]
    out["Categoria"] = df["Categoria_final"]

    if "Id Item tesouraria" in df.columns:
        out["Descrição"] = df["Id Item tesouraria"].astype(str) + " " + df["Descrição"].astype(str)
    else:
        out["Descrição"] = df["Descrição"]

    out["Cliente/Fornecedor"] = ""
    out["CNPJ/CPF Cliente/Fornecedor"] = ""
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
            "Cliente/Fornecedor",
            "CNPJ/CPF Cliente/Fornecedor",
            "Centro de Custo",
            "Observações"
        ]
    ]

    return out

# ============================
# CARREGAR ARQUIVO W4
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
# INTERFACE STREAMLIT
# ============================

st.title("🎄 Conversor W4 🎄")
st.markdown("### Envie o arquivo W4 (CSV ou Excel)")

arq_w4 = st.file_uploader(
    "Selecione o arquivo W4",
    type=["csv", "xlsx", "xls"]
)

if arq_w4:
    if st.button("Converter arquivo"):
        try:
            df_w4 = carregar_arquivo_w4(arq_w4)
            df_final = converter_w4(df_w4, df_cat_prep)

            st.success("Arquivo convertido com sucesso!")

            buffer = BytesIO()
            df_final.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)

            st.download_button(
                label="🎁 Baixar arquivo convertido",
                data=buffer,
                file_name="conta_azul_convertido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro: {e}")

else:
    st.info("Faça o upload do arquivo acima.")
