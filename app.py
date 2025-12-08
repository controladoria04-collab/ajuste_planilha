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

def converter_valor(valor_str, is_despesa):
    if pd.isna(valor_str):
        return ""
    texto = str(valor_str).strip().lstrip("+- ")
    texto_num = texto.replace(",", ".")
    try:
        numero = float(texto_num)
    except:
        return valor_str
    if is_despesa:
        numero = -numero
    formatado = "{:,.2f}".format(abs(numero))
    formatado = formatado.replace(",", "X").replace(".", ",").replace("X", ".")
    if numero < 0:
        formatado = "-" + formatado
    return formatado

# ============================
# FUNÇÃO DE CONVERSÃO (CORRIGIDA)
# ============================

def converter_w4(df_w4, df_categorias_prep):

    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError("Coluna 'Detalhe Conta / Objeto' não existe no arquivo W4.")

    col_cat = "Detalhe Conta / Objeto"

    mascara_transfer = df_w4[col_cat].astype(str).str.contains(
        "Transferência Entre Disponíveis", case=False, na=False
    )
    df = df_w4.loc[~mascara_transfer].copy()

    col_desc_cat = "Descrição da categoria financeira"
    df["nome_base_w4"] = df[col_cat].astype(str).apply(normalize_text)

    df = df.merge(
        df_categorias_prep[["nome_base", col_desc_cat]],
        left_on="nome_base_w4",
        right_on="nome_base",
        how="left"
    )

    df["Categoria_final"] = df[col_desc_cat].where(df[col_desc_cat].notna(), df[col_cat])

    if "Processo" in df.columns:
        proc_lower = df["Processo"].astype(str).str.lower()
        mask_emp = proc_lower.str.contains("emprestimo", na=False)
        df.loc[mask_emp, "Categoria_final"] = df.loc[mask_emp, "Processo"]

    fluxo = df.get("Fluxo", pd.Series("", index=df.index)).astype(str).str.lower()
    detalhe_lower = df[col_cat].astype(str).str.lower()

    cond_receita = fluxo.str.contains("receita", na=False)
    cond_despesa_fluxo = fluxo.str.contains("despesa", na=False)
    cond_despesa_palavra = (
        ~cond_receita & ~cond_despesa_fluxo &
        (detalhe_lower.str.contains("custo", na=False) | detalhe_lower.str.contains("despesa", na=False))
    )

    df["is_despesa"] = cond_despesa_fluxo | cond_despesa_palavra

    df["Valor_str_final"] = [
        converter_valor(v, d)
        for v, d in zip(df["Valor total"], df["is_despesa"])
    ]

    # 🔥 TODAS AS DATAS = Data da Tesouraria
    data_tes = formatar_data_coluna(df["Data da Tesouraria"])

    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes

    # 🔥 Colocar ID ANTES da descrição — NOME CORRIGIDO
    if "Id Item tesouraria" in df.columns:
        out["Descrição"] = df["Id Item tesouraria"].astype(str) + " " + df["Descrição"].astype(str)
    else:
        out["Descrição"] = df["Descrição"]

    out["Categoria"] = df["Categoria_final"]
    out["Valor"] = df["Valor_str_final"]

    return out

# ============================
# FUNÇÃO PARA CARREGAR ARQUIVO DO W4
# ============================

def carregar_arquivo_w4(arquivo):
    if arquivo.name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(arquivo)
    else:
        return pd.read_csv(arquivo, sep=";", encoding="latin1")

# ============================
# CARREGA CATEGORIAS
# ============================

CATEGORIAS_ARQ = "categorias_contabeis.xlsx"
df_cat_raw = pd.read_excel(CATEGORIAS_ARQ)
df_cat_prep = preparar_categorias(df_cat_raw)

# ============================
# INTERFACE STREAMLIT
# ============================

st.title("🎄 Conversor W4 🎄")

st.markdown("""
### Envie o arquivo W4 (CSV ou Excel)
Ele será convertido automaticamente para o formato aceito pelo Conta Azul.
""")

arquivo_w4 = st.file_uploader("Selecione o arquivo W4", type=["csv", "xlsx", "xls"])

if arquivo_w4:
    if st.button("Converter arquivo"):
        try:
            df_w4 = carregar_arquivo_w4(arquivo_w4)
            df_final = converter_w4(df_w4, df_cat_prep)

            st.success("Arquivo convertido com sucesso! 🎅✨")

            buffer = BytesIO()
            df_final.to_excel(buffer, index=False)
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
    st.info("Faça o upload do arquivo W4 acima 🎄")
