import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import re
from ofxparse import OfxParser

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
}
.stButton>button {
    background-color: #b30000;
    color: white;
    border-radius: 10px;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

# ============================
# FUNÇÕES AUXILIARES EXISTENTES
# ============================

def normalize_text(texto):
    texto = str(texto).lower().strip()
    texto = ''.join(
        c for c in unicodedata.normalize('NFKD', texto)
        if not unicodedata.combining(c)
    )
    texto = re.sub(r'[^a-z0-9]+', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()


def preparar_categorias(df_cat):
    col = "Descrição da categoria financeira"
    df = df_cat.copy()

    def tirar_codigo(txt):
        txt = str(txt).strip()
        parts = txt.split(" ", 1)
        if len(parts) == 2 and any(ch.isdigit() for ch in parts[0]):
            return parts[1].strip()
        return txt

    df["nome_base"] = df[col].apply(tirar_codigo).apply(normalize_text)
    return df


def formatar_data_coluna(serie):
    datas = pd.to_datetime(serie, errors="coerce")
    return datas.dt.strftime("%d/%m/%Y")


def converter_valor(valor_str, is_despesa):
    if pd.isna(valor_str):
        return ""
    if isinstance(valor_str, (int, float)):
        base = f"{valor_str:.2f}".replace(".", ",")
    else:
        base = str(valor_str).strip()
    base_sem_sinal = base.lstrip("+- ").strip()
    return ("-" if is_despesa else "") + base_sem_sinal


# ============================
# FUNÇÃO PRINCIPAL EXISTENTE
# ============================

def converter_w4(df_w4, df_categorias_prep):

    col_cat = "Detalhe Conta / Objeto"

    df = df_w4.loc[
        ~df_w4[col_cat].astype(str)
        .str.contains("Transferência Entre Disponíveis", case=False, na=False)
    ].copy()

    col_desc_cat = "Descrição da categoria financeira"
    df["nome_base_w4"] = df[col_cat].astype(str).apply(normalize_text)

    df = df.merge(
        df_categorias_prep[["nome_base", col_desc_cat]],
        left_on="nome_base_w4",
        right_on="nome_base",
        how="left"
    )

    df["Categoria_final"] = df[col_desc_cat].where(
        df[col_desc_cat].notna(), df[col_cat]
    )

    fluxo = df.get("Fluxo", "").astype(str).str.lower()
    df["is_despesa"] = fluxo.str.contains("despesa", na=False)
    df.loc[fluxo.str.contains("receita", na=False), "is_despesa"] = False

    df["Valor_str_final"] = [
        converter_valor(v, d)
        for v, d in zip(df["Valor total"], df["is_despesa"])
    ]

    data_tes = formatar_data_coluna(df["Data da Tesouraria"])

    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes
    out["Valor"] = df["Valor_str_final"]
    out["Categoria"] = df["Categoria_final"]
    out["Descrição"] = df["Descrição"]

    out["Cliente/Fornecedor"] = ""
    out["CNPJ/CPF Cliente/Fornecedor"] = ""
    out["Centro de Custo"] = ""
    out["Observações"] = ""

    return out, pd.DataFrame()

# ============================
# NOVAS FUNÇÕES – OFX
# ============================

def carregar_ofx(arq_ofx):
    ofx = OfxParser.parse(arq_ofx)
    conta = ofx.accounts[0]

    dados = []
    for t in conta.statement.transactions:
        if float(t.amount) > 0:  # SOMENTE RECEITAS
            dados.append({
                "data": t.date.strftime("%d/%m/%Y"),
                "valor": round(float(t.amount), 2),
                "descricao_ofx": t.memo or "",
                "usado": False
            })
    return pd.DataFrame(dados)


def valor_str_para_float(v):
    if pd.isna(v):
        return None
    try:
        return float(str(v).replace(".", "").replace(",", "."))
    except:
        return None


# ============================
# QUEBRA APENAS COLETA DE MISSA
# ============================

def quebrar_coleta_missa(df_convertido, df_ofx):
    novas_linhas = []
    relatorio = []

    for data in df_convertido["Data de Competência"].dropna().unique():

        w4_dia = df_convertido[
            df_convertido["Data de Competência"] == data
        ].copy()

        ofx_dia = df_ofx[df_ofx["data"] == data].copy()

        for idx, row in w4_dia.iterrows():

            valor = valor_str_para_float(row["Valor"])
            categoria = str(row["Categoria"]).lower()

            # Apenas receitas
            if valor is None or valor <= 0:
                novas_linhas.append(row)
                continue

            # Apenas coleta de missa
            if "coleta" not in categoria:
                novas_linhas.append(row)
                continue

            # Se valor existe no OFX → NÃO quebra
            if not ofx_dia.empty:
                if (ofx_dia["valor"].round(2) == round(valor, 2)).any():
                    novas_linhas.append(row)
                    continue

            # Tentar quebrar
            soma = 0
            usados = []

            for _, tx in ofx_dia.iterrows():
                soma = round(soma + tx["valor"], 2)
                usados.append(tx)

                if abs(soma - valor) < 0.01:
                    for u in usados:
                        nova = row.copy()
                        nova["Valor"] = f"{u['valor']:.2f}".replace(".", ",")
                        nova["Descrição"] = (
                            f"{row['Descrição']} | {u['descricao_ofx']}"
                        ).strip(" |")
                        novas_linhas.append(nova)

                    relatorio.append({
                        "Data": data,
                        "Categoria": row["Categoria"],
                        "Valor Original": row["Valor"],
                        "Quebrado em": len(usados)
                    })
                    break
            else:
                novas_linhas.append(row)

    return pd.DataFrame(novas_linhas), pd.DataFrame(relatorio)

# ============================
# CARREGAR ARQUIVO W4
# ============================

def carregar_arquivo_w4(arq):
    if arq.name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(arq)
    return pd.read_csv(arq, sep=";", encoding="latin1")

# ============================
# CATEGORIAS
# ============================

df_cat_raw = pd.read_excel("categorias_contabeis.xlsx")
df_cat_prep = preparar_categorias(df_cat_raw)

# ============================
# INTERFACE
# ============================

st.title("🎄 Conversor W4 🎄")
st.markdown("### Envie o arquivo W4 e o OFX")

arq_w4 = st.file_uploader("Arquivo W4", ["csv", "xlsx", "xls"])
arq_ofx = st.file_uploader("Arquivo OFX", ["ofx"])

if arq_w4 and st.button("Converter arquivo"):
    try:
        df_w4 = carregar_arquivo_w4(arq_w4)
        df_final, _ = converter_w4(df_w4, df_cat_prep)

        if arq_ofx:
            df_ofx = carregar_ofx(arq_ofx)
            df_final, rel = quebrar_coleta_missa(df_final, df_ofx)

            if not rel.empty:
                st.markdown("### 📊 Coletas de missa quebradas")
                st.dataframe(rel, use_container_width=True)

        buffer = BytesIO()
        df_final.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)

        st.download_button(
            "🎁 Baixar arquivo convertido",
            buffer,
            "conta_azul_convertido.xlsx"
        )

        st.success("Arquivo convertido com sucesso!")

    except Exception as e:
        st.error(f"Erro: {e}")
