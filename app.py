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

    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError("Coluna 'Detalhe Conta / Objeto' não existe no W4.")

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

    fluxo = df.get("Fluxo", pd.Series("", index=df.index)).astype(str).str.lower()
    fluxo_vazio = fluxo.str.strip().isin(["", "none", "nan"])
    cond_fluxo_receita = fluxo.str.contains("receita", na=False)
    cond_fluxo_despesa = fluxo.str.contains("despesa", na=False)

    df["is_despesa"] = cond_fluxo_despesa
    df.loc[cond_fluxo_receita, "is_despesa"] = False

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

    if "Id Item tesouraria" in df.columns:
        out["Descrição"] = (
            df["Id Item tesouraria"].astype(str)
            + " "
            + df["Descrição"].astype(str)
        )
    else:
        out["Descrição"] = df["Descrição"]

    out["Cliente/Fornecedor"] = ""
    out["CNPJ/CPF Cliente/Fornecedor"] = ""
    out["Centro de Custo"] = ""
    out["Observações"] = ""

    return out, pd.DataFrame()


# ============================
# NOVAS FUNÇÕES (ACRÉSCIMO)
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
    s = str(v).replace(".", "").replace(",", ".").strip()
    try:
        return float(s)
    except:
        return None


def quebrar_receitas_unificadas(df_convertido, df_ofx):
    novas_linhas = []
    relatorio = []

    for data in df_convertido["Data de Competência"].dropna().unique():

        w4_dia = df_convertido[
            df_convertido["Data de Competência"] == data
        ].copy()

        ofx_dia = df_ofx[df_ofx["data"] == data].copy()

        if ofx_dia.empty:
            novas_linhas.extend(w4_dia.to_dict("records"))
            continue

        # 1️⃣ Manter receitas que já existem no extrato
        for idx, row in w4_dia.iterrows():
            valor_w4 = valor_str_para_float(row["Valor"])
            if valor_w4 is None or valor_w4 <= 0:
                continue

            mask = (
                (~ofx_dia["usado"]) &
                (ofx_dia["valor"].round(2) == round(valor_w4, 2))
            )

            if mask.any():
                ofx_dia.loc[mask.idxmax(), "usado"] = True
                novas_linhas.append(row)
                w4_dia.drop(idx, inplace=True)

        # 2️⃣ Tentar quebrar montantes
        ofx_restante = ofx_dia[~ofx_dia["usado"]]

        for _, row in w4_dia.iterrows():
            valor_w4 = valor_str_para_float(row["Valor"])
            if valor_w4 is None or valor_w4 <= 0:
                novas_linhas.append(row)
                continue

            soma = 0
            usados = []

            for idx_ofx, tx in ofx_restante.iterrows():
                soma = round(soma + tx["valor"], 2)
                usados.append(idx_ofx)

                if abs(soma - valor_w4) < 0.01:
                    for i in usados:
                        tx = ofx_restante.loc[i]
                        nova = row.copy()
                        nova["Valor"] = f"{tx['valor']:.2f}".replace(".", ",")

                        desc_base = str(row["Descrição"]).strip()
                        nova["Descrição"] = (
                            f"{desc_base} | {tx['descricao_ofx']}"
                        ).strip(" |")

                        novas_linhas.append(nova)
                        ofx_restante.loc[i, "usado"] = True

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
# CARREGAR CATEGORIAS
# ============================

df_cat_raw = pd.read_excel("categorias_contabeis.xlsx")
df_cat_prep = preparar_categorias(df_cat_raw)

# ============================
# INTERFACE
# ============================

st.title("🎄 Conversor W4 🎄")
st.markdown("### Envie o arquivo W4 e, se desejar, o OFX")

arq_w4 = st.file_uploader(
    "Selecione o arquivo W4",
    type=["csv", "xlsx", "xls"]
)

arq_ofx = st.file_uploader(
    "Selecione o OFX (opcional – apenas receitas)",
    type=["ofx"]
)

if arq_w4:
    if st.button("Converter arquivo"):
        try:
            df_w4 = carregar_arquivo_w4(arq_w4)
            df_final, _ = converter_w4(df_w4, df_cat_prep)

            if arq_ofx:
                df_ofx = carregar_ofx(arq_ofx)
                df_final, relatorio = quebrar_receitas_unificadas(
                    df_final, df_ofx
                )

                if not relatorio.empty:
                    st.markdown("### 📊 Receitas quebradas automaticamente")
                    st.dataframe(relatorio, use_container_width=True)

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
