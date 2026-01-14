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
# FUNÇÕES AUXILIARES
# ============================
def normalize_text(texto):
    texto = str(texto).lower().strip()
    texto = "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


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


def converter_valor_numerico(valor, is_despesa):
    if pd.isna(valor):
        return None
    try:
        valor = float(valor)
    except:
        return None
    return -abs(valor) if is_despesa else abs(valor)

# ============================
# FUNÇÃO PRINCIPAL
# ============================
def converter_w4(df_w4, df_categorias_prep):
    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError("Coluna 'Detalhe Conta / Objeto' não existe no W4.")

    col_cat = "Detalhe Conta / Objeto"

    df = df_w4.loc[
        ~df_w4[col_cat].astype(str).str.contains(
            "Transferência Entre Disponíveis",
            case=False,
            na=False
        )
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
        df[col_desc_cat].notna(),
        df[col_cat]
    )

    fluxo = df.get("Fluxo", pd.Series("", index=df.index)).astype(str).str.lower()
    fluxo_vazio = fluxo.str.strip().isin(["", "none", "nan"])

    cond_fluxo_receita = fluxo.str.contains("receita", na=False)
    cond_fluxo_despesa = fluxo.str.contains("despesa", na=False)

    proc_original = df.get("Processo", pd.Series("", index=df.index)).astype(str)
    proc = proc_original.str.lower()
    proc = proc.apply(
        lambda t: unicodedata.normalize("NFKD", t)
        .encode("ascii", "ignore")
        .decode("ascii")
    )

    pessoa = df.get("Pessoa", pd.Series("", index=df.index)).astype(str)

    cond_emprestimo = proc.str.contains("emprestimo", na=False)
    cond_pag_emp = proc.str.contains("pagamento", na=False) & cond_emprestimo
    cond_rec_emp = proc.str.contains("recebimento", na=False) & cond_emprestimo

    df.loc[cond_pag_emp, "Categoria_final"] = (
        proc_original[cond_pag_emp] + " " + pessoa[cond_pag_emp]
    )

    df.loc[cond_rec_emp, "Categoria_final"] = (
        proc_original[cond_rec_emp] + " " + pessoa[cond_rec_emp]
    )

    df.loc[
        cond_emprestimo & ~cond_pag_emp & ~cond_rec_emp,
        "Categoria_final"
    ] = proc_original[cond_emprestimo]

    detalhe_lower = df[col_cat].astype(str).str.lower()

    cond_palavra_despesa = (
        fluxo_vazio
        & ~cond_rec_emp
        & (
            detalhe_lower.str.contains("custo", na=False)
            | detalhe_lower.str.contains("despesa", na=False)
        )
    )

    cond_imobilizado = fluxo.str.contains("imobilizado", na=False)

    df["is_despesa"] = (
        cond_fluxo_despesa
        | cond_pag_emp
        | cond_palavra_despesa
        | cond_imobilizado
    )

    df.loc[cond_fluxo_receita | cond_rec_emp, "is_despesa"] = False

    cond_sem_def = (
        df["is_despesa"].isna()
        | (
            (df["is_despesa"] == False)
            & ~cond_fluxo_receita
            & ~cond_rec_emp
            & ~cond_imobilizado
            & ~cond_palavra_despesa
        )
    )

    cond_pag_proc = proc.str.contains("pagamento", na=False)
    cond_rec_proc = proc.str.contains("recebimento", na=False)

    df.loc[cond_sem_def & cond_pag_proc, "is_despesa"] = True
    df.loc[cond_sem_def & cond_rec_proc, "is_despesa"] = False

    df["Valor_final"] = [
        converter_valor_numerico(v, d)
        for v, d in zip(df["Valor total"], df["is_despesa"])
    ]

    df_ignorados = pd.DataFrame()

    if "Id Item tesouraria" in df.columns:
        ids = df["Id Item tesouraria"]
        ids_limpo = ids.astype(str).str.strip()

        mask_duplicado = (
            ids.notna()
            & (ids_limpo != "")
            & ids_limpo.duplicated(keep="first")
        )

        df_ignorados = df.loc[mask_duplicado].copy()
        df = df.loc[~mask_duplicado].copy()

    data_tes = formatar_data_coluna(df["Data da Tesouraria"])

    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes
    out["Valor"] = df["Valor_final"]
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

    return out, df_ignorados

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
st.title("Conversor W4")
st.markdown("### Envie o arquivo W4 (CSV ou Excel)")

arq_w4 = st.file_uploader(
    "Selecione o arquivo W4",
    type=["csv", "xlsx", "xls"]
)

if arq_w4:
    try:
        df_w4 = carregar_arquivo_w4(arq_w4)
        df_final, _ = converter_w4(df_w4, df_cat_prep)

        buffer = BytesIO()
        df_final.to_excel(
            buffer,
            index=False,
            engine="openpyxl"
        )
        buffer.seek(0)

        st.download_button(
            label="📥 Baixar arquivo convertido",
            data=buffer,
            file_name="conta_azul_convertido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro: {e}")
else:
    st.info("Faça o upload do arquivo acima.")
