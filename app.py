import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import re

st.set_page_config(page_title="Conversor W4 → Conta Azul", layout="centered")

# ===========================================================
# FUNÇÕES
# ===========================================================

def normalize_text(texto):
    texto = str(texto).lower().strip()
    texto = ''.join(
        c for c in unicodedata.normalize('NFKD', texto)
        if not unicodedata.combining(c)
    )
    texto = re.sub(r'[^a-z0-9]+', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto


def preparar_categorias(df_cat):
    col_desc = "Descrição da categoria financeira"
    df = df_cat.copy()

    def tirar_codigo_inicial(texto):
        texto = str(texto).strip()
        partes = texto.split(" ", 1)
        # se primeira parte é código contábil
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

    texto = str(valor_str).strip()
    texto = texto.lstrip("+- ")
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


def converter_w4(df_w4, df_categorias_prep):
    # Categoria fixa
    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError("Coluna 'Detalhe Conta / Objeto' não existe no arquivo W4.")
    col_cat = "Detalhe Conta / Objeto"

    # Remove transferências
    mascara_transfer = df_w4[col_cat].astype(str).str.contains(
        "Transferência Entre Disponíveis", case=False, na=False
    )
    df = df_w4.loc[~mascara_transfer].copy()

    # Mapeamento de categorias
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

    # Processo com "emprestimo"
    if "Processo" in df.columns:
        proc_lower = df["Processo"].astype(str).str.lower()
        mask_emp = proc_lower.str.contains("emprestimo", na=False)
        df.loc[mask_emp, "Categoria_final"] = df.loc[mask_emp, "Processo"]

    # Determinar despesas
    fluxo = df.get("Fluxo", pd.Series("", index=df.index)).astype(str).str.lower()
    detalhe_lower = df[col_cat].astype(str).str.lower()

    cond_receita = fluxo.str.contains("receita", na=False)
    cond_despesa_fluxo = fluxo.str.contains("despesa", na=False)
    cond_despesa_palavra = (
        ~cond_receita & ~cond_despesa_fluxo &
        (detalhe_lower.str.contains("custo", na=False) |
         detalhe_lower.str.contains("despesa", na=False))
    )

    df["is_despesa"] = cond_despesa_fluxo | cond_despesa_palavra

    # Converter valores BR corretamente
    df["Valor_str_final"] = [
        converter_valor(v, d)
        for v, d in zip(df["Valor total"], df["is_despesa"])
    ]

    # Datas
    data_comp = formatar_data_coluna(df["Data da Tesouraria"])
    data_venc = formatar_data_coluna(df.get("Data de Vencimento", df["Data da Tesouraria"]))
    data_pag = formatar_data_coluna(df.get("Data de Pagamento", df["Data da Tesouraria"]))

    # Planilha final
    out = pd.DataFrame()
    out["Data de Competência"] = data_comp
    out["Data de Vencimento"] = data_venc
    out["Data de Pagamento"] = data_pag
    out["Descrição"] = df["Descrição"]
    out["Categoria"] = df["Categoria_final"]
    out["Valor"] = df["Valor_str_final"]

    return out


def carregar_arquivo_w4(arquivo):
    if arquivo.name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(arquivo)
    else:
        return pd.read_csv(arquivo, sep=";", encoding="latin1")


# ===========================================================
# CARREGAR CATEGORIAS FIXAS
# ===========================================================

CATEGORIAS_ARQ = "categorias_contabeis.xlsx"
try:
    df_cat_raw = pd.read_excel(CATEGORIAS_ARQ)
    df_cat_prep = preparar_categorias(df_cat_raw)
except Exception as e:
    st.error(f"Erro ao carregar '{CATEGORIAS_ARQ}': {e}")
    st.stop()

# ===========================================================
# INTERFACE STREAMLIT
# ===========================================================

st.title("🔄 Conversor W4 → Conta Azul")

st.markdown("""
### Regras aplicadas:
- Categorias vêm de **'Detalhe Conta / Objeto'**
- Datas convertidas para **dd/mm/yyyy**
- Valores convertidos corretamente (ex.: `1306,59` → `1.306,59`)
- Despesas ficam **negativas**
- Se **Processo** contém "emprestimo", essa será a categoria do lançamento
- Remove **Transferência Entre Disponíveis**
""")

arquivo_w4 = st.file_uploader(
    "Envie o arquivo W4 (CSV ou Excel)",
    type=["csv", "xlsx", "xls"]
)

if arquivo_w4:
    if st.button("Converter"):
        try:
            df_w4 = carregar_arquivo_w4(arquivo_w4)
            df_final = converter_w4(df_w4, df_cat_prep)

            st.success("Conversão concluída!")
            st.dataframe(df_final.head(20))

            buffer = BytesIO()
            df_final.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                label="⬇️ Baixar arquivo convertido",
                data=buffer,
                file_name="conta_azul_import.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro: {e}")

else:
    st.info("Envie o arquivo W4 acima.")
