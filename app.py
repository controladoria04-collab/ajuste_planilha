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
# FUNÇÕES AUXILIARES (COLUNAS / TEXTO)
# ============================

CENTROS_CUSTO_VALIDOS = {"DIACONIA", "FORTALEZA", "BRASIL", "EXTERIOR"}


def clean_colname(name: str) -> str:
    # remove NBSP e espaços no fim
    return str(name).replace("\u00A0", " ").strip()


def make_unique_columns(cols):
    """
    Garante que não existam colunas duplicadas.
    Ex.: "Descrição", "Descrição" -> "Descrição", "Descrição__2"
    """
    cleaned = [clean_colname(c) for c in cols]
    counts = {}
    out = []
    for c in cleaned:
        counts[c] = counts.get(c, 0) + 1
        if counts[c] == 1:
            out.append(c)
        else:
            out.append(f"{c}__{counts[c]}")
    return out


def col(df: pd.DataFrame, name: str) -> pd.Series:
    """
    Sempre devolve uma SERIES.
    Se a coluna estiver duplicada no arquivo de origem,
    o make_unique_columns já terá deixado a primeira sem sufixo.
    """
    name = clean_colname(name)
    if name not in df.columns:
        raise ValueError(f"Coluna '{name}' não existe no W4. Colunas encontradas: {list(df.columns)}")

    s = df[name]
    # segurança extra (caso raro): se virar DataFrame, pega a primeira
    if isinstance(s, pd.DataFrame):
        return s.iloc[:, 0]
    return s


def normalize_text(texto):
    texto = str(texto).lower().strip()
    texto = ''.join(
        c for c in unicodedata.normalize('NFKD', texto)
        if not unicodedata.combining(c)
    )
    texto = re.sub(r'[^a-z0-9]+', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()


def preparar_categorias(df_cat):
    col_desc = "Descrição da categoria financeira"
    df = df_cat.copy()

    def tirar_codigo(txt):
        txt = str(txt).strip()
        parts = txt.split(" ", 1)
        if len(parts) == 2 and any(ch.isdigit() for ch in parts[0]):
            return parts[1].strip()
        return txt

    df["nome_base"] = df[col_desc].apply(tirar_codigo).apply(normalize_text)
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


def extrair_centro_custo(cliente: str) -> str:
    # pega o sufixo depois do último " - "
    partes = [p.strip() for p in str(cliente).split(" - ")]
    if partes and partes[-1].upper() in CENTROS_CUSTO_VALIDOS:
        return partes[-1].upper()
    return ""


# ============================
# MAPEAMENTO PREVIDÊNCIA (UPLOAD)
# ============================

def carregar_mapeamento_previdencia(arq_map):
    """
    Excel/CSV com colunas:
    - Cliente
    - Padrao

    Retorna lista de regras: (padrao_norm, cliente, centro_custo)
    """
    if arq_map is None:
        raise ValueError("Para o setor 'Previdência Brasil', envie o arquivo de mapeamento (Cliente e Padrao).")

    if arq_map.name.lower().endswith((".xlsx", ".xls")):
        dfm = pd.read_excel(arq_map)
    else:
        dfm = pd.read_csv(arq_map, sep=";", encoding="latin1")

    dfm.columns = make_unique_columns(dfm.columns)

    if "Cliente" not in dfm.columns or "Padrao" not in dfm.columns:
        raise ValueError("Arquivo de mapeamento precisa ter colunas: Cliente e Padrao")

    dfm["Cliente"] = dfm["Cliente"].astype(str).str.strip()
    dfm["Padrao"] = dfm["Padrao"].astype(str).str.strip()
    dfm = dfm[(dfm["Cliente"] != "") & (dfm["Padrao"] != "")]

    if dfm.empty:
        raise ValueError("Arquivo de mapeamento está vazio ou sem dados válidos.")

    regras = []
    for cliente, padrao in zip(dfm["Cliente"], dfm["Padrao"]):
        regras.append((
            normalize_text(padrao),
            cliente,
            extrair_centro_custo(cliente)
        ))

    # padrões maiores primeiro (evita casar em um padrão curto antes do correto)
    regras.sort(key=lambda x: len(x[0]), reverse=True)
    return regras


# ============================
# FUNÇÃO PRINCIPAL
# ============================

def converter_w4(df_w4, df_categorias_prep, setor, regras_previdencia=None):

    # segurança: colunas saneadas e únicas
    df_w4 = df_w4.copy()
    df_w4.columns = make_unique_columns(df_w4.columns)

    col_cat = "Detalhe Conta / Objeto"
    col_val = "Valor total"
    col_data = "Data da Tesouraria"
    col_id = "Id Item tesouraria"
    col_desc = "Descrição"

    # valida colunas que você citou
    _ = col(df_w4, col_cat)
    _ = col(df_w4, col_val)
    _ = col(df_w4, col_data)
    # id e descrição podem existir ou não, mas você disse que existem — então validamos também:
    _ = col(df_w4, col_id)
    _ = col(df_w4, col_desc)

    df = df_w4.loc[
        ~col(df_w4, col_cat).astype(str).str.contains(
            "Transferência Entre Disponíveis",
            case=False,
            na=False
        )
    ].copy()

    # ============================
    # CATEGORIAS BASE (categorias_contabeis.xlsx)
    # ============================

    col_desc_cat = "Descrição da categoria financeira"
    df["nome_base_w4"] = col(df, col_cat).astype(str).apply(normalize_text)

    df = df.merge(
        df_categorias_prep[["nome_base", col_desc_cat]],
        left_on="nome_base_w4",
        right_on="nome_base",
        how="left"
    )

    df["Categoria_final"] = df[col_desc_cat].where(
        df[col_desc_cat].notna(),
        col(df, col_cat).astype(str)
    )

    # ============================
    # PREVIDÊNCIA: trocar categoria + cliente + centro de custo
    # ============================

    df["ClienteFornecedor_final"] = ""
    df["CentroCusto_final"] = ""

    if setor == "Previdência Brasil":
        if not regras_previdencia:
            raise ValueError("Regras de Previdência não carregadas.")

        detalhe_norm = col(df, col_cat).astype(str).apply(normalize_text)

        def buscar(txt_norm: str):
            for padrao_norm, cliente, centro in regras_previdencia:
                if padrao_norm and padrao_norm in txt_norm:
                    return cliente, centro
            return "", ""

        pares = detalhe_norm.apply(buscar)
        df["ClienteFornecedor_final"] = pares.apply(lambda x: x[0])
        df["CentroCusto_final"] = pares.apply(lambda x: x[1])

        achou = df["ClienteFornecedor_final"].astype(str).str.strip().ne("")
        df.loc[achou, "Categoria_final"] = "11318 - Repasse Recebido Fundo de Previdência"

    # ============================
    # PROCESSO / RECEITA / DESPESA (SEU ORIGINAL)
    # ============================

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

    # cuidado: sempre atribuir ESCALAR ou SERIES alinhada
    df.loc[cond_pag_emp, "Categoria_final"] = (
        proc_original[cond_pag_emp].astype(str) + " " + pessoa[cond_pag_emp].astype(str)
    )
    df.loc[cond_rec_emp, "Categoria_final"] = (
        proc_original[cond_rec_emp].astype(str) + " " + pessoa[cond_rec_emp].astype(str)
    )
    df.loc[
        cond_emprestimo & ~cond_pag_emp & ~cond_rec_emp,
        "Categoria_final"
    ] = proc_original[cond_emprestimo & ~cond_pag_emp & ~cond_rec_emp].astype(str)

    detalhe_lower = col(df, col_cat).astype(str).str.lower()

    cond_palavra_despesa = (
        fluxo_vazio &
        ~(cond_rec_emp) &
        (
            detalhe_lower.str.contains("custo", na=False) |
            detalhe_lower.str.contains("despesa", na=False)
        )
    )

    cond_imobilizado = fluxo.str.contains("imobilizado", na=False)

    df["is_despesa"] = (
        cond_fluxo_despesa |
        cond_pag_emp |
        cond_palavra_despesa |
        cond_imobilizado
    )

    df.loc[cond_fluxo_receita | cond_rec_emp, "is_despesa"] = False

    cond_sem_def = (
        df["is_despesa"].isna() |
        (
            (df["is_despesa"] == False) &
            (~cond_fluxo_receita) &
            (~cond_rec_emp) &
            (~cond_imobilizado) &
            (~cond_palavra_despesa)
        )
    )

    cond_pag_proc = proc.str.contains("pagamento", na=False)
    cond_rec_proc = proc.str.contains("recebimento", na=False)

    df.loc[cond_sem_def & cond_pag_proc, "is_despesa"] = True
    df.loc[cond_sem_def & cond_rec_proc, "is_despesa"] = False

    # ============================
    # VALORES
    # ============================

    df["Valor_str_final"] = [
        converter_valor(v, d)
        for v, d in zip(col(df, col_val), df["is_despesa"])
    ]

    # ============================
    # IGNORAR DUPLICADOS PELO ID (SEU ORIGINAL)
    # ============================

    df_ignorados = pd.DataFrame()
    ids = col(df, col_id)
    ids_limpo = ids.astype(str).str.strip()

    mask_duplicado = (
        ids.notna() &
        (ids_limpo != "") &
        ids_limpo.duplicated(keep="first")
    )

    if mask_duplicado.any():
        df_ignorados = df.loc[mask_duplicado].copy()
        df = df.loc[~mask_duplicado].copy()

    # ============================
    # DATAS
    # ============================

    data_tes = formatar_data_coluna(col(df, col_data))

    # ============================
    # SAÍDA
    # ============================

    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes
    out["Valor"] = df["Valor_str_final"]
    out["Categoria"] = df["Categoria_final"]

    out["Descrição"] = (
        col(df, col_id).astype(str) + " " + col(df, col_desc).astype(str)
    )

    # Preencher cliente/centro conforme setor
    if setor == "Previdência Brasil":
        out["Cliente/Fornecedor"] = df["ClienteFornecedor_final"]
        out["Centro de Custo"] = df["CentroCusto_final"]
    else:
        out["Cliente/Fornecedor"] = ""
        out["Centro de Custo"] = ""

    out["CNPJ/CPF Cliente/Fornecedor"] = ""
    out["Observações"] = ""

    # ============================
    # SAÍDA IGNORADOS
    # ============================

    out_ignorados = pd.DataFrame()
    if not df_ignorados.empty:
        data_tes_ign = formatar_data_coluna(col(df_ignorados, col_data))

        out_ignorados = pd.DataFrame()
        out_ignorados["Data de Competência"] = data_tes_ign
        out_ignorados["Data de Vencimento"] = data_tes_ign
        out_ignorados["Data de Pagamento"] = data_tes_ign
        out_ignorados["Valor"] = df_ignorados["Valor_str_final"]
        out_ignorados["Categoria"] = df_ignorados["Categoria_final"]

        out_ignorados["Descrição"] = (
            col(df_ignorados, col_id).astype(str) + " " + col(df_ignorados, col_desc).astype(str)
        )

        if setor == "Previdência Brasil":
            out_ignorados["Cliente/Fornecedor"] = df_ignorados["ClienteFornecedor_final"]
            out_ignorados["Centro de Custo"] = df_ignorados["CentroCusto_final"]
        else:
            out_ignorados["Cliente/Fornecedor"] = ""
            out_ignorados["Centro de Custo"] = ""

        out_ignorados["CNPJ/CPF Cliente/Fornecedor"] = ""
        out_ignorados["Observações"] = ""

    return out, out_ignorados


# ============================
# CARREGAR ARQUIVO W4
# ============================

def carregar_arquivo_w4(arq):
    if arq.name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(arq)
    else:
        df = pd.read_csv(arq, sep=";", encoding="latin1")

    df.columns = make_unique_columns(df.columns)
    return df


# ============================
# CARREGAR CATEGORIAS
# ============================

df_cat_raw = pd.read_excel("categorias_contabeis.xlsx")
df_cat_raw.columns = make_unique_columns(df_cat_raw.columns)
df_cat_prep = preparar_categorias(df_cat_raw)

# ============================
# INTERFACE
# ============================

st.title("🎄 Conversor W4 🎄")
st.markdown("### Selecione o setor e envie o arquivo W4")

setor = st.selectbox(
    "Selecione o setor",
    ["Ass. Comunitária", "Sinodalidade", "Previdência Brasil"]
)

arq_map = None
if setor == "Previdência Brasil":
    st.markdown("### Upload do mapeamento (Previdência)")
    arq_map = st.file_uploader(
        "Envie o Excel/CSV de mapeamento (colunas: Cliente e Padrao)",
        type=["csv", "xlsx", "xls"],
        key="map"
    )

st.markdown("### Upload do W4")
arq_w4 = st.file_uploader(
    "Selecione o arquivo W4",
    type=["csv", "xlsx", "xls"],
    key="w4"
)

if arq_w4:
    if st.button("Converter arquivo"):
        try:
            df_w4 = carregar_arquivo_w4(arq_w4)

            regras_previdencia = None
            if setor == "Previdência Brasil":
                regras_previdencia = carregar_mapeamento_previdencia(arq_map)

            df_final, df_ignorados_preview = converter_w4(
                df_w4,
                df_cat_prep,
                setor=setor,
                regras_previdencia=regras_previdencia
            )

            st.success("Arquivo convertido com sucesso!")

            col_esq, col_dir = st.columns([2, 1])

            with col_esq:
                buffer = BytesIO()
                df_final.to_excel(buffer, index=False, engine="openpyxl")
                buffer.seek(0)

                st.download_button(
                    label="🎁 Baixar arquivo convertido",
                    data=buffer,
                    file_name="conta_azul_convertido.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col_dir:
                if not df_ignorados_preview.empty:
                    st.warning(f"⚠️ {len(df_ignorados_preview)} linhas ignoradas (ID duplicado).")
                    st.dataframe(df_ignorados_preview.head(30))
                else:
                    st.info("Nenhuma linha ignorada.")

        except Exception as e:
            st.error(f"Erro: {e}")
else:
    st.info("Faça o upload do arquivo acima.")
