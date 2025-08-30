import os
from datetime import date, datetime
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st

# --------------------- Configuração Visual ---------------------
pio.templates.default = "plotly_dark"

st.set_page_config(
    page_title="Dashboard de RH",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado (simples e estável)
st.markdown("""
    <style>
    .stDownloadButton button { border-radius: 8px; padding: 6px 12px; }
    .metric-label { font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Dashboard de Recursos Humanos")
st.markdown("---")


# --------------------- Funções Utilitárias ---------------------
def brl(x: float) -> str:
    """Formata um número como moeda BRL sem encurtar."""
    try:
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "R$ 0,00"


def safe_div(a: float, b: float) -> float:
    """Divisão segura (evita ZeroDivision)."""
    return a / b if b else 0.0


def get_turnover_rate(df: pd.DataFrame) -> float:
    """Calcula a taxa de turnover (anualizada)."""
    if df is None or df.empty:
        return 0.0
    if "data_de_contratacao" not in df.columns or "data_de_demissao" not in df.columns:
        return 0.0

    current_year = date.today().year
    desligados_ano = df[
        (df["status"] == "Desligado") & (df["data_de_demissao"].dt.year == current_year)
    ].shape[0]

    total_inicio = df[df["data_de_contratacao"].dt.year < current_year].shape[0]
    total_fim = df[df["status"] == "Ativo"].shape[0]
    avg_headcount = (total_inicio + total_fim) / 2.0

    return (desligados_ano / avg_headcount) * 100 if avg_headcount > 0 else 0.0


@st.cache_data
def prepare_df(df: pd.DataFrame) -> pd.DataFrame:
    """Padroniza e cria colunas derivadas para o DataFrame."""
    if df is None:
        return pd.DataFrame()

    # Normaliza strings
    obj_cols = df.select_dtypes(include="object").columns
    df[obj_cols] = df[obj_cols].apply(lambda x: x.astype(str).str.strip())

    # Normaliza nomes de colunas
    df.columns = [
        c.strip().lower()
        .replace(" ", "_")
        .replace("á", "a")
        .replace("ã", "a")
        .replace("ç", "c")
        .replace("é", "e")
        .replace("ê", "e")
        for c in df.columns
    ]

    # Datas
    DATE_COLS = ["data_de_nascimento", "data_de_contratacao", "data_de_demissao"]
    for c in DATE_COLS:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], dayfirst=True, errors="coerce")

    # Sexo padronizado
    if "sexo" in df.columns:
        df["sexo"] = (
            df["sexo"].astype(str).str.upper()
            .replace({"MASCULINO": "Masculino", "FEMININO": "Feminino", "M": "Masculino", "F": "Feminino"})
        )

    # Colunas numéricas obrigatórias
    NUMERIC_COLS = ["salario_base", "impostos", "beneficios", "vt", "vr"]
    for col in NUMERIC_COLS:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Avaliação do funcionário (default = 8.23)
    if "avaliacao_do_funcionario" in df.columns:
        df["avaliacao_do_funcionario"] = pd.to_numeric(
            df["avaliacao_do_funcionario"], errors="coerce"
        ).fillna(8.23)
    else:
        df["avaliacao_do_funcionario"] = 8.23

    today = pd.Timestamp(date.today())

    # Idade
    if "data_de_nascimento" in df.columns:
        df["idade"] = ((today - df["data_de_nascimento"]).dt.days // 365).clip(lower=0)

    # Tempo de casa (meses)
    if "data_de_contratacao" in df.columns:
        meses = (today.year - df["data_de_contratacao"].dt.year) * 12 + \
                (today.month - df["data_de_contratacao"].dt.month)
        df["tempo_de_casa_(meses)"] = pd.Series(meses, index=df.index).clip(lower=0)

    # Status
    if "data_de_demissao" in df.columns:
        df["status"] = np.where(df["data_de_demissao"].notna(), "Desligado", "Ativo")
    else:
        df["status"] = "Ativo"

    # Custo total mensal
    df["custo_total_mensal"] = df[["salario_base", "impostos", "beneficios", "vt", "vr"]].sum(axis=1)

    return df


@st.cache_data
def load_data(path_or_bytes) -> pd.DataFrame:
    """Carrega dados do Excel (caminho ou upload)."""
    if path_or_bytes is None:
        return pd.DataFrame()
    df = pd.read_excel(path_or_bytes, sheet_name=0, engine="openpyxl")
    return prepare_df(df)


# --------------------- Sidebar: Fonte de Dados ---------------------
DEFAULT_EXCEL_PATH = "BaseFuncionarios.xlsx"
with st.sidebar:
    st.header("⚙️ Fonte de dados")
    uploaded = st.file_uploader("Carregar Excel (.xlsx)", type=["xlsx"])
    caminho_manual = st.text_input("Ou caminho local do arquivo", value=DEFAULT_EXCEL_PATH)
    st.divider()

    df = None
    fonte = None
    if uploaded is not None:
        try:
            df = load_data(uploaded)
            fonte = "Upload"
        except Exception as e:
            st.error(f"Erro ao ler arquivo (Upload): {e}")
            st.stop()
    else:
        try:
            if not os.path.exists(caminho_manual):
                st.warning(f"Arquivo não encontrado: **{caminho_manual}**")
            else:
                df = load_data(caminho_manual)
                fonte = "Caminho Local"
        except Exception as e:
            st.error(f"Erro ao ler arquivo (Caminho): {e}")
            st.stop()

    if df is not None and not df.empty:
        st.success(f"Dados carregados via **{fonte}**")
        st.caption(f"Linhas: {len(df)} | Colunas: {len(df.columns)}")
    else:
        if fonte:
            st.info("Arquivo carregado, mas sem dados válidos.")


# --------------------- Filtros da Sidebar ---------------------
if df is not None and not df.empty:
    st.sidebar.header("🗂️ Filtros")

    def msel(col_name: str, display_name: str = None):
        """Cria um multiselect para uma coluna, se ela existir."""
        if display_name is None:
            display_name = col_name.replace("_", " ").title()
        if col_name in df.columns:
            vals = sorted(df[col_name].dropna().unique().tolist())
            return st.sidebar.multiselect(display_name, vals, default=[])
        return []

    area_sel = msel("area", "Área")
    nivel_sel = msel("nivel", "Nível")
    cargo_sel = msel("cargo", "Cargo")
    sexo_sel = msel("sexo", "Sexo")
    status_sel = msel("status", "Status")

    nome_busca = st.sidebar.text_input("Buscar por Nome Completo") if "nome_completo" in df.columns else ""

    # Sliders
    faixa_idade = None
    if "idade" in df.columns and not df["idade"].dropna().empty:
        ida_min, ida_max = int(df["idade"].min()), int(df["idade"].max())
        faixa_idade = st.sidebar.slider("Faixa Etária", ida_min, ida_max, (ida_min, ida_max))

    faixa_sal = None
    if "salario_base" in df.columns and not df["salario_base"].dropna().empty:
        sal_min, sal_max = float(df["salario_base"].min()), float(df["salario_base"].max())
        faixa_sal = st.sidebar.slider(
            "Faixa de Salário Base",
            float(sal_min),
            float(sal_max),
            (float(sal_min), float(sal_max)),
            format="R$ %.2f"
        )

    # Aplica filtros
    df_f = df.copy()
    def apply_in(df_, col, values):
        if values and col in df_.columns:
            return df_[df_[col].isin(values)]
        return df_

    df_f = apply_in(df_f, "area", area_sel)
    df_f = apply_in(df_f, "nivel", nivel_sel)
    df_f = apply_in(df_f, "cargo", cargo_sel)
    df_f = apply_in(df_f, "sexo", sexo_sel)
    df_f = apply_in(df_f, "status", status_sel)

    if nome_busca and "nome_completo" in df_f.columns:
        df_f = df_f[df_f["nome_completo"].str.contains(nome_busca, case=False, na=False, regex=False)]

    if faixa_idade and "idade" in df_f.columns:
        df_f = df_f[(df_f["idade"] >= faixa_idade[0]) & (df_f["idade"] <= faixa_idade[1])]

    if faixa_sal and "salario_base" in df_f.columns:
        df_f = df_f[(df_f["salario_base"] >= faixa_sal[0]) & (df_f["salario_base"] <= faixa_sal[1])]

    # --------------------- KPIs ---------------------
    def k_headcount_ativo(d): return int((d["status"] == "Ativo").sum())
    def k_desligados(d): return int((d["status"] == "Desligado").sum())
    def k_folha(d): return float(d.loc[d["status"] == "Ativo", "salario_base"].sum())
    def k_custo_total(d): return float(d.loc[d["status"] == "Ativo", "custo_total_mensal"].sum())
    def k_idade_media(d): return float(d["idade"].mean())
    def k_avaliacao_media(d): return float(d["avaliacao_do_funcionario"].mean())

    tab1, tab2, tab3 = st.tabs(["Dashboard Geral", "Análise de Desligamentos", "Tabela e Exportação"])

    with tab1:
        st.subheader("Métricas Chave")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("👥 Headcount Ativo", f"{k_headcount_ativo(df_f)}")
            st.metric("➡️ Desligados (Total)", f"{k_desligados(df_f)}")
        with col2:
            st.metric("💸 Folha Salarial", brl(k_folha(df_f)))
            st.metric("💰 Custo Total", brl(k_custo_total(df_f)))
        with col3:
            st.metric("👵 Idade Média", f"{k_idade_media(df_f):.1f} anos")
            st.metric("⭐ Avaliação Média", f"{k_avaliacao_media(df_f):.2f}")
        with col4:
            total_rows = len(df_f)
            desligados_pct = (k_desligados(df_f) / total_rows * 100) if total_rows else 0.0
            st.metric("📉 % Desligados", f"{desligados_pct:.2f}%")

        st.markdown("---")
        st.subheader("Gráficos")
        c1, c2 = st.columns(2)
        with c1:
            if "area" in df_f.columns:
                d = df_f.groupby("area").size().reset_index(name="Headcount")
                fig = px.bar(d, x="area", y="Headcount", text_auto=True)
                fig.update_traces(texttemplate="%{text}")  # números completos
                fig.update_yaxes(tickformat="")  # sem abreviação
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            if "cargo" in df_f.columns:
                d = df_f.groupby("cargo")["salario_base"].mean().reset_index()
                fig = px.bar(d, x="cargo", y="salario_base", text_auto=".2f")
                fig.update_yaxes(tickformat=".2f")
                st.plotly_chart(fig, use_container_width=True)

        # Novo gráfico: Distribuição por Sexo (cores personalizadas)
        if "sexo" in df_f.columns:
            st.subheader("Distribuição por Sexo")
            fig_genero = px.pie(
                df_f,
                names="sexo",
                title="Distribuição por Sexo",
                color="sexo",
                color_discrete_map={
                    "Masculino": "darkblue",
                    "Feminino": "deeppink"
                }
            )
            st.plotly_chart(fig_genero, use_container_width=True)

    with tab3:
        st.dataframe(df_f, use_container_width=True)