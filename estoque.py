from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import io
from msal import ConfidentialClientApplication

# ======================================================
# CONFIGURAÇÃO DA PÁGINA
# ======================================================
st.set_page_config(page_title="COMEX Vision", page_icon="", layout="wide")

# Custom CSS para Design Executivo (Fundo Roxo com Texto Branco para Alto Contraste)
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    
    /* Cards Principais - Estilo Executivo Roxo com Escrita Branca */
    .metric-card {
        background-color: #4B2A85 !important;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
        margin-bottom: 15px;
        text-align: center;
        border: none !important;
    }
    
    .metric-title {
        font-size: 13px !important;
        color: #E2D9F3 !important;
        font-weight: 700 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 8px;
    }
    
    .metric-value {
        font-size: 26px !important;
        color: #FFFFFF !important;
        font-weight: 800 !important;
    }
    
    /* Alertas Rápidos de Lotes - Fontes Ampliadas e Contraste Máximo */
    .alert-box {
        background-color: #5A339B !important;
        border-left: 6px solid #FFB800 !important;
        padding: 22px;
        border-radius: 8px;
        margin-bottom: 15px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    .alert-title {
        color: #FFB800 !important;
        font-size: 19px !important;
        font-weight: 800 !important;
        display: block;
        margin-bottom: 12px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .alert-text {
        color: #FFFFFF !important;
        font-size: 15px !important;
        font-weight: 500 !important;
        line-height: 1.6 !important;
        display: block;
    }
    
    .alert-text strong {
        color: #E2D9F3 !important;
        font-weight: 700 !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# ======================================================
# FUNÇÃO DE FORMATAÇÃO EXECUTIVA (K / MM)
# ======================================================
def formatar_valor_comex(valor, sufixo_moeda=True):
    prefixo = "R$ " if sufixo_moeda else ""
    if abs(valor) >= 1_000_000:
        return f"{prefixo}{valor / 1_000_000:.2f} MM"
    elif abs(valor) >= 1_000:
        return f"{prefixo}{valor / 1_000:.2f} K"
    else:
        return f"{prefixo}{valor:.2f}"


# ======================================================
# FUNÇÕES AUXILIARES - ESTOQUE EM TRÂNSITO
# ======================================================
def formatar_kg(valor):
    """Formata peso em kg/toneladas para leitura executiva."""
    try:
        valor = float(valor)
    except (TypeError, ValueError):
        valor = 0

    if abs(valor) >= 1_000_000:
        return f"{valor / 1_000_000:.2f} mil t"
    elif abs(valor) >= 1_000:
        return f"{valor / 1_000:.2f} t"
    else:
        return f"{valor:.0f} kg"


def converter_numero(valor):
    """Converte número vindo do Excel, tratando padrão brasileiro e americano."""
    if pd.isna(valor):
        return 0
    if isinstance(valor, (int, float)):
        return valor

    texto = str(valor).strip()
    if texto == "":
        return 0

    # Exemplo BR: 1.234,56 -> 1234.56
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")

    return pd.to_numeric(texto, errors="coerce") if pd.notna(texto) else 0


def preparar_dados_transito(df_ped_compra):
    """
    Prepara a aba PED. COMPRA para o painel de Estoque em Trânsito.

    Mapeamento solicitado:
    - Coluna E: filial ou matriz
    - Coluna G: descrição do produto
    - Coluna H: quantidade
    - Coluna K: fornecedor
    - Coluna T: data de entrega no armazém
    """
    if df_ped_compra is None or df_ped_compra.empty:
        return pd.DataFrame()

    df_transito = df_ped_compra.copy()

    # Remove linhas totalmente vazias
    df_transito = df_transito.dropna(how="all")

    # Posições das colunas no Excel (zero-index no pandas)
    posicoes = {
        "Filial": 4,  # Coluna E
        "Produto": 6,  # Coluna G
        "Quantidade_kg": 7,  # Coluna H
        "Fornecedor": 10,  # Coluna K
        "Data_Entrega_Armazem": 19,  # Coluna T
    }

    # Garante que a aba tenha as colunas mínimas esperadas
    maior_posicao = max(posicoes.values())
    if len(df_transito.columns) <= maior_posicao:
        return pd.DataFrame()

    df_transito = df_transito.rename(
        columns={df_transito.columns[pos]: nome for nome, pos in posicoes.items()}
    )

    colunas_base = [
        "Filial",
        "Produto",
        "Quantidade_kg",
        "Fornecedor",
        "Data_Entrega_Armazem",
    ]
    df_transito = df_transito[colunas_base].copy()

    df_transito["Data_Entrega_Armazem"] = pd.to_datetime(
        df_transito["Data_Entrega_Armazem"], errors="coerce", dayfirst=True
    )
    df_transito["Quantidade_kg"] = (
        df_transito["Quantidade_kg"].apply(converter_numero).fillna(0)
    )

    # Considera apenas registros com data de entrega e quantidade positiva
    df_transito = df_transito.dropna(subset=["Data_Entrega_Armazem"])
    df_transito = df_transito[df_transito["Quantidade_kg"] > 0]

    if df_transito.empty:
        return df_transito

    df_transito["Mes_Entrega"] = (
        df_transito["Data_Entrega_Armazem"].dt.to_period("M").dt.to_timestamp()
    )
    df_transito["Mes_Entrega_Label"] = df_transito["Data_Entrega_Armazem"].dt.strftime(
        "%m/%Y"
    )
    df_transito["Dias_Para_Entrega"] = (
        df_transito["Data_Entrega_Armazem"] - DATA_HOJE
    ).dt.days

    return df_transito


# ======================================================
# CONFIGURAÇÃO SHAREPOINT (EDITE AQUI SE NECESSÁRIO)
# ======================================================
SHAREPOINT_HOST = "kempartsquimica.sharepoint.com"
SITE_PATH = "/sites/IMPORTACAO"
FILE_PATH = "/BASE DASHBOARD/PLANEJAMENTO DE COMPRAS 30.06.2026_FRANKLIN.xlsx"


# ======================================================
# AUTENTICAÇÃO E LEITURA DO SHAREPOINT
# ======================================================
def get_access_token():
    app = ConfidentialClientApplication(
        st.secrets["AZURE_CLIENT_ID"],
        authority=f"https://login.microsoftonline.com/{st.secrets['AZURE_TENANT_ID']}",
        client_credential=st.secrets["AZURE_CLIENT_SECRET"],
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise Exception(f"Erro de autenticação: {result.get('error_description')}")
    return result["access_token"]


@st.cache_data(ttl=3600)  # Atualiza a cada 1 hora
def carregar_dados_sharepoint():
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    # Obter ID do site
    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOST}:{SITE_PATH}"
    site_resp = requests.get(site_url, headers=headers)
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]

    # Baixar o arquivo Excel
    file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{FILE_PATH}:/content"
    file_resp = requests.get(file_url, headers=headers)
    file_resp.raise_for_status()

    # Ler com pandas direto da memória
    arquivo_excel = io.BytesIO(file_resp.content)
    df_estoque = pd.read_excel(arquivo_excel, sheet_name="ESTOQUE", header=1)

    # A aba PED. COMPRA alimenta o módulo de Estoque em Trânsito.
    # Caso a aba não exista ou esteja com problema, o painel de estoque atual continua funcionando.
    try:
        arquivo_excel.seek(0)
        df_ped_compra = pd.read_excel(arquivo_excel, sheet_name="PED. COMPRA", header=1)
    except Exception:
        df_ped_compra = pd.DataFrame()

    return df_estoque, df_ped_compra


# ======================================================
# CAPA / TOPO PREMIUM
# ======================================================
BASE_DIR = Path(__file__).parent
CAPA = BASE_DIR / "assets" / "comexcapa.png"
DATA_HOJE = pd.to_datetime("2026-06-30")

if CAPA.exists():
    st.image(str(CAPA), use_container_width=True)

st.subheader("Inteligência e Gestão de Estoques de Importação")
st.caption(
    f"Análise Estratégica Baseada em Dados Atualizados | Data de Referência: {DATA_HOJE.strftime('%d/%m/%Y')}"
)
st.divider()

# ======================================================
# LEITURA E TRATAMENTO DE DADOS
# ======================================================
try:
    with st.spinner("Carregando dados do SharePoint..."):
        df, df_ped_compra = carregar_dados_sharepoint()

    df = df.dropna(subset=["Produto", "Dt. Entrada", "Data Validad"])
    df["Dt. Entrada"] = pd.to_datetime(df["Dt. Entrada"], errors="coerce")
    df["Data Validad"] = pd.to_datetime(df["Data Validad"], errors="coerce")

    df["Saldo 1a.U.M."] = pd.to_numeric(df["Saldo 1a.U.M."], errors="coerce").fillna(0)
    df["C Unitario"] = pd.to_numeric(df["C Unitario"], errors="coerce").fillna(0)
    df["Custo total estoque"] = df["Saldo 1a.U.M."] * df["C Unitario"]

    # 1. Regra de Giro de Lote
    df["Dias_no_Estoque"] = (DATA_HOJE - df["Dt. Entrada"]).dt.days
    df["Classificacao_Giro"] = df["Dias_no_Estoque"].apply(
        lambda x: "Baixo Giro" if x > 120 else "Giro Regular"
    )

    # 2. Regra de Saúde dos Lotes
    df["Shelf life "] = (
        pd.to_numeric(df["Shelf life "], errors="coerce").fillna(1).replace(0, 1)
    )
    df["Dias_Restantes"] = (df["Data Validad"] - DATA_HOJE).dt.days

    def calcular_saude_lote(row):
        if row["Dias_Restantes"] < 0:
            return "VENCIDO"
        elif row["Dias_Restantes"] < (row["Shelf life "] * 0.5):
            return "À VENCER"
        else:
            return "VALIDADE OK"

    df["Classificacao_Saude"] = df.apply(calcular_saude_lote, axis=1)

    # ======================================================
    # TRATAMENTO DA ABA PED. COMPRA - ESTOQUE EM TRÂNSITO
    # ======================================================
    df_transito = preparar_dados_transito(df_ped_compra)

    # ======================================================
    # FILTROS EM LINHA (OTIMIZADOS COM OPÇÃO "TODOS")
    # ======================================================
    st.markdown("###  Parâmetros de Consulta")

    lista_produtos_base = sorted(df["Descricao"].dropna().unique())
    opcoes_produto = ["TODOS"] + lista_produtos_base

    opcoes_filial = ["TODAS"] + sorted(list(df["Filial"].dropna().unique()))

    col_p1, col_p2 = st.columns([2, 1])

    with col_p1:
        produtos_selecionados = st.multiselect(
            " Escolha os Produtos para Análise de Lotes:",
            options=opcoes_produto,
            default=["TODOS"],
        )
    with col_p2:
        filial_selecionada = st.multiselect(
            " Filtrar Filiais:", options=opcoes_filial, default=["TODAS"]
        )

    # --- LÓGICA DE FILTRAGEM INTELIGENTE ---
    if "TODOS" in produtos_selecionados or not produtos_selecionados:
        df_filtrado = df.copy()
    else:
        df_filtrado = df[df["Descricao"].isin(produtos_selecionados)]

    if not ("TODAS" in filial_selecionada or not filial_selecionada):
        df_filtrado = df_filtrado[df_filtrado["Filial"].isin(filial_selecionada)]
        df_macro = df[df["Filial"].isin(filial_selecionada)]
    else:
        df_macro = df.copy()

    st.write("")

    # ======================================================
    # REQUISITOS DA GERENTE (INSIGHTS CRÍTICOS DO PRODUTO)
    # ======================================================
    if not df_filtrado.empty:
        st.markdown("###  Raio-X Operacional dos Itens Selecionados")

        lote_prox_vencido = df_filtrado.sort_values(by="Data Validad").iloc[0]
        lote_mais_antigo = df_filtrado.sort_values(by="Dt. Entrada").iloc[0]

        c_ins1, c_ins2 = st.columns(2)

        with c_ins1:
            qtd_formatada = formatar_valor_comex(
                lote_prox_vencido["Saldo 1a.U.M."], sufixo_moeda=False
            )
            st.markdown(
                f"""
            <div class="alert-box">
                <span class="alert-title"> Lote Mais Próximo do Vencimento</span>
                <span class="alert-text">
                    <strong>Produto:</strong> {lote_prox_vencido["Descricao"]}<br>
                    <strong>Lote / Fornecedor:</strong> {lote_prox_vencido["Lote"]} ({lote_prox_vencido["Fornecedor"]})<br>
                    <strong>Vencimento:</strong> {lote_prox_vencido["Data Validad"].strftime("%d/%m/%Y")}<br>
                    <strong>Quantidade:</strong> {qtd_formatada} Un. | <strong>Status:</strong> {lote_prox_vencido["Classificacao_Saude"]}
                </span>
            </div>
            """,
                unsafe_allow_html=True,
            )

        with c_ins2:
            custo_antigo_formatado = formatar_valor_comex(
                lote_mais_antigo["Custo total estoque"]
            )
            st.markdown(
                f"""
            <div class="alert-box" style="border-left-color: #00E5FF !important;">
                <span class="alert-title" style="color: #00E5FF !important;"> Lote Mais Antigo em Estoque</span>
                <span class="alert-text">
                    <strong>Produto:</strong> {lote_mais_antigo["Descricao"]}<br>
                    <strong>Lote / Entrada:</strong> {lote_mais_antigo["Lote"]} ({lote_mais_antigo["Dt. Entrada"].strftime("%d/%m/%Y")})<br>
                    <strong>Tempo de Casa:</strong> {lote_mais_antigo["Dias_no_Estoque"]} dias retido<br>
                    <strong>Capital Imobilizado:</strong> {custo_antigo_formatado}
                </span>
            </div>
            """,
                unsafe_allow_html=True,
            )

    else:
        st.warning("Sem registros encontrados para os filtros aplicados.")

    st.divider()

    # ======================================================
    # TABS DE NAVEGAÇÃO PRINCIPAL
    # ======================================================
    tab_visao_geral, tab_detalhes_lote, tab_graficos, tab_transito = st.tabs(
        [
            " VISÃO GERAL DO PRODUTO",
            " TODOS OS LOTES DESTE PRODUTO",
            " ANÁLISE GRÁFICA MACRO",
            " ESTOQUE EM TRÂNSITO",
        ]
    )

    # --------------------------------------------------
    # TAB 1: VISÃO GERAL E KPIs
    # --------------------------------------------------
    with tab_visao_geral:
        vlr_total = (
            df_filtrado["Custo total estoque"].sum() if not df_filtrado.empty else 0
        )
        vlr_vencido = (
            df_filtrado[df_filtrado["Classificacao_Saude"] == "VENCIDO"][
                "Custo total estoque"
            ].sum()
            if not df_filtrado.empty
            else 0
        )
        vlr_risco = (
            df_filtrado[df_filtrado["Classificacao_Saude"] == "À VENCER"][
                "Custo total estoque"
            ].sum()
            if not df_filtrado.empty
            else 0
        )
        vlr_baixo_giro = (
            df_filtrado[df_filtrado["Classificacao_Giro"] == "Baixo Giro"][
                "Custo total estoque"
            ].sum()
            if not df_filtrado.empty
            else 0
        )

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(
                f'<div class="metric-card"><div class="metric-title">Custo Total em Estoque</div><div class="metric-value">{formatar_valor_comex(vlr_total)}</div></div>',
                unsafe_allow_html=True,
            )
        with col2:
            st.markdown(
                f'<div class="metric-card"><div class="metric-title">Custo Total Vencido</div><div class="metric-value">{formatar_valor_comex(vlr_vencido)}</div></div>',
                unsafe_allow_html=True,
            )
        with col3:
            st.markdown(
                f'<div class="metric-card"><div class="metric-title">Custo em Risco (À Vencer)</div><div class="metric-value">{formatar_valor_comex(vlr_risco)}</div></div>',
                unsafe_allow_html=True,
            )
        with col4:
            st.markdown(
                f'<div class="metric-card"><div class="metric-title">Estoque de Baixo Giro</div><div class="metric-value">{formatar_valor_comex(vlr_baixo_giro)}</div></div>',
                unsafe_allow_html=True,
            )

    # --------------------------------------------------
    # TAB 2: DETALHES DOS LOTES DO PRODUTO
    # --------------------------------------------------
    with tab_detalhes_lote:
        st.markdown("#### Lista Completa de Lotes")
        if not df_filtrado.empty:
            df_exibir = df_filtrado[
                [
                    "Filial",
                    "Descricao",
                    "Lote",
                    "Lote Fornec.",
                    "Dt. Entrada",
                    "Data Validad",
                    "Saldo 1a.U.M.",
                    "C Unitario",
                    "Custo total estoque",
                    "Classificacao_Giro",
                    "Classificacao_Saude",
                ]
            ].copy()

            df_exibir["Dt. Entrada"] = df_exibir["Dt. Entrada"].dt.strftime("%d/%m/%Y")
            df_exibir["Data Validad"] = df_exibir["Data Validad"].dt.strftime(
                "%d/%m/%Y"
            )

            st.dataframe(
                df_exibir.style.format(
                    {
                        "Saldo 1a.U.M.": "{:,.2f}",
                        "C Unitario": "R$ {:,.2f}",
                        "Custo total estoque": "R$ {:,.2f}",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("Nenhum lote para listar.")

    # --------------------------------------------------
    # TAB 3: VISÃO MACRO GRÁFICA + INSPEÇÃO EXECUTIVA (DRILL-DOWN)
    # --------------------------------------------------
    with tab_graficos:
        if not df_macro.empty:
            g1, g2 = st.columns(2)

            def obtener_dados_pizza(df_base, coluna_grupo):
                res_grupo = (
                    df_base.groupby(coluna_grupo)["Custo total estoque"]
                    .sum()
                    .reset_index()
                )
                maiores_produtos = []
                for cat in res_grupo[coluna_grupo]:
                    sub_df = df_base[df_base[coluna_grupo] == cat]
                    if not sub_df.empty:
                        idx_max = sub_df["Custo total estoque"].idxmax()
                        prod_nome = sub_df.loc[idx_max, "Descricao"]
                        prod_val = sub_df.loc[idx_max, "Custo total estoque"]
                        maiores_produtos.append(
                            f"{prod_nome} ({formatar_valor_comex(prod_val)})"
                        )
                    else:
                        maiores_produtos.append("Nenhum")

                res_grupo["Maior_Impacto"] = maiores_produtos
                return res_grupo

            # --- GRÁFICO 1: VELOCIDADE DE GIRO ---
            with g1:
                st.markdown("#####  Giro Total do Estoque Macro")
                dados_giro = obtener_dados_pizza(df_macro, "Classificacao_Giro")

                fig_giro = px.pie(
                    dados_giro,
                    names="Classificacao_Giro",
                    values="Custo total estoque",
                    hole=0.4,
                    color_discrete_sequence=["#E91010", "#07E249"],
                    custom_data=["Maior_Impacto"],
                )
                fig_giro.update_traces(
                    textinfo="percent+label",
                    textposition="auto",
                    insidetextorientation="horizontal",
                    insidetextfont=dict(size=13, color="#FFFFFF", family="Arial Black"),
                    outsidetextfont=dict(
                        size=13, color="#FFFFFF", family="Arial Black"
                    ),
                    hovertemplate="<b>%{label}</b><br>Valor Total: R$ %{value:,.2f}<br><b>Maior Impacto:</b> %{customdata[0]}<extra></extra>",
                )
                fig_giro.update_layout(
                    showlegend=False, margin=dict(t=30, b=30, l=30, r=30)
                )
                st.plotly_chart(fig_giro, use_container_width=True)

            # --- GRÁFICO 2: SAÚDE DO LOTE ---
            with g2:
                st.markdown("#####  Saúde Geral dos Lotes Macro (%)")
                dados_saude = obtener_dados_pizza(df_macro, "Classificacao_Saude")
                cores_saude = {
                    "VALIDADE OK": "#07E249",
                    "À VENCER": "#E9DA0D",
                    "VENCIDO": "#E91010",
                }

                fig_saude = px.pie(
                    dados_saude,
                    names="Classificacao_Saude",
                    values="Custo total estoque",
                    hole=0.4,
                    color="Classificacao_Saude",
                    color_discrete_map=cores_saude,
                    custom_data=["Maior_Impacto"],
                )
                fig_saude.update_traces(
                    textinfo="percent+label",
                    textposition="auto",
                    insidetextorientation="horizontal",
                    insidetextfont=dict(size=13, color="#FFFFFF", family="Arial Black"),
                    outsidetextfont=dict(
                        size=13, color="#FFFFFF", family="Arial Black"
                    ),
                    hovertemplate="<b>%{label}</b><br>Valor Total: R$ %{value:,.2f}<br><b>Maior Impacto:</b> %{customdata[0]}<extra></extra>",
                )
                fig_saude.update_layout(
                    showlegend=False, margin=dict(t=30, b=30, l=30, r=30)
                )
                st.plotly_chart(fig_saude, use_container_width=True)

            # --- CENTRAL DE AUDITORIA E INSIGHTS PARA A DIRETORIA ---
            st.write("")
            st.markdown("---")
            st.markdown("###  Central de Insights Gerenciais (Drill-Down)")
            st.caption(
                "Selecione um indicador crítico dos gráficos acima para auditar os gargalos reais da operação."
            )

            categoria_analise = st.selectbox(
                " Qual indicador macro deseja auditar detalhadamente?",
                options=[
                    "Selecione uma métrica...",
                    "Estoque de Baixo Giro (Imobilizado > 120 dias)",
                    "Lotes Vencidos / À Vencer (Risco Financeiro)",
                ],
            )

            if categoria_analise == "Estoque de Baixo Giro (Imobilizado > 120 dias)":
                df_gargalo = df_macro[df_macro["Classificacao_Giro"] == "Baixo Giro"]

                if not df_gargalo.empty:
                    vlr_parado = df_gargalo["Custo total estoque"].sum()
                    tempo_medio = int(df_gargalo["Dias_no_Estoque"].mean())

                    st.markdown(f"####  Diagnóstico: Custo de Oportunidade Bloqueado")

                    c_b1, c_b2, c_b3 = st.columns(3)
                    c_b1.metric(
                        "Capital Total Parado", formatar_valor_comex(vlr_parado)
                    )
                    c_b2.metric("Tempo Médio de Retenção", f"{tempo_medio} dias")
                    c_b3.metric(
                        "Itens Críticos Afetados",
                        f"{df_gargalo['Produto'].nunique()} SKUs",
                    )

                    st.markdown("#####  Top 5 Produtos com Maior Capital Imobilizado")
                    df_top_giro = (
                        df_gargalo.groupby("Descricao")
                        .agg(
                            Capital_Retido=("Custo total estoque", "sum"),
                            Dias_Max_Estoque=("Dias_no_Estoque", "max"),
                            Volume_Total=("Saldo 1a.U.M.", "sum"),
                        )
                        .sort_values(by="Capital_Retido", ascending=False)
                        .head(5)
                        .reset_index()
                    )

                    df_top_giro["Capital_Retido"] = df_top_giro["Capital_Retido"].apply(
                        lambda x: formatar_valor_comex(x)
                    )
                    df_top_giro["Volume_Total"] = df_top_giro["Volume_Total"].apply(
                        lambda x: f"{x:,.0f} Un."
                    )
                    df_top_giro.columns = [
                        "Descrição do Produto",
                        "Capital Retido",
                        "Idade Máxima do Lote",
                        "Volume Total em Estoque",
                    ]

                    st.table(df_top_giro)
                else:
                    st.success(
                        "Parabéns! Não existem produtos com baixo giro na seleção atual."
                    )

            elif categoria_analise == "Lotes Vencidos / À Vencer (Risco Financeiro)":
                df_gargalo = df_macro[
                    df_macro["Classificacao_Saude"].isin(["VENCIDO", "À VENCER"])
                ]

                if not df_gargalo.empty:
                    vlr_vencido_macro = df_gargalo[
                        df_gargalo["Classificacao_Saude"] == "VENCIDO"
                    ]["Custo total estoque"].sum()
                    vlr_a_vencer_macro = df_gargalo[
                        df_gargalo["Classificacao_Saude"] == "À VENCER"
                    ]["Custo total estoque"].sum()

                    st.markdown("#### 🩺 Diagnóstico: Risco de Perda Material")

                    c_s1, c_s2 = st.columns(2)
                    c_s1.metric(
                        "Prejuízo Real (Já Vencido)",
                        formatar_valor_comex(vlr_vencido_macro),
                    )
                    c_s2.metric(
                        "Risco Comercial (À Vencer)",
                        formatar_valor_comex(vlr_a_vencer_macro),
                    )

                    st.markdown(
                        "#####  Lotes Críticos Exigindo Ação Comercial Imediata"
                    )
                    df_top_saude = (
                        df_gargalo.sort_values(by="Dias_Restantes", ascending=True)
                        .head(5)[
                            [
                                "Filial",
                                "Descricao",
                                "Lote",
                                "Data Validad",
                                "Dias_Restantes",
                                "Custo total estoque",
                                "Classificacao_Saude",
                            ]
                        ]
                        .copy()
                    )

                    df_top_saude["Data Validad"] = df_top_saude[
                        "Data Validad"
                    ].dt.strftime("%d/%m/%Y")
                    df_top_saude["Custo total estoque"] = df_top_saude[
                        "Custo total estoque"
                    ].apply(lambda x: formatar_valor_comex(x))
                    df_top_saude["Dias_Restantes"] = df_top_saude[
                        "Dias_Restantes"
                    ].apply(
                        lambda x: (
                            f"Vencido há {abs(x)} dias"
                            if x < 0
                            else f"{x} dias restantes"
                        )
                    )
                    df_top_saude.columns = [
                        "Filial",
                        "Descrição",
                        "Número do Lote",
                        "Data de Vencimento",
                        "Janela de Tempo",
                        "Custo do Lote",
                        "Status",
                    ]

                    st.dataframe(
                        df_top_saude, use_container_width=True, hide_index=True
                    )
                else:
                    st.success(
                        "Excelente! Nenhum lote com vencimento crítico identificado na seleção atual."
                    )
        else:
            st.info("Dados indisponíveis para gerar a análise macro dos gráficos.")

    # --------------------------------------------------
    # TAB 4: ESTOQUE EM TRÂNSITO
    # --------------------------------------------------
    with tab_transito:
        st.markdown("###  Estoque em Trânsito")
        st.caption(
            "Visão executiva dos pedidos de compra/importação em andamento, considerando a data de entrega no armazém como disponibilidade do material."
        )

        if df_transito.empty:
            st.warning(
                "Não foram encontrados dados válidos na aba PED. COMPRA. Verifique se as colunas E, G, H, K e T estão preenchidas corretamente."
            )
        else:
            # Filtros específicos do módulo em trânsito
            col_tf1, col_tf2 = st.columns([1, 1])

            opcoes_filial_transito = ["TODAS"] + sorted(
                df_transito["Filial"].dropna().astype(str).unique()
            )
            opcoes_fornecedor_transito = ["TODOS"] + sorted(
                df_transito["Fornecedor"].dropna().astype(str).unique()
            )

            with col_tf1:
                filial_transito = st.multiselect(
                    "Filtrar filial/matriz:",
                    options=opcoes_filial_transito,
                    default=["TODAS"],
                )
            with col_tf2:
                fornecedor_transito = st.multiselect(
                    "Filtrar fornecedor:",
                    options=opcoes_fornecedor_transito,
                    default=["TODOS"],
                )

            df_transito_filtrado = df_transito.copy()

            if not ("TODAS" in filial_transito or not filial_transito):
                df_transito_filtrado = df_transito_filtrado[
                    df_transito_filtrado["Filial"].astype(str).isin(filial_transito)
                ]

            if not ("TODOS" in fornecedor_transito or not fornecedor_transito):
                df_transito_filtrado = df_transito_filtrado[
                    df_transito_filtrado["Fornecedor"]
                    .astype(str)
                    .isin(fornecedor_transito)
                ]

            if df_transito_filtrado.empty:
                st.info("Sem registros para os filtros selecionados.")
            else:
                total_kg_transito = df_transito_filtrado["Quantidade_kg"].sum()
                total_registros_transito = len(df_transito_filtrado)
                total_fornecedores_transito = df_transito_filtrado[
                    "Fornecedor"
                ].nunique()
                proxima_entrega = df_transito_filtrado["Data_Entrega_Armazem"].min()
                dias_proxima_entrega = int((proxima_entrega - DATA_HOJE).days)

                # Mês de maior volume previsto
                volume_mes = (
                    df_transito_filtrado.groupby("Mes_Entrega", as_index=False)[
                        "Quantidade_kg"
                    ]
                    .sum()
                    .sort_values("Mes_Entrega")
                )
                maior_mes = volume_mes.loc[volume_mes["Quantidade_kg"].idxmax()]
                maior_mes_label = maior_mes["Mes_Entrega"].strftime("%m/%Y")
                maior_mes_kg = maior_mes["Quantidade_kg"]

                top_fornecedor = (
                    df_transito_filtrado.groupby("Fornecedor", dropna=False)[
                        "Quantidade_kg"
                    ]
                    .sum()
                    .sort_values(ascending=False)
                )
                nome_top_fornecedor = str(top_fornecedor.index[0])
                perc_top_fornecedor = (
                    (top_fornecedor.iloc[0] / total_kg_transito * 100)
                    if total_kg_transito > 0
                    else 0
                )

                # KPIs executivos
                k1, k2, k3, k4 = st.columns(4)
                with k1:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Total em Trânsito</div><div class="metric-value">{formatar_kg(total_kg_transito)}</div></div>',
                        unsafe_allow_html=True,
                    )
                with k2:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Registros em Aberto</div><div class="metric-value">{total_registros_transito}</div></div>',
                        unsafe_allow_html=True,
                    )
                with k3:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Fornecedores</div><div class="metric-value">{total_fornecedores_transito}</div></div>',
                        unsafe_allow_html=True,
                    )
                with k4:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Próxima Entrega</div><div class="metric-value">{proxima_entrega.strftime("%d/%m/%Y")}</div></div>',
                        unsafe_allow_html=True,
                    )

                st.markdown(
                    f"""
                    <div class="alert-box" style="border-left-color: #00E5FF !important;">
                        <span class="alert-title" style="color: #00E5FF !important;"> Resumo Executivo do Trânsito</span>
                        <span class="alert-text">
                            <strong>{formatar_kg(total_kg_transito)}</strong> de produtos estão em trânsito/importação.<br>
                            O maior volume está previsto para <strong>{maior_mes_label}</strong>, com <strong>{formatar_kg(maior_mes_kg)}</strong>.<br>
                            O fornecedor <strong>{nome_top_fornecedor}</strong> concentra <strong>{perc_top_fornecedor:.1f}%</strong> do volume em trânsito.<br>
                            A próxima entrega está prevista para <strong>{proxima_entrega.strftime("%d/%m/%Y")}</strong> ({dias_proxima_entrega} dias a partir da data de referência).
                        </span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                st.markdown("####  Entrada prevista por mês")

                volume_mes["Mes_Label"] = volume_mes["Mes_Entrega"].dt.strftime("%m/%Y")
                fig_mes = px.bar(
                    volume_mes,
                    x="Mes_Label",
                    y="Quantidade_kg",
                    text="Quantidade_kg",
                    labels={
                        "Mes_Label": "Mês de entrega",
                        "Quantidade_kg": "Quantidade em trânsito (kg)",
                    },
                    title="Volume em trânsito por mês de entrega no armazém",
                )
                fig_mes.update_traces(
                    marker_color="#4B2A85",
                    texttemplate="%{text:,.0f} kg",
                    textposition="outside",
                    hovertemplate="<b>%{x}</b><br>Quantidade: %{y:,.2f} kg<extra></extra>",
                )
                fig_mes.update_layout(
                    height=430,
                    showlegend=False,
                    margin=dict(t=60, b=40, l=40, r=40),
                    yaxis_title="Quantidade (kg)",
                    xaxis_title="Mês de entrega",
                )
                st.plotly_chart(fig_mes, use_container_width=True)

                g_for, g_filial = st.columns(2)

                with g_for:
                    st.markdown("####  Top fornecedores por volume em trânsito")
                    dados_fornecedor = (
                        df_transito_filtrado.groupby(
                            "Fornecedor", dropna=False, as_index=False
                        )["Quantidade_kg"]
                        .sum()
                        .sort_values("Quantidade_kg", ascending=False)
                        .head(10)
                    )
                    fig_fornecedor = px.bar(
                        dados_fornecedor,
                        x="Quantidade_kg",
                        y="Fornecedor",
                        orientation="h",
                        text="Quantidade_kg",
                        labels={
                            "Quantidade_kg": "Quantidade (kg)",
                            "Fornecedor": "Fornecedor",
                        },
                    )
                    fig_fornecedor.update_traces(
                        marker_color="#5A339B",
                        texttemplate="%{text:,.0f} kg",
                        textposition="outside",
                        hovertemplate="<b>%{y}</b><br>Quantidade: %{x:,.2f} kg<extra></extra>",
                    )
                    fig_fornecedor.update_layout(
                        height=430,
                        yaxis={"categoryorder": "total ascending"},
                        margin=dict(t=20, b=40, l=20, r=40),
                    )
                    st.plotly_chart(fig_fornecedor, use_container_width=True)

                with g_filial:
                    st.markdown("####  Distribuição por filial/matriz")
                    dados_filial = (
                        df_transito_filtrado.groupby(
                            "Filial", dropna=False, as_index=False
                        )["Quantidade_kg"]
                        .sum()
                        .sort_values("Quantidade_kg", ascending=False)
                    )
                    fig_filial = px.pie(
                        dados_filial,
                        names="Filial",
                        values="Quantidade_kg",
                        hole=0.45,
                        color_discrete_sequence=[
                            "#4B2A85",
                            "#5A339B",
                            "#00E5FF",
                            "#FFB800",
                            "#E91010",
                        ],
                    )
                    fig_filial.update_traces(
                        textinfo="percent+label",
                        hovertemplate="<b>%{label}</b><br>Quantidade: %{value:,.2f} kg<extra></extra>",
                    )
                    fig_filial.update_layout(
                        height=430, margin=dict(t=20, b=40, l=20, r=20)
                    )
                    st.plotly_chart(fig_filial, use_container_width=True)

                st.markdown("####  Detalhamento do estoque em trânsito")
                tabela_transito = df_transito_filtrado[
                    [
                        "Data_Entrega_Armazem",
                        "Filial",
                        "Produto",
                        "Fornecedor",
                        "Quantidade_kg",
                        "Dias_Para_Entrega",
                    ]
                ].copy()

                tabela_transito = tabela_transito.sort_values("Data_Entrega_Armazem")
                tabela_transito["Data_Entrega_Armazem"] = tabela_transito[
                    "Data_Entrega_Armazem"
                ].dt.strftime("%d/%m/%Y")
                tabela_transito["Quantidade_kg"] = tabela_transito["Quantidade_kg"].map(
                    lambda x: f"{x:,.2f} kg"
                )
                tabela_transito["Dias_Para_Entrega"] = tabela_transito[
                    "Dias_Para_Entrega"
                ].map(
                    lambda x: (
                        f"Entrega vencida há {abs(int(x))} dias"
                        if x < 0
                        else f"{int(x)} dias"
                    )
                )
                tabela_transito.columns = [
                    "Entrega no Armazém",
                    "Filial/Matriz",
                    "Produto",
                    "Fornecedor",
                    "Quantidade",
                    "Dias até Entrega",
                ]

                st.dataframe(tabela_transito, use_container_width=True, hide_index=True)

except Exception as erro:
    st.error(f"Erro ao processar e estruturar o painel de COMEX: {erro}")
