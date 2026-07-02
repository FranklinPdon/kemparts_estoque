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
    df = pd.read_excel(io.BytesIO(file_resp.content), sheet_name="ESTOQUE", header=1)
    return df


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
        df = carregar_dados_sharepoint()

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
    tab_visao_geral, tab_detalhes_lote, tab_graficos = st.tabs(
        [
            " VISÃO GERAL DO PRODUTO",
            " TODOS OS LOTES DESTE PRODUTO",
            " ANÁLISE GRÁFICA MACRO",
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
                    "VALIDADE OK": "#08D81A",
                    "À VENCER": "#E9DA0D",
                    "VENCIDO": "#B71C1C",
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

except Exception as erro:
    st.error(f"Erro ao processar e estruturar o painel de COMEX: {erro}")
