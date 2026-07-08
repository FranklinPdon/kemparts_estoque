from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import io
import hmac
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

    # Tenta localizar automaticamente uma coluna de container/processo na aba PED. COMPRA.
    # Se a planilha ainda não tiver essa informação, o painel cria uma referência executiva
    # agrupando fornecedor + filial + data de entrega.
    coluna_ref_container = None
    palavras_chave_container = [
        "container",
        "cntr",
        "processo",
        "embarque",
        "invoice",
        "pedido",
        "bl",
        "referencia",
        "referência",
    ]

    for coluna in df_transito.columns:
        nome_coluna = str(coluna).strip().lower()
        if any(palavra in nome_coluna for palavra in palavras_chave_container):
            coluna_ref_container = coluna
            break

    colunas_base = [
        "Filial",
        "Produto",
        "Quantidade_kg",
        "Fornecedor",
        "Data_Entrega_Armazem",
    ]

    if coluna_ref_container is not None:
        df_transito["Container_Ref"] = df_transito[coluna_ref_container].astype(str)
        colunas_base.append("Container_Ref")
    else:
        df_transito["Container_Ref"] = ""
        colunas_base.append("Container_Ref")

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

    # Caso não exista uma coluna real de container/processo, cria uma referência consolidada
    # para permitir uma visão gerencial por "embarque previsto".
    sem_ref = (
        df_transito["Container_Ref"].isna()
        | (df_transito["Container_Ref"].astype(str).str.strip() == "")
        | (df_transito["Container_Ref"].astype(str).str.lower().isin(["nan", "none"]))
    )

    df_transito.loc[sem_ref, "Container_Ref"] = (
        "EMB-"
        + df_transito.loc[sem_ref, "Fornecedor"].astype(str).str[:12].str.upper()
        + "-"
        + df_transito.loc[sem_ref, "Filial"].astype(str).str[:6].str.upper()
        + "-"
        + df_transito.loc[sem_ref, "Data_Entrega_Armazem"].dt.strftime("%Y%m%d")
    )

    def classificar_status_container(dias):
        if dias < 0:
            return "🔴 Atrasado"
        elif dias <= 15:
            return "🟢 Chega em até 15 dias"
        elif dias <= 45:
            return "🟡 Chega em 16 a 45 dias"
        else:
            return "🔵 Futuro"

    df_transito["Status_Entrega"] = df_transito["Dias_Para_Entrega"].apply(
        classificar_status_container
    )

    return df_transito


# ======================================================
# FUNÇÕES AUXILIARES - CUSTOS E PREVISÃO DE ESTOQUE
# ======================================================
def formatar_percentual(valor):
    try:
        return f"{float(valor):.1f}%"
    except (TypeError, ValueError):
        return "0,0%"


def preparar_historico_custos(df_hist_venda):
    """
    Prepara a aba Hist. Venda para cálculo de custo médio por KG mês a mês.

    Mapeamento solicitado:
    - Coluna AM: custo por KG
    - Coluna F: data de emissão da nota fiscal
    - Coluna Q: quantidade
    - Coluna M: descrição do produto
    """
    if df_hist_venda is None or df_hist_venda.empty:
        return pd.DataFrame()

    df_hist = df_hist_venda.copy().dropna(how="all")

    posicoes = {
        "Data_Emissao_NF": 5,  # Coluna F
        "Produto": 12,  # Coluna M
        "Quantidade_kg": 16,  # Coluna Q
        "Custo_KG": 38,  # Coluna AM
    }

    if len(df_hist.columns) <= max(posicoes.values()):
        return pd.DataFrame()

    df_hist = df_hist.rename(
        columns={df_hist.columns[pos]: nome for nome, pos in posicoes.items()}
    )

    df_hist = df_hist[
        list(posicoes.values())
        if False
        else ["Data_Emissao_NF", "Produto", "Quantidade_kg", "Custo_KG"]
    ].copy()

    df_hist["Data_Emissao_NF"] = pd.to_datetime(
        df_hist["Data_Emissao_NF"], errors="coerce", dayfirst=True
    )
    df_hist["Quantidade_kg"] = (
        df_hist["Quantidade_kg"].apply(converter_numero).fillna(0)
    )
    df_hist["Custo_KG"] = df_hist["Custo_KG"].apply(converter_numero).fillna(0)

    df_hist = df_hist.dropna(subset=["Data_Emissao_NF", "Produto"])
    df_hist = df_hist[(df_hist["Quantidade_kg"] > 0) & (df_hist["Custo_KG"] > 0)]

    if df_hist.empty:
        return df_hist

    df_hist["Mes"] = df_hist["Data_Emissao_NF"].dt.to_period("M").dt.to_timestamp()
    df_hist["Custo_Total_Calculado"] = df_hist["Quantidade_kg"] * df_hist["Custo_KG"]

    return df_hist


def preparar_previsao_custos(df_previsao_custo):
    """
    Prepara a aba Previsão de custo para projeção de custo futuro.

    Mapeamento solicitado:
    - Coluna H: descrição do produto
    - Coluna AO: custo por KG
    - Coluna AP: mês esperado de entrega
    """
    if df_previsao_custo is None or df_previsao_custo.empty:
        return pd.DataFrame()

    df_prev = df_previsao_custo.copy().dropna(how="all")

    posicoes = {
        "Produto": 7,  # Coluna H
        "Custo_KG_Projetado": 40,  # Coluna AO
        "Mes_Entrega": 41,  # Coluna AP
    }

    if len(df_prev.columns) <= max(posicoes.values()):
        return pd.DataFrame()

    df_prev = df_prev.rename(
        columns={df_prev.columns[pos]: nome for nome, pos in posicoes.items()}
    )

    df_prev = df_prev[["Produto", "Custo_KG_Projetado", "Mes_Entrega"]].copy()
    df_prev["Mes_Entrega"] = pd.to_datetime(
        df_prev["Mes_Entrega"], errors="coerce", dayfirst=True
    )
    df_prev["Custo_KG_Projetado"] = (
        df_prev["Custo_KG_Projetado"].apply(converter_numero).fillna(0)
    )

    df_prev = df_prev.dropna(subset=["Produto", "Mes_Entrega"])
    df_prev = df_prev[df_prev["Custo_KG_Projetado"] > 0]

    if df_prev.empty:
        return df_prev

    df_prev["Mes"] = df_prev["Mes_Entrega"].dt.to_period("M").dt.to_timestamp()

    return df_prev


def consolidar_custo_medio_historico(df_hist_custos):
    if df_hist_custos is None or df_hist_custos.empty:
        return pd.DataFrame()

    df_consolidado = df_hist_custos.groupby(["Produto", "Mes"], as_index=False).agg(
        Quantidade_kg=("Quantidade_kg", "sum"),
        Custo_Total_Calculado=("Custo_Total_Calculado", "sum"),
    )
    df_consolidado["Custo_KG"] = (
        df_consolidado["Custo_Total_Calculado"] / df_consolidado["Quantidade_kg"]
    )
    df_consolidado["Tipo"] = "Histórico"
    return df_consolidado


def consolidar_custo_medio_projetado(df_previsao_custos):
    if df_previsao_custos is None or df_previsao_custos.empty:
        return pd.DataFrame()

    df_proj = df_previsao_custos.groupby(["Produto", "Mes"], as_index=False).agg(
        Custo_KG=("Custo_KG_Projetado", "mean")
    )
    df_proj["Tipo"] = "Projeção"
    return df_proj


def normalizar_produtos_visao(produtos_selecionados):
    """
    Padroniza a seleção da visão executiva.

    Regra importante:
    - Se apenas TODOS estiver selecionado, usa visão macro.
    - Se TODOS + produto(s) estiverem selecionados, o sistema ignora TODOS
      e considera somente os produtos específicos.
    Isso evita que o botão Ver mais mostre Top 5 geral quando o usuário
    já escolheu um produto específico.
    """
    if produtos_selecionados is None:
        return ["TODOS"]

    if isinstance(produtos_selecionados, str):
        produtos = [produtos_selecionados]
    else:
        produtos = list(produtos_selecionados)

    produtos = [str(p).strip() for p in produtos if str(p).strip() != ""]

    if not produtos:
        return ["TODOS"]

    produtos_especificos = [p for p in produtos if p.upper() != "TODOS"]

    if produtos_especificos:
        return produtos_especificos

    return ["TODOS"]


def visao_todos_produtos(produtos_selecionados):
    produtos = normalizar_produtos_visao(produtos_selecionados)
    return len(produtos) == 0 or (
        len(produtos) == 1 and str(produtos[0]).upper() == "TODOS"
    )


def montar_previsao_estoque_ate_fim_ano(
    df_estoque_base, df_transito_base, produtos_selecionados
):
    """Monta a previsão mensal de estoque até dezembro, usando estoque atual + entradas em trânsito."""
    produtos = normalizar_produtos_visao(produtos_selecionados)
    todos_produtos = visao_todos_produtos(produtos)

    estoque_base = df_estoque_base.copy()
    transito_base = (
        df_transito_base.copy() if df_transito_base is not None else pd.DataFrame()
    )

    if not todos_produtos:
        estoque_base = estoque_base[
            estoque_base["Descricao"].astype(str).isin([str(p) for p in produtos])
        ]
        if not transito_base.empty:
            transito_base = transito_base[
                transito_base["Produto"].astype(str).isin([str(p) for p in produtos])
            ]

    estoque_atual_kg = (
        estoque_base["Saldo 1a.U.M."].sum() if not estoque_base.empty else 0
    )

    ano_ref = int(DATA_HOJE.year)
    mes_inicial = pd.Timestamp(DATA_HOJE.year, DATA_HOJE.month, 1)
    meses = pd.date_range(
        start=mes_inicial, end=pd.Timestamp(ano_ref, 12, 1), freq="MS"
    )
    df_previsao = pd.DataFrame({"Mes": meses})

    if not transito_base.empty:
        entradas_mes = (
            transito_base.groupby("Mes_Entrega", as_index=False)["Quantidade_kg"]
            .sum()
            .rename(
                columns={"Mes_Entrega": "Mes", "Quantidade_kg": "Entrada_Prevista_kg"}
            )
        )
    else:
        entradas_mes = pd.DataFrame(columns=["Mes", "Entrada_Prevista_kg"])

    df_previsao = df_previsao.merge(entradas_mes, on="Mes", how="left")
    df_previsao["Entrada_Prevista_kg"] = df_previsao["Entrada_Prevista_kg"].fillna(0)
    df_previsao["Estoque_Projetado_kg"] = (
        estoque_atual_kg + df_previsao["Entrada_Prevista_kg"].cumsum()
    )
    df_previsao["Mes_Label"] = df_previsao["Mes"].dt.strftime("%m/%Y")

    return estoque_atual_kg, df_previsao


def filtrar_transito_por_produto(df_transito_base, produtos_selecionados):
    """Filtra o trânsito pelos produtos selecionados na visão executiva."""
    if df_transito_base is None or df_transito_base.empty:
        return pd.DataFrame()
    produtos = normalizar_produtos_visao(produtos_selecionados)
    if visao_todos_produtos(produtos):
        return df_transito_base.copy()
    return df_transito_base[
        df_transito_base["Produto"].astype(str).isin([str(p) for p in produtos])
    ].copy()


def tabela_estoque_atual_detalhe(df_estoque_base, produtos_selecionados):
    """Detalhe inteligente do card Estoque Atual: macro em TODOS, detalhe quando há produto filtrado."""
    if df_estoque_base is None or df_estoque_base.empty:
        return pd.DataFrame(), "Sem dados para detalhar."

    produtos = normalizar_produtos_visao(produtos_selecionados)
    if visao_todos_produtos(produtos):
        tabela = (
            df_estoque_base.groupby("Descricao", as_index=False)["Saldo 1a.U.M."]
            .sum()
            .sort_values("Saldo 1a.U.M.", ascending=False)
            .head(5)
        )
        tabela.columns = ["Produto", "Estoque Atual"]
        tabela["Estoque Atual"] = tabela["Estoque Atual"].map(lambda x: formatar_kg(x))
        return tabela, "Top 5 produtos por volume em estoque"

    base = df_estoque_base[
        df_estoque_base["Descricao"].astype(str).isin([str(p) for p in produtos])
    ].copy()
    if base.empty:
        return pd.DataFrame(), "Sem dados para os produtos selecionados."

    tabela = (
        base.groupby("Descricao", as_index=False)
        .agg(
            Estoque_Atual=("Saldo 1a.U.M.", "sum"),
            Lotes=("Lote", "nunique"),
            Fornecedores=("Fornecedor", "nunique"),
            Validade_Mais_Proxima=("Data Validad", "min"),
        )
        .sort_values("Estoque_Atual", ascending=False)
    )
    tabela["Estoque_Atual"] = tabela["Estoque_Atual"].map(lambda x: formatar_kg(x))
    tabela["Validade_Mais_Proxima"] = pd.to_datetime(
        tabela["Validade_Mais_Proxima"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")
    tabela.columns = [
        "Produto",
        "Estoque atual",
        "Lotes",
        "Fornecedores",
        "Validade mais próxima",
    ]
    return tabela, "Resumo dos produtos selecionados"


def tabela_entradas_detalhe(df_transito_base, produtos_selecionados):
    """Detalhe inteligente do card Entradas até Dezembro."""
    if df_transito_base is None or df_transito_base.empty:
        return pd.DataFrame(), "Sem entradas previstas até dezembro."

    produtos = normalizar_produtos_visao(produtos_selecionados)
    ano_ref = int(DATA_HOJE.year)
    limite = pd.Timestamp(ano_ref, 12, 31)
    base = df_transito_base[df_transito_base["Data_Entrega_Armazem"] <= limite].copy()

    if not visao_todos_produtos(produtos):
        base = base[base["Produto"].astype(str).isin([str(p) for p in produtos])]

    if base.empty:
        return (
            pd.DataFrame(),
            "Sem entradas previstas até dezembro para a seleção atual.",
        )

    if visao_todos_produtos(produtos):
        tabela = (
            base.groupby("Produto", as_index=False)["Quantidade_kg"]
            .sum()
            .sort_values("Quantidade_kg", ascending=False)
            .head(5)
        )
        tabela.columns = ["Produto", "Entrada Prevista"]
        tabela["Entrada Prevista"] = tabela["Entrada Prevista"].map(
            lambda x: formatar_kg(x)
        )
        return tabela, "Top 5 produtos com maior entrada prevista"

    tabela = (
        base.groupby(["Produto", "Mes_Entrega"], as_index=False)["Quantidade_kg"]
        .sum()
        .sort_values(["Mes_Entrega", "Quantidade_kg"], ascending=[True, False])
    )
    tabela["Mês"] = tabela["Mes_Entrega"].dt.strftime("%m/%Y")
    tabela["Quantidade"] = tabela["Quantidade_kg"].map(lambda x: formatar_kg(x))
    tabela = tabela[["Mês", "Produto", "Quantidade"]]
    return tabela, "Entradas previstas dos produtos selecionados"


def tabela_previsao_fim_ano_detalhe(
    df_estoque_base, df_transito_base, produtos_selecionados, df_previsao_estoque
):
    """Detalhe inteligente do card Previsão Fim do Ano."""
    produtos = normalizar_produtos_visao(produtos_selecionados)

    if visao_todos_produtos(produtos):
        tabela = tabela_top_previsao_fim_ano(df_estoque_base, df_transito_base)
        return tabela, "Top 5 produtos por estoque projetado até dezembro"

    if df_previsao_estoque is None or df_previsao_estoque.empty:
        return pd.DataFrame(), "Sem previsão para os produtos selecionados."

    tabela = df_previsao_estoque[
        ["Mes_Label", "Entrada_Prevista_kg", "Estoque_Projetado_kg"]
    ].copy()
    tabela["Entrada_Prevista_kg"] = tabela["Entrada_Prevista_kg"].map(
        lambda x: formatar_kg(x)
    )
    tabela["Estoque_Projetado_kg"] = tabela["Estoque_Projetado_kg"].map(
        lambda x: formatar_kg(x)
    )
    tabela.columns = ["Mês", "Entrada prevista", "Estoque projetado"]
    return tabela, "Evolução mensal da seleção atual"


def tabela_proximas_entradas_detalhe(df_transito_base, produtos_selecionados):
    """Detalhe inteligente do card Próxima Entrada."""
    if df_transito_base is None or df_transito_base.empty:
        return pd.DataFrame(), "Sem próximas entradas previstas."

    produtos = normalizar_produtos_visao(produtos_selecionados)
    base = df_transito_base.copy()
    if not visao_todos_produtos(produtos):
        base = base[base["Produto"].astype(str).isin([str(p) for p in produtos])]

    if base.empty:
        return pd.DataFrame(), "Sem próximas entradas para a seleção atual."

    base = base.sort_values("Data_Entrega_Armazem").head(5)
    tabela = base[
        ["Data_Entrega_Armazem", "Produto", "Fornecedor", "Quantidade_kg"]
    ].copy()
    tabela["Data_Entrega_Armazem"] = tabela["Data_Entrega_Armazem"].dt.strftime(
        "%d/%m/%Y"
    )
    tabela["Quantidade_kg"] = tabela["Quantidade_kg"].map(lambda x: formatar_kg(x))
    tabela.columns = ["Entrega", "Produto", "Fornecedor", "Quantidade"]
    titulo = (
        "Próximas 5 entradas previstas"
        if visao_todos_produtos(produtos)
        else "Próximas entradas da seleção atual"
    )
    return tabela, titulo


def tabela_top_estoque_atual(df_estoque_base):
    """Retorna os 5 principais produtos por volume atual em estoque."""
    tabela, _ = tabela_estoque_atual_detalhe(df_estoque_base, ["TODOS"])
    return tabela


def tabela_top_entradas(df_transito_base):
    """Retorna os 5 principais produtos por volume de entrada prevista."""
    tabela, _ = tabela_entradas_detalhe(df_transito_base, ["TODOS"])
    return tabela


def tabela_top_previsao_fim_ano(df_estoque_base, df_transito_base):
    """Retorna os 5 principais produtos por estoque projetado até dezembro."""
    if df_estoque_base is None or df_estoque_base.empty:
        estoque_prod = pd.DataFrame(columns=["Produto", "Estoque_Atual"])
    else:
        estoque_prod = (
            df_estoque_base.groupby("Descricao", as_index=False)["Saldo 1a.U.M."]
            .sum()
            .rename(columns={"Descricao": "Produto", "Saldo 1a.U.M.": "Estoque_Atual"})
        )

    if df_transito_base is None or df_transito_base.empty:
        entradas_prod = pd.DataFrame(columns=["Produto", "Entradas"])
    else:
        ano_ref = int(DATA_HOJE.year)
        limite = pd.Timestamp(ano_ref, 12, 31)
        base = df_transito_base[
            df_transito_base["Data_Entrega_Armazem"] <= limite
        ].copy()
        entradas_prod = (
            base.groupby("Produto", as_index=False)["Quantidade_kg"]
            .sum()
            .rename(columns={"Quantidade_kg": "Entradas"})
            if not base.empty
            else pd.DataFrame(columns=["Produto", "Entradas"])
        )

    tabela = estoque_prod.merge(entradas_prod, on="Produto", how="outer").fillna(0)
    if tabela.empty:
        return pd.DataFrame()
    tabela["Previsao_Fim_Ano"] = tabela["Estoque_Atual"] + tabela["Entradas"]
    tabela = tabela.sort_values("Previsao_Fim_Ano", ascending=False).head(5)
    tabela = tabela[["Produto", "Previsao_Fim_Ano"]]
    tabela.columns = ["Produto", "Previsão Fim do Ano"]
    tabela["Previsão Fim do Ano"] = tabela["Previsão Fim do Ano"].map(
        lambda x: formatar_kg(x)
    )
    return tabela


def tabela_proximas_entradas(df_transito_base):
    """Retorna as 5 próximas entradas previstas."""
    tabela, _ = tabela_proximas_entradas_detalhe(df_transito_base, ["TODOS"])
    return tabela


# ======================================================
# CONFIGURAÇÃO SHAREPOINT (EDITE AQUI SE NECESSÁRIO)
# ======================================================
SHAREPOINT_HOST = "kempartsquimica.sharepoint.com"
SITE_PATH = "/sites/IMPORTACAO"
FILE_PATH = "/BASE DASHBOARD/PLANEJAMENTO DE COMPRAS 30.06.2026_FRANKLIN.xlsx"
USERS_FILE_PATH = "/BASE DASHBOARD/USUARIOS_ESTOQUEVISION.xlsx"


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


# ======================================================
# CONTROLE DE ACESSO - LOGIN E PERFIS
# ======================================================
def normalizar_sim_nao(valor):
    texto = str(valor).strip().upper()
    return texto in ["SIM", "S", "YES", "Y", "TRUE", "1"]


def preparar_usuarios(df_usuarios):
    """Padroniza a planilha USUARIOS_ESTOQUEVISION.xlsx."""
    if df_usuarios is None or df_usuarios.empty:
        return pd.DataFrame()

    dfu = df_usuarios.copy()
    dfu.columns = [str(c).strip().upper().replace(" ", "_") for c in dfu.columns]

    colunas_obrigatorias = ["LOGIN", "NOME", "SENHA", "PERFIL", "ATIVO"]
    for col in colunas_obrigatorias:
        if col not in dfu.columns:
            raise Exception(
                f"Coluna obrigatória ausente na planilha de usuários: {col}"
            )

    dfu["LOGIN"] = dfu["LOGIN"].astype(str).str.strip().str.lower()
    dfu["NOME"] = dfu["NOME"].astype(str).str.strip()
    dfu["SENHA"] = dfu["SENHA"].astype(str)
    dfu["PERFIL"] = dfu["PERFIL"].astype(str).str.strip().str.upper()
    dfu["ATIVO"] = dfu["ATIVO"].apply(normalizar_sim_nao)

    # Se a planilha tiver uma coluna explícita de permissão, ela prevalece.
    # Caso contrário, o sistema libera custos para ADMIN, DIRETORIA, SUPPLY e FINANCEIRO.
    if "PODE_VER_CUSTOS" in dfu.columns:
        dfu["PODE_VER_CUSTOS"] = dfu["PODE_VER_CUSTOS"].apply(normalizar_sim_nao)
    elif "PODE_VER_CUSTO" in dfu.columns:
        dfu["PODE_VER_CUSTOS"] = dfu["PODE_VER_CUSTO"].apply(normalizar_sim_nao)
    else:
        perfis_autorizados = [
            "ADMIN",
            "ADMINISTRADOR",
            "DIRETORIA",
            "SUPPLY",
            "SUPPLY CHAIN",
            "SUPPLY CHEN",
            "FINANCEIRO",
        ]
        dfu["PODE_VER_CUSTOS"] = dfu["PERFIL"].isin(perfis_autorizados)

    return dfu


@st.cache_data(ttl=3600)
def carregar_usuarios_sharepoint():
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOST}:{SITE_PATH}"
    site_resp = requests.get(site_url, headers=headers)
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]

    file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{USERS_FILE_PATH}:/content"
    file_resp = requests.get(file_url, headers=headers)
    file_resp.raise_for_status()

    df_usuarios = pd.read_excel(io.BytesIO(file_resp.content))
    return preparar_usuarios(df_usuarios)


def autenticar_usuario(login, senha, df_usuarios):
    login = str(login).strip().lower()
    senha = str(senha)

    if df_usuarios is None or df_usuarios.empty:
        return None

    usuario = df_usuarios[
        (df_usuarios["LOGIN"] == login) & (df_usuarios["ATIVO"] == True)
    ]
    if usuario.empty:
        return None

    registro = usuario.iloc[0]
    senha_planilha = str(registro["SENHA"])

    if not hmac.compare_digest(senha, senha_planilha):
        return None

    return {
        "login": registro["LOGIN"],
        "nome": registro["NOME"],
        "perfil": registro["PERFIL"],
        "pode_ver_custos": bool(registro["PODE_VER_CUSTOS"]),
    }


def tela_login():
    """Exibe tela de login antes do painel."""
    st.markdown("## 🔐 Acesso ao ESTOQUEVISION")
    st.caption(
        "Informe seu login e senha para acessar o painel conforme seu perfil de permissão."
    )

    with st.form("form_login"):
        login = st.text_input("Login / e-mail")
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar")

    if entrar:
        try:
            df_usuarios = carregar_usuarios_sharepoint()
            usuario = autenticar_usuario(login, senha, df_usuarios)

            if usuario is None:
                st.error("Login ou senha inválidos, ou usuário inativo.")
            else:
                st.session_state["usuario_logado"] = usuario
                st.rerun()
        except Exception as erro_login:
            st.error(f"Erro ao validar acesso: {erro_login}")

    st.stop()


def barra_usuario_logado(usuario):
    col_user, col_sair = st.columns([5, 1])
    with col_user:
        st.caption(f"👋 Bem-vindo, **{usuario.get('nome')}**")
    with col_sair:
        if st.button("Sair"):
            st.session_state.pop("usuario_logado", None)
            st.rerun()


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

    # A aba Hist. Venda alimenta o custo médio histórico mês a mês.
    try:
        arquivo_excel.seek(0)
        df_hist_venda = pd.read_excel(arquivo_excel, sheet_name="Hist. Venda", header=1)
    except Exception:
        df_hist_venda = pd.DataFrame()

    # A aba Previsão de custo alimenta a projeção futura de custo por KG.
    try:
        arquivo_excel.seek(0)
        df_previsao_custo = pd.read_excel(
            arquivo_excel, sheet_name="Previsão de custo", header=1
        )
    except Exception:
        df_previsao_custo = pd.DataFrame()

    return df_estoque, df_ped_compra, df_hist_venda, df_previsao_custo


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
# LOGIN DO SISTEMA
# ======================================================
if "usuario_logado" not in st.session_state:
    tela_login()

usuario_logado = st.session_state["usuario_logado"]
pode_ver_custos = usuario_logado.get("pode_ver_custos", False)
barra_usuario_logado(usuario_logado)
st.divider()

# ======================================================
# LEITURA E TRATAMENTO DE DADOS
# ======================================================
try:
    with st.spinner("Carregando dados do SharePoint..."):
        df, df_ped_compra, df_hist_venda, df_previsao_custo = (
            carregar_dados_sharepoint()
        )

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
    # TRATAMENTO DAS ABAS DE CUSTOS
    # ======================================================
    df_hist_custos = preparar_historico_custos(df_hist_venda)
    df_previsao_custos = preparar_previsao_custos(df_previsao_custo)
    df_custo_hist_mensal = consolidar_custo_medio_historico(df_hist_custos)
    df_custo_proj_mensal = consolidar_custo_medio_projetado(df_previsao_custos)

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
    # VISÃO EXECUTIVA - TOTAL DE ESTOQUE E PREVISÃO ATÉ O FIM DO ANO
    # ======================================================
    st.markdown("###  Visão Executiva de Estoque")
    st.caption(
        "Resumo solicitado pela diretoria: estoque total atual, filtro por produto e previsão mensal até o fim do ano."
    )

    opcoes_produto_visao = ["TODOS"] + lista_produtos_base
    produto_visao = st.multiselect(
        "Selecione um ou mais produtos para analisar o estoque total e a previsão até o fim do ano:",
        options=opcoes_produto_visao,
        default=["TODOS"],
        key="produto_visao_executiva",
    )

    # Regra de leitura do detalhe:
    # - TODOS selecionado: visão macro para diretoria.
    # - Produtos específicos: visão focada apenas nos itens selecionados.
    produtos_visao_normalizados = normalizar_produtos_visao(produto_visao)
    visao_macro_todos = visao_todos_produtos(produtos_visao_normalizados)

    estoque_atual_kg, df_previsao_estoque = montar_previsao_estoque_ate_fim_ano(
        df_filtrado, df_transito, produtos_visao_normalizados
    )

    entrada_total_ate_fim_ano = df_previsao_estoque["Entrada_Prevista_kg"].sum()
    estoque_proj_fim_ano = (
        df_previsao_estoque["Estoque_Projetado_kg"].iloc[-1]
        if not df_previsao_estoque.empty
        else estoque_atual_kg
    )
    meses_com_entrada = df_previsao_estoque[
        df_previsao_estoque["Entrada_Prevista_kg"] > 0
    ]
    prox_entrada_label = (
        meses_com_entrada.iloc[0]["Mes_Label"]
        if not meses_com_entrada.empty
        else "Sem previsão"
    )

    df_transito_visao = filtrar_transito_por_produto(
        df_transito, produtos_visao_normalizados
    )

    ve1, ve2, ve3, ve4 = st.columns(4)
    with ve1:
        st.markdown(
            f'<div class="metric-card"><div class="metric-title">Estoque Atual</div><div class="metric-value">{formatar_kg(estoque_atual_kg)}</div></div>',
            unsafe_allow_html=True,
        )
        with st.popover("Ver mais"):
            tabela, titulo = tabela_estoque_atual_detalhe(
                df_filtrado, produtos_visao_normalizados
            )
            st.markdown(f"**{titulo}**")
            if tabela.empty:
                st.info(titulo)
            else:
                st.dataframe(tabela, use_container_width=True, hide_index=True)

    with ve2:
        st.markdown(
            f'<div class="metric-card"><div class="metric-title">Entradas até Dezembro</div><div class="metric-value">{formatar_kg(entrada_total_ate_fim_ano)}</div></div>',
            unsafe_allow_html=True,
        )
        with st.popover("Ver mais"):
            tabela, titulo = tabela_entradas_detalhe(
                df_transito, produtos_visao_normalizados
            )
            st.markdown(f"**{titulo}**")
            if tabela.empty:
                st.info(titulo)
            else:
                st.dataframe(tabela, use_container_width=True, hide_index=True)

    with ve3:
        st.markdown(
            f'<div class="metric-card"><div class="metric-title">Previsão Fim do Ano</div><div class="metric-value">{formatar_kg(estoque_proj_fim_ano)}</div></div>',
            unsafe_allow_html=True,
        )
        with st.popover("Ver mais"):
            tabela, titulo = tabela_previsao_fim_ano_detalhe(
                df_filtrado,
                df_transito,
                produtos_visao_normalizados,
                df_previsao_estoque,
            )
            st.markdown(f"**{titulo}**")
            if tabela.empty:
                st.info(titulo)
            else:
                st.dataframe(tabela, use_container_width=True, hide_index=True)

    with ve4:
        st.markdown(
            f'<div class="metric-card"><div class="metric-title">Próxima Entrada</div><div class="metric-value">{prox_entrada_label}</div></div>',
            unsafe_allow_html=True,
        )
        with st.popover("Ver mais"):
            tabela, titulo = tabela_proximas_entradas_detalhe(
                df_transito, produtos_visao_normalizados
            )
            st.markdown(f"**{titulo}**")
            if tabela.empty:
                st.info(titulo)
            else:
                st.dataframe(tabela, use_container_width=True, hide_index=True)

    if not df_previsao_estoque.empty:
        fig_prev_estoque = px.line(
            df_previsao_estoque,
            x="Mes_Label",
            y="Estoque_Projetado_kg",
            markers=True,
            text="Estoque_Projetado_kg",
            labels={
                "Mes_Label": "Mês",
                "Estoque_Projetado_kg": "Estoque projetado (kg)",
            },
            title="Previsão de estoque até o fim do ano",
        )
        fig_prev_estoque.update_traces(
            line=dict(width=4),
            texttemplate="%{text:,.0f} kg",
            textposition="top center",
            hovertemplate="<b>%{x}</b><br>Estoque projetado: %{y:,.2f} kg<extra></extra>",
        )
        fig_prev_estoque.update_layout(
            height=380,
            margin=dict(t=60, b=40, l=40, r=40),
            yaxis_title="Estoque projetado (kg)",
            xaxis_title="Mês",
        )
        st.plotly_chart(fig_prev_estoque, use_container_width=True)

    st.divider()

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
                    {f"<strong>Capital Imobilizado:</strong> {custo_antigo_formatado}" if pode_ver_custos else ""}
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
    abas_principais = [
        " VISÃO GERAL DO PRODUTO",
        " TODOS OS LOTES DESTE PRODUTO",
        " ANÁLISE GRÁFICA MACRO",
        " ESTOQUE EM TRÂNSITO",
        " CONTAINERS",
    ]

    if pode_ver_custos:
        abas_principais.append(" CUSTOS")

    abas_renderizadas = st.tabs(abas_principais)
    tab_visao_geral = abas_renderizadas[0]
    tab_detalhes_lote = abas_renderizadas[1]
    tab_graficos = abas_renderizadas[2]
    tab_transito = abas_renderizadas[3]
    tab_containers = abas_renderizadas[4]

    if pode_ver_custos:
        tab_custos = abas_renderizadas[5]

    # --------------------------------------------------
    # TAB 1: VISÃO GERAL E KPIs
    # --------------------------------------------------
    with tab_visao_geral:
        if pode_ver_custos:
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
        else:
            total_skus = (
                df_filtrado["Produto"].nunique() if not df_filtrado.empty else 0
            )
            volume_total = (
                df_filtrado["Saldo 1a.U.M."].sum() if not df_filtrado.empty else 0
            )
            lotes_vencidos = (
                len(df_filtrado[df_filtrado["Classificacao_Saude"] == "VENCIDO"])
                if not df_filtrado.empty
                else 0
            )
            qtd_baixo_giro = (
                len(df_filtrado[df_filtrado["Classificacao_Giro"] == "Baixo Giro"])
                if not df_filtrado.empty
                else 0
            )

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(
                    f'<div class="metric-card"><div class="metric-title">SKUs em Estoque</div><div class="metric-value">{total_skus}</div></div>',
                    unsafe_allow_html=True,
                )
            with col2:
                st.markdown(
                    f'<div class="metric-card"><div class="metric-title">Volume Total</div><div class="metric-value">{formatar_valor_comex(volume_total, sufixo_moeda=False)}</div></div>',
                    unsafe_allow_html=True,
                )
            with col3:
                st.markdown(
                    f'<div class="metric-card"><div class="metric-title">Lotes Vencidos</div><div class="metric-value">{lotes_vencidos}</div></div>',
                    unsafe_allow_html=True,
                )
            with col4:
                st.markdown(
                    f'<div class="metric-card"><div class="metric-title">Registros Baixo Giro</div><div class="metric-value">{qtd_baixo_giro}</div></div>',
                    unsafe_allow_html=True,
                )

    # --------------------------------------------------
    # TAB 2: DETALHES DOS LOTES DO PRODUTO
    # --------------------------------------------------
    with tab_detalhes_lote:
        st.markdown("#### Lista Completa de Lotes")
        if not df_filtrado.empty:
            colunas_lotes = [
                "Filial",
                "Descricao",
                "Lote",
                "Lote Fornec.",
                "Dt. Entrada",
                "Data Validad",
                "Saldo 1a.U.M.",
                "Classificacao_Giro",
                "Classificacao_Saude",
            ]

            if pode_ver_custos:
                colunas_lotes.insert(7, "C Unitario")
                colunas_lotes.insert(8, "Custo total estoque")

            df_exibir = df_filtrado[colunas_lotes].copy()

            df_exibir["Dt. Entrada"] = df_exibir["Dt. Entrada"].dt.strftime("%d/%m/%Y")
            df_exibir["Data Validad"] = df_exibir["Data Validad"].dt.strftime(
                "%d/%m/%Y"
            )

            formatos_lotes = {"Saldo 1a.U.M.": "{:,.2f}"}
            if pode_ver_custos:
                formatos_lotes.update(
                    {
                        "C Unitario": "R$ {:,.2f}",
                        "Custo total estoque": "R$ {:,.2f}",
                    }
                )

            st.dataframe(
                df_exibir.style.format(formatos_lotes),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("Nenhum lote para listar.")

    # --------------------------------------------------
    # TAB 3: VISÃO MACRO GRÁFICA + INSPEÇÃO EXECUTIVA (DRILL-DOWN)
    # --------------------------------------------------
    with tab_graficos:
        if not pode_ver_custos:
            st.write("")
        elif not df_macro.empty:
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

    # --------------------------------------------------
    # TAB 5: CONTAINERS / EMBARQUES
    # --------------------------------------------------
    with tab_containers:
        st.markdown("###  Torre de Controle de Containers")
        st.caption(
            "Visão executiva dos containers/embarques vinculados aos pedidos de compra em andamento. "
            "Quando não houver número de container na planilha, o painel consolida uma referência por fornecedor, filial e data prevista."
        )

        if df_transito.empty:
            st.warning(
                "Não foram encontrados dados válidos na aba PED. COMPRA para montar a visão de containers."
            )
        else:
            col_cf1, col_cf2, col_cf3 = st.columns([1, 1, 1])

            opcoes_status_container = ["TODOS"] + sorted(
                df_transito["Status_Entrega"].dropna().astype(str).unique()
            )
            opcoes_fornecedor_container = ["TODOS"] + sorted(
                df_transito["Fornecedor"].dropna().astype(str).unique()
            )
            opcoes_filial_container = ["TODAS"] + sorted(
                df_transito["Filial"].dropna().astype(str).unique()
            )

            with col_cf1:
                status_container = st.multiselect(
                    "Filtrar status:",
                    options=opcoes_status_container,
                    default=["TODOS"],
                    key="status_container",
                )
            with col_cf2:
                fornecedor_container = st.multiselect(
                    "Filtrar fornecedor:",
                    options=opcoes_fornecedor_container,
                    default=["TODOS"],
                    key="fornecedor_container",
                )
            with col_cf3:
                filial_container = st.multiselect(
                    "Filtrar filial/matriz:",
                    options=opcoes_filial_container,
                    default=["TODAS"],
                    key="filial_container",
                )

            df_container_base = df_transito.copy()

            if not ("TODOS" in status_container or not status_container):
                df_container_base = df_container_base[
                    df_container_base["Status_Entrega"]
                    .astype(str)
                    .isin(status_container)
                ]

            if not ("TODOS" in fornecedor_container or not fornecedor_container):
                df_container_base = df_container_base[
                    df_container_base["Fornecedor"]
                    .astype(str)
                    .isin(fornecedor_container)
                ]

            if not ("TODAS" in filial_container or not filial_container):
                df_container_base = df_container_base[
                    df_container_base["Filial"].astype(str).isin(filial_container)
                ]

            if df_container_base.empty:
                st.info("Sem containers/embarques para os filtros selecionados.")
            else:
                df_containers = (
                    df_container_base.groupby("Container_Ref", dropna=False)
                    .agg(
                        Fornecedor=(
                            "Fornecedor",
                            lambda x: ", ".join(sorted(set(x.astype(str)))[:3]),
                        ),
                        Filial=(
                            "Filial",
                            lambda x: ", ".join(sorted(set(x.astype(str)))[:3]),
                        ),
                        Data_Entrega_Armazem=("Data_Entrega_Armazem", "min"),
                        Quantidade_kg=("Quantidade_kg", "sum"),
                        Produtos=("Produto", "nunique"),
                        Status_Entrega=("Status_Entrega", lambda x: x.iloc[0]),
                    )
                    .reset_index()
                )

                df_containers["Dias_Para_Entrega"] = (
                    df_containers["Data_Entrega_Armazem"] - DATA_HOJE
                ).dt.days
                df_containers["Mes_Entrega"] = (
                    df_containers["Data_Entrega_Armazem"]
                    .dt.to_period("M")
                    .dt.to_timestamp()
                )

                total_containers = df_containers["Container_Ref"].nunique()
                volume_total_containers = df_containers["Quantidade_kg"].sum()
                containers_atrasados = df_containers[
                    df_containers["Dias_Para_Entrega"] < 0
                ]["Container_Ref"].nunique()
                proxima_chegada_container = df_containers["Data_Entrega_Armazem"].min()

                cc1, cc2, cc3, cc4 = st.columns(4)
                with cc1:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Containers / Embarques</div><div class="metric-value">{total_containers}</div></div>',
                        unsafe_allow_html=True,
                    )
                with cc2:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Volume Consolidado</div><div class="metric-value">{formatar_kg(volume_total_containers)}</div></div>',
                        unsafe_allow_html=True,
                    )
                with cc3:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Atrasados</div><div class="metric-value">{containers_atrasados}</div></div>',
                        unsafe_allow_html=True,
                    )
                with cc4:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Próxima Chegada</div><div class="metric-value">{proxima_chegada_container.strftime("%d/%m/%Y")}</div></div>',
                        unsafe_allow_html=True,
                    )

                # Insight executivo de containers
                volume_por_mes_container = (
                    df_containers.groupby("Mes_Entrega", as_index=False)[
                        "Quantidade_kg"
                    ]
                    .sum()
                    .sort_values("Mes_Entrega")
                )
                mes_pico_container = volume_por_mes_container.loc[
                    volume_por_mes_container["Quantidade_kg"].idxmax()
                ]

                top_container_fornecedor = (
                    df_containers.groupby("Fornecedor", dropna=False)["Quantidade_kg"]
                    .sum()
                    .sort_values(ascending=False)
                )
                fornecedor_pico_container = str(top_container_fornecedor.index[0])
                perc_fornecedor_pico_container = (
                    top_container_fornecedor.iloc[0] / volume_total_containers * 100
                    if volume_total_containers > 0
                    else 0
                )

                st.markdown(
                    f"""
                    <div class="alert-box" style="border-left-color: #FFB800 !important;">
                        <span class="alert-title"> Resumo Executivo de Containers</span>
                        <span class="alert-text">
                            Existem <strong>{total_containers}</strong> containers/embarques monitorados, totalizando <strong>{formatar_kg(volume_total_containers)}</strong>.<br>
                            O maior volume está concentrado em <strong>{mes_pico_container["Mes_Entrega"].strftime("%m/%Y")}</strong>, com <strong>{formatar_kg(mes_pico_container["Quantidade_kg"])}</strong> previstos.<br>
                            O principal fornecedor no pipeline é <strong>{fornecedor_pico_container}</strong>, representando <strong>{perc_fornecedor_pico_container:.1f}%</strong> do volume consolidado.<br>
                            Há <strong>{containers_atrasados}</strong> containers/embarques com entrega vencida em relação à data de referência.
                        </span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                st.markdown("####  Containers / embarques previstos por mês")

                volume_por_mes_container["Mes_Label"] = volume_por_mes_container[
                    "Mes_Entrega"
                ].dt.strftime("%m/%Y")
                volume_por_mes_container["Containers"] = volume_por_mes_container[
                    "Mes_Entrega"
                ].map(df_containers.groupby("Mes_Entrega")["Container_Ref"].nunique())

                fig_container_mes = px.bar(
                    volume_por_mes_container,
                    x="Mes_Label",
                    y="Quantidade_kg",
                    text="Containers",
                    labels={
                        "Mes_Label": "Mês de entrega",
                        "Quantidade_kg": "Volume previsto (kg)",
                    },
                    title="Volume consolidado por mês de chegada no armazém",
                )
                fig_container_mes.update_traces(
                    marker_color="#4B2A85",
                    texttemplate="%{text} embarques",
                    textposition="outside",
                    hovertemplate="<b>%{x}</b><br>Volume: %{y:,.2f} kg<br>Embarques: %{text}<extra></extra>",
                )
                fig_container_mes.update_layout(
                    height=430,
                    showlegend=False,
                    margin=dict(t=60, b=40, l=40, r=40),
                    yaxis_title="Volume (kg)",
                    xaxis_title="Mês de entrega",
                )
                st.plotly_chart(fig_container_mes, use_container_width=True)

                col_ct1, col_ct2 = st.columns(2)

                with col_ct1:
                    st.markdown("####  Status dos embarques")
                    dados_status = (
                        df_containers.groupby("Status_Entrega", as_index=False)
                        .agg(
                            Containers=("Container_Ref", "nunique"),
                            Quantidade_kg=("Quantidade_kg", "sum"),
                        )
                        .sort_values("Containers", ascending=False)
                    )
                    fig_status = px.pie(
                        dados_status,
                        names="Status_Entrega",
                        values="Containers",
                        hole=0.45,
                        color_discrete_sequence=[
                            "#E91010",
                            "#FFB800",
                            "#07E249",
                            "#00E5FF",
                            "#4B2A85",
                        ],
                    )
                    fig_status.update_traces(
                        textinfo="percent+label",
                        hovertemplate="<b>%{label}</b><br>Containers/embarques: %{value}<extra></extra>",
                    )
                    fig_status.update_layout(
                        height=420, margin=dict(t=20, b=30, l=20, r=20)
                    )
                    st.plotly_chart(fig_status, use_container_width=True)

                with col_ct2:
                    st.markdown("####  Top fornecedores por volume consolidado")
                    dados_fornecedor_container = (
                        df_containers.groupby("Fornecedor", as_index=False)[
                            "Quantidade_kg"
                        ]
                        .sum()
                        .sort_values("Quantidade_kg", ascending=False)
                        .head(10)
                    )
                    fig_fornecedor_container = px.bar(
                        dados_fornecedor_container,
                        x="Quantidade_kg",
                        y="Fornecedor",
                        orientation="h",
                        text="Quantidade_kg",
                        labels={
                            "Quantidade_kg": "Volume (kg)",
                            "Fornecedor": "Fornecedor",
                        },
                    )
                    fig_fornecedor_container.update_traces(
                        marker_color="#5A339B",
                        texttemplate="%{text:,.0f} kg",
                        textposition="outside",
                        hovertemplate="<b>%{y}</b><br>Volume: %{x:,.2f} kg<extra></extra>",
                    )
                    fig_fornecedor_container.update_layout(
                        height=420,
                        yaxis={"categoryorder": "total ascending"},
                        margin=dict(t=20, b=40, l=20, r=40),
                    )
                    st.plotly_chart(fig_fornecedor_container, use_container_width=True)

                st.markdown("####  Próximas chegadas de containers/embarques")

                proximas_chegadas = df_containers.sort_values(
                    "Data_Entrega_Armazem"
                ).head(12)

                tabela_containers = proximas_chegadas[
                    [
                        "Container_Ref",
                        "Fornecedor",
                        "Filial",
                        "Data_Entrega_Armazem",
                        "Quantidade_kg",
                        "Produtos",
                        "Dias_Para_Entrega",
                        "Status_Entrega",
                    ]
                ].copy()

                tabela_containers["Data_Entrega_Armazem"] = tabela_containers[
                    "Data_Entrega_Armazem"
                ].dt.strftime("%d/%m/%Y")
                tabela_containers["Quantidade_kg"] = tabela_containers[
                    "Quantidade_kg"
                ].map(lambda x: f"{x:,.2f} kg")
                tabela_containers["Dias_Para_Entrega"] = tabela_containers[
                    "Dias_Para_Entrega"
                ].map(
                    lambda x: (
                        f"Atrasado há {abs(int(x))} dias" if x < 0 else f"{int(x)} dias"
                    )
                )

                tabela_containers.columns = [
                    "Container / Embarque",
                    "Fornecedor",
                    "Filial/Matriz",
                    "Chegada no Armazém",
                    "Volume",
                    "Produtos",
                    "Prazo",
                    "Status",
                ]

                st.dataframe(
                    tabela_containers, use_container_width=True, hide_index=True
                )

                with st.expander("Ver base completa de containers/embarques"):
                    tabela_completa_container = df_containers.sort_values(
                        "Data_Entrega_Armazem"
                    ).copy()
                    tabela_completa_container["Data_Entrega_Armazem"] = (
                        tabela_completa_container["Data_Entrega_Armazem"].dt.strftime(
                            "%d/%m/%Y"
                        )
                    )
                    tabela_completa_container["Quantidade_kg"] = (
                        tabela_completa_container["Quantidade_kg"].map(
                            lambda x: f"{x:,.2f} kg"
                        )
                    )
                    tabela_completa_container["Dias_Para_Entrega"] = (
                        tabela_completa_container["Dias_Para_Entrega"].map(
                            lambda x: (
                                f"Atrasado há {abs(int(x))} dias"
                                if x < 0
                                else f"{int(x)} dias"
                            )
                        )
                    )
                    tabela_completa_container = tabela_completa_container[
                        [
                            "Container_Ref",
                            "Fornecedor",
                            "Filial",
                            "Data_Entrega_Armazem",
                            "Quantidade_kg",
                            "Produtos",
                            "Dias_Para_Entrega",
                            "Status_Entrega",
                        ]
                    ]
                    tabela_completa_container.columns = [
                        "Container / Embarque",
                        "Fornecedor",
                        "Filial/Matriz",
                        "Chegada no Armazém",
                        "Volume",
                        "Produtos",
                        "Prazo",
                        "Status",
                    ]
                    st.dataframe(
                        tabela_completa_container,
                        use_container_width=True,
                        hide_index=True,
                    )

    # --------------------------------------------------
    # TAB 6: CUSTOS - RESTRITA A PERFIS AUTORIZADOS
    # --------------------------------------------------
    if pode_ver_custos:
        with tab_custos:
            st.markdown("###  Custos")
            st.caption(
                "Visão restrita de custo médio histórico por KG e projeção de custo futuro por produto."
            )

            if df_custo_hist_mensal.empty and df_custo_proj_mensal.empty:
                st.warning(
                    "Não foram encontrados dados válidos nas abas Hist. Venda e Previsão de custo. "
                    "Verifique se as colunas solicitadas estão preenchidas corretamente."
                )
            else:
                produtos_custos = sorted(
                    set(
                        df_custo_hist_mensal.get("Produto", pd.Series(dtype=str))
                        .dropna()
                        .astype(str)
                        .unique()
                    )
                    | set(
                        df_custo_proj_mensal.get("Produto", pd.Series(dtype=str))
                        .dropna()
                        .astype(str)
                        .unique()
                    )
                )
                produto_custo = st.selectbox(
                    "Selecione o produto para análise de custo:",
                    options=["TODOS"] + produtos_custos,
                    index=0,
                    key="produto_custos",
                )

                hist_custo_filtrado = df_custo_hist_mensal.copy()
                proj_custo_filtrado = df_custo_proj_mensal.copy()

                if produto_custo != "TODOS":
                    hist_custo_filtrado = hist_custo_filtrado[
                        hist_custo_filtrado["Produto"].astype(str) == str(produto_custo)
                    ]
                    proj_custo_filtrado = proj_custo_filtrado[
                        proj_custo_filtrado["Produto"].astype(str) == str(produto_custo)
                    ]

                # Quando seleciona TODOS, consolida por mês com média ponderada no histórico e média simples na projeção.
                if produto_custo == "TODOS":
                    if not df_hist_custos.empty:
                        hist_plot = df_hist_custos.groupby("Mes", as_index=False).agg(
                            Quantidade_kg=("Quantidade_kg", "sum"),
                            Custo_Total_Calculado=("Custo_Total_Calculado", "sum"),
                        )
                        hist_plot["Custo_KG"] = (
                            hist_plot["Custo_Total_Calculado"]
                            / hist_plot["Quantidade_kg"]
                        )
                        hist_plot["Tipo"] = "Histórico"
                    else:
                        hist_plot = pd.DataFrame()

                    if not df_previsao_custos.empty:
                        proj_plot = df_previsao_custos.groupby(
                            "Mes", as_index=False
                        ).agg(Custo_KG=("Custo_KG_Projetado", "mean"))
                        proj_plot["Tipo"] = "Projeção"
                    else:
                        proj_plot = pd.DataFrame()
                else:
                    hist_plot = hist_custo_filtrado.copy()
                    proj_plot = proj_custo_filtrado.copy()

                ultimo_custo_hist = (
                    hist_plot.sort_values("Mes")["Custo_KG"].iloc[-1]
                    if not hist_plot.empty
                    else 0
                )
                proximo_custo_proj = (
                    proj_plot.sort_values("Mes")["Custo_KG"].iloc[0]
                    if not proj_plot.empty
                    else 0
                )
                variacao_proj = (
                    ((proximo_custo_proj - ultimo_custo_hist) / ultimo_custo_hist) * 100
                    if ultimo_custo_hist > 0 and proximo_custo_proj > 0
                    else 0
                )
                qtd_hist_total = (
                    hist_plot["Quantidade_kg"].sum()
                    if "Quantidade_kg" in hist_plot.columns and not hist_plot.empty
                    else 0
                )

                cst1, cst2, cst3, cst4 = st.columns(4)
                with cst1:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Último Custo Médio</div><div class="metric-value">R$ {ultimo_custo_hist:,.2f}/kg</div></div>',
                        unsafe_allow_html=True,
                    )
                with cst2:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Próximo Custo Projetado</div><div class="metric-value">R$ {proximo_custo_proj:,.2f}/kg</div></div>',
                        unsafe_allow_html=True,
                    )
                with cst3:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Variação Projetada</div><div class="metric-value">{variacao_proj:.1f}%</div></div>',
                        unsafe_allow_html=True,
                    )
                with cst4:
                    st.markdown(
                        f'<div class="metric-card"><div class="metric-title">Volume Histórico</div><div class="metric-value">{formatar_kg(qtd_hist_total)}</div></div>',
                        unsafe_allow_html=True,
                    )

                st.markdown(
                    f"""
                    <div class="alert-box" style="border-left-color: #00E5FF !important;">
                        <span class="alert-title" style="color: #00E5FF !important;"> Resumo Executivo de Custos</span>
                        <span class="alert-text">
                            O último custo médio histórico apurado é de <strong>R$ {ultimo_custo_hist:,.2f}/kg</strong>.<br>
                            A próxima projeção de custo é de <strong>R$ {proximo_custo_proj:,.2f}/kg</strong>.<br>
                            A variação projetada em relação ao último histórico é de <strong>{variacao_proj:.1f}%</strong>.<br>
                            Esta aba é restrita a perfis autorizados e não aparece para usuários sem permissão de custos.
                        </span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                dados_grafico_custos = []
                if not hist_plot.empty:
                    temp_hist = hist_plot[["Mes", "Custo_KG", "Tipo"]].copy()
                    dados_grafico_custos.append(temp_hist)
                if not proj_plot.empty:
                    temp_proj = proj_plot[["Mes", "Custo_KG", "Tipo"]].copy()
                    dados_grafico_custos.append(temp_proj)

                if dados_grafico_custos:
                    df_grafico_custos = pd.concat(
                        dados_grafico_custos, ignore_index=True
                    )
                    df_grafico_custos = df_grafico_custos.sort_values("Mes")
                    df_grafico_custos["Mes_Label"] = df_grafico_custos[
                        "Mes"
                    ].dt.strftime("%m/%Y")

                    st.markdown("####  Custo médio por KG mês a mês")
                    fig_custos = px.line(
                        df_grafico_custos,
                        x="Mes_Label",
                        y="Custo_KG",
                        color="Tipo",
                        markers=True,
                        text="Custo_KG",
                        labels={
                            "Mes_Label": "Mês",
                            "Custo_KG": "Custo por KG",
                            "Tipo": "Origem",
                        },
                        title="Histórico e projeção de custo por KG",
                    )
                    fig_custos.update_traces(
                        line=dict(width=4),
                        texttemplate="R$ %{text:,.2f}",
                        textposition="top center",
                        hovertemplate="<b>%{x}</b><br>Custo: R$ %{y:,.2f}/kg<extra></extra>",
                    )
                    fig_custos.update_layout(
                        height=430,
                        margin=dict(t=60, b=40, l=40, r=40),
                        yaxis_title="Custo por KG",
                        xaxis_title="Mês",
                    )
                    st.plotly_chart(fig_custos, use_container_width=True)

                col_hist, col_proj = st.columns(2)
                with col_hist:
                    st.markdown("####  Histórico por produto")
                    if hist_custo_filtrado.empty:
                        st.info("Sem histórico de custo para o filtro selecionado.")
                    else:
                        tabela_hist = hist_custo_filtrado.sort_values(
                            "Mes", ascending=False
                        ).copy()
                        tabela_hist["Mes"] = tabela_hist["Mes"].dt.strftime("%m/%Y")
                        tabela_hist["Custo_KG"] = tabela_hist["Custo_KG"].map(
                            lambda x: f"R$ {x:,.2f}/kg"
                        )
                        if "Quantidade_kg" in tabela_hist.columns:
                            tabela_hist["Quantidade_kg"] = tabela_hist[
                                "Quantidade_kg"
                            ].map(lambda x: f"{x:,.2f} kg")
                        colunas_hist = [
                            c
                            for c in ["Produto", "Mes", "Quantidade_kg", "Custo_KG"]
                            if c in tabela_hist.columns
                        ]
                        tabela_hist = tabela_hist[colunas_hist].head(30)
                        tabela_hist.columns = [
                            "Produto",
                            "Mês",
                            "Quantidade",
                            "Custo médio",
                        ][: len(tabela_hist.columns)]
                        st.dataframe(
                            tabela_hist, use_container_width=True, hide_index=True
                        )

                with col_proj:
                    st.markdown("####  Projeção futura")
                    if proj_custo_filtrado.empty:
                        st.info("Sem projeção de custo para o filtro selecionado.")
                    else:
                        tabela_proj = proj_custo_filtrado.sort_values("Mes").copy()
                        tabela_proj["Mes"] = tabela_proj["Mes"].dt.strftime("%m/%Y")
                        tabela_proj["Custo_KG"] = tabela_proj["Custo_KG"].map(
                            lambda x: f"R$ {x:,.2f}/kg"
                        )
                        tabela_proj = tabela_proj[["Produto", "Mes", "Custo_KG"]].head(
                            30
                        )
                        tabela_proj.columns = [
                            "Produto",
                            "Mês previsto",
                            "Custo projetado",
                        ]
                        st.dataframe(
                            tabela_proj, use_container_width=True, hide_index=True
                        )

except Exception as erro:
    st.error(f"Erro ao processar e estruturar o painel de COMEX: {erro}")
