import io
import sqlite3
import pandas as pd
import streamlit as st
import re  

# ==========================================
# TRATAMENTO E PADRONIZAÇÃO DE CLIENTES
# ==========================================
def padronizar_cliente(nome):
    if not isinstance(nome, str):
        return nome
    
    nome = nome.upper().strip()
    nome = re.sub(r'\s+', ' ', nome)
    
    # DSM e SAVINA
    if 'DSM' in nome or 'DS&M' in nome: return 'DSM'
    if 'SAVINA' in nome or 'SANIVA' in nome: return 'SAVINA'

    # CLIENTES DA PLANILHA
    if 'TONACRIL' in nome: return 'TONACRIL'
    if 'WS CARDOSO' in nome or 'W S CARDOSO' in nome: return 'WS CARDOSO'
    if 'ACQUACORES' in nome or 'ACQUA CORES' in nome: return 'ACQUACORES'
    if 'LABORSAN' in nome: return 'LABORSAN'
    if 'BIG MASSA' in nome or 'BIGMASSA' in nome: return 'BIG MASSA'
    if 'FS DE MORAIS' in nome or 'F S DE MORAIS' in nome or 'FS DE MORAES' in nome: return 'FS DE MORAIS'
    if 'WILTON' in nome: return 'WILTON IND'
    if 'WESTROCK' in nome or 'WEST ROCK' in nome: return 'WESTROCK'
    if 'JA LARA' in nome or 'J A LARA' in nome or 'J.A. LARA' in nome: return 'JA LARA'
    if 'PROTELIM' in nome: return 'PROTELIM'
    if 'GENESIS' in nome or 'GÊNESIS' in nome: return 'GENESIS'
    if 'GARIN' in nome: return 'GARIN'
    if 'OUROCOLOR' in nome or 'OURO COLOR' in nome: return 'OUROCOLOR'
    if 'ROYAL MARK' in nome or 'ROYALMARK' in nome: return 'ROYAL MARK'
    if 'SAINT-GOBAIN' in nome or 'SAINT GOBAIN' in nome or 'SAINTGOBAIN' in nome: return 'SAINT-GOBAIN'
    if 'AGRO QUIM' in nome or 'AGROQUIM' in nome or 'AGRO QUÍM' in nome: return 'AGROQUIMICA'
    
    # OUTROS CLIENTES
    if 'ACTEGA' in nome or 'AZTEGA' in nome: return 'ACTEGA'
    if 'SUN CHEMICAL' in nome: return 'SUN CHEMICAL'
    if 'ALLTAK' in nome: return 'ALLTAK'
    if 'COLOR DEX' in nome or 'COLORDEX' in nome: return 'COLORDEX'
    if 'FLEXOINK' in nome: return 'FLEXOINK'
    if 'UIRAPURU' in nome: return 'UIRAPURU'
    if 'ENGEFORTE' in nome: return 'ENGEFORTE'
    if 'VIACOLOR' in nome: return 'VIACOLOR'
    if 'BORDEAUX' in nome: return 'BORDEAUX'
    if 'MARTINS' in nome: return 'MARTINS'

    return nome

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Painel Executivo de Vendas", page_icon="📊", layout="wide"
)

USUARIOS_PERMITIDOS = st.secrets["passwords"]

# Inicialização do estado de autenticação
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False


def tela_login():
    """Renderiza a interface inicial de login."""
    st.markdown("## 🔒 Acesso ao Sistema de Inteligência de Vendas")
    st.markdown("Por favor, insira suas credenciais para acessar o painel.")

    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        with st.form("form_login"):
            usuario_input = st.text_input("Usuário").strip()
            senha_input = st.text_input("Senha", type="password")
            btn_entrar = st.form_submit_button("Entrar no Painel", type="primary")

            if btn_entrar:
                # Validação direta respeitando maiúsculas e minúsculas
                if (
                    usuario_input in USUARIOS_PERMITIDOS
                    and USUARIOS_PERMITIDOS[usuario_input] == senha_input
                ):
                    st.session_state["autenticado"] = True
                    st.session_state["usuario_logado"] = usuario_input
                    st.success("Acesso liberado!")
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos. Verifique maiúsculas e minúsculas.")


# --- CONTROLE DE ACESSO ---
if not st.session_state["autenticado"]:
    tela_login()
    st.stop()

# --- A PARTIR DESTE PONTO O USUÁRIO ESTÁ AUTENTICADO ---

# Barra Lateral: Identificação do Usuário e Logout
st.sidebar.markdown(f"👤 *Usuário Ativo:* {st.session_state['usuario_logado']}")
if st.sidebar.button("🚪 Sair / Logout"):
    st.session_state["autenticado"] = False
    st.session_state.pop("usuario_logado", None)
    st.rerun()

st.title("📊 Painel de Análise de Saídas e Clientes (2020 - 2024)")


# --- FUNÇÃO AUXILIAR PARA EXPORTAR EXCEL ---
def gerar_excel(df, nome_aba="Dados"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nome_aba)
    return output.getvalue()


# --- CONSULTAS COM CACHE DE BANCO ---
@st.cache_data(ttl=3600)
def obter_filtros_iniciais():
    conn = sqlite3.connect("estoque.db")

    anos = pd.read_sql_query(
        "SELECT DISTINCT ANO_ORIGEM FROM movimentacao_vendas WHERE ANO_ORIGEM IS NOT NULL AND TRIM(ANO_ORIGEM) != '' ORDER BY ANO_ORIGEM DESC",
        conn,
    )["ANO_ORIGEM"].tolist()

 # 1. Carrega o DataFrame com os clientes do banco
df_clientes_raw = pd.read_sql_query(
    """
    SELECT DISTINCT NOME_CLIENTE
    FROM movimentacao_vendas
    WHERE NOME_CLIENTE IS NOT NULL
        AND LOWER(TRIM(NOME_CLIENTE)) NOT IN ('não informado', 'nao informado', 'produção')
    """,
    conn
)

# 2. Aplica a padronização para limpar duplicados e erros
df_clientes_raw['NOME_CLIENTE'] = df_clientes_raw['NOME_CLIENTE'].apply(padronizar_cliente)

# 3. Gera a lista final sem duplicidades e ordenada para o dropdown
clientes = sorted(df_clientes_raw['NOME_CLIENTE'].dropna().unique().tolist())

conn.close()
return anos, clientes


try:
    anos_disponiveis, clientes_disponiveis = obter_filtros_iniciais()

    # --- BARRA LATERAL (FILTROS CRUZADOS DINÂMICOS) ---
    st.sidebar.header("🔍 Filtros de Pesquisa")

    ano_selecionado = st.sidebar.multiselect(
        "Selecione o(s) Ano(s)",
        options=anos_disponiveis,
        default=[anos_disponiveis[0]] if anos_disponiveis else [],
    )

    cliente_selecionado = st.sidebar.multiselect(
        "Filtrar por Cliente",
        options=clientes_disponiveis,
        placeholder="Busque o nome do cliente...",
    )

    # Lógica de Filtro Cruzado: Busca apenas produtos comprados pelos clientes selecionados
    conn = sqlite3.connect("estoque.db")

    condicoes_prods_disponiveis = [
        "NOME_DO_PRODUTO IS NOT NULL",
        "LOWER(TRIM(NOME_DO_PRODUTO)) NOT IN ('null', '', 'none')",
    ]
    if cliente_selecionado:
        clis_fmt = "', '".join(
            [c.replace("'", "''") for c in cliente_selecionado]
        )
        condicoes_prods_disponiveis.append(f"NOME_CLIENTE IN ('{clis_fmt}')")
    if ano_selecionado:
        anos_fmt = "', '".join(ano_selecionado)
        condicoes_prods_disponiveis.append(f"ANO_ORIGEM IN ('{anos_fmt}')")

    where_prods = "WHERE " + " AND ".join(condicoes_prods_disponiveis)
    produtos_disponiveis = pd.read_sql_query(
        f"SELECT DISTINCT NOME_DO_PRODUTO FROM movimentacao_vendas {where_prods} ORDER BY NOME_DO_PRODUTO ASC",
        conn,
    )["NOME_DO_PRODUTO"].tolist()

    produto_selecionado = st.sidebar.multiselect(
        "Filtrar por Produto",
        options=produtos_disponiveis,
        placeholder="Busque o produto...",
    )

    # --- CONSTRUÇÃO DA CLÁUSULA WHERE PRINCIPAL ---
    condicoes_base = []

    if ano_selecionado:
        anos_fmt = "', '".join(ano_selecionado)
        condicoes_base.append(f"ANO_ORIGEM IN ('{anos_fmt}')")

    if cliente_selecionado:
        clis_fmt = "', '".join(
            [c.replace("'", "''") for c in cliente_selecionado]
        )
        condicoes_base.append(f"NOME_CLIENTE IN ('{clis_fmt}')")

    if produto_selecionado:
        prods_fmt = "', '".join(
            [p.replace("'", "''") for p in produto_selecionado]
        )
        condicoes_base.append(f"NOME_DO_PRODUTO IN ('{prods_fmt}')")

    # Exclusão global de registros inválidos
    condicoes_base.append("NOME_CLIENTE IS NOT NULL")
    condicoes_base.append(
        "LOWER(TRIM(NOME_CLIENTE)) NOT IN ('não informado', 'nao informado', 'produção', 'producao', 'null', '', 'none')"
    )
    condicoes_base.append("NOME_DO_PRODUTO IS NOT NULL")
    condicoes_base.append(
        "LOWER(TRIM(NOME_DO_PRODUTO)) NOT IN ('null', '', 'none')"
    )

    where_sql = "WHERE " + " AND ".join(condicoes_base)

    # --- CONSULTAS DE DADOS ---

    # 1. KPIs Gerais
    query_totais = f"SELECT SUM(QUANTIDADE_KG) as TOTAL_KG, COUNT(*) as TOTAL_REGISTROS, COUNT(DISTINCT NOME_CLIENTE) as TOTAL_CLIENTES FROM movimentacao_vendas {where_sql}"
    df_totais = pd.read_sql_query(query_totais, conn)

 # 2. Ranking de Todos os Clientes
query_clientes = f"""
    SELECT 
        NOME_CLIENTE AS "Cliente",
        SUM(QUANTIDADE_KG) AS "Volume_KG"
    FROM movimentacao_vendas
    {where_sql} AND UPPER(NOME_CLIENTE) NOT LIKE '%PRODUÇÃO%'
    GROUP BY NOME_CLIENTE
"""

df_todos_clientes = pd.read_sql_query(query_clientes, conn)

# --- APLICA A NOSSA PADRONIZAÇÃO E SOMA OS VOLUMES ---
df_todos_clientes['Cliente'] = df_todos_clientes['Cliente'].apply(padronizar_cliente)
df_todos_clientes = (
    df_todos_clientes.groupby('Cliente', as_index=False)['Volume_KG']
    .sum()
    .sort_values(by='Volume_KG', ascending=False)
)
    # 3. Ranking de Todos os Produtos do Filtro
    query_produtos = f"""
        SELECT NOME_DO_PRODUTO AS "Produto", SUM(QUANTIDADE_KG) AS "Volume_KG"
        FROM movimentacao_vendas
        {where_sql}
        GROUP BY NOME_DO_PRODUTO
        ORDER BY "Volume_KG" DESC
    """
    df_todos_produtos = pd.read_sql_query(query_produtos, conn)

    # 4. Dados para Gráfico Temporal Interativo
    query_grafico = f"""
        SELECT ANO_ORIGEM, MES_ORIGEM, SUM(QUANTIDADE_KG) AS TOTAL_KG
        FROM movimentacao_vendas
        {where_sql}
        GROUP BY ANO_ORIGEM, MES_ORIGEM
    """
    df_grafico = pd.read_sql_query(query_grafico, conn)

    conn.close()

    # --- RENDERIZAÇÃO DA PÁGINA ---

    # KPI HEADERS
    col1, col2, col3 = st.columns(3)
    total_kg = df_totais["TOTAL_KG"].iloc[0] or 0
    total_pedidos = df_totais["TOTAL_REGISTROS"].iloc[0] or 0
    total_clientes = df_totais["TOTAL_CLIENTES"].iloc[0] or 0

    col1.metric(
        "Volume Total (KG)",
        f"{total_kg:,.2f} kg"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", "."),
    )
    col2.metric(
        "Total de Saídas/Registros", f"{total_pedidos:,}".replace(",", ".")
    )
    col3.metric(
        "Clientes Reais Atendidos", f"{total_clientes:,}".replace(",", ".")
    )

    st.markdown("---")

    # CARTOES DE DESTAQUE (MAIS COMPRADO E MENOS COMPRADO)
    if not df_todos_produtos.empty:
        prod_mais_vendido = df_todos_produtos.iloc[0]
        prod_menos_vendido = df_todos_produtos.iloc[-1]

        vol_mais_fmt = (
            f"{prod_mais_vendido['Volume_KG']:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
        vol_menos_fmt = (
            f"{prod_menos_vendido['Volume_KG']:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

        m1, m2 = st.columns(2)
        m1.info(
            f"🏆 *Mais Comprado:* {prod_mais_vendido['Produto']} — *{vol_mais_fmt} kg*"
        )
        m2.warning(
            f"🔻 *Menos Comprado:* {prod_menos_vendido['Produto']} — *{vol_menos_fmt} kg*"
        )

    st.markdown("---")

    # --- SEÇÃO 1: RANKINGS E EXPORTAÇÕES ---
    c1, c2 = st.columns(2)

    # TABELA DE CLIENTES
    with c1:
        st.subheader("🏆 Ranking de Clientes")
        if not df_todos_clientes.empty:
            df_cli_exibir = df_todos_clientes.copy()
            df_cli_exibir.insert(
                0, "Posição", [f"{i+1}º" for i in range(len(df_cli_exibir))]
            )

            excel_cli = gerar_excel(df_cli_exibir, "Clientes")

            df_cli_exibir["Volume (KG)"] = df_cli_exibir["Volume_KG"].apply(
                lambda x: f"{x:,.2f} kg"
                .replace(",", "X")
                .replace(".", ",")
                .replace("X", ".")
            )
            df_cli_exibir.drop(columns=["Volume_KG"], inplace=True)

            st.dataframe(
                df_cli_exibir.head(10),
                use_container_width=True,
                hide_index=True,
            )

            col_exp1, col_dl1 = st.columns([2, 1])
            with col_dl1:
                st.download_button(
                    label="📥 Exportar Clientes (Excel)",
                    data=excel_cli,
                    file_name="relatorio_clientes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            with st.expander(
                f"Ver lista completa ({len(df_cli_exibir)} clientes)"
            ):
                st.dataframe(
                    df_cli_exibir, use_container_width=True, hide_index=True
                )
        else:
            st.info("Nenhum cliente encontrado.")

    # TABELA DE PRODUTOS
    with c2:
        st.subheader("📦 Ranking de Produtos")
        if not df_todos_produtos.empty:
            df_prod_exibir = df_todos_produtos.copy()
            df_prod_exibir.insert(
                0, "Posição", [f"{i+1}º" for i in range(len(df_prod_exibir))]
            )

            excel_prod = gerar_excel(df_prod_exibir, "Produtos")

            df_prod_exibir["Volume (KG)"] = df_prod_exibir["Volume_KG"].apply(
                lambda x: f"{x:,.2f} kg"
                .replace(",", "X")
                .replace(".", ",")
                .replace("X", ".")
            )
            df_prod_exibir.drop(columns=["Volume_KG"], inplace=True)

            st.dataframe(
                df_prod_exibir.head(10),
                use_container_width=True,
                hide_index=True,
            )

            col_exp2, col_dl2 = st.columns([2, 1])
            with col_dl2:
                st.download_button(
                    label="📥 Exportar Produtos (Excel)",
                    data=excel_prod,
                    file_name="relatorio_produtos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            with st.expander(
                f"Ver lista completa ({len(df_prod_exibir)} produtos)"
            ):
                st.dataframe(
                    df_prod_exibir, use_container_width=True, hide_index=True
                )
        else:
            st.info("Nenhum produto encontrado.")

    st.markdown("---")

    # --- SEÇÃO 2: GRÁFICO INTERATIVO DE INSIGHTS ---
    st.subheader("📈 Evolução Temporal e Insights de Volume (KG)")

    if not df_grafico.empty:
        ordem_meses = [
            "JANEIRO",
            "FEVEREIRO",
            "MARÇO",
            "ABRIL",
            "MAIO",
            "JUNHO",
            "JULHO",
            "AGOSTO",
            "SETEMBRO",
            "OUTUBRO",
            "NOVEMBRO",
            "DEZEMBRO",
        ]
        df_grafico["MES_ORIGEM"] = pd.Categorical(
            df_grafico["MES_ORIGEM"], categories=ordem_meses, ordered=True
        )
        df_grafico.sort_values(
            ["ANO_ORIGEM", "MES_ORIGEM"], inplace=True
        )

        pivot_grafico = df_grafico.pivot(
            index="MES_ORIGEM", columns="ANO_ORIGEM", values="TOTAL_KG"
        ).fillna(0)

        st.line_chart(pivot_grafico, height=380, use_container_width=True)
    else:
        st.info("Dados insuficientes para gerar a evolução gráfica.")

except Exception as e:
    st.error(f"Erro ao carregar o painel: {e}")
