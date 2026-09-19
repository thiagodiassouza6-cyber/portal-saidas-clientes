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
    
    # Trava 1: Elimina PRODUCAO / PRODUÇÃO descartando o registro (retorna None)
    if 'PRODUC' in nome:
        return None
    
    # --- REGRA DE OURO: BUSCA PELA PALAVRA-CHAVE PRINCIPAL ---
    
    # ACN (Captura ACN QUIMICA, ACN REPRESENTAÇ, ACN IND, etc.)
    if 'ACN' in nome: return 'ACN QUIMICA'

    # DSM e SAVINA
    if 'DSM' in nome or 'DS&M' in nome: return 'DSM'
    if 'SAVINA' in nome or 'SANIVA' in nome: return 'SAVINA'

    # CLIENTES UNIFICADOS
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
    
    # NOVOS CLIENTES ADICIONADOS
    if 'ACQUAPLUF' in nome: return 'ACQUAPLUF'
    if 'ALTHA COR' in nome or 'ALTHACOR' in nome: return 'ALTHA COR'
    if 'ARTCRIL' in nome: return 'ARTCRIL'
    if 'ATA ASSESSORIA' in nome: return 'ATA ASSESSORIA'
    if 'ATLAS COPCO' in nome: return 'ATLAS COPCO'
    if 'BASF' in nome: return 'BASF'
    if 'BELAFIX' in nome or 'BELLAFIX' in nome or 'BELLA FIX' in nome: return 'BELAFIX'
    if 'COOPERUNI' in nome: return 'COOPERUNI'
    if 'DECORPOL' in nome: return 'DECORPOL'
    if 'E A DA SILVA' in nome: return 'E A DA SILVA'
    if 'ECOPACK' in nome: return 'ECOPACK'
    if 'FERSAL' in nome: return 'FERSAL'

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

    df_clientes_raw = pd.read_sql_query(
        """
        SELECT DISTINCT NOME_CLIENTE
        FROM movimentacao_vendas
        WHERE NOME_CLIENTE IS NOT NULL
            AND LOWER(TRIM(NOME_CLIENTE)) NOT IN ('não informado', 'nao informado', 'produção', 'producao')
            AND UPPER(NOME_CLIENTE) NOT LIKE '%PRODUC%'
        """,
        conn,
    )

    df_clientes_raw['NOME_CLIENTE'] = df_clientes_raw['NOME_CLIENTE'].apply(padronizar_cliente)
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

    # --- CARREGAMENTO E FILTRAGEM VIA PANDAS ---
    conn = sqlite3.connect("estoque.db")

    # Busca a base de vendas respeitando filtro de ano se houver
    condicoes_sql = ["NOME_CLIENTE IS NOT NULL", "NOME_DO_PRODUTO IS NOT NULL"]
    condicoes_sql.append("LOWER(TRIM(NOME_CLIENTE)) NOT IN ('não informado', 'nao informado', 'produção', 'producao', 'null', '', 'none')")
    condicoes_sql.append("UPPER(NOME_CLIENTE) NOT LIKE '%PRODUC%'")
    condicoes_sql.append("LOWER(TRIM(NOME_DO_PRODUTO)) NOT IN ('null', '', 'none')")

    if ano_selecionado:
        anos_fmt = "', '".join(ano_selecionado)
        condicoes_sql.append(f"ANO_ORIGEM IN ('{anos_fmt}')")

    where_clause = "WHERE " + " AND ".join(condicoes_sql)

    query_base = f"""
        SELECT ANO_ORIGEM, MES_ORIGEM, NOME_CLIENTE, NOME_DO_PRODUTO, QUANTIDADE_KG
        FROM movimentacao_vendas
        {where_clause}
    """

    df_vendas = pd.read_sql_query(query_base, conn)
    conn.close()

    # Aplica a padronização no DataFrame completo
    df_vendas['NOME_CLIENTE'] = df_vendas['NOME_CLIENTE'].apply(padronizar_cliente)

    # Trava 2: Remove registros descartados (PRODUÇÃO / None) do DataFrame Pandas
    df_vendas = df_vendas[df_vendas['NOME_CLIENTE'].notnull()]

    # Atualiza lista de produtos dinamicamente baseado nos clientes selecionados
    if cliente_selecionado:
        produtos_disponiveis = sorted(df_vendas[df_vendas['NOME_CLIENTE'].isin(cliente_selecionado)]['NOME_DO_PRODUTO'].dropna().unique().tolist())
    else:
        produtos_disponiveis = sorted(df_vendas['NOME_DO_PRODUTO'].dropna().unique().tolist())

    produto_selecionado = st.sidebar.multiselect(
        "Filtrar por Produto",
        options=produtos_disponiveis,
        placeholder="Busque o produto...",
    )

    # Aplica os filtros de Cliente e Produto selecionados
    if cliente_selecionado:
        df_vendas = df_vendas[df_vendas['NOME_CLIENTE'].isin(cliente_selecionado)]

    if produto_selecionado:
        df_vendas = df_vendas[df_vendas['NOME_DO_PRODUTO'].isin(produto_selecionado)]

    # --- PROCESSAMENTO DOS KPIs E TABELAS ---

    # 1. KPIs Gerais
    total_kg = df_vendas["QUANTIDADE_KG"].sum() if not df_vendas.empty else 0
    total_pedidos = len(df_vendas)
    total_clientes = df_vendas["NOME_CLIENTE"].nunique() if not df_vendas.empty else 0

    # 2. Ranking de Clientes
    if not df_vendas.empty:
        # Trava direta na tabela: remove qualquer variação de PRODUCAO/PRODUÇÃO
        df_vendas_limpo = df_vendas[~df_vendas['NOME_CLIENTE'].astype(str).str.upper().str.contains('PRODUC', na=False)]
        
        df_todos_clientes = (
            df_vendas_limpo.groupby("NOME_CLIENTE", as_index=False)["QUANTIDADE_KG"]
            .sum()
            .rename(columns={"NOME_CLIENTE": "Cliente", "QUANTIDADE_KG": "Volume_KG"})
            .sort_values(by="Volume_KG", ascending=False)
        )
    else:
        df_todos_clientes = pd.DataFrame(columns=["Cliente", "Volume_KG"])

    # 3. Ranking de Produtos
    if not df_vendas.empty:
        df_todos_produtos = (
            df_vendas.groupby("NOME_DO_PRODUTO", as_index=False)["QUANTIDADE_KG"]
            .sum()
            .rename(columns={"NOME_DO_PRODUTO": "Produto", "QUANTIDADE_KG": "Volume_KG"})
            .sort_values(by="Volume_KG", ascending=False)
        )
    else:
        df_todos_produtos = pd.DataFrame(columns=["Produto", "Volume_KG"])

    # 4. Dados para Gráfico Temporal Interativo
    if not df_vendas.empty:
        df_grafico = (
            df_vendas.groupby(["ANO_ORIGEM", "MES_ORIGEM"], as_index=False)["QUANTIDADE_KG"]
            .sum()
            .rename(columns={"QUANTIDADE_KG": "TOTAL_KG"})
        )
    else:
        df_grafico = pd.DataFrame(columns=["ANO_ORIGEM", "MES_ORIGEM", "TOTAL_KG"])

    # --- RENDERIZAÇÃO DA PÁGINA ---

    # KPI HEADERS
    col1, col2, col3 = st.columns(3)

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