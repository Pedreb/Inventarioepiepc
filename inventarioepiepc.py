import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime
import io
from typing import List, Dict, Optional

# Configuração da página
st.set_page_config(
    page_title="Sistema de Inventário",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado para responsividade e design avançado
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    /* Reset e configurações gerais */
    .main {
        font-family: 'Inter', sans-serif;
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
    }

    /* Header com gradiente */
    .header-gradient {
        background: linear-gradient(135deg, #F7931E 0%, #000000 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.5rem;
        font-weight: 700;
        text-align: center;
        margin-bottom: 2rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }

    /* Cards responsivos */
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 15px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.1);
        border: 1px solid rgba(247, 147, 30, 0.1);
        margin-bottom: 1rem;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }

    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 12px 40px rgba(247, 147, 30, 0.2);
    }

    /* Botões personalizados */
    .stButton > button {
        background: linear-gradient(135deg, #F7931E 0%, #ff7b00 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 25px;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(247, 147, 30, 0.3);
    }

    .stButton > button:hover {
        background: linear-gradient(135deg, #e8840d 0%, #f7931e 100%);
        box-shadow: 0 6px 20px rgba(247, 147, 30, 0.4);
        transform: translateY(-2px);
    }

    /* Selectbox personalizado */
    .stSelectbox > div > div {
        background: white;
        border-radius: 10px;
        border: 2px solid rgba(247, 147, 30, 0.2);
        transition: border-color 0.3s ease;
    }

    .stSelectbox > div > div:focus-within {
        border-color: #F7931E;
        box-shadow: 0 0 0 3px rgba(247, 147, 30, 0.1);
    }

    /* Input fields */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input {
        background: white;
        border-radius: 10px;
        border: 2px solid rgba(247, 147, 30, 0.2);
        transition: all 0.3s ease;
        font-size: 1rem;
        padding: 0.75rem;
    }

    .stTextInput > div > div > input:focus,
    .stNumberInput > div > div > input:focus {
        border-color: #F7931E;
        box-shadow: 0 0 0 3px rgba(247, 147, 30, 0.1);
    }

    /* File uploader */
    .stFileUploader > div {
        background: white;
        border-radius: 15px;
        border: 2px dashed rgba(247, 147, 30, 0.3);
        padding: 2rem;
        text-align: center;
        transition: all 0.3s ease;
    }

    .stFileUploader > div:hover {
        border-color: #F7931E;
        background: rgba(247, 147, 30, 0.05);
    }

    /* Tabelas */
    .stDataFrame {
        background: white;
        border-radius: 15px;
        overflow: hidden;
        box-shadow: 0 8px 32px rgba(0,0,0,0.1);
    }

    /* Success/Error messages */
    .stSuccess {
        background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        border-radius: 10px;
        color: white;
        padding: 1rem;
    }

    .stError {
        background: linear-gradient(135deg, #dc3545 0%, #fd7e14 100%);
        border-radius: 10px;
        color: white;
        padding: 1rem;
    }

    /* Sidebar personalizada */
    .css-1d391kg {
        background: linear-gradient(180deg, #F7931E 0%, #000000 100%);
    }

    .css-1d391kg .css-17lntkn {
        color: white;
        font-weight: 600;
    }

    /* Responsividade */
    @media (max-width: 768px) {
        .header-gradient {
            font-size: 2rem;
        }

        .metric-card {
            margin-bottom: 0.5rem;
            padding: 1rem;
        }

        .stButton > button {
            width: 100%;
            margin-bottom: 1rem;
        }
    }

    /* Animações */
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(30px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }

    .fade-in-up {
        animation: fadeInUp 0.6s ease-out;
    }
</style>
""", unsafe_allow_html=True)


class DatabaseManager:
    """Gerenciador do banco de dados SQLite"""

    def __init__(self, db_path: str = "inventario.db"):
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        """Inicializa as tabelas do banco de dados"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Tabela de materiais
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS materiais (
                codigo TEXT PRIMARY KEY,
                descricao TEXT NOT NULL,
                data_cadastro DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Tabela de inventários
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS inventarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                responsavel TEXT NOT NULL,
                data_inventario DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Tabela de itens do inventário
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS inventario_itens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                inventario_id INTEGER,
                codigo_material TEXT,
                quantidade INTEGER NOT NULL,
                FOREIGN KEY (inventario_id) REFERENCES inventarios (id),
                FOREIGN KEY (codigo_material) REFERENCES materiais (codigo)
            )
        """)

        conn.commit()
        conn.close()

    def inserir_materiais(self, materiais: List[Dict[str, str]]) -> bool:
        """Insere materiais no banco de dados"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            for material in materiais:
                cursor.execute("""
                    INSERT OR REPLACE INTO materiais (codigo, descricao)
                    VALUES (?, ?)
                """, (material['codigo'], material['descricao']))

            conn.commit()
            conn.close()
            return True
        except Exception as e:
            st.error(f"Erro ao inserir materiais: {str(e)}")
            return False

    def obter_materiais(self) -> pd.DataFrame:
        """Obtém todos os materiais cadastrados"""
        conn = sqlite3.connect(self.db_path)
        df = pd.read_sql_query("SELECT * FROM materiais ORDER BY codigo", conn)
        conn.close()
        return df

    def criar_inventario(self, responsavel: str) -> int:
        """Cria um novo inventário e retorna o ID"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("""
            INSERT INTO inventarios (responsavel, data_inventario)
            VALUES (?, ?)
        """, (responsavel, datetime.now()))

        inventario_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return inventario_id

    def adicionar_item_inventario(self, inventario_id: int, codigo_material: str, quantidade: int) -> bool:
        """Adiciona um item ao inventário"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            cursor.execute("""
                INSERT INTO inventario_itens (inventario_id, codigo_material, quantidade)
                VALUES (?, ?, ?)
            """, (inventario_id, codigo_material, quantidade))

            conn.commit()
            conn.close()
            return True
        except Exception as e:
            st.error(f"Erro ao adicionar item: {str(e)}")
            return False

    def obter_itens_inventario(self, inventario_id: int) -> pd.DataFrame:
        """Obtém os itens de um inventário"""
        conn = sqlite3.connect(self.db_path)
        query = """
            SELECT 
                ii.codigo_material,
                m.descricao,
                ii.quantidade
            FROM inventario_itens ii
            JOIN materiais m ON ii.codigo_material = m.codigo
            WHERE ii.inventario_id = ?
            ORDER BY ii.codigo_material
        """
        df = pd.read_sql_query(query, conn, params=(inventario_id,))
        conn.close()
        return df

    def obter_inventario_info(self, inventario_id: int) -> Optional[Dict]:
        """Obtém informações do inventário"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("""
            SELECT responsavel, data_inventario
            FROM inventarios
            WHERE id = ?
        """, (inventario_id,))

        result = cursor.fetchone()
        conn.close()

        if result:
            return {
                'responsavel': result[0],
                'data_inventario': result[1]
            }
        return None

    def obter_todos_inventarios(self) -> pd.DataFrame:
        """Obtém todos os inventários realizados"""
        conn = sqlite3.connect(self.db_path)
        query = """
            SELECT 
                i.id,
                i.responsavel,
                i.data_inventario,
                COUNT(ii.id) as total_itens,
                SUM(ii.quantidade) as quantidade_total
            FROM inventarios i
            LEFT JOIN inventario_itens ii ON i.id = ii.inventario_id
            GROUP BY i.id, i.responsavel, i.data_inventario
            ORDER BY i.data_inventario DESC
        """
        df = pd.read_sql_query(query, conn)
        conn.close()

        if not df.empty:
            # Formatar data para padrão brasileiro
            df['data_formatada'] = pd.to_datetime(df['data_inventario']).dt.strftime('%d/%m/%Y')
            # Criar nome do inventário no formato solicitado
            df['nome_inventario'] = df.apply(
                lambda row: f"Inventário Rezende Energia - {row['responsavel']} - {row['data_formatada']}",
                axis=1
            )

        return df


def processar_excel_materiais(arquivo_excel) -> Optional[List[Dict[str, str]]]:
    """Processa o arquivo Excel de materiais"""
    try:
        df = pd.read_excel(arquivo_excel, header=0)

        if df.shape[1] < 2:
            st.error("O arquivo deve ter pelo menos 2 colunas: código e descrição")
            return None

        # Pega as duas primeiras colunas
        df = df.iloc[:, :2]
        df.columns = ['codigo', 'descricao']

        # Remove linhas vazias
        df = df.dropna()

        # Converte para string e remove espaços
        df['codigo'] = df['codigo'].astype(str).str.strip()
        df['descricao'] = df['descricao'].astype(str).str.strip()

        return df.to_dict('records')

    except Exception as e:
        st.error(f"Erro ao processar arquivo Excel: {str(e)}")
        return None


def gerar_excel_inventario(inventario_id: int, db_manager: DatabaseManager) -> io.BytesIO:
    """Gera arquivo Excel do inventário"""
    # Obter informações do inventário
    info = db_manager.obter_inventario_info(inventario_id)
    itens = db_manager.obter_itens_inventario(inventario_id)

    # Formatar data para padrão brasileiro
    data_formatada = datetime.fromisoformat(info['data_inventario']).strftime('%d/%m/%Y %H:%M')

    # Criar arquivo Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Aba com os itens
        if not itens.empty:
            # Renomear colunas para português
            itens_export = itens.copy()
            itens_export.columns = ['Código do Material', 'Descrição', 'Quantidade']
            itens_export.to_excel(writer, sheet_name='Itens do Inventário', index=False)
        else:
            # Criar aba vazia se não houver itens
            pd.DataFrame(columns=['Código do Material', 'Descrição', 'Quantidade']).to_excel(
                writer, sheet_name='Itens do Inventário', index=False
            )

        # Aba com informações gerais
        info_df = pd.DataFrame({
            'Empresa': ['Rezende Energia'],
            'Responsável': [info['responsavel']],
            'Data do Inventário': [data_formatada],
            'Total de Itens Diferentes': [len(itens)],
            'Quantidade Total': [itens['quantidade'].sum() if not itens.empty else 0]
        })
        info_df.to_excel(writer, sheet_name='Informações Gerais', index=False)

        # Aba com resumo por código (se houver itens)
        if not itens.empty:
            resumo_df = itens.groupby('codigo_material').agg({
                'descricao': 'first',
                'quantidade': 'sum'
            }).reset_index()
            resumo_df.columns = ['Código do Material', 'Descrição', 'Quantidade Total']
            resumo_df.to_excel(writer, sheet_name='Resumo por Código', index=False)

    buffer.seek(0)
    return buffer


# Inicializar o gerenciador de banco de dados
if 'db_manager' not in st.session_state:
    st.session_state.db_manager = DatabaseManager()

# Inicializar estados da sessão
if 'inventario_ativo' not in st.session_state:
    st.session_state.inventario_ativo = None
if 'itens_adicionados' not in st.session_state:
    st.session_state.itens_adicionados = []


def main():
    # Título principal
    st.markdown('<h1 class="header-gradient fade-in-up">📋 Sistema de Inventário</h1>',
                unsafe_allow_html=True)

    # Sidebar para navegação
    st.sidebar.title("🔧 Menu Principal")
    opcao = st.sidebar.selectbox(
        "Selecione uma opção:",
        ["📦 Cadastro de Materiais", "📋 Rotina de Inventário", "📊 Relatórios"]
    )

    if opcao == "📦 Cadastro de Materiais":
        tela_cadastro_materiais()
    elif opcao == "📋 Rotina de Inventário":
        tela_rotina_inventario()
    elif opcao == "📊 Relatórios":
        tela_relatorios()


def tela_cadastro_materiais():
    st.markdown('<div class="fade-in-up">', unsafe_allow_html=True)

    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📤 Upload de Planilha de Materiais")
        st.info("Faça upload de um arquivo Excel (.xlsx) com duas colunas: código do item e descrição")

        arquivo = st.file_uploader(
            "Escolha o arquivo Excel",
            type=['xlsx'],
            help="Primeira coluna: Código do Item | Segunda coluna: Descrição"
        )

        if arquivo is not None:
            with st.spinner("Processando arquivo..."):
                materiais = processar_excel_materiais(arquivo)

            if materiais:
                st.success(f"✅ {len(materiais)} materiais encontrados no arquivo!")

                # Preview dos dados
                df_preview = pd.DataFrame(materiais).head(10)
                st.subheader("📋 Preview dos Dados")
                st.dataframe(df_preview, use_container_width=True)

                if len(materiais) > 10:
                    st.info(f"Mostrando apenas os primeiros 10 itens. Total: {len(materiais)}")

                col_btn1, col_btn2 = st.columns(2)

                with col_btn1:
                    if st.button("💾 Importar Materiais", key="importar"):
                        with st.spinner("Importando materiais..."):
                            sucesso = st.session_state.db_manager.inserir_materiais(materiais)

                        if sucesso:
                            st.success("✅ Materiais importados com sucesso!")
                            st.rerun()

                with col_btn2:
                    if st.button("❌ Cancelar", key="cancelar"):
                        st.rerun()

    with col2:
        # Estatísticas
        materiais_df = st.session_state.db_manager.obter_materiais()

        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric(
            "📦 Total de Materiais",
            len(materiais_df),
            delta=None
        )
        st.markdown('</div>', unsafe_allow_html=True)

    # Lista de materiais cadastrados
    if not materiais_df.empty:
        st.subheader("📋 Materiais Cadastrados")
        st.dataframe(materiais_df, use_container_width=True)
    else:
        st.info("Nenhum material cadastrado ainda. Faça o upload de uma planilha para começar.")

    st.markdown('</div>', unsafe_allow_html=True)


def tela_rotina_inventario():
    st.markdown('<div class="fade-in-up">', unsafe_allow_html=True)

    materiais_df = st.session_state.db_manager.obter_materiais()

    if materiais_df.empty:
        st.warning("⚠️ Não há materiais cadastrados. Cadastre materiais primeiro na aba 'Cadastro de Materiais'.")
        return

    # Verificar se há inventário ativo
    if st.session_state.inventario_ativo is None:
        st.subheader("🆕 Iniciar Nova Rotina de Inventário")

        col1, col2 = st.columns([2, 1])

        with col1:
            responsavel = st.text_input(
                "👤 Nome do Responsável",
                placeholder="Digite o nome do responsável pelo inventário"
            )

        with col2:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.write("📅 **Data do Inventário**")
            st.write(datetime.now().strftime("%d/%m/%Y %H:%M"))
            st.markdown('</div>', unsafe_allow_html=True)

        if st.button("🚀 Iniciar Inventário", disabled=not responsavel.strip()):
            inventario_id = st.session_state.db_manager.criar_inventario(responsavel)
            st.session_state.inventario_ativo = inventario_id
            st.session_state.itens_adicionados = []
            st.success(f"✅ Inventário iniciado! ID: {inventario_id}")
            st.rerun()

    else:
        # Inventário ativo
        inventario_id = st.session_state.inventario_ativo
        info = st.session_state.db_manager.obter_inventario_info(inventario_id)

        st.subheader(f"📋 Inventário em Andamento - ID: {inventario_id}")

        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.metric("👤 Responsável", info['responsavel'])
            st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.metric("📅 Data", datetime.fromisoformat(info['data_inventario']).strftime("%d/%m/%Y %H:%M"))
            st.markdown('</div>', unsafe_allow_html=True)

        with col3:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.metric("📦 Itens Adicionados", len(st.session_state.itens_adicionados))
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("---")

        # Adicionar itens ao inventário
        st.subheader("➕ Adicionar Item ao Inventário")

        # Filtrar materiais não adicionados
        materiais_disponiveis = materiais_df[
            ~materiais_df['codigo'].isin(st.session_state.itens_adicionados)
        ]

        if materiais_disponiveis.empty:
            st.info("✅ Todos os materiais já foram adicionados ao inventário!")
        else:
            col_select, col_qty = st.columns([3, 1])

            with col_select:
                opcoes_material = materiais_disponiveis.apply(
                    lambda x: f"{x['codigo']} - {x['descricao']}", axis=1
                ).tolist()

                material_selecionado = st.selectbox(
                    "🔍 Selecione o Material",
                    options=opcoes_material,
                    key="select_material"
                )

            with col_qty:
                quantidade = st.number_input(
                    "📊 Quantidade",
                    min_value=0,
                    value=1,
                    step=1,
                    key="input_quantidade"
                )

            if st.button("➕ Adicionar ao Inventário"):
                if material_selecionado and quantidade >= 0:
                    codigo_material = material_selecionado.split(" - ")[0]

                    sucesso = st.session_state.db_manager.adicionar_item_inventario(
                        inventario_id, codigo_material, quantidade
                    )

                    if sucesso:
                        st.session_state.itens_adicionados.append(codigo_material)
                        st.success(f"✅ Item {codigo_material} adicionado com quantidade {quantidade}!")
                        st.rerun()

        # Mostrar itens já adicionados
        itens_inventario = st.session_state.db_manager.obter_itens_inventario(inventario_id)

        if not itens_inventario.empty:
            st.subheader("📋 Itens no Inventário Atual")
            st.dataframe(itens_inventario, use_container_width=True)

        # Botões de ação
        st.markdown("---")
        col_finalizar, col_excel, col_cancelar = st.columns(3)

        with col_finalizar:
            if st.button("✅ Finalizar Inventário"):
                st.session_state.inventario_ativo = None
                st.session_state.itens_adicionados = []
                st.success("✅ Inventário finalizado com sucesso!")
                st.rerun()

        with col_excel:
            if st.button("📊 Gerar Excel") and not itens_inventario.empty:
                excel_buffer = gerar_excel_inventario(inventario_id, st.session_state.db_manager)

                st.download_button(
                    label="⬇️ Download Excel",
                    data=excel_buffer.getvalue(),
                    file_name=f"inventario_{inventario_id}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        with col_cancelar:
            if st.button("❌ Cancelar Inventário"):
                st.session_state.inventario_ativo = None
                st.session_state.itens_adicionados = []
                st.warning("⚠️ Inventário cancelado!")
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)


def tela_relatorios():
    st.markdown('<div class="fade-in-up">', unsafe_allow_html=True)
    st.subheader("📊 Relatórios e Estatísticas")

    materiais_df = st.session_state.db_manager.obter_materiais()
    inventarios_df = st.session_state.db_manager.obter_todos_inventarios()

    # Métricas gerais
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric(
            "📦 Total de Materiais",
            len(materiais_df)
        )
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric(
            "📋 Inventários Realizados",
            len(inventarios_df)
        )
        st.markdown('</div>', unsafe_allow_html=True)

    with col3:
        total_itens_inventariados = inventarios_df['quantidade_total'].sum() if not inventarios_df.empty else 0
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric(
            "📊 Total Itens Inventariados",
            int(total_itens_inventariados) if pd.notna(total_itens_inventariados) else 0
        )
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # Lista de inventários realizados
    if not inventarios_df.empty:
        st.subheader("📋 Inventários Realizados")

        for index, inventario in inventarios_df.iterrows():
            col_info, col_btn = st.columns([4, 1])

            with col_info:
                st.markdown(f"""
                <div class="metric-card">
                    <h4 style="color: #F7931E; margin-bottom: 0.5rem;">{inventario['nome_inventario']}</h4>
                    <div style="display: flex; gap: 2rem; flex-wrap: wrap;">
                        <span><strong>📅 Data:</strong> {inventario['data_formatada']}</span>
                        <span><strong>📦 Itens:</strong> {inventario['total_itens']}</span>
                        <span><strong>📊 Quantidade Total:</strong> {int(inventario['quantidade_total']) if pd.notna(inventario['quantidade_total']) else 0}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            with col_btn:
                # Botão para exportar Excel
                excel_buffer = gerar_excel_inventario(inventario['id'], st.session_state.db_manager)
                nome_arquivo = f"inventario_rezende_energia_{inventario['responsavel'].replace(' ', '_')}_{inventario['data_formatada'].replace('/', '')}.xlsx"

                st.download_button(
                    label="📊 Excel",
                    data=excel_buffer.getvalue(),
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{inventario['id']}",
                    help=f"Baixar inventário de {inventario['responsavel']}"
                )

        st.markdown("---")
    else:
        st.info("📝 Nenhum inventário foi realizado ainda.")

    # Lista de materiais cadastrados
    if not materiais_df.empty:
        st.subheader("📦 Materiais Cadastrados")

        # Opção de exportar lista de materiais
        col_title, col_export = st.columns([3, 1])

        with col_title:
            st.write(f"Total de **{len(materiais_df)}** materiais cadastrados no sistema")

        with col_export:
            # Gerar Excel dos materiais
            buffer_materiais = io.BytesIO()
            with pd.ExcelWriter(buffer_materiais, engine='openpyxl') as writer:
                materiais_df.to_excel(writer, sheet_name='Materiais Cadastrados', index=False)
            buffer_materiais.seek(0)

            st.download_button(
                label="📊 Excel Materiais",
                data=buffer_materiais.getvalue(),
                file_name=f"materiais_rezende_energia_{datetime.now().strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Baixar lista completa de materiais"
            )

        # Exibir tabela de materiais
        st.dataframe(materiais_df, use_container_width=True)
    else:
        st.info("📦 Nenhum material cadastrado ainda. Vá para a aba 'Cadastro de Materiais' para começar.")

    st.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()