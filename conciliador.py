import streamlit as st
import pandas as pd
from io import BytesIO

# --- Configuração da Página ---
st.set_page_config(
    page_title="Conciliador de Vendas",
    page_icon="📊",
    layout="wide"
)

# --- Estilização Customizada (Opcional) ---
st.markdown("""
    <style>
    .main {
        padding-top: 2rem;
    }
    .stButton>button {
        width: 100%;
        margin-top: 1rem;
    }
    </style>
    """, unsafe_allow_html=True)

# --- Funções Auxiliares ---

@st.cache_data
def load_data(file):
    """Carrega dados de CSV ou Excel."""
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file)
        elif file.name.endswith(('.xls', '.xlsx')):
            return pd.read_excel(file)
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {e}")
        return None

def standardize_df(df, mapping, system_name_col, system_name_val):
    """Renomeia colunas e adiciona coluna de origem."""
    df_std = df.rename(columns=mapping)
    # Manter apenas as colunas mapeadas
    cols_to_keep = list(mapping.values())
    
    # Verificar se todas as colunas existem
    available_cols = [c for c in cols_to_keep if c in df_std.columns]
    df_std = df_std[available_cols]
    
    # Adicionar coluna de identificação
    df_std[system_name_col] = system_name_val
    
    return df_std

def get_operator_suggestion(name):
    """Retorna sugestão de padronização baseada em regras."""
    name_upper = str(name).upper()
    
    rules = {
        'TICKET': 'TICKET',
        'BIG': 'BIGCARD',
        'MAESTRO': 'MASTERCARD',
        'MASTER': 'MASTERCARD',
        'ALELO': 'ALELO',  # Alelo deve vir antes de Elo para evitar falso positivo
        'ELO': 'ELO',
        'VISA': 'VISA',
        'VR': 'VR',
        'SODEXO': 'PLUXEE',
        'CABAL': 'CABAL',
        'AMEX': 'AMEX',
        'COMPROCARD': 'COMPROCARD',
        'UPBRASIL': 'UPBRASIL',
        'PIX': 'PIX',
        'POLICARD': 'POLICARD',
        'NÃO INFORMADO': 'NÃO INFORMADO'
    }
    
    for key, value in rules.items():
        if key in name_upper:
            return value
    return ""

# --- Interface Principal ---

st.title("📊 Conciliador de Vendas")
st.markdown("Faça o upload das planilhas do **ERP** e da **Operadora** para iniciar a conciliação.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Arquivo 1: Sistema ERP")
    file_erp = st.file_uploader("Upload Planilha ERP", type=['csv', 'xlsx', 'xls'], key="erp")
    erp_name = st.selectbox("Nome do Sistema ERP", options=["Brajan", "Uniplus", "StarTwo", "Exodus", "Outro"])
    
    if erp_name == "Outro":
        erp_name = st.text_input("Digite o nome do ERP", value="Meu ERP")

with col2:
    st.subheader("📁 Arquivo 2: Vendas Operadora")
    file_operadora = st.file_uploader("Upload Planilha Operadora", type=['csv', 'xlsx', 'xls'], key="operadora")
    operadora_source = st.selectbox("Fonte dos Dados", options=["Sistema Conciliadora", "Outros"])
    
    if operadora_source == "Outros":
        operadora_source = st.text_input("Digite a Fonte dos Dados", value="Portal Operadora")

# --- Lógica de Processamento ---

# --- Inicialização do Session State ---
if 'step' not in st.session_state:
    st.session_state.step = 0
if 'df_erp_processed' not in st.session_state:
    st.session_state.df_erp_processed = None
if 'df_op_processed' not in st.session_state:
    st.session_state.df_op_processed = None

# --- Lógica de Processamento ---

if file_erp and file_operadora:
    df_erp = load_data(file_erp)
    df_operadora = load_data(file_operadora)

    if df_erp is not None and df_operadora is not None:
        st.divider()
        
        # --- ETAPA 0: Configuração e Processamento Inicial ---
        if st.session_state.step == 0:
            st.header("🛠️ Mapeamento de Colunas")
            
            # Colunas Padrão Requeridas (ERP)
            required_columns_erp = {
                "data_venda": "Data da Venda",
                "valor_venda": "Valor da Venda",
                "operadora": "Operadora",
                "nsu": "NSU",
                "parcelas": "Parcelas"
            }

            # Colunas Padrão Requeridas (Operadora) - Sem Parcelas
            required_columns_op = {
                "data_venda": "Data da Venda",
                "valor_venda": "Valor da Venda",
                "operadora": "Operadora",
                "nsu": "NSU"
            }

            # Mapeamentos Automáticos (Padrão)
            ERP_MAPPINGS = {
                "Brajan": {
                    "data_venda": "Data Efetiva",
                    "valor_venda": "Vr. Venda",
                    "operadora": "Bandeira",
                    "nsu": "Autorização",
                    "parcelas": "Pc"
                }
            }

            OPERADORA_MAPPINGS = {
                "Sistema Conciliadora": {
                    "data_venda": "Venda",
                    "valor_venda": "Bruto",
                    "operadora": "Produto",
                    "nsu": "NSU"
                }
            }

            col_map1, col_map2 = st.columns(2)

            mapping_erp = {}
            mapping_operadora = {}

            # Mapeamento ERP
            with col_map1:
                st.subheader(f"Mapear: {erp_name}")
                st.info("Selecione a coluna correspondente no seu arquivo.")
                
                # Obter mapeamento padrão se existir
                current_erp_map = ERP_MAPPINGS.get(erp_name, {})
                
                for key, label in required_columns_erp.items():
                    # Tentar encontrar o índice da coluna padrão
                    default_col_name = current_erp_map.get(key)
                    default_index = 0
                    
                    if default_col_name and default_col_name in df_erp.columns:
                        default_index = list(df_erp.columns).index(default_col_name)
                    
                    mapping_erp[key] = st.selectbox(
                        f"{label} ({erp_name})",
                        options=df_erp.columns,
                        index=default_index,
                        key=f"erp_{key}"
                    )

            # Mapeamento Operadora
            with col_map2:
                st.subheader(f"Mapear: {operadora_source}")
                st.info("Selecione a coluna correspondente no seu arquivo.")
                
                # Obter mapeamento padrão se existir
                current_op_map = OPERADORA_MAPPINGS.get(operadora_source, {})

                for key, label in required_columns_op.items():
                    # Tentar encontrar o índice da coluna padrão
                    default_col_name = current_op_map.get(key)
                    default_index = 0
                    
                    if default_col_name and default_col_name in df_operadora.columns:
                        default_index = list(df_operadora.columns).index(default_col_name)

                    mapping_operadora[key] = st.selectbox(
                        f"{label} ({operadora_source})",
                        options=df_operadora.columns,
                        index=default_index,
                        key=f"op_{key}"
                    )

            st.divider()
            st.divider()
            st.header("⚙️ Opções de Conciliação")
            
            col_opt1, col_opt2 = st.columns(2)
            with col_opt1:
                check_erp_filter = st.checkbox("Filtrar Vendas Parceladas (Apenas Parcelas = 1)", value=True)
                check_date_fmt = st.checkbox("Padronizar formato de datas (DD/MM/AAAA)", value=True)
            with col_opt2:
                check_op_agg = st.checkbox("Unificar Vendas por NSU (Somar Valores)", value=True)
                check_val_fmt = st.checkbox("Padronizar formato de valores (R$ 0.00)", value=True)
            
            check_depara = st.checkbox("Padronizar Operadoras (De-Para)", value=True)

            st.divider()
            st.header("📅 Filtros e Processamento")

            # Filtro de Data
            col_date1, col_date2 = st.columns(2)
            with col_date1:
                start_date = st.date_input("Data Início")
            with col_date2:
                end_date = st.date_input("Data Fim")

            if st.button("🚀 Processar Dados", type="primary"):
                try:
                    # Inverter o dicionário para usar no rename (De: Original -> Para: Padrão)
                    rename_map_erp = {v: k for k, v in mapping_erp.items()}
                    rename_map_op = {v: k for k, v in mapping_operadora.items()}

                    # Padronização
                    df_erp_std = standardize_df(df_erp, rename_map_erp, "origem", erp_name)
                    df_op_std = standardize_df(df_operadora, rename_map_op, "origem", operadora_source)

                    # --- Regras de Negócio ---
                    
                    # 1. ERP: Filtrar Parcelas == '1'
                    if check_erp_filter:
                        # Converter para string e remover decimais (.0) caso existam
                        if 'parcelas' in df_erp_std.columns:
                            df_erp_std['parcelas'] = df_erp_std['parcelas'].astype(str).str.replace(r'\.0$', '', regex=True)
                            df_erp_std = df_erp_std[df_erp_std['parcelas'] == '1']

                    # 2. Operadora: Unificar por NSU (Somar Valor)
                    df_op_std['valor_venda'] = pd.to_numeric(df_op_std['valor_venda'], errors='coerce').fillna(0)
                    
                    if check_op_agg:
                        if 'nsu' in df_op_std.columns:
                            df_op_std = df_op_std.groupby('nsu', as_index=False).agg({
                                'data_venda': 'first',
                                'valor_venda': 'sum',
                                'operadora': 'first',
                                'origem': 'first'
                            })

                    # Conversão de Data para filtragem
                    df_erp_std['data_venda'] = pd.to_datetime(df_erp_std['data_venda'], errors='coerce')
                    df_op_std['data_venda'] = pd.to_datetime(df_op_std['data_venda'], errors='coerce')

                    # Filtragem por Data
                    start_date = pd.to_datetime(start_date)
                    end_date = pd.to_datetime(end_date)

                    mask_erp = (df_erp_std['data_venda'] >= start_date) & (df_erp_std['data_venda'] <= end_date)
                    mask_op = (df_op_std['data_venda'] >= start_date) & (df_op_std['data_venda'] <= end_date)

                    df_erp_final = df_erp_std.loc[mask_erp].copy()
                    df_op_final = df_op_std.loc[mask_op].copy()

                    # --- Formatação de Dados (Opcional) ---
                    # Aplicar formatação APÓS filtragem e ANTES de salvar no session state
                    # Isso converte para string, o que é bom para visualização e merge exato
                    
                    if check_date_fmt:
                        df_erp_final['data_venda'] = df_erp_final['data_venda'].dt.strftime('%d/%m/%Y')
                        df_op_final['data_venda'] = df_op_final['data_venda'].dt.strftime('%d/%m/%Y')
                    
                    if check_val_fmt:
                        # Função auxiliar para formatar moeda
                        def format_currency(val):
                            return f"R$ {val:.2f}"
                        
                        # Garantir que é numérico antes de formatar
                        df_erp_final['valor_venda'] = pd.to_numeric(df_erp_final['valor_venda'], errors='coerce').fillna(0)
                        df_op_final['valor_venda'] = pd.to_numeric(df_op_final['valor_venda'], errors='coerce').fillna(0)
                        
                        df_erp_final['valor_venda'] = df_erp_final['valor_venda'].apply(format_currency)
                        df_op_final['valor_venda'] = df_op_final['valor_venda'].apply(format_currency)

                    # Salvar no Session State e avançar
                    st.session_state.df_erp_processed = df_erp_final
                    st.session_state.df_op_processed = df_op_final
                    
                    # Lógica de Pulo do De-Para
                    if check_depara:
                        st.session_state.step = 1
                    else:
                        st.session_state.step = 2
                    
                    st.rerun()

                except Exception as e:
                    st.error(f"Ocorreu um erro durante o processamento: {e}")
                    st.exception(e)

        # --- ETAPA 1: De-Para de Operadoras ---
        elif st.session_state.step == 1:
            st.header("🔄 Padronização de Operadoras (De-Para)")
            st.info("Identifique e padronize os nomes das operadoras encontradas. Deixe em branco para manter o original.")

            df_erp_final = st.session_state.df_erp_processed
            df_op_final = st.session_state.df_op_processed

            # Obter valores únicos
            unique_ops_erp = sorted(df_erp_final['operadora'].astype(str).unique())
            unique_ops_op = sorted(df_op_final['operadora'].astype(str).unique())
            
            # Combinar valores únicos de ambas as planilhas para facilitar
            all_unique_ops = sorted(list(set(unique_ops_erp + unique_ops_op)))

            st.write(f"Foram encontradas {len(all_unique_ops)} variações de operadoras.")

            # Formulário para De-Para
            with st.form("de_para_form"):
                col_depara1, col_depara2 = st.columns(2)
                
                mapping_changes = {}
                
                # Distribuir inputs em duas colunas
                for i, op_name in enumerate(all_unique_ops):
                    col = col_depara1 if i % 2 == 0 else col_depara2
                    
                    # Obter sugestão automática
                    suggestion = get_operator_suggestion(op_name)
                    
                    new_val = col.text_input(f"De: '{op_name}' Para:", value=suggestion, key=f"depara_{i}")
                    if new_val.strip():
                        mapping_changes[op_name] = new_val.strip()

                submitted = st.form_submit_button("✅ Finalizar e Visualizar Resultados")
                
                if submitted:
                    # Aplicar mudanças
                    if mapping_changes:
                        df_erp_final['operadora'] = df_erp_final['operadora'].astype(str).replace(mapping_changes)
                        df_op_final['operadora'] = df_op_final['operadora'].astype(str).replace(mapping_changes)
                        st.success(f"{len(mapping_changes)} padronizações aplicadas!")
                    
                    # Atualizar Session State
                    st.session_state.df_erp_processed = df_erp_final
                    st.session_state.df_op_processed = df_op_final
                    st.session_state.step = 2
                    st.rerun()

        # --- ETAPA 2: Resultados da Conciliação ---
        elif st.session_state.step == 2:
            st.header("📊 Resultados da Conciliação")
            st.info("Avaliação considerando as variáveis: 'Data da Venda', 'Valor da Venda' e 'Operadora'")

            df_erp = st.session_state.df_erp_processed.copy()
            df_op = st.session_state.df_op_processed.copy()

            # --- Lógica de Conciliação 1-para-1 ---
            
            # Definir colunas chave para conciliação
            match_cols = ['data_venda', 'valor_venda', 'operadora']
            
            # Criar ID único para tratar duplicatas (cumcount)
            # Isso garante que se houver 2 vendas iguais no ERP e 2 na Operadora, elas sejam casadas 1-a-1
            df_erp['match_id'] = df_erp.groupby(match_cols).cumcount()
            df_op['match_id'] = df_op.groupby(match_cols).cumcount()

            # Realizar o Merge (Left Join: ERP -> Operadora)
            # Indicador=True cria a coluna '_merge' que diz se encontrou ou não
            df_merged = pd.merge(
                df_erp,
                df_op[match_cols + ['match_id']], # Trazer apenas as chaves para verificar existência
                on=match_cols + ['match_id'],
                how='left',
                indicator=True
            )

            # Definir Status
            # both = Encontrado em ambos (CONCILIADO)
            # left_only = Encontrado apenas no ERP (DIVERGENTE)
            df_merged['status_conciliacao'] = df_merged['_merge'].map({
                'both': 'CONCILIADO',
                'left_only': 'DIVERGENTE'
            })

            # Remover colunas auxiliares
            df_merged = df_merged.drop(columns=['match_id', '_merge'])

            # Exibir Métricas
            total_vendas = len(df_merged)
            total_conciliado = len(df_merged[df_merged['status_conciliacao'] == 'CONCILIADO'])
            total_divergente = len(df_merged[df_merged['status_conciliacao'] == 'DIVERGENTE'])
            
            col_met1, col_met2, col_met3 = st.columns(3)
            col_met1.metric("Total de Vendas (ERP)", total_vendas)
            col_met2.metric("Conciliadas", total_conciliado)
            col_met3.metric("Divergentes", total_divergente)

            st.divider()
            st.subheader("📊 Resultado por adquirente - No período")

            # --- Cálculo do Resumo por Operadora ---
            # Precisamos garantir que os valores sejam numéricos para somar
            # Se a formatação foi aplicada, 'valor_venda' será string (R$ ...). Precisamos limpar.
            
            def clean_currency(val):
                if isinstance(val, str):
                    # Remover 'R$', espaços e substituir vírgula por ponto se necessário (embora nosso format use ponto)
                    clean_val = val.replace('R$', '').replace(' ', '')
                    return float(clean_val)
                return float(val)

            df_erp_calc = df_erp.copy()
            df_op_calc = df_op.copy()

            df_erp_calc['valor_venda_num'] = df_erp_calc['valor_venda'].apply(clean_currency)
            df_op_calc['valor_venda_num'] = df_op_calc['valor_venda'].apply(clean_currency)

            # Agrupamento
            erp_summary = df_erp_calc.groupby('operadora')['valor_venda_num'].sum().reset_index()
            op_summary = df_op_calc.groupby('operadora')['valor_venda_num'].sum().reset_index()

            # Merge dos resumos
            summary_merged = pd.merge(erp_summary, op_summary, on='operadora', how='outer', suffixes=('_erp', '_op')).fillna(0)

            # Cálculo da Divergência
            summary_merged['divergencia'] = summary_merged['valor_venda_num_erp'] - summary_merged['valor_venda_num_op']

            # Renomear e Formatar para Exibição
            summary_display = summary_merged.rename(columns={
                'operadora': 'Operadora',
                'valor_venda_num_erp': 'Valor de venda total no ERP',
                'valor_venda_num_op': 'Valor de venda total na Operadora',
                'divergencia': 'Divergência de valor'
            })

            # Aplicar formatação de moeda para exibição
            def fmt_currency(x): return f"R$ {x:.2f}"
            
            summary_display['Valor de venda total no ERP'] = summary_display['Valor de venda total no ERP'].apply(fmt_currency)
            summary_display['Valor de venda total na Operadora'] = summary_display['Valor de venda total na Operadora'].apply(fmt_currency)
            summary_display['Divergência de valor'] = summary_display['Divergência de valor'].apply(fmt_currency)

            st.dataframe(summary_display, use_container_width=True)

            st.divider()
            st.subheader("📊 Resultado por adquirente - Diário")

            # --- Cálculo do Resumo Diário por Operadora ---
            # Agrupamento por Data e Operadora
            erp_daily = df_erp_calc.groupby(['data_venda', 'operadora'])['valor_venda_num'].sum().reset_index()
            op_daily = df_op_calc.groupby(['data_venda', 'operadora'])['valor_venda_num'].sum().reset_index()

            # Merge dos resumos diários
            daily_merged = pd.merge(erp_daily, op_daily, on=['data_venda', 'operadora'], how='outer', suffixes=('_erp', '_op')).fillna(0)

            # Cálculo da Divergência
            daily_merged['divergencia'] = daily_merged['valor_venda_num_erp'] - daily_merged['valor_venda_num_op']

            # Filtros para o Resumo Diário
            col_filter1, col_filter2 = st.columns(2)
            
            # Obter valores únicos para os filtros
            unique_dates = sorted(daily_merged['data_venda'].astype(str).unique())
            unique_ops_daily = sorted(daily_merged['operadora'].astype(str).unique())

            with col_filter1:
                selected_dates = st.multiselect("Filtrar por Data:", unique_dates, default=unique_dates)
            with col_filter2:
                selected_ops = st.multiselect("Filtrar por Operadora:", unique_ops_daily, default=unique_ops_daily)

            # Aplicar Filtros
            daily_filtered = daily_merged.copy()
            
            # Converter data para string para filtrar corretamente com o multiselect
            daily_filtered['data_venda_str'] = daily_filtered['data_venda'].astype(str)
            
            if selected_dates:
                daily_filtered = daily_filtered[daily_filtered['data_venda_str'].isin(selected_dates)]
            if selected_ops:
                daily_filtered = daily_filtered[daily_filtered['operadora'].isin(selected_ops)]

            # Renomear e Formatar para Exibição
            daily_display = daily_filtered.rename(columns={
                'data_venda': 'Data',
                'operadora': 'Operadora',
                'valor_venda_num_erp': 'Valor de venda total no ERP',
                'valor_venda_num_op': 'Valor de venda total na Operadora',
                'divergencia': 'Divergência de valor'
            })

            # Selecionar colunas finais
            cols_to_show = ['Data', 'Operadora', 'Valor de venda total no ERP', 'Valor de venda total na Operadora', 'Divergência de valor']
            daily_display = daily_display[cols_to_show]

            # Aplicar formatação de moeda
            daily_display['Valor de venda total no ERP'] = daily_display['Valor de venda total no ERP'].apply(fmt_currency)
            daily_display['Valor de venda total na Operadora'] = daily_display['Valor de venda total na Operadora'].apply(fmt_currency)
            daily_display['Divergência de valor'] = daily_display['Divergência de valor'].apply(fmt_currency)

            st.dataframe(daily_display, use_container_width=True)

            st.divider()
            st.subheader("Detalhamento")
            
            # Filtro de visualização
            filter_status = st.multiselect("Filtrar por Status:", ['CONCILIADO', 'DIVERGENTE'], default=['CONCILIADO', 'DIVERGENTE'])
            
            if filter_status:
                st.dataframe(df_merged[df_merged['status_conciliacao'].isin(filter_status)])
            else:
                st.dataframe(df_merged)

            st.divider()
            with st.expander("📂 Visualizar Dados Processados (Detalhes)"):
                tab_erp, tab_op = st.tabs(["Dados ERP (Processados)", "Dados Operadora (Processados)"])
                
                with tab_erp:
                    st.dataframe(df_erp)
                    st.caption(f"Total de registros: {len(df_erp)}")
                
                with tab_op:
                    st.dataframe(df_op)
                    st.caption(f"Total de registros: {len(df_op)}")

            if st.button("🔄 Reiniciar Processo"):
                st.session_state.step = 0
                st.session_state.df_erp_processed = None
                st.session_state.df_op_processed = None
                st.rerun()

        # --- ETAPA 2: Resultados Finais ---
        elif st.session_state.step == 2:
            st.header("📊 Resultados da Conciliação")
            
            df_erp_final = st.session_state.df_erp_processed
            df_op_final = st.session_state.df_op_processed

            # Métricas
            m1, m2 = st.columns(2)
            m1.metric(label=f"Linhas {erp_name}", value=len(df_erp_final))
            m2.metric(label=f"Linhas {operadora_source}", value=len(df_op_final))

            # Previews
            tab1, tab2 = st.tabs([f"Dados {erp_name}", f"Dados {operadora_source}"])
            
            with tab1:
                st.dataframe(df_erp_final, use_container_width=True)
            
            with tab2:
                st.dataframe(df_op_final, use_container_width=True)
            
            if st.button("🔄 Reiniciar Processo"):
                st.session_state.step = 0
                st.session_state.df_erp_processed = None
                st.session_state.df_op_processed = None
                st.rerun()

    else:
        st.warning("Por favor, faça o upload de arquivos válidos.")

else:
    st.info("👆 Aguardando upload dos arquivos para exibir opções de mapeamento.")
