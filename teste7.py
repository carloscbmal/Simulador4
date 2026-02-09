import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io
import math

# ==========================================
# CONFIGURAÇÕES E CONSTANTES
# ==========================================

HIERARQUIA = ['SD 1', 'CB', '3º SGT', '2º SGT', '1º SGT', 'SUB TEN', 
              '2º TEN', '1º TEN', 'CAP', 'MAJ', 'TEN CEL', 'CEL']

# Tempo mínimo (anos)
TEMPO_MINIMO = {
    'SD 1': 5, 'CB': 3, '3º SGT': 3, '2º SGT': 3, '1º SGT': 2,
    'SUB TEN': 2, '2º TEN': 3, '1º TEN': 3, 'CAP': 3, 'MAJ': 3, 'TEN CEL': 30
}

# Postos que aceitam promoção automática por tempo (Excedente)
POSTOS_COM_EXCEDENTE = ['CB', '3º SGT', '2º SGT', '2º TEN', '1º TEN', 'CAP']

# --- LIMITES DE VAGAS ---

# 1. QOA/QPC (Baseado no código anterior)
VAGAS_QOA = {
    'SD 1': 600, 'CB': 600, '3º SGT': 573, '2º SGT': 409, '1º SGT': 245,
    'SUB TEN': 96, '2º TEN': 34, '1º TEN': 29, 'CAP': 24, 'MAJ': 10, 'TEN CEL': 1, 'CEL': 9999
}

# 2. QOMT/QPMT (Condutores)
VAGAS_QOMT = {
    'SD 1': 999, 'CB': 999, '3º SGT': 999, # Assumindo ilimitado na base ou definir valor
    '2º SGT': 68, '1º SGT': 49, 'SUB TEN': 19, 
    '2º TEN': 14, '1º TEN': 11, 'CAP': 8, 'MAJ': 4, 'TEN CEL': 2, 'CEL': 0
}

# 3. QOM/QPM (Músicos)
VAGAS_QOM = {
    'SD 1': 999, 'CB': 999, # Assumindo ilimitado na base
    '3º SGT': 1, '2º SGT': 13, '1º SGT': 10, 'SUB TEN': 5, 
    '2º TEN': 11, '1º TEN': 9, 'CAP': 6, 'MAJ': 4, 'TEN CEL': 2, 'CEL': 0
}

# ==========================================
# FUNÇÕES DE LÓGICA
# ==========================================

def carregar_dados(arquivo_path):
    """Carrega e trata os dados do Excel."""
    try:
        # Tenta ler do upload ou do disco local
        df = pd.read_excel(arquivo_path)
    except Exception:
        return None

    # Tratamento de erros e tipos
    cols_numericas = ['Matricula', 'Pos_Hierarquica']
    for col in cols_numericas:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    cols_datas = ['Ultima_promocao', 'Data_Admissao', 'Data_Nascimento']
    for col in cols_datas:
        df[col] = pd.to_datetime(df[col], dayfirst=True)

    if 'Excedente' not in df.columns:
        df['Excedente'] = ""
    df['Excedente'] = df['Excedente'].fillna("")
    
    return df

def get_anos(data_ref, data_origem):
    if pd.isna(data_origem): return 0
    return relativedelta(data_ref, data_origem).years

def executar_simulacao_quadro(df_input, vagas_limite_base, data_alvo, tempo_aposentadoria, 
                              matriculas_foco, vagas_extras_dict=None):
    """
    Executa a simulação para um quadro específico.
    
    Args:
        vagas_extras_dict: Dicionário opcional com vagas migradas de outros quadros.
                           Formato: {data_ciclo: {'Posto': qtd, ...}}
    Returns:
        df_final, df_inativos, historico_foco, sobras_por_ciclo
    """
    df = df_input.copy()
    data_atual = pd.to_datetime(datetime.now().strftime('%d/%m/%Y'), dayfirst=True)
    
    # Gerar Ciclos
    datas_ciclo = []
    for ano in range(data_atual.year, data_alvo.year + 1):
        for mes, dia in [(6, 26), (11, 29)]:
            d = pd.Timestamp(year=ano, month=mes, day=dia)
            if data_atual <= d <= data_alvo:
                datas_ciclo.append(d)
    datas_ciclo.sort()

    historicos = {m: [] for m in matriculas_foco} if matriculas_foco else {}
    df_inativos = pd.DataFrame()
    sobras_por_ciclo = {} # Para registrar vagas não preenchidas

    for data_referencia in datas_ciclo:
        sobras_deste_ciclo = {}
        
        # Recupera vagas extras vindas de migração (se houver) para esta data
        extras_hoje = {}
        if vagas_extras_dict and data_referencia in vagas_extras_dict:
            extras_hoje = vagas_extras_dict[data_referencia]

        # A) PROMOÇÕES (Topo -> Base)
        for i in range(len(HIERARQUIA) - 1):
            posto_atual = HIERARQUIA[i]
            proximo_posto = HIERARQUIA[i+1]
            
            candidatos = df[df['Posto_Graduacao'] == posto_atual].sort_values('Pos_Hierarquica')
            
            # Limite Total = Limite Fixo + Vagas Migradas
            limite_atual = vagas_limite_base.get(proximo_posto, 9999) + extras_hoje.get(proximo_posto, 0)
            
            # Contagem de vagas
            ocupados_reais = len(df[(df['Posto_Graduacao'] == proximo_posto) & (df['Excedente'] != "x")])
            vagas_disponiveis = limite_atual - ocupados_reais
            
            candidatos_aptos_count = 0
            
            for idx, militar in candidatos.iterrows():
                anos_no_posto = relativedelta(data_referencia, militar['Ultima_promocao']).years
                promoveu = False
                
                # Regra Excedente (6 anos)
                if posto_atual in POSTOS_COM_EXCEDENTE and anos_no_posto >= 6:
                    df.at[idx, 'Posto_Graduacao'] = proximo_posto
                    df.at[idx, 'Ultima_promocao'] = data_referencia
                    df.at[idx, 'Excedente'] = "x"
                    promoveu = True
                
                # Promoção Normal (Tem vaga e tem tempo)
                elif anos_no_posto >= TEMPO_MINIMO.get(posto_atual, 99) and vagas_disponiveis > 0:
                    df.at[idx, 'Posto_Graduacao'] = proximo_posto
                    df.at[idx, 'Ultima_promocao'] = data_referencia
                    df.at[idx, 'Excedente'] = ""
                    vagas_disponiveis -= 1 # Consome vaga
                    promoveu = True
                
                # Contagem para saber se sobraram vagas por FALTA de gente
                if anos_no_posto >= TEMPO_MINIMO.get(posto_atual, 99):
                    candidatos_aptos_count += 1

                if promoveu and militar['Matricula'] in historicos:
                    historicos[militar['Matricula']].append(f"✅ {data_referencia.strftime('%d/%m/%Y')}: Promovido a {proximo_posto}")

            # REGISTRO DE SOBRAS (Para migração)
            # Se ainda existem vagas (vagas_disponiveis > 0) e processamos todos os candidatos
            # Significa que faltou militar para ocupar.
            if vagas_disponiveis > 0:
                sobras_deste_ciclo[proximo_posto] = int(vagas_disponiveis)

        sobras_por_ciclo[data_referencia] = sobras_deste_ciclo

        # B) ABSORÇÃO DE EXCEDENTES
        for posto in HIERARQUIA:
            limite_atual = vagas_limite_base.get(posto, 9999) + extras_hoje.get(posto, 0)
            ativos_normais = len(df[(df['Posto_Graduacao'] == posto) & (df['Excedente'] != "x")])
            vagas_abertas = limite_atual - ativos_normais
            
            if vagas_abertas > 0:
                excedentes = df[(df['Posto_Graduacao'] == posto) & (df['Excedente'] == "x")].sort_values('Pos_Hierarquica')
                for idx_exc in excedentes.head(int(vagas_abertas)).index:
                    df.at[idx_exc, 'Excedente'] = ""
                    m_id = df.at[idx_exc, 'Matricula']
                    if m_id in historicos:
                        historicos[m_id].append(f"ℹ️ {data_referencia.strftime('%d/%m/%Y')}: Ocupou vaga comum em {posto}")

        # C) APOSENTADORIA
        idade = df['Data_Nascimento'].apply(lambda x: get_anos(data_referencia, x))
        servico = df['Data_Admissao'].apply(lambda x: get_anos(data_referencia, x))
        
        mask_apo = (idade >= 63) | (servico >= tempo_aposentadoria)
        
        if mask_apo.any():
            militares_aposentando = df[mask_apo]
            for m_foco in historicos:
                if m_foco in militares_aposentando['Matricula'].values:
                    historicos[m_foco].append(f"🛑 {data_referencia.strftime('%d/%m/%Y')}: APOSENTADO")
            
            df_inativos = pd.concat([df_inativos, militares_aposentando.copy()], ignore_index=True)
            df = df[~mask_apo].copy()

    return df, df_inativos, historicos, sobras_por_ciclo

# ==========================================
# INTERFACE STREAMLIT
# ==========================================

def main():
    st.set_page_config(page_title="Simulador Multi-Quadros", layout="wide")
    st.title("🎖️ Simulador de Promoção Militar (Multi-Quadros)")

    st.sidebar.header("1. Selecione o Quadro")
    tipo_simulacao = st.sidebar.radio(
        "Escolha o critério de pesquisa:",
        ("QOA/QPC (Administrativo)", "QOMT/QPMT (Condutores)", "QOM/QPM (Músicos)")
    )

    # Definição de arquivos necessários baseados na escolha
    # Se for QOA, precisa dos 3. Se for os outros, só precisa do específico.
    
    st.sidebar.header("2. Upload de Arquivos")
    
    file_militares = st.sidebar.file_uploader("Carregar 'militares.xlsx' (QOA)", type=["xlsx"])
    file_condutores = st.sidebar.file_uploader("Carregar 'condutores.xlsx' (QOMT)", type=["xlsx"])
    file_musicos = st.sidebar.file_uploader("Carregar 'musicos.xlsx' (QOM)", type=["xlsx"])

    # Carregamento dos DataFrames
    df_militares = pd.read_excel(file_militares) if file_militares else carregar_dados('militares.xlsx')
    df_condutores = pd.read_excel(file_condutores) if file_condutores else carregar_dados('condutores.xlsx')
    df_musicos = pd.read_excel(file_musicos) if file_musicos else carregar_dados('musicos.xlsx')

    # Validação de Arquivos
    df_ativo = None
    if tipo_simulacao == "QOA/QPC (Administrativo)":
        if df_militares is None:
            st.warning("Para simular QOA, o arquivo 'militares.xlsx' é obrigatório.")
            return
        if df_condutores is None or df_musicos is None:
            st.info("⚠️ Atenção: Para cálculo preciso do QOA, é ideal ter 'condutores.xlsx' e 'musicos.xlsx' para migração de vagas. A simulação rodará sem migração se não forem fornecidos.")
        df_ativo = df_militares
        
    elif tipo_simulacao == "QOMT/QPMT (Condutores)":
        if df_condutores is None:
            st.error("Por favor, carregue 'condutores.xlsx'.")
            return
        df_ativo = df_condutores

    elif tipo_simulacao == "QOM/QPM (Músicos)":
        if df_musicos is None:
            st.error("Por favor, carregue 'musicos.xlsx'.")
            return
        df_ativo = df_musicos

    # Tratamento Inicial do DF Ativo para pegar as matrículas
    df_ativo['Matricula'] = pd.to_numeric(df_ativo['Matricula'], errors='coerce')
    lista_matriculas = sorted(df_ativo['Matricula'].dropna().unique().astype(int))

    # Parâmetros de Simulação
    st.sidebar.divider()
    matriculas_foco = st.sidebar.multiselect(
        "Matrículas para acompanhar:",
        options=lista_matriculas,
        max_selections=5
    )

    data_alvo_input = st.sidebar.date_input(
        "Data Alvo:", 
        value=datetime(2030, 12, 31)
    )
    
    tempo_aposentadoria = st.sidebar.slider("Tempo p/ Aposentadoria:", 30, 35, 35)
    
    botao_executar = st.sidebar.button("🚀 Iniciar Simulação")

    if botao_executar:
        data_alvo = pd.to_datetime(data_alvo_input)
        
        with st.spinner('Processando simulação...'):
            
            # ---------------------------------------------------------
            # CENÁRIO 1 e 2: Condutores ou Músicos (Simples)
            # ---------------------------------------------------------
            if tipo_simulacao == "QOMT/QPMT (Condutores)":
                df_final, df_inativos, historicos, _ = executar_simulacao_quadro(
                    df_ativo, VAGAS_QOMT, data_alvo, tempo_aposentadoria, matriculas_foco
                )

            elif tipo_simulacao == "QOM/QPM (Músicos)":
                df_final, df_inativos, historicos, _ = executar_simulacao_quadro(
                    df_ativo, VAGAS_QOM, data_alvo, tempo_aposentadoria, matriculas_foco
                )

            # ---------------------------------------------------------
            # CENÁRIO 3: QOA/QPC (Complexo - Com Migração)
            # ---------------------------------------------------------
            else: # QOA
                vagas_migradas = {} # {data: {posto: qtd}}

                # Passo 1: Simular Condutores (se disponível) para pegar sobras
                if df_condutores is not None:
                    # Roda simulação sem focar em matrícula, apenas para pegar 'sobras'
                    _, _, _, sobras_condutores = executar_simulacao_quadro(
                        df_condutores, VAGAS_QOMT, data_alvo, tempo_aposentadoria, []
                    )
                    
                    # Migração Integral (100%)
                    for data, vagas_dict in sobras_condutores.items():
                        if data not in vagas_migradas: vagas_migradas[data] = {}
                        for posto, qtd in vagas_dict.items():
                            vagas_migradas[data][posto] = vagas_migradas[data].get(posto, 0) + qtd

                # Passo 2: Simular Músicos (se disponível) para pegar sobras
                if df_musicos is not None:
                    _, _, _, sobras_musicos = executar_simulacao_quadro(
                        df_musicos, VAGAS_QOM, data_alvo, tempo_aposentadoria, []
                    )
                    
                    # Migração Parcial (Regra do Arredondamento)
                    oficiais = ['2º TEN', '1º TEN', 'CAP', 'MAJ', 'TEN CEL']
                    pracas = ['SD 1', 'CB', '3º SGT', '2º SGT', '1º SGT', 'SUB TEN']
                    
                    for data, vagas_dict in sobras_musicos.items():
                        if data not in vagas_migradas: vagas_migradas[data] = {}
                        
                        for posto, qtd in vagas_dict.items():
                            qtd_migrar = 0
                            if posto in pracas:
                                qtd_migrar = qtd # 100%
                            elif posto in oficiais:
                                qtd_migrar = math.ceil(qtd / 2) # 50% arredondado p/ cima
                            
                            if qtd_migrar > 0:
                                vagas_migradas[data][posto] = vagas_migradas[data].get(posto, 0) + qtd_migrar

                # Passo 3: Simular QOA usando as vagas extras
                df_final, df_inativos, historicos, _ = executar_simulacao_quadro(
                    df_ativo, VAGAS_QOA, data_alvo, tempo_aposentadoria, matriculas_foco, 
                    vagas_extras_dict=vagas_migradas
                )

            # ---------------------------------------------------------
            # EXIBIÇÃO DE RESULTADOS
            # ---------------------------------------------------------
            st.success("Simulação concluída!")
            
            if matriculas_foco:
                st.subheader("📊 Histórico Individual")
                abas = st.tabs([str(m) for m in matriculas_foco])
                for i, m in enumerate(matriculas_foco):
                    with abas[i]:
                        if not historicos[m]:
                            st.info("Sem alterações.")
                        else:
                            for evento in historicos[m]:
                                st.write(evento)
                        
                        # Status Final
                        if m in df_final['Matricula'].values:
                            res = df_final[df_final['Matricula'] == m].iloc[0]
                            st.success(f"Status Final: {res['Posto_Graduacao']} ({'EXCEDENTE' if res['Excedente']=='x' else 'ATIVO'})")
                        else:
                            st.warning("Status Final: APOSENTADO/RESERVA")

            st.divider()
            
            # Downloads
            def to_excel(df):
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False)
                return out.getvalue()

            c1, c2 = st.columns(2)
            c1.download_button(f"Baixar Ativos {tipo_simulacao}", to_excel(df_final), "Ativos.xlsx")
            c2.download_button(f"Baixar Inativos {tipo_simulacao}", to_excel(df_inativos), "Inativos.xlsx")

if __name__ == "__main__":
    main()
