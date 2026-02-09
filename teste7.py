import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io
import math
import os

# ==========================================
# CONFIGURAÇÕES E CONSTANTES
# ==========================================

HIERARQUIA = ['SD 1', 'CB', '3º SGT', '2º SGT', '1º SGT', 'SUB TEN', 
              '2º TEN', '1º TEN', 'CAP', 'MAJ', 'TEN CEL', 'CEL']

TEMPO_MINIMO = {
    'SD 1': 5, 'CB': 3, '3º SGT': 3, '2º SGT': 3, '1º SGT': 2,
    'SUB TEN': 2, '2º TEN': 3, '1º TEN': 3, 'CAP': 3, 'MAJ': 3, 'TEN CEL': 30
}

POSTOS_COM_EXCEDENTE = ['CB', '3º SGT', '2º SGT', '2º TEN', '1º TEN', 'CAP']

# --- LIMITES DE VAGAS ---

# 1. QOA/QPC (Administrativo)
VAGAS_QOA = {
    'SD 1': 600, 'CB': 600, '3º SGT': 573, '2º SGT': 409, '1º SGT': 245,
    'SUB TEN': 96, '2º TEN': 34, '1º TEN': 29, 'CAP': 24, 'MAJ': 10, 'TEN CEL': 1, 'CEL': 9999
}

# 2. QOMT/QPMT (Condutores)
VAGAS_QOMT = {
    'SD 1': 999, 'CB': 999, '3º SGT': 999,
    '2º SGT': 68, '1º SGT': 49, 'SUB TEN': 19, 
    '2º TEN': 14, '1º TEN': 11, 'CAP': 8, 'MAJ': 4, 'TEN CEL': 2, 'CEL': 0
}

# 3. QOM/QPM (Músicos)
VAGAS_QOM = {
    'SD 1': 999, 'CB': 999,
    '3º SGT': 1, '2º SGT': 13, '1º SGT': 10, 'SUB TEN': 5, 
    '2º TEN': 11, '1º TEN': 9, 'CAP': 6, 'MAJ': 4, 'TEN CEL': 2, 'CEL': 0
}

# ==========================================
# FUNÇÕES DE LÓGICA
# ==========================================

def carregar_dados(nome_arquivo):
    """Carrega automaticamente o arquivo da pasta local."""
    if not os.path.exists(nome_arquivo):
        return None

    try:
        df = pd.read_excel(nome_arquivo)
        
        # Tratamento de erros e tipos
        cols_numericas = ['Matricula', 'Pos_Hierarquica']
        for col in cols_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        cols_datas = ['Ultima_promocao', 'Data_Admissao', 'Data_Nascimento']
        for col in cols_datas:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], dayfirst=True)

        if 'Excedente' not in df.columns:
            df['Excedente'] = ""
        df['Excedente'] = df['Excedente'].fillna("")
        
        return df
    except Exception as e:
        st.error(f"Erro ao ler {nome_arquivo}: {e}")
        return None

def get_anos(data_ref, data_origem):
    if pd.isna(data_origem): return 0
    return relativedelta(data_ref, data_origem).years

def executar_simulacao_quadro(df_input, vagas_limite_base, data_alvo, tempo_aposentadoria, 
                              matriculas_foco, vagas_extras_dict=None):
    """
    Executa a simulação para um quadro específico.
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
    sobras_por_ciclo = {} 

    for data_referencia in datas_ciclo:
        sobras_deste_ciclo = {}
        extras_hoje = {}
        
        if vagas_extras_dict and data_referencia in vagas_extras_dict:
            extras_hoje = vagas_extras_dict[data_referencia]

        # A) PROMOÇÕES
        for i in range(len(HIERARQUIA) - 1):
            posto_atual = HIERARQUIA[i]
            proximo_posto = HIERARQUIA[i+1]
            
            candidatos = df[df['Posto_Graduacao'] == posto_atual].sort_values('Pos_Hierarquica')
            limite_atual = vagas_limite_base.get(proximo_posto, 9999) + extras_hoje.get(proximo_posto, 0)
            
            ocupados_reais = len(df[(df['Posto_Graduacao'] == proximo_posto) & (df['Excedente'] != "x")])
            vagas_disponiveis = limite_atual - ocupados_reais
            
            candidatos_aptos_count = 0
            
            for idx, militar in candidatos.iterrows():
                anos_no_posto = relativedelta(data_referencia, militar['Ultima_promocao']).years
                promoveu = False
                
                # Regra Excedente
                if posto_atual in POSTOS_COM_EXCEDENTE and anos_no_posto >= 6:
                    df.at[idx, 'Posto_Graduacao'] = proximo_posto
                    df.at[idx, 'Ultima_promocao'] = data_referencia
                    df.at[idx, 'Excedente'] = "x"
                    promoveu = True
                
                # Promoção Normal
                elif anos_no_posto >= TEMPO_MINIMO.get(posto_atual, 99) and vagas_disponiveis > 0:
                    df.at[idx, 'Posto_Graduacao'] = proximo_posto
                    df.at[idx, 'Ultima_promocao'] = data_referencia
                    df.at[idx, 'Excedente'] = ""
                    vagas_disponiveis -= 1
                    promoveu = True
                
                if anos_no_posto >= TEMPO_MINIMO.get(posto_atual, 99):
                    candidatos_aptos_count += 1

                if promoveu and militar['Matricula'] in historicos:
                    historicos[militar['Matricula']].append(f"✅ {data_referencia.strftime('%d/%m/%Y')}: Promovido a {proximo_posto}")

            if vagas_disponiveis > 0:
                sobras_deste_ciclo[proximo_posto] = int(vagas_disponiveis)

        sobras_por_ciclo[data_referencia] = sobras_deste_ciclo

        # B) ABSORÇÃO
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
    st.title("🎖️ Simulador de Promoção Militar")

    # CARGA AUTOMÁTICA DOS ARQUIVOS
    df_militares = carregar_dados('militares.xlsx')
    df_condutores = carregar_dados('condutores.xlsx')
    df_musicos = carregar_dados('musicos.xlsx')

    # EXIBIÇÃO DO STATUS DOS ARQUIVOS NA SIDEBAR
    st.sidebar.header("📂 Status dos Arquivos")
    
    def check_status(df, nome):
        if df is not None:
            st.sidebar.success(f"✅ {nome} carregado")
            return True
        else:
            st.sidebar.error(f"❌ {nome} ausente")
            return False

    has_militares = check_status(df_militares, "militares.xlsx")
    has_condutores = check_status(df_condutores, "condutores.xlsx")
    has_musicos = check_status(df_musicos, "musicos.xlsx")

    st.sidebar.divider()
    st.sidebar.header("⚙️ Configuração")
    
    tipo_simulacao = st.sidebar.radio(
        "Escolha o Quadro:",
        ("QOA/QPC (Administrativo)", "QOMT/QPMT (Condutores)", "QOM/QPM (Músicos)")
    )

    # Validação baseada na escolha
    df_ativo = None

    if tipo_simulacao == "QOA/QPC (Administrativo)":
        if not has_militares:
            st.warning("O arquivo 'militares.xlsx' é necessário para esta simulação.")
            return
        df_ativo = df_militares
        # Avisos sobre migração
        if not has_condutores or not has_musicos:
            st.info("⚠️ Aviso: Para o cálculo correto de vagas migradas para o QOA, é necessário ter 'condutores.xlsx' e 'musicos.xlsx' na pasta. A simulação rodará apenas com as vagas fixas.")

    elif tipo_simulacao == "QOMT/QPMT (Condutores)":
        if not has_condutores:
            st.warning("O arquivo 'condutores.xlsx' é necessário para esta simulação.")
            return
        df_ativo = df_condutores

    elif tipo_simulacao == "QOM/QPM (Músicos)":
        if not has_musicos:
            st.warning("O arquivo 'musicos.xlsx' é necessário para esta simulação.")
            return
        df_ativo = df_musicos

    # Preparação para filtro de matrícula
    if df_ativo is not None:
        lista_matriculas = sorted(df_ativo['Matricula'].dropna().unique().astype(int))
        
        matriculas_foco = st.sidebar.multiselect(
            "Matrículas para acompanhar:",
            options=lista_matriculas,
            max_selections=5
        )

        # ALTERAÇÃO REALIZADA AQUI: max_value definido para o final de 2060
        data_alvo_input = st.sidebar.date_input(
            "Data Alvo:", 
            value=datetime(2030, 12, 31),
            max_value=datetime(2060, 12, 31)
        )
        
        tempo_aposentadoria = st.sidebar.slider("Tempo p/ Aposentadoria:", 30, 35, 35)
        
        botao_executar = st.sidebar.button("🚀 Iniciar Simulação")

        if botao_executar:
            data_alvo = pd.to_datetime(data_alvo_input)
            
            with st.spinner('Simulando...'):
                
                # SIMULAÇÃO CONDUTORES
                if tipo_simulacao == "QOMT/QPMT (Condutores)":
                    df_final, df_inativos, historicos, _ = executar_simulacao_quadro(
                        df_ativo, VAGAS_QOMT, data_alvo, tempo_aposentadoria, matriculas_foco
                    )

                # SIMULAÇÃO MÚSICOS
                elif tipo_simulacao == "QOM/QPM (Músicos)":
                    df_final, df_inativos, historicos, _ = executar_simulacao_quadro(
                        df_ativo, VAGAS_QOM, data_alvo, tempo_aposentadoria, matriculas_foco
                    )

                # SIMULAÇÃO QOA (COM MIGRAÇÃO)
                else:
                    vagas_migradas = {} 

                    # 1. Simular Condutores (Background)
                    if has_condutores:
                        _, _, _, sobras_condutores = executar_simulacao_quadro(
                            df_condutores, VAGAS_QOMT, data_alvo, tempo_aposentadoria, []
                        )
                        for data, vagas_dict in sobras_condutores.items():
                            if data not in vagas_migradas: vagas_migradas[data] = {}
                            for posto, qtd in vagas_dict.items():
                                vagas_migradas[data][posto] = vagas_migradas[data].get(posto, 0) + qtd

                    # 2. Simular Músicos (Background)
                    if has_musicos:
                        _, _, _, sobras_musicos = executar_simulacao_quadro(
                            df_musicos, VAGAS_QOM, data_alvo, tempo_aposentadoria, []
                        )
                        oficiais = ['2º TEN', '1º TEN', 'CAP', 'MAJ', 'TEN CEL']
                        pracas = ['SD 1', 'CB', '3º SGT', '2º SGT', '1º SGT', 'SUB TEN']
                        
                        for data, vagas_dict in sobras_musicos.items():
                            if data not in vagas_migradas: vagas_migradas[data] = {}
                            for posto, qtd in vagas_dict.items():
                                qtd_migrar = 0
                                if posto in pracas: qtd_migrar = qtd
                                elif posto in oficiais: qtd_migrar = math.ceil(qtd / 2)
                                
                                if qtd_migrar > 0:
                                    vagas_migradas[data][posto] = vagas_migradas[data].get(posto, 0) + qtd_migrar

                    # 3. Simular QOA Real
                    df_final, df_inativos, historicos, _ = executar_simulacao_quadro(
                        df_ativo, VAGAS_QOA, data_alvo, tempo_aposentadoria, matriculas_foco, 
                        vagas_extras_dict=vagas_migradas
                    )

                # RESULTADOS
                st.success("Simulação Concluída!")
                
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
                            
                            if m in df_final['Matricula'].values:
                                res = df_final[df_final['Matricula'] == m].iloc[0]
                                st.success(f"Status Final: {res['Posto_Graduacao']} ({'EXCEDENTE' if res['Excedente']=='x' else 'ATIVO'})")
                            else:
                                st.warning("Status Final: APOSENTADO/RESERVA")

                st.divider()
                
                def to_excel(df):
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False)
                    return out.getvalue()

                c1, c2 = st.columns(2)
                c1.download_button(f"Baixar Ativos {tipo_simulacao}", to_excel(df_final), "Ativos_Final.xlsx")
                c2.download_button(f"Baixar Inativos {tipo_simulacao}", to_excel(df_inativos), "Inativos_Final.xlsx")

if __name__ == "__main__":
    main()
