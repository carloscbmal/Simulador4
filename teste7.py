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

VAGAS_QOA = {
    'SD 1': 600, 'CB': 600, '3º SGT': 573, '2º SGT': 409, '1º SGT': 245,
    'SUB TEN': 96, '2º TEN': 34, '1º TEN': 29, 'CAP': 24, 'MAJ': 10, 'TEN CEL': 1, 'CEL': 9999
}

VAGAS_QOMT = {
    'SD 1': 999, 'CB': 999, '3º SGT': 999,
    '2º SGT': 68, '1º SGT': 49, 'SUB TEN': 19, 
    '2º TEN': 14, '1º TEN': 11, 'CAP': 8, 'MAJ': 4, 'TEN CEL': 2, 'CEL': 0
}

VAGAS_QOM = {
    'SD 1': 999, 'CB': 999,
    '3º SGT': 1, '2º SGT': 13, '1º SGT': 10, 'SUB TEN': 5, 
    '2º TEN': 11, '1º TEN': 9, 'CAP': 6, 'MAJ': 4, 'TEN CEL': 2, 'CEL': 0
}

# ==========================================
# FUNÇÕES DE LÓGICA
# ==========================================

def carregar_dados(nome_arquivo):
    if not os.path.exists(nome_arquivo):
        return None
    try:
        df = pd.read_excel(nome_arquivo)
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
    df = df_input.copy()
    data_atual = pd.to_datetime(datetime.now().strftime('%d/%m/%Y'), dayfirst=True)
    
    datas_ciclo = []
    for ano in range(data_atual.year, data_alvo.year + 1):
        for mes, dia in [(6, 26), (11, 29)]:
            d = pd.Timestamp(year=ano, month=mes, day=dia)
            if data_atual <= d <= data_alvo:
                datas_ciclo.append(d)
    datas_ciclo.sort()

    historicos = {m: [] for m in matriculas_foco} if matriculas_foco else {}
    df_inativos = pd.DataFrame()

    for data_referencia in datas_ciclo:
        sobras_deste_ciclo = {}
        extras_hoje = (vagas_extras_dict or {}).get(data_referencia, {})

        # A) PROMOÇÕES
        for i in range(len(HIERARQUIA) - 1):
            posto_atual = HIERARQUIA[i]
            proximo_posto = HIERARQUIA[i+1]
            candidatos = df[df['Posto_Graduacao'] == posto_atual].sort_values('Pos_Hierarquica')
            limite_atual = vagas_limite_base.get(proximo_posto, 9999) + extras_hoje.get(proximo_posto, 0)
            ocupados_reais = len(df[(df['Posto_Graduacao'] == proximo_posto) & (df['Excedente'] != "x")])
            vagas_disponiveis = limite_atual - ocupados_reais
            
            for idx, militar in candidatos.iterrows():
                anos_no_posto = relativedelta(data_referencia, militar['Ultima_promocao']).years
                promoveu = False
                if posto_atual in POSTOS_COM_EXCEDENTE and anos_no_posto >= 6:
                    df.at[idx, 'Posto_Graduacao'] = proximo_posto
                    df.at[idx, 'Ultima_promocao'] = data_referencia
                    df.at[idx, 'Excedente'] = "x"
                    promoveu = True
                elif anos_no_posto >= TEMPO_MINIMO.get(posto_atual, 99) and vagas_disponiveis > 0:
                    df.at[idx, 'Posto_Graduacao'] = proximo_posto
                    df.at[idx, 'Ultima_promocao'] = data_referencia
                    df.at[idx, 'Excedente'] = ""
                    vagas_disponiveis -= 1
                    promoveu = True
                if promoveu and militar['Matricula'] in historicos:
                    historicos[militar['Matricula']].append(f"✅ {data_referencia.strftime('%d/%m/%Y')}: Promovido a {proximo_posto}")

        # B) ABSORÇÃO
        for posto in HIERARQUIA:
            limite_atual = vagas_limite_base.get(posto, 9999) + extras_hoje.get(posto, 0)
            vagas_abertas = limite_atual - len(df[(df['Posto_Graduacao'] == posto) & (df['Excedente'] != "x")])
            if vagas_abertas > 0:
                excedentes = df[(df['Posto_Graduacao'] == posto) & (df['Excedente'] == "x")].sort_values('Pos_Hierarquica')
                for idx_exc in excedentes.head(int(vagas_abertas)).index:
                    df.at[idx_exc, 'Excedente'] = ""

        # C) APOSENTADORIA (CORREÇÃO TYPEERROR)
        # Convertemos explicitamente para numérico para evitar conflito de tipos no Pandas
        idade = pd.to_numeric(df['Data_Nascimento'].apply(lambda x: get_anos(data_referencia, x)))
        servico = pd.to_numeric(df['Data_Admissao'].apply(lambda x: get_anos(data_referencia, x)))
        
        mask_apo = (idade >= 63) | (servico >= tempo_aposentadoria)
        
        if mask_apo.any():
            df_inativos = pd.concat([df_inativos, df[mask_apo].copy()], ignore_index=True)
            df = df[~mask_apo].copy()

    return df, df_inativos, historicos, {}

# ==========================================
# INTERFACE
# ==========================================

def main():
    st.set_page_config(page_title="Simulador Multi-Quadros", layout="wide")
    st.title("🎖️ Simulador de Promoção Militar")

    df_militares = carregar_dados('militares.xlsx')
    df_condutores = carregar_dados('condutores.xlsx')
    df_musicos = carregar_dados('musicos.xlsx')

    st.sidebar.header("⚙️ Configuração")
    tipo_simulacao = st.sidebar.radio("Quadro:", ("QOA/QPC", "QOMT/QPMT", "QOM/QPM"))

    df_ativo = df_militares if tipo_simulacao == "QOA/QPC" else (df_condutores if tipo_simulacao == "QOMT/QPMT" else df_musicos)

    if df_ativo is not None:
        # --- DATA LIMITE PARA FINAL DE 2060 ---
        data_alvo_input = st.sidebar.date_input(
            "Data Alvo:", 
            value=datetime(2030, 12, 31),
            max_value=datetime(2060, 12, 31)
        )
        tempo_aposentadoria = st.sidebar.slider("Tempo p/ Aposentadoria:", 30, 35, 35)
        
        if st.sidebar.button("🚀 Iniciar Simulação"):
            data_alvo = pd.to_datetime(data_alvo_input)
            with st.spinner('Processando...'):
                df_final, df_inativos, _, _ = executar_simulacao_quadro(df_ativo, VAGAS_QOA if tipo_simulacao == "QOA/QPC" else (VAGAS_QOMT if tipo_simulacao == "QOMT/QPMT" else VAGAS_QOM), data_alvo, tempo_aposentadoria, [])
                
                st.success("Simulação Concluída!")
                
                def to_excel(df):
                    out = io.BytesIO()
                    df.to_excel(out, index=False, engine='xlsxwriter')
                    return out.getvalue()
                
                st.download_button("📥 Baixar Resultado (Ativos)", to_excel(df_final), "Ativos_Simulacao.xlsx")
                st.download_button("📥 Baixar Resultado (Inativos)", to_excel(df_inativos), "Inativos_Simulacao.xlsx")

if __name__ == "__main__":
    main()
