import pandas as pd
import streamlit as st
import os
import openpyxl

# Diretório onde os arquivos Excel estão localizados
dir_path = os.path.dirname(os.path.realpath(__file__))

# Caminhos relativos para os arquivos Excel
excel_file_rh = os.path.join(dir_path, 'Diretoria Operações 11.02.24.xlsx')
excel_file_base = os.path.join(dir_path, 'De-Para Bases_2024.xlsx')
excel_file_mps = os.path.join(dir_path, 'MPS Vigente 14_06_2023_homologação inicial.xlsx')

# Verifica se os arquivos Excel existem
if not os.path.exists(excel_file_rh) or not os.path.exists(excel_file_base) or not os.path.exists(excel_file_mps):
    st.error('Erro: Um ou mais arquivos Excel não foram encontrados.')
else:
    # Lê os arquivos Excel
    df_RH = pd.read_excel(excel_file_rh)
    df_base = pd.read_excel(excel_file_base, sheet_name='RH ORIGINAL')
    df_MPS = pd.read_excel(excel_file_mps, sheet_name='BASE ROSA', index_col=0)

    Nivel_MPS = ['Básico', 'Avançado']
    Treinamento = ['MPS Avançado – Capacitação Avançada para EPS - Emissor de Permissão de Serviço e Formação de SEEC - Supervisor de Entrada em Espaço Confinado - NR33']

    nome_coluna = df_MPS.columns[0]
    df_RH['CS'] = df_RH['CS'].astype(str)

    opcao_text_input = st.sidebar.radio("O colaborador é terceiro?", options=["Sim", "Não"], index=1)

    if opcao_text_input == 'Sim':
        CS_colaborador = st.sidebar.text_input('Insira somente os números do TR do colaborador:', key='cs_text_input')
        CS_CORRIGIDO = 'CS' + CS_colaborador
        Filtro_nome = st.sidebar.text_input("Insira o nome:", key='nome_text_input')
        cargo_colaborador = sorted(df_RH['Descrição Cargo'].unique())
        Filtro_cargo = st.sidebar.selectbox("Selecione o cargo:", cargo_colaborador)
        base_colaborador = df_base['RH'].unique()
        Filtro_base = st.sidebar.selectbox("Selecione a base:", base_colaborador)
        nucleo_colaborador = df_base.loc[df_base['RH'] == Filtro_base, 'NUCLEO'].values.tolist()
        Filtro_nucleo = st.sidebar.selectbox("Selecione o núcleo:", nucleo_colaborador)
        Gerencia_colaborador = df_base.loc[df_base['NUCLEO'] == Filtro_nucleo, 'GERENCIA'].values.tolist()
        Filtro_gerencia = st.sidebar.selectbox("Selecione a gerência:", Gerencia_colaborador)
        Filtro_Treinamento = st.sidebar.selectbox("Selecione o Treinamento:", Treinamento)
        Filtro_nivel = st.sidebar.selectbox("Selecione o Treinamento:", Nivel_MPS)
        base_ur = None
        if Filtro_base:
                base_ur = df_base.loc[df_base['RH'] == Filtro_base, 'UR'].iloc[0]
    else:
            CS_colaborador = st.sidebar.text_input('Insira o CS do colaborador:')
            nome_colaborador = sorted(df_RH['Nome Completo'])
            Nome_filtro = df_RH.loc[df_RH['CS'] == CS_colaborador, 'Nome Completo'].values.tolist()
            Filtro_nome = st.sidebar.selectbox("Selecione o nome:", Nome_filtro)
            CS_CORRIGIDO = 'CS' + CS_colaborador
            cargo_colaborador = df_RH.loc[df_RH['Nome Completo'] == Filtro_nome, 'Descrição Cargo'].values.tolist()
            Filtro_cargo = st.sidebar.selectbox("Selecione o cargo:", cargo_colaborador)

            if CS_colaborador:
                Filtro_base = df_RH.loc[df_RH['CS'] == CS_colaborador, 'Unidade'].values.tolist()[0]
            else:
                base_colaborador = df_RH['Unidade'].unique()
                Filtro_base = st.sidebar.selectbox("Selecione a base:", base_colaborador)

            base_ur = None
            if Filtro_base:
                base_ur = df_base.loc[df_base['RH'] == Filtro_base, 'UR'].iloc[0]
            nucleo_colaborador = df_base.loc[df_base['RH'] == Filtro_base, 'NUCLEO'].values.tolist()
            Filtro_nucleo = st.sidebar.selectbox("Selecione o núcleo:", nucleo_colaborador)
            Gerencia_colaborador = df_base.loc[df_base['NUCLEO'] == Filtro_nucleo, 'GERENCIA'].values.tolist()
            Filtro_gerencia = st.sidebar.selectbox("Selecione a gerência:", Gerencia_colaborador)
            Filtro_Treinamento = st.sidebar.selectbox("Selecione o Treinamento:", Treinamento)
            Filtro_nivel = st.sidebar.selectbox("Selecione o Treinamento:", Nivel_MPS)

    if st.sidebar.button('Adicionar nome'):
        colunas_desejadas = ['Usuário - Nome do Usuário', 'Campos Calculados - Nome e Sobrenome do Usuário', 'Usuário - Cargo', 'Treinamento - Título do Treinamento', 'Usuário - Unidade', 'Usuário - Localização', 'Gerência', 'Nível']
        lista_adicionar = [CS_CORRIGIDO, Filtro_nome, Filtro_cargo, Filtro_Treinamento, Filtro_base, base_ur, Filtro_gerencia, Filtro_nivel]
        nova_linha = pd.DataFrame([lista_adicionar], columns=colunas_desejadas)
        df_MPS = pd.concat([df_MPS, nova_linha])
        df_MPS.to_excel(excel_file_mps, sheet_name='BASE ROSA')
        st.success('Nome adicionado')

    # Exibindo o DataFrame df_MPS
    df_MPS
