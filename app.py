import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import openpyxl
from io import BytesIO
import requests

st.set_page_config(page_title="Análise de Serviços Técnicos", layout="wide")

FORMAS_PAGAMENTO_VALIDAS = [
    'Check', 'American Express', 'Apple Pay', 'Discover',
    'Master Card', 'Visa', 'Zelle', 'Cash', 'Invoice'
]

INVALID_CLIENTS = ['SERVICES IN:', 'BNS PROFIT:', 'Total']


def format_currency(value):
    """Formata valores como moeda USD com 2 casas decimais"""
    if pd.isna(value):
        return None
    return f"${value:,.2f}"


def process_spreadsheet(file):
    all_weeks_data = {}
    if isinstance(file, str) and file.startswith('http'):
        response = requests.get(file)
        file = BytesIO(response.content)
    elif isinstance(file, BytesIO):
        file.seek(0)

    xls = pd.ExcelFile(file)
    for sheet_name in xls.sheet_names:
        if sheet_name.startswith('WEEK'):
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            technician_blocks = []
            current_block = []
            collecting = False
            for idx, row in df.iterrows():
                if any('NAME:' in str(cell) for cell in row.values):
                    if current_block:
                        technician_blocks.append(current_block)
                        current_block = []
                    collecting = True
                if collecting:
                    current_block.append(row)
            if current_block:
                technician_blocks.append(current_block)

            week_data = []
            for block in technician_blocks:
                name_row = next((row for row in block if any('NAME:' in str(cell) for cell in row.values)), None)

                if name_row is None:
                    continue
                name_col = next(
                    (i for i, cell in enumerate(name_row.values) if isinstance(cell, str) and 'NAME:' in cell), None)

                technician_info = {
                    'Semana': sheet_name,
                    'Nome': name_row[name_col + 1] if name_col is not None else None,
                    'Categoria': name_row[name_col + 3] if name_col is not None else None,
                    'Origem': name_row[name_col + 5] if name_col is not None and 'From:' in str(
                        name_row[name_col + 4]) else None
                }

                header_row = next((i for i, row in enumerate(block) if all(
                    keyword in str(row.values) for keyword in ['Schedule', 'DATE', 'SERVICE'])), None)
                if header_row is None:
                    continue

                days_data = []
                for i in range(header_row + 1, len(block)):
                    day_row = block[i]
                    for day_idx, day_col in enumerate(
                            [(1, 9), (10, 18), (19, 27), (28, 36), (37, 45), (46, 54), (55, 63)]):
                        start_col, end_col = day_col
                        day_data = day_row[start_col:end_col + 1].values
                        client_name = str(day_data[0]).strip() if pd.notna(day_data[0]) else ''

                        # Ignorar linhas inválidas
                        if not client_name or client_name.upper() in [c.upper() for c in INVALID_CLIENTS]:
                            continue

                        # Verificar se é um atendimento válido (tem valor de serviço)
                        if pd.notna(day_data[2]) and str(day_data[2]).strip() and str(day_data[2]).strip() != 'nan':
                            try:
                                service_value = float(day_data[2])
                            except:
                                service_value = np.nan
                            if not np.isnan(service_value):
                                pagamento = day_data[5] if pd.notna(day_data[5]) and str(
                                    day_data[5]).strip() in FORMAS_PAGAMENTO_VALIDAS else None
                                tip_value = float(day_data[3]) if pd.notna(day_data[3]) else 0
                                day_info = {
                                    'Dia': ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'][
                                        day_idx],
                                    'Data': day_data[1],
                                    'Cliente': client_name,
                                    'Serviço': service_value,
                                    'Gorjeta': tip_value,
                                    'Pets': day_data[4] if pd.notna(day_data[4]) else 0,
                                    'Pagamento': pagamento,
                                    'ID Pagamento': day_data[6] if pd.notna(day_data[6]) else None,
                                    'Verificado': day_data[7] if pd.notna(day_data[7]) else False,
                                    'Realizado': True
                                }
                                days_data.append({**technician_info, **day_info})
                        elif pd.notna(day_data[0]):
                            if client_name.upper() in [c.upper() for c in INVALID_CLIENTS]:
                                continue
                            day_info = {
                                'Dia': ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'][day_idx],
                                'Data': day_data[1],
                                'Cliente': client_name,
                                'Serviço': 0,
                                'Gorjeta': 0,
                                'Pets': 0,
                                'Pagamento': None,
                                'ID Pagamento': None,
                                'Verificado': False,
                                'Realizado': False
                            }
                            days_data.append({**technician_info, **day_info})

                week_data.extend(days_data)
            if week_data:
                all_weeks_data[sheet_name] = pd.DataFrame(week_data)

    if all_weeks_data:
        combined_data = pd.concat(all_weeks_data.values(), ignore_index=True)
        combined_data['Data'] = pd.to_datetime(combined_data['Data'], errors='coerce')
        combined_data['Serviço'] = pd.to_numeric(combined_data['Serviço'], errors='coerce')
        combined_data['Gorjeta'] = pd.to_numeric(combined_data['Gorjeta'], errors='coerce').fillna(0)
        combined_data['Pets'] = pd.to_numeric(combined_data['Pets'], errors='coerce').fillna(0)
        combined_data = combined_data.dropna(subset=['Data'])
        combined_data = combined_data[
            ~combined_data['Cliente'].astype(str).str.strip().str.upper().isin([c.upper() for c in INVALID_CLIENTS])]
        return combined_data
    return pd.DataFrame()


# Configuração da sidebar
st.sidebar.markdown("""
<div style="text-align: center; margin-bottom: 20px;">
    <img src="https://i.imgur.com/tlb2Bcy.png" 
         alt="Logo da Empresa" 
</div>
""", unsafe_allow_html=True)

st.sidebar.title("🔍 Filtros")

# Main content
st.title("📊 BNS - PORTAL DE ANÁLISES DE DADOS FINANCEIROS")

uploaded_files = st.sidebar.file_uploader("Carregue uma ou mais planilhas Excel", type=['xlsx'],
                                          accept_multiple_files=True)
url_input = st.sidebar.text_input("Ou cole a URL de uma planilha online")

all_dataframes = []
if uploaded_files or url_input:
    files_to_process = uploaded_files if uploaded_files else [url_input]
    for file in files_to_process:
        df = process_spreadsheet(file)
        if not df.empty:
            all_dataframes.append(df)

    if all_dataframes:
        data = pd.concat(all_dataframes, ignore_index=True)
        data = data[data['Nome'].notna() & (data['Nome'].astype(str).str.strip() != '')]
        data = data[~data['Cliente'].astype(str).str.strip().str.upper().isin([c.upper() for c in INVALID_CLIENTS])]

        # Filtros na sidebar
        st.sidebar.header("Filtrar por:")

        # Filtrar por abas (Semana)
        weeks = data['Semana'].unique()
        selected_weeks = st.sidebar.multiselect(
            "Selecione as abas (semanas):",
            options=weeks,
            default=list(weeks)
        )

        # Filtrar por técnico
        technicians = data['Nome'].unique()
        selected_techs = st.sidebar.multiselect(
            "Selecione os técnicos:",
            options=technicians,
            default=list(technicians)
        )

        # Filtrar por categoria
        categories = data['Categoria'].unique()
        selected_categories = st.sidebar.multiselect(
            "Selecione as categorias:",
            options=categories,
            default=list(categories)
        )

        # Aplicar filtros
        if selected_weeks:
            data = data[data['Semana'].isin(selected_weeks)]
        if selected_techs:
            data = data[data['Nome'].isin(selected_techs)]
        if selected_categories:
            data = data[data['Categoria'].isin(selected_categories)]

        if data.empty:
            st.warning("Nenhum dado encontrado com os filtros selecionados.")
            st.stop()

        st.success("✅ Planilhas processadas com sucesso!")

        if st.checkbox("🔍 Mostrar dados brutos"):
            st.dataframe(data)

        st.header("📈 Métricas Gerais")
        completed_services = data[data['Realizado']]
        not_completed = data[(data['Realizado'] == False) & (data['Cliente'].notna())]

        # Corrigido: Calcular dias trabalhados corretamente (1 por dia com atendimento, por técnico por semana)
        dias_trabalhados = completed_services.groupby(['Nome', 'Semana', 'Data']).size().reset_index()
        dias_trabalhados = dias_trabalhados.groupby(['Nome', 'Semana']).size().reset_index(name='Dias Trabalhados')

        # Agrupar por técnico e semana para calcular totais
        weekly_totals = completed_services.groupby(['Nome', 'Semana', 'Categoria']).agg({
            'Serviço': 'sum',
            'Gorjeta': 'sum',
            'Dia': 'count'
        }).reset_index()

        # Juntar com os dias trabalhados corretamente calculados
        weekly_totals = pd.merge(weekly_totals, dias_trabalhados, on=['Nome', 'Semana'], how='left')


        # Função para calcular pagamento e lucro baseado na categoria (agora considerando semana)
        def calcular_pagamento_semanal(row):
            categoria = row['Categoria']
            servico = row['Serviço']
            gorjeta = row['Gorjeta']
            dias_trabalhados = row['Dias Trabalhados']

            if categoria == 'Registering':
                pagamento = 0.00
                lucro = servico + gorjeta
            elif categoria == 'Technician':
                pagamento = servico * 0.20 + gorjeta
                lucro = servico * 0.80
            elif categoria == 'Training':
                pagamento = 80 * dias_trabalhados  # $80 por dia trabalhado
                lucro = servico + gorjeta - pagamento
            elif categoria == 'Coordinator':
                pagamento = servico * 0.25 + gorjeta
                lucro = servico * 0.75
            elif categoria == 'Started':
                # Aplicar cálculo por semana individualmente
                valor_comissao = servico * 0.20 + gorjeta
                valor_minimo = 150 * dias_trabalhados
                pagamento = max(valor_minimo, valor_comissao)
                lucro = servico + gorjeta - pagamento
            else:
                pagamento = 0
                lucro = servico + gorjeta

            return pd.Series([pagamento, lucro])


        # Aplicar cálculo de pagamento e lucro semanal
        weekly_totals[['Pagamento Tecnico', 'Lucro Empresa']] = weekly_totals.apply(
            calcular_pagamento_semanal, axis=1)


        # Agora calcular para cada atendimento individual (proporcional)
        def calcular_pagamento_individual(row, weekly_data):
            tech_week_data = weekly_data[
                (weekly_data['Nome'] == row['Nome']) &
                (weekly_data['Semana'] == row['Semana'])
                ]

            if len(tech_week_data) == 0:
                return pd.Series([0, row['Serviço'] + row['Gorjeta']])

            total_pagamento = tech_week_data['Pagamento Tecnico'].iloc[0]
            total_servico = tech_week_data['Serviço'].iloc[0]

            if total_servico == 0:
                return pd.Series([0, row['Serviço'] + row['Gorjeta']])

            # Pagamento proporcional ao serviço realizado neste atendimento
            pagamento = (row['Serviço'] / total_servico) * total_pagamento
            lucro = row['Serviço'] + row['Gorjeta'] - pagamento

            return pd.Series([pagamento, lucro])


        # Aplicar cálculo proporcional para cada atendimento
        completed_services[['Pagamento Tecnico', 'Lucro Empresa']] = completed_services.apply(
            lambda x: calcular_pagamento_individual(x, weekly_totals), axis=1)

        total_lucro = completed_services['Lucro Empresa'].sum()
        total_pagamentos = completed_services['Pagamento Tecnico'].sum()

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Realizados", len(completed_services))
        col2.metric("Não Realizados", len(not_completed))
        col3.metric("Total em Serviços", format_currency(completed_services['Serviço'].sum()))
        col4.metric("Total em Gorjetas", format_currency(completed_services['Gorjeta'].sum()))
        col5.metric("Lucro da Empresa", format_currency(total_lucro))

        # LAYOUT COM COLUNAS - AGORA COM CÁLCULOS À ESQUERDA E ANÁLISE À DIREITA
        col_calculos, col_analise = st.columns([1, 2])  # Proporção 1:2

        with col_calculos:
            st.header("💰 Cálculos Semanais")

            # Formatar valores monetários para exibição
            weekly_totals_display = weekly_totals.copy()
            weekly_totals_display['Serviço'] = weekly_totals_display['Serviço'].apply(format_currency)
            weekly_totals_display['Gorjeta'] = weekly_totals_display['Gorjeta'].apply(format_currency)
            weekly_totals_display['Pagamento Tecnico'] = weekly_totals_display['Pagamento Tecnico'].apply(
                format_currency)
            weekly_totals_display['Lucro Empresa'] = weekly_totals_display['Lucro Empresa'].apply(format_currency)

            weekly_totals_display = weekly_totals_display.rename(columns={
                'Nome': 'Técnico',
                'Semana': 'Semana',
                'Categoria': 'Categoria',
                'Serviço': 'Total Serviços',
                'Gorjeta': 'Total Gorjetas',
                'Pagamento Tecnico': 'Pagamento Semanal',
                'Lucro Empresa': 'Lucro da Empresa',
                'Dias Trabalhados': 'Dias Trabalhados'
            })

            st.dataframe(weekly_totals_display)

        with col_analise:
            st.header("👨‍🔧 Análise por Técnico")

            # Agrupar por técnico e categoria, mantendo os cálculos semanais separados
            tech_summary = weekly_totals.groupby(['Nome', 'Categoria']).agg({
                'Serviço': 'sum',
                'Gorjeta': 'sum',
                'Pagamento Tecnico': 'sum',
                'Lucro Empresa': 'sum',
                'Dia': 'sum',
                'Dias Trabalhados': 'sum'
            }).reset_index()

            # Ajustar nomes das colunas
            tech_summary.columns = ['Técnico', 'Categoria', 'Total Serviços',
                                    'Total Gorjetas', 'Total Pagamento', 'Lucro Empresa',
                                    'Atendimentos', 'Dias Trabalhados']

            tech_summary['Média Atendimento'] = tech_summary['Total Serviços'] / tech_summary['Atendimentos']
            tech_summary['Gorjeta Média'] = tech_summary['Total Gorjetas'] / tech_summary['Atendimentos']

            # Formatar valores monetários
            tech_summary['Total Serviços'] = tech_summary['Total Serviços'].apply(format_currency)
            tech_summary['Total Gorjetas'] = tech_summary['Total Gorjetas'].apply(format_currency)
            tech_summary['Total Pagamento'] = tech_summary['Total Pagamento'].apply(format_currency)
            tech_summary['Lucro Empresa'] = tech_summary['Lucro Empresa'].apply(format_currency)
            tech_summary['Média Atendimento'] = tech_summary['Média Atendimento'].apply(format_currency)
            tech_summary['Gorjeta Média'] = tech_summary['Gorjeta Média'].apply(format_currency)

            st.dataframe(tech_summary.sort_values('Atendimentos', ascending=False))

        st.subheader("📈 Evolução Semanal por Técnico")
        fig_evolucao = px.line(
            weekly_totals,
            x='Semana',
            y='Serviço',
            color='Nome',
            markers=True,
            title='Evolução de Serviços por Técnico',
            labels={'Serviço': 'Valor em Serviços ($)', 'Semana': 'Semana'}
        )
        fig_evolucao.update_traces(hovertemplate="<b>%{x}</b><br>Valor: $%{y:,.2f}")
        st.plotly_chart(fig_evolucao, use_container_width=True)

        # Técnico da Semana
        if len(selected_weeks) == 1:  # Mostrar apenas se uma semana estiver selecionada
            semana_atual = selected_weeks[0]
            tech_da_semana = \
                weekly_totals[weekly_totals['Semana'] == semana_atual].sort_values('Serviço',
                                                                                   ascending=False).iloc[0]

            st.subheader("🏆 Técnico da Semana")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Técnico", tech_da_semana['Nome'])
            col2.metric("Total em Serviços", format_currency(tech_da_semana['Serviço']))
            col3.metric("Pagamento Semanal", format_currency(tech_da_semana['Pagamento Tecnico']))
            col4.metric("Lucro Empresa", format_currency(tech_da_semana['Lucro Empresa']))

        fig_pagamento = px.bar(
            weekly_totals,
            x='Pagamento Tecnico',
            y='Nome',
            color='Semana',
            barmode='group',
            title='Pagamento Semanal por Técnico',
            labels={'Pagamento Tecnico': 'Pagamento ($)', 'Nome': 'Técnico'}
        )
        fig_pagamento.update_traces(texttemplate='$%{x:,.2f}', textposition='outside')
        fig_pagamento.update_layout(hovermode="x unified")
        st.plotly_chart(fig_pagamento, use_container_width=True)

        # Gráfico de atendimentos por técnico
        tech_summary_graph = tech_summary.copy()
        tech_summary_graph['Atendimentos'] = pd.to_numeric(tech_summary_graph['Atendimentos'], errors='coerce')

        fig1 = px.bar(tech_summary_graph.sort_values('Atendimentos'),
                      x='Atendimentos', y='Técnico',
                      title='Atendimentos por Técnico',
                      color='Categoria',
                      labels={'Atendimentos': 'Quantidade'})
        fig1.update_traces(hovertemplate="<b>%{y}</b><br>Atendimentos: %{x}<br>Categoria: %{marker.color}")
        st.plotly_chart(fig1, use_container_width=True)

        # Gráfico de gorjetas por técnico
        fig2 = px.bar(tech_summary_graph.sort_values('Total Gorjetas'),
                      x='Total Gorjetas', y='Técnico',
                      title='Gorjetas por Técnico',
                      color='Categoria',
                      labels={'Total Gorjetas': 'Valor Gorjetas ($)'})
        fig2.update_traces(hovertemplate="<b>%{y}</b><br>Total Gorjetas: $%{x:,.2f}<br>Categoria: %{marker.color}")
        st.plotly_chart(fig2, use_container_width=True)

        st.header("⚠️ Atendimentos Não Realizados")
        if not not_completed.empty:
            st.warning(f"{len(not_completed)} atendimentos não realizados.")
            st.dataframe(not_completed[['Nome', 'Dia', 'Data', 'Cliente']])
        else:
            st.success("Todos os agendamentos foram realizados!")

        st.header("💳 Métodos de Pagamento")
        valid_payments = completed_services[completed_services['Pagamento'].isin(FORMAS_PAGAMENTO_VALIDAS)]
        invalid_payments = completed_services[
            ~completed_services['Pagamento'].isin(FORMAS_PAGAMENTO_VALIDAS) & completed_services['Pagamento'].notna()]

        # Criar colunas para métricas
        col1, col2, col3 = st.columns(3)
        col1.metric("Válidos", len(valid_payments))
        col2.metric("Inválidos", len(invalid_payments))
        col3.metric("Formas de Pagamento", len(valid_payments['Pagamento'].unique()))

        if not valid_payments.empty:
            # Criar dataframe com informações detalhadas
            payment_methods = valid_payments.groupby('Pagamento').agg({
                'Serviço': ['sum', 'count'],
                'Gorjeta': 'sum',
                'Cliente': 'count',
                'Lucro Empresa': 'sum'
            }).reset_index()

            # Renomear colunas para melhor visualização
            payment_methods.columns = ['Pagamento', 'Total Serviços', 'Qtd Usos', 'Total Gorjetas',
                                       'Total Atendimentos', 'Lucro Empresa']

            # Calcular valores totais
            payment_methods['Total Geral'] = payment_methods['Total Serviços'] + payment_methods['Total Gorjetas']

            # Calcular porcentagem de uso
            total_usos = payment_methods['Qtd Usos'].sum()
            payment_methods['% Uso'] = (payment_methods['Qtd Usos'] / total_usos * 100).round(2)

            # Formatar valores monetários
            payment_methods['Total Serviços'] = payment_methods['Total Serviços'].apply(format_currency)
            payment_methods['Total Gorjetas'] = payment_methods['Total Gorjetas'].apply(format_currency)
            payment_methods['Lucro Empresa'] = payment_methods['Lucro Empresa'].apply(format_currency)
            payment_methods['Total Geral'] = payment_methods['Total Geral'].apply(format_currency)
            payment_methods['% Uso'] = payment_methods['% Uso'].astype(str) + '%'

            # Mostrar tabela detalhada
            st.subheader("Detalhes por Método de Pagamento")
            st.dataframe(payment_methods.sort_values('Qtd Usos', ascending=False))

            # Criar gráficos
            tab1, tab2 = st.tabs(["Valor Total", "Quantidade de Usos"])

            with tab1:
                # Dataframe para gráfico (valores numéricos)
                payment_graph = valid_payments.groupby('Pagamento').agg({
                    'Serviço': 'sum',
                    'Gorjeta': 'sum',
                    'Lucro Empresa': 'sum'
                }).reset_index()
                payment_graph['Total'] = payment_graph['Serviço'] + payment_graph['Gorjeta']

                fig_total = px.bar(payment_graph.sort_values('Total'),
                                   x='Total', y='Pagamento',
                                   title='Valor Total por Método de Pagamento (Serviços + Gorjetas)',
                                   color='Serviço',
                                   color_continuous_scale='Peach',
                                   labels={'Total': 'Valor Total ($)', 'Serviço': 'Valor Serviços ($)'})
                fig_total.update_traces(
                    hovertemplate="<b>%{y}</b><br>Total: $%{x:,.2f}<br>Serviços: $%{marker.color:,.2f}")
                st.plotly_chart(fig_total, use_container_width=True)

            with tab2:
                payment_count = valid_payments['Pagamento'].value_counts().reset_index()
                payment_count.columns = ['Pagamento', 'Qtd Usos']

                # Calcular porcentagem para o gráfico
                total = payment_count['Qtd Usos'].sum()
                payment_count['% Uso'] = (payment_count['Qtd Usos'] / total * 100).round(2)

                fig_qtd = px.bar(payment_count.sort_values('Qtd Usos'),
                                 x='Qtd Usos', y='Pagamento',
                                 title='Quantidade de Usos por Método de Pagamento',
                                 color='Qtd Usos',
                                 color_continuous_scale='Peach',
                                 labels={'Qtd Usos': 'Quantidade de Usos'},
                                 text='% Uso')

                fig_qtd.update_traces(
                    texttemplate='%{text}%',
                    textposition='outside',
                    hovertemplate="<b>%{y}</b><br>Usos: %{x}<br>% do Total: %{text}%"
                )
                st.plotly_chart(fig_qtd, use_container_width=True)

        if not invalid_payments.empty:
            st.warning("Pagamentos inválidos encontrados:")
            st.dataframe(invalid_payments[['Nome', 'Data', 'Cliente', 'Pagamento']])

        st.header("📅 Análise por Dia da Semana")
        day_summary = completed_services.groupby('Dia').agg({
            'Serviço': ['count', 'sum'],
            'Gorjeta': 'sum',
            'Pets': 'sum',
            'Lucro Empresa': 'sum'
        }).reset_index()
        day_summary.columns = ['Dia', 'Atendimentos', 'Total Serviços', 'Total Gorjetas', 'Total Pets', 'Lucro Empresa']
        day_summary['Dia'] = pd.Categorical(day_summary['Dia'],
                                            categories=['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta',
                                                        'Sábado'],
                                            ordered=True)
        day_summary = day_summary.sort_values('Dia')

        # Formatar valores monetários para exibição
        day_summary_display = day_summary.copy()
        day_summary_display['Total Serviços'] = day_summary_display['Total Serviços'].apply(format_currency)
        day_summary_display['Total Gorjetas'] = day_summary_display['Total Gorjetas'].apply(format_currency)
        day_summary_display['Lucro Empresa'] = day_summary_display['Lucro Empresa'].apply(format_currency)

        st.dataframe(day_summary_display)

        fig7 = px.bar(day_summary, x='Dia', y='Atendimentos',
                      title='Atendimentos por Dia da Semana',
                      labels={'Atendimentos': 'Quantidade'})
        fig7.update_traces(hovertemplate="<b>%{x}</b><br>Atendimentos: %{y}")
        st.plotly_chart(fig7, use_container_width=True)

        st.header("📤 Exportar Dados")
        if st.button("Exportar CSV"):
            csv = data.to_csv(index=False).encode('utf-8')
            st.download_button("📁 Baixar CSV", data=csv, file_name="servicos_tecnicos.csv", mime="text/csv")

st.markdown("""
    <style>
    .stMetricValue { font-size: 22px; }
    .stDataFrame th, .stDataFrame td { padding: 8px 10px; }
    .css-1aumxhk { background-color: #f9f9f9; padding: 20px; border-radius: 10px; }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
---
<small>Desenvolvido por Alan Salviano | Análise de Planilhas de Serviços Técnicos</small>
""", unsafe_allow_html=True)