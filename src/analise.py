import pandas as pd
import plotly.express as px

# =========================
# LEITURA
# =========================
df = pd.read_excel('data/processed/opps_corrigido.xlsx')

df['Close_Date'] = pd.to_datetime(df['Close_Date'], dayfirst=True)
df['Created_Date'] = pd.to_datetime(df['Created_Date'], dayfirst=True)

df['Month_Year'] = df['Close_Date'].dt.strftime('%Y-%m')

# =========================
# LAYOUT PADRÃO
# =========================
layout_dark = {
    'template': 'plotly_dark',
    'paper_bgcolor': '#1e293b',
    'plot_bgcolor': '#1e293b',
    'font': {'color': '#f8fafc', 'family': 'Inter'}
}

# =========================
# 1. Receita MoM
# =========================
receita_mom = df[df['Stage'] == 'Closed Won'] \
    .groupby('Month_Year')['Amount'].sum().reset_index()

fig1 = px.bar(receita_mom, x='Month_Year', y='Amount',
              title='Receita Closed Won Mensal (2026)',
              color_discrete_sequence=['#f97316'])
fig1.update_layout(layout_dark)

insight1 = "A receita apresenta concentração em determinados meses, indicando possível sazonalidade ou fechamento concentrado de deals."

# =========================
# 2. Participação por fonte
# =========================
part_source = df[df['Stage'] == 'Closed Won'] \
    .groupby('Lead_Source_Category')['Amount'].sum().reset_index()

fig2 = px.pie(part_source, values='Amount', names='Lead_Source_Category',
              hole=0.5,
              title='Participação de Receita por Fonte de Lead',
              color_discrete_sequence=px.colors.sequential.Oranges_r)
fig2.update_layout(layout_dark)

insight2 = "Algumas fontes concentram a maior parte da receita, indicando maior eficiência desses canais."

# =========================
# 3. Top 10 oportunidades abertas (TABELA BONITA)
# =========================
top_opps = df[df['Stage'] != 'Closed Won'] \
    .nlargest(10, 'Amount')[['Account_Name', 'Stage', 'Amount', 'Close_Date']]

# formatação
top_opps['Amount'] = top_opps['Amount'].apply(lambda x: f"${x:,.0f}")
top_opps.columns = ['Cliente', 'Estágio', 'Valor', 'Data de Fechamento']

table_top_opps = top_opps.to_html(index=False, classes='custom-table', border=0)

insight3 = "O pipeline possui alta concentração de valor em poucas oportunidades, indicando risco de dependência em grandes deals."

# =========================
# 4. Win rate
# =========================
counts = df.groupby(['Lead_Source_Category', 'Stage']).size().unstack(fill_value=0)
counts['Total'] = counts.sum(axis=1)
counts['Win_Rate'] = (counts.get('Closed Won', 0) / counts['Total'] * 100).round(1)

fig4 = px.bar(counts.reset_index(), x='Lead_Source_Category', y='Win_Rate',
              title='Win Rate (%) por Fonte',
              color_discrete_sequence=['#fb923c'])
fig4.update_layout(layout_dark)

insight4 = "Diferenças relevantes de conversão entre canais indicam oportunidades de otimização."

# =========================
# 5. Ticket médio
# =========================
ticket = df.groupby('Type')['Amount'].mean().reset_index()

fig5 = px.bar(ticket, x='Type', y='Amount',
              title='Ticket Médio por Tipo de Projeto',
              color_discrete_sequence=['#f97316'])
fig5.update_layout(layout_dark)

insight5 = "Projetos possuem tickets distintos, refletindo diferentes perfis de negócio."

# =========================
# 6. Pipeline por estágio
# =========================
pipeline = df[df['Stage'] != 'Closed Won'] \
    .groupby('Stage')['Amount'].sum().reset_index()

fig6 = px.bar(pipeline, x='Stage', y='Amount',
              title='Pipeline em Aberto por Estágio',
              color_discrete_sequence=['#f97316'])
fig6.update_layout(layout_dark)

insight6 = "O volume financeiro concentra-se em etapas intermediárias do funil."

# =========================
# 7. Mix negócio
# =========================
df['Tipo_Negocio'] = df['Type'].apply(
    lambda x: 'Upsell' if 'Upsell' in x else 'New Business'
)

mix = df[df['Stage'] == 'Closed Won'] \
    .groupby(['Month_Year', 'Tipo_Negocio'])['Amount'].sum().reset_index()

fig7 = px.bar(mix, x='Month_Year', y='Amount',
              color='Tipo_Negocio',
              title='Mix New Business vs Upsell',
              barmode='stack')
fig7.update_layout(layout_dark)

insight7 = "A proporção entre novos negócios e upsell varia ao longo do tempo."

# =========================
# 8. Ciclo de venda
# =========================
df['Ciclo_Venda'] = (df['Close_Date'] - df['Created_Date']).dt.days

ciclo = df[df['Stage'] == 'Closed Won'] \
    .groupby('Lead_Office')['Ciclo_Venda'].mean().reset_index()

fig8 = px.bar(ciclo, x='Lead_Office', y='Ciclo_Venda',
              title='Ciclo de Venda Médio por Escritório',
              color_discrete_sequence=['#f97316'])
fig8.update_layout(layout_dark)

insight8 = "Diferenças no ciclo de venda indicam variações de eficiência operacional."

# =========================
# 9. Top clientes
# =========================
top_clients = df[df['Stage'] == 'Closed Won'] \
    .groupby('Account_Name')['Amount'].sum() \
    .nlargest(10).reset_index()

fig9 = px.bar(top_clients, x='Amount', y='Account_Name',
              orientation='h',
              title='Top 10 Clientes',
              color_discrete_sequence=['#f97316'])
fig9.update_layout(layout_dark)

insight9 = "A receita está concentrada em poucos clientes estratégicos."

# =========================
# 10. Idade pipeline
# =========================
df_open = df[df['Stage'] != 'Closed Won'].copy()
df_open['Pipeline_Age'] = (pd.Timestamp.today() - df_open['Created_Date']).dt.days

fig10 = px.histogram(df_open, x='Pipeline_Age',
                     title='Idade do Pipeline',
                     color_discrete_sequence=['#f97316'])
fig10.update_layout(layout_dark)

insight10 = "Existem oportunidades envelhecidas que podem estar estagnadas."

# =========================
# COMPONENTE CARD
# =========================
def card(grafico, insight):
    return f"""
    <div class="chart-card">
        {grafico.to_html(full_html=False, include_plotlyjs='cdn')}
        <p class="insight">{insight}</p>
    </div>
    """

# =========================
# HTML FINAL
# =========================
html_final = f"""
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{
            background-color: #0f172a;
            font-family: 'Inter', sans-serif;
            color: white;
            padding: 40px;
        }}

        .header {{
            border-bottom: 2px solid #f97316;
            margin-bottom: 30px;
        }}

        .grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }}

        .chart-card {{
            background: #1e293b;
            border-radius: 12px;
            padding: 15px;
        }}

        .insight {{
            font-size: 14px;
            color: #94a3b8;
            margin-top: 20px;
        }}

        .custom-table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            margin-top: 20px;
        }}

        .custom-table th {{
            background-color: #f97316;
            color: white;
            padding: 10px;
            text-align: left;
        }}

        .custom-table td {{
            padding: 10px;
            border-bottom: 1px solid #334155;
            background-color: #1e293b;
        }}

        .grid {{
            grid-template-columns: 1fr;
        }}

    </style>
    <title>Análise Comercial | Operations</title>
</head>

<body>
    <div class="header">
        <h1>📊 Análise de Performance Comercial - 2026</h1>
        <p>Dados tratados e normalizados para análise de pipeline e receita.</p>
    </div>

    <div class="grid">
        {card(fig1, insight1)}
        {card(fig2, insight2)}

        <div class="chart-card">
            <h4>Top 10 Oportunidades em Aberto</h4>
            {table_top_opps}
            <p class="insight">{insight3}</p>
        </div>

        {card(fig4, insight4)}
        {card(fig5, insight5)}
        {card(fig6, insight6)}
        {card(fig7, insight7)}
        {card(fig8, insight8)}
        {card(fig9, insight9)}
        {card(fig10, insight10)}

    </div>
</body>
</html>
"""
with open('reports/analise.html', 'w', encoding='utf-8') as f:
    f.write(html_final)

print("✅ Análise completa gerada!")