import pandas as pd
from wcwidth import width
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
    'paper_bgcolor': 'rgba(0,0,0,0)',
    'plot_bgcolor': 'rgba(0,0,0,0)',
    'font': {'color': '#f8fafc', 'family': 'Arial'},
    'autosize': True, 
    'margin': {'t': 40, 'b': 40, 'l': 60, 'r': 40}
}

layout_compacto = layout_dark.copy()
layout_compacto['margin'] = {'t': 40, 'b': 30, 'l': 30, 'r': 10}
layout_compacto['font'] = {'size': 10, 'color': '#f8fafc'}

# =========================
# 1. Receita MoM
# =========================
receita_mom = df[df['Stage'] == 'Closed Won'] \
    .groupby('Month_Year')['Amount'].sum().reset_index()

fig1 = px.bar(receita_mom, x='Month_Year', y='Amount',
              title='Receita Closed Won Mensal (2026)',
              color_discrete_sequence=['#f97316'])

fig1.update_layout(layout_dark)

fig1.update_xaxes(type='category', categoryorder='category ascending')
fig1.update_layout(bargap=0.4)

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
fig2.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.1))

insight2 = "Algumas fontes concentram a maior parte da receita, indicando maior eficiência desses canais."

# =========================
# 3. Top 10 oportunidades abertas
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
fig4.update_layout(layout_compacto)

insight4 = "Diferenças relevantes de conversão entre canais indicam oportunidades de otimização."

# =========================
# 5. Ticket médio
# =========================
ticket = df.groupby('Type')['Amount'].mean().reset_index()

fig5 = px.bar(ticket, x='Type', y='Amount',
              title='Ticket Médio por Tipo de Projeto',
              color_discrete_sequence=['#f97316'])
fig5.update_layout(layout_compacto)

insight5 = "Projetos possuem tickets distintos, refletindo diferentes perfis de negócio."

# =========================
# 6. Pipeline por estágio
# =========================
pipeline = df[df['Stage'] != 'Closed Won'] \
    .groupby('Stage')['Amount'].sum().reset_index()

fig6 = px.bar(pipeline, x='Stage', y='Amount',
              title='Pipeline em Aberto por Estágio',
              color_discrete_sequence=['#f97316'])
fig6.update_layout(layout_compacto)

insight6 = "O volume financeiro concentra-se em etapas intermediárias do funil."

total_revenue = receita_mom['Amount'].sum()
avg_win_rate = counts['Win_Rate'].mean()
total_pipeline = pipeline['Amount'].sum()

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
fig7.update_layout(layout_compacto)

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
fig8.update_layout(layout_compacto)

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
fig9.update_layout(layout_compacto)

insight9 = "A receita está concentrada em poucos clientes estratégicos."

# =========================
# 10. Idade pipeline
# =========================
df_open = df[df['Stage'] != 'Closed Won'].copy()
df_open['Pipeline_Age'] = (pd.Timestamp.today() - df_open['Created_Date']).dt.days

fig10 = px.histogram(df_open, x='Pipeline_Age',
                     title='Idade do Pipeline',
                     color_discrete_sequence=['#f97316'])
fig10.update_layout(layout_compacto)

insight10 = "Existem oportunidades envelhecidas que podem estar estagnadas."

# =========================
# COMPONENTE CARD
# =========================
def card(grafico, insight):
    config_plotly = {'responsive': True}
    
    return f"""
    <div class="chart-card">
        <div style="width: 100%; overflow: hidden;">
            {grafico.to_html(full_html=False, include_plotlyjs='cdn', config=config_plotly)}
        </div>
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
            font-family: 'Arial', sans-serif; 
            color: white; 
            padding: 20px; 
            margin: 0; 
        }}

        .header {{ 
            margin-bottom: 20px; 
            padding-left: 10px;
        }}

        .header h1 {{ 
            color: #f97316; 
            margin: 0; 
            font-size: 24px; 
        }}

        .kpi-container {{ 
            display: flex; 
            flex-wrap: wrap; 
            gap: 15px; 
            margin-bottom: 20px; 
        }}

        .kpi-card {{ 
            background: #1e293b; 
            padding: 10px 20px; 
            border-radius: 8px; 
            min-width: 200px; 
            flex: 1; 
            border-left: 4px solid #f97316;
            display: flex; 
            flex-direction: column; 
            justify-content: center; 
            height: 60px;
        }}

        .kpi-card h2 {{
            margin: 0; 
            font-size: 11px; 
            color: #94a3b8; 
            text-transform: uppercase; 
        }}
        
        .kpi-card p {{ 
            margin: 2px 0 0; 
            font-size: 18px; 
            font-weight: bold; 
            color: #f97316; 
        }}

        .grid {{
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); 
            gap: 20px; 
            align-items: stretch; 
        }}

        .grid-top {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 20px;
        }}

        .grid-bottom {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            margin-top: 20px;
        }}

        .full-bottom {{
            grid-column: span 3;
            margin-top: 10px;
        }}

        .chart-card {{
            background: #1e293b;
            border-radius: 12px;
            padding: 15px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}

        .insight {{
            background: rgba(15, 23, 42, 0.4); 
            padding: 12px; 
            border-radius: 6px;
            font-size: 12px; 
            color: #cbd5e1; 
            margin-top: auto; 
            border-left: 2px solid #f97316;
        }}

        .custom-table {{ 
            width: 100%; 
            border-collapse: collapse; 
            font-size: 12px; 
        }}

        .custom-table th {{ 
            background-color: #f97316; 
            color: white; 
            padding: 8px; 
            text-align: left; 
        }}

        .custom-table td {{ 
            padding: 8px; 
            border-bottom: 1px solid #334155; 
        }}

        @media (max-width: 1100px) {{
            .grid-bottom {{
                grid-template-columns: 1fr 1fr; /* 2 colunas em tablets */
            }}
        }}

        @media (max-width: 768px) {{
            .grid-bottom {{
                grid-template-columns: 1fr; /* 1 coluna no celular */
            }}
        }}
    </style>
    <title>Análise Comercial | Operations</title>
</head>
<body>
    <div class="header">
        <h1>📊 Performance Comercial - Monks 2026</h1>
        <p>Visão executiva de pipeline e conversão.</p>
    </div>

    <div class="kpi-container">
        <div class="kpi-card"><h2>Receita Total (Won)</h2><p>${total_revenue:,.0f}</p></div>
        <div class="kpi-card"><h2>Win Rate Médio</h2><p>{avg_win_rate:.1f}%</p></div>
        <div class="kpi-card"><h2>Volume em Pipeline</h2><p>${total_pipeline:,.0f}</p></div>
    </div>

    <div class="grid-top">
        {card(fig1, insight1)}
        {card(fig2, insight2)}
    </div>

    <div class="chart-card" style="margin-bottom: 20px;">
        <h4 style="margin:0;">Top 10 Oportunidades em Aberto</h4>
        {table_top_opps}
        <p class="insight">{insight3}</p>
    </div>

    <div class="grid-bottom">
        {card(fig4, insight4)}
        {card(fig5, insight5)}
        {card(fig6, insight6)}
        {card(fig7, insight7)}
        {card(fig8, insight8)}
        {card(fig9, insight9)}
        
        <div class="chart-card full-bottom">
            <div style="width: 100%; overflow: hidden;">
                {fig10.to_html(full_html=False, include_plotlyjs='cdn', config={'responsive': True})}
            </div>
            <p class="insight">{insight10}</p>
        </div>
    </div>
</body>
</html>
"""

with open('reports/analise.html', 'w', encoding='utf-8') as f:
    f.write(html_final)

print("✅ Análise completa gerada!")