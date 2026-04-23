import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

df = pd.read_excel('opps_corrigido.xlsx')

for col in ['Amount', 'Total_Product_Amount']:
    df[col] = df[col].astype(str).str.replace(',', '.').str.strip()
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

df['Close_Date'] = pd.to_datetime(df['Close_Date'], dayfirst=True)
df['Created_Date'] = pd.to_datetime(df['Created_Date'], dayfirst=True)

df['Month_Year'] = df['Close_Date'].dt.strftime('%Y-%m')

df_unique = df.drop_duplicates(subset='Opportunity_ID').copy()

layout_dark = {
    'template': 'plotly_dark',
    'paper_bgcolor': '#1e293b',
    'plot_bgcolor': '#1e293b',
    'font': {'color': '#f8fafc', 'family': 'Inter'},
    'margin': {'t': 50, 'b': 50, 'l': 50, 'r': 50}
}

receita_mom = df_unique[df_unique['Stage'] == 'Closed Won'].groupby('Month_Year')['Amount'].sum().reset_index()

fig1 = px.bar(receita_mom, x='Month_Year', y='Amount', title='Receita Closed Won Mensal (2026)',
             color_discrete_sequence=['#f97316'])
fig1.update_xaxes(type='category')
fig1.update_layout(layout_dark)
fig1.update_traces(texttemplate='%{y:.2s}', textposition='outside')

part_source = df_unique[df_unique['Stage'] == 'Closed Won'].groupby('Lead_Source_Category')['Amount'].sum().reset_index()

fig2 = px.pie(part_source, values='Amount', names='Lead_Source_Category', hole=0.5,
             title='Participação de Receita por Fonte de Lead',
             color_discrete_sequence=px.colors.sequential.Oranges_r)
fig2.update_layout(layout_dark)

counts = df_unique.groupby(['Lead_Source_Category', 'Stage']).size().unstack(fill_value=0)

if 'Closed Won' in counts.columns:
    counts['Total'] = counts.sum(axis=1)
    counts['Win_Rate'] = (counts['Closed Won'] / counts['Total'] * 100).round(1)
    fig4 = px.bar(counts.reset_index(), x='Lead_Source_Category', y='Win_Rate', 
                 title='Win Rate % por Categoria de Origem',
                 color_discrete_sequence=['#fb923c'])
    fig4.update_layout(layout_dark)

df_unique['Ciclo_Venda'] = (df_unique['Close_Date'] - df_unique['Created_Date']).dt.days

ciclo_office = df_unique[df_unique['Stage'] == 'Closed Won'].groupby('Lead_Office')['Ciclo_Venda'].mean().reset_index()

fig8 = px.bar(ciclo_office, x='Lead_Office', y='Ciclo_Venda', title='Ciclo de Venda Médio (Dias) por Escritório',
             color_discrete_sequence=['#f97316'])
fig8.update_layout(layout_dark)

graphs_html = ""
for f in [fig1, fig2, fig4, fig8]:
    graphs_html += f.to_html(full_html=False, include_plotlyjs='cdn')

html_final = f"""
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ background-color: #0f172a; font-family: 'Inter', sans-serif; color: white; padding: 40px; }}
        .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        .header {{ border-bottom: 2px solid #f97316; margin-bottom: 30px; padding-bottom: 10px; }}
        .chart-card {{ background: #1e293b; border-radius: 12px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 Análise de Performance Comercial - Monks 2026</h1>
        <p>Dados limpos e normalizados via script de auditoria.</p>
    </div>
    <div class="grid">
        {graphs_html}
    </div>
</body>
</html>
"""

with open('analise.html', 'w', encoding='utf-8') as f:
    f.write(html_final)

print("Análise concluída! Abra o arquivo analise.html.")