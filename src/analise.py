import os
import copy
import pandas as pd
import plotly.express as px

# =========================
# LEITURA E TRATAMENTO
# =========================
df = pd.read_excel('data/processed/opps_corrigido.xlsx')

df['Close_Date'] = pd.to_datetime(df['Close_Date'], dayfirst=True)
df['Created_Date'] = pd.to_datetime(df['Created_Date'], dayfirst=True)

df['Month_Year'] = df['Close_Date'].dt.strftime('%Y-%m')

# =========================
# LAYOUT PLOTLY (.MONKS DESIGN SYSTEM)
# =========================
FONT_FAMILY = "Inter, system-ui, -apple-system, sans-serif"

layout_base = {
    'paper_bgcolor': 'rgba(0,0,0,0)',
    'plot_bgcolor': 'rgba(0,0,0,0)',
    'font': {'color': '#1A1A1A', 'family': FONT_FAMILY, 'size': 12},
    'title': {
        'font': {'size': 15, 'color': '#1A1A1A', 'family': FONT_FAMILY},
        'x': 0,
        'xanchor': 'left',
        'y': 0.96
    },
    'autosize': True,
    'height': 360,
    'margin': {'t': 56, 'b': 48, 'l': 56, 'r': 24},
    'xaxis': {
        'showgrid': False,
        'linecolor': '#E2E1DC',
        'automargin': True,
        'tickfont': {'size': 11, 'color': '#5A5A55'}
    },
    'yaxis': {
        'gridcolor': '#E8E7E2',
        'zerolinecolor': '#E2E1DC',
        'automargin': True,
        'tickfont': {'size': 11, 'color': '#5A5A55'}
    },
    'hoverlabel': {
        'bgcolor': '#1A1A1A',
        'font_size': 12,
        'font_family': FONT_FAMILY,
        'font_color': '#F3F3F1'
    }
}

def aplicar_layout(fig, height=360, margin_override=None):
    cfg = copy.deepcopy(layout_base)
    cfg['height'] = height
    if margin_override:
        cfg['margin'].update(margin_override)
    fig.update_layout(cfg)
    fig.update_xaxes(automargin=True)
    fig.update_yaxes(automargin=True)
    return fig

# Paleta editorial monocromática (.Monks) - Modo Claro
paleta_cinzas = ['#1A1A1A', '#4A4A48', '#787874', '#A8A8A2', '#D4D3CD']

# =========================
# 1. Receita MoM
# =========================
receita_mom = df[df['Stage'] == 'Closed Won'] \
    .groupby('Month_Year')['Amount'].sum().reset_index()

fig1 = px.bar(
    receita_mom, x='Month_Year', y='Amount',
    title='Receita Closed Won Mensal (2026)',
    color_discrete_sequence=['#1A1A1A']
)
aplicar_layout(fig1, height=370)
fig1.update_xaxes(type='category', categoryorder='category ascending', title=None)
fig1.update_yaxes(title=None, tickprefix="$")
fig1.update_layout(bargap=0.38)
fig1.update_traces(marker_line_width=0)

insight1 = "A receita apresenta concentração em determinados meses, indicando possível sazonalidade ou fechamento concentrado de deals."

# =========================
# 2. Participação por fonte
# =========================
part_source = df[df['Stage'] == 'Closed Won'] \
    .groupby('Lead_Source_Category')['Amount'].sum().reset_index()

fig2 = px.pie(
    part_source, values='Amount', names='Lead_Source_Category',
    hole=0.58,
    title='Participação de Receita por Fonte de Lead',
    color_discrete_sequence=paleta_cinzas
)
aplicar_layout(fig2, height=370, margin_override={'t': 56, 'b': 72, 'l': 24, 'r': 24})
fig2.update_layout(
    showlegend=True,
    legend=dict(orientation="h", y=-0.22, x=0.5, xanchor="center", font=dict(size=11))
)
fig2.update_traces(
    textposition='inside',
    textinfo='percent',
    marker=dict(line=dict(color='#FFFFFF', width=1.5))
)

insight2 = "Algumas fontes concentram a maior parte da receita, indicando maior eficiência desses canais."

# =========================
# 3. Top 10 oportunidades abertas
# =========================
top_opps = df[df['Stage'] != 'Closed Won'] \
    .nlargest(10, 'Amount')[['Account_Name', 'Stage', 'Amount', 'Close_Date']].copy()

top_opps['Amount'] = top_opps['Amount'].apply(lambda x: f"${x:,.0f}")
top_opps['Close_Date'] = top_opps['Close_Date'].dt.strftime('%d/%m/%Y')
top_opps.columns = ['Cliente', 'Estágio do Funil', 'Valor Estimado', 'Previsão de Fechamento']

table_top_opps = top_opps.to_html(index=False, classes='custom-table', border=0)

insight3 = "O pipeline possui alta concentração de valor em poucas oportunidades, indicando risco de dependência em grandes deals."

# =========================
# 4. Win rate
# =========================
counts = df.groupby(['Lead_Source_Category', 'Stage']).size().unstack(fill_value=0)
counts['Total'] = counts.sum(axis=1)
counts['Win_Rate'] = (counts.get('Closed Won', 0) / counts['Total'] * 100).round(1)

fig4 = px.bar(
    counts.reset_index(), x='Lead_Source_Category', y='Win_Rate',
    title='Win Rate (%) por Fonte',
    color_discrete_sequence=['#1A1A1A']
)
aplicar_layout(fig4, height=350)
fig4.update_xaxes(title=None)
fig4.update_yaxes(title=None, ticksuffix="%")
fig4.update_layout(bargap=0.35)

insight4 = "Diferenças relevantes de conversão entre canais indicam oportunidades de otimização."

# =========================
# 5. Ticket médio
# =========================
ticket = df.groupby('Type')['Amount'].mean().reset_index()

fig5 = px.bar(
    ticket, x='Type', y='Amount',
    title='Ticket Médio por Tipo de Projeto',
    color_discrete_sequence=['#1A1A1A']
)
aplicar_layout(fig5, height=350)
fig5.update_xaxes(title=None)
fig5.update_yaxes(title=None, tickprefix="$")
fig5.update_layout(bargap=0.4)

insight5 = "Projetos possuem tickets distintos, refletindo diferentes perfis de negócio."

# =========================
# 6. Pipeline por estágio
# =========================
pipeline = df[df['Stage'] != 'Closed Won'] \
    .groupby('Stage')['Amount'].sum().reset_index()

fig6 = px.bar(
    pipeline, x='Stage', y='Amount',
    title='Pipeline em Aberto por Estágio',
    color_discrete_sequence=['#1A1A1A']
)
aplicar_layout(fig6, height=350)
fig6.update_xaxes(title=None, tickangle=-20)
fig6.update_yaxes(title=None, tickprefix="$")
fig6.update_layout(bargap=0.35)

insight6 = "O volume financeiro concentra-se em etapas intermediárias do funil."

total_revenue = receita_mom['Amount'].sum()
avg_win_rate = counts['Win_Rate'].mean()
total_pipeline = pipeline['Amount'].sum()
total_open_deals = len(df[df['Stage'] != 'Closed Won'])

# =========================
# 7. Mix negócio
# =========================
df['Tipo_Negocio'] = df['Type'].apply(
    lambda x: 'Upsell' if 'Upsell' in str(x) else 'New Business'
)

mix = df[df['Stage'] == 'Closed Won'] \
    .groupby(['Month_Year', 'Tipo_Negocio'])['Amount'].sum().reset_index()

fig7 = px.bar(
    mix, x='Month_Year', y='Amount',
    color='Tipo_Negocio',
    title='Mix New Business vs Upsell',
    barmode='stack',
    color_discrete_sequence=['#1A1A1A', '#888884']
)
aplicar_layout(fig7, height=350, margin_override={'t': 72, 'b': 48, 'l': 56, 'r': 24})
fig7.update_layout(
    legend=dict(orientation="h", y=1.14, x=0, xanchor="left", title=None),
    bargap=0.35
)
fig7.update_xaxes(type='category', categoryorder='category ascending', title=None)
fig7.update_yaxes(title=None, tickprefix="$")

insight7 = "A proporção entre novos negócios e upsell varia ao longo do tempo."

# =========================
# 8. Ciclo de venda
# =========================
df['Ciclo_Venda'] = (df['Close_Date'] - df['Created_Date']).dt.days

ciclo = df[df['Stage'] == 'Closed Won'] \
    .groupby('Lead_Office')['Ciclo_Venda'].mean().reset_index()

fig8 = px.bar(
    ciclo, x='Lead_Office', y='Ciclo_Venda',
    title='Ciclo de Venda Médio por Escritório (Dias)',
    color_discrete_sequence=['#1A1A1A']
)
aplicar_layout(fig8, height=360)
fig8.update_xaxes(title=None)
fig8.update_yaxes(title=None, ticksuffix="d")
fig8.update_layout(bargap=0.42)

insight8 = "Diferenças no ciclo de venda indicam variações de eficiência operacional."

# =========================
# 9. Top clientes
# =========================
top_clients = df[df['Stage'] == 'Closed Won'] \
    .groupby('Account_Name')['Amount'].sum() \
    .nlargest(10).sort_values(ascending=True).reset_index()

fig9 = px.bar(
    top_clients, x='Amount', y='Account_Name',
    orientation='h',
    title='Top 10 Clientes (Closed Won)',
    color_discrete_sequence=['#1A1A1A']
)
aplicar_layout(fig9, height=360, margin_override={'t': 56, 'b': 44, 'l': 140, 'r': 28})
fig9.update_xaxes(title=None, tickprefix="$", showgrid=True, gridcolor='#E8E7E2')
fig9.update_yaxes(title=None, showgrid=False)

insight9 = "A receita está concentrada em poucos clientes estratégicos."

# =========================
# 10. Idade pipeline
# =========================
df_open = df[df['Stage'] != 'Closed Won'].copy()
df_open['Pipeline_Age'] = (pd.Timestamp.today() - df_open['Created_Date']).dt.days

fig10 = px.histogram(
    df_open, x='Pipeline_Age',
    title='Distribuição de Idade do Pipeline (Dias em Aberto)',
    color_discrete_sequence=['#1A1A1A']
)
aplicar_layout(fig10, height=340)
fig10.update_xaxes(title="Dias desde a criação")
fig10.update_yaxes(title="Qtd. Oportunidades")
fig10.update_layout(bargap=0.12)

insight10 = "Existem oportunidades envelhecidas que podem estar estagnadas no funil comercial."

# =========================
# COMPONENTE CARD (.MONKS BENTO CARD)
# =========================
def card(grafico, insight, badge_code, categoria, extra_class=""):
    config_plotly = {'responsive': True, 'displayModeBar': False}
    html_plot = grafico.to_html(
        full_html=False,
        include_plotlyjs=False,
        config=config_plotly,
        default_width='100%'
    )

    return f"""
    <article class="chart-card {extra_class}" data-category="{categoria}">
        <div class="card-meta">
            <span class="module-tag">{badge_code}</span>
            <span class="status-dot"></span>
        </div>
        <div class="plot-wrapper">
            {html_plot}
        </div>
        <div class="insight-box">
            <span class="insight-label">INSIGHT EXECUTIVO</span>
            <p class="insight-text">{insight}</p>
        </div>
    </article>
    """

# =========================
# HTML FINAL (DESIGN SYSTEM IGUAL AO INDEX.HTML)
# =========================
html_final = f"""<!DOCTYPE html>
<html lang="pt-br" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Performance Comercial | .Monks Operations Case</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@500;600;700&display=swap" rel="stylesheet">
    <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
    <style>
        :root {{
            --bg-canvas: #F3F3F1;
            --bg-secondary: #C9C6C1;
            --surface-card: #FFFFFF;
            --surface-subtle: #F8F8F6;
            --text-main: #1A1A1A;
            --text-secondary: #4A4A46;
            --text-muted: #787873;
            --border-light: #E2E1DC;
            --border-strong: #1A1A1A;
            --shadow-card: 0 12px 32px -12px rgba(26, 26, 26, 0.08);
            --radius-lg: 16px;
            --radius-md: 10px;
            --radius-pill: 50px;
            --transition: 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        }}

        [data-theme="dark"] {{
            --bg-canvas: #111110;
            --bg-secondary: #1B1B19;
            --surface-card: #1A1A18;
            --surface-subtle: #242421;
            --text-main: #F3F3F1;
            --text-secondary: #B5B5B0;
            --text-muted: #8E8E88;
            --border-light: #2C2C28;
            --border-strong: #F3F3F1;
            --shadow-card: 0 12px 32px -12px rgba(0, 0, 0, 0.45);
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}

        body {{
            background-color: var(--bg-canvas);
            background-image: radial-gradient(var(--border-light) 1px, transparent 1px);
            background-size: 24px 24px;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Helvetica', Arial, sans-serif;
            color: var(--text-main);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            transition: background-color var(--transition), color var(--transition);
        }}

        /* BARRA SUPERIOR PADRONIZADA COM A HOME */
        .topbar {{
            position: sticky;
            top: 0;
            z-index: 100;
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            background: color-mix(in srgb, var(--bg-canvas) 88%, transparent);
            border-bottom: 1px solid var(--border-light);
            padding: 16px 32px;
        }}

        .topbar-inner {{
            max-width: 1400px;
            margin: 0 auto;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            flex-wrap: wrap;
        }}

        .brand {{
            display: flex;
            align-items: center;
            gap: 12px;
            text-decoration: none;
            color: var(--text-main);
        }}

        .brand-mark {{
            font-weight: 900;
            font-size: 20px;
            letter-spacing: -0.04em;
        }}

        .brand-divider {{
            width: 1px;
            height: 18px;
            background: var(--border-light);
        }}

        .brand-tag {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: var(--text-muted);
        }}

        .nav-actions {{
            display: flex;
            align-items: center;
            gap: 12px;
            flex-wrap: wrap;
        }}

        .btn {{
            background: var(--text-main);
            color: var(--bg-canvas);
            padding: 11px 24px;
            text-decoration: none;
            border-radius: var(--radius-pill);
            font-weight: 700;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 0.04em;
            transition: transform 0.25s ease, background 0.25s ease, color 0.25s ease;
            border: 2px solid var(--text-main);
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }}

        .btn:hover {{
            background: transparent;
            color: var(--text-main);
            transform: translateY(-3px);
        }}

        .btn-outline {{
            background: transparent;
            color: var(--text-main);
            border: 1.5px solid var(--border-light);
            padding: 10px 18px;
        }}

        .btn-outline:hover {{
            border-color: var(--text-main);
            transform: translateY(-2px);
        }}

        /* CONTAINER PRINCIPAL */
        .main-container {{
            max-width: 1400px;
            width: 100%;
            margin: 0 auto;
            padding: 40px 32px 64px;
            flex: 1;
        }}

        /* CABEÇALHO HERO */
        .hero {{
            display: flex;
            flex-direction: column;
            gap: 16px;
            padding-bottom: 28px;
            border-bottom: 2px solid var(--text-main);
            margin-bottom: 32px;
        }}

        .hero-top {{
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            gap: 24px;
            flex-wrap: wrap;
        }}

        .eyebrow {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.12em;
            color: var(--text-muted);
            margin-bottom: 8px;
            display: block;
        }}

        .hero h1 {{
            font-size: clamp(28px, 4vw, 44px);
            font-weight: 900;
            text-transform: uppercase;
            letter-spacing: -0.03em;
            line-height: 1.05;
        }}

        .hero p {{
            color: var(--text-secondary);
            font-size: 16px;
            max-width: 680px;
            line-height: 1.5;
        }}

        /* FILTROS DE NAVEGAÇÃO INTERNA */
        .filter-bar {{
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }}

        .filter-btn {{
            background: var(--surface-card);
            color: var(--text-secondary);
            border: 1px solid var(--border-light);
            padding: 8px 16px;
            border-radius: var(--radius-pill);
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
        }}

        .filter-btn:hover,
        .filter-btn.active {{
            background: var(--text-main);
            color: var(--bg-canvas);
            border-color: var(--text-main);
        }}


        .kpi-container {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 24px;
            margin-bottom: 28px;
        }}

        .kpi-card {{
            background: var(--surface-card);
            padding: 24px 28px;
            border-radius: var(--radius-lg);
            border: 1px solid var(--border-light);
            box-shadow: var(--shadow-card);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            gap: 10px;
            position: relative;
            overflow: hidden;
            transition: transform var(--transition), border-color var(--transition);
        }}

        .kpi-card::before {{
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            width: 4px;
            height: 100%;
            background: var(--text-main);
        }}

        .kpi-card:hover {{
            transform: translateY(-4px);
            border-color: var(--text-main);
        }}

        .kpi-meta {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 8px;
        }}

        .kpi-index {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            color: var(--text-muted);
            font-weight: 600;
            white-space: nowrap;
        }}

        .kpi-card h2 {{
            font-size: 12px;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-weight: 700;
        }}

        .kpi-value {{
            font-family: 'JetBrains Mono', monospace;
            font-size: clamp(26px, 2.8vw, 36px);
            font-weight: 700;
            color: var(--text-main);
            letter-spacing: -0.03em;
            font-variant-numeric: tabular-nums;
        }}

        .kpi-sub {{
            font-size: 12px;
            color: var(--text-secondary);
        }}

        .bento-grid {{
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 24px;
            align-items: stretch;
        }}

        .col-span-2 {{
            grid-column: 1 / -1;
        }}

        /* CARDS DE GRÁFICOS */
        .chart-card {{
            background: var(--surface-card);
            border-radius: var(--radius-lg);
            padding: 24px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: var(--shadow-card);
            border: 1px solid var(--border-light);
            min-width: 0;
            overflow: hidden;
            transition: transform var(--transition), border-color var(--transition);
        }}

        .chart-card:hover {{
            border-color: var(--text-muted);
        }}

        .card-meta {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 10px;
        }}

        .module-tag {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: var(--text-muted);
            background: var(--surface-subtle);
            padding: 4px 12px;
            border-radius: var(--radius-pill);
            border: 1px solid var(--border-light);
        }}

        .status-dot {{
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: var(--text-main);
            opacity: 0.4;
        }}

        .plot-wrapper {{
            width: 100%;
            min-height: 350px;
            margin-bottom: 14px;
            position: relative;
        }}

        .plot-wrapper .plotly-graph-div,
        .plot-wrapper .js-plotly-plot,
        .plot-wrapper .plot-container {{
            width: 100% !important;
        }}

        /* CAIXA DE INSIGHT */
        .insight-box {{
            background: var(--surface-subtle);
            padding: 14px 16px;
            border-radius: var(--radius-md);
            border-left: 3px solid var(--text-main);
            margin-top: auto;
        }}

        .insight-label {{
            display: block;
            font-family: 'JetBrains Mono', monospace;
            font-size: 10px;
            font-weight: 700;
            letter-spacing: 0.1em;
            color: var(--text-muted);
            margin-bottom: 4px;
        }}

        .insight-text {{
            font-size: 14px;
            line-height: 1.45;
            color: var(--text-secondary);
            font-weight: 500;
        }}

        /* TABELA EXECUTIVA */
        .table-header-bar {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            flex-wrap: wrap;
            gap: 12px;
        }}

        .table-title {{
            font-size: 18px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: -0.02em;
        }}

        .table-responsive {{
            width: 100%;
            overflow-x: auto;
            margin-bottom: 18px;
            border-radius: var(--radius-md);
            border: 1px solid var(--border-light);
        }}

        .custom-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            text-align: left;
        }}

        .custom-table th {{
            background-color: var(--text-main);
            color: var(--bg-canvas);
            padding: 14px 16px;
            text-transform: uppercase;
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            letter-spacing: 0.08em;
            font-weight: 700;
        }}

        .custom-table td {{
            padding: 13px 16px;
            border-bottom: 1px solid var(--border-light);
            color: var(--text-secondary);
            font-variant-numeric: tabular-nums;
        }}

        .custom-table td:nth-child(3),
        .custom-table td:nth-child(4) {{
            font-family: 'JetBrains Mono', monospace;
            font-weight: 600;
            color: var(--text-main);
        }}

        .custom-table tr:last-child td {{
            border-bottom: none;
        }}

        .custom-table tbody tr:hover td {{
            background-color: var(--surface-subtle);
        }}

        .footer {{
            border-top: 1px solid var(--border-light);
            padding: 24px 32px;
            font-size: 12px;
            color: var(--text-muted);
            background: var(--surface-card);
        }}

        .footer-inner {{
            max-width: 1400px;
            margin: 0 auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
        }}

        .hidden-card {{
            display: none !important;
        }}

        @media (max-width: 1024px) {{
            .kpi-container {{
                grid-template-columns: 1fr;
            }}
            .bento-grid {{
                grid-template-columns: 1fr;
            }}
        }}

        @media (max-width: 768px) {{
            .topbar {{
                padding: 14px 20px;
            }}
            .main-container {{
                padding: 24px 20px 48px;
            }}
        }}
    </style>
</head>
<body>

    <!-- BARRA DE NAVEGAÇÃO SUPERIOR -->
    <header class="topbar">
        <div class="topbar-inner">
            <a href="../index.html" class="brand">
                <span class="brand-mark">.Monks</span>
                <span class="brand-divider"></span>
                <span class="brand-tag">CASE OPERATIONS // 2026</span>
            </a>
            <div class="nav-actions">
                <button class="btn btn-outline" id="themeToggle" type="button">Alternar Tema</button>
                <a href="relatorio_erros.html" class="btn btn-outline">Ver Auditoria →</a>
                <a href="../index.html" class="btn">← Voltar para o Início</a>
            </div>
        </div>
    </header>

    <!-- CONTEÚDO PRINCIPAL -->
    <main class="main-container">
        <section class="hero">
            <div class="hero-top">
                <div>
                    <span class="eyebrow">MÓDULO 01 // BUSINESS INTELLIGENCE & PIPELINE</span>
                    <h1>Performance Comercial — Monks 2026</h1>
                </div>
                <div class="filter-bar">
                    <button class="filter-btn active" onclick="filterCards('all', this)">Visão Completa</button>
                    <button class="filter-btn" onclick="filterCards('receita', this)">Receita & Clientes</button>
                    <button class="filter-btn" onclick="filterCards('pipeline', this)">Pipeline & Risco</button>
                    <button class="filter-btn" onclick="filterCards('eficiencia', this)">Eficiência & Conversão</button>
                </div>
            </div>
            <p>Visão executiva consolidada de geração de receita, eficiência de conversão por canal e saúde do funil comercial gerada automaticamente via pipeline Python.</p>
        </section>

        <!-- KPIS EXECUTIVOS -->
        <section class="kpi-container">
            <div class="kpi-card">
                <div class="kpi-meta">
                    <h2>Receita Total (Closed Won)</h2>
                    <span class="kpi-index">KPI // 01</span>
                </div>
                <p class="kpi-value">${total_revenue:,.0f}</p>
                <span class="kpi-sub">Volume financeiro total convertido</span>
            </div>

            <div class="kpi-card">
                <div class="kpi-meta">
                    <h2>Win Rate Médio</h2>
                    <span class="kpi-index">KPI // 02</span>
                </div>
                <p class="kpi-value">{avg_win_rate:.1f}%</p>
                <span class="kpi-sub">Taxa média de conversão entre canais</span>
            </div>

            <div class="kpi-card">
                <div class="kpi-meta">
                    <h2>Volume em Pipeline</h2>
                    <span class="kpi-index">KPI // 03</span>
                </div>
                <p class="kpi-value">${total_pipeline:,.0f}</p>
                <span class="kpi-sub">{total_open_deals} oportunidades ativas em aberto</span>
            </div>
        </section>

        <!-- BENTO GRID UNIFICADO -->
        <section class="bento-grid">
            {card(fig1, insight1, "MÉTRICA 01 // SAZONALIDADE", "receita")}
            {card(fig2, insight2, "MÉTRICA 02 // CANAIS DE AQUISIÇÃO", "receita")}

            <!-- TABELA TOP 10 OPORTUNIDADES (LARGURA TOTAL) -->
            <article class="chart-card col-span-2" data-category="pipeline">
                <div class="card-meta">
                    <span class="module-tag">MÉTRICA 03 // CONCENTRAÇÃO DE PIPELINE</span>
                    <span class="status-dot"></span>
                </div>
                <div class="table-header-bar">
                    <h4 class="table-title">Top 10 Oportunidades em Aberto</h4>
                </div>
                <div class="table-responsive">
                    {table_top_opps}
                </div>
                <div class="insight-box">
                    <span class="insight-label">INSIGHT EXECUTIVO</span>
                    <p class="insight-text">{insight3}</p>
                </div>
            </article>

            {card(fig4, insight4, "MÉTRICA 04 // EFICIÊNCIA", "eficiencia")}
            {card(fig5, insight5, "MÉTRICA 05 // VALOR MÉDIO", "receita")}
            {card(fig6, insight6, "MÉTRICA 06 // FUNIL ATIVO", "pipeline")}
            {card(fig7, insight7, "MÉTRICA 07 // EXPANSÃO VS AQUISIÇÃO", "receita")}
            {card(fig8, insight8, "MÉTRICA 08 // VELOCIDADE REGIONAL", "eficiencia")}
            {card(fig9, insight9, "MÉTRICA 09 // CONTAS ESTRATÉGICAS", "receita")}
            {card(fig10, insight10, "MÉTRICA 10 // AGING DO FUNIL", "pipeline", "col-span-2")}
        </section>
    </main>

    <!-- RODAPÉ -->
    <footer class="footer">
        <div class="footer-inner">
            <span><strong>.Monks</strong> — Operations Case 2026</span>
            <span>Desenvolvido por Mateus Lima Santos</span>
        </div>
    </footer>

    <!-- SCRIPTS DE TEMA E FILTROS -->
    <script>
        const themeToggle = document.getElementById('themeToggle');
        const htmlEl = document.documentElement;

        const LIGHT_PIE_PALETTE = ['#1A1A1A', '#4A4A48', '#787874', '#A8A8A2', '#D4D3CD'];
        const DARK_PIE_PALETTE  = ['#F3F3F1', '#C9C6C1', '#9A9892', '#6E6D68', '#454540'];

        const LIGHT_SERIES_COLORS = ['#1A1A1A', '#888884', '#B5B5B0'];
        const DARK_SERIES_COLORS  = ['#F3F3F1', '#787873', '#4A4A46'];

        function updatePlotlyTheme(isDark) {{
            const textColor = isDark ? '#F3F3F1' : '#1A1A1A';
            const mutedColor = isDark ? '#B5B5B0' : '#5A5A55';
            const gridColor = isDark ? '#2C2C28' : '#E8E7E2';
            const lineColor = isDark ? '#3A3A35' : '#E2E1DC';
            const cardBg    = isDark ? '#1A1A18' : '#FFFFFF';
            const hoverBg   = isDark ? '#F3F3F1' : '#1A1A1A';
            const hoverTxt  = isDark ? '#1A1A1A' : '#F3F3F1';

            const plots = document.querySelectorAll('.js-plotly-plot');
            plots.forEach(plot => {{
                // 1. Atualiza textos, grades e eixos (Layout)
                Plotly.relayout(plot, {{
                    'font.color': textColor,
                    'title.font.color': textColor,
                    'legend.font.color': textColor,
                    'xaxis.tickfont.color': mutedColor,
                    'yaxis.tickfont.color': mutedColor,
                    'xaxis.title.font.color': mutedColor,
                    'yaxis.title.font.color': mutedColor,
                    'xaxis.linecolor': lineColor,
                    'yaxis.linecolor': lineColor,
                    'xaxis.gridcolor': gridColor,
                    'yaxis.gridcolor': gridColor,
                    'yaxis.zerolinecolor': lineColor,
                    'hoverlabel.bgcolor': hoverBg,
                    'hoverlabel.font.color': hoverTxt
                }});

                // 2. Atualiza as cores das barras, histogramas e fatias (Traces)
                if (!plot.data) return;
                plot.data.forEach((trace, idx) => {{
                    if (trace.type === 'pie') {{
                        Plotly.restyle(plot, {{
                            'marker.colors': [isDark ? DARK_PIE_PALETTE : LIGHT_PIE_PALETTE],
                            'marker.line.color': cardBg,
                            'textfont.color': isDark ? '#111110' : '#FFFFFF'
                        }}, [idx]);
                    }} else if (trace.type === 'bar' || trace.type === 'histogram') {{
                        const palette = isDark ? DARK_SERIES_COLORS : LIGHT_SERIES_COLORS;
                        const barColor = palette[idx % palette.length];
                        Plotly.restyle(plot, {{
                            'marker.color': barColor
                        }}, [idx]);
                    }}
                }});
            }});
        }}

        themeToggle.addEventListener('click', () => {{
            const currentTheme = htmlEl.getAttribute('data-theme');
            const nextTheme = currentTheme === 'dark' ? 'light' : 'dark';
            htmlEl.setAttribute('data-theme', nextTheme);
            localStorage.setItem('monks-theme', nextTheme);
            updatePlotlyTheme(nextTheme === 'dark');
        }});

        // Sincroniza o tema salvo e garante redimensionamento correto dos gráficos
        window.addEventListener('load', () => {{
            const savedTheme = localStorage.getItem('monks-theme');
            if (savedTheme === 'dark') {{
                htmlEl.setAttribute('data-theme', 'dark');
                updatePlotlyTheme(true);
            }}
            setTimeout(() => window.dispatchEvent(new Event('resize')), 100);
        }});

        // Filtro dinâmico de categorias
        function filterCards(category, btnEl) {{
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            btnEl.classList.add('active');

            document.querySelectorAll('.chart-card').forEach(card => {{
                if (category === 'all' || card.getAttribute('data-category') === category) {{
                    card.classList.remove('hidden-card');
                }} else {{
                    card.classList.add('hidden-card');
                }}
            }});
            setTimeout(() => window.dispatchEvent(new Event('resize')), 50);
        }}
    </script>
</body>
</html>
"""

os.makedirs('reports', exist_ok=True)
with open('reports/analise.html', 'w', encoding='utf-8') as f:
    f.write(html_final)

print("✅ Análise completa gerada em 'reports/analise.html'!")