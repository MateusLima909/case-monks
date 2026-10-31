import os
import pandas as pd

# =========================
# LEITURA
# =========================
df = pd.read_excel('data/raw/opps_corrupted.xlsx')
total_linhas = len(df)

# =========================
# LIMPEZA NUMÉRICA
# =========================
for col in ['Amount', 'Total_Product_Amount']:
    df[col] = (
        df[col]
        .astype(str)
        .str.replace(',', '.')
        .str.strip()
    )
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# =========================
# CORREÇÃO DE AMOUNT
# =========================
real_amounts = df.groupby('Opportunity_ID')['Total_Product_Amount'].sum().reset_index()
real_amounts.columns = ['Opportunity_ID', 'Corrected_Amount']

df = df.merge(real_amounts, on='Opportunity_ID', how='left')
df['Amount_Original'] = df['Amount']
df['Amount'] = df['Corrected_Amount']
df = df.drop(columns=['Corrected_Amount'])

# =========================
# NORMALIZAÇÃO LEAD SOURCE
# =========================
df['Lead_Source_Normalizado'] = df['Lead_Source'].astype(str).str.strip().str.lower()

mapeamento_taxonomia = {
    'inbound - website - media.monks (deprecated)': 'Inbound',
    'inbound - marketing (deprecated)': 'Inbound',
    'inbound - website (deprecated)': 'Inbound',
    'marketing - lead scoring/nurturing': 'Inbound',
    'website - sales form': 'Inbound',
    'prospecting - personal (deprecated)': 'Outbound',
    'referral - internal': 'Referral',
    'referral - personal': 'Referral',
    'referral - s4 company': 'Referral',
    'referral - partner - google marketing platform': 'Referral',
    'referral - partner - google cloud': 'Referral',
    'inbound - current client (deprecated)': 'Customer Success',
    'existing client - account/growth activity': 'Customer Success',
    'prospecting - current client (deprecated)': 'Customer Success',
    'industry event': 'Event',
    'mh lead (deprecated)': 'Unknown',
    "don't know/other (deprecated)": 'Unknown'
}

df['Lead_Source_Category'] = df['Lead_Source_Normalizado'].map(mapeamento_taxonomia)

# =========================
# VALIDAÇÕES
# =========================
cidades_canonicas = ['Sao Carlos, BR', 'Votorantim, BR', 'Sao Paulo, BR']
estagios_canonicos = [
    'Opportunity Identified', 'Qualified', 'Registration',
    'Pitching', 'Pitched', 'Negotiation', 'Closed Won'
]

df['Erro_Cidade'] = ~df['Lead_Office'].isin(cidades_canonicas)
df['Erro_Estagio'] = ~df['Stage'].isin(estagios_canonicos)
df['Erro_Taxonomia'] = df['Lead_Source_Category'].isna()

def definir_motivo(row):
    motivos = []
    if row['Erro_Taxonomia']:
        motivos.append("Fonte de Lead Inválida")
    if row['Erro_Cidade']:
        motivos.append("Cidade Fora do Padrão")
    if row['Erro_Estagio']:
        motivos.append("Estágio Inválido")
    return ", ".join(motivos)

df['Motivo_do_Erro'] = df.apply(definir_motivo, axis=1)

# =========================
# FILTRO DE TIPOS
# =========================
tipos_validos = [
    'Project - Competitive',
    'Project - Not Competitive',
    'Change Order/Upsell'
]

df = df[df['Type'].isin(tipos_validos)]

# =========================
# DEDUPLICAÇÃO
# =========================
df_final = df[df['Motivo_do_Erro'] == ""].copy()

df_final = df_final.groupby('Opportunity_ID').agg({
    'Account_Name': 'first',
    'Stage': 'first',
    'Lead_Source_Category': 'first',
    'Lead_Office': 'first',
    'Type': 'first',
    'Created_Date': 'first',
    'Close_Date': 'first',
    'Amount': 'first'
}).reset_index()

total_deals = len(df_final)

# =========================
# EXPORTAÇÃO
# =========================
os.makedirs('data/processed', exist_ok=True)
df_final.to_excel('data/processed/opps_corrigido.xlsx', index=False)

erros = df[df['Motivo_do_Erro'] != ""].copy()

# =========================
# MÉTRICAS
# =========================
total_erros = len(erros)
score_qualidade = round((1 - (total_erros / total_linhas)) * 100, 1)

qtd_taxonomia = int(df['Erro_Taxonomia'].sum())
qtd_cidade_estagio = int(((df['Erro_Cidade']) | (df['Erro_Estagio'])).sum())

pct_taxonomia = round((qtd_taxonomia / total_linhas) * 100, 1)

# =========================
# TABELA HTML E FORMATAÇÃO
# =========================
colunas_auditoria = [
    'Opportunity_ID',
    'Account_Name',
    'Lead_Source',
    'Lead_Office',
    'Stage',
    'Motivo_do_Erro'
]

erros_html = erros[colunas_auditoria].head(10).copy()

def formatar_badges(motivo_str):
    if not motivo_str:
        return ""
    motivos = motivo_str.split(", ")
    badges = [f'<span class="status-badge">{m}</span>' for m in motivos]
    return " ".join(badges)

erros_html['Motivo_do_Erro'] = erros_html['Motivo_do_Erro'].apply(formatar_badges)
erros_html.columns = [
    'Opportunity ID',
    'Cliente',
    'Fonte de Lead (Bruta)',
    'Escritório',
    'Estágio',
    'Anomalias Detectadas'
]

tabela_html = erros_html.to_html(index=False, escape=False, border=0, classes="data-table")

# =========================
# TEMPLATE HTML (.MONKS DESIGN SYSTEM)
# =========================
html_template = f"""<!DOCTYPE html>
<html lang="pt-br" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Auditoria de Dados | .Monks Operations Case</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@500;600;700&display=swap" rel="stylesheet">
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
            --text-muted: #787873;
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
            line-height: 1.6;
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

        /* BOTÕES PÍLULA (.BTN IGUAL À PÁGINA INICIAL) */
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
            gap: 14px;
            padding-bottom: 28px;
            border-bottom: 2px solid var(--text-main);
            margin-bottom: 32px;
        }}

        .eyebrow {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.12em;
            color: var(--text-muted);
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
            max-width: 760px;
            line-height: 1.5;
        }}

        /* KPIS DE AUDITORIA (BENTO SUMMARY) */
        .summary {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 20px;
            margin-bottom: 32px;
        }}

        .kpi-card {{
            background: var(--surface-card);
            padding: 24px 26px;
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
        }}

        .kpi-index {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            color: var(--text-muted);
            font-weight: 600;
        }}

        .kpi-card h2 {{
            color: var(--text-muted);
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }}

        .kpi-card h3 {{
            font-family: 'JetBrains Mono', monospace;
            color: var(--text-main);
            font-size: clamp(28px, 3vw, 36px);
            font-weight: 700;
            letter-spacing: -0.03em;
            font-variant-numeric: tabular-nums;
            line-height: 1.1;
        }}

        .kpi-sub {{
            font-size: 12px;
            color: var(--text-secondary);
        }}

        /* SEÇÃO DE DIAGNÓSTICO E REGRAS DE NEGÓCIO */
        .bento-section {{
            background: var(--surface-card);
            padding: 28px;
            border-radius: var(--radius-lg);
            border: 1px solid var(--border-light);
            margin-bottom: 32px;
            box-shadow: var(--shadow-card);
        }}

        .card-meta {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
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

        .section-title {{
            font-size: 20px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: -0.02em;
            margin-bottom: 20px;
            color: var(--text-main);
        }}

        .problem-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }}

        .problem-item {{
            background: var(--surface-subtle);
            border: 1px solid var(--border-light);
            border-radius: var(--radius-md);
            padding: 18px 20px;
            display: flex;
            flex-direction: column;
            gap: 6px;
        }}

        .problem-tag {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: var(--text-main);
        }}

        .problem-desc {{
            font-size: 14px;
            color: var(--text-secondary);
            line-height: 1.45;
        }}

        .impact-box {{
            background: var(--surface-subtle);
            padding: 16px 20px;
            border-left: 3px solid var(--text-main);
            border-radius: var(--radius-md);
        }}

        .impact-label {{
            display: block;
            font-family: 'JetBrains Mono', monospace;
            font-size: 10px;
            font-weight: 700;
            letter-spacing: 0.1em;
            color: var(--text-muted);
            margin-bottom: 4px;
            text-transform: uppercase;
        }}

        .impact-text {{
            font-size: 14px;
            color: var(--text-secondary);
            font-weight: 500;
        }}

        /* TABELA DE AUDITORIA */
        .table-toolbar {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 16px;
            margin-bottom: 18px;
            flex-wrap: wrap;
        }}

        .search-input {{
            background: var(--surface-subtle);
            border: 1px solid var(--border-light);
            color: var(--text-main);
            padding: 10px 16px;
            border-radius: var(--radius-pill);
            font-family: 'Inter', sans-serif;
            font-size: 13px;
            min-width: 280px;
            outline: none;
            transition: border-color 0.2s ease;
        }}

        .search-input:focus {{
            border-color: var(--text-main);
        }}

        .table-responsive {{
            width: 100%;
            overflow-x: auto;
            border-radius: var(--radius-md);
            border: 1px solid var(--border-light);
        }}

        .data-table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 13px;
        }}

        .data-table thead th {{
            background-color: var(--text-main);
            color: var(--bg-canvas);
            padding: 14px 18px;
            font-family: 'JetBrains Mono', monospace;
            font-weight: 700;
            text-transform: uppercase;
            font-size: 11px;
            letter-spacing: 0.08em;
        }}

        .data-table tbody td {{
            padding: 14px 18px;
            border-bottom: 1px solid var(--border-light);
            color: var(--text-secondary);
            vertical-align: middle;
        }}

        .data-table tbody td:first-child {{
            font-family: 'JetBrains Mono', monospace;
            font-weight: 600;
            color: var(--text-main);
            font-size: 12px;
        }}

        .data-table tbody tr:last-child td {{
            border-bottom: none;
        }}

        .data-table tbody tr:hover td {{
            background-color: var(--surface-subtle);
        }}

        .status-badge {{
            display: inline-block;
            background: var(--surface-subtle);
            color: var(--text-main);
            padding: 4px 10px;
            border: 1px solid var(--text-main);
            border-radius: var(--radius-pill);
            font-family: 'JetBrains Mono', monospace;
            font-size: 10px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            white-space: nowrap;
            margin: 2px 4px 2px 0;
        }}

        /* RODAPÉ PADRONIZADO */
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

        @media (max-width: 768px) {{
            .topbar {{
                padding: 14px 20px;
            }}
            .main-container {{
                padding: 24px 20px 48px;
            }}
            .search-input {{
                width: 100%;
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
                <a href="analise.html" class="btn btn-outline">Ver Dashboard →</a>
                <a href="../index.html" class="btn">← Voltar para o Início</a>
            </div>
        </div>
    </header>

    <!-- CONTEÚDO PRINCIPAL -->
    <main class="main-container">
        <section class="hero">
            <span class="eyebrow">MÓDULO 02 // DATA QUALITY & GOVERNANÇA ETL</span>
            <h1>Relatório de Auditoria de Dados</h1>
            <p>Inspeção automatizada de integridade cadastral, reconciliação financeira por soma de produtos, normalização de taxonomia de canais e expurgo de registros inconsistentes.</p>
        </section>

        <!-- KPIS DE QUALIDADE -->
        <section class="summary">
            <div class="kpi-card">
                <div class="kpi-meta">
                    <h2>Total Analisado</h2>
                    <span class="kpi-index">AUDIT // 01</span>
                </div>
                <h3>{len(df)}</h3>
                <span class="kpi-sub">De {total_linhas} linhas brutas iniciais</span>
            </div>

            <div class="kpi-card">
                <div class="kpi-meta">
                    <h2>Registros com Erro</h2>
                    <span class="kpi-index">AUDIT // 02</span>
                </div>
                <h3>{len(erros)}</h3>
                <span class="kpi-sub">Oportunidades retidas na malha de validação</span>
            </div>

            <div class="kpi-card">
                <div class="kpi-meta">
                    <h2>Score de Qualidade</h2>
                    <span class="kpi-index">AUDIT // 03</span>
                </div>
                <h3>{score_qualidade}%</h3>
                <span class="kpi-sub">Índice de conformidade da base</span>
            </div>

            <div class="kpi-card">
                <div class="kpi-meta">
                    <h2>Deals Validados</h2>
                    <span class="kpi-index">AUDIT // 04</span>
                </div>
                <h3>{total_deals}</h3>
                <span class="kpi-sub">Oportunidades únicas em opps_corrigido.xlsx</span>
            </div>
        </section>

        <!-- DIAGNÓSTICO DE PROBLEMAS -->
        <section class="bento-section">
            <div class="card-meta">
                <span class="module-tag">DIAGNÓSTICO // REGRAS DE HIGIENIZAÇÃO</span>
                <span class="status-dot"></span>
            </div>
            <h2 class="section-title">Principais Problemas Detectados e Tratamento</h2>

            <div class="problem-grid">
                <div class="problem-item">
                    <span class="problem-tag">01 // Taxonomia</span>
                    <p class="problem-desc"><strong>{qtd_taxonomia} registros ({pct_taxonomia}%)</strong> apresentavam nomenclatura legada ou fora do padrão &rarr; mapeados e normalizados em categorias canônicas.</p>
                </div>

                <div class="problem-item">
                    <span class="problem-tag">02 // Financeiro</span>
                    <p class="problem-desc">Divergência identificada entre o valor de cabeçalho e os itens &rarr; <strong>Amount</strong> recalculado via soma exata de <code>Total_Product_Amount</code> por deal.</p>
                </div>

                <div class="problem-item">
                    <span class="problem-tag">03 // Estrutura</span>
                    <p class="problem-desc"><strong>{qtd_cidade_estagio} registros inválidos</strong> com escritórios ou estágios do funil fora do padrão operacional &rarr; isolados e expurgados da base analítica.</p>
                </div>
            </div>

            <div class="impact-box">
                <span class="impact-label">IMPACTO EXECUTIVO DA HIGIENIZAÇÃO</span>
                <p class="impact-text">Foram identificadas inconsistências que impactavam diretamente a confiabilidade do pipeline comercial. Após o tratamento e a deduplicação por Opportunity_ID, a base foi padronizada, permitindo análises precisas de receita, canais de aquisição e performance operacional.</p>
            </div>
        </section>

        <!-- TABELA DE REGISTROS COM INCONSISTÊNCIAS -->
        <section class="bento-section" style="margin-bottom: 0;">
            <div class="card-meta">
                <span class="module-tag">AMOSTRAGEM // LOG DE EXCEÇÕES (TOP 10)</span>
                <span class="status-dot"></span>
            </div>
            <div class="table-toolbar">
                <h2 class="section-title" style="margin-bottom: 0;">Exemplos Reais de Registros com Inconsistências</h2>
                <input type="text" id="auditSearch" class="search-input" placeholder="Filtrar por cliente, escritório ou erro...">
            </div>

            <div class="table-responsive">
                {tabela_html}
            </div>
        </section>
    </main>

    <!-- RODAPÉ PADRONIZADO -->
    <footer class="footer">
        <div class="footer-inner">
            <span><strong>.Monks</strong> — Operations Case 2026</span>
            <span>Desenvolvido por Mateus Lima Santos</span>
        </div>
    </footer>

    <!-- SCRIPTS DE TEMA E BUSCA NA TABELA -->
    <script>
        const themeToggle = document.getElementById('themeToggle');
        const htmlEl = document.documentElement;

        themeToggle.addEventListener('click', () => {{
            const currentTheme = htmlEl.getAttribute('data-theme');
            const nextTheme = currentTheme === 'dark' ? 'light' : 'dark';
            htmlEl.setAttribute('data-theme', nextTheme);
            localStorage.setItem('monks-theme', nextTheme);
        }});

        // Sincroniza o tema salvo na página inicial ou no dashboard
        const savedTheme = localStorage.getItem('monks-theme');
        if (savedTheme === 'dark') {{
            htmlEl.setAttribute('data-theme', 'dark');
        }}

        // Filtro rápido na tabela de auditoria
        const searchInput = document.getElementById('auditSearch');
        searchInput.addEventListener('input', function() {{
            const query = this.value.toLowerCase();
            const rows = document.querySelectorAll('.data-table tbody tr');
            rows.forEach(row => {{
                const text = row.innerText.toLowerCase();
                row.style.display = text.includes(query) ? '' : 'none';
            }});
        }});
    </script>
</body>
</html>
"""

os.makedirs('reports', exist_ok=True)
with open('reports/relatorio_erros.html', 'w', encoding='utf-8') as f:
    f.write(html_template)

print("✅ Limpeza e relatório concluídos em 'reports/relatorio_erros.html'!")