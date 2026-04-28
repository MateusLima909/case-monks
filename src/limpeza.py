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
    if row['Erro_Taxonomia']: motivos.append("Fonte de Lead Inválida")
    if row['Erro_Cidade']: motivos.append("Cidade Fora do Padrão")
    if row['Erro_Estagio']: motivos.append("Estágio Inválido")
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
df_final.to_excel('data/processed/opps_corrigido.xlsx', index=False)

erros = df[df['Motivo_do_Erro'] != ""].copy()

# =========================
# MÉTRICAS
# =========================
total_erros = len(erros)
score_qualidade = round((1 - (total_erros / total_linhas)) * 100, 1)

qtd_taxonomia = df['Erro_Taxonomia'].sum()
qtd_cidade_estagio = ((df['Erro_Cidade']) | (df['Erro_Estagio'])).sum()

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

# Transformar a string de erros em badges HTML independentes
def formatar_badges(motivo_str):
    if not motivo_str:
        return ""
    motivos = motivo_str.split(", ")
    badges = [f'<span class="status-badge">{m}</span>' for m in motivos]
    return " ".join(badges)

erros_html['Motivo_do_Erro'] = erros_html['Motivo_do_Erro'].apply(formatar_badges)

# Convertendo para HTML com escape=False para permitir a renderização das tags <span>
tabela_html = erros_html.to_html(index=False, escape=False, border=0, classes="data-table")

# =========================
# TEMPLATE HTML
# =========================
html_template = f"""
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <style>
        :root {{
            --bg-color: #C9C6C1; 
            --text-main: #1A1A1A; 
            --text-secondary: #222222; 
            --text-muted: #888888; 
            --white: #FFFFFF; 
            --border-light: #E5E5E0;
            --hover-bg: #F5F5F0;
        }}

        body {{
            font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            margin: 0; 
            padding: 60px 40px; 
            background-color: var(--bg-color); 
            color: var(--text-main); 
            line-height: 1.6;
        }}

        .container {{
            max-width: 1100px; 
            margin: auto; 
        }}

        .header-title {{
            font-size: 32px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: -1px; 
            margin-bottom: 40px; 
            border-bottom: 3px solid var(--text-main);
            padding-bottom: 15px;
        }}

        .summary {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px; 
            margin-bottom: 40px; 
        }}

        .card {{
            background: var(--text-main); 
            padding: 24px; 
            border-radius: 12px;
            text-align: left;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }}

        .card h2 {{
            color: var(--text-muted); 
            font-size: 11px; 
            font-weight: 700;
            letter-spacing: 1.5px; 
            text-transform: uppercase;
            margin: 0 0 8px 0;
        }}

        .card h3 {{
            color: var(--white); 
            font-size: 36px; 
            font-weight: 800;
            margin: 0; 
        }}

        .section {{
            background: var(--white);
            padding: 32px;
            border-radius: 12px;
            border: 1px solid var(--border-light);
            margin-bottom: 40px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }}

        .section h2 {{
            font-size: 18px;
            font-weight: 800;
            text-transform: uppercase;
            margin-top: 0;
            margin-bottom: 20px;
            color: var(--text-main);
        }}

        .problem-list {{
            list-style: none;
            padding: 0;
            margin: 0 0 24px 0;
        }}

        .problem-list li {{
            font-size: 15px; 
            color: var(--text-secondary);
            padding: 12px 0;
            border-bottom: 1px solid var(--border-light);
        }}

        .problem-list li:last-child {{
            border-bottom: none;
        }}

        .problem-list strong {{
            color: var(--text-main);
            display: inline-block;
            width: 100px;
        }}

        .impact-box {{
            background-color: #F3F3F1;
            padding: 16px 20px;
            border-left: 4px solid var(--text-main);
            border-radius: 0 8px 8px 0;
            font-size: 14px;
            color: var(--text-secondary);
        }}

        .table-section-title {{
            font-size: 14px;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 16px;
        }}
        
        .data-table {{
            width: 100%; 
            border-collapse: separate; 
            border-spacing: 0;
            background-color: var(--white);
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid var(--border-light);
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }}

        .data-table thead th {{
            background-color: var(--text-main); 
            color: var(--white); 
            padding: 16px 20px; 
            font-weight: 700; 
            text-transform: uppercase; 
            font-size: 11px; 
            letter-spacing: 1px;
            text-align: left;
        }}

        .data-table tbody td {{
            padding: 16px 20px; 
            border-bottom: 1px solid var(--border-light); 
            color: var(--text-secondary); 
            font-size: 14px; 
            vertical-align: middle;
        }}

        .data-table tbody tr:last-child td {{
            border-bottom: none;
        }}

        .data-table tbody tr:hover td {{
            background-color: var(--hover-bg); 
        }}
        
        .status-badge {{
            display: inline-block;
            background: var(--white);
            color: var(--text-main);
            padding: 4px 10px;
            border: 1px solid var(--text-main); 
            border-radius: 6px;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            white-space: nowrap;
            margin: 2px 4px 2px 0;
        }}

        .btn-back {{
            position: relative;
            top: -30px;
            left: 90%;
            background: #1A1A1A;
            color: #F3F3F1;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 30px;
            font-size: 12px;
            font-weight: bold;
            text-transform: uppercase;
            border: 1px solid #F3F3F1;
            z-index: 1000;
            transition: 0.3s;
        }}

        .btn-back:hover {{
            background: #F3F3F1;
            color: #1A1A1A;
        }}
    </style>
    <title>Auditoria de Dados | Operations</title>
</head>
<body>
    <a href="../index.html" class="btn-back">Voltar para o Início</a>
    <div class="container">
        <div class="header-title">Relatório de Auditoria de Dados</div>
        
        <div class="summary">
            <div class="card">
                <h2>Total Analisado</h2>
                <h3>{len(df)}</h3>
            </div>
            <div class="card">
                <h2>Registros com Erro</h2>
                <h3>{len(erros)}</h3>
            </div>
            <div class="card">
                <h2>Score de Qualidade</h2>
                <h3>{score_qualidade}%</h3>
            </div>
        </div>

        <div class="section">
            <h2>Principais Problemas</h2>
            <ul class="problem-list">
                <li><strong>Taxonomia:</strong> {qtd_taxonomia} registros ({pct_taxonomia}%) fora do padrão &rarr; normalizados.</li>
                <li><strong>Financeiro:</strong> divergência entre valores &rarr; corrigido via soma dos produtos.</li>
                <li><strong>Estrutura:</strong> {qtd_cidade_estagio} registros inválidos &rarr; removidos.</li>
            </ul>
            <div class="impact-box">
                Foram identificadas inconsistências que impactavam diretamente a confiabilidade do pipeline comercial. Após o tratamento, a base foi padronizada, permitindo análises mais precisas de receita, canais de aquisição e performance operacional.
            </div>
        </div>

        <div class="table-section-title">
            Exemplos reais de registros com inconsistências:
        </div>
        
        {tabela_html}

    </div>
</body>
</html>
"""

with open('reports/relatorio_erros.html', 'w', encoding='utf-8') as f:
    f.write(html_template)

print("✅ Limpeza e relatório concluídos!")