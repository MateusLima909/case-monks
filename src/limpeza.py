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
# TABELA HTML
# =========================
colunas_auditoria = [
    'Opportunity_ID',
    'Account_Name',
    'Lead_Source',
    'Lead_Office',
    'Stage',
    'Motivo_do_Erro'
]

tabela_html = erros[colunas_auditoria].head(10).to_html(index=False)

for motivo in ["Fonte de Lead Inválida", "Cidade Fora do Padrão", "Estágio Inválido"]:
    tabela_html = tabela_html.replace(
        motivo,
        f'<span class="status-badge">{motivo}</span>'
    )

html_template = f"""
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <style>
        body {{ 
            font-family: 'Inter', 'Segoe UI', Roboto, sans-serif; 
            margin: 40px; 
            background-color: #0f172a; 
            color: #f8fafc; 
        }}
        .container {{ 
            max-width: 1200px; 
            margin: auto; 
            background: #1e293b; 
            padding: 40px; 
            border-radius: 16px; 
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.5); 
        }}
        h1 {{ 
            color: #f97316; 
            font-size: 28px; 
            margin-bottom: 30px; 
            display: flex; 
            align-items: center; 
            gap: 10px;
        }}
        .summary {{ 
            display: flex; 
            gap: 20px; 
            margin-bottom: 40px; 
        }}
        .card {{ 
            background: #334155; 
            padding: 20px; 
            border-radius: 12px; 
            border-top: 4px solid #f97316; 
            flex: 1; 
            transition: transform 0.2s;
        }}
        .card:hover {{ transform: translateY(-5px); }}
        .card h2 {{ margin: 0; color: #94a3b8; font-size: 12px; letter-spacing: 1px; }}
        .card h3 {{ margin: 8px 0 0; font-weight: 800; color: #f8fafc; }}

        .section {{
            margin-top: 30px;
        }}

        .section h2 {{
            color: #f97316;
            margin-bottom: 10px;
        }}

        .section p {{
            color: #cbd5e1;
            line-height: 1.6;
            margin-bottom: 10px;
        }}
        
        table {{ 
            width: 100%; 
            border-collapse: separate; 
            border-spacing: 0; 
            margin-top: 20px; 
            border-radius: 8px; 
            overflow: hidden; 
        }}
        th {{ 
            background-color: #f97316; 
            color: #ffffff; 
            text-align: left; 
            padding: 16px; 
            font-weight: 600; 
            text-transform: uppercase; 
            font-size: 12px; 
        }}
        td {{ 
            padding: 14px 16px; 
            border-bottom: 1px solid #334155; 
            background-color: #1e293b; 
            color: #cbd5e1; 
            font-size: 14px; 
        }}
        tr:hover td {{ background-color: #334155; color: #f8fafc; }}
        
        .status-badge {{
            background: rgba(249, 115, 22, 0.1);
            color: #f97316;
            padding: 4px 12px;
            border-radius: 9999px;
            font-size: 12px;
            font-weight: 600;
        }}
    </style>
    <title>Auditoria de Dados | Operations</title>
</head>
<body>
    <div class="container">
        <h1><span>📊</span> Relatório de Auditoria de Dados</h1>
        
        <div class="summary">
            <div class="card">
                <h2>TOTAL ANALISADO</h2>
                <h3>{len(df)}</h3>
            </div>
            <div class="card">
                <h2>REGISTROS COM ERRO</h2>
                <h3>{len(erros)}</h3>
            </div>
            <div class="card">
                <h2>SCORE DE QUALIDADE</h2>
                <h3>{round((1 - (len(erros)/len(df)))*100, 1)}%</h3>
            </div>
        </div>

        <div class="card">
            <h2 style="color: #f97316; font-size: 20px;">Principais Problemas</h2>
            <p><b>Taxonomia:</b> {qtd_taxonomia} registros ({pct_taxonomia}%) fora do padrão → normalizados.</p>
            <p><b>Financeiro:</b> divergência entre valores → corrigido via soma dos produtos.</p>
            <p><b>Estrutura:</b> {qtd_cidade_estagio} registros inválidos → removidos.</p>

            <p>
                Foram identificadas inconsistências que impactavam diretamente a confiabilidade do pipeline comercial. 
                Após o tratamento, a base foi padronizada, permitindo análises mais precisas de receita, canais de aquisição 
                e performance operacional.
            </p>
        </div>

        <p style="color: #94a3b8; margin-top: 30px;">
            Exemplos reais de registros com inconsistências identificadas:
        </p>
        
        {tabela_html}

    </div>
</body>
</html>
"""

with open('reports/relatorio_erros.html', 'w', encoding='utf-8') as f:
    f.write(html_template)

print("✅ Limpeza e relatório concluídos!")