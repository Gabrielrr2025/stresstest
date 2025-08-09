import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
import math
import datetime
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Finhealth • VaR Calculator", page_icon="📊", layout="wide")

# =============== CSS MODERNO E ELEGANTE ===============
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

:root {
    --primary: #0066ff;
    --primary-dark: #0052d4;
    --success: #10b981;
    --warning: #f59e0b;
    --error: #ef4444;
    --bg-main: #fafbfc;
    --bg-card: #ffffff;
    --text-primary: #111827;
    --text-secondary: #6b7280;
    --text-muted: #9ca3af;
    --border: #e5e7eb;
    --border-light: #f3f4f6;
    --shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
    --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
    --gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    --gradient-success: linear-gradient(135deg, #10b981 0%, #059669 100%);
}

* {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* Background e layout principal */
html, body, [data-testid="stAppViewContainer"] {
    background: var(--bg-main);
    color: var(--text-primary);
}

/* Header principal */
.main-header {
    background: var(--gradient);
    color: white;
    padding: 2rem 2.5rem;
    border-radius: 20px;
    margin-bottom: 2rem;
    box-shadow: var(--shadow-lg);
    text-align: center;
}

.main-header h1 {
    font-size: 2.5rem;
    font-weight: 700;
    margin: 0 0 0.5rem 0;
    text-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.main-header .subtitle {
    font-size: 1.1rem;
    opacity: 0.9;
    font-weight: 400;
}

/* Cards e containers */
.section-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 1.5rem;
    margin-bottom: 1.5rem;
    box-shadow: var(--shadow);
    transition: all 0.3s ease;
}

.section-card:hover {
    box-shadow: var(--shadow-lg);
    transform: translateY(-2px);
}

.section-title {
    font-size: 1.25rem;
    font-weight: 600;
    color: var(--text-primary);
    margin-bottom: 1rem;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid var(--border-light);
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

/* KPIs modernos */
.kpi-container {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 1rem;
    margin-bottom: 2rem;
}

.kpi-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.5rem;
    text-align: center;
    box-shadow: var(--shadow);
    transition: all 0.3s ease;
    position: relative;
    overflow: hidden;
}

.kpi-card::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 4px;
    background: var(--gradient);
}

.kpi-card:hover {
    transform: translateY(-3px);
    box-shadow: var(--shadow-lg);
}

.kpi-value {
    font-size: 2rem;
    font-weight: 700;
    color: var(--primary);
    margin-bottom: 0.25rem;
}

.kpi-label {
    font-size: 0.875rem;
    color: var(--text-secondary);
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.kpi-subtitle {
    font-size: 0.75rem;
    color: var(--text-muted);
    margin-top: 0.25rem;
}

/* Status badges */
.status-badge {
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
    padding: 0.5rem 1rem;
    border-radius: 8px;
    font-size: 0.875rem;
    font-weight: 500;
    margin-bottom: 1rem;
}

.status-success {
    background: rgba(16, 185, 129, 0.1);
    color: var(--success);
    border: 1px solid rgba(16, 185, 129, 0.2);
}

.status-warning {
    background: rgba(245, 158, 11, 0.1);
    color: var(--warning);
    border: 1px solid rgba(245, 158, 11, 0.2);
}

.status-error {
    background: rgba(239, 68, 68, 0.1);
    color: var(--error);
    border: 1px solid rgba(239, 68, 68, 0.2);
}

/* Progress bar customizada */
.progress-container {
    background: var(--border-light);
    border-radius: 8px;
    height: 8px;
    overflow: hidden;
    margin: 1rem 0;
}

.progress-bar {
    height: 100%;
    background: var(--gradient-success);
    border-radius: 8px;
    transition: width 0.3s ease;
}

/* Botões customizados */
.stButton > button {
    background: var(--gradient) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.75rem 2rem !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
    transition: all 0.3s ease !important;
    box-shadow: var(--shadow) !important;
}

.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: var(--shadow-lg) !important;
}

.stButton > button:disabled {
    background: var(--text-muted) !important;
    cursor: not-allowed !important;
    transform: none !important;
}

/* Form inputs */
.stNumberInput > div > div > input,
.stTextInput > div > div > input,
.stSelectbox > div > div > select,
.stDateInput > div > div > input {
    border: 2px solid var(--border) !important;
    border-radius: 8px !important;
    padding: 0.75rem 1rem !important;
    font-size: 0.875rem !important;
    transition: border-color 0.3s ease !important;
}

.stNumberInput > div > div > input:focus,
.stTextInput > div > div > input:focus,
.stSelectbox > div > div > select:focus,
.stDateInput > div > div > input:focus {
    border-color: var(--primary) !important;
    box-shadow: 0 0 0 3px rgba(0, 102, 255, 0.1) !important;
}

/* Dataframes */
.stDataFrame {
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow: hidden;
    box-shadow: var(--shadow);
}

/* Charts */
.js-plotly-plot {
    border-radius: 12px;
    overflow: hidden;
    box-shadow: var(--shadow);
}

/* Sidebar */
.css-1d391kg {
    background: var(--bg-card) !important;
    border-right: 1px solid var(--border) !important;
}

/* Expanders */
.streamlit-expanderHeader {
    background: var(--border-light) !important;
    border-radius: 8px !important;
    margin-bottom: 0.5rem !important;
}

/* Footer */
.footer {
    text-align: center;
    padding: 2rem;
    color: var(--text-muted);
    font-size: 0.875rem;
    margin-top: 3rem;
    border-top: 1px solid var(--border);
}

/* Responsive */
@media (max-width: 768px) {
    .main-header {
        padding: 1.5rem;
    }
    
    .main-header h1 {
        font-size: 2rem;
    }
    
    .kpi-container {
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
        gap: 0.75rem;
    }
    
    .section-card {
        padding: 1rem;
    }
}

/* Animations */
@keyframes slideIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.section-card {
    animation: slideIn 0.3s ease-out;
}

/* Hide Streamlit elements */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.stDeployButton {display: none;}
</style>
""", unsafe_allow_html=True)

# =============== HEADER PRINCIPAL ===============
st.markdown("""
<div class="main-header">
    <h1>📊 Finhealth VaR Calculator</h1>
    <div class="subtitle">Análise de Risco Paramétrica por Classe de Ativo • Compliance CVM/B3</div>
</div>
""", unsafe_allow_html=True)

# =============== CONFIGURAÇÕES E DADOS ===============

VOL_PADRAO = {
    "Ações (Ibovespa)": 0.25,
    "Juros-Pré": 0.08,
    "Câmbio (Dólar)": 0.15,
    "Cupom Cambial": 0.12,
    "Crédito Privado": 0.05,
    "Multimercado": 0.18,
    "Outros": 0.10
}

CENARIOS_PADRAO = {
    "Ibovespa": -0.15,
    "Juros-Pré": 0.02,
    "Cupom Cambial": -0.01,
    "Dólar": -0.05,
    "Outros": -0.03
}

DESC_CENARIO = {
    "Ibovespa": "Queda de 15% no IBOVESPA",
    "Juros-Pré": "Alta de 200 bps na taxa de juros",
    "Cupom Cambial": "Queda de 1% no cupom cambial",
    "Dólar": "Queda de 5% no dólar",
    "Outros": "Queda de 3% em outros ativos"
}

# =============== LAYOUT PRINCIPAL ===============

left, right = st.columns([1.1, 2.2])

with left:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">🏢 Dados do Fundo</div>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        cnpj = st.text_input("CNPJ *", placeholder="00.000.000/0001-00")
        data_ref = st.date_input("Data de Referência *", value=datetime.date.today())
    with col2:
        nome_fundo = st.text_input("Nome do Fundo *", placeholder="Fundo XPTO")
        pl = st.number_input("Patrimônio Líquido (R$) *", min_value=0.0, value=1000000.0, step=1000.0, format="%.2f")

    st.markdown("""
    <div class="section-card">
        <div class="section-title">⚙️ Parâmetros de VaR</div>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        horizonte_dias = st.selectbox("Horizonte", [1, 10, 21], index=2)
    with col2:
        conf_label = st.selectbox("Confiança", ["95%", "99%"], index=0)
    
    z = 1.65 if conf_label == "95%" else 2.33

    st.markdown("""
    <div class="section-card">
        <div class="section-title">📊 Alocação da Carteira</div>
    </div>
    """, unsafe_allow_html=True)
    
    carteira, soma = [], 0.0
    
    for classe, vol in VOL_PADRAO.items():
        col1, col2 = st.columns([2, 1])
        with col1:
            perc = st.number_input(
                f"{classe} (%)", 
                0.0, 100.0, 0.0, 1.0, 
                key=f"p_{classe}",
                help=f"Volatilidade padrão: {vol:.0%}"
            )
        with col2:
            vol_a = st.number_input(
                "Vol", 
                0.0, 1.0, float(vol), 0.01, 
                format="%.2f", 
                key=f"v_{classe}",
                help="Volatilidade anual"
            )
        
        if perc > 0:
            carteira.append({
                "classe": classe,
                "%PL": perc,
                "vol_anual": float(vol_a)
            })
            soma += perc

    # Progress bar da alocação
    progress_pct = min(soma / 100, 1.0)
    st.markdown(f"""
    <div class="progress-container">
        <div class="progress-bar" style="width: {progress_pct * 100}%"></div>
    </div>
    """, unsafe_allow_html=True)

    # Status da alocação
    if soma == 100:
        st.markdown('<div class="status-badge status-success">✅ Alocação perfeita: 100%</div>', unsafe_allow_html=True)
    elif soma > 100:
        st.markdown(f'<div class="status-badge status-error">❌ Alocação excede: {soma:.1f}%</div>', unsafe_allow_html=True)
    elif soma > 0:
        st.markdown(f'<div class="status-badge status-warning">⚠️ Alocação parcial: {soma:.1f}%</div>', unsafe_allow_html=True)

    # Validação
    campos_ok = bool(cnpj.strip() and nome_fundo.strip() and pl > 0)
    pode_calcular = campos_ok and soma > 0 and soma <= 100

    calcular = st.button("🚀 Calcular VaR & Compliance", disabled=not pode_calcular)

with right:
    if not pode_calcular:
        st.markdown("""
        <div class="section-card">
            <div class="section-title">ℹ️ Instruções</div>
            <p>Para calcular o VaR, preencha:</p>
            <ul>
                <li><strong>CNPJ e Nome do Fundo</strong></li>
                <li><strong>Patrimônio Líquido</strong> maior que zero</li>
                <li><strong>Alocação da carteira</strong> (soma entre 1% e 100%)</li>
            </ul>
            <p>O sistema calculará automaticamente:</p>
            <ul>
                <li>VaR paramétrico por classe</li>
                <li>Cenários de estresse</li>
                <li>Respostas para CVM/B3</li>
                <li>Relatórios em Excel</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    elif calcular:
        # =============== CÁLCULOS DE VAR ===============
        
        for item in carteira:
            vol_d = item["vol_anual"] / np.sqrt(252)
            var_pct = z * vol_d * np.sqrt(horizonte_dias)
            item["VaR_%"] = var_pct * 100
            item["VaR_R$"] = pl * (item["%PL"] / 100) * var_pct
        
        df_var = pd.DataFrame(carteira)
        var_total = df_var["VaR_R$"].sum()
        var_total_pct = (var_total / pl) if pl > 0 else 0.0

        # =============== KPIS PRINCIPAIS ===============
        
        st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card">
                <div class="kpi-value">{var_total_pct*100:.2f}%</div>
                <div class="kpi-label">VaR ({conf_label} / {horizonte_dias}d)</div>
                <div class="kpi-subtitle">Método Paramétrico</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value">R$ {var_total:,.0f}</div>
                <div class="kpi-label">VaR em Reais</div>
                <div class="kpi-subtitle">Perda potencial</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value">{len(carteira)}</div>
                <div class="kpi-label">Classes Ativas</div>
                <div class="kpi-subtitle">Em uso na carteira</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value">{soma:.1f}%</div>
                <div class="kpi-label">Alocação Total</div>
                <div class="kpi-subtitle">do patrimônio</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # =============== TABELA DE RESULTADOS ===============
        
        st.markdown("""
        <div class="section-card">
            <div class="section-title">📈 VaR por Classe de Ativo</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Formatar tabela
        df_display = df_var.copy()
        df_display['%PL'] = df_display['%PL'].apply(lambda x: f"{x:.1f}%")
        df_display['vol_anual'] = df_display['vol_anual'].apply(lambda x: f"{x:.2%}")
        df_display['VaR_%'] = df_display['VaR_%'].apply(lambda x: f"{x:.2f}%")
        df_display['VaR_R$'] = df_display['VaR_R$'].apply(lambda x: f"R$ {x:,.0f}")
        
        df_display.columns = ['Classe de Ativo', 'Alocação', 'Volatilidade Anual', 'VaR (%)', 'VaR (R$)']
        
        st.dataframe(df_display, use_container_width=True)

        # =============== GRÁFICOS ===============
        
        col1, col2 = st.columns(2)
        
        with col1:
            fig_pie = px.pie(
                df_var, 
                values="%PL", 
                names="classe",
                title="Distribuição da Carteira",
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig_pie.update_layout(
                font=dict(family="Inter, sans-serif"),
                title_font_size=16,
                legend=dict(orientation="v", y=0.5)
            )
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col2:
            fig_bar = px.bar(
                df_var, 
                x="classe", 
                y="VaR_R$",
                title="VaR por Classe (R$)",
                color="VaR_R$",
                color_continuous_scale="Blues"
            )
            fig_bar.update_layout(
                font=dict(family="Inter, sans-serif"),
                title_font_size=16,
                xaxis_title="",
                yaxis_title="VaR (R$)"
            )
            fig_bar.update_xaxis(tickangle=45)
            st.plotly_chart(fig_bar, use_container_width=True)

        # =============== CENÁRIOS DE ESTRESSE ===============
        
        st.markdown("""
        <div class="section-card">
            <div class="section-title">⚠️ Cenários de Estresse</div>
        </div>
        """, unsafe_allow_html=True)
        
        res_estresse = []
        for fator, choque in CENARIOS_PADRAO.items():
            impacto_total = 0
            for item in carteira:
                if fator.lower() in item['classe'].lower():
                    impacto = choque * (item['%PL'] / 100)
                    impacto_total += impacto
            
            res_estresse.append({
                "Fator de Risco": fator,
                "Descrição": DESC_CENARIO[fator],
                "Choque": f"{choque:+.1%}",
                "Impacto (% PL)": f"{impacto_total*100:+.2f}%",
                "Impacto (R$)": f"R$ {impacto_total * pl:+,.0f}"
            })
        
        df_estresse = pd.DataFrame(res_estresse)
        st.dataframe(df_estresse, use_container_width=True)

        # =============== RESPOSTAS CVM/B3 ===============
        
        st.markdown("""
        <div class="section-card">
            <div class="section-title">🏛️ Relatório de Compliance CVM/B3</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Calcular VaR para 21 dias e 95% (padrão CVM)
        var21_total = 0.0
        for item in carteira:
            vol_d = item["vol_anual"] / np.sqrt(252)
            var21_total += pl * (item["%PL"] / 100) * (1.65 * vol_d * np.sqrt(21))
        var21_pct = (var21_total / pl) if pl > 0 else 0.0

        # Obter impactos de estresse
        get_impacto = lambda fator: next((float(r["Impacto (% PL)"].rstrip('%')) for r in res_estresse if r["Fator de Risco"] == fator), 0.0)
        pior_stress = min([float(r["Impacto (% PL)"].rstrip('%')) for r in res_estresse], default=0.0)

        df_respostas_cvm = pd.DataFrame({
            "Pergunta": [
                "Qual é o VAR (Valor de risco) de um dia como percentual do PL calculado para 21 dias úteis e 95% de confiança?",
                "Qual classe de modelos foi utilizada para o cálculo do VAR reportado na questão anterior?",
                "Considerando os cenários de estresse definidos pela BM&FBOVESPA para o fator primitivo de risco (FPR) IBOVESPA que gere o pior resultado para o fundo, indique o cenário utilizado.",
                "Considerando os cenários de estresse definidos pela BM&FBOVESPA para o fator primitivo de risco (FPR) Juros-Pré que gere o pior resultado para o fundo, indique o cenário utilizado.",
                "Considerando os cenários de estresse definidos pela BM&FBOVESPA para o fator primitivo de risco (FPR) Cupom Cambial que gere o pior resultado para o fundo, indique o cenário utilizado.",
                "Considerando os cenários de estresse definidos pela BM&FBOVESPA para o fator primitivo de risco (FPR) Dólar que gere o pior resultado para o fundo, indique o cenário utilizado.",
                "Considerando os cenários de estresse definidos pela BM&FBOVESPA para o fator primitivo de risco (FPR) Outros que gere o pior resultado para o fundo, indique o cenário utilizado.",
                "Qual a variação diária percentual esperada para o valor da cota?",
                "Qual a variação diária percentual esperada para o valor da cota do fundo no pior cenário de estresse definido pelo seu administrador?",
                "Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% na taxa anual de juros (pré)?",
                "Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% na taxa de câmbio (US$/Real)?",
                "Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% no preço das ações (IBOVESPA)?"
            ],
            "Resposta": [
                f"{var21_pct*100:.4f}%",
                "Paramétrico - Delta Normal (sem correlação)",
                DESC_CENARIO.get("Ibovespa", "—"),
                DESC_CENARIO.get("Juros-Pré", "—"),
                DESC_CENARIO.get("Cupom Cambial", "—"),
                DESC_CENARIO.get("Dólar", "—"),
                DESC_CENARIO.get("Outros", "—"),
                f"{df_var['VaR_%'].mean():.4f}%",
                f"{pior_stress:.4f}%",
                f"{get_impacto('Juros-Pré'):.4f}%",
                f"{get_impacto('Dólar'):.4f}%",
                f"{get_impacto('Ibovespa'):.4f}%"
            ]
        })
        
        st.dataframe(df_respostas_cvm, use_container_width=True, height=400)

        # =============== DOWNLOADS ===============
        
        st.markdown("""
        <div class="section-card">
            <div class="section-title">📥 Downloads e Relatórios</div>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Excel completo
            excel_output = BytesIO()
            with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
                # Metadados
                df_meta = pd.DataFrame({
                    "Campo": ["CNPJ", "Fundo", "Data", "PL (R$)", "Confiança", "Horizonte", "Método"],
                    "Valor": [cnpj, nome_fundo, data_ref.strftime("%d/%m/%Y"), f"R$ {pl:,.2f}", 
                             conf_label, f"{horizonte_dias} dias", "Paramétrico Delta-Normal"]
                })
                df_meta.to_excel(writer, sheet_name='Metadados', index=False)
                
                # Resultados
                df_var.to_excel(writer, sheet_name='VaR_por_Classe', index=False)
                df_estresse.to_excel(writer, sheet_name='Cenarios_Estresse', index=False)
                df_respostas_cvm.to_excel(writer, sheet_name='Respostas_CVM_B3', index=False)
            
            excel_output.seek(0)
            
            st.download_button(
                "📊 Relatório Completo",
                data=excel_output,
                file_name=f"relatorio_var_{nome_fundo.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            # Apenas respostas CVM
            excel_cvm = BytesIO()
            df_respostas_cvm.to_excel(excel_cvm, index=False, engine='openpyxl')
            excel_cvm.seek(0)
            
            st.download_button(
                "🏛️ Respostas CVM/B3",
                data=excel_cvm,
                file_name=f"respostas_cvm_{nome_fundo.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col3:
            # Template preenchido
            template_uploaded = st.file_uploader("📋 Upload Template B3/CVM", type=["xlsx"], help="Faça upload do template oficial para preenchimento automático")
            
            if template_uploaded is not None:
                try:
                    output_template = BytesIO()
                    wb = openpyxl.load_workbook(template_uploaded)
                    ws = wb.active
                    
                    # Preencher template automaticamente
                    for col in range(3, ws.max_column + 1):
                        pergunta_template = ws.cell(row=3, column=col).value
                        if pergunta_template:
                            pergunta_text = str(pergunta_template).strip()
                            for _, row in df_respostas_cvm.iterrows():
                                if row["Pergunta"].strip()[:50] in pergunta_text[:50]:
                                    ws.cell(row=6, column=col).value = row["Resposta"]
                                    break
                    
                    wb.save(output_template)
                    output_template.seek(0)
                    
                    st.download_button(
                        "📄 Template Preenchido",
                        data=output_template,
                        file_name=f"template_preenchido_{nome_fundo.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except Exception as e:
                    st.error(f"❌ Erro ao processar template: {str(e)}")
                    st.info("💡 Verifique se o arquivo está no formato correto")

# =============== FOOTER ===============
st.markdown("""
<div class="footer">
    <p>Desenvolvido com ❤️ por <strong>Finhealth</strong></p>
    <p>Análise de risco profissional • Compliance CVM/B3 • Relatórios automatizados</p>
</div>
""", unsafe_allow_html=True)
