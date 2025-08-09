import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
import datetime
import plotly.express as px

# ===================== CONFIG =====================
st.set_page_config(page_title="Finhealth • VaR Calculator", page_icon="📊", layout="wide")

# ===================== CSS =====================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
:root {
  --primary:#0066ff; --primary-dark:#0052d4;
  --success:#10b981; --warning:#f59e0b; --error:#ef4444;
  --bg-main:#fafbfc; --bg-card:#ffffff;
  --text-primary:#111827; --text-secondary:#6b7280; --text-muted:#9ca3af;
  --border:#e5e7eb; --border-light:#f3f4f6;
  --shadow:0 1px 3px rgba(0,0,0,.1), 0 1px 2px rgba(0,0,0,.06);
  --shadow-lg:0 10px 15px -3px rgba(0,0,0,.1), 0 4px 6px -2px rgba(0,0,0,.05);
  --gradient:linear-gradient(135deg,#667eea 0%,#764ba2 100%);
  --gradient-success:linear-gradient(135deg,#10b981 0%,#059669 100%);
}
*{font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif;}
html, body, [data-testid="stAppViewContainer"]{background:var(--bg-main); color:var(--text-primary);}
.main-header{background:var(--gradient); color:#fff; padding:2rem 2.5rem; border-radius:20px; margin-bottom:2rem; box-shadow:var(--shadow-lg); text-align:center;}
.main-header h1{font-size:2.4rem; font-weight:700; margin:0 0 .5rem 0; text-shadow:0 2px 4px rgba(0,0,0,.1);}
.main-header .subtitle{font-size:1.05rem; opacity:.95; font-weight:400;}
.section-card{background:var(--bg-card); border:1px solid var(--border); border-radius:16px; padding:1.25rem; margin-bottom:1rem; box-shadow:var(--shadow); transition:all .3s ease; animation:slideIn .3s ease-out;}
.section-card:hover{box-shadow:var(--shadow-lg); transform:translateY(-2px);}
.section-title{font-size:1.1rem; font-weight:600; color:var(--text-primary); margin-bottom:.75rem; padding-bottom:.5rem; border-bottom:2px solid var(--border-light); display:flex; align-items:center; gap:.5rem;}
.kpi-container{display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:1rem; margin-bottom:1.25rem;}
.kpi-card{background:var(--bg-card); border:1px solid var(--border); border-radius:12px; padding:1.25rem; text-align:center; box-shadow:var(--shadow); position:relative; overflow:hidden;}
.kpi-card::before{content:''; position:absolute; top:0; left:0; right:0; height:4px; background:var(--gradient);}
.kpi-value{font-size:1.8rem; font-weight:700; color:var(--primary); margin-bottom:.25rem;}
.kpi-label{font-size:.85rem; color:var(--text-secondary); font-weight:600; text-transform:uppercase; letter-spacing:.4px;}
.kpi-subtitle{font-size:.75rem; color:var(--text-muted); margin-top:.25rem;}
.status-badge{display:inline-flex; align-items:center; gap:.5rem; padding:.5rem 1rem; border-radius:8px; font-size:.9rem; font-weight:500; margin:.25rem 0 .75rem;}
.status-success{background:rgba(16,185,129,.1); color:var(--success); border:1px solid rgba(16,185,129,.2);}
.status-warning{background:rgba(245,158,11,.1); color:var(--warning); border:1px solid rgba(245,158,11,.2);}
.status-error{background:rgba(239,68,68,.1); color:var(--error); border:1px solid rgba(239,68,68,.2);}
.progress-container{background:var(--border-light); border-radius:8px; height:8px; overflow:hidden; margin:.5rem 0 1rem;}
.progress-bar{height:100%; background:var(--gradient-success); border-radius:8px; transition:width .3s ease;}
.stButton > button{background:var(--gradient)!important; color:#fff!important; border:none!important; border-radius:10px!important; padding:.7rem 1.6rem!important; font-weight:700!important; font-size:1rem!important; transition:all .2s ease!important; box-shadow:var(--shadow)!important;}
.stButton > button:hover{transform:translateY(-2px)!important; box-shadow:var(--shadow-lg)!important;}
.stButton > button:disabled{background:var(--text-muted)!important;}
[data-testid="stSidebar"]{background:var(--bg-card)!important; border-right:1px solid var(--border)!important;}
.streamlit-expanderHeader{background:var(--border-light)!important; border-radius:8px!important; margin-bottom:.5rem!important;}
.js-plotly-plot{border-radius:12px; overflow:hidden; box-shadow:var(--shadow);}
.footer{text-align:center; padding:2rem; color:var(--text-muted); font-size:.9rem; margin-top:2rem; border-top:1px solid var(--border);}
@media (max-width:768px){.main-header{padding:1.5rem}.main-header h1{font-size:2rem}.section-card{padding:1rem}}
@keyframes slideIn{from{opacity:0; transform:translateY(16px)} to{opacity:1; transform:translateY(0)}}
#MainMenu{visibility:hidden} footer{visibility:hidden} header{visibility:hidden} .stDeployButton{display:none}
</style>
""", unsafe_allow_html=True)

# ===================== HEADER =====================
st.markdown("""
<div class="main-header">
  <h1>📊 Finhealth VaR Calculator</h1>
  <div class="subtitle">Análise de Risco Paramétrica por Classe de Ativo • Compliance CVM/B3</div>
</div>
""", unsafe_allow_html=True)

# ===================== CONSTANTES =====================
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

# Classe -> Fator de Risco (para estresse)
FATOR_MAP = {
    "Ações (Ibovespa)": "Ibovespa",
    "Juros-Pré": "Juros-Pré",
    "Câmbio (Dólar)": "Dólar",
    "Cupom Cambial": "Cupom Cambial",
    "Crédito Privado": "Outros",
    "Multimercado": "Outros",
    "Outros": "Outros"
}

# ===================== HELPERS =====================
def z_value(conf_label: str) -> float:
    return 1.644854 if conf_label == "95%" else 2.326347

def brl(x: float, casas: int = 0) -> str:
    s = f"{x:,.{casas}f}"
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")

def pct(x: float, casas: int = 2) -> str:
    return f"{x*100:.{casas}f}%"

def montar_correlacao(classes):
    n = len(classes)
    base = np.full((n, n), 0.20, dtype=float)
    np.fill_diagonal(base, 1.0)
    return pd.DataFrame(base, index=classes, columns=classes)

def var_portfolio(pl, pesos, sigma_d, h, z, corr=None):
    """
    Retorna (var_total_pct, var_total_R$ , sigma_port_d).
    - pesos: np.array shape (n,)
    - sigma_d: vols diárias (n,)
    - corr: matriz de correlação (n,n) ou None (ρ=0)
    """
    w = np.array(pesos)
    s = np.array(sigma_d)
    if corr is None:
        sigma_port_d = np.sqrt(np.sum((w * s) ** 2))
    else:
        D = np.diag(s)
        Sigma = D @ corr @ D
        sigma2 = float(w @ Sigma @ w)
        sigma_port_d = np.sqrt(max(sigma2, 0.0))
    var_total_pct = z * sigma_port_d * np.sqrt(h)  # fração do PL
    return var_total_pct, pl * var_total_pct, sigma_port_d

def impacto_por_fator(fator, carteira_rows, choque):
    """
    Impacto em % do PL para um choque no fator.
    Usa: impacto = choque * sensibilidade * peso
    """
    impacto = 0.0
    for it in carteira_rows:
        classe = it["classe"]
        if FATOR_MAP.get(classe) == fator:
            impacto += choque * it.get("sens", 1.0) * (it["%PL"] / 100.0)
    return impacto  # fração do PL

# ===================== ESTADO =====================
if "rodar" not in st.session_state:
    st.session_state.rodar = False
if "corr_df" not in st.session_state:
    st.session_state.corr_df = None

# ===================== LAYOUT =====================
left, right = st.columns([1.05, 2.0])

with left:
    with st.form("entrada"):
        st.markdown('<div class="section-card"><div class="section-title">🏢 Dados do Fundo</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            cnpj = st.text_input("CNPJ *", placeholder="00.000.000/0001-00")
            data_ref = st.date_input("Data de Referência *", value=datetime.date.today())
        with col2:
            nome_fundo = st.text_input("Nome do Fundo *", placeholder="Fundo XPTO")
            pl = st.number_input("Patrimônio Líquido (R$) *", min_value=0.0, value=1_000_000.0, step=1_000.0, format="%.2f")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="section-card"><div class="section-title">⚙️ Parâmetros</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            horizonte_dias = st.selectbox("Horizonte", [1, 10, 21], index=2,
                help="Período em dias úteis para o VaR (ex.: 21 ≈ 1 mês de pregões).")
        with col2:
            conf_label = st.selectbox("Confiança", ["95%", "99%"], index=0,
                help="Nível de confiança para o VaR (z crítico 95%≈1,645 | 99%≈2,326).")
        metodologia = st.selectbox(
            "Metodologia",
            ["Paramétrico Delta-Normal (ρ=0)", "Paramétrico Delta-Normal (com correlação)"],
            index=0,
            help="Delta-Normal assume retornos ~Normais. ρ=0 soma em quadratura (sem correlação). "
                 "Com correlação usa matriz Corr (editável) e covariância Σ = D·Corr·D."
        )
        usar_corr = (metodologia == "Paramétrico Delta-Normal (com correlação)")

        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="section-card"><div class="section-title">📊 Alocação e Volatilidades</div>', unsafe_allow_html=True)
        carteira, soma = [], 0.0

        st.caption("Informe a alocação (% do PL), a vol anual e (opcional) a sensibilidade ao choque do fator (β/Δ/DV01 normalizado).")

        for classe, vol in VOL_PADRAO.items():
            c1, c2, c3 = st.columns([1.2, .8, .8])
            with c1:
                perc = st.number_input(f"{classe} (%)", 0.0, 100.0, 0.0, 0.5, key=f"p_{classe}", help=f"Vol padrão sugerida: {vol:.0%}")
            with c2:
                vol_a = st.number_input("Vol Anual", 0.0, 2.0, float(vol), 0.01, format="%.2f", key=f"v_{classe}",
                                        help="Volatilidade anual (ex.: 0,25 = 25% a.a.)")
            with c3:
                sens = st.number_input("Sensibilidade", -10.0, 10.0, 1.0, 0.1, key=f"s_{classe}",
                                       help="Elasticidade ao fator (β/Δ/DV01 normalizado). Use 1.0 se não souber.")

            if perc > 0:
                carteira.append({"classe": classe, "%PL": perc, "vol_anual": float(vol_a), "sens": float(sens)})
                soma += perc

        completar_caixa = st.checkbox("Completar com Caixa (σ≈0) quando soma < 100%", value=True)
        normalizar_quando_maior = st.checkbox("Bloquear/alertar quando soma > 100% (não normalizar)", value=True)

        # Barra de progresso
        progress_pct = min(soma / 100.0, 1.0)
        st.markdown(f"""
        <div class="progress-container"><div class="progress-bar" style="width: {progress_pct*100:.1f}%"></div></div>
        """, unsafe_allow_html=True)

        # Status
        if soma == 100:
            st.markdown('<div class="status-badge status-success">✅ Alocação perfeita: 100%</div>', unsafe_allow_html=True)
        elif soma > 100:
            st.markdown(f'<div class="status-badge status-error">❌ Alocação excede: {soma:.1f}%</div>', unsafe_allow_html=True)
        elif soma > 0:
            st.markdown(f'<div class="status-badge status-warning">⚠️ Alocação parcial: {soma:.1f}%</div>', unsafe_allow_html=True)

        # Correlação (se aplicável)
        classes_ativas = [it["classe"] for it in carteira]
        if usar_corr and classes_ativas:
            if st.session_state.corr_df is None or list(st.session_state.corr_df.index) != classes_ativas:
                st.session_state.corr_df = montar_correlacao(classes_ativas)
            st.markdown('<div class="section-card"><div class="section-title">🔗 Matriz de Correlação</div>', unsafe_allow_html=True)
            st.caption("Edite a matriz (1 na diagonal). A simetria é ajustada automaticamente.")
            corr_edit = st.data_editor(
                st.session_state.corr_df.round(2),
                num_rows="fixed",
                use_container_width=True
            )
            # Força simetria e diagonal = 1
            corr_np = corr_edit.to_numpy(dtype=float)
            corr_sym = (corr_np + corr_np.T) / 2.0
            np.fill_diagonal(corr_sym, 1.0)
            st.session_state.corr_df = pd.DataFrame(corr_sym, index=classes_ativas, columns=classes_ativas)
            st.markdown('</div>', unsafe_allow_html=True)

        campos_ok = bool(cnpj.strip() and nome_fundo.strip() and pl > 0 and soma > 0)
        pode_calcular = campos_ok and (soma <= 100 or not normalizar_quando_maior)

        botao = st.form_submit_button("🚀 Calcular VaR & Compliance", disabled=not pode_calcular)
        if botao:
            # Completa com caixa se habilitado
            if soma < 100 and completar_caixa:
                caixa = 100 - soma
                carteira.append({"classe": "Caixa", "%PL": caixa, "vol_anual": 0.0001, "sens": 0.0})
                soma = 100.0
            st.session_state.rodar = True

with right:
    if not st.session_state.rodar:
        st.markdown("""
        <div class="section-card">
            <div class="section-title">ℹ️ Instruções</div>
            <p>Para calcular o VaR, preencha:</p>
            <ul>
                <li><strong>CNPJ</strong> e <strong>Nome do Fundo</strong></li>
                <li><strong>Patrimônio Líquido</strong> maior que zero</li>
                <li><strong>Alocação da carteira</strong> somando até 100%</li>
                <li><strong>Volatilidades anuais</strong> por classe (padrões sugeridos)</li>
                <li>Se desejar, edite a <strong>matriz de correlação</strong></li>
            </ul>
            <p>O sistema calculará automaticamente:</p>
            <ul>
                <li>VaR do portfólio (soma em quadratura ou com correlação)</li>
                <li>VaR por classe (isolado)</li>
                <li>Cenários de estresse com sensibilidade</li>
                <li>Respostas para CVM/B3</li>
                <li>Relatórios em Excel</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    elif st.session_state.rodar:
        # ===================== CÁLCULOS =====================
        z = z_value(conf_label)
        h = int(horizonte_dias)

        # Arrays
        pesos = np.array([it["%PL"]/100 for it in carteira], dtype=float)
        sigma_d = np.array([it["vol_anual"]/np.sqrt(252) for it in carteira], dtype=float)

        corr = None
        if usar_corr:
            # Remonta matriz para as classes da carteira (inclusive Caixa, se houver)
            idx = [it["classe"] for it in carteira]
            # Se corr_df não cobre alguma classe (ex.: Caixa), cria base e injeta
            if st.session_state.corr_df is None or list(st.session_state.corr_df.index) != [c for c in classes_ativas]:
                base = montar_correlacao(idx)
                st.session_state.corr_df = base
            else:
                # expandir para incluir "Caixa" se não existir
                corr_curr = st.session_state.corr_df
                if set(idx) != set(corr_curr.index):
                    base = montar_correlacao(idx)
                    for i in base.index:
                        for j in base.columns:
                            if i in corr_curr.index and j in corr_curr.columns:
                                base.loc[i, j] = float(corr_curr.loc[i, j])
                    st.session_state.corr_df = base
            corr = st.session_state.corr_df.to_numpy(dtype=float)

        var_total_pct, var_total_rs, sigma_port_d = var_portfolio(pl, pesos, sigma_d, h, z, corr=corr)

        # VaR por classe (isolado)
        var_classes_rs = []
        var_classes_pct = []
        for w_i, sdi in zip(pesos, sigma_d):
            v_pct = z * sdi * np.sqrt(h) * w_i   # fração do PL
            var_classes_pct.append(v_pct)
            var_classes_rs.append(pl * v_pct)

        # DataFrame principal
        df_var = pd.DataFrame({
            "classe": [it["classe"] for it in carteira],
            "%PL": [it["%PL"] for it in carteira],
            "vol_anual": [it["vol_anual"] for it in carteira],
            "VaR_%": [v*100 for v in var_classes_pct],
            "VaR_R$": var_classes_rs
        })

        # ===================== KPIs =====================
        st.markdown('<div class="section-card"><div class="section-title">📌 Indicadores</div>', unsafe_allow_html=True)
        k1 = f"{var_total_pct*100:.2f}%"
        k2 = brl(var_total_rs, 0)
        k3 = len(df_var)
        k4 = sum([it["%PL"] for it in carteira])
        k5 = f"{sigma_port_d*100:.2f}%"
        st.markdown(f"""
        <div class="kpi-container">
          <div class="kpi-card"><div class="kpi-value">{k1}</div><div class="kpi-label">VaR ({conf_label} / {h}d)</div><div class="kpi-subtitle">Portfólio</div></div>
          <div class="kpi-card"><div class="kpi-value">{k2}</div><div class="kpi-label">VaR em Reais</div><div class="kpi-subtitle">Perda potencial</div></div>
          <div class="kpi-card"><div class="kpi-value">{k5}</div><div class="kpi-label">σ diário da cota</div><div class="kpi-subtitle">Volatilidade</div></div>
          <div class="kpi-card"><div class="kpi-value">{k3}</div><div class="kpi-label">Classes Ativas</div><div class="kpi-subtitle">em uso</div></div>
          <div class="kpi-card"><div class="kpi-value">{k4:.1f}%</div><div class="kpi-label">Alocação Total</div><div class="kpi-subtitle">do PL</div></div>
        </div>
        """, unsafe_allow_html=True)

        # ===================== TABELA VAR =====================
        st.markdown('<div class="section-card"><div class="section-title">📈 VaR por Classe de Ativo (isolado)</div>', unsafe_allow_html=True)
        df_display = df_var.copy()
        df_display['%PL'] = df_display['%PL'].map(lambda x: f"{x:.1f}%")
        df_display['vol_anual'] = df_display['vol_anual'].map(lambda x: f"{x:.2%}")
        df_display['VaR_%'] = df_display['VaR_%'].map(lambda x: f"{x:.2f}%")
        df_display['VaR_R$'] = df_var['VaR_R$'].map(lambda x: brl(x, 0))
        df_display.columns = ['Classe de Ativo', 'Alocação', 'Volatilidade Anual', 'VaR (%)', 'VaR (R$)']
        st.dataframe(df_display, use_container_width=True)

        # ===================== GRÁFICOS =====================
        colg1, colg2 = st.columns(2)
        with colg1:
            fig_pie = px.pie(df_var, values="%PL", names="classe", title="Distribuição da Carteira")
            fig_pie.update_layout(font=dict(family="Inter, sans-serif"), title_font_size=16, legend=dict(orientation="v", y=0.5))
            st.plotly_chart(fig_pie, use_container_width=True)
        with colg2:
            fig_bar = px.bar(df_var, x="classe", y="VaR_R$", title="VaR por Classe (R$)", color="VaR_R$", color_continuous_scale="Blues")
            fig_bar.update_layout(font=dict(family="Inter, sans-serif"), title_font_size=16, xaxis_title="", yaxis_title="VaR (R$)")
            fig_bar.update_xaxes(tickangle=45)
            st.plotly_chart(fig_bar, use_container_width=True)

        # ===================== ESTRESSE =====================
        st.markdown('<div class="section-card"><div class="section-title">⚠️ Cenários de Estresse</div>', unsafe_allow_html=True)
        resultados_estresse = []
        for fator, choque in CENARIOS_PADRAO.items():
            impacto_pct = impacto_por_fator(fator, carteira, choque)      # fração do PL
            resultados_estresse.append({
                "Fator de Risco": fator,
                "Descrição": DESC_CENARIO[fator],
                "Choque": choque,
                "Impacto_%PL": impacto_pct,
                "Impacto_R$": impacto_pct * pl
            })
        df_estresse = pd.DataFrame(resultados_estresse)

        df_estresse_view = df_estresse.copy()
        df_estresse_view["Choque"] = df_estresse_view["Choque"].map(lambda x: f"{x:+.1%}")
        df_estresse_view["Impacto (% PL)"] = df_estresse_view["Impacto_%PL"].map(lambda x: f"{x*100:+.2f}%")
        df_estresse_view["Impacto (R$)"] = df_estresse_view["Impacto_R$"].map(lambda x: brl(x, 0))
        df_estresse_view = df_estresse_view[["Fator de Risco","Descrição","Choque","Impacto (% PL)","Impacto (R$)"]]
        st.dataframe(df_estresse_view, use_container_width=True)

        # ===================== COMPLIANCE CVM/B3 =====================
        st.markdown('<div class="section-card"><div class="section-title">🏛️ Relatório de Compliance CVM/B3</div>', unsafe_allow_html=True)

        z95 = 1.644854
        var21_pct = z95 * sigma_port_d * np.sqrt(21) * 100.0  # em %
        # Variação diária esperada (interpretação: σ diário, não VaR)
        variacao_diaria_pct = sigma_port_d * 100.0

        # Pior cenário dos definidos
        pior_stress_pct = 0.0
        if not df_estresse.empty:
            pior_stress_pct = float(df_estresse["Impacto_%PL"].min() * 100.0)

        # Sensibilidades unitárias de -1% em fatores específicos:
        def impacto_unit(fator, unit=-0.01):
            return impacto_por_fator(fator, carteira, unit) * 100.0  # em %

        resp_rows = [
            ("Qual é o VAR (Valor de risco) de um dia como percentual do PL calculado para 21 dias úteis e 95% de confiança?",
             f"{var21_pct:.4f}%"),
            ("Qual classe de modelos foi utilizada para o cálculo do VAR reportado na questão anterior?",
             "Paramétrico - Delta Normal" + (" (com correlação)" if usar_corr else " (ρ=0, sem correlação)")),
            ("Considerando os cenários de estresse definidos pela BM&FBOVESPA para o fator primitivo de risco (FPR) IBOVESPA que gere o pior resultado para o fundo, indique o cenário utilizado.",
             DESC_CENARIO.get("Ibovespa", "—")),
            ("Considerando os cenários de estresse definidos pela BM&FBOVESPA para o fator primitivo de risco (FPR) Juros-Pré que gere o pior resultado para o fundo, indique o cenário utilizado.",
             DESC_CENARIO.get("Juros-Pré", "—")),
            ("Considerando os cenários de estresse definidos pela BM&FBOVESPA para o fator primitivo de risco (FPR) Cupom Cambial que gere o pior resultado para o fundo, indique o cenário utilizado.",
             DESC_CENARIO.get("Cupom Cambial", "—")),
            ("Considerando os cenários de estresse definidos pela BM&FBOVESPA para o fator primitivo de risco (FPR) Dólar que gere o pior resultado para o fundo, indique o cenário utilizado.",
             DESC_CENARIO.get("Dólar", "—")),
            ("Considerando os cenários de estresse definidos pela BM&FBOVESPA para o fator primitivo de risco (FPR) Outros que gere o pior resultado para o fundo, indique o cenário utilizado.",
             DESC_CENARIO.get("Outros", "—")),
            ("Qual a variação diária percentual esperada para o valor da cota?",
             f"{variacao_diaria_pct:.4f}%"),
            ("Qual a variação diária percentual esperada para o valor da cota do fundo no pior cenário de estresse definido pelo seu administrador?",
             f"{pior_stress_pct:.4f}%"),
            ("Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% na taxa anual de juros (pré)?",
             f"{impacto_unit('Juros-Pré', -0.01):.4f}%"),
            ("Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% na taxa de câmbio (US$/Real)?",
             f"{impacto_unit('Dólar', -0.01):.4f}%"),
            ("Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% no preço das ações (IBOVESPA)?",
             f"{impacto_unit('Ibovespa', -0.01):.4f}%"),
        ]
        df_respostas_cvm = pd.DataFrame(resp_rows, columns=["Pergunta","Resposta"])
        st.dataframe(df_respostas_cvm, use_container_width=True, height=420)

        # ===================== DOWNLOADS =====================
        st.markdown('<div class="section-card"><div class="section-title">📥 Downloads e Relatórios</div>', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)

        with col1:
            excel_output = BytesIO()
            with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
                # Metadados
                df_meta = pd.DataFrame({
                    "Campo": ["CNPJ", "Fundo", "Data", "PL (R$)", "Confiança", "Horizonte", "Método"],
                    "Valor": [cnpj, nome_fundo, data_ref.strftime("%d/%m/%Y"), brl(pl, 2),
                              conf_label, f"{h} dias", "Paramétrico Delta-Normal " + ("(com correlação)" if usar_corr else "(ρ=0)")]
                })
                df_meta.to_excel(writer, sheet_name='Metadados', index=False)
                # Resultados VaR
                df_var.to_excel(writer, sheet_name='VaR_por_Classe', index=False)
                # Estresse
                df_estresse.to_excel(writer, sheet_name='Cenarios_Estresse_raw', index=False)
                df_estresse_view.to_excel(writer, sheet_name='Cenarios_Estresse', index=False)
                # Respostas CVM
                df_respostas_cvm.to_excel(writer, sheet_name='Respostas_CVM_B3', index=False)
                # Sumário
                pd.DataFrame({
                    "Métrica":["VaR_port_%","VaR_port_R$","Sigma_diario_%"],
                    "Valor":[f"{var_total_pct*100:.4f}%", brl(var_total_rs, 2), f"{sigma_port_d*100:.4f}%"]
                }).to_excel(writer, sheet_name='Sumario', index=False)
            excel_output.seek(0)
            st.download_button("📊 Relatório Completo (Excel)", data=excel_output,
                               file_name=f"relatorio_var_{nome_fundo.replace(' ', '_')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with col2:
            excel_cvm = BytesIO()
            df_respostas_cvm.to_excel(excel_cvm, index=False, engine='openpyxl')
            excel_cvm.seek(0)
            st.download_button("🏛️ Respostas CVM/B3 (Excel)", data=excel_cvm,
                               file_name=f"respostas_cvm_{nome_fundo.replace(' ', '_')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with col3:
            template_uploaded = st.file_uploader("📋 Upload Template B3/CVM", type=["xlsx"],
                                                 help="Suba o template oficial para preenchimento automático (matching por trechos da pergunta).")
            if template_uploaded is not None:
                try:
                    output_template = BytesIO()
                    wb = openpyxl.load_workbook(template_uploaded)
                    ws = wb.active
                    # Procura perguntas na linha 3 (ajuste se necessário)
                    perguntas_template = {}
                    for col in range(3, ws.max_column + 1):
                        perguntas_template[col] = str(ws.cell(row=3, column=col).value or "").strip().lower()

                    # Preenche linha 6 com respostas mapeadas
                    for _, row in df_respostas_cvm.iterrows():
                        p = row["Pergunta"].strip().lower()
                        for col, ptemp in perguntas_template.items():
                            # matching simples por substring
                            if p[:50] in ptemp or ptemp[:50] in p:
                                ws.cell(row=6, column=col).value = row["Resposta"]
                                break

                    wb.save(output_template)
                    output_template.seek(0)
                    st.download_button("📄 Template Preenchido", data=output_template,
                                       file_name=f"template_preenchido_{nome_fundo.replace(' ', '_')}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"❌ Erro ao processar template: {str(e)}")
                    st.info("💡 Verifique se o arquivo e as células (linha 3 perguntas / linha 6 respostas) estão conforme esperado.")

# ===================== FOOTER =====================
st.markdown("""
<div class="footer">
  <p>Feito com ❤️ <strong>Finhealth</strong></p>
  <p>Análise de risco profissional • Compliance CVM/B3 • Relatórios automatizados</p>
</div>
""", unsafe_allow_html=True)
