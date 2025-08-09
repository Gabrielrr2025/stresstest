import streamlit as st
import pandas as pd
import numpy as np
import datetime
from io import BytesIO
import openpyxl
import plotly.express as px

# ===================== CONFIG =====================
st.set_page_config(page_title="Finhealth • VaR", page_icon="📊", layout="wide")

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
CENARIOS_PADRAO = {"Ibovespa": -0.15, "Juros-Pré": 0.02, "Cupom Cambial": -0.01, "Dólar": -0.05, "Outros": -0.03}
DESC_CENARIO = {
    "Ibovespa": "Queda de 15% no IBOVESPA",
    "Juros-Pré": "Alta de 200 bps na taxa de juros",
    "Cupom Cambial": "Queda de 1% no cupom cambial",
    "Dólar": "Queda de 5% no dólar",
    "Outros": "Queda de 3% em outros ativos"
}
# Mapeia classe -> fator para estresse
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
def z_value(level: str) -> float:
    return 1.644854 if level == "95%" else 2.326347

def brl(x: float, casas: int = 0) -> str:
    s = f"{x:,.{casas}f}"
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")

def var_portfolio(pl, pesos, sigma_d, h, z, corr=None):
    w = np.array(pesos, dtype=float)
    s = np.array(sigma_d, dtype=float)
    if corr is None:
        sigma_port_d = np.sqrt(np.sum((w * s) ** 2))
    else:
        D = np.diag(s); Sigma = D @ corr @ D
        sigma2 = float(w @ Sigma @ w)
        sigma_port_d = np.sqrt(max(sigma2, 0.0))
    var_total_pct = z * sigma_port_d * np.sqrt(h)
    return var_total_pct, var_total_pct * pl, sigma_port_d

def montar_correlacao(classes):
    n = len(classes)
    base = np.full((n, n), 0.20, dtype=float)
    np.fill_diagonal(base, 1.0)
    return pd.DataFrame(base, index=classes, columns=classes)

def impacto_por_fator(fator, carteira_rows, choque):
    impacto = 0.0
    for it in carteira_rows:
        if FATOR_MAP.get(it["classe"]) == fator:
            impacto += choque * it.get("sens", 1.0) * (it["%PL"]/100.0)
    return impacto  # fração do PL

def label(texto: str, missing: bool=False):
    st.markdown(f'<div class="lbl{" missing" if missing else ""}">{texto}</div>', unsafe_allow_html=True)

# ===================== ESTADO =====================
if "rodar" not in st.session_state: st.session_state.rodar = False
if "corr_df" not in st.session_state: st.session_state.corr_df = None
if "tentou" not in st.session_state: st.session_state.tentou = False

# ===================== SIDEBAR (Parâmetros + Tema) =====================
with st.sidebar:
    st.header("⚙️ Parâmetros")

    tema = st.selectbox(
        "Tema",
        ["Claro", "Escuro"],
        index=0,
        help="Altera a aparência do site."
    )

    horizonte_dias = st.selectbox(
        "Horizonte (dias úteis)",
        [1, 10, 21], index=2,
        help="Período considerado para o cálculo do VaR."
    )
    conf_label = st.selectbox(
        "Confiança",
        ["95%", "99%"], index=0,
        help="Probabilidade associada ao nível de perda estimada."
    )
    metodologia = st.selectbox(
        "Metodologia",
        ["Sem correlação (soma em quadratura)", "Com correlação (matriz de correlação)"],
        index=0,
        help="Define se o portfólio considera dependência entre classes de ativos."
    )
    usar_corr = metodologia.startswith("Com correlação")

# ===================== TEMA (CSS dinâmico) =====================
CSS_LIGHT = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
:root{
  --bg:#fafbfc; --card:#ffffff; --text:#111827; --muted:#6b7280; --line:#e5e7eb;
  --primary:#075aff; --ok:#10b981; --warn:#f59e0b; --err:#ef4444;
}
*{font-family:'Inter',system-ui,-apple-system,BlinkMacSystemFont}
[data-testid="stAppViewContainer"]{background:var(--bg)}
.block-container{max-width:1100px; padding-top:1rem}
.card{background:var(--card); border:1px solid var(--line); border-radius:14px; padding:1rem 1.2rem; margin-bottom:1rem}
.h1{font-size:1.6rem; font-weight:700; margin:0 0 .25rem}
.h2{font-size:1.05rem; font-weight:700; color:var(--text); border-bottom:1px solid #f2f3f5; padding-bottom:.35rem; margin-bottom:.7rem}
.kpi{background:var(--card); border:1px solid var(--line); border-radius:12px; padding:1rem; text-align:center}
.kpv{font-size:1.5rem; font-weight:700; color:var(--primary)}
.kpl{font-size:.8rem; text-transform:uppercase; letter-spacing:.4px; color:var(--muted); font-weight:700}
.progress{height:8px; background:#f3f4f6; border-radius:8px; overflow:hidden; margin:.5rem 0 .7rem}
.progress > div{height:100%; background:linear-gradient(90deg,#22c55e,#16a34a)}
.badge{display:inline-block; padding:.35rem .6rem; border-radius:8px; font-weight:600; font-size:.85rem; border:1px solid}
.ok{color:var(--ok); background:rgba(16,185,129,.08); border-color:rgba(16,185,129,.25)}
.warn{color:var(--warn); background:rgba(245,158,11,.08); border-color:rgba(245,158,11,.25)}
.err{color:var(--err); background:rgba(239,68,68,.08); border-color:rgba(239,68,68,.25)}
.lbl{font-weight:600; margin-bottom:4px}
.lbl.missing{color:var(--err)}
.help-err{color:var(--err); font-size:.85rem; margin-top:.25rem}
.js-plotly-plot{border:1px solid var(--line); border-radius:12px}
footer, #MainMenu, header{visibility:hidden}
.footer{color:#6b7280; text-align:center; padding:1.6rem 0 1rem; border-top:1px solid #ececec; margin-top:1.2rem}
</style>
"""
CSS_DARK = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
:root{
  --bg:#0b1020; --card:#11182a; --text:#eef2ff; --muted:#a5b4fc; --line:#1f2a44;
  --primary:#7c9cff; --ok:#34d399; --warn:#f59e0b; --err:#f87171;
}
*{font-family:'Inter',system-ui,-apple-system,BlinkMacSystemFont}
[data-testid="stAppViewContainer"]{background:var(--bg)}
.block-container{max-width:1100px; padding-top:1rem}
.card{background:var(--card); border:1px solid var(--line); border-radius:14px; padding:1rem 1.2rem; margin-bottom:1rem}
.h1{font-size:1.6rem; font-weight:700; color:var(--text); margin:0 0 .25rem}
.h2{font-size:1.05rem; font-weight:700; color:var(--text); border-bottom:1px solid #243352; padding-bottom:.35rem; margin-bottom:.7rem}
.kpi{background:var(--card); border:1px solid var(--line); border-radius:12px; padding:1rem; text-align:center}
.kpv{font-size:1.5rem; font-weight:700; color:var(--primary)}
.kpl{font-size:.8rem; text-transform:uppercase; letter-spacing:.4px; color:var(--muted); font-weight:700}
.progress{height:8px; background:#1a2744; border-radius:8px; overflow:hidden; margin:.5rem 0 .7rem}
.progress > div{height:100%; background:linear-gradient(90deg,#22c55e,#16a34a)}
.badge{display:inline-block; padding:.35rem .6rem; border-radius:8px; font-weight:600; font-size:.85rem; border:1px solid}
.ok{color:var(--ok); background:rgba(52,211,153,.12); border-color:rgba(52,211,153,.25)}
.warn{color:var(--warn); background:rgba(245,158,11,.12); border-color:rgba(245,158,11,.25)}
.err{color:var(--err); background:rgba(248,113,113,.12); border-color:rgba(248,113,113,.25)}
.lbl{font-weight:600; margin-bottom:4px; color:var(--text)}
.lbl.missing{color:var(--err)}
.help-err{color:var(--err); font-size:.85rem; margin-top:.25rem}
.js-plotly-plot{border:1px solid var(--line); border-radius:12px; background:var(--card)}
footer, #MainMenu, header{visibility:hidden}
.footer{color:#9aa6ff; text-align:center; padding:1.6rem 0 1rem; border-top:1px solid #243352; margin-top:1.2rem}
</style>
"""
st.markdown(CSS_DARK if tema == "Escuro" else CSS_LIGHT, unsafe_allow_html=True)
plotly_template = "plotly_dark" if tema == "Escuro" else "plotly_white"

# ===================== CABEÇALHO =====================
st.markdown('<div class="card"><div class="h1">📊 Finhealth VaR</div>'
            '<div style="color:var(--muted)">Risco paramétrico por classe • Relatórios e respostas CVM/B3</div></div>',
            unsafe_allow_html=True)

# ===================== DADOS DO FUNDO + ALOCAÇÃO =====================
with st.form("form_fundo"):
    st.markdown('<div class="card"><div class="h2">🏢 Dados do Fundo</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        label("CNPJ *", missing=(st.session_state.tentou and not st.session_state.get("cnpj_val", "").strip()))
        cnpj = st.text_input("", placeholder="00.000.000/0001-00", label_visibility="collapsed")
        if st.session_state.tentou and not cnpj.strip():
            st.markdown('<div class="help-err">Informe o CNPJ.</div>', unsafe_allow_html=True)

        label("Nome do Fundo *", missing=(st.session_state.tentou and not st.session_state.get("nome_val", "").strip()))
        nome_fundo = st.text_input("", placeholder="Fundo XPTO", label_visibility="collapsed")
        if st.session_state.tentou and not nome_fundo.strip():
            st.markdown('<div class="help-err">Informe o nome do fundo.</div>', unsafe_allow_html=True)

    with c2:
        label("Data de Referência *")
        data_ref = st.date_input("", value=datetime.date.today(), label_visibility="collapsed")

        pl_missing = st.session_state.tentou and (st.session_state.get("pl_val", 0.0) <= 0)
        label("Patrimônio Líquido (R$) *", missing=pl_missing)
        pl = st.number_input("", min_value=0.0, value=1_000_000.0, step=1_000.0, format="%.2f",
                             label_visibility="collapsed")
        if st.session_state.tentou and pl <= 0:
            st.markdown('<div class="help-err">Informe um valor maior que zero.</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card"><div class="h2">📊 Alocação por Classe</div>', unsafe_allow_html=True)
    st.caption("Informe a distribuição por classe, a volatilidade anual sugerida e, se aplicável, a sensibilidade (elasticidade ao fator).")

    carteira, soma = [], 0.0
    faltas_vol = {}

    for classe, vol_sugerida in VOL_PADRAO.items():
        a, b, c = st.columns([1.2, .9, .9])
        with a:
            label(f"{classe} (%)")
            perc = st.number_input("", min_value=0.0, max_value=100.0, value=0.0, step=0.5,
                                   key=f"p_{classe}", label_visibility="collapsed")
        with b:
            label("Volatilidade Anual")
            vol_a = st.number_input("", min_value=0.0, max_value=2.0, value=float(vol_sugerida),
                                    step=0.01, format="%.2f", key=f"v_{classe}", label_visibility="collapsed")
        with c:
            label("Sensibilidade")
            sens = st.number_input("", min_value=-10.0, max_value=10.0, value=1.0, step=0.1,
                                   key=f"s_{classe}", label_visibility="collapsed")

        if perc > 0:
            carteira.append({"classe": classe, "%PL": perc, "vol_anual": float(vol_a), "sens": float(sens)})
            soma += perc
            if st.session_state.tentou and vol_a <= 0:
                faltas_vol[classe] = True
                st.markdown(f'<div class="help-err">Volatilidade obrigatória para "{classe}".</div>', unsafe_allow_html=True)

    # Barra + status
    st.markdown(f'<div class="progress"><div style="width:{min(soma,100):.1f}%"></div></div>', unsafe_allow_html=True)
    if soma == 100:
        st.markdown('<span class="badge ok">✅ Alocação total: 100%</span>', unsafe_allow_html=True)
    elif soma > 100:
        st.markdown(f'<span class="badge err">❌ A soma ultrapassa 100% ({soma:.1f}%).</span>', unsafe_allow_html=True)
    elif soma > 0:
        st.markdown(f'<span class="badge warn">⚠️ A soma está em {soma:.1f}%. Complete até 100%.</span>', unsafe_allow_html=True)

    completar_caixa = st.checkbox("Completar automaticamente com Caixa quando a soma for menor que 100%", value=True)

    submit = st.form_submit_button("🚀 Calcular")
    if submit:
        st.session_state.tentou = True
        missing_msgs = []
        if not cnpj.strip(): missing_msgs.append("CNPJ")
        if not nome_fundo.strip(): missing_msgs.append("Nome do Fundo")
        if pl <= 0: missing_msgs.append("Patrimônio Líquido maior que zero")
        if soma == 0: missing_msgs.append("Informar ao menos uma classe na alocação")
        if soma > 100: missing_msgs.append("Soma da alocação não pode exceder 100%")
        for classe in faltas_vol:
            missing_msgs.append(f'Volatilidade anual para "{classe}"')

        if missing_msgs:
            st.session_state.rodar = False
            st.error("Por favor, corrija os campos destacados:\n- " + "\n- ".join(missing_msgs))
        else:
            if soma < 100 and completar_caixa:
                carteira.append({"classe": "Caixa", "%PL": 100 - soma, "vol_anual": 0.0001, "sens": 0.0})
                soma = 100.0
            st.session_state.rodar = True
            st.session_state.inputs = {"cnpj": cnpj, "nome": nome_fundo, "data": data_ref, "pl": pl, "carteira": carteira}
            st.success("Cálculo concluído. Veja os resultados abaixo.")

# ===================== RESULTADOS =====================
if st.session_state.rodar:
    data = st.session_state.inputs
    pl = data["pl"]
    carteira = data["carteira"]

    # Arrays
    pesos = np.array([it["%PL"]/100 for it in carteira], dtype=float)
    sigma_d = np.array([it["vol_anual"]/np.sqrt(252) for it in carteira], dtype=float)
    classes = [it["classe"] for it in carteira]

    # Correlação (opcional)
    corr = None
    if usar_corr:
        if (st.session_state.corr_df is None) or (list(st.session_state.corr_df.index) != classes):
            st.session_state.corr_df = montar_correlacao(classes)
        with st.expander("🔗 Matriz de correlação (opcional)"):
            st.caption("A matriz deve ser simétrica e ter 1 na diagonal.")
            edit = st.data_editor(st.session_state.corr_df.round(2), num_rows="fixed", use_container_width=True)
            M = edit.to_numpy(float); M = (M + M.T)/2.0; np.fill_diagonal(M, 1.0)
            st.session_state.corr_df = pd.DataFrame(M, index=classes, columns=classes)
        corr = st.session_state.corr_df.to_numpy(float)

    # Cálculo VaR portfólio
    z = z_value(conf_label); h = int(horizonte_dias)
    var_pct, var_rs, sigma_port_d = var_portfolio(pl, pesos, sigma_d, h, z, corr=corr)

    # VaR isolado por classe (exibição)
    var_cls_pct = (z * sigma_d * np.sqrt(h)) * pesos     # fração do PL
    var_cls_rs = var_cls_pct * pl
    df_var = pd.DataFrame({
        "classe": classes,
        "%PL": [it["%PL"] for it in carteira],
        "vol_anual": [it["vol_anual"] for it in carteira],
        "VaR_%": var_cls_pct * 100,
        "VaR_R$": var_cls_rs
    })

    # KPIs
    st.markdown('<div class="card"><div class="h2">📌 Indicadores</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    cols[0].markdown(f'<div class="kpi"><div class="kpv">{var_pct*100:.2f}%</div><div class="kpl">VaR ({conf_label} / {h}d)</div></div>', unsafe_allow_html=True)
    cols[1].markdown(f'<div class="kpi"><div class="kpv">{brl(var_rs,0)}</div><div class="kpl">VaR em Reais</div></div>', unsafe_allow_html=True)
    cols[2].markdown(f'<div class="kpi"><div class="kpv">{sigma_port_d*100:.2f}%</div><div class="kpl">σ diário da cota</div></div>', unsafe_allow_html=True)
    cols[3].markdown(f'<div class="kpi"><div class="kpv">{sum([it["%PL"] for it in carteira]):.1f}%</div><div class="kpl">Alocação total</div></div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Tabela
    st.markdown('<div class="card"><div class="h2">📈 VaR por Classe (isolado)</div>', unsafe_allow_html=True)
    df_show = df_var.copy()
    df_show["%PL"] = df_show["%PL"].map(lambda x: f"{x:.1f}%")
    df_show["vol_anual"] = df_show["vol_anual"].map(lambda x: f"{x:.2%}")
    df_show["VaR_%"] = df_show["VaR_%"].map(lambda x: f"{x:.2f}%")
    df_show["VaR (R$)"] = df_var["VaR_R$"].map(lambda x: brl(x, 0))
    df_show = df_show.drop(columns=["VaR_R$"]).rename(columns={"classe":"Classe de Ativo","vol_anual":"Volatilidade Anual"})
    st.dataframe(df_show, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Gráficos
    g1, g2 = st.columns(2)
    with g1:
        fig = px.pie(df_var, values="%PL", names="classe", title="Distribuição da Carteira", template=plotly_template)
        fig.update_layout(height=360)
        st.plotly_chart(fig, use_container_width=True)
    with g2:
        fig2 = px.bar(df_var, x="classe", y="VaR_R$", title="VaR por Classe (R$)",
                      color="VaR_R$", color_continuous_scale="Blues", template=plotly_template)
        fig2.update_layout(xaxis_title="", yaxis_title="VaR (R$)", height=360)
        fig2.update_xaxes(tickangle=45)
        st.plotly_chart(fig2, use_container_width=True)

    # Estresse
    st.markdown('<div class="card"><div class="h2">⚠️ Cenários de Estresse</div>', unsafe_allow_html=True)
    est_rows = []
    for fator, choque in CENARIOS_PADRAO.items():
        impacto = impacto_por_fator(fator, carteira, choque)   # fração do PL
        est_rows.append({
            "Fator": fator,
            "Descrição": DESC_CENARIO[fator],
            "Choque": f"{choque:+.1%}",
            "Impacto (% PL)": f"{impacto*100:+.2f}%",
            "Impacto (R$)": brl(impacto*pl, 0)
        })
    st.dataframe(pd.DataFrame(est_rows), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ===================== COMPLIANCE CVM/B3 =====================
    st.markdown('<div class="card"><div class="h2">🏛️ Respostas CVM/B3</div>', unsafe_allow_html=True)
    z95 = 1.644854
    var21_pct = z95 * sigma_port_d * np.sqrt(21) * 100.0  # em %
    brutos = [impacto_por_fator(f, carteira, ch) for f, ch in CENARIOS_PADRAO.items()]
    pior_stress_pct = (min(brutos) * 100.0) if brutos else 0.0

    def imp_unit(fator, unit=-0.01):
        return impacto_por_fator(fator, carteira, unit) * 100.0  # em %

    # Principal fator (pondera exposição e sensibilidade)
    excluidos = {"Ibovespa", "Juros-Pré", "Dólar"}
    expos = {}
    for it in carteira:
        fator = FATOR_MAP.get(it["classe"])
        if fator:
            expos[fator] = expos.get(fator, 0.0) + (it["%PL"]/100.0)*abs(it.get("sens", 1.0))
    principal = max(expos, key=expos.get) if expos else None

    if principal in excluidos:
        resp_outros_composta = "Não aplicável (principal fator é juros, câmbio ou bolsa)"
        resp_outros_fator = "—"
        resp_outros_pct = "—"
        explicacao_outros = f"Obs.: Principal fator identificado: {principal}. Como ele já está entre juros, câmbio ou bolsa, as três últimas linhas não se aplicam."
    else:
        if principal:
            var_outros_pct = imp_unit(principal, -0.01)
            resp_outros_composta = f"{var_outros_pct:.4f}% (Fator: {principal})"
            resp_outros_fator = principal
            resp_outros_pct = f"{var_outros_pct:.4f}%"
        else:
            resp_outros_composta = "—"
            resp_outros_fator = "—"
            resp_outros_pct = "—"
        explicacao_outros = "Obs.: As três últimas linhas só se aplicam quando o principal fator não é juros, câmbio nem bolsa."

    df_cvm = pd.DataFrame({
        "Pergunta": [
            "Qual é o VAR (Valor de risco) de um dia como percentual do PL calculado para 21 dias úteis e 95% de confiança?",
            "Qual classe de modelos foi utilizada para o cálculo do VAR reportado na questão anterior?",
            "Considerando os cenários de estresse definidos pela BM&FBOVESPA para o FPR IBOVESPA que gere o pior resultado para o fundo, indique o cenário utilizado.",
            "Considerando os cenários de estresse definidos pela BM&FBOVESPA para o FPR Juros-Pré que gere o pior resultado para o fundo, indique o cenário utilizado.",
            "Considerando os cenários de estresse definidos pela BM&FBOVESPA para o FPR Cupom Cambial que gere o pior resultado para o fundo, indique o cenário utilizado.",
            "Considerando os cenários de estresse definidos pela BM&FBOVESPA para o FPR Dólar que gere o pior resultado para o fundo, indique o cenário utilizado.",
            "Considerando os cenários de estresse definidos pela BM&FBOVESPA para o FPR Outros que gere o pior resultado para o fundo, indique o cenário utilizado.",
            "Qual a variação diária percentual esperada para o valor da cota?",
            "Qual a variação diária percentual esperada para o valor da cota do fundo no pior cenário de estresse definido pelo seu administrador?",
            "Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% na taxa anual de juros (pré)?",
            "Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% na taxa de câmbio (US$/Real)?",
            "Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% no preço das ações (IBOVESPA)?",
            # Novas:
            "Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% no principal fator de risco que o fundo está exposto, caso não seja nenhum dos 3 citados anteriormente (juros, câmbio, bolsa)? Considerar o último dia útil do mês de referência. Informar também qual foi o fator de risco considerado.",
            "Indicar o fator de risco",
            "Variação diária percentual esperada"
        ],
        "Resposta": [
            f"{var21_pct:.4f}%",
            "Paramétrico - Delta-Normal " + ("(com correlação)" if usar_corr else "(ρ=0, sem correlação)"),
            DESC_CENARIO["Ibovespa"],
            DESC_CENARIO["Juros-Pré"],
            DESC_CENARIO["Cupom Cambial"],
            DESC_CENARIO["Dólar"],
            DESC_CENARIO["Outros"],
            f"{sigma_port_d*100:.4f}%",
            f"{pior_stress_pct:.4f}%",
            f"{imp_unit('Juros-Pré', -0.01):.4f}%",
            f"{imp_unit('Dólar', -0.01):.4f}%",
            f"{imp_unit('Ibovespa', -0.01):.4f}%",
            resp_outros_composta,
            resp_outros_fator,
            resp_outros_pct
        ]
    })
    st.dataframe(df_cvm, use_container_width=True)
    st.caption(explicacao_outros)
    st.markdown('</div>', unsafe_allow_html=True)

    # Downloads
    st.markdown('<div class="card"><div class="h2">📥 Downloads</div>', unsafe_allow_html=True)
    colA, colB, colC = st.columns(3)
    with colA:
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            pd.DataFrame({
                "Campo":["CNPJ","Fundo","Data","PL (R$)","Confiança","Horizonte","Método"],
                "Valor":[data["cnpj"], data["nome"], data["data"].strftime("%d/%m/%Y"), brl(pl,2),
                         conf_label, f"{h} dias",
                         "Delta-Normal " + ("com correlação" if usar_corr else "ρ=0")]
            }).to_excel(w, sheet_name="Metadados", index=False)
            df_var.to_excel(w, sheet_name="VaR_por_Classe", index=False)
            pd.DataFrame(est_rows).to_excel(w, sheet_name="Cenarios_Estresse", index=False)
            df_cvm.to_excel(w, sheet_name="Respostas_CVM_B3", index=False)
            pd.DataFrame({
                "Métrica":["VaR_port_%","VaR_port_R$","Sigma_diario_%"],
                "Valor":[f"{var_pct*100:.4f}%", brl(var_rs,2), f"{sigma_port_d*100:.4f}%"]
            }).to_excel(w, sheet_name="Sumario", index=False)
        out.seek(0)
        st.download_button("📊 Relatório Completo (Excel)", data=out,
                           file_name=f"relatorio_var_{data['nome'].replace(' ','_')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with colB:
        out2 = BytesIO()
        df_cvm.to_excel(out2, index=False, engine="openpyxl")
        out2.seek(0)
        st.download_button("🏛️ Respostas CVM/B3 (Excel)", data=out2,
                           file_name=f"respostas_cvm_{data['nome'].replace(' ','_')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with colC:
        template = st.file_uploader("📋 Template B3/CVM", type=["xlsx"],
                                    help="Carregue o modelo oficial para preenchimento automático.")
        if template is not None:
            try:
                out_t = BytesIO()
                wb = openpyxl.load_workbook(template); ws = wb.active
                mapa = {col: str(ws.cell(row=3, column=col).value or "").strip().lower()
                        for col in range(3, ws.max_column+1)}
                for _, row in df_cvm.iterrows():
                    p = row["Pergunta"].strip().lower()
                    for col, txt in mapa.items():
                        if p[:50] in txt or txt[:50] in p:
                            ws.cell(row=6, column=col).value = row["Resposta"]; break
                wb.save(out_t); out_t.seek(0)
                st.download_button("📄 Template Preenchido", data=out_t,
                                   file_name=f"template_preenchido_{data['nome'].replace(' ','_')}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Erro ao processar template: {e}")

# ===================== RODAPÉ =====================
st.markdown('<div class="footer">Feito com ❤️ <b>Finhealth</b></div>', unsafe_allow_html=True)

