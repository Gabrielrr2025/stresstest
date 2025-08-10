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

# ===================== ESTADO =====================
if "rodar" not in st.session_state: st.session_state.rodar = False
if "corr_df" not in st.session_state: st.session_state.corr_df = None
if "tentou" not in st.session_state: st.session_state.tentou = False

# ===================== SIDEBAR (Parâmetros formais) =====================
with st.sidebar:
    st.header("⚙️ Parâmetros")

    horizonte_dias = st.selectbox("Horizonte (dias úteis)", [1, 10, 21], index=2,
                                  help="Período considerado para o cálculo do VaR.")
    conf_label = st.selectbox(
        "Confiança",
        ["95%", "99%"], index=0,
        help=(
            "Nível de confiança do VaR (quantil da distribuição de perdas).\n"
            "• 95% (z≈1,645): aceita 5% de chance de exceder o VaR. Normalmente usado em relatórios e limites operacionais.\n"
            "• 99% (z≈2,326): aceita 1% de chance de exceder o VaR. Mais conservador; útil para estresse regulatório.\n"
            "A política de risco do fundo deve definir o nível adotado."
        )
    )
    metodologia = st.selectbox(
        "Metodologia",
        [
            "VaR Paramétrico (Delta-Normal, ρ=0)",
            "VaR Paramétrico (Delta-Normal, com correlação)"
        ],
        index=0,
        help="Modelo paramétrico variância-covariância. A versão com correlação utiliza matriz Corr para agregação."
    )
    usar_corr = metodologia.endswith("com correlação")

# ===================== CABEÇALHO =====================
st.title("📊 Finhealth VaR")
st.caption("Risco paramétrico por classe • Relatórios e respostas CVM/B3")

# ===================== DADOS DO FUNDO + ALOCAÇÃO =====================
with st.form("form_fundo"):
    st.subheader("🏢 Dados do Fundo")

    c1, c2 = st.columns(2)
    with c1:
        cnpj = st.text_input("CNPJ *", placeholder="00.000.000/0001-00")
        if st.session_state.tentou and not cnpj.strip():
            st.markdown("<div style='color:#d00000'>Informe o CNPJ.</div>", unsafe_allow_html=True)

        nome_fundo = st.text_input("Nome do Fundo *", placeholder="Fundo XPTO")
        if st.session_state.tentou and not nome_fundo.strip():
            st.markdown("<div style='color:#d00000'>Informe o nome do fundo.</div>", unsafe_allow_html=True)

    with c2:
        data_ref = st.date_input("Data de Referência *", value=datetime.date.today())

        pl = st.number_input("Patrimônio Líquido (R$) *", min_value=0.0, value=1_000_000.0,
                             step=1_000.0, format="%.2f")
        if st.session_state.tentou and pl <= 0:
            st.markdown("<div style='color:#d00000'>Informe um valor maior que zero.</div>", unsafe_allow_html=True)

    st.subheader("📊 Alocação por Classe")
    st.caption(
        "Informe a distribuição por classe, a volatilidade anual sugerida e, se aplicável, a sensibilidade."
    )
    with st.expander("ℹ️ O que é Sensibilidade (β)?", expanded=False):
        st.write(
            "- **Definição:** elasticidade do valor da classe ao seu fator de risco. "
            "β=1,0 ⇒ choque de **-1%** no fator gera **-1%** na parcela do PL dessa classe.\n"
            "- **Exemplos:** β=0,5 (metade do efeito); β=-1,0 (efeito inverso). Em juros, β pode refletir **DV01** normalizado.\n"
            "- **Uso no estresse:** impacto = choque × β × (% da classe no PL)."
        )

    carteira, soma = [], 0.0
    faltas_vol = {}

    for classe, vol_sugerida in VOL_PADRAO.items():
        a, b, c = st.columns([1.2, .9, .9])
        with a:
            perc = st.number_input(f"{classe} (%)", min_value=0.0, max_value=100.0,
                                   value=0.0, step=0.5, key=f"p_{classe}")
        with b:
            vol_a = st.number_input("Volatilidade Anual", min_value=0.0, max_value=2.0,
                                    value=float(vol_sugerida), step=0.01, format="%.2f", key=f"v_{classe}")
        with c:
            sens = st.number_input("Sensibilidade (β)", min_value=-10.0, max_value=10.0,
                                   value=1.0, step=0.1, key=f"s_{classe}")

        if perc > 0:
            carteira.append({"classe": classe, "%PL": perc, "vol_anual": float(vol_a), "sens": float(sens)})
            soma += perc
            if st.session_state.tentou and vol_a <= 0:
                faltas_vol[classe] = True
                st.markdown(f"<div style='color:#d00000'>Volatilidade obrigatória para \"{classe}\".</div>", unsafe_allow_html=True)

    # Barra + status (padrão Streamlit)
    st.progress(min(int(soma), 100) / 100)
    if soma == 100:
        st.success("✅ Alocação total: 100%")
    elif soma > 100:
        st.error(f"❌ A soma ultrapassa 100% ({soma:.1f}%).")
    elif soma > 0:
        st.warning(f"⚠️ A soma está em {soma:.1f}%. Complete até 100%.")
    else:
        if st.session_state.tentou:
            st.error("❌ Informe ao menos uma alocação.")

    completar_caixa = st.checkbox("Completar automaticamente com Caixa quando a soma for menor que 100%", value=True)

    submit = st.form_submit_button("🚀 Calcular")
    if submit:
        st.session_state.tentou = True
        missing = []

        if not cnpj.strip(): missing.append("CNPJ")
        if not nome_fundo.strip(): missing.append("Nome do Fundo")
        if pl <= 0: missing.append("Patrimônio Líquido maior que zero")
        if soma == 0: missing.append("Informar ao menos uma classe na alocação")
        if soma > 100: missing.append("Soma da alocação não pode exceder 100%")
        for classe in faltas_vol:
            missing.append(f'Volatilidade anual para "{classe}"')

        if missing:
            st.session_state.rodar = False
            st.error("Por favor, corrija:\n- " + "\n- ".join(missing))
        else:
            if soma < 100 and completar_caixa:
                carteira.append({"classe": "Caixa", "%PL": 100 - soma, "vol_anual": 0.0001, "sens": 0.0})
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
    if "com correlação" in st.session_state.get("metodologia_texto", "").lower():
        usar_corr = True
    # mas preferimos a flag da sidebar:
    # (mantida da seleção anterior)
    if 'usar_corr' in globals():
        pass
    if usar_corr:
        if (st.session_state.corr_df is None) or (list(st.session_state.corr_df.index) != classes):
            st.session_state.corr_df = montar_correlacao(classes)
        with st.expander("🔗 Matriz de correlação (opcional)"):
            st.caption("A matriz deve ser simétrica e ter 1 na diagonal. Ajuste se necessário.")
            edit = st.data_editor(st.session_state.corr_df.round(2), num_rows="fixed", use_container_width=True)
            M = edit.to_numpy(float); M = (M + M.T)/2.0; np.fill_diagonal(M, 1.0)
            st.session_state.corr_df = pd.DataFrame(M, index=classes, columns=classes)
        corr = st.session_state.corr_df.to_numpy(float)

    # Cálculo VaR portfólio
    h = int(horizonte_dias); z = z_value(conf_label)
    var_pct, var_rs, sigma_port_d = var_portfolio(pl, pesos, sigma_d, h, z, corr=corr)

    # VaR isolado por classe (exibição)
    var_cls_pct = (z * sigma_d * np.sqrt(h)) * pesos     # fração do PL
    var_cls_rs = var_cls_pct * pl
    df_var = pd.DataFrame({
        "Classe de Ativo": classes,
        "Alocação (%)": [it["%PL"] for it in carteira],
        "Volatilidade Anual": [it["vol_anual"] for it in carteira],
        "VaR (%)": var_cls_pct * 100,
        "VaR (R$)": var_cls_rs
    })
    df_show = df_var.copy()
    df_show["Alocação (%)"] = df_show["Alocação (%)"].map(lambda x: f"{x:.1f}%")
    df_show["Volatilidade Anual"] = df_show["Volatilidade Anual"].map(lambda x: f"{x:.2%}")
    df_show["VaR (%)"] = df_show["VaR (%)"].map(lambda x: f"{x:.2f}%")
    df_show["VaR (R$)"] = df_show["VaR (R$)"].map(lambda x: brl(x, 0))

    # KPIs
    st.subheader("📌 Indicadores")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("VaR do Portfólio", f"{var_pct*100:.2f}%", f"{h} dias • {conf_label}")
    k2.metric("VaR em Reais", brl(var_rs, 0))
    k3.metric("σ diário da cota", f"{sigma_port_d*100:.2f}%")
    k4.metric("Alocação total", f"{sum([it['%PL'] for it in carteira]):.1f}%")

    st.subheader("📈 VaR por Classe (isolado)")
    st.dataframe(df_show, use_container_width=True)

    # Gráficos (tema padrão)
    g1, g2 = st.columns(2)
    with g1:
        fig = px.pie(df_var, values="Alocação (%)", names="Classe de Ativo", title="Distribuição da Carteira", template="plotly_white")
        fig.update_layout(height=360)
        st.plotly_chart(fig, use_container_width=True)
    with g2:
        fig2 = px.bar(df_var, x="Classe de Ativo", y="VaR (R$)", title="VaR por Classe (R$)",
                      color="VaR (R$)", color_continuous_scale="Blues", template="plotly_white")
        fig2.update_layout(xaxis_title="", yaxis_title="VaR (R$)", height=360)
        fig2.update_xaxes(tickangle=45)
        st.plotly_chart(fig2, use_container_width=True)

    # Estresse
    st.subheader("⚠️ Cenários de Estresse")
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

    # ===================== COMPLIANCE CVM/B3 =====================
    st.subheader("🏛️ Respostas CVM/B3")
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
            "Qual a variação diária percentual esperada para o patrimônio do fundo caso ocorra uma variação negativa de 1% no principal fator de risco que o fundo está exposto, caso não seja nenhum dos 3 citados anteriormente (juros, câmbio, bolsa)? Considerar o último dia útil do mês de referência. Informar também qual foi o fator de risco considerado.",
            "Indicar o fator de risco",
            "Variação diária percentual esperada"
        ],
        "Resposta": [
            f"{var21_pct:.4f}%",
            "VaR Paramétrico (Delta-Normal)" + (" — com correlação" if usar_corr else " — ρ=0, sem correlação"),
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

    # ===================== DOWNLOADS (sem upload) =====================
    st.subheader("📥 Downloads")
    colA, colB = st.columns(2)
    with colA:
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            pd.DataFrame({
                "Campo":["CNPJ","Fundo","Data","PL (R$)","Confiança","Horizonte","Método"],
                "Valor":[data["cnpj"], data["nome"], data["data"].strftime("%d/%m/%Y"), brl(pl,2),
                         conf_label, f"{h} dias",
                         "VaR Paramétrico (Delta-Normal, com correlação)" if usar_corr else "VaR Paramétrico (Delta-Normal, ρ=0)"]
            }).to_excel(w, sheet_name="Metadados", index=False)
            # Exporta dados “crus” também:
            pd.DataFrame(carteira).to_excel(w, sheet_name="Carteira_Input", index=False)
            df_var.to_excel(w, sheet_name="VaR_por_Classe_raw", index=False)
            pd.DataFrame(est_rows).to_excel(w, sheet_name="Cenarios_Estresse", index=False)
            df_cvm.to_excel(w, sheet_name="Respostas_CVM_B3", index=False)
        out.seek(0)
        st.download_button(
            "📊 Relatório Completo (Excel)",
            data=out,
            file_name=f"relatorio_var_{data['nome'].replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with colB:
        out2 = BytesIO()
        df_cvm.to_excel(out2, index=False, engine="openpyxl")
        out2.seek(0)
        st.download_button(
            "🏛️ Respostas CVM/B3 (Excel)",
            data=out2,
            file_name=f"respostas_cvm_{data['nome'].replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ===================== RODAPÉ =====================
st.caption("Feito com ❤️ Finhealth")
