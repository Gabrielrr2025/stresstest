
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
import math
import datetime
import plotly.express as px

st.set_page_config(page_title="VaR & Estresse • Paramétrico por Classe", page_icon="📊", layout="wide")

# =============== Estilo claro minimalista ===============
st.markdown("""
<style>
:root{
  --bg:#ffffff; --panel:#f8fafc; --text:#0f172a; --muted:#475569; --border:#e2e8f0; --accent:#10b981;
}
html, body, [data-testid="stAppViewContainer"]{background:var(--bg);color:var(--text)}
h1,h2,h3{color:#0b1324}
.badge{display:inline-flex;align-items:center;gap:.5rem;border:1px solid var(--border);padding:.35rem .6rem;border-radius:999px;background:#f1f5f9;color:var(--muted);font-size:.8rem}
.section{margin-top:.25rem}
.card{background:var(--panel);border:1px solid var(--border);border-radius:16px;padding:14px}
.kpi{display:flex;flex-direction:column;gap:4px;border:1px solid var(--border);border-radius:14px;padding:12px;background:#fff}
.kpi .l{color:var(--muted);font-size:.75rem}.kpi .v{font-size:1.35rem;font-weight:700}.kpi .s{color:var(--muted);font-size:.75rem}
hr.soft{border:none;height:1px;background:var(--border);margin:6px 0 14px}
</style>
""", unsafe_allow_html=True)

st.markdown("# 📊 VaR Paramétrico por Classe (sem correlação)")
st.markdown("<div class='badge'>Claro • limpo • sem tickers • focado em classes</div>", unsafe_allow_html=True)
st.markdown("<hr class='soft'/>", unsafe_allow_html=True)

left, right = st.columns([1.05,2.0])

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

with left:
    st.markdown("### Painel de Controles")
    with st.container():
        st.markdown("#### Dados do Fundo")
        with st.container():
            c1,c2 = st.columns(2)
            with c1:
                cnpj = st.text_input("CNPJ do Fundo", placeholder="00.000.000/0001-00")
                data_ref = st.date_input("Data de Referência", value=datetime.date.today())
            with c2:
                nome_fundo = st.text_input("Nome do Fundo", placeholder="Fundo XPTO")
                pl = st.number_input("Patrimônio Líquido (R$)", min_value=0.0, value=0.0, step=1000.0, format="%.2f")

    with st.container():
        st.markdown("#### Parâmetros de VaR")
        c1,c2 = st.columns(2)
        with c1:
            horizonte_dias = st.selectbox("Horizonte de VaR (dias)", [1,10,21], index=2,
                help="Para a pergunta CVM usamos 21 dias e 95%.")
        with c2:
            conf_label = st.selectbox("Nível de confiança", ["95%","99%"], index=0,
                help="Método: Paramétrico (Delta-Normal) **sem correlação**.")
        z = 1.65 if conf_label=="95%" else 2.33

    with st.container():
        st.markdown("#### Alocação por Classe")
        st.caption("Vol_Anual em decimal (0.25 = 25% a.a.). Ajuste conforme suas estimativas internas.")
        cols = st.columns(2)
        carteira, soma = [], 0.0
        for i,(classe, vol) in enumerate(VOL_PADRAO.items()):
            with cols[i%2]:
                perc = st.number_input(f"{classe} — % do PL", 0.0, 100.0, 0.0, 1.0, key=f"p_{classe}")
                vol_a = st.number_input(f"Vol anual {classe}", 0.0, 5.0, float(vol), 0.01, format="%.2f", key=f"v_{classe}")
                if perc>0:
                    carteira.append({"classe":classe,"%PL":perc,"vol_anual":float(vol_a)})
                soma += perc
        if soma==100: st.success(f"Soma: {soma:.1f}%")
        elif soma>100: st.error(f"Soma: {soma:.1f}% (excede 100%)")
        elif soma>0: st.info(f"Soma: {soma:.1f}%")

    with st.container():
        st.markdown("#### Cenários de Estresse (FPR)")
        df_cen = pd.DataFrame([{"Fator":k,"Choque":v,"Descrição":DESC_CENARIO[k]} for k,v in CENARIOS_PADRAO.items()])
        df_cen = st.data_editor(df_cen, num_rows="dynamic", use_container_width=True)

    calcular = st.button("Calcular", type="primary", use_container_width=True, disabled=pl<=0 or soma<=0 or soma>100)

with right:
    st.markdown("### Resultados")
    if pl<=0 or soma<=0 or soma>100:
        st.info("Preencha CNPJ, nome, PL e a alocação (soma > 0 e ≤ 100%).")
    elif calcular:
        # VaR por classe
        for it in carteira:
            vol_d = it["vol_anual"]/np.sqrt(252)
            var_pct = z*vol_d*np.sqrt(horizonte_dias)
            it["VaR_%"] = var_pct*100
            it["VaR_R$"] = pl*(it["%PL"]/100)*var_pct
        df_var = pd.DataFrame(carteira)
        var_total = df_var["VaR_R$"].sum()
        var_total_pct = (var_total/pl) if pl>0 else 0.0

        # KPIs
        k1,k2,k3,k4 = st.columns(4)
        with k1: st.markdown(f"<div class='kpi'><div class='l'>VaR ({conf_label}/{horizonte_dias}d)</div><div class='v'>{var_total_pct*100:.2f}%</div><div class='s'>Delta-Normal</div></div>", unsafe_allow_html=True)
        with k2: st.markdown(f"<div class='kpi'><div class='l'>VaR (R$)</div><div class='v'>R$ {var_total:,.0f}</div><div class='s'>Perda potencial</div></div>", unsafe_allow_html=True)
        with k3: st.markdown(f"<div class='kpi'><div class='l'>Classes</div><div class='v'>{(df_var['%PL']>0).sum()}</div><div class='s'>em uso</div></div>", unsafe_allow_html=True)
        with k4: st.markdown(f"<div class='kpi'><div class='l'>Soma %PL</div><div class='v'>{soma:.1f}%</div><div class='s'>deve ≤ 100%</div></div>", unsafe_allow_html=True)

        st.markdown("#### VaR por Classe")
        st.dataframe(df_var.style.format({"%PL":"{:.1f}%","vol_anual":"{:.2%}","VaR_%":"{:.2f}","VaR_R$":"R$ {:,.0f}"}), use_container_width=True)

        c1,c2 = st.columns(2)
        with c1:
            st.plotly_chart(px.pie(df_var, values="%PL", names="classe", title="Distribuição da Carteira"), use_container_width=True)
        with c2:
            st.plotly_chart(px.bar(df_var, x="classe", y="VaR_R$", title="VaR por Classe (R$)"), use_container_width=True)

        # Estresse
        res = []
        for _,r in df_cen.iterrows():
            fator, choque = str(r["Fator"]), float(r["Choque"])
            impacto = sum(choque*(it["%PL"]/100) for it in carteira if fator.lower() in it["classe"].lower())
            res.append({"Fator de Risco":fator,"Descrição":str(r.get("Descrição","")),"Choque":choque,"Impacto % do PL":impacto*100,"Impacto (R$)":impacto*pl})
        df_estresse = pd.DataFrame(res)
        st.markdown("#### Estresse por Fator de Risco")
        st.dataframe(df_estresse.style.format({"Choque":"{:+.2%}","Impacto % do PL":"{:+.4f}","Impacto (R$)":"R$ {:,.0f}"}), use_container_width=True)

        # Respostas CVM/B3
        var21_total = 0.0
        for it in carteira:
            vol_d = it["vol_anual"]/np.sqrt(252)
            var21_total += pl*(it["%PL"]/100)*(1.65*vol_d*np.sqrt(21))
        var21_pct = (var21_total/pl) if pl>0 else 0.0

        get = lambda fator: df_estresse.loc[df_estresse["Fator de Risco"]==fator,"Impacto % do PL"].values[0] if (df_estresse["Fator de Risco"]==fator).any() else 0.0
        pior_stress = float(df_estresse["Impacto % do PL"].min()) if not df_estresse.empty else 0.0

        df_resp = pd.DataFrame({
            "Pergunta":[
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
            "Resposta":[
                f"{var21_pct*100:.4f}%",
                "Paramétrico - Delta Normal (sem correlação)",
                str(df_estresse.loc[df_estresse['Fator de Risco']=='Ibovespa','Descrição'].values[0]) if (df_estresse['Fator de Risco']=='Ibovespa').any() else "—",
                str(df_estresse.loc[df_estresse['Fator de Risco']=='Juros-Pré','Descrição'].values[0]) if (df_estresse['Fator de Risco']=='Juros-Pré').any() else "—",
                str(df_estresse.loc[df_estresse['Fator de Risco']=='Cupom Cambial','Descrição'].values[0]) if (df_estresse['Fator de Risco']=='Cupom Cambial').any() else "—",
                str(df_estresse.loc[df_estresse['Fator de Risco']=='Dólar','Descrição'].values[0]) if (df_estresse['Fator de Risco']=='Dólar').any() else "—",
                str(df_estresse.loc[df_estresse['Fator de Risco']=='Outros','Descrição'].values[0]) if (df_estresse['Fator de Risco']=='Outros').any() else "—",
                f"{df_var['VaR_%'].mean():.4f}%",
                f"{pior_stress:.4f}%",
                f"{get('Juros-Pré'):.4f}%",
                f"{get('Dólar'):.4f}%",
                f"{get('Ibovespa'):.4f}%"
            ]
        })
        st.markdown("#### Respostas CVM/B3")
        st.dataframe(df_resp, use_container_width=True)

        st.markdown("#### Downloads")
        c1,c2 = st.columns(2)
        with c1:
            b = BytesIO(); df_resp.to_excel(b, index=False, engine="openpyxl"); b.seek(0)
            st.download_button("Baixar Respostas (XLSX)", b, file_name=f"respostas_{nome_fundo.replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with c2:
            tpl = st.file_uploader("Template B3/CVM (opcional)", type=["xlsx"])
            if tpl is not None:
                try:
                    out = BytesIO()
                    wb = openpyxl.load_workbook(tpl); ws = wb.active
                    for col in range(3, ws.max_column+1):
                        pergunta = ws.cell(row=3, column=col).value
                        if pergunta:
                            txt = str(pergunta).strip()
                            for _, r in df_resp.iterrows():
                                if r["Pergunta"].strip()[:50] in txt[:50]:
                                    ws.cell(row=6, column=col).value = r["Resposta"]
                                    break
                    wb.save(out); out.seek(0)
                    st.download_button("Baixar Template Preenchido (XLSX)", out,
                        file_name=f"perfil_mensal_{nome_fundo.replace(' ','_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.warning(f"Erro ao preencher template: {e}")
