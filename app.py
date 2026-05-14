"""
Painel Web — Prazos Preclusivos
Execute: streamlit run app.py
"""

import streamlit as st
from datetime import date
from gerar_relatorio_prazos import gerar_relatorio, SUPERVISORES

# ─── CONFIGURAÇÃO DA PÁGINA ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Prazos Preclusivos",
    page_icon="⚖️",
    layout="centered",
)

st.title("⚖️ Gerador de Relatório — Prazos Preclusivos")
st.caption(f"Controladoria Jurídica · {date.today().strftime('%d/%m/%Y')}")

# ─── UPLOAD ──────────────────────────────────────────────────────────────────
st.divider()
arquivo = st.file_uploader(
    "Selecione a planilha de entrada (.xlsx)",
    type=["xlsx"],
    help="Planilha exportada do LegalOne — CJ Conferência Prazos Preclusivos Pendentes",
)

if arquivo:
    st.success(f"✅ Arquivo carregado: **{arquivo.name}**")

    if st.button("🚀 Gerar Relatório", type="primary", use_container_width=True):
        with st.spinner("Processando... aguarde."):
            try:
                output_bytes, resumo = gerar_relatorio(arquivo.read())
            except Exception as e:
                st.error(f"Erro ao processar: {e}")
                st.stop()

        # ─── MÉTRICAS ────────────────────────────────────────────────────────
        st.divider()
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de linhas", resumo['total'])
        col2.metric("Erros de data", resumo['data_erros'],
                    delta=None if resumo['data_erros'] == 0 else f"⚠️ {resumo['data_erros']}",
                    delta_color="inverse")
        col3.metric("Não mapeados", resumo['nao_mapeados'],
                    delta=None if resumo['nao_mapeados'] == 0 else f"⚠️ {resumo['nao_mapeados']}",
                    delta_color="inverse")

        # ─── TABELA RESUMO ───────────────────────────────────────────────────
        st.subheader("Resumo por Supervisor")
        sup_data = []
        for sup in SUPERVISORES:
            if sup not in resumo['por_supervisor']:
                continue
            c = resumo['por_supervisor'][sup]
            sup_data.append({
                "Supervisor": sup.title(),
                "Total": c['total'],
                "D-1 🟡": c['amarelo'],
                "Fatal 🌸": c['rosa'],
                "Vencido 🔴": c['vermelho'],
                "OK 🟢": c['verde'],
                "Em Atraso": c['atraso'],
            })
        if sup_data:
            st.dataframe(sup_data, use_container_width=True, hide_index=True)

        # ─── DOWNLOAD ────────────────────────────────────────────────────────
        st.divider()
        nome_saida = arquivo.name.replace('.xlsx', '_RELATORIO.xlsx')
        st.download_button(
            label="⬇️ Baixar Relatório (.xlsx)",
            data=output_bytes,
            file_name=nome_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        if resumo['nao_mapeados'] > 0:
            st.warning(
                f"⚠️ **{resumo['nao_mapeados']} executor(es) não mapeado(s).** "
                "Verifique a aba RESUMO do relatório para detalhes."
            )
        if resumo['data_erros'] > 0:
            st.warning(
                f"⚠️ **{resumo['data_erros']} data(s) fora do D-1 útil.** "
                "Células marcadas em vermelho no relatório."
            )

# ─── LEGENDA ─────────────────────────────────────────────────────────────────
with st.expander("📋 Legenda de cores"):
    st.markdown("""
| Cor | Significado |
|-----|------------|
| 🟢 Verde | Dentro do prazo |
| 🟡 Amarelo | D-1 — agir hoje |
| 🌸 Rosa | Prazo fatal é hoje |
| 🔴 Vermelho escuro | Prazo vencido / pendente de baixa |
| 🟠 Laranja | Justificativa preenchida ou "não" na descrição |
| 🔴 Célula Data | Data cadastrada diverge do D-1 calculado |
    """)
