"""
Painel Web — Prazos Preclusivos
Execute: streamlit run app.py
"""

import streamlit as st
from datetime import date
from gerar_relatorio_prazos import gerar_relatorio, ORDEM_SUPERVISORES

st.set_page_config(
    page_title="Prazos Preclusivos",
    page_icon="⚖️",
    layout="centered",
)

st.title("⚖️ Gerador de Relatório — Prazos Preclusivos")
st.caption(f"Controladoria Jurídica · {date.today().strftime('%d/%m/%Y')}")

st.divider()
arquivo = st.file_uploader(
    "Selecione a planilha de entrada (.xlsx)",
    type=["xlsx"],
    help="Planilha exportada do LegalOne — CJ Conferência Prazos Preclusivos Pendentes",
)

if arquivo:
    st.success(f"✅ Arquivo carregado: **{arquivo.name}**")

    if st.button("🚀 Gerar Relatório", type="primary", use_container_width=True):
        with st.spinner("Processando..."):
            try:
                output_bytes, resumo = gerar_relatorio(arquivo.read())
            except Exception as e:
                st.error(f"Erro ao processar: {e}")
                st.stop()

        # Métricas
        st.divider()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total", resumo['total'])
        c2.metric("Laranjas ⚠️", resumo['laranjas'])
        c3.metric("Erros de data ⚠️", resumo['data_erros'])
        c4.metric("Não mapeados ⚠️", resumo['nao_mapeados'])

        # Tabela resumo
        st.subheader("Resumo por Supervisor")
        tabela = []
        for sup in ORDEM_SUPERVISORES:
            if sup not in resumo['por_supervisor']:
                continue
            c = resumo['por_supervisor'][sup]
            tabela.append({
                "Supervisor": sup.title(),
                "Total": c['total'],
                "D-1 🟡": c['amarelo'],
                "Fatal 🌸": c['rosa'],
                "Vencido 🔴": c['vermelho'],
                "OK 🟢": c['verde'],
                "Em Atraso": c['atraso'],
            })
        if tabela:
            st.dataframe(tabela, use_container_width=True, hide_index=True)

        # Download
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
            st.warning(f"⚠️ {resumo['nao_mapeados']} executor(es) não mapeado(s) — ver aba RESUMO.")
        if resumo['data_erros'] > 0:
            st.warning(f"⚠️ {resumo['data_erros']} data(s) fora do D-1 útil — células em vermelho.")

# Legenda
with st.expander("📋 Legenda de cores"):
    st.markdown("""
| Cor | Significado |
|-----|------------|
| 🟢 Verde | Dentro do prazo (D-1 ainda não chegou) |
| 🟡 Amarelo | D-1 — hoje é o dia de cumprir, fatal amanhã |
| 🌸 Rosa | Prazo fatal é hoje |
| 🔴 Vermelho escuro | Fatal já passou — perda de prazo |
| 🟠 Laranja | Justificativa preenchida ou "não" na descrição |
| 🔴 Célula Data | Data cadastrada ≠ D-1 útil calculado |
    """)
