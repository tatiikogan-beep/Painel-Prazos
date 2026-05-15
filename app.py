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
    layout="wide",
)

st.title("⚖️ Gerador de Relatório — Prazos Preclusivos")
st.caption(f"Controladoria Jurídica · {date.today().strftime('%d/%m/%Y')}")

st.divider()

col_up, col_info = st.columns([2, 1])

with col_up:
    arquivo = st.file_uploader(
        "Selecione a planilha de entrada (.xlsx)",
        type=["xlsx"],
        help="Planilha exportada do LegalOne — CJ Conferência Prazos Preclusivos Pendentes",
    )

with col_info:
    with st.expander("📋 Legenda de cores"):
        st.markdown("""
| Cor | Significado |
|-----|------------|
| 🟢 Verde | Dentro do prazo |
| 🟡 Amarelo | D-1: agir hoje |
| 🌸 Rosa | Fatal hoje |
| 🔴 Vermelho | Prazo perdido |
| 🟠 Laranja | Justificativa / "não" na descrição |
| 🔴 Célula Data | Data ≠ D-1 útil calculado |
        """)

if arquivo:
    st.success(f"✅ **{arquivo.name}** carregado.")

    if st.button("🚀 Gerar Relatório", type="primary", use_container_width=True):
        with st.spinner("Processando..."):
            try:
                output_bytes, resumo = gerar_relatorio(arquivo.read())
            except Exception as e:
                st.error(f"Erro ao processar o arquivo: {e}")
                st.stop()

        st.divider()

        # ── PAINEL DE ALERTAS ─────────────────────────────────────────
        alertas = resumo.get('alertas', [])
        tem_alerta = bool(alertas)

        if tem_alerta:
            st.error(
                "### ⚠️ ATENÇÃO — DIVERGÊNCIAS IDENTIFICADAS\n\n"
                "O relatório contém informações que **divergem das regras padrão** "
                "ou **não constam no cadastro do sistema**. "
                "**Comunique imediatamente à gestão do sistema** para análise e regularização.",
                icon="🚨"
            )

            for tipo, dados in alertas:

                if tipo == 'NAO_MAPEADO':
                    with st.expander(
                        f"🔴 {len(dados)} executante(s) NÃO MAPEADO(S) — não constam nas regras padrão",
                        expanded=True
                    ):
                        st.warning(
                            "Os executantes abaixo **não foram reconhecidos** pelo sistema. "
                            "Suas linhas foram **excluídas do relatório**. "
                            "Informe a gestão do sistema para incluí-los no cadastro.",
                        )
                        import pandas as pd
                        df_nm = pd.DataFrame(dados, columns=["Executante", "Processo"])
                        st.dataframe(df_nm, use_container_width=True, hide_index=True)

                elif tipo == 'DATA_DIVERGENTE':
                    with st.expander(
                        f"🔴 {len(dados)} data(s) FORA DO D-1 ÚTIL — divergência de cadastro",
                        expanded=True
                    ):
                        st.warning(
                            "As linhas abaixo foram cadastradas com data diferente do D-1 útil calculado. "
                            "A **célula Data está marcada em vermelho** no relatório. "
                            "Informe a gestão do sistema para verificação e correção.",
                        )
                        import pandas as pd
                        df_div = pd.DataFrame(dados, columns=["Supervisor", "Executante", "Divergência"])
                        st.dataframe(df_div, use_container_width=True, hide_index=True)

                elif tipo == 'COLUNA_NOVA':
                    with st.expander(
                        f"🟡 {len(dados)} coluna(s) não prevista(s) na documentação",
                        expanded=False
                    ):
                        st.warning(
                            "As colunas abaixo existem no arquivo de entrada mas **não constam "
                            "nas regras padrão do sistema**. Informe a gestão para verificar "
                            "se devem ser incluídas na documentação.",
                        )
                        for col in dados:
                            st.code(col)

        else:
            st.success("✅ Nenhuma divergência identificada. Relatório dentro das regras padrão.")

        # ── MÉTRICAS ─────────────────────────────────────────────────
        st.divider()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total de linhas", resumo['total'])
        c2.metric("🟠 Laranjas", resumo['laranjas'])
        c3.metric("🔴 Erros de data", resumo['data_erros'])
        c4.metric("❌ Não mapeados", resumo['nao_mapeados'])

        # ── TABELA RESUMO ─────────────────────────────────────────────
        st.subheader("Resumo por Supervisor")
        tabela = []
        for sup in ORDEM_SUPERVISORES:
            if sup not in resumo['por_supervisor']:
                continue
            c = resumo['por_supervisor'][sup]
            tabela.append({
                "Supervisor":  sup.title(),
                "Total":       c['total'],
                "D-1 🟡":     c['amarelo'],
                "Fatal 🌸":   c['rosa'],
                "Vencido 🔴": c['vermelho'],
                "OK 🟢":      c['verde'],
                "Em Atraso":  c['atraso'],
            })
        if tabela:
            st.dataframe(tabela, use_container_width=True, hide_index=True)

        # ── DOWNLOAD ──────────────────────────────────────────────────
        st.divider()
        nome_saida = arquivo.name.replace('.xlsx', f'_RELATORIO_{date.today().strftime("%d-%m-%Y")}.xlsx')
        st.download_button(
            label="⬇️ Baixar Relatório (.xlsx)",
            data=output_bytes,
            file_name=nome_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
