"""
Gerador de Escala Médica — UAI Luizote (Psiquiatria)
Interface Web — Streamlit

Deploy gratuito: https://share.streamlit.io
"""

import io, sys, os, tempfile, datetime
import streamlit as st

sys.path.insert(0, os.path.dirname(__file__))
from gerador_escala import (
    gerar_escala, gerar_excel, nome_periodo,
    MESES_ABREV, MESES_PT
)

# ============================================================
st.set_page_config(
    page_title="Escala UAI Luizote — Psiquiatria",
    page_icon="🏥",
    layout="centered",
)

st.title("🏥 Escala UAI Luizote — Psiquiatria")
st.caption("Gerador institucional · Período: dia 16 → dia 15 do mês seguinte")
st.divider()

# ── Parâmetros ───────────────────────────────────────────────
MESES_NOMES = {
    1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril",
    5:"Maio", 6:"Junho", 7:"Julho", 8:"Agosto",
    9:"Setembro", 10:"Outubro", 11:"Novembro", 12:"Dezembro"
}

hoje = datetime.date.today()

col1, col2 = st.columns(2)
with col1:
    mes = st.selectbox(
        "Mês de referência",
        options=list(MESES_NOMES.keys()),
        format_func=lambda x: MESES_NOMES[x],
        index=hoje.month - 1,
    )
with col2:
    ano = st.number_input("Ano", min_value=2024, max_value=2030,
                          value=hoje.year, step=1)

mariana_ativa = not st.checkbox(
    "Mariana em atestado (desativar plantões)",
    value=True,
)

st.divider()

# ── Geração e download ───────────────────────────────────────
if st.button("🗓️ Gerar Escala", type="primary", use_container_width=True):
    with st.spinner("Gerando escala..."):
        dados = gerar_escala(ano, mes, mariana_ativa=mariana_ativa)

        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_path = tmp.name
        gerar_excel(dados, tmp_path)
        with open(tmp_path, 'rb') as f:
            excel_bytes = f.read()
        os.unlink(tmp_path)

    if mes == 12:
        nome_arquivo = f"ESCALA UAI {MESES_ABREV[mes]}_{MESES_PT[1]} {ano+1}.xlsx"
    else:
        nome_arquivo = f"ESCALA UAI {MESES_ABREV[mes]}_{MESES_PT[mes+1]} {ano}.xlsx"

    st.success(f"✅ {nome_periodo(ano, mes)} — escala gerada.")

    st.download_button(
        label=f"⬇️ Baixar {nome_arquivo}",
        data=excel_bytes,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
