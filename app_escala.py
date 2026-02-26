"""
Gerador de Escala Médica — UAI Luizote (Psiquiatria)
Interface Web — Streamlit

Deploy gratuito: https://share.streamlit.io
"""

import io
import sys
import os
import streamlit as st
import pandas as pd

# Importar módulo principal (deve estar no mesmo diretório)
sys.path.insert(0, os.path.dirname(__file__))
from gerador_escala import (
    gerar_escala, gerar_excel, nome_periodo, contar_plantoes,
    MEDICOS_CONFIG, DIAS_SEMANA_PT, MESES_PT
)

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Escala UAI Luizote — Psiquiatria",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    /* Cabeçalho */
    .titulo-principal {
        font-size: 1.6rem;
        font-weight: 700;
        color: #1a5c2a;
        margin-bottom: 0;
    }
    .subtitulo {
        font-size: 0.9rem;
        color: #555;
        margin-top: 0;
        margin-bottom: 1.5rem;
    }

    /* Cards de contador */
    .card-meta-ok {
        background: #e8f5e9;
        border-left: 5px solid #2e7d32;
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 8px;
    }
    .card-atestado {
        background: #fffde7;
        border-left: 5px solid #f9a825;
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 8px;
    }
    .card-falta {
        background: #ffebee;
        border-left: 5px solid #c62828;
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 8px;
    }
    .card-fixo {
        background: #f5f5f5;
        border-left: 5px solid #9e9e9e;
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 8px;
    }
    .card-nome {
        font-size: 0.85rem;
        font-weight: 600;
        color: #212121;
        margin: 0;
    }
    .card-contagem {
        font-size: 1.1rem;
        font-weight: 700;
        margin: 2px 0 0 0;
    }
    .card-status {
        font-size: 0.78rem;
        color: #555;
        margin: 0;
    }

    /* Seções */
    .secao-titulo {
        font-size: 1.1rem;
        font-weight: 700;
        color: #1a5c2a;
        border-bottom: 2px solid #a5d6a7;
        padding-bottom: 4px;
        margin: 1.5rem 0 0.8rem 0;
    }
    .rpa-titulo {
        color: #e65100;
        border-bottom: 2px solid #ffcc80;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# SIDEBAR — PARÂMETROS
# ============================================================
with st.sidebar:
    st.markdown("## ⚙️ Parâmetros da Escala")
    st.divider()

    MESES_NOMES = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }

    import datetime
    hoje = datetime.date.today()

    mes = st.selectbox(
        "Mês de referência",
        options=list(MESES_NOMES.keys()),
        format_func=lambda x: MESES_NOMES[x],
        index=hoje.month - 1,
    )

    ano = st.number_input(
        "Ano",
        min_value=2024,
        max_value=2030,
        value=hoje.year,
        step=1,
    )

    st.divider()
    st.markdown("**Médicos com situação especial**")

    mariana_ativa = not st.checkbox(
        "Mariana em atestado",
        value=True,
        help="Mariana Zanatta Bechara — quando marcado, fica sem plantões (atestado/afastamento).",
    )

    st.divider()
    gerar = st.button("🗓️ Gerar Escala", type="primary", use_container_width=True)

    st.markdown("---")
    st.markdown(
        "<small>UAI Luizote · Psiquiatria<br>"
        "Período: dia 16 → dia 15 do mês seguinte</small>",
        unsafe_allow_html=True,
    )


# ============================================================
# ÁREA PRINCIPAL
# ============================================================
st.markdown('<p class="titulo-principal">🏥 Gerador de Escala — UAI Luizote</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitulo">Psiquiatria · Preenchimento automático por regras individuais</p>', unsafe_allow_html=True)

if not gerar:
    st.info("👈 Configure os parâmetros na barra lateral e clique em **Gerar Escala**.")
    st.stop()


# ── Geração ──────────────────────────────────────────────────
with st.spinner("Gerando escala..."):
    dados = gerar_escala(ano, mes, mariana_ativa=mariana_ativa)

periodo_str = nome_periodo(ano, mes)

mes_fim = mes + 1 if mes < 12 else 1
ano_fim = ano if mes < 12 else ano + 1
st.success(
    f"✅ Escala gerada: **{periodo_str}** "
    f"(16/{mes:02d}/{ano} → 15/{mes_fim:02d}/{ano_fim})"
)


# ============================================================
# SEÇÃO 1 — CONTADOR DE PLANTÕES
# ============================================================
st.markdown('<p class="secao-titulo">📊 Contador de Plantões</p>', unsafe_allow_html=True)

config_por_nome = {c['nome']: c for c in MEDICOS_CONFIG}

ORDEM = [
    'Gustavo Garcia Gonçalves',
    'Mariana Zanatta Bechara',
    'Maurício Rosa de Almeida Junior',
    'Sergio Monteiro Faim',
    'Melissa Maria R Nascimento',
    'Bruna Silva Freitas',
    'Laura Jorge Diniz Povoa',
    'Valquiria Alves Souza',
]

# Dividir em 2 colunas
col_a, col_b = st.columns(2)

for i, nome in enumerate(ORDEM):
    esc   = dados['escalas'].get(nome, {})
    total = contar_plantoes(esc)
    cfg   = config_por_nome.get(nome, {})
    meta  = cfg.get('meta')
    regra = cfg.get('regra', '')
    col   = col_a if i % 2 == 0 else col_b

    # ── Determinar tipo de card ──
    if regra == 'licenca':
        card_class   = 'card-fixo'
        status_label = '🩺 Licença / Maternidade'
        contagem_str = f"0 plantões"
        meta_str     = "─"

    elif not dados['mariana_ativa'] and nome == 'Mariana Zanatta Bechara':
        card_class   = 'card-atestado'
        status_label = f'⚠️ Atestado — meta: {int(meta)} plantões'
        contagem_str = "0 plantões"
        meta_str     = str(int(meta))

    elif regra == 'bruna':
        card_class   = 'card-fixo'
        status_label = '🌅 Plantão matutino fixo (Seg–Sex)'
        contagem_str = f"{total:.1f} plantões".replace(".0", "")
        meta_str     = "─"

    elif meta is None:
        card_class   = 'card-fixo'
        status_label = '─'
        contagem_str = f"{total:.1f}".replace(".0", "")
        meta_str     = "─"

    else:
        meta_num = float(meta)
        if total >= meta_num:
            card_class   = 'card-meta-ok'
            status_label = '✅ Meta atingida'
        elif total >= meta_num * 0.75:
            card_class   = 'card-atestado'
            faltam       = round(meta_num - total, 1)
            status_label = f'⚠️ Faltam {faltam:.1f} plantão(ões)'.replace(".0 ", " ")
        else:
            card_class   = 'card-falta'
            faltam       = round(meta_num - total, 1)
            status_label = f'✗ Faltam {faltam:.1f} plantão(ões)'.replace(".0 ", " ")

        total_fmt    = int(total) if total == int(total) else total
        meta_fmt     = int(meta_num) if meta_num == int(meta_num) else meta_num
        contagem_str = f"{total_fmt} / {meta_fmt} plantões"
        meta_str     = str(meta_fmt)

    # ── Barra de progresso (só para quem tem meta numérica) ──
    with col:
        st.markdown(
            f'<div class="{card_class}">'
            f'<p class="card-nome">{nome}</p>'
            f'<p class="card-contagem">{contagem_str}</p>'
            f'<p class="card-status">{status_label}</p>'
            f'</div>',
            unsafe_allow_html=True,
        )
        if meta and meta > 0 and regra not in ('licenca', 'bruna'):
            pct = min(float(total) / float(meta), 1.0)
            st.progress(pct)


# ============================================================
# SEÇÃO 2 — RELATÓRIO DE DATAS DESCOBERTAS (RPA)
# ============================================================
st.markdown('<p class="secao-titulo rpa-titulo">📅 Datas Descobertas — Oferta de RPA</p>', unsafe_allow_html=True)

slots_vagos = dados['slots_vagos']
todas_datas = sorted(
    set(slots_vagos['M']) | set(slots_vagos['T']) | set(slots_vagos['N'])
)

if not todas_datas:
    st.success("✅ Todos os turnos estão cobertos — nenhuma data para RPA.")
else:
    dias_semana_full = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']

    linhas = []
    for d in todas_datas:
        vm = d in slots_vagos['M']
        vt = d in slots_vagos['T']
        vn = d in slots_vagos['N']
        vagos = []
        if vm: vagos.append('Manhã')
        if vt: vagos.append('Tarde')
        if vn: vagos.append('Noite')
        linhas.append({
            'Data': d.strftime('%d/%m/%Y'),
            'Dia da Semana': dias_semana_full[d.weekday()],
            'Manhã (M)': '○ Vago' if vm else '● Coberto',
            'Tarde (T)': '○ Vago' if vt else '● Coberto',
            'Noite (N)': '○ Vago' if vn else '● Coberto',
            'Turnos Vagos': ', '.join(vagos),
            '_fds': d.weekday() >= 5,
        })

    df = pd.DataFrame(linhas)

    # Estilização da tabela
    def estilizar(row):
        base = [''] * len(row)
        is_fds = row.get('_fds', False)
        fundo  = 'background-color: #fffde7' if is_fds else 'background-color: #fff'
        for i, col in enumerate(row.index):
            if col == '_fds':
                base[i] = 'display: none'
                continue
            val = str(row[col])
            if '○' in val:
                base[i] = 'background-color: #fff3e0; color: #e65100; font-weight: 600'
            elif '●' in val:
                base[i] = 'color: #2e7d32'
            else:
                base[i] = fundo
        return base

    df_view = df.drop(columns=['_fds'])
    styled  = df_view.style.apply(estilizar, axis=1)

    st.dataframe(styled, use_container_width=True, hide_index=True)

    # Totais
    total_m = sum(1 for d in todas_datas if d in slots_vagos['M'])
    total_t = sum(1 for d in todas_datas if d in slots_vagos['T'])
    total_n = sum(1 for d in todas_datas if d in slots_vagos['N'])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🌅 Manhãs vagas",  total_m)
    c2.metric("🌇 Tardes vagas",  total_t)
    c3.metric("🌙 Noites vagas",  total_n)
    c4.metric("📌 Total de slots", total_m + total_t + total_n)

    st.caption("○ Vago = oferecer plantão ao RPA   ·   ● Coberto = turno já alocado   ·   Linhas amarelas = finais de semana")

    # ── Botão: gerar lista de plantões disponíveis (copiável) ──
    st.markdown("")
    if st.button("📋 Gerar lista de plantões disponíveis para RPA", use_container_width=True):
        dias_semana_copia = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
        linhas_texto = [f"PLANTÕES DISPONÍVEIS — {periodo_str}",
                        f"Período: 16/{mes:02d}/{ano} → 15/{mes_fim:02d}/{ano_fim}",
                        "─" * 46]
        for d in todas_datas:
            vagos_txt = []
            if d in slots_vagos['M']: vagos_txt.append('Manhã')
            if d in slots_vagos['T']: vagos_txt.append('Tarde')
            if d in slots_vagos['N']: vagos_txt.append('Noite')
            dia_str   = dias_semana_copia[d.weekday()]
            data_str  = d.strftime('%d/%m/%Y')
            turnos_str = '  ·  '.join(vagos_txt)
            linhas_texto.append(f"{data_str}  ({dia_str})  —  {turnos_str}")
        linhas_texto.append("─" * 46)
        total_slots_txt = (len(slots_vagos['M']) + len(slots_vagos['T']) + len(slots_vagos['N']))
        linhas_texto.append(f"Total: {len(todas_datas)} datas  |  {total_slots_txt} slots vagos")

        st.text_area(
            label="Selecione o texto abaixo e copie (Ctrl+A → Ctrl+C):",
            value="\n".join(linhas_texto),
            height=300,
            key="texto_rpa_copiavel",
        )


# ============================================================
# SEÇÃO 3 — ALERTAS TÉCNICOS
# ============================================================
if dados['alertas']:
    st.markdown('<p class="secao-titulo" style="color:#c62828; border-color:#ef9a9a">⚠️ Alertas</p>', unsafe_allow_html=True)
    for alerta in dados['alertas']:
        st.error(alerta)


# ============================================================
# SEÇÃO 4 — DOWNLOAD DO EXCEL
# ============================================================
st.markdown('<p class="secao-titulo">📥 Exportar Escala</p>', unsafe_allow_html=True)

# Gerar Excel em memória (sem salvar em disco)
buffer = io.BytesIO()

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Reutilizar a função gerar_excel passando um caminho temporário
import tempfile, os
with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
    tmp_path = tmp.name

gerar_excel(dados, tmp_path)
with open(tmp_path, 'rb') as f:
    excel_bytes = f.read()
os.unlink(tmp_path)

if mes == 12:
    from gerador_escala import MESES_ABREV, MESES_PT as _MESES_PT
    nome_arquivo = f"ESCALA UAI {MESES_ABREV[mes]}_{_MESES_PT[1]} {ano + 1}.xlsx"
else:
    from gerador_escala import MESES_ABREV, MESES_PT as _MESES_PT
    nome_arquivo = f"ESCALA UAI {MESES_ABREV[mes]}_{_MESES_PT[mes + 1]} {ano}.xlsx"

st.download_button(
    label=f"⬇️ Baixar {nome_arquivo}",
    data=excel_bytes,
    file_name=nome_arquivo,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary",
)
