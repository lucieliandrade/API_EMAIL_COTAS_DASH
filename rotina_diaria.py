"""
rotina_diaria.py - Dashboard Rotina Diária · RI
================================================
Checklist interativo das tarefas do dia atual, carregadas do Rotina.xlsx.
Salva estado em JSON para persistência entre sessões.

Rodar:
    python -m streamlit run rotina_diaria.py --server.port 8502 --server.address 0.0.0.0
"""

import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from collections import OrderedDict
from streamlit_autorefresh import st_autorefresh

# ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────
XLSX_PATH = r"C:\Users\lucieli.andrade\OneDrive - Capitania S.A\RI - DADOS COMPARTILHADOS EM PLANILHA\Rotina.xlsx"
JSON_DIR  = r"Z:\Relações com Investidores - NOVO\codigos\cotas\json"
COR_PRIM  = "#1C57A8"
DIAS_PT   = {0: "Segunda-feira", 1: "Terça-feira", 2: "Quarta-feira",
             3: "Quinta-feira", 4: "Sexta-feira"}

st.set_page_config(page_title="Rotina Diária · RI", layout="wide", page_icon="✅")
st_autorefresh(interval=300_000, key="autorefresh_rotina")  # 5 min

today     = datetime.today()
today_str = today.strftime('%Y%m%d')
dia_idx   = today.weekday()  # 0=Seg … 4=Sex


# ── PERSISTÊNCIA ──────────────────────────────────────────────────────────────
def _state_path():
    return os.path.join(JSON_DIR, f"rotina_{today_str}.json")


def carregar_estado():
    p = _state_path()
    if os.path.exists(p):
        with open(p, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def salvar_estado(estado: dict):
    os.makedirs(JSON_DIR, exist_ok=True)
    with open(_state_path(), 'w', encoding='utf-8') as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)


# ── LER TAREFAS DO DIA ────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def carregar_tarefas(dia: int) -> OrderedDict:
    """
    Lê Rotina.xlsx e retorna OrderedDict {secao: [tarefas]} para o dia (0=Seg..4=Sex).
    Seções = linhas onde todas as colunas de dias são NaN.
    Tarefas = linhas com 'X' ou 'x' na coluna do dia.
    """
    df = pd.read_excel(XLSX_PATH, sheet_name='Diária', header=0)

    all_day_cols = ['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5']
    day_col      = all_day_cols[dia]

    # Row 0 = cabeçalho visual do Excel ("Diariamente", "Segunda-feira" …) → pula
    df = df.iloc[1:].reset_index(drop=True)

    result: OrderedDict = OrderedDict()
    secao_atual = 'Geral'
    result[secao_atual] = []

    for _, row in df.iterrows():
        raw = row['Unnamed: 0']
        tarefa = str(raw).strip() if pd.notna(raw) else ''
        if not tarefa or tarefa.lower() == 'nan':
            continue

        # Linha onde TODOS os dias são NaN = título de seção
        if all(pd.isna(row[c]) for c in all_day_cols):
            secao_atual = tarefa
            if secao_atual not in result:
                result[secao_atual] = []
            continue

        val = str(row[day_col]).strip().lower() if pd.notna(row[day_col]) else ''
        if val == 'x':
            result[secao_atual].append(tarefa)

    # Remove seções sem tarefas para o dia
    return OrderedDict((k, v) for k, v in result.items() if v)


# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Header */
.header-bar {
    background: linear-gradient(90deg, #1C57A8 0%, #2e7dd1 100%);
    border-radius: 12px;
    padding: 18px 28px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 20px;
}
.header-title { color: white; font-size: 22px; font-weight: 700; margin: 0; }
.header-sub   { color: rgba(255,255,255,.75); font-size: 13px; margin: 4px 0 0; }
.header-date  { color: white; font-size: 20px; font-weight: 700; text-align: right; margin: 0; }
.header-day   { color: rgba(255,255,255,.8); font-size: 13px; text-align: right; margin: 2px 0 0; }

/* Progresso */
.prog-wrap   { margin-bottom: 18px; }
.prog-label  { font-size: 14px; font-weight: 600; color: #1a2540; margin-bottom: 6px; }
.prog-pct    { font-size: 13px; font-weight: 500; color: #64748b; margin-left: 8px; }
.prog-bg     { background: #f1f5f9; border-radius: 99px; height: 10px; overflow: hidden; }
.prog-fill   { height: 10px; border-radius: 99px; transition: width .4s ease;
               background: linear-gradient(90deg, #1C57A8, #2e7dd1); }
.prog-fill.done { background: linear-gradient(90deg, #22c55e, #16a34a); }

/* Título de seção */
.secao-header {
    border-left: 4px solid #1C57A8;
    padding: 4px 0 4px 12px;
    margin: 22px 0 8px;
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .09em;
    color: #1C57A8;
}
.secao-cnt { font-weight: 400; color: #64748b; margin-left: 6px; font-size: 12px; }

/* Strikethrough em checkboxes marcados */
div[data-testid="stCheckbox"]:has(input:checked) label p,
div[data-testid="stCheckbox"]:has(input:checked) label span {
    text-decoration: line-through !important;
    color: #94a3b8 !important;
}

/* Separador de seção */
.secao-sep { border: none; border-top: 1px solid #e8edf5; margin: 20px 0 0; }
</style>
""", unsafe_allow_html=True)


# ── FIM DE SEMANA ─────────────────────────────────────────────────────────────
if dia_idx > 4:
    st.markdown("""
    <div class="header-bar">
      <div>
        <p class="header-title">✅ Rotina Diária · RI</p>
        <p class="header-sub">Capitânia Investimentos · Relações com Investidores</p>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.info("Hoje é fim de semana. Rotina disponível de Segunda a Sexta.")
    st.stop()


# ── CARREGAR DADOS ────────────────────────────────────────────────────────────
secoes       = carregar_tarefas(dia_idx)
todas_tarefas = [t for ts in secoes.values() for t in ts]
total         = len(todas_tarefas)

# Índice estável: tarefa → posição na lista flat
flat_idx: dict[str, int] = {}
_seen: dict[str, int] = {}
for i, t in enumerate(todas_tarefas):
    if t not in _seen:
        _seen[t] = i
    flat_idx[t] = _seen[t]  # primeiro occurrence wins


# ── INICIALIZAR SESSION STATE (uma vez por dia) ───────────────────────────────
if st.session_state.get('rotina_date') != today_str:
    st.session_state['rotina_date'] = today_str
    estado_salvo = carregar_estado()
    for t in todas_tarefas:
        k = f"cb_{flat_idx[t]}"
        if k not in st.session_state:  # não sobrescreve se já existe nessa sessão
            st.session_state[k] = estado_salvo.get(t, False)


# ── CALLBACK DE SALVAMENTO ────────────────────────────────────────────────────
def _on_change():
    """Salva o estado atual de todos os checkboxes no JSON."""
    estado = {t: bool(st.session_state.get(f"cb_{flat_idx[t]}", False))
              for t in todas_tarefas}
    salvar_estado(estado)


# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="header-bar">
  <div>
    <p class="header-title">✅ Rotina Diária · RI</p>
    <p class="header-sub">Capitânia Investimentos · Relações com Investidores</p>
  </div>
  <div>
    <p class="header-date">{today.strftime('%d/%m/%Y')}</p>
    <p class="header-day">{DIAS_PT[dia_idx]}</p>
  </div>
</div>
""", unsafe_allow_html=True)


# ── BARRA DE PROGRESSO ────────────────────────────────────────────────────────
concluidas = sum(1 for t in todas_tarefas if st.session_state.get(f"cb_{flat_idx[t]}", False))
pct        = int(concluidas / total * 100) if total else 0
done_class = "prog-fill done" if pct == 100 else "prog-fill"

st.markdown(f"""
<div class="prog-wrap">
  <span class="prog-label">{concluidas} / {total} tarefas concluídas</span>
  <span class="prog-pct">{pct}%</span>
  <div class="prog-bg">
    <div class="{done_class}" style="width:{pct}%"></div>
  </div>
</div>
""", unsafe_allow_html=True)


# ── BOTÕES DE AÇÃO ────────────────────────────────────────────────────────────
col_limpar, col_att, _ = st.columns([1, 1, 5])

with col_limpar:
    if st.button("🗑  Limpar tudo", use_container_width=True):
        for t in todas_tarefas:
            st.session_state[f"cb_{flat_idx[t]}"] = False
        salvar_estado({})
        st.rerun()

with col_att:
    if st.button("🔄  Atualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()


# ── CHECKLIST POR SEÇÃO ───────────────────────────────────────────────────────
for secao, tarefas in secoes.items():
    n_feitas = sum(1 for t in tarefas if st.session_state.get(f"cb_{flat_idx[t]}", False))
    n_total  = len(tarefas)

    st.markdown(f"""
    <hr class="secao-sep">
    <div class="secao-header">
        {secao}<span class="secao-cnt">{n_feitas} / {n_total}</span>
    </div>
    """, unsafe_allow_html=True)

    for tarefa in tarefas:
        k = f"cb_{flat_idx[tarefa]}"
        st.checkbox(tarefa, key=k, on_change=_on_change)


# ── RODAPÉ ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<hr style='border:none; border-top:1px solid #e8edf5; margin:32px 0 10px'>
<p style='text-align:center; color:#aaa; font-size:12px'>
  Capitânia Investimentos · Relações com Investidores &nbsp;|&nbsp;
  Atualizado em {today.strftime('%d/%m/%Y %H:%M')}
</p>
""", unsafe_allow_html=True)
