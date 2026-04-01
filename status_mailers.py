import streamlit as st
import pandas as pd
import json
import os
import glob
from datetime import datetime, timedelta
from streamlit_autorefresh import st_autorefresh

JSON_DIR  = r"Z:\Relações com Investidores - NOVO\codigos\cotas\json"
TIPO_FUNDOS = r"X:\BDM\Novo Modelo de Carteiras\Tipo_Fundos.xlsx"
DIAS_PT   = {0: "Segunda", 1: "Terça", 2: "Quarta", 3: "Quinta", 4: "Sexta"}
DIAS_ABR  = {0: "Seg", 1: "Ter", 2: "Qua", 3: "Qui", 4: "Sex"}
COR_PRIM  = "#1C57A8"

st.set_page_config(page_title="Mailers · Cotas Diárias", layout="wide", page_icon="📬")

# Auto-refresh a cada 2 minutos (120.000 ms)
st_autorefresh(interval=120_000, key="autorefresh")

# ── CSS ──────────────────────────────────────────────────────────────────────
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
    margin-bottom: 24px;
}
.header-title { color: white; font-size: 22px; font-weight: 700; margin: 0; }
.header-sub   { color: rgba(255,255,255,0.75); font-size: 13px; margin: 0; }
.header-week  { color: white; font-size: 15px; font-weight: 600; text-align: right; }

/* Day cards */
.day-card {
    background: white;
    border-radius: 10px;
    padding: 14px 16px;
    border: 1.5px solid #e8edf5;
    text-align: center;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    transition: border-color .2s;
}
.day-card.hoje { border-color: #1C57A8; box-shadow: 0 2px 10px rgba(28,87,168,0.15); }
.day-card.futuro { opacity: .45; }
.day-name  { font-size: 12px; font-weight: 600; color: #6b7a99; text-transform: uppercase; letter-spacing: .05em; }
.day-date  { font-size: 20px; font-weight: 700; color: #1a2540; margin: 2px 0 8px; }
.day-count { font-size: 13px; color: #444; margin-bottom: 8px; }
.day-count span { font-weight: 700; font-size: 16px; }
.day-count .total { color: #888; }

.progress-bg { background: #e8edf5; border-radius: 99px; height: 7px; overflow: hidden; }
.progress-fill { height: 7px; border-radius: 99px; transition: width .4s ease; }
.progress-ok   { background: linear-gradient(90deg,#28a745,#5cb85c); }
.progress-pend { background: linear-gradient(90deg,#e8a000,#ffc107); }
.progress-zero { background: #dc3545; }

.badge-ok   { background:#d4edda; color:#155724; border-radius:99px; padding:2px 10px; font-size:11px; font-weight:600; }
.badge-pend { background:#fff3cd; color:#856404; border-radius:99px; padding:2px 10px; font-size:11px; font-weight:600; }
.badge-zero { background:#f8d7da; color:#721c24; border-radius:99px; padding:2px 10px; font-size:11px; font-weight:600; }
.badge-fut  { background:#e8edf5; color:#6b7a99; border-radius:99px; padding:2px 10px; font-size:11px; font-weight:600; }

/* Filters */
.filter-bar {
    background: #f7f9fc;
    border-radius: 10px;
    padding: 14px 20px;
    margin-bottom: 16px;
    border: 1px solid #e8edf5;
}

/* Table */
.tabela-wrapper { border-radius: 10px; overflow: hidden; border: 1px solid #e8edf5; }
thead th {
    background: #1C57A8 !important;
    color: white !important;
    font-weight: 600 !important;
    text-align: center !important;
    padding: 10px 8px !important;
    font-size: 13px !important;
}
tbody tr:nth-child(even) td { background: #f7f9fc; }
tbody tr:hover td { background: #eef3fb !important; }
td { padding: 7px 10px !important; font-size: 13px !important; }
.cel-ok   { background: #d4edda !important; color: #155724; text-align:center; font-size:15px; }
.cel-pend { background: #f8d7da !important; color: #721c24; text-align:center; font-size:15px; }
.cel-fut  { background: #f2f4f8 !important; color: #bbb;    text-align:center; font-size:13px; }
.cel-nome { font-weight: 500; color: #1a2540; }

/* Nav buttons */
div[data-testid="column"] button {
    width: 100%;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 600;
    padding: 6px 0;
}

/* Expander pendentes */
.pendente-tag {
    display:inline-block;
    background:#fff3cd;
    color:#6d4c0a;
    border:1px solid #ffc107;
    border-radius:6px;
    padding:3px 10px;
    font-size:12px;
    font-weight:500;
    margin:3px;
}
</style>
""", unsafe_allow_html=True)


# ── DADOS ────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def get_fundos():
    df = pd.read_excel(TIPO_FUNDOS, usecols="A,E,F,I,J,K,L")
    df = df[df["Encerrado"].isna() & df["modelo_mailer"].notna()].copy()
    df["fundo"] = df["fundo"].replace("CAPITANIA FCOPEL", "FCopel")
    df["ADM"] = df["ADM"].fillna("Outro")
    return df[["fundo", "ADM"]].sort_values("fundo").reset_index(drop=True)

fundo_df = get_fundos()
fundos    = fundo_df["fundo"].tolist()
adms      = sorted(fundo_df["ADM"].unique().tolist())


# ── SEMANA ───────────────────────────────────────────────────────────────────
if "offset" not in st.session_state:
    st.session_state.offset = 0

today  = datetime.today()
monday = today - timedelta(days=today.weekday()) + timedelta(weeks=st.session_state.offset)
dias   = [monday + timedelta(days=i) for i in range(5)]

semana_label = f"{monday.strftime('%d/%m')} — {dias[-1].strftime('%d/%m/%Y')}"
semana_atual = st.session_state.offset == 0


# ── HEADER ───────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="header-bar">
  <div>
    <p class="header-title">📬 Mailers · Cotas Diárias</p>
    <p class="header-sub">Capitânia Investimentos · Relações com Investidores</p>
  </div>
  <div class="header-week">
    {'Semana atual' if semana_atual else 'Semana'}<br>
    {semana_label}
  </div>
</div>
""", unsafe_allow_html=True)


# ── NAVEGAÇÃO ────────────────────────────────────────────────────────────────
c1, c2, c3, c4, c5 = st.columns([1.2, 1.2, 3, 1.2, 1.2])
with c1:
    if st.button("◀  Semana anterior", use_container_width=True):
        st.session_state.offset -= 1
        st.rerun()
with c2:
    if not semana_atual:
        if st.button("🏠  Semana atual", use_container_width=True):
            st.session_state.offset = 0
            st.rerun()
with c4:
    if st.button("Semana seguinte  ▶", use_container_width=True):
        st.session_state.offset += 1
        st.rerun()
with c5:
    if st.button("🔄  Atualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()


# ── CARREGAR JSONs ────────────────────────────────────────────────────────────
# Coluna = data calendário do envio. JSON = D-1 dessa data (referência da cota).
# Ex: coluna Qua 01/04 → carrega processados_20260331.json (D-1 = 31/03)
#     coluna Seg 30/03 → carrega processados_20260327.json (D-1 de seg = sex anterior)
def ref_de(d: datetime) -> datetime:
    """Retorna a data de referência (D-1 útil) de um dia calendário."""
    if d.weekday() == 0:
        return d - timedelta(days=3)   # segunda → sexta anterior
    return d - timedelta(days=1)

status = {}
for d in dias:
    d_str  = d.strftime("%Y%m%d")
    d_ref  = ref_de(d)
    path   = os.path.join(JSON_DIR, f"processados_{d_ref.strftime('%Y%m%d')}.json")
    if d.date() > today.date():
        status[d_str] = None
    elif os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            status[d_str] = set(json.load(f))
    else:
        status[d_str] = set()

total = len(fundos)


# ── CARDS DOS DIAS ────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
cols_cards = st.columns(5)

for i, d in enumerate(dias):
    d_str  = d.strftime("%Y%m%d")
    e_hoje = d.date() == today.date()
    futuro = status[d_str] is None

    with cols_cards[i]:
        if futuro:
            pct, badge = 0, "badge-fut"
            count_html = "<span style='color:#bbb'>—</span>"
            bar_class  = ""
        else:
            env = sum(1 for f in fundos if f in status[d_str])
            pct = int(env / total * 100) if total else 0
            count_html = f"<span>{env}</span> <span class='total'>/ {total}</span>"
            if env == total:
                badge, bar_class = "badge-ok", "progress-ok"
            elif env > 0:
                badge, bar_class = "badge-pend", "progress-pend"
            else:
                badge, bar_class = "badge-zero", "progress-zero"

        card_class = "day-card hoje" if e_hoje else ("day-card futuro" if futuro else "day-card")

        if futuro:
            badge_label = "Aguardando"
        elif env == total:
            badge_label = "Completo"
        elif env > 0:
            badge_label = f"{total - env} pendentes"
        else:
            badge_label = "Nenhum enviado"

        st.markdown(f"""
        <div class="{card_class}">
          <div class="day-name">{DIAS_PT[i]}</div>
          <div class="day-date">{d.strftime('%d/%m')}</div>
          <div class="day-count">{count_html}</div>
          <div class="progress-bg">
            <div class="progress-fill {bar_class}" style="width:{pct}%"></div>
          </div>
          <div style="margin-top:8px"><span class="{badge}">{badge_label}</span></div>
        </div>
        """, unsafe_allow_html=True)


# ── FILTROS ───────────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="filter-bar">', unsafe_allow_html=True)
    fc1, fc2, fc3 = st.columns([2, 2, 3])
    with fc1:
        adm_sel = st.selectbox("ADM", ["Todos"] + adms, label_visibility="visible")
    with fc2:
        hoje_str = today.strftime("%Y%m%d")
        status_opts = ["Todos", "✅ Enviado hoje", "❌ Pendente hoje"]
        status_sel  = st.selectbox("Status hoje", status_opts)
    with fc3:
        busca = st.text_input("Buscar fundo", placeholder="Digite o nome do fundo...")
    st.markdown('</div>', unsafe_allow_html=True)

# Aplicar filtros
df_filtrado = fundo_df.copy()
if adm_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado["ADM"] == adm_sel]
if status_sel == "✅ Enviado hoje" and hoje_str in status and status[hoje_str] is not None:
    df_filtrado = df_filtrado[df_filtrado["fundo"].isin(status[hoje_str])]
elif status_sel == "❌ Pendente hoje" and hoje_str in status and status[hoje_str] is not None:
    df_filtrado = df_filtrado[~df_filtrado["fundo"].isin(status[hoje_str])]
if busca:
    df_filtrado = df_filtrado[df_filtrado["fundo"].str.contains(busca, case=False, na=False)]

fundos_filtrados = df_filtrado["fundo"].tolist()


# ── TABELA ────────────────────────────────────────────────────────────────────
colunas = [f"{DIAS_ABR[i]}<br>{d.strftime('%d/%m')}" for i, d in enumerate(dias)]

linhas = []
for fundo in fundos_filtrados:
    linha = {"Fundo": fundo}
    for i, d in enumerate(dias):
        d_str = d.strftime("%Y%m%d")
        col   = colunas[i]
        if status[d_str] is None:
            linha[col] = "·"
        elif fundo in status[d_str]:
            linha[col] = "✅"
        else:
            linha[col] = "❌"
    linhas.append(linha)

df_tab = pd.DataFrame(linhas).set_index("Fundo") if linhas else pd.DataFrame()

if df_tab.empty:
    st.info("Nenhum fundo encontrado com os filtros selecionados.")
else:
    def colorir(val):
        if val == "✅":
            return "background-color:#d4edda; color:#155724; text-align:center; font-size:15px; font-weight:600"
        elif val == "❌":
            return "background-color:#f8d7da; color:#721c24; text-align:center; font-size:15px"
        return "background-color:#f2f4f8; color:#ccc; text-align:center"

    hoje_col = f"{DIAS_ABR[today.weekday()]}<br>{today.strftime('%d/%m')}"

    def destacar_hoje(col):
        if col.name == hoje_col:
            return ["border-left: 3px solid #1C57A8; border-right: 3px solid #1C57A8"] * len(col)
        return [""] * len(col)

    styled = (
        df_tab.style
        .applymap(colorir)
        .apply(destacar_hoje, axis=0)
        .set_table_styles([
            {"selector": "thead th", "props": [
                ("background-color", COR_PRIM), ("color", "white"),
                ("font-weight", "600"), ("text-align", "center"),
                ("padding", "10px 8px"), ("font-size", "13px")
            ]},
            {"selector": "tbody td", "props": [("font-size", "13px"), ("padding", "6px 10px")]},
            {"selector": "tbody tr:hover td", "props": [("background-color", "#eef3fb")]},
            {"selector": "tbody tr:nth-child(even) td", "props": [("background-color", "#f7f9fc")]},
        ])
        .set_properties(**{"text-align": "left"}, subset=["Fundo"] if "Fundo" in df_tab.columns else [])
    )

    altura = min(80 + len(fundos_filtrados) * 36, 850)
    st.markdown(f"**{len(fundos_filtrados)} fundos** exibidos", unsafe_allow_html=False)
    st.dataframe(styled, use_container_width=True, height=altura)


# ── PENDENTES HOJE ────────────────────────────────────────────────────────────
if hoje_str in status and status[hoje_str] is not None:
    pendentes = [f for f in fundos if f not in status[hoje_str]]
    if pendentes:
        with st.expander(f"❌  {len(pendentes)} fundos pendentes hoje"):
            tags = "".join(f'<span class="pendente-tag">{f}</span>' for f in pendentes)
            st.markdown(tags, unsafe_allow_html=True)
    else:
        st.success("Todos os fundos foram enviados hoje.")

# ── RODAPÉ ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<hr style='border:none; border-top:1px solid #e8edf5; margin:28px 0 10px'>
<p style='text-align:center; color:#aaa; font-size:12px'>
Capitânia Investimentos · Relações com Investidores &nbsp;|&nbsp;
Atualizado em {today.strftime('%d/%m/%Y %H:%M')}
</p>
""", unsafe_allow_html=True)
