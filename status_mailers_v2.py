import streamlit as st
import pandas as pd
import json
import os
import glob
import re
from datetime import datetime, timedelta
from streamlit_autorefresh import st_autorefresh
import holidays

JSON_DIR  = r"Z:\Relações com Investidores - NOVO\codigos\cotas\json"
PDF_DIR   = r"Z:\Relações com Investidores - NOVO\codigos\cotas\PDFs"
TIPO_FUNDOS = r"X:\BDM\Novo Modelo de Carteiras\Tipo_Fundos.xlsx"
DIAS_PT   = {0: "Segunda", 1: "Terça", 2: "Quarta", 3: "Quinta", 4: "Sexta"}
DIAS_ABR  = {0: "Seg", 1: "Ter", 2: "Qua", 3: "Qui", 4: "Sex"}
COR_PRIM  = "#1C57A8"
ROBO_LOG  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "robo_log.txt")

MANUAIS_LISTA = [
    "FCopel", "FCopel_Imob", "Sabesprev", "CAPITANIA REIT",
    "OPOR IMOB FII", "OPOR IMOB SUBCLA", "OPOR IMOB SUBCLB", "OPOR IMOB SUBCLC",
]


st.set_page_config(page_title="RI | Dash Cotas", layout="wide", page_icon="📬")

# ── LOGIN ───────────────────────────────────────────────────────────────────
LOGIN_USER = "RI"
LOGIN_PASS = "Capitania2025!"

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_left, col_form, col_right = st.columns([1.5, 1, 1.5])
    with col_form:
        st.markdown(f"""
        <div style="text-align:center; margin-bottom:24px;">
            <span style="font-size:36px;">📬</span>
            <h2 style="color:{COR_PRIM}; margin:8px 0 4px;">Mailers · Cotas Diárias</h2>
            <p style="color:#64748b; font-size:13px;">Capitania Investimentos</p>
        </div>
        """, unsafe_allow_html=True)
        with st.form("login_form"):
            usuario = st.text_input("Usuário")
            senha   = st.text_input("Senha", type="password")
            entrar  = st.form_submit_button("Entrar", use_container_width=True)
            if entrar:
                if usuario == LOGIN_USER and senha == LOGIN_PASS:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")
    st.stop()

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
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    transition: transform .15s, box-shadow .15s;
    position: relative;
}
.day-card:hover { transform: translateY(-2px); box-shadow: 0 6px 18px rgba(0,0,0,0.11); }
.day-card.hoje  { box-shadow: 0 4px 20px rgba(28,87,168,0.18); }
.day-card.futuro { opacity: .5; }

.card-accent {
    height: 5px;
    width: 100%;
    background: #e8edf5;
}
.card-accent.ok   { background: linear-gradient(90deg, #22c55e, #16a34a); }
.card-accent.pend { background: linear-gradient(90deg, #f59e0b, #d97706); }
.card-accent.hoje-pend { background: linear-gradient(90deg, #1C57A8, #2e7dd1); }
.card-accent.zero { background: linear-gradient(90deg, #ef4444, #dc2626); }
.card-accent.fut  { background: #e8edf5; }

.card-inner { padding: 16px 14px 14px; text-align: center; }

.day-name { font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: .08em; }
.day-date { font-size: 26px; font-weight: 800; color: #1a2540; margin: 4px 0 2px; line-height: 1; }
.day-date.hoje-color { color: #1C57A8; }

.day-pct  { font-size: 13px; font-weight: 700; color: #64748b; margin-bottom: 10px; }

.progress-bg   { background: #f1f5f9; border-radius: 99px; height: 6px; overflow: hidden; margin-bottom: 12px; }
.progress-fill { height: 6px; border-radius: 99px; transition: width .5s ease; }
.progress-ok   { background: linear-gradient(90deg,#22c55e,#16a34a); }
.progress-pend { background: linear-gradient(90deg,#f59e0b,#d97706); }
.progress-hoje { background: linear-gradient(90deg,#1C57A8,#2e7dd1); }
.progress-zero { background: #ef4444; }

.status-pill {
    display: inline-block;
    padding: 3px 12px;
    border-radius: 99px;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: .02em;
}
.pill-ok   { background:#dcfce7; color:#15803d; }
.pill-pend { background:#fef3c7; color:#92400e; }
.pill-hoje { background:#dbeafe; color:#1e40af; }
.pill-zero { background:#fee2e2; color:#991b1b; }
.pill-fut     { background:#f1f5f9; color:#94a3b8; }
.card-accent.feriado { background: linear-gradient(90deg, #818cf8, #6366f1); }
.pill-feriado { background:#e0e7ff; color:#3730a3; }

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
    df = df[~df["fundo"].isin(["PETROS RFCP"])]
    df["fundo"] = df["fundo"].replace("CAPITANIA FCOPEL", "FCopel")
    df["ADM"] = df["ADM"].fillna("Outro")
    df["Tipo"] = "Auto"
    # Site com template: robo gera rascunho, dash mostra como Site
    fundos_site = ["BNYCL12879", "CSHG MAGIS II", "BNY12748", "BNYCL12975",
                   "CAPIT D INC FIC", "PORTFOLIO FIDC", "CAPITANIA PREV BP",
                   "CAPITANIA YIELD 120", "INFRA ADV CLA", "XP INFRA90"]
    df.loc[df["fundo"].isin(fundos_site), "Tipo"] = "Site"
    # Fundos de envio manual e Site sem template
    extras = pd.DataFrame([
        {"fundo": "FCopel",          "ADM": "Itau",     "Tipo": "Manual"},
        {"fundo": "FCopel_Imob",     "ADM": "Itau",     "Tipo": "Manual"},
        {"fundo": "Sabesprev",       "ADM": "Itau",     "Tipo": "Manual"},
        {"fundo": "CAPITANIA REIT",  "ADM": "BNYM",     "Tipo": "Manual"},
        {"fundo": "OPOR IMOB FII",   "ADM": "XP",       "Tipo": "Manual"},
        {"fundo": "OPOR IMOB SUBCLA","ADM": "XP",       "Tipo": "Manual"},
        {"fundo": "OPOR IMOB SUBCLB","ADM": "XP",       "Tipo": "Manual"},
        {"fundo": "OPOR IMOB SUBCLC","ADM": "XP",       "Tipo": "Manual"},
        {"fundo": "CAPIT REIT FI",   "ADM": "BNYM",     "Tipo": "Site"},
        {"fundo": "CAPIT MULTIPREV", "ADM": "BNYM",     "Tipo": "Site"},
        {"fundo": "CAPIT PREMIUM",   "ADM": "BNYM",     "Tipo": "Site"},
        {"fundo": "CAPIT PREV FDR",  "ADM": "BNYM",     "Tipo": "Site"},
        {"fundo": "CAPITANIA TOP",   "ADM": "BNYM",     "Tipo": "Site"},
    ])
    df = pd.concat([df[["fundo", "ADM", "Tipo"]], extras], ignore_index=True)
    df = df.drop_duplicates(subset="fundo")
    return df.sort_values("fundo").reset_index(drop=True)

fundo_df = get_fundos()
fundos    = fundo_df["fundo"].tolist()
adms      = sorted(fundo_df["ADM"].unique().tolist())
tipos     = sorted(fundo_df["Tipo"].unique().tolist())


# ── SEMANA ───────────────────────────────────────────────────────────────────
if "offset" not in st.session_state:
    st.session_state.offset = 0

today  = datetime.today()

# Feriados nacionais + SP para o ano corrente e adjacentes
_anos_feriados = [today.year - 1, today.year, today.year + 1]
feriados_br = holidays.country_holidays('BR', subdiv='SP', years=_anos_feriados)

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


# ── STATUS DO ROBÔ ───────────────────────────────────────────────────────────
def get_robo_status():
    if not os.path.exists(ROBO_LOG):
        return "desconhecido", "—", "#888"
    try:
        with open(ROBO_LOG, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()
        import re
        timestamps = re.findall(r'\[(\d{2}:\d{2}:\d{2})\]', conteudo)
        if not timestamps:
            return "desconhecido", "—", "#888"
        ultimo = timestamps[-1]
        h, m, s = map(int, ultimo.split(":"))
        ultima_dt = today.replace(hour=h, minute=m, second=s, microsecond=0)
        diff = (today - ultima_dt).total_seconds()
        if diff < 0:
            diff += 86400
        if diff <= 300:   # até 5 min
            return "Ativo", ultimo, "#28a745"
        elif diff <= 600:
            return "Lento", ultimo, "#ffc107"
        else:
            return "Parado", ultimo, "#dc3545"
    except Exception:
        return "desconhecido", "—", "#888"

robo_estado, robo_ultima, robo_cor = get_robo_status()
st.markdown(f"""
<div style="display:flex; align-items:center; gap:10px; margin-bottom:16px;">
  <div style="width:11px; height:11px; border-radius:50%; background:{robo_cor}; box-shadow:0 0 6px {robo_cor};"></div>
  <span style="font-size:13px; color:#444; font-weight:500;">
    Robô: <strong style="color:{robo_cor}">{robo_estado}</strong>
    &nbsp;·&nbsp; Última verificação: <strong>{robo_ultima}</strong>
  </span>
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
    """Retorna a data de referência (D-1 útil) de um dia calendário, pulando fins de semana e feriados."""
    prev = d - timedelta(days=1)
    while prev.weekday() >= 5 or prev.date() in feriados_br:
        prev -= timedelta(days=1)
    return prev

status     = {}
erros      = {}
horarios   = {}
timestamps = {}   # {d_str: {fundo: {"dt": datetime, "atrasado": bool}}}
manuais_aprovados = {}  # {d_str: set de fundos manuais aprovados}

for d in dias:
    d_str = d.strftime("%Y%m%d")
    d_ref = ref_de(d).strftime("%Y%m%d")

    def _load(prefix):
        path = os.path.join(JSON_DIR, f"{prefix}_{d_ref}.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    # Carregar aprovados manuais do dia
    _aprov_path = os.path.join(JSON_DIR, f"aprovados_{d_ref}.json")
    _aprov_manual = set()
    if os.path.exists(_aprov_path):
        with open(_aprov_path, "r", encoding="utf-8") as f:
            _aprov_manual = set(json.load(f).get("manual", []))
    manuais_aprovados[d_str] = _aprov_manual

    if d.date() in feriados_br:
        status[d_str]     = "feriado"
        erros[d_str]      = {}
        horarios[d_str]   = {}
        timestamps[d_str] = {}
    elif d.date() > today.date():
        status[d_str]     = None
        erros[d_str]      = _load("erros")
        horarios[d_str]   = _load("horarios")
        timestamps[d_str] = {}
    else:
        # Escaneia pasta de PDFs: quais fundos foram gerados e quando
        pdfs = glob.glob(os.path.join(PDF_DIR, f"*_{d_ref}.pdf"))
        processados = set()
        ts_dia = {}
        for p in pdfs:
            nome = os.path.basename(p).rsplit(f"_{d_ref}.pdf", 1)[0]
            processados.add(nome)
            dt_criacao = datetime.fromtimestamp(os.path.getmtime(p))
            atrasado   = dt_criacao.date() > d.date()
            ts_dia[nome] = {"dt": dt_criacao, "atrasado": atrasado}
        status[d_str]     = processados
        erros[d_str]      = _load("erros")
        horarios[d_str]   = _load("horarios")
        timestamps[d_str] = ts_dia

total = len(fundos)


# ── CARDS DOS DIAS ────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
cols_cards = st.columns(5)

for i, d in enumerate(dias):
    d_str  = d.strftime("%Y%m%d")
    e_hoje = d.date() == today.date()
    futuro = status[d_str] is None

    with cols_cards[i]:
        if status[d_str] == "feriado":
            nome_feriado = feriados_br.get(d.date(), "Feriado")
            st.markdown(f"""
            <div class="day-card">
              <div class="card-accent feriado"></div>
              <div class="card-inner">
                <div class="day-name">{DIAS_PT[i]}</div>
                <div class="day-date">{d.strftime('%d/%m')}</div>
                <div class="day-pct">—</div>
                <div class="progress-bg"><div class="progress-fill" style="width:0%"></div></div>
                <span class="status-pill pill-feriado" title="{nome_feriado}">🏖️ Feriado</span>
              </div>
            </div>
            """, unsafe_allow_html=True)
            continue
        if futuro:
            pct        = 0
            accent     = "fut"
            bar_class  = ""
            pill_class = "pill-fut"
            pill_label = "Aguardando"
            pct_label  = "—"
        else:
            env = sum(1 for f in fundos if f in status[d_str])
            pct = int(env / total * 100) if total else 0
            pct_label = f"{env} / {total}"
            if env == total:
                accent, bar_class, pill_class = "ok", "progress-ok", "pill-ok"
                pill_label = "Completo"
            elif env > 0 and e_hoje:
                accent, bar_class, pill_class = "hoje-pend", "progress-hoje", "pill-hoje"
                pill_label = f"{total - env} pendentes"
            elif env > 0:
                accent, bar_class, pill_class = "pend", "progress-pend", "pill-pend"
                pill_label = f"{total - env} pendentes"
            else:
                accent, bar_class, pill_class = "zero", "progress-zero", "pill-zero"
                pill_label = "Nenhum enviado"

        card_class  = "day-card hoje" if e_hoje else ("day-card futuro" if futuro else "day-card")
        date_class  = "day-date hoje-color" if e_hoje else "day-date"

        st.markdown(f"""
        <div class="{card_class}">
          <div class="card-accent {accent}"></div>
          <div class="card-inner">
            <div class="day-name">{DIAS_PT[i]}</div>
            <div class="{date_class}">{d.strftime('%d/%m')}</div>
            <div class="day-pct">{pct_label}</div>
            <div class="progress-bg">
              <div class="progress-fill {bar_class}" style="width:{pct}%"></div>
            </div>
            <span class="status-pill {pill_class}">{pill_label}</span>
          </div>
        </div>
        """, unsafe_allow_html=True)


# ── FILTROS ───────────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="filter-bar">', unsafe_allow_html=True)
    fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 3])
    with fc1:
        adm_sel = st.selectbox("ADM", ["Todos"] + adms, label_visibility="visible")
    with fc2:
        tipo_sel = st.selectbox("Tipo", ["Todos"] + tipos, label_visibility="visible")
    with fc3:
        hoje_str = today.strftime("%Y%m%d")
        status_opts = ["Todos", "✅ Enviado hoje", "❌ Pendente hoje"]
        status_sel  = st.selectbox("Status hoje", status_opts)
    with fc4:
        busca = st.text_input("Buscar fundo", placeholder="Digite o nome do fundo...")
    st.markdown('</div>', unsafe_allow_html=True)

# Aplicar filtros
df_filtrado = fundo_df.copy()
if adm_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado["ADM"] == adm_sel]
if tipo_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Tipo"] == tipo_sel]
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
    tipo = fundo_df.loc[fundo_df["fundo"] == fundo, "Tipo"].values[0]
    linha = {"Fundo": fundo, "Tipo": tipo}
    for i, d in enumerate(dias):
        d_str = d.strftime("%Y%m%d")
        col   = colunas[i]
        if status[d_str] == "feriado":
            linha[col] = "🏖️"
        elif status[d_str] is None:
            linha[col] = "·"
        elif fundo in status[d_str]:
            ts = timestamps[d_str].get(fundo)
            if ts and ts["atrasado"]:
                linha[col] = f"⚠️ {ts['dt'].strftime('%d/%m %H:%M')}"
            else:
                hora = ts["dt"].strftime("%H:%M") if ts else ""
                linha[col] = f"✅ {hora}" if hora else "✅"
        else:
            if fundo in MANUAIS_LISTA and fundo in manuais_aprovados.get(d_str, set()):
                linha[col] = "ENVIAR"
            else:
                motivo = erros[d_str].get(fundo, "")
                linha[col] = f"❌ {motivo}" if motivo else "❌"
    linhas.append(linha)

df_tab = pd.DataFrame(linhas).set_index("Fundo") if linhas else pd.DataFrame()

if df_tab.empty:
    st.info("Nenhum fundo encontrado com os filtros selecionados.")
else:
    def colorir(val):
        v = str(val)
        if v.startswith("⚠️"):
            return "background-color:#fef3c7; color:#92400e; text-align:center; font-size:12px; font-weight:600"
        elif v.startswith("✅"):
            return "background-color:#d4edda; color:#155724; text-align:center; font-size:13px; font-weight:600"
        elif v == "ENVIAR":
            return "background-color:#dbeafe; color:#1e40af; text-align:center; font-size:13px; font-weight:700"
        elif v.startswith("❌"):
            return "background-color:#f8d7da; color:#721c24; text-align:center; font-size:12px"
        elif v.startswith("🏖️"):
            return "background-color:#e0e7ff; color:#3730a3; text-align:center; font-size:13px"
        return "background-color:#f2f4f8; color:#ccc; text-align:center"

    hoje_col = f"{DIAS_ABR[today.weekday()]}<br>{today.strftime('%d/%m')}"

    def destacar_hoje(col):
        if col.name == hoje_col:
            return ["border-left: 3px solid #1C57A8; border-right: 3px solid #1C57A8"] * len(col)
        return [""] * len(col)

    styled = (
        df_tab.style
        .map(colorir)
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
    st.dataframe(
        styled,
        width="stretch",
        height=altura,
        column_config={"Fundo": st.column_config.TextColumn("Fundo", width=220)}
    )


# ── PENDENTES HOJE ────────────────────────────────────────────────────────────
if hoje_str in status and status[hoje_str] is not None and status[hoje_str] != "feriado":
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
