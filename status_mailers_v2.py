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

# INTRAG (esteira de boletagem Itau Vida)
INTRAG_PASTA = r"Z:\Relações com Investidores - NOVO\Boletas Fundos\INTRAG"
INTRAG_HEARTBEAT = os.path.join(INTRAG_PASTA, "agendador_heartbeat.txt")
INTRAG_PROCESSADOS = os.path.join(INTRAG_PASTA, "processados_intrag.txt")
INTRAG_ESTADO_MANUAL = os.path.join(INTRAG_PASTA, "esteira_estado.json")
INTRAG_PASTA_NET = r"N:\Middle\Resgates\Codigos_movimentacoes_adm\Código Itaú"
DIAS_PT   = {0: "Segunda", 1: "Terça", 2: "Quarta", 3: "Quinta", 4: "Sexta"}
DIAS_ABR  = {0: "Seg", 1: "Ter", 2: "Qua", 3: "Qui", 4: "Sex"}
COR_PRIM  = "#1C57A8"
ROBO_LOG  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "robo_log.txt")

MANUAIS_LISTA = [
    "CAPITANIA FCOPEL", "FCopel_Imob", "Sabesprev", "CAPITANIA REIT", "PETROS RFCP",
    "OPOR IMOB FII", "OPOR IMOB SUBCLA", "OPOR IMOB SUBCLB", "OPOR IMOB SUBCLC",
]

# Mapa de nomes antigos (em JSONs do scan_outlook) para o nome exibido no dash.
# Usado ao ler aprovados_*.json para normalizar nomes historicos.
DASH_DISPLAY_NAME = {
    "FCopel": "CAPITANIA FCOPEL",
}


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


# ── INTRAG (esteira de boletagem) ────────────────────────────────────────────
def _intrag_proc_hoje():
    hoje_iso = datetime.now().date().isoformat()
    if not os.path.exists(INTRAG_PROCESSADOS):
        return None
    try:
        with open(INTRAG_PROCESSADOS, 'r', encoding='utf-8') as f:
            for linha in reversed(list(f)):
                p = linha.strip().split('|')
                if len(p) >= 2 and p[0] == hoje_iso:
                    return {'tipo': p[1], 'ts': p[2] if len(p) > 2 else ''}
    except Exception:
        pass
    return None


def _intrag_heartbeat():
    if not os.path.exists(INTRAG_HEARTBEAT):
        return None, None
    try:
        with open(INTRAG_HEARTBEAT, 'r', encoding='utf-8') as f:
            partes = f.read().strip().split('|')
        ts = datetime.strptime(partes[0], '%Y-%m-%dT%H:%M:%S')
        estado = partes[1] if len(partes) > 1 else None
        if ts.date() != datetime.now().date():
            return None, None
        return ts, estado
    except Exception:
        return None, None


def _intrag_txts_hoje():
    yyyymmdd = datetime.now().strftime('%Y%m%d')
    nomes = [
        f'Passivo_ItauVida_FIE_{yyyymmdd}.txt',
        f'Ativo_FIE_FIFE_{yyyymmdd}.txt',
        f'Passivo_FIE_FIFE_{yyyymmdd}.txt',
    ]
    return sum(1 for n in nomes if os.path.exists(os.path.join(INTRAG_PASTA, n)))


def _intrag_arquivo_net():
    """Retorna (existe, mtime) de qualquer arquivo do dia na pasta net."""
    yyyymmdd = datetime.now().strftime('%Y%m%d')
    if not os.path.isdir(INTRAG_PASTA_NET):
        return False, None
    try:
        for nome in os.listdir(INTRAG_PASTA_NET):
            if nome.startswith(yyyymmdd):
                caminho = os.path.join(INTRAG_PASTA_NET, nome)
                if os.path.isfile(caminho):
                    try:
                        mtime = datetime.fromtimestamp(os.path.getmtime(caminho))
                    except Exception:
                        mtime = None
                    return True, mtime
    except Exception:
        pass
    return False, None


def _intrag_estado_manual_hoje():
    hoje_iso = datetime.now().date().isoformat()
    if not os.path.exists(INTRAG_ESTADO_MANUAL):
        return {}
    try:
        with open(INTRAG_ESTADO_MANUAL, 'r', encoding='utf-8') as f:
            return json.load(f).get(hoje_iso, {})
    except Exception:
        return {}


def _intrag_marcar(chave, valor):
    hoje_iso = datetime.now().date().isoformat()
    todo = {}
    if os.path.exists(INTRAG_ESTADO_MANUAL):
        try:
            with open(INTRAG_ESTADO_MANUAL, 'r', encoding='utf-8') as f:
                todo = json.load(f)
        except Exception:
            todo = {}
    todo.setdefault(hoje_iso, {})
    if valor:
        todo[hoje_iso][chave] = {'feito': True, 'ts': datetime.now().strftime('%H:%M:%S')}
    else:
        todo[hoje_iso].pop(chave, None)
    try:
        os.makedirs(os.path.dirname(INTRAG_ESTADO_MANUAL), exist_ok=True)
        with open(INTRAG_ESTADO_MANUAL, 'w', encoding='utf-8') as f:
            json.dump(todo, f, indent=2, ensure_ascii=False)
    except Exception as e:
        st.warning(f"Falha ao salvar estado INTRAG: {e}")


def _intrag_step_card(col, num, titulo, accent, icon, sub):
    cores = {'ok': '#22c55e', 'pend': '#f59e0b', 'zero': '#ef4444', 'fut': '#94a3b8'}
    cor = cores.get(accent, '#94a3b8')
    with col:
        st.markdown(f"""
        <div style='background:white;border-radius:10px;padding:10px 6px;
                    box-shadow:0 2px 6px rgba(0,0,0,0.06);text-align:center;
                    border-top:4px solid {cor};margin-bottom:6px;min-height:130px'>
          <div style='font-size:10px;color:#94a3b8;font-weight:700;letter-spacing:.05em'>STEP {num}</div>
          <div style='font-size:12px;color:#1a2540;font-weight:600;margin-top:4px;height:32px'>{titulo}</div>
          <div style='font-size:24px;margin:4px 0'>{icon}</div>
          <div style='font-size:10.5px;color:#64748b;height:14px'>{sub}</div>
        </div>
        """, unsafe_allow_html=True)


def render_intrag_esteira():
    """Renderiza a esteira INTRAG de 7 steps."""
    proc = _intrag_proc_hoje()
    hb_ts, hb_estado = _intrag_heartbeat()
    n_txts = _intrag_txts_hoje()
    manual = _intrag_estado_manual_hoje()
    arq_net_existe, arq_net_mtime = _intrag_arquivo_net()

    agora = datetime.now()
    is_dia_util = agora.weekday() < 5

    # Step 1 - Email Itau
    if not is_dia_util:
        s1 = ('fut', '🏖️', 'fim de semana')
    elif proc and proc['tipo'] == 'sucesso':
        hr = proc['ts'].split()[1][:5] if ' ' in proc['ts'] else ''
        s1 = ('ok', '✅', f'recebido {hr}')
    elif proc and proc['tipo'] == 'sem_movimento':
        s1 = ('ok', '✅', 'sem movimento')
    elif proc and proc['tipo'] == 'fim_dia':
        s1 = ('zero', '❌', '17h sem email')
    elif hb_ts and (agora - hb_ts).total_seconds() / 60 < 7:
        s1 = ('pend', '⏳', f"vivo {hb_ts.strftime('%H:%M')}")
    elif agora.hour < 13:
        s1 = ('fut', '⏰', 'inicia 13h')
    elif hb_ts:
        delta_min = int((agora - hb_ts).total_seconds() / 60)
        s1 = ('zero', '⚠️', f'parado {delta_min}min')
    else:
        s1 = ('zero', '⚠️', 'robô off')

    # Step 2 - TXTs gerados
    if not is_dia_util:
        s2 = ('fut', '🏖️', '-')
    elif proc and proc['tipo'] == 'sem_movimento':
        s2 = ('ok', '🏖️', 'sem mov')
    elif n_txts == 3:
        s2 = ('ok', '✅', '3/3 TXTs')
    elif n_txts > 0:
        s2 = ('pend', '⏳', f'{n_txts}/3 TXTs')
    else:
        s2 = ('fut', '·', 'aguarda step 1')

    # Steps 3-6: manuais
    def _step_manual(chave):
        info = manual.get(chave)
        if info and info.get('feito'):
            ts = info.get('ts', '')
            return ('ok', '✅', f'feito {ts[:5]}', True)
        # se ainda nao chegou ao step 2, mostra como futuro
        if not (n_txts == 3 or (proc and proc['tipo'] == 'sem_movimento')):
            return ('fut', '⏳', 'aguarda TXTs', False)
        return ('pend', '⏳', 'pendente', False)

    s3 = _step_manual('subiu_passivo_itau')
    s4 = _step_manual('subiu_ativo_fife')
    s5 = _step_manual('subiu_passivo_fife')
    s6 = _step_manual('liquidado')

    # Step 7 - Arquivo na pasta net (auto)
    if not is_dia_util:
        s7 = ('fut', '🏖️', '-')
    elif arq_net_existe:
        hr = arq_net_mtime.strftime('%H:%M') if arq_net_mtime else ''
        s7 = ('ok', '✅', f'criado {hr}' if hr else 'OK')
    elif s6[3]:  # liquidacao marcada mas arquivo ainda nao apareceu
        s7 = ('pend', '⏳', 'aguardando')
    else:
        s7 = ('fut', '·', 'aguarda step 6')

    st.markdown("<br>", unsafe_allow_html=True)
    titulo_col, btn_intrag_col = st.columns([5, 2])
    with titulo_col:
        st.markdown("### 🏦 Esteira INTRAG · Boletas Itaú Vida")
    with btn_intrag_col:
        st.markdown("<div style='padding-top:10px'></div>", unsafe_allow_html=True)

        # Botao 1: abre pasta (so funciona na maquina servidor = Lucieli)
        if st.button("📁 abrir pasta", key="intrag_abrir_pasta", use_container_width=True, help=INTRAG_PASTA_NET):
            try:
                os.startfile(INTRAG_PASTA_NET)
            except Exception as e:
                st.warning(f"Falha ao abrir: {e}")

        # Botao 2: copia caminho pro clipboard (funciona pra qualquer usuario)
        caminho_js = INTRAG_PASTA_NET.replace('\\', '\\\\').replace("'", "\\'")
        st.markdown(f"""
        <button onclick="
        navigator.clipboard.writeText('{caminho_js}').then(()=>{{
          this.innerHTML='✅ caminho copiado! cole no Win+R';
          setTimeout(()=>{{this.innerHTML='📋 copiar caminho rede';}},3000);
        }});
        " style="
        background:#e8edf5;color:#1C57A8;border:1px solid #c7d2e6;padding:7px 12px;
        border-radius:8px;cursor:pointer;font-size:12px;font-weight:600;
        width:100%;margin-top:6px;font-family:Inter,sans-serif;">
        📋 copiar caminho rede
        </button>
        """, unsafe_allow_html=True)

    cols = st.columns(7)
    _intrag_step_card(cols[0], '1', 'Email Itaú', *s1)
    _intrag_step_card(cols[1], '2', 'TXTs gerados', *s2)
    _intrag_step_card(cols[2], '3', 'Passivo Itaú→FIE', s3[0], s3[1], s3[2])
    _intrag_step_card(cols[3], '4', 'Ativo FIE→FIFE', s4[0], s4[1], s4[2])
    _intrag_step_card(cols[4], '5', 'Passivo FIE→FIFE', s5[0], s5[1], s5[2])
    _intrag_step_card(cols[5], '6', 'Liquidação', s6[0], s6[1], s6[2])
    _intrag_step_card(cols[6], '7', 'Arquivo pasta net', *s7)

    if is_dia_util:
        # Checkboxes manuais (steps 3-6, step 7 e auto)
        ck_cols = st.columns(7)
        ck_cols[0].markdown("<div style='font-size:10px;color:#94a3b8;text-align:center'>(auto)</div>", unsafe_allow_html=True)
        ck_cols[1].markdown("<div style='font-size:10px;color:#94a3b8;text-align:center'>(auto)</div>", unsafe_allow_html=True)
        for col, chave, marcado in [
            (ck_cols[2], 'subiu_passivo_itau', s3[3]),
            (ck_cols[3], 'subiu_ativo_fife', s4[3]),
            (ck_cols[4], 'subiu_passivo_fife', s5[3]),
            (ck_cols[5], 'liquidado', s6[3]),
        ]:
            with col:
                novo = st.checkbox('feito', value=marcado, key=f'intrag_{chave}', label_visibility='collapsed')
                if novo != marcado:
                    _intrag_marcar(chave, novo)
                    st.rerun()
        ck_cols[6].markdown("<div style='font-size:10px;color:#94a3b8;text-align:center'>(auto)</div>", unsafe_allow_html=True)


# ── DADOS ────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def get_fundos():
    df = pd.read_excel(TIPO_FUNDOS, usecols="A,E,F,I,J,K,L")
    df = df[df["Encerrado"].isna() & df["modelo_mailer"].notna()].copy()
    df = df[~df["fundo"].isin(["PETROS RFCP"])]
    df["ADM"] = df["ADM"].fillna("Outro")
    df["Tipo"] = "Auto"
    # Site com template: robo gera rascunho, dash mostra como Site
    fundos_site = ["BNYCL12879", "CSHG MAGIS II", "BNY12748", "BNYCL12975",
                   "CAPIT D INC FIC", "PORTFOLIO FIDC", "CAPITANIA PREV BP",
                   "CAPITANIA YIELD 120", "INFRA ADV CLA", "XP INFRA90"]
    df.loc[df["fundo"].isin(fundos_site), "Tipo"] = "Site"
    # Fundos de envio manual e Site sem template
    extras = pd.DataFrame([
        {"fundo": "CAPITANIA FCOPEL","ADM": "Itau",     "Tipo": "Manual"},
        {"fundo": "FCopel_Imob",     "ADM": "Itau",     "Tipo": "Manual"},
        {"fundo": "Sabesprev",       "ADM": "Itau",     "Tipo": "Manual"},
        {"fundo": "CAPITANIA REIT",  "ADM": "BNYM",     "Tipo": "Manual"},
        {"fundo": "PETROS RFCP",     "ADM": "Bradesco", "Tipo": "Manual"},
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

# Banner grande se o robo estiver Parado (> 10 min sem atividade)
if robo_estado == "Parado":
    st.markdown(f"""
    <div style="background:#f8d7da; border-left:5px solid #dc3545;
                padding:14px 18px; border-radius:6px; margin-bottom:16px;">
      <div style="font-size:15px; font-weight:700; color:#721c24;">
        🚨 Robô PARADO - última verificação: {robo_ultima}
      </div>
      <div style="font-size:12px; color:#721c24; margin-top:6px;">
        O robô não está processando emails de aprovação. Verifique se o processo Python esta
        ativo (Gerenciador de Tarefas) ou reinicie via mailer_robo.bat na pasta Startup.
        Um rascunho de alerta deve ter sido criado no seu Outlook pelo watchdog_robo.py.
      </div>
    </div>
    """, unsafe_allow_html=True)

# ── LOG DO ROBÔ ──────────────────────────────────────────────────────────────
if os.path.exists(ROBO_LOG):
    with st.expander("📋 Log do robô"):
        try:
            with open(ROBO_LOG, "r", encoding="utf-8", errors="ignore") as f:
                linhas_log = f.readlines()
            # Últimas 30 linhas, filtrar só as úteis
            ultimas = [l.rstrip() for l in linhas_log[-30:] if l.strip()]
            st.code("\n".join(ultimas), language=None)
        except Exception:
            st.warning("Não foi possível ler o log.")

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
aguardando = {}   # {d_str: {fundo: {"desde": datetime, "motivo": str}}} - cota nao chegou no banco
orfas      = {}   # {d_str: {fundo: {"iniciado": datetime}}} - tentativa sem resultado, requer revisao

for d in dias:
    d_str = d.strftime("%Y%m%d")
    d_ref = ref_de(d).strftime("%Y%m%d")

    def _load(prefix):
        path = os.path.join(JSON_DIR, f"{prefix}_{d_ref}.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    # Carregar aprovados do dia (scan_outlook gera este JSON)
    _aprov_path = os.path.join(JSON_DIR, f"aprovados_{d_ref}.json")
    _aprov_manual = set()
    _aprov_site = set()
    _aprov_manual_erros = {}
    if os.path.exists(_aprov_path):
        with open(_aprov_path, "r", encoding="utf-8") as f:
            _aprov_data = json.load(f)
            # Normaliza nomes internos -> nomes exibidos no dash (ex: FCopel -> CAPITANIA FCOPEL)
            _aprov_manual = set(DASH_DISPLAY_NAME.get(m, m) for m in _aprov_data.get("manual", []))
            _aprov_site = set(_aprov_data.get("site", []))
            _aprov_manual_erros = {
                DASH_DISPLAY_NAME.get(k, k): v
                for k, v in _aprov_data.get("manual_erros", {}).items()
            }
    manuais_aprovados[d_str] = _aprov_manual

    # Carregar aguardando (fundos com cota ausente no banco COTAS_CAP)
    _aguard_raw = _load("aguardando")
    _aguard_dia = {}
    for _f, _info in (_aguard_raw or {}).items():
        try:
            _aguard_dia[_f] = {
                "desde": datetime.fromisoformat(_info["desde"]),
                "motivo": _info.get("motivo", ""),
            }
        except Exception:
            pass
    aguardando[d_str] = _aguard_dia

    # Carregar orfas: tentativa iniciada sem resultado. Precisa revisao humana.
    _tent_raw = _load("tentativas")
    _orfas_dia = {}
    # So e orfa se NAO esta em processados. Processados vem do PDF abaixo (ver bloco else),
    # entao aqui carregamos a lista bruta e filtramos apos.
    for _f, _info in (_tent_raw or {}).items():
        try:
            _orfas_dia[_f] = {"iniciado": datetime.fromisoformat(_info["iniciado"])}
        except Exception:
            pass
    orfas[d_str] = _orfas_dia

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
        # Manual: check apenas se scan_outlook encontrou email "COTA DIARIA"
        for fm in _aprov_manual:
            processados.add(fm)
        # Site sem template: aprovacao no site nao significa envio - so marcar
        # como processado se nao estiver aguardando dado (cota do banco).
        # Aprovacao no site + fundo em aguardando = email ainda NAO foi enviado.
        for fs in _aprov_site:
            if fs not in processados and fs not in _aguard_dia:
                processados.add(fs)
        status[d_str]     = processados
        # Filtrar orfas: se fundo ja esta em processados, nao eh mais orfa
        orfas[d_str] = {f: info for f, info in orfas[d_str].items() if f not in processados}
        erros[d_str]      = _load("erros")
        # Adicionar erros de validação de PDFs manuais
        for fundo_erro, motivo in _aprov_manual_erros.items():
            # Extrair datas para mensagem legível
            import re as _re
            _m = _re.search(r'PDF=(\d{4})(\d{2})(\d{2}).*esperado=(\d{4})(\d{2})(\d{2})', motivo)
            if _m:
                erros[d_str][fundo_erro] = f"Data errada ({_m.group(3)}/{_m.group(2)} ≠ {_m.group(6)}/{_m.group(5)})"
            else:
                erros[d_str][fundo_erro] = f"Data errada: {motivo}"
        horarios[d_str]   = _load("horarios")
        timestamps[d_str] = ts_dia

total = len(fundos)


# ── BANNER: TENTATIVAS ORFAS (REQUER REVISAO HUMANA) ─────────────────────────
_hoje_str = today.strftime("%Y%m%d")
_orfas_hoje = orfas.get(_hoje_str, {})
if _orfas_hoje:
    _linhas_orfas = []
    for _f, _info in sorted(_orfas_hoje.items()):
        _linhas_orfas.append(f"<b>{_f}</b> (iniciada {_info['iniciado'].strftime('%H:%M')})")
    st.markdown(f"""
    <div style="background:#f8d7da; border-left:5px solid #dc3545;
                padding:14px 18px; border-radius:6px; margin-bottom:16px;">
      <div style="font-size:15px; font-weight:700; color:#721c24;">
        🚨 {len(_orfas_hoje)} tentativa(s) ORFA(S) - requerem revisao humana
      </div>
      <div style="font-size:12px; color:#721c24; margin-top:6px;">
        {', '.join(_linhas_orfas)}
      </div>
      <div style="font-size:11px; color:#721c24; margin-top:8px; font-style:italic;">
        O robo comecou a processar esses fundos mas foi interrompido antes de confirmar
        o resultado. Verificar no Outlook se o rascunho foi aberto. Para destravar:
        deletar a entrada em tentativas_{_hoje_str}.json.
      </div>
    </div>
    """, unsafe_allow_html=True)

# ── BANNER: FUNDOS AGUARDANDO COTA NO BANCO ─────────────────────────────────
_aguard_hoje = aguardando.get(_hoje_str, {})
if _aguard_hoje:
    _linhas = []
    for _f, _info in sorted(_aguard_hoje.items()):
        _min = int((datetime.now() - _info["desde"]).total_seconds() / 60)
        _linhas.append(f"<b>{_f}</b> ({_min} min)")
    _lista_html = ", ".join(_linhas)
    _total = len(_aguard_hoje)
    st.markdown(f"""
    <div style="background:#fed7aa; border-left:5px solid #c2410c;
                padding:12px 16px; border-radius:6px; margin-bottom:16px;">
      <div style="font-size:14px; font-weight:700; color:#9a3412; margin-bottom:6px;">
        ⏳ {_total} fundo(s) aguardando cota ser lancada no banco COTAS_CAP
      </div>
      <div style="font-size:12px; color:#7c2d12;">
        {_lista_html}
      </div>
      <div style="font-size:11px; color:#7c2d12; margin-top:6px; font-style:italic;">
        Apos 25 min, o robo cria automaticamente um rascunho de cobranca no Outlook.
      </div>
    </div>
    """, unsafe_allow_html=True)


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


# ── ESTEIRA INTRAG ────────────────────────────────────────────────────────────
render_intrag_esteira()


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
            if fundo in orfas.get(d_str, {}):
                linha[col] = "🚨 ORFA - revisar Outlook"
            elif fundo in MANUAIS_LISTA and fundo in manuais_aprovados.get(d_str, set()):
                linha[col] = "ENVIAR"
            elif fundo in aguardando.get(d_str, {}):
                info_ag = aguardando[d_str][fundo]
                min_decorridos = int((datetime.now() - info_ag["desde"]).total_seconds() / 60)
                linha[col] = f"⏳ aguardando COTAS_CAP {min_decorridos}min"
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
        if v.startswith("🚨"):
            return "background-color:#f8d7da; color:#721c24; text-align:center; font-size:11px; font-weight:700"
        elif v.startswith("⏳"):
            return "background-color:#fed7aa; color:#9a3412; text-align:center; font-size:12px; font-weight:600"
        elif v.startswith("⚠️"):
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
