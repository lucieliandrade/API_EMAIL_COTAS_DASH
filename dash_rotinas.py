import streamlit as st
import json
import os
from datetime import datetime, date
from streamlit_autorefresh import st_autorefresh

# ─────────────────────────────────────────────────────────────────────────────
# Dash de ROTINAS DIARIAS (checklist operacional, D0 = hoje).
# App Streamlit independente do dash de cotas (status_mailers_v2.py). Roda em
# OUTRA porta (8503) com watchdog proprio, entao um nao derruba/estraga o outro.
# So MONITORA: a equipe marca cada item (Pendente / Feito / N/A). Nao roda script.
# Estado de cada dia salvo em JSON (rotinas_<YYYYMMDD>.json) -> historico por data.
# ─────────────────────────────────────────────────────────────────────────────

COR_PRIM = "#1C57A8"

# Pasta de estado: local ao projeto, isolada do dash de cotas. Uma so instancia do
# Streamlit (no servidor) le/grava aqui, entao o estado e compartilhado pela equipe.
ESTADO_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rotinas_estado")

OPCOES = ["Pendente", "Feito", "N/A"]
ICONE_ESTADO = {"Pendente": "⬜", "Feito": "✅", "N/A": "➖"}
DIAS_PT = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira",
           "Sexta-feira", "Sábado", "Domingo"]


# ── CHECKLIST ────────────────────────────────────────────────────────────────
# Ordem dos itens e dos ids e FIXA: os ids viram chave no JSON de estado.
CHECKLIST = [
    {"secao": "DIARIAMENTE", "icone": "🗓️", "itens": [
        {"id": "d1", "txt": "Liquidações e-mail"},
        {"id": "d2", "txt": "Boletar no administrador (resgates ref. e-mail liquidações)"},
        {"id": "d3", "txt": "PL fundos XP"},
        {"id": "d4", "txt": "Cotas fundos listados"},
        {"id": "d5", "txt": "Checar saldos nos FICs para pagamento dos resgates do dia"},
    ]},
    {"secao": "MOVIMENTAÇÕES", "icone": "🔄", "itens": [
        {"id": "m1",  "txt": "ICATU — mandamos tela CETIP?"},
        {"id": "m2",  "txt": "ICATU — BRE boletou?"},
        {"id": "m3",  "txt": "Zerou CREDPREV?"},
        {"id": "m4",  "txt": "PREV XP"},
        {"id": "m5",  "txt": "PREV XP — BRE boletou?"},
        {"id": "m6",  "txt": "REIT PREV XP?"},
        {"id": "m7",  "txt": "Zerou REIT PREV XP?"},
        {"id": "m8",  "txt": "INFLAÇÃO XP?"},
        {"id": "m9",  "txt": "Boletou INFLAÇÃO XP?"},
        {"id": "m10", "txt": "PREVIDENCE XP — mandou para o Zuniga?"},
        {"id": "m11", "txt": "Avisar o Zuniga (movimentos Prev Itaú)"},
        {"id": "m12", "txt": "Outros cotistas na Mellon?"},
        {"id": "m13", "txt": "BTG CAP"},
        {"id": "m14", "txt": "BTG PREV"},
        {"id": "m15", "txt": "BTG SA"},
        {"id": "m16", "txt": "BTG ALTERNATIVES"},
        {"id": "m17", "txt": "BRADESCO?"},
        {"id": "m18", "txt": "PREVs ITAÚ (passivo — Itaú Vida e Prev)", "obs": "Checar se tem R$ na conta"},
        {"id": "m19", "txt": "Zerou ITAÚ FIES × FIFES (Ativo e Passivo)", "obs": "Se aplicação: boletar 1º no ativo e só depois no passivo"},
        {"id": "m20", "txt": "Outros fundos no Itaú?"},
        {"id": "m21", "txt": "Respondeu todas as movimentações liquidadas?"},
        {"id": "m22", "txt": "MELLON — zerou?"},
    ]},
    {"secao": "PRÉVIA", "icone": "📝", "itens": [
        {"id": "p1",  "txt": "MELLON txt", "obs": "Trocou qtd de cotas por financeiro? · Atenção com movimentações entre fundos"},
        {"id": "p2",  "txt": "Zeragens dos FICs nos Masters — MELLON"},
        {"id": "p3",  "txt": "Zeragens dos FICs nos Masters — BRADESCO"},
        {"id": "p4",  "txt": "Zeragens dos FICs nos Masters — ITAÚ (Previs)"},
        {"id": "p5",  "txt": "Zeragens dos FIES nos FIFES — ITAÚ"},
        {"id": "p6",  "txt": "ITAÚ Previs — liquidou"},
        {"id": "p7",  "txt": "Zeragens dos FIES nos FIFES — ITAÚ (liquidou)"},
        {"id": "p8",  "txt": "ITAÚ — Sabesprev"},
        {"id": "p9",  "txt": "ITAÚ — FAPES"},
        {"id": "p10", "txt": "XP Renda 90 (EG)"},
        {"id": "p11", "txt": "Yield 120 (EG)"},
        {"id": "p12", "txt": "Infra Renda Adv (EG)"},
        {"id": "p13", "txt": "Infra Geral Advisory (Feeder)"},
        {"id": "p14", "txt": "BTG Login CAP"},
        {"id": "p15", "txt": "BTG Login PREV"},
        {"id": "p16", "txt": "BTG Login SA"},
        {"id": "p17", "txt": "BTG Login ALTERNATIVES"},
        {"id": "p18", "txt": "BTG/ITAÚ — FCOPEL", "obs": "Atenção"},
    ]},
    {"secao": "NET", "icone": "📊", "itens": [
        {"id": "n1",  "txt": "MELLON txt", "obs": "Trocou qtd de cotas por financeiro? · Atenção com movimentações entre fundos"},
        {"id": "n2",  "txt": "Manual — ITAÚ (todos os fundos)"},
        {"id": "n3",  "txt": "Manual — BTG Login CAP"},
        {"id": "n4",  "txt": "Manual — BTG Login PREV"},
        {"id": "n5",  "txt": "Manual — BTG Login SA"},
        {"id": "n6",  "txt": "Manual — BTG Login ALTERNATIVES"},
        {"id": "n7",  "txt": "Manual — BRADESCO"},
        {"id": "n8",  "txt": "Manual — XP Renda 90 (EG)"},
        {"id": "n9",  "txt": "Manual — Yield 120 (EG)"},
        {"id": "n10", "txt": "Manual — Infra Renda Adv (EG)"},
    ]},
    {"secao": "CONFERÊNCIA FINAL", "icone": "✅", "itens": [
        {"id": "f1", "txt": "Checou TODAS as datas de liquidação? Alguma divergência?", "obs": "ATENÇÃO TOTAL"},
        {"id": "f2", "txt": "Verificar se as movimentações fazem sentido"},
        {"id": "f3", "txt": "Tudo foi liquidado? Se sim, pode enviar por e-mail"},
        {"id": "f4", "txt": "Não deixar campo de data da coluna A em branco"},
    ]},
]

TODOS_IDS = [it["id"] for sec in CHECKLIST for it in sec["itens"]]


# ── ESTADO (JSON por dia) ────────────────────────────────────────────────────
def _estado_path(d_str):
    return os.path.join(ESTADO_DIR, f"rotinas_{d_str}.json")


def carregar_estado(d_str):
    """Le o JSON do dia. {} se nao existir (= tudo Pendente)."""
    p = _estado_path(d_str)
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def salvar_estado(d_str):
    """Grava o estado do dia a partir do que esta na sessao (callback dos radios)."""
    os.makedirs(ESTADO_DIR, exist_ok=True)
    estado = {}
    for _id in TODOS_IDS:
        chave = f"{d_str}::{_id}"
        if chave in st.session_state:
            estado[_id] = st.session_state[chave]
    estado["_atualizado_em"] = datetime.now().isoformat(timespec="seconds")
    with open(_estado_path(d_str), "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)


def reiniciar_dia(d_str):
    for _id in TODOS_IDS:
        st.session_state[f"{d_str}::{_id}"] = "Pendente"
    salvar_estado(d_str)


# ── PAGINA / LOGIN ───────────────────────────────────────────────────────────
st.set_page_config(page_title="RI | Rotinas Diárias", layout="wide", page_icon="🧾")

LOGIN_USER = "RI"
LOGIN_PASS = "Capitania2025!"
AUTO_LOGIN_TOKEN = "ri-dash-local"

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
try:
    if not st.session_state.autenticado and st.query_params.get("k") == AUTO_LOGIN_TOKEN:
        st.session_state.autenticado = True
except Exception:
    pass

if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_left, col_form, col_right = st.columns([1.5, 1, 1.5])
    with col_form:
        st.markdown(f"""
        <div style="text-align:center; margin-bottom:24px;">
            <span style="font-size:36px;">🧾</span>
            <h2 style="color:{COR_PRIM}; margin:8px 0 4px;">Rotinas Diárias</h2>
            <p style="color:#64748b; font-size:13px;">Capitânia Investimentos</p>
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

# Auto-refresh a cada 2 minutos (espelha mudancas de outras pessoas da equipe).
st_autorefresh(interval=120_000, key="autorefresh")


# ── CSS (mesmo visual do dash de cotas) ──────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

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
.header-sub   { color: rgba(255,255,255,0.75); font-size: 13px; margin: 0; }
.header-day   { color: white; font-size: 15px; font-weight: 600; text-align: right; }

.sec-title { font-size: 15px; font-weight: 700; color: #1a2540; letter-spacing:.02em; }
.sec-meta  { font-size: 12px; color: #64748b; font-weight: 600; }

.item-txt  { font-size: 14px; color: #1a2540; font-weight: 500; }
.item-txt.feito { color:#94a3b8; text-decoration: line-through; }
.item-txt.na    { color:#94a3b8; }
.item-obs  { font-size: 11px; color:#b45309; font-style: italic; margin-top:1px; }

.progress-bg   { background:#f1f5f9; border-radius:99px; height:10px; overflow:hidden; margin:6px 0 2px; }
.progress-fill { height:10px; border-radius:99px; background: linear-gradient(90deg,#22c55e,#16a34a); transition: width .4s ease; }

div[data-testid="stRadio"] label p { font-size: 12px; }
</style>
""", unsafe_allow_html=True)


# ── SIDEBAR: dia + reiniciar ─────────────────────────────────────────────────
hoje = date.today()
with st.sidebar:
    st.markdown(f"### 🧾 Rotinas Diárias")
    dia_sel = st.date_input("Dia", value=hoje, max_value=hoje, format="DD/MM/YYYY")
    if dia_sel != hoje:
        st.warning("Visualizando um dia anterior (histórico).")
    st.caption("D0 = hoje. Cada dia começa zerado; o estado fica salvo por data.")
    st.divider()
    if st.button("🔄 Reiniciar este dia", use_container_width=True):
        reiniciar_dia(dia_sel.strftime("%Y%m%d"))
        st.rerun()

d_str = dia_sel.strftime("%Y%m%d")
estado = carregar_estado(d_str)

# Inicializa a sessao a partir do JSON salvo (so na 1a vez de cada chave/dia).
for _id in TODOS_IDS:
    chave = f"{d_str}::{_id}"
    if chave not in st.session_state:
        st.session_state[chave] = estado.get(_id, "Pendente")


def _val(_id):
    return st.session_state.get(f"{d_str}::{_id}", "Pendente")


# ── PROGRESSO GERAL ──────────────────────────────────────────────────────────
n_total   = len(TODOS_IDS)
n_feito   = sum(1 for i in TODOS_IDS if _val(i) == "Feito")
n_na      = sum(1 for i in TODOS_IDS if _val(i) == "N/A")
n_pend    = n_total - n_feito - n_na
aplicaveis = n_total - n_na
pct = int(round(100 * n_feito / aplicaveis)) if aplicaveis else 100

ts_label = ""
if isinstance(estado.get("_atualizado_em"), str):
    try:
        ts_label = "atualizado " + datetime.fromisoformat(estado["_atualizado_em"]).strftime("%H:%M")
    except Exception:
        ts_label = ""

st.markdown(f"""
<div class="header-bar">
  <div>
    <p class="header-title">🧾 Rotinas Diárias</p>
    <p class="header-sub">Capitânia Investimentos · Relações com Investidores</p>
  </div>
  <div class="header-day">
    {'Hoje' if dia_sel == hoje else 'Dia'}<br>
    {DIAS_PT[dia_sel.weekday()]}, {dia_sel.strftime('%d/%m/%Y')}
  </div>
</div>
""", unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)
c1.metric("✅ Feito", n_feito)
c2.metric("⬜ Pendente", n_pend)
c3.metric("➖ N/A", n_na)
c4.metric("Concluído", f"{pct}%", help="Feitos ÷ itens aplicáveis (exclui os N/A)")

st.markdown(f"""
<div class="progress-bg"><div class="progress-fill" style="width:{pct}%;"></div></div>
<div style="text-align:right; font-size:11px; color:#94a3b8;">{ts_label}</div>
""", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)


# ── SECOES / ITENS ───────────────────────────────────────────────────────────
def render_item(it):
    _id = it["id"]
    estado_atual = _val(_id)
    cls = {"Feito": "feito", "N/A": "na"}.get(estado_atual, "")
    col_txt, col_radio = st.columns([6, 4])
    with col_txt:
        st.markdown(
            f"<div class='item-txt {cls}'>{ICONE_ESTADO[estado_atual]} {it['txt']}</div>"
            + (f"<div class='item-obs'>⚠️ {it['obs']}</div>" if it.get("obs") else ""),
            unsafe_allow_html=True,
        )
    with col_radio:
        st.radio(
            it["txt"], OPCOES,
            key=f"{d_str}::{_id}",
            horizontal=True,
            label_visibility="collapsed",
            on_change=salvar_estado, args=(d_str,),
        )


for sec in CHECKLIST:
    ids = [it["id"] for it in sec["itens"]]
    s_feito = sum(1 for i in ids if _val(i) == "Feito")
    s_na    = sum(1 for i in ids if _val(i) == "N/A")
    s_aplic = len(ids) - s_na
    completa = s_aplic > 0 and s_feito == s_aplic
    selo = "✔️" if completa else f"{s_feito}/{s_aplic}"
    with st.container(border=True):
        st.markdown(
            f"<span class='sec-title'>{sec['icone']} {sec['secao']}</span> "
            f"&nbsp;<span class='sec-meta'>· {selo} feitos"
            + (f" · {s_na} N/A" if s_na else "") + "</span>",
            unsafe_allow_html=True,
        )
        for it in sec["itens"]:
            render_item(it)

st.caption("As marcações são salvas automaticamente. Auto-atualiza a cada 2 min.")
