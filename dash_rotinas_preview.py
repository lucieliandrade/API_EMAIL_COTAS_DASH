import streamlit as st
from datetime import date
from rotinas_checklist import CHECKLIST, TODOS_IDS, OPCOES, ICONE_ESTADO, DIAS_PT

# ─────────────────────────────────────────────────────────────────────────────
# PREVIEW de layouts do dash de rotinas. NAO salva nada (estado so em memoria).
# Serve so para a usuaria escolher o visual. Roda em porta separada (nao toca
# no dash de cotas 8502 nem no dash de rotinas 8503).
# ─────────────────────────────────────────────────────────────────────────────

COR_PRIM = "#1C57A8"
COR_ESTADO = {"Pendente": ("#fef3c7", "#92400e"),
              "Feito":    ("#dcfce7", "#15803d"),
              "N/A":      ("#e2e8f0", "#475569")}
NEXT = {"Pendente": "Feito", "Feito": "N/A", "N/A": "Pendente"}

st.set_page_config(page_title="Preview · Rotinas", layout="wide", page_icon="🎨")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.header-bar { background: linear-gradient(90deg,#1C57A8 0%,#2e7dd1 100%);
  border-radius:12px; padding:16px 26px; display:flex; align-items:center;
  justify-content:space-between; margin-bottom:16px; }
.header-title { color:white; font-size:21px; font-weight:700; margin:0; }
.header-sub { color:rgba(255,255,255,.75); font-size:13px; margin:0; }
.header-day { color:white; font-size:14px; font-weight:600; text-align:right; }
.sec-title { font-size:15px; font-weight:700; color:#1a2540; }
.sec-meta { font-size:12px; color:#64748b; font-weight:600; }
.item-txt { font-size:14px; color:#1a2540; font-weight:500; }
.item-txt.feito { color:#94a3b8; text-decoration:line-through; }
.item-txt.na { color:#94a3b8; }
.item-obs { font-size:11px; color:#b45309; font-style:italic; }
.progress-bg { background:#f1f5f9; border-radius:99px; height:10px; overflow:hidden; margin:6px 0; }
.progress-fill { height:10px; border-radius:99px; background:linear-gradient(90deg,#22c55e,#16a34a); }
.pill { display:inline-block; padding:3px 10px; border-radius:99px; font-size:12px; font-weight:700; }
div[data-testid="stRadio"] label p { font-size:12px; }
/* deixa os botoes de status mais justos */
div[data-testid="column"] button { padding:2px 6px; }
</style>
""", unsafe_allow_html=True)

# ── Estado em memoria (sem persistencia), com amostra inicial para visualizar ──
if "pv_init" not in st.session_state:
    amostra_feito = {"d1", "d3", "m1", "m4", "m13", "m21", "p1", "p8", "p9", "n1", "n7"}
    amostra_na    = {"m6", "m17", "p13", "n9"}
    for _id in TODOS_IDS:
        st.session_state[f"pv::{_id}"] = (
            "Feito" if _id in amostra_feito else "N/A" if _id in amostra_na else "Pendente"
        )
    st.session_state["pv_init"] = True


def val(_id):
    return st.session_state.get(f"pv::{_id}", "Pendente")


def set_radio(_id):
    st.session_state[f"pv::{_id}"] = st.session_state[f"r::{_id}"]


def cycle(_id):
    st.session_state[f"pv::{_id}"] = NEXT[val(_id)]


def contagem(ids):
    feito = sum(1 for i in ids if val(i) == "Feito")
    na    = sum(1 for i in ids if val(i) == "N/A")
    aplic = len(ids) - na
    return feito, na, aplic


def progresso_geral():
    feito, na, aplic = contagem(TODOS_IDS)
    pct = int(round(100 * feito / aplic)) if aplic else 100
    pend = len(TODOS_IDS) - feito - na
    return feito, pend, na, pct


# ── Cabecalho + seletor de layout ────────────────────────────────────────────
hoje = date.today()
feito, pend, na, pct = progresso_geral()
st.markdown(f"""
<div class="header-bar">
  <div><p class="header-title">🧾 Rotinas Diárias</p>
  <p class="header-sub">Capitânia · Relações com Investidores · PREVIEW de layout (não salva)</p></div>
  <div class="header-day">Hoje<br>{DIAS_PT[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')}</div>
</div>
""", unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)
c1.metric("✅ Feito", feito)
c2.metric("⬜ Pendente", pend)
c3.metric("➖ N/A", na)
c4.metric("Concluído", f"{pct}%")
st.markdown(f"<div class='progress-bg'><div class='progress-fill' style='width:{pct}%;'></div></div>",
            unsafe_allow_html=True)

layout = st.radio(
    "Layout para visualizar",
    ["1 · Abas + 2 colunas (radio compacto)",
     "2 · Tela única, clique pra alternar",
     "3 · Abas + clique pra alternar"],
    horizontal=True,
)
filtro = st.pills("Mostrar", OPCOES, selection_mode="multi", default=OPCOES,
                  format_func=lambda o: f"{ICONE_ESTADO[o]} {o}", key="pv_filtro")
filtro_efetivo = filtro if filtro else OPCOES  # nada selecionado = mostra tudo
st.caption("Marque à vontade — é só preview, nada é salvo. Troque o layout/filtro acima.")
st.divider()


def filtra(itens):
    return [it for it in itens if val(it["id"]) in filtro_efetivo]


# ── Componentes de item ──────────────────────────────────────────────────────
def item_radio(it):
    """Item com radio compacto (Pendente/Feito/N/A)."""
    _id = it["id"]; e = val(_id)
    cls = {"Feito": "feito", "N/A": "na"}.get(e, "")
    ct, cr = st.columns([6, 4])
    with ct:
        st.markdown(f"<div class='item-txt {cls}'>{ICONE_ESTADO[e]} {it['txt']}</div>"
                    + (f"<div class='item-obs'>⚠️ {it['obs']}</div>" if it.get('obs') else ""),
                    unsafe_allow_html=True)
    with cr:
        st.radio(it["txt"], OPCOES, index=OPCOES.index(e), key=f"r::{_id}",
                 horizontal=True, label_visibility="collapsed", on_change=set_radio, args=(_id,))


def item_cycle(it):
    """Item com UM botao de status que alterna no clique."""
    _id = it["id"]; e = val(_id)
    cls = {"Feito": "feito", "N/A": "na"}.get(e, "")
    cb, ct = st.columns([1.4, 8.6])
    with cb:
        st.button(f"{ICONE_ESTADO[e]} {e}", key=f"b::{_id}", use_container_width=True,
                  on_click=cycle, args=(_id,),
                  type=("primary" if e == "Pendente" else "secondary"))
    with ct:
        st.markdown(f"<div class='item-txt {cls}' style='padding-top:6px'>{it['txt']}</div>"
                    + (f"<div class='item-obs'>⚠️ {it['obs']}</div>" if it.get('obs') else ""),
                    unsafe_allow_html=True)


def cabecalho_secao(sec):
    feito, na, aplic = contagem([it["id"] for it in sec["itens"]])
    selo = "✔️" if (aplic > 0 and feito == aplic) else f"{feito}/{aplic}"
    st.markdown(f"<span class='sec-title'>{sec['icone']} {sec['secao']}</span> "
                f"&nbsp;<span class='sec-meta'>· {selo} feitos"
                + (f" · {na} N/A" if na else "") + "</span>", unsafe_allow_html=True)


def label_aba(sec):
    feito, na, aplic = contagem([it["id"] for it in sec["itens"]])
    return f"{sec['secao']} ({feito}/{aplic})"


if filtro and len(filtro) < len(OPCOES):
    st.caption(f"🔎 Filtrando: {', '.join(filtro_efetivo)}")

# ── LAYOUT 1: abas + 2 colunas com radio compacto ───────────────────────────
if layout.startswith("1"):
    abas = st.tabs([label_aba(s) for s in CHECKLIST])
    for aba, sec in zip(abas, CHECKLIST):
        with aba:
            cabecalho_secao(sec)
            vis = filtra(sec["itens"])
            if not vis:
                st.info("Nenhum item no filtro selecionado.")
                continue
            colA, colB = st.columns(2)
            for idx, it in enumerate(vis):
                with (colA if idx % 2 == 0 else colB):
                    with st.container(border=True):
                        item_radio(it)

# ── LAYOUT 2: tela unica, clique pra alternar ────────────────────────────────
elif layout.startswith("2"):
    st.info("Clique no botão de status (esquerda) para alternar: Pendente → Feito → N/A.")
    algum = False
    for sec in CHECKLIST:
        vis = filtra(sec["itens"])
        if not vis:
            continue
        algum = True
        with st.container(border=True):
            cabecalho_secao(sec)
            for it in vis:
                item_cycle(it)
    if not algum:
        st.info("Nenhum item no filtro selecionado.")

# ── LAYOUT 3: abas + clique pra alternar ─────────────────────────────────────
else:
    abas = st.tabs([label_aba(s) for s in CHECKLIST])
    for aba, sec in zip(abas, CHECKLIST):
        with aba:
            cabecalho_secao(sec)
            st.caption("Clique no status para alternar: Pendente → Feito → N/A.")
            vis = filtra(sec["itens"])
            if not vis:
                st.info("Nenhum item no filtro selecionado.")
                continue
            for it in vis:
                item_cycle(it)
