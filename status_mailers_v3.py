r"""
===============================================================================
 DASH DE COTAS DIÁRIAS  ·  Capitânia Investimentos / Relações com Investidores
===============================================================================

 O QUE É ISTO?
   Um painel web (Streamlit) que mostra o status do envio diário das cotas dos
   fundos. Ele NÃO envia e-mail e NÃO decide nada — funciona como o painel de um
   carro: apenas LÊ arquivos deixados pelos robôs e pinta a tela de verde/
   amarelo/vermelho. Quem faz o trabalho são outros programas:
     - scan_outlook.py  -> lê o Outlook e gera os .json de aprovações/erros
     - mailer_robo.py   -> processa as cotas e gera os PDFs
     - watchdog_robo.py -> vigia se o robô principal travou
   Se um robô para, o dash mostra "Parado", mas ele mesmo continua funcionando.

 COMO RODA
   auto_dash.bat  ->  streamlit run  ->  http://192.168.3.81:8502  (login RI)
   Recarrega sozinho a cada 2 min (st_autorefresh).

 MAPA DO CÓDIGO (procure por estes marcadores "── ... ──" ao navegar):
   LOGIN ................. tela de usuário/senha + auto-login por token (?k=)
   CSS .................. estilos visuais dos cards e da tabela
   INTRAG .............. esteira de 7 steps das boletas Itaú Vida
   ENVIO DIÁRIO ........ XMLs Mellon (ICATU / Aquila / BASF)
   DADOS ............... get_fundos(): lê Tipo_Fundos.xlsx (lista de fundos)
   SEMANA .............. calcula os 5 dias e o calendário de feriados B3/BVMF
   HEADER .............. cabeçalho azul
   STATUS DO ROBÔ ...... a "bolinha" (verde/amarelo/vermelho) + banner se parado
   CARREGAR JSONs ...... lê PDFs e .json de cada dia -> monta o status de cada fundo
   BANNERS ............. alertas de órfãs e de fundos aguardando cota (COTAS_CAP)
   CARDS DOS DIAS ...... os 5 cartões de resumo da semana
   FILTROS / TABELA .... a grade fundo × dia (coração do dash)
   PENDENTES HOJE ...... lista os fundos que ainda não saíram

 CONCEITO QUE MAIS CONFUNDE — a data de referência (D-1):
   A COLUNA é o dia do ENVIO, mas a COTA é sempre do dia útil ANTERIOR.
   Ex.: coluna Quarta 01/04 -> procura o arquivo de 31/03 (terça).
        coluna Segunda     -> procura o de sexta (pula o fim de semana).
   Por isso usamos o calendário B3/BVMF: ele pula feriados e fins de semana
   corretamente ao calcular esse D-1 (função ref_de()).

 TIPOS DE FUNDO (coluna "Tipo"): muda como o dash confirma o envio:
   Auto   -> confirma sozinho pelo PDF que aparece na pasta de PDFs
   Site   -> confirma pela aprovação registrada pelo scan_outlook
   Manual -> alguém envia na mão; aparece "ENVIAR" quando o e-mail é aprovado

 ONDE FICAM AS FONTES DE DADOS (se a pasta cair, a seção fica vazia — não é bug):
   Lista de fundos ....... X:\...\Tipo_Fundos.xlsx
   Aprovações/erros ...... Z:\...\cotas\json\*.json   (scan_outlook)
   PDFs das cotas ........ Z:\...\cotas\PDFs\         (mailer_robo)
   Log do robô ........... robo_log.txt (ao lado deste arquivo)

 >>> MANUAL COMPLETO, em linguagem para qualquer pessoa: documentacao/MANUAL_DASH_COTAS.md <<<
===============================================================================
"""

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
# Scripts/runtime do robo ficam em _robo_automatico\ (TXTs gerados continuam na raiz)
INTRAG_ROBO_DIR = os.path.join(INTRAG_PASTA, "_robo_automatico")
INTRAG_HEARTBEAT = os.path.join(INTRAG_ROBO_DIR, "agendador_heartbeat.txt")
INTRAG_PROCESSADOS = os.path.join(INTRAG_ROBO_DIR, "processados_intrag.txt")
INTRAG_ESTADO_MANUAL = os.path.join(INTRAG_ROBO_DIR, "esteira_estado.json")
INTRAG_PASTA_NET = r"N:\Middle\Resgates\Codigos_movimentacoes_adm\Código Itaú"
# Destino do encaminhamento do e-mail do Itau (step 2 - Email Zuniga)
ZUNIGA_DEST = "lucieli.andrade@capitania.net"
# Remetente do encaminhamento: sempre o e-mail generico compartilhado (os 3 usam),
# nao a conta pessoal de quem disparou. Deixa o remetente uniforme para o Zuniga.
ZUNIGA_FROM = "invest@capitaniainvestimentos.com.br"

# Envio Diário (XMLs Mellon para ICATU / Aquila / BASF)
ENVIO_DIARIO_PASTA_XML = r"X:\RI + BACK - PILOTO XML\Mellon_API_Diariamente(RI)"
ENVIO_DIARIO_PASTA_CARTEIRAS = r"X:\#CapitaniaRFE\Operational\BatimentoCotas\Carteiras_BNYM"
ENVIO_DIARIO_DIR = r"Z:\Relações com Investidores - NOVO\codigos\Envio_Diário"
ENVIO_DIARIO_LOG = os.path.join(ENVIO_DIARIO_DIR, "enviados.json")

ENVIO_DIARIO_INVEST = "invest@capitaniainvestimentos.com.br"
ENVIO_DIARIO_CLIENTES = {
    "ICATU": {
        "codigos": ["FD26498249000162", "FD27239065000140", "FD30338838000150"],
        "assunto": "ICATU | XML - {data}",
        "to":  ENVIO_DIARIO_INVEST,
        "cc":  ENVIO_DIARIO_INVEST,
        "bcc": "ControledeInvestimentos@icatuseguros.com.br",
    },
    "Aquila 6 e 7": {
        "codigos": ["FD17898668000109", "FD42870959000128"],
        "assunto": "Aquila 6 e 7 | CARTEIRA - {data}",
        "to":  ENVIO_DIARIO_INVEST,
        "cc":  ENVIO_DIARIO_INVEST,
        "bcc": "Claudia.vieira@linde.com;Isadora.monteiro@linde.com;Wilson.bispo@linde.com",
        "extras": [
            os.path.join(ENVIO_DIARIO_PASTA_CARTEIRAS, "CAPIT AQUILA 6_{data}.xlsx"),
            os.path.join(ENVIO_DIARIO_PASTA_CARTEIRAS, "BNY11585_{data}.xlsx"),
        ],
    },
    "BASF": {
        "codigos": ["FD18447898000106", "FD21732670000172"],
        "assunto": "BASF | XML - {data}",
        "to":  ENVIO_DIARIO_INVEST,
        "cc":  ENVIO_DIARIO_INVEST,
        "bcc": "gabriel.a.silva@basf.com",
    },
}

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


st.set_page_config(page_title="RI | Dash Cotas (v2 nova)", layout="wide", page_icon="🆕")

# ── LOGIN ───────────────────────────────────────────────────────────────────
LOGIN_USER = "RI"
LOGIN_PASS = "Capitania2025!"

# Token de auto-login: mantem o MESMO login/senha (RI/Capitania2025!), mas permite
# que a janela dedicada do dash (keepalive_dash_chrome, na maquina servidora) entre
# direto sem redigitar. Outras maquinas continuam com o login normal por formulario.
AUTO_LOGIN_TOKEN = "ri-dash-local"

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

# Auto-login somente quando a URL traz o token correto (usado so pela janela dedicada).
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


def _intrag_encaminhar_zuniga():
    """Abre no Outlook o ENCAMINHAMENTO do e-mail do Itaú (assunto INSTRU+CAPITANIA,
    com PDF anexo, recebido hoje na Caixa de Entrada) para ZUNIGA_DEST.
    Usa o MESMO critério que o robô usa para achar o e-mail. Mantém o anexo original
    e NÃO altera nada — só encaminha. Display (NUNCA Send): abre como rascunho para
    revisar e enviar manualmente. Retorna (True, None) ou (False, motivo)."""
    try:
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        ns = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
        inbox = ns.GetDefaultFolder(6)
        hoje = datetime.now().strftime('%m/%d/%Y')
        try:
            itens = inbox.Items.Restrict(f"[ReceivedTime] >= '{hoje} 00:00'")
            itens.Sort("[ReceivedTime]", True)
        except Exception:
            itens = inbox.Items
        alvo = None
        for mail in itens:
            try:
                au = (getattr(mail, 'Subject', '') or '').upper()
                if au.startswith('RE:') or au.startswith('RES:'):
                    continue
                if 'INSTRU' not in au or 'CAPITANIA' not in au:
                    continue
                tem_pdf = any(
                    str(mail.Attachments.Item(i).FileName).lower().endswith('.pdf')
                    for i in range(1, mail.Attachments.Count + 1)
                )
                if not tem_pdf:
                    continue
                alvo = mail
                break
            except Exception:
                continue
        if alvo is None:
            return False, ("e-mail do Itaú (INSTRUÇÃO CAPITANIA com PDF) não encontrado "
                           "na Caixa de Entrada de hoje")
        fwd = alvo.Forward()
        fwd.To = ZUNIGA_DEST
        try:
            fwd.SentOnBehalfOfName = ZUNIGA_FROM  # remetente = invest@ (generico), nao conta pessoal
        except Exception:
            pass
        fwd.Display()
        return True, None
    except Exception as e:
        return False, str(e)


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
    """Renderiza a esteira INTRAG de 8 steps."""
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

    # Step 2 - Email Zuniga (manual): apos receber o email/PDF do Itau, encaminhar
    # para lucieli.andrade@capitania.net (so encaminhar, copiando esse endereco).
    # So fica pendente depois que o email do Itau chega (step 1 = sucesso).
    zi = manual.get('email_zuniga')
    if not is_dia_util:
        sz = ('fut', '🏖️', '-', False)
    elif zi and zi.get('feito'):
        sz = ('ok', '✅', f"feito {zi.get('ts', '')[:5]}", True)
    elif proc and proc['tipo'] == 'sem_movimento':
        sz = ('ok', '🏖️', 'sem mov', False)
    elif proc and proc['tipo'] == 'sucesso':
        sz = ('pend', '📧', 'encaminhar', False)
    else:
        sz = ('fut', '⏳', 'aguarda email', False)

    # Step 3 - TXTs gerados
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

    # Resumo no titulo do expander
    todos_steps = [s1, sz, s2, s3, s4, s5, s6, s7]
    n_ok = sum(1 for s in todos_steps if s[0] == 'ok')
    n_pend = sum(1 for s in todos_steps if s[0] == 'pend')
    if not is_dia_util:
        resumo = "🏖️ fim de semana"
    elif n_ok == 8:
        resumo = "✅ 8/8 etapas concluídas"
    elif n_pend > 0 or n_ok > 0:
        resumo = f"⏳ {n_ok}/8 etapas concluídas"
    else:
        resumo = "⏰ aguardando início"

    label_expander = f"🏦 Esteira INTRAG · Boletas Itaú Vida   —   {resumo}"

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander(label_expander, expanded=False):
        btn_pasta_col, code_col = st.columns([1, 5])
        with btn_pasta_col:
            if st.button("📁 abrir pasta", key="intrag_abrir_pasta", use_container_width=True, help=INTRAG_PASTA_NET):
                try:
                    os.startfile(INTRAG_PASTA_NET)
                except Exception as e:
                    st.warning(f"Falha ao abrir: {e}")
        with code_col:
            st.code(INTRAG_PASTA_NET, language=None)

        cols = st.columns(8)
        _intrag_step_card(cols[0], '1', 'Email Itaú', *s1)
        _intrag_step_card(cols[1], '2', 'Email Zúñiga', sz[0], sz[1], sz[2])
        _intrag_step_card(cols[2], '3', 'TXTs gerados', *s2)
        _intrag_step_card(cols[3], '4', 'Passivo Itaú→FIE', s3[0], s3[1], s3[2])
        _intrag_step_card(cols[4], '5', 'Ativo FIE→FIFE', s4[0], s4[1], s4[2])
        _intrag_step_card(cols[5], '6', 'Passivo FIE→FIFE', s5[0], s5[1], s5[2])
        _intrag_step_card(cols[6], '7', 'Liquidação', s6[0], s6[1], s6[2])
        _intrag_step_card(cols[7], '8', 'Arquivo pasta net', *s7)

        # Step 2 (Email Zúñiga): botão para abrir o rascunho de ENCAMINHAMENTO do
        # e-mail do Itaú (com o PDF) para ZUNIGA_DEST. Só aparece depois que o e-mail
        # chegou (proc=sucesso) ou o step já foi marcado (permite reabrir). Display.
        if is_dia_util and ((proc and proc.get('tipo') == 'sucesso') or sz[3]):
            fcol1, fcol2 = st.columns([2, 4])
            with fcol1:
                if st.button("📧 Encaminhar Itaú → Zúñiga", key="intrag_fwd_zuniga",
                             use_container_width=True,
                             help=f"Abre o encaminhamento do e-mail do Itaú (com o PDF) para "
                                  f"{ZUNIGA_DEST} no Outlook DO SERVIDOR (máquina da Lucieli), "
                                  f"não na sua tela. Quem envia precisa estar no servidor."):
                    ok_fwd, msg_fwd = _intrag_encaminhar_zuniga()
                    if ok_fwd:
                        st.toast("Rascunho de encaminhamento aberto no Outlook", icon="✉️")
                    else:
                        st.error(f"Rascunho NÃO aberto: {msg_fwd}")
            with fcol2:
                st.markdown(
                    f"<div style='padding-top:8px;font-size:11px;color:#94a3b8'>"
                    f"Encaminha o e-mail do Itaú (assunto INSTRUÇÃO CAPITANIA, com PDF) para "
                    f"<code style='font-size:10px'>{ZUNIGA_DEST}</code> — abre como rascunho "
                    f"<b>no Outlook do servidor</b> (máquina da Lucieli), não na tela de quem "
                    f"clicou. Só 1 pessoa encaminha; revise, envie e marque ✓ no step 2.</div>",
                    unsafe_allow_html=True)

        if is_dia_util:
            hoje_iso = agora.date().isoformat()
            ck_cols = st.columns(8)
            ck_cols[0].markdown("<div style='font-size:10px;color:#94a3b8;text-align:center'>(auto)</div>", unsafe_allow_html=True)
            ck_cols[2].markdown("<div style='font-size:10px;color:#94a3b8;text-align:center'>(auto)</div>", unsafe_allow_html=True)
            for col, chave, marcado in [
                (ck_cols[1], 'email_zuniga', sz[3]),
                (ck_cols[3], 'subiu_passivo_itau', s3[3]),
                (ck_cols[4], 'subiu_ativo_fife', s4[3]),
                (ck_cols[5], 'subiu_passivo_fife', s5[3]),
                (ck_cols[6], 'liquidado', s6[3]),
            ]:
                with col:
                    # key inclui data: evita session_state do dia anterior re-marcar no dia novo
                    wkey = f'intrag_{hoje_iso}_{chave}'
                    skey = f'_seen_{wkey}'
                    # Sincroniza com o JSON compartilhado (Z:\...\esteira_estado.json).
                    # Se OUTRA pessoa do time mudou o estado no arquivo desde a ultima vez
                    # que ESTA sessao/aba viu, forca o checkbox a refletir o disco. Sem isso,
                    # o Streamlit ignora o value= apos o primeiro render e o visual congela
                    # por sessao — o dado sincroniza, mas o check nao aparece para os outros.
                    # marcado == skey => JSON nao mudou do nosso ponto de vista, entao uma
                    # diferenca no widget e o clique da propria pessoa: nao sobrescrever.
                    if st.session_state.get(skey) != marcado:
                        st.session_state[wkey] = marcado
                        st.session_state[skey] = marcado
                    novo = st.checkbox('feito', key=wkey, label_visibility='collapsed')
                    if novo != marcado:
                        _intrag_marcar(chave, novo)
                        st.session_state[skey] = novo
                        st.rerun()
            ck_cols[7].markdown("<div style='font-size:10px;color:#94a3b8;text-align:center'>(auto)</div>", unsafe_allow_html=True)


# ── ENVIO DIÁRIO (XMLs Mellon: ICATU / Aquila / BASF) ────────────────────────
def _envio_data_default():
    """D-1 útil (pula fim de semana e feriado de mercado B3/BVMF). Default ao abrir o dash.
    Usa o calendario B3 (nacionais + Corpus Christi, Carnaval e Sexta-feira Santa)."""
    anos = [datetime.now().year - 1, datetime.now().year]
    br_holidays = holidays.financial_holidays('BVMF', years=anos)
    d = datetime.now().date() - timedelta(days=1)
    while d.weekday() >= 5 or d in br_holidays:
        d -= timedelta(days=1)
    return d


@st.cache_data(ttl=60, show_spinner=False)
def _envio_listar_pasta_xml():
    """Lista arquivos da pasta Mellon (em rede X:) com cache de 60s."""
    if not os.path.isdir(ENVIO_DIARIO_PASTA_XML):
        return None  # pasta off
    try:
        return tuple(os.listdir(ENVIO_DIARIO_PASTA_XML))
    except Exception:
        return None


@st.cache_data(ttl=60, show_spinner=False)
def _envio_arquivo_existe(caminho):
    """os.path.exists com cache de 60s — para extras de rede."""
    return os.path.exists(caminho)


def _envio_buscar_arquivos(cliente_cfg, data_yyyymmdd):
    """Para um cliente, retorna {
        'xmls_ok':[caminho], 'xmls_falt':[codigo],
        'extras_ok':[caminho], 'extras_falt':[nome],
    }"""
    xmls_ok, xmls_falt = [], []
    extras_ok, extras_falt = [], []

    arquivos_pasta = _envio_listar_pasta_xml()
    if arquivos_pasta is None:
        return {'xmls_ok': [], 'xmls_falt': cliente_cfg['codigos'],
                'extras_ok': [], 'extras_falt': [], 'pasta_off': True}

    for cod in cliente_cfg['codigos']:
        padrao = f"{cod}_{data_yyyymmdd}"
        achou = next((a for a in arquivos_pasta if a.startswith(padrao)), None)
        if achou:
            xmls_ok.append(os.path.join(ENVIO_DIARIO_PASTA_XML, achou))
        else:
            xmls_falt.append(cod)

    for tpl in cliente_cfg.get('extras', []):
        caminho = tpl.format(data=data_yyyymmdd)
        if _envio_arquivo_existe(caminho):
            extras_ok.append(caminho)
        else:
            extras_falt.append(os.path.basename(caminho))

    return {'xmls_ok': xmls_ok, 'xmls_falt': xmls_falt,
            'extras_ok': extras_ok, 'extras_falt': extras_falt, 'pasta_off': False}


@st.cache_data(ttl=60, show_spinner=False)
def _envio_log_ler():
    if not os.path.exists(ENVIO_DIARIO_LOG):
        return {}
    try:
        with open(ENVIO_DIARIO_LOG, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _envio_log_marcar(data_yyyymmdd, cliente):
    # Le DIRETO do disco (nao do cache de 60s) para nao apagar a marcacao que
    # outra pessoa do time acabou de gravar no mesmo arquivo compartilhado.
    todo = {}
    if os.path.exists(ENVIO_DIARIO_LOG):
        try:
            with open(ENVIO_DIARIO_LOG, 'r', encoding='utf-8') as f:
                todo = json.load(f)
        except Exception:
            todo = {}
    todo.setdefault(data_yyyymmdd, {})
    todo[data_yyyymmdd][cliente] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        os.makedirs(ENVIO_DIARIO_DIR, exist_ok=True)
        with open(ENVIO_DIARIO_LOG, 'w', encoding='utf-8') as f:
            json.dump(todo, f, indent=2, ensure_ascii=False)
        _envio_log_ler.clear()
    except Exception as e:
        st.warning(f"Falha ao salvar log de envio: {e}")


def _envio_log_desmarcar(data_yyyymmdd, cliente):
    todo = dict(_envio_log_ler())
    if data_yyyymmdd in todo and cliente in todo[data_yyyymmdd]:
        todo[data_yyyymmdd] = dict(todo[data_yyyymmdd])
        todo[data_yyyymmdd].pop(cliente)
        if not todo[data_yyyymmdd]:
            todo.pop(data_yyyymmdd)
        try:
            with open(ENVIO_DIARIO_LOG, 'w', encoding='utf-8') as f:
                json.dump(todo, f, indent=2, ensure_ascii=False)
            _envio_log_ler.clear()
        except Exception as e:
            st.warning(f"Falha ao atualizar log: {e}")


def _envio_saudacao():
    h = datetime.now().hour
    if h < 12:  return "bom dia,"
    if h < 18:  return "boa tarde,"
    return "boa noite,"


def _envio_corpo_html():
    saud = _envio_saudacao()
    return f"""<html><body style="font-family: Verdana; font-size: 10pt; color: black;">
<p>Prezados, {saud}</p>
<p>Seguem os arquivos.</p>
<p>Atenciosamente,</p>
<p style="color: #1B51A3;">
Relações com Investidores<br>
Capitânia Investimentos<br>
Tel: 55-11-2853-8888<br>
<a href="mailto:invest@capitaniainvestimentos.com.br" style="color:#1B51A3;text-decoration:none;">invest@capitaniainvestimentos.com.br</a><br>
<a href="https://www.capitaniainvestimentos.com.br" style="color:#1B51A3;text-decoration:none;">www.capitaniainvestimentos.com.br</a>
</p>
<p style="font-size:9pt;color:#1B51A3;">
Informação confidencial para uso exclusivo pelo destinatário da mensagem. Confidential information for exclusive use by recipient.
</p></body></html>"""


def _envio_abrir_outlook(cliente_cfg, anexos, data_exibicao, n_esperado):
    """Abre o rascunho no Outlook (Display, NUNCA Send). So a Lucieli usa o dash,
    na maquina servidora, entao o rascunho abre direto no Outlook dela.

    TUDO-OU-NADA: nunca abre rascunho incompleto. Revalida na HORA (sem cache) que
    todos os n_esperado arquivos ainda existem e que cada anexo foi de fato
    adicionado; se faltar/falhar qualquer um, DESCARTA o rascunho e retorna erro -
    o cache de 60s pode estar defasado (arquivo movido/renomeado ou Mellon ainda
    gravando o XML no momento do clique).
    Retorna (True, None) ou (False, motivo)."""
    # 1. Revalida quantidade e existencia real (os.path.exists = tempo real, sem cache)
    if len(anexos) != n_esperado:
        return False, f"esperado {n_esperado} arquivos, mas so {len(anexos)} encontrados - rascunho NAO aberto"
    sumiram = [os.path.basename(c) for c in anexos if not os.path.exists(c)]
    if sumiram:
        return False, f"arquivo(s) sumiram desde a ultima leitura: {', '.join(sumiram)} - rascunho NAO aberto (clique em Atualizar)"
    try:
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)
        email.To = cliente_cfg['to']
        email.CC = cliente_cfg['cc']
        if cliente_cfg.get('bcc'):
            email.BCC = cliente_cfg['bcc']
        email.Subject = cliente_cfg['assunto'].format(data=data_exibicao)
        email.HTMLBody = _envio_corpo_html()
        anexados = 0
        for caminho in anexos:
            try:
                email.Attachments.Add(os.path.abspath(caminho))
                anexados += 1
            except Exception as e:
                # Aborta: nao deixa abrir rascunho faltando anexo. Descarta o item.
                try:
                    email.Close(1)  # olDiscard
                except Exception:
                    pass
                return False, f"falha ao anexar {os.path.basename(caminho)}: {e} - rascunho NAO aberto"
        if anexados != n_esperado:
            try:
                email.Close(1)
            except Exception:
                pass
            return False, f"anexou {anexados}/{n_esperado} - rascunho NAO aberto"
        email.Display()
        return True, None
    except Exception as e:
        return False, str(e)


_ENVIO_PALETA = {
    'ok':      {'cor': '#22c55e', 'bg': '#dcfce7', 'fg': '#166534', 'icon': '✅', 'pill': 'PRONTO'},
    'pend':    {'cor': '#f59e0b', 'bg': '#fef3c7', 'fg': '#92400e', 'icon': '⏳', 'pill': 'AGUARDANDO'},
    'zero':    {'cor': '#ef4444', 'bg': '#fee2e2', 'fg': '#991b1b', 'icon': '⚠️', 'pill': 'ERRO'},
    'enviado': {'cor': '#1C57A8', 'bg': '#dbeafe', 'fg': '#1e40af', 'icon': '📧', 'pill': 'ENVIADO'},
}


@st.cache_data(ttl=120, show_spinner=False)
def _envio_ja_enviado(assunto_exato):
    """Confere na CAIXA DE ENTRADA padrao se ja existe e-mail com este assunto
    exato = prova de que o cliente JA foi enviado (To/CC=invest@ entrega copia a
    todos do time). Mais confiavel que qualquer flag manual. Cache 120s.
    Retorna (True, 'DD/MM HH:MM', remetente) ou (False, None, None)."""
    try:
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        ns = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
        inbox = ns.GetDefaultFolder(6)
        itens = inbox.Items
        try:
            res = itens.Restrict("[Subject] = '" + assunto_exato.replace("'", "''") + "'")
        except Exception:
            res = itens
        for msg in res:
            try:
                if str(msg.Subject).strip() == assunto_exato:
                    return True, msg.ReceivedTime.strftime('%d/%m %H:%M'), str(msg.SenderName)
            except Exception:
                continue
        return False, None, None
    except Exception:
        return False, None, None


def render_envio_diario():
    """Seção 'Envio Diário · XMLs Mellon' — ICATU / Aquila / BASF (versão compacta)."""
    if 'envio_data' not in st.session_state:
        st.session_state.envio_data = _envio_data_default()
    data_sel = st.session_state.envio_data
    data_yyyymmdd = data_sel.strftime('%Y%m%d')
    data_exibicao = data_sel.strftime('%d/%m/%Y')

    status_clientes = {}
    for nome, cfg in ENVIO_DIARIO_CLIENTES.items():
        info = _envio_buscar_arquivos(cfg, data_yyyymmdd)
        falt = info['xmls_falt'] + info['extras_falt']
        n_total = len(cfg['codigos']) + len(cfg.get('extras', []))
        n_ok = len(info['xmls_ok']) + len(info['extras_ok'])
        # Fonte da verdade: a caixa de entrada (cópia invest@). Se ja saiu, mostra enviado.
        assunto = cfg['assunto'].format(data=data_exibicao)
        ja, hr_env, remet = _envio_ja_enviado(assunto)
        if ja:
            sub = f'enviado {hr_env}' + (f' · {remet.split()[0]}' if remet else '')
            status_clientes[nome] = ('enviado', sub, info, n_ok, n_total)
        elif info.get('pasta_off'):
            status_clientes[nome] = ('zero', 'pasta X: offline', info, 0, n_total)
        elif not falt:
            status_clientes[nome] = ('ok', f'{n_ok}/{n_total} arquivos', info, n_ok, n_total)
        else:
            status_clientes[nome] = ('pend', f'{n_ok}/{n_total} arquivos', info, n_ok, n_total)

    n_enviados = sum(1 for s in status_clientes.values() if s[0] == 'enviado')
    n_prontos  = sum(1 for s in status_clientes.values() if s[0] == 'ok')
    n_total_c  = len(status_clientes)
    if n_enviados == n_total_c:
        resumo = f"✅ {n_enviados}/{n_total_c} enviados"
    elif n_prontos:
        resumo = f"📧 {n_prontos} pronto · {n_total_c - n_enviados - n_prontos} aguardando"
    else:
        resumo = f"⏳ aguardando ({n_total_c - n_enviados})"

    label = f"📤 Envio Diário · XMLs Mellon   ·   {data_exibicao}   ·   {resumo}"

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander(label, expanded=False):
        # Topo: date picker pequeno
        c1, c2 = st.columns([1, 5])
        with c1:
            nova = st.date_input("Data ref.", value=data_sel, key="envio_date_input",
                                 format="DD/MM/YYYY", label_visibility="collapsed")
            if nova != data_sel:
                st.session_state.envio_data = nova
                st.rerun()
        with c2:
            st.markdown(
                f"<div style='padding-top:8px;font-size:11px;color:#94a3b8'>"
                f"Detecta automaticamente quando os XMLs chegam em "
                f"<code style='font-size:10px'>{ENVIO_DIARIO_PASTA_XML}</code></div>",
                unsafe_allow_html=True)

        st.markdown("<div style='margin:6px 0'></div>", unsafe_allow_html=True)

        # Linha por cliente
        for nome, cfg in ENVIO_DIARIO_CLIENTES.items():
            accent, sub, info, n_ok, n_total = status_clientes[nome]
            pal = _ENVIO_PALETA[accent]
            pct = int(100 * n_ok / n_total) if n_total else 0
            ja_enviado = (accent == 'enviado')

            cnome, cstatus, cprog, cacao = st.columns([2, 2, 3, 3])

            with cnome:
                st.markdown(
                    f"<div style='padding-top:8px;font-size:13px;font-weight:700;color:#1a2540'>"
                    f"{pal['icon']} {nome}</div>",
                    unsafe_allow_html=True)

            with cstatus:
                st.markdown(
                    f"<div style='padding-top:6px'>"
                    f"<span style='background:{pal['bg']};color:{pal['fg']};"
                    f"padding:4px 12px;border-radius:14px;font-size:10.5px;font-weight:700;"
                    f"letter-spacing:.04em'>{pal['pill']}</span>"
                    f"<span style='font-size:11px;color:#64748b;margin-left:8px'>{sub}</span>"
                    f"</div>",
                    unsafe_allow_html=True)

            with cprog:
                st.markdown(
                    f"<div style='padding-top:14px'>"
                    f"<div style='background:#e8edf5;height:5px;border-radius:3px;overflow:hidden'>"
                    f"<div style='width:{pct}%;height:5px;background:{pal['cor']};"
                    f"border-radius:3px;transition:width .3s'></div></div></div>",
                    unsafe_allow_html=True)

            with cacao:
                if ja_enviado:
                    st.markdown(
                        "<div style='padding-top:10px;font-size:11px;color:#1e40af;"
                        "font-weight:600'>✔ confirmado na caixa</div>",
                        unsafe_allow_html=True)
                elif accent == 'ok':
                    # 100% dos arquivos -> abre o rascunho no Outlook (Display, NUNCA Send).
                    # Nao marca flag: quando o e-mail for enviado, a copia cai na caixa
                    # (invest@) e o card vira 'ENVIADO' sozinho - prova real de envio.
                    if st.button("📧 Abrir rascunho", key=f"envio_draft_{nome}",
                                 use_container_width=True, type="primary",
                                 help="Abre o rascunho no Outlook. Revise e envie — ao enviar, "
                                      "o card vira ENVIADO automaticamente (detecta na caixa)."):
                        anexos = info['xmls_ok'] + info['extras_ok']
                        ok, err = _envio_abrir_outlook(cfg, anexos, data_exibicao, n_total)
                        if ok:
                            st.toast(f"Rascunho {nome} aberto no Outlook", icon="✉️")
                        else:
                            st.error(f"Erro Outlook: {err}")
                else:
                    # mostra chips dos faltantes (apenas códigos curtos)
                    falt = info['xmls_falt'] + [os.path.basename(x).replace(f'_{data_yyyymmdd}.xlsx', '') for x in info['extras_falt']]
                    chips = ''.join(
                        f"<span style='display:inline-block;background:#fef3c7;color:#92400e;"
                        f"padding:2px 8px;border-radius:10px;font-size:10px;font-weight:600;"
                        f"margin:1px 2px'>{c}</span>"
                        for c in falt[:3])
                    if len(falt) > 3:
                        chips += f"<span style='font-size:10px;color:#94a3b8'> +{len(falt)-3}</span>"
                    st.markdown(f"<div style='padding-top:10px'>{chips}</div>", unsafe_allow_html=True)

            st.markdown(
                "<hr style='margin:6px 0;border:none;border-top:1px solid #f1f5f9'>",
                unsafe_allow_html=True)


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

# Calendario B3/BVMF: e o calendario que determina quando EXISTE cota (dias de
# mercado/NAV). Ja inclui TODOS os feriados nacionais + Corpus Christi, Carnaval e
# Sexta-feira Santa. Exclui de proposito feriados puramente estaduais/municipais de
# SP (raros, ex: 09/07), que sao dias de mercado normais com cota.
# Sem o calendario certo, o D-1 util do dash (ref_de) calcula a data de referencia
# errada em torno desses feriados e a coluna do dia fica vazia / cota na coluna errada.
_anos_feriados = [today.year - 1, today.year, today.year + 1]
feriados_br = holidays.financial_holidays('BVMF', years=_anos_feriados)

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


# Cache de leitura I/O - reduz lentidao em reruns frequentes (autorefresh,
# interacao com filtros). TTL 60s mantem dado fresco mas evita rescan a cada
# clique. Auto-refresh do dash e a cada 120s, entao 60s nao atrapalha.
@st.cache_data(ttl=60, show_spinner=False)
def _load_json_cached(path: str):
    """Carrega JSON do disco, ou {} se nao existir. Cacheado 60s."""
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


@st.cache_data(ttl=60, show_spinner=False)
def _scan_pdfs_dia(d_ref: str):
    """Retorna lista (nome_fundo, mtime_unix) para PDFs do dia. Cache 60s.
    O calculo de 'atrasado' depende do dia de envio (d, nao d_ref), entao
    fica fora desse cache - quem chama compara mtime com d.date()."""
    pdfs = glob.glob(os.path.join(PDF_DIR, f"*_{d_ref}.pdf"))
    resultado = []
    for p in pdfs:
        nome = os.path.basename(p).rsplit(f"_{d_ref}.pdf", 1)[0]
        try:
            mtime = os.path.getmtime(p)
        except OSError:
            continue
        resultado.append((nome, mtime))
    return resultado


@st.cache_data(ttl=120, show_spinner=False)
def _scan_cotas_email(drefs, dt_ini_iso, dt_fim_iso):
    """Detecta cotas JA enviadas olhando a CAIXA DE ENTRADA do Outlook.

    Todo e-mail 'COTA DIÁRIA | ...' tem invest@ em copia, entao a mensagem volta
    para a caixa de entrada. Isso vale tanto para o que o robô envia quanto para o
    que e' enviado MANUALMENTE (exceções) - que e' justamente o que o robô/PDF nao
    enxerga. O nome do anexo segue sempre o padrao {fundo}_{YYYYMMDD}.pdf (o mesmo
    {fundo} que o dash usa), entao da' para saber qual fundo foi enviado em cada
    data de referencia.

    Varre apenas a janela de datas da semana exibida (ordena do mais novo para o
    mais antigo e para quando passa da janela). Retorna {d_ref: {fundo: dt_envio}}.
    Cache 120s (igual ao auto-refresh). Se o Outlook estiver indisponivel, retorna
    tudo vazio e o dash continua funcionando pelos PDFs/JSONs normalmente."""
    res = {dr: {} for dr in drefs}
    drefs_set = set(drefs)
    try:
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        ns = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
        inbox = ns.GetDefaultFolder(6)
        itens = inbox.Items
        itens.Sort("[ReceivedTime]", True)  # mais novo -> mais antigo
        dt_ini = datetime.fromisoformat(dt_ini_iso)
        dt_fim = datetime.fromisoformat(dt_fim_iso)
        padrao = re.compile(r'(.+)_(\d{8})\.pdf$', re.IGNORECASE)
        for msg in itens:
            try:
                rt = msg.ReceivedTime
                rt_naive = datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)
            except Exception:
                continue
            if rt_naive > dt_fim:
                continue   # mais novo que a janela -> pula
            if rt_naive < dt_ini:
                break      # passou da janela (lista ordenada desc) -> encerra
            try:
                # 'COTA DI' pega 'COTA DIÁRIA'/'COTA DIARIA' e exclui 'Relatório de Cotas'
                if 'COTA DI' not in str(msg.Subject).upper():
                    continue
                for i in range(1, msg.Attachments.Count + 1):
                    m = padrao.match(msg.Attachments.Item(i).FileName)
                    if not m:
                        continue
                    fundo, dref = m.group(1), m.group(2)
                    if dref in drefs_set and fundo not in res[dref]:
                        res[dref][fundo] = rt_naive
            except Exception:
                continue
    except Exception:
        pass
    return res

status     = {}
erros      = {}
horarios   = {}
timestamps = {}   # {d_str: {fundo: {"dt": datetime, "atrasado": bool}}}
manuais_aprovados = {}  # {d_str: set de fundos manuais aprovados}
aguardando = {}   # {d_str: {fundo: {"desde": datetime, "motivo": str}}} - cota nao chegou no banco
orfas      = {}   # {d_str: {fundo: {"iniciado": datetime}}} - tentativa sem resultado, requer revisao

# Cotas detectadas pelo e-mail 'COTA DIÁRIA' na caixa de entrada (robô + manuais).
# Varre so a janela de datas da semana exibida, de uma vez (cacheado 120s).
_drefs_semana = tuple(sorted({ref_de(d).strftime("%Y%m%d") for d in dias}))
_dt_ini_scan = datetime(dias[0].year, dias[0].month, dias[0].day)
_dt_fim_scan = datetime(dias[-1].year, dias[-1].month, dias[-1].day) + timedelta(days=1)
cotas_email = _scan_cotas_email(_drefs_semana, _dt_ini_scan.isoformat(), _dt_fim_scan.isoformat())

for d in dias:
    d_str = d.strftime("%Y%m%d")
    d_ref = ref_de(d).strftime("%Y%m%d")

    def _load(prefix):
        return _load_json_cached(os.path.join(JSON_DIR, f"{prefix}_{d_ref}.json"))

    # Carregar aprovados do dia (scan_outlook gera este JSON)
    _aprov_data = _load_json_cached(os.path.join(JSON_DIR, f"aprovados_{d_ref}.json"))
    _aprov_manual = set()
    _aprov_site = set()
    _aprov_manual_erros = {}
    if _aprov_data:
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
        # Escaneia pasta de PDFs (cacheado 60s para reduzir lentidao em Z:\)
        processados = set()
        ts_dia = {}
        for nome, mtime in _scan_pdfs_dia(d_ref):
            processados.add(nome)
            dt_criacao = datetime.fromtimestamp(mtime)
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
        # COTA DIÁRIA detectada na caixa de entrada (robô OU manual). Pega as
        # exceções enviadas na mao, que nao geram PDF na pasta. So preenche o que
        # ainda nao foi detectado pelo PDF; marca 'manual' para sinalizar (✋) que
        # nao veio do robô.
        for _fm, _dtm in cotas_email.get(d_ref, {}).items():
            if _fm not in processados:
                processados.add(_fm)
                ts_dia[_fm] = {"dt": _dtm, "atrasado": False, "manual": True}
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


# Fundos que o robo escreve nos JSONs (aguardando/tentativas) mas que NAO sao
# tratados pelo dash - nao devem aparecer nos banners de aviso.
IGNORAR_AVISOS = {"ARTON FOF", "ARTON JP", "CAPITANIA FIAGRO", "FIAGRO XP"}

# ── BANNER: TENTATIVAS ORFAS (REQUER REVISAO HUMANA) ─────────────────────────
_hoje_str = today.strftime("%Y%m%d")
_orfas_hoje = {
    _f: _info
    for _f, _info in orfas.get(_hoje_str, {}).items()
    if _f not in IGNORAR_AVISOS
}
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

def _aguard_motivo_curto(motivo):
    """Classifica o motivo do 'aguardando' para exibir o rótulo certo:
    - 'carteira não bate'/batimento => a cota JÁ está na base, o que trava é o
      batimento (COTAS_CAP). NÃO é falta de cota.
    - senão => cota ainda não foi lançada no COTAS_CAP.
    Evita mostrar 'aguardando COTAS_CAP' (parece falta de cota) quando na verdade
    a cota chegou e só o batimento está pendente."""
    m = (motivo or "").lower()
    if "batimento" in m or "carteira" in m or "bate" in m:
        return "carteira não bate"
    return "aguardando cota"


# ── BANNER: FUNDOS AGUARDANDO COTA NO BANCO ─────────────────────────────────
# So mostra fundos que AINDA nao foram processados/enviados hoje. Se o fundo ja
# aparece como ✅ na tabela (esta no set de processados), some do banner sozinho -
# mesmo que a entrada em aguardando_<ref>.json nao tenha sido removida pelo robo.
# Tambem esconde os fundos de outros times (IGNORAR_AVISOS).
_proc_hoje = status.get(_hoje_str)
_proc_hoje = _proc_hoje if isinstance(_proc_hoje, set) else set()
_aguard_hoje = {
    _f: _info
    for _f, _info in aguardando.get(_hoje_str, {}).items()
    if _f not in _proc_hoje and _f not in IGNORAR_AVISOS
}
if _aguard_hoje:
    _linhas = []
    for _f, _info in sorted(_aguard_hoje.items()):
        _min = int((datetime.now() - _info["desde"]).total_seconds() / 60)
        _rot = _aguard_motivo_curto(_info.get("motivo", ""))
        _linhas.append(f"<b>{_f}</b> — {_rot} ({_min} min)")
    _lista_html = "<br>".join(_linhas)
    _total = len(_aguard_hoje)
    st.markdown(f"""
    <div style="background:#fed7aa; border-left:5px solid #c2410c;
                padding:12px 16px; border-radius:6px; margin-bottom:16px;">
      <div style="font-size:14px; font-weight:700; color:#9a3412; margin-bottom:6px;">
        ⏳ {_total} fundo(s) aguardando envio (cota ou batimento no COTAS_CAP)
      </div>
      <div style="font-size:12px; color:#7c2d12;">
        {_lista_html}
      </div>
      <div style="font-size:11px; color:#7c2d12; margin-top:6px; font-style:italic;">
        "aguardando cota" = cota ainda não lançada no COTAS_CAP (após 25 min o robô cria rascunho de cobrança).
        "carteira não bate" = cota já está na base, o que trava é o batimento — verificar a carteira.
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


# ── ENVIO DIÁRIO (XMLs Mellon) ────────────────────────────────────────────────
render_envio_diario()


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
            elif ts and ts.get("manual"):
                # Enviado MANUALMENTE (exceção) - detectado pelo e-mail, sem PDF.
                linha[col] = f"✅ {ts['dt'].strftime('%H:%M')} ✋"
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
                linha[col] = f"⏳ {_aguard_motivo_curto(info_ag.get('motivo', ''))} {min_decorridos}min"
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
