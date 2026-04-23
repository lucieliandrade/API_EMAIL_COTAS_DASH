"""
mailer_robo.py - Robo automatico de mailers de cotas diarias
============================================================
Monitora a caixa de entrada do Outlook a cada 2 minutos.
Quando encontra emails de aprovacao ("Carteiras Aprovadas - Fundos [ADM] - Sistema Backoffice"),
extrai a data e os fundos aprovados, e executa o mailer automaticamente.

REGRA: NUNCA envia email direto. Sempre Display() = rascunho para revisao manual.

Uso:
    python mailer_robo.py

Para parar: Ctrl+C
"""

import win32com.client as win32
from datetime import datetime
import re
import os
import json
import sys
import time
import subprocess
import traceback
import msvcrt
import threading
import ctypes

INTERVALO_MINUTOS = 2
MAX_FALHAS_POR_FUNDO = 3        # apos 3 falhas seguidas, para de tentar no dia
TIMEOUT_CICLO_SEG = 600          # watchdog do ciclo: MAIOR que timeout do subprocess (300s), evita matar o robo no meio e causar duplicidade
DIRETORIO = r"Z:\Relações com Investidores - NOVO\codigos\cotas"

# Alerta de cobranca quando fundos ficam muito tempo com dado ausente no banco
ALERTA_COBRANCA_MINUTOS = 25        # apos este tempo aguardando, gera o PRIMEIRO rascunho de cobranca
ALERTA_COBRANCA_INTERVALO_MIN = 120 # apos o primeiro, gera novo rascunho a cada X min se ainda tiver pendente
ALERTA_COBRANCA_EMAIL = "lucieli.andrade@capitaniainvestimentos.com.br"  # destinatario do rascunho

# Motivos de erro do mailer que indicam "dado ausente no banco" (nao culpa do robo/codigo).
# Esses erros NAO contam como falha permanente: quando o dado aparecer, o robo reprocessa automaticamente.
PADROES_DADO_AUSENTE = (
    "não consta no COTAS_CAP",
    "nao consta no COTAS_CAP",
    "sem dados para o dia",
    "igual a 0 ou NaN no COTAS_CAP",
    "valores zerados na tabela",
)

def eh_dado_ausente(motivo):
    """True se o motivo indica que um dado externo (cota, benchmark) ainda nao chegou."""
    if not motivo:
        return False
    m = str(motivo)
    return any(p in m for p in PADROES_DADO_AUSENTE)

# Fundos manuais: enviados por email com PDF, NAO pelo mailer automatico.
# O robo deve pular esses fundos (scan_outlook.py detecta o envio deles).
FUNDOS_MANUAIS = {
    "FCopel", "FCopel_Imob", "Sabesprev", "CAPITANIA REIT", "PETROS RFCP",
    "OPOR IMOB FII", "OPOR IMOB SUBCLA", "OPOR IMOB SUBCLB", "OPOR IMOB SUBCLC",
    "CAPITANIA FCOPEL",  # alias do FCopel: nome usado no email de aprovacao do Itau
}
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_MAILER = os.path.join(SCRIPT_DIR, "mailer_v_auto.py")
LOCK_FILE = os.path.join(SCRIPT_DIR, "mailer_robo.lock")

# ── TRAVA: impede mais de 1 instancia (com deteccao de processo morto) ────
def _adquirir_lock():
    """Tenta adquirir o lock. Se o processo anterior morreu/travou, retoma."""
    # Checar se o lock existe e se o PID anterior ainda esta vivo
    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, 'r') as f:
                conteudo = f.read().strip()
            if conteudo.isdigit():
                pid_antigo = int(conteudo)
                # Verificar se o processo ainda existe
                kernel32 = ctypes.windll.kernel32
                handle = kernel32.OpenProcess(0x0400, False, pid_antigo)  # PROCESS_QUERY_INFORMATION
                if handle:
                    kernel32.CloseHandle(handle)
                    # Processo existe — checar se o lock esta velho (> 10 min = travado)
                    idade = time.time() - os.path.getmtime(LOCK_FILE)
                    if idade > 600:
                        print(f"  Lock antigo (PID {pid_antigo}, {idade/60:.0f} min). Matando processo travado...")
                        try:
                            kernel32.TerminateProcess(
                                kernel32.OpenProcess(0x0001, False, pid_antigo), 1)  # PROCESS_TERMINATE
                        except:
                            pass
                        time.sleep(2)
                    else:
                        print(f"BLOQUEADO: Outra instancia (PID {pid_antigo}) rodando ha {idade/60:.0f} min. Encerrando.")
                        sys.exit(0)
                # else: processo morreu, lock e orfao — prosseguir
        except:
            pass  # arquivo corrompido, prosseguir

    # Adquirir o lock
    fh = open(LOCK_FILE, 'w')
    try:
        msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
    except (OSError, IOError):
        print("BLOQUEADO: Outra instancia do robo ja esta rodando. Encerrando.")
        sys.exit(0)
    fh.write(str(os.getpid()))
    fh.flush()
    return fh

_lock_fh = _adquirir_lock()
LOG_PATH = os.path.join(SCRIPT_DIR, "robo_log.txt")

# Redireciona print para o arquivo de log E para o console
class _Tee:
    def __init__(self, *streams): self.streams = streams
    def write(self, data):
        for s in self.streams:
            try: s.write(data); s.flush()
            except: pass
    def flush(self):
        for s in self.streams:
            try: s.flush()
            except: pass

_log_file = open(LOG_PATH, 'a', encoding='utf-8', buffering=1)
sys.stdout = _Tee(sys.__stdout__, _log_file)
sys.stderr = _Tee(sys.__stderr__, _log_file)


################################## LEITURA DE EMAILS DO OUTLOOK ##################################

def ler_emails_aprovacao():
    """
    Le TODOS os emails de aprovacao da Caixa de Entrada do Outlook recebidos hoje.
    Suporta todos os ADMs: Bradesco, BTG, BNYM/Mellon, Itau, XP.

    Returns:
        Lista de dicts: [{'adm': str, 'data_ref': str (DD/MM), 'fundos': list, 'msg': MailItem}]
    """
    outlook = win32.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # Mais recentes primeiro

    hoje = datetime.today().date()
    resultados = []

    for msg in messages:
        try:
            received_date = msg.ReceivedTime.date()
            # Parar quando chegar em emails de dias anteriores
            if received_date < hoje:
                break
            if received_date != hoje:
                continue

            subject = str(msg.Subject)

            # Verificar se e email de aprovacao
            if 'Carteiras Aprovadas' not in subject or 'Sistema Backoffice' not in subject:
                continue

            # Extrair ADM do assunto: "Carteiras Aprovadas - Fundos Bradesco - Sistema Backoffice"
            match_adm = re.search(r'Fundos\s+(.+?)\s*-\s*Sistema', subject)
            if not match_adm:
                continue
            adm = match_adm.group(1).strip()

            body = str(msg.Body)

            # Extrair data de referencia: "referentes a DD/MM"
            match_data = re.search(r'referentes a (\d{1,2}/\d{2})', body)
            data_ref = match_data.group(1) if match_data else None

            # Extrair nomes dos fundos: entre "Carteiras Aprovadas" e "Atenciosamente"
            match_fundos = re.search(r'Carteiras Aprovadas\s*(.*?)Atenciosamente', body, re.DOTALL)
            fundos = []
            if match_fundos:
                linhas = match_fundos.group(1).strip().split('\n')
                fundos = [l.strip() for l in linhas if l.strip()]

            if data_ref and fundos:
                resultados.append({
                    'adm': adm,
                    'data_ref': data_ref,
                    'fundos': fundos,
                    'msg': msg
                })

        except Exception:
            continue

    return resultados


def mover_email_para_cotas(msg):
    """Move um email da Caixa de Entrada para a pasta ri_middle-cotas no Outlook."""
    try:
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        pasta_destino = None

        # Tentar como subpasta da Caixa de Entrada: ***RI_MIDDLE > COTAS
        try:
            pasta_destino = inbox.Folders("***RI_MIDDLE").Folders("COTAS")
        except:
            pass

        # Fallback sem asteriscos
        if pasta_destino is None:
            try:
                pasta_destino = inbox.Folders("RI_MIDDLE").Folders("COTAS")
            except:
                pass

        if pasta_destino is None:
            print("    Pasta RI_MIDDLE > COTAS nao encontrada no Outlook.")
            return False

        msg.Move(pasta_destino)
        return True

    except Exception as e:
        print(f"    Erro ao mover email: {e}")
        return False


################################## CONTROLE DE DUPLICIDADE ##################################

def get_arquivo_processados(data_ref_yyyymmdd):
    """Retorna o caminho do arquivo de controle de processados para a data."""
    pasta_json = os.path.join(DIRETORIO, "json")
    os.makedirs(pasta_json, exist_ok=True)
    return os.path.join(pasta_json, f"processados_{data_ref_yyyymmdd}.json")


def carregar_processados(data_ref_yyyymmdd):
    """Carrega o conjunto de fundos ja processados para a data."""
    arquivo = get_arquivo_processados(data_ref_yyyymmdd)
    if os.path.exists(arquivo):
        with open(arquivo, 'r', encoding='utf-8') as f:
            return set(json.load(f))
    return set()


def salvar_processados(data_ref_yyyymmdd, fundos):
    """Marca uma lista de fundos como processados para a data."""
    arquivo = get_arquivo_processados(data_ref_yyyymmdd)
    processados = carregar_processados(data_ref_yyyymmdd)
    processados.update(fundos)
    with open(arquivo, 'w', encoding='utf-8') as f:
        json.dump(list(processados), f, ensure_ascii=False, indent=2)


def salvar_erro(data_ref_yyyymmdd, fundo, motivo):
    """Salva o erro no JSON para o dash exibir."""
    pasta_json = os.path.join(DIRETORIO, "json")
    os.makedirs(pasta_json, exist_ok=True)
    arquivo = os.path.join(pasta_json, f"erros_{data_ref_yyyymmdd}.json")
    erros = {}
    if os.path.exists(arquivo):
        with open(arquivo, 'r', encoding='utf-8') as f:
            erros = json.load(f)
    erros[fundo] = motivo
    with open(arquivo, 'w', encoding='utf-8') as f:
        json.dump(erros, f, ensure_ascii=False, indent=2)


################################## CONTROLE DE AGUARDANDO DADO ##################################

def _path_aguardando(data_ref_yyyymmdd):
    pasta_json = os.path.join(DIRETORIO, "json")
    os.makedirs(pasta_json, exist_ok=True)
    return os.path.join(pasta_json, f"aguardando_{data_ref_yyyymmdd}.json")

def _path_alerta_cobranca(data_ref_yyyymmdd):
    pasta_json = os.path.join(DIRETORIO, "json")
    os.makedirs(pasta_json, exist_ok=True)
    return os.path.join(pasta_json, f"alerta_cobranca_{data_ref_yyyymmdd}.json")

def carregar_aguardando(data_ref_yyyymmdd):
    """Retorna {fundo: timestamp_iso_primeira_deteccao}."""
    arquivo = _path_aguardando(data_ref_yyyymmdd)
    if os.path.exists(arquivo):
        with open(arquivo, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def registrar_aguardando(data_ref_yyyymmdd, fundo, motivo):
    """Marca o fundo como aguardando dado. Mantem o timestamp da PRIMEIRA deteccao para calcular o tempo total."""
    arquivo = _path_aguardando(data_ref_yyyymmdd)
    ag = carregar_aguardando(data_ref_yyyymmdd)
    if fundo not in ag:
        ag[fundo] = {
            "desde": datetime.now().isoformat(timespec='seconds'),
            "motivo": motivo,
        }
        with open(arquivo, 'w', encoding='utf-8') as f:
            json.dump(ag, f, ensure_ascii=False, indent=2)

def remover_aguardando(data_ref_yyyymmdd, fundo):
    """Remove o fundo do aguardando (ex: quando a cota finalmente chegou e o fundo processou OK)."""
    arquivo = _path_aguardando(data_ref_yyyymmdd)
    ag = carregar_aguardando(data_ref_yyyymmdd)
    if fundo in ag:
        del ag[fundo]
        with open(arquivo, 'w', encoding='utf-8') as f:
            json.dump(ag, f, ensure_ascii=False, indent=2)

def carregar_historico_alertas(data_ref_yyyymmdd):
    """Retorna lista de alertas ja enviados hoje: [{'em': iso, 'fundos': [...]}, ...]"""
    path = _path_alerta_cobranca(data_ref_yyyymmdd)
    if not os.path.exists(path):
        return []
    with open(path, 'r', encoding='utf-8') as f:
        dados = json.load(f)
    # Compatibilidade com formato antigo (dict unico ao inves de lista)
    if isinstance(dados, dict) and "alertas" in dados:
        return dados["alertas"]
    if isinstance(dados, dict) and "enviado_em" in dados:
        return [{"em": dados["enviado_em"], "fundos": dados.get("fundos", [])}]
    return []

def deve_enviar_alerta(data_ref_yyyymmdd):
    """Retorna True se for hora de enviar um novo alerta de cobranca:
    - Primeiro alerta: nunca foi enviado antes.
    - Re-cobranca: ultimo alerta foi ha >= ALERTA_COBRANCA_INTERVALO_MIN."""
    historico = carregar_historico_alertas(data_ref_yyyymmdd)
    if not historico:
        return True  # sem alerta anterior, entao dispara
    # Tempo desde o ultimo alerta
    try:
        ultimo = datetime.fromisoformat(historico[-1]["em"])
    except Exception:
        return True
    return (datetime.now() - ultimo).total_seconds() >= ALERTA_COBRANCA_INTERVALO_MIN * 60

def marcar_alerta_enviado(data_ref_yyyymmdd, fundos):
    """Append de um novo alerta no historico. Permite re-cobrancas a cada ALERTA_COBRANCA_INTERVALO_MIN."""
    historico = carregar_historico_alertas(data_ref_yyyymmdd)
    historico.append({
        "em": datetime.now().isoformat(timespec='seconds'),
        "fundos": list(fundos),
    })
    with open(_path_alerta_cobranca(data_ref_yyyymmdd), 'w', encoding='utf-8') as f:
        json.dump({"alertas": historico}, f, ensure_ascii=False, indent=2)

def criar_rascunho_cobranca(data_completa, fundos_aguardando_info, tentativa=1):
    """Cria UM rascunho no Outlook com a lista dos fundos pendentes no banco COTAS_CAP.
    tentativa=1 -> primeiro alerta, tentativa>1 -> re-cobranca.
    REGRA: apenas Display() - nunca Send. Usuario revisa e envia manualmente."""
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)  # 0 = MailItem

        data_br = datetime.strptime(data_completa, '%Y-%m-%d').strftime('%d/%m')
        if tentativa == 1:
            mail.Subject = f"Cotas pendentes no COTAS_CAP - ref {data_br}"
            cabecalho = (f"<p>Os fundos abaixo tiveram carteira aprovada pelo Backoffice mas a cota de "
                         f"<b>{data_br}</b> ainda nao foi lancada na tabela <b>COTAS_CAP</b> do banco:</p>")
        else:
            mail.Subject = f"REITERACAO #{tentativa} - Cotas pendentes no COTAS_CAP - ref {data_br}"
            cabecalho = (f"<p><b>Reiteracao (tentativa {tentativa}):</b> os fundos abaixo continuam pendentes "
                         f"de lancamento da cota de <b>{data_br}</b> na tabela <b>COTAS_CAP</b>:</p>")
        mail.To = ALERTA_COBRANCA_EMAIL

        linhas_html = []
        for fundo, info in sorted(fundos_aguardando_info.items()):
            desde = datetime.fromisoformat(info["desde"])
            decorrido_min = int((datetime.now() - desde).total_seconds() / 60)
            linhas_html.append(
                f"<li><b>{fundo}</b> — aguardando ha {decorrido_min} min "
                f"(desde {desde.strftime('%H:%M')})</li>"
            )

        mail.HTMLBody = f"""
        {cabecalho}
        <ul>
            {''.join(linhas_html)}
        </ul>
        <p>Sem a cota no banco, o mailer de cotas diarias nao consegue ser gerado.</p>
        <p>Enviar para o time responsavel.</p>
        """
        mail.Display()  # NUNCA trocar por Send() - usuario revisa e envia manualmente
        return True
    except Exception as e:
        print(f"  ERRO ao criar rascunho de cobranca: {e}")
        return False

def avaliar_alerta_cobranca():
    """Verifica todos os arquivos aguardando_*.json e, para cada data com fundos pendentes
    ha mais de ALERTA_COBRANCA_MINUTOS (e sem alerta enviado hoje), cria um rascunho."""
    pasta_json = os.path.join(DIRETORIO, "json")
    if not os.path.isdir(pasta_json):
        return

    agora = datetime.now()
    for nome in os.listdir(pasta_json):
        if not nome.startswith("aguardando_") or not nome.endswith(".json"):
            continue
        data_json = nome[len("aguardando_"):-len(".json")]  # ex: 20260422

        # Ja alertou recentemente? Pula. Senao, pode ser primeiro alerta ou re-cobranca.
        if not deve_enviar_alerta(data_json):
            continue

        ag = carregar_aguardando(data_json)
        # Filtrar apenas os fundos com >= ALERTA_COBRANCA_MINUTOS de espera
        pendentes = {}
        for fundo, info in ag.items():
            try:
                desde = datetime.fromisoformat(info["desde"])
            except Exception:
                continue
            if (agora - desde).total_seconds() >= ALERTA_COBRANCA_MINUTOS * 60:
                pendentes[fundo] = info

        if not pendentes:
            continue

        # Converter data_json (YYYYMMDD) para YYYY-MM-DD
        data_completa = f"{data_json[:4]}-{data_json[4:6]}-{data_json[6:]}"
        tentativa = len(carregar_historico_alertas(data_json)) + 1
        rotulo = "cobranca" if tentativa == 1 else f"RE-COBRANCA #{tentativa}"
        print(f"\n  [ALERTA] {len(pendentes)} fundo(s) aguardando cota para ref {data_completa}. Criando rascunho de {rotulo}...")
        if criar_rascunho_cobranca(data_completa, pendentes, tentativa=tentativa):
            marcar_alerta_enviado(data_json, pendentes.keys())
            print(f"  [ALERTA] Rascunho criado no Outlook (Display). Proxima re-cobranca em {ALERTA_COBRANCA_INTERVALO_MIN} min se ainda pendente.")


################################## CONTROLE DE FALHAS ##################################

_falhas_hoje = {}  # {fundo: contagem} — resetado quando muda o dia
_dia_falhas = None

def _get_falhas():
    """Retorna o dict de falhas do dia, resetando se mudou o dia."""
    global _falhas_hoje, _dia_falhas
    hoje = datetime.today().strftime('%Y%m%d')
    if _dia_falhas != hoje:
        _falhas_hoje = {}
        _dia_falhas = hoje
    return _falhas_hoje

def registrar_falha(fundo):
    """Registra +1 falha para o fundo. Retorna True se atingiu o limite."""
    falhas = _get_falhas()
    falhas[fundo] = falhas.get(fundo, 0) + 1
    return falhas[fundo] >= MAX_FALHAS_POR_FUNDO

def fundo_bloqueado(fundo):
    """Retorna True se o fundo ja falhou demais hoje."""
    falhas = _get_falhas()
    return falhas.get(fundo, 0) >= MAX_FALHAS_POR_FUNDO


################################## WATCHDOG (ANTI-TRAVAMENTO) ##################################

_watchdog_timer = None

def _watchdog_matar():
    """Chamado pelo timer se o ciclo excedeu o timeout. Mata o processo."""
    print(f"\n  WATCHDOG: ciclo travou por mais de {TIMEOUT_CICLO_SEG}s. Reiniciando processo...")
    sys.stdout.flush()
    os._exit(99)  # sai imediatamente — Task Scheduler ou Startup reinicia

def watchdog_iniciar():
    global _watchdog_timer
    watchdog_cancelar()
    _watchdog_timer = threading.Timer(TIMEOUT_CICLO_SEG, _watchdog_matar)
    _watchdog_timer.daemon = True
    _watchdog_timer.start()

def watchdog_cancelar():
    global _watchdog_timer
    if _watchdog_timer is not None:
        _watchdog_timer.cancel()
        _watchdog_timer = None


################################## CICLO PRINCIPAL ##################################

def processar_ciclo():
    """Um ciclo completo: ler emails, processar fundos novos, registrar."""

    agora = datetime.now().strftime('%H:%M:%S')
    print(f"\n{'='*60}")
    print(f"[{agora}] VERIFICANDO EMAILS DE APROVACAO...")
    print(f"{'='*60}")

    # 0. Atualizar JSON de aprovacoes para o dash (somente leitura)
    try:
        from scan_outlook import scan
        scan()
    except Exception as e:
        print(f"  Aviso: scan_outlook falhou ({e})")

    # 1. Ler emails de aprovacao (TODOS os ADMs)
    emails = ler_emails_aprovacao()

    if not emails:
        print("  Nenhum email de aprovacao encontrado na caixa de entrada.")
        return

    print(f"\n  {len(emails)} email(s) de aprovacao encontrado(s):")
    for i, e in enumerate(emails, 1):
        print(f"    {i}. ADM={e['adm']}, Data={e['data_ref']}, Fundos: {', '.join(e['fundos'])}")

    # 2. Agrupar fundos por data de referencia
    ano_atual = str(datetime.today().year)
    datas_fundos = {}

    for e in emails:
        partes = e['data_ref'].split('/')
        data_ref_completa = f"{ano_atual}-{partes[1]}-{partes[0].zfill(2)}"
        data_ref_json = f"{ano_atual}{partes[1]}{partes[0].zfill(2)}"

        if data_ref_json not in datas_fundos:
            datas_fundos[data_ref_json] = {
                'data_completa': data_ref_completa,
                'fundos': [],
                'emails': []
            }

        for f in e['fundos']:
            if f in FUNDOS_MANUAIS:
                continue  # fundo manual — enviado por PDF, nao pelo mailer
            if f not in datas_fundos[data_ref_json]['fundos']:
                datas_fundos[data_ref_json]['fundos'].append(f)
        datas_fundos[data_ref_json]['emails'].append(e)

    # 3. Processar cada data
    for data_json, info in datas_fundos.items():
        processados = carregar_processados(data_json)
        fundos_novos = [f for f in info['fundos'] if f not in processados]

        if not fundos_novos:
            print(f"\n  Data {info['data_completa']}: todos os {len(info['fundos'])} fundos ja foram processados.")
            continue

        # Mostrar status
        print(f"\n  Data {info['data_completa']}: {len(fundos_novos)} fundo(s) novo(s) de {len(info['fundos'])} total:")
        for f in info['fundos']:
            if f in processados:
                print(f"    [JA FEITO] {f}")
            elif fundo_bloqueado(f):
                print(f"    [PARADO]   {f}  (falhou {MAX_FALHAS_POR_FUNDO}x, nao tenta mais hoje)")
            else:
                print(f"    [NOVO]     {f}")

        # 4. Processar FUNDO A FUNDO (evita que 1 erro derrube o lote)
        fundos_a_tentar = [f for f in fundos_novos if not fundo_bloqueado(f)]
        if not fundos_a_tentar:
            print(f"\n  Todos os fundos novos ja atingiram o limite de {MAX_FALHAS_POR_FUNDO} falhas.")
            continue

        print(f"\n  Executando mailer para {len(fundos_a_tentar)} fundo(s)...")

        for fundo in fundos_a_tentar:
            # Reler processados antes de cada fundo (evita duplicidade)
            if fundo in carregar_processados(data_json):
                print(f"\n    [{fundo}] JA PROCESSADO (pulando)")
                continue

            resultado_path = os.path.join(DIRETORIO, "json", f"resultado_{data_json}_{datetime.now().strftime('%H%M%S')}.json")

            cmd = [
                sys.executable,
                SCRIPT_MAILER,
                '--data', info['data_completa'],
                '--fundos', fundo,
                '--resultado', resultado_path
            ]

            print(f"\n    [{fundo}] Processando...")

            try:
                subprocess.run(cmd, timeout=300)  # 5 min por fundo

                # Ler resultado
                if os.path.exists(resultado_path):
                    with open(resultado_path, 'r', encoding='utf-8') as f:
                        dados_resultado = json.load(f)
                    os.remove(resultado_path)

                    # Compatibilidade: formato novo = dict {ok, erros}; formato antigo = lista
                    if isinstance(dados_resultado, dict):
                        fundos_ok = dados_resultado.get("ok", [])
                        erros_motivo = dados_resultado.get("erros", {})
                    else:
                        fundos_ok = dados_resultado
                        erros_motivo = {}

                    if fundo in fundos_ok:
                        salvar_processados(data_json, [fundo])
                        remover_aguardando(data_json, fundo)  # cota chegou, sai do aguardando
                        print(f"    [{fundo}] OK")
                    else:
                        motivo_real = erros_motivo.get(fundo, "mailer nao processou (sem motivo reportado)")
                        # Se for "dado ausente" (cota/benchmark nao chegou no banco), NAO conta como falha permanente.
                        # O robo vai tentar de novo no proximo ciclo ate o dado aparecer.
                        if eh_dado_ausente(motivo_real):
                            registrar_aguardando(data_json, fundo, motivo_real)  # timestamp da primeira deteccao
                            salvar_erro(data_json, fundo, f"AGUARDANDO DADO para COTAS_CAP: {motivo_real}")
                            print(f"    [{fundo}] AGUARDANDO DADO para COTAS_CAP - {motivo_real}")
                        else:
                            bloqueou = registrar_falha(fundo)
                            salvar_erro(data_json, fundo, motivo_real)
                            print(f"    [{fundo}] ERRO - {motivo_real} (falha {_get_falhas().get(fundo,0)}/{MAX_FALHAS_POR_FUNDO})")
                            if bloqueou:
                                print(f"    [{fundo}] PARADO - nao tenta mais hoje")
                else:
                    bloqueou = registrar_falha(fundo)
                    motivo = "script falhou (sem resultado)"
                    salvar_erro(data_json, fundo, motivo)
                    print(f"    [{fundo}] ERRO - {motivo} (falha {_get_falhas().get(fundo,0)}/{MAX_FALHAS_POR_FUNDO})")
                    if bloqueou:
                        print(f"    [{fundo}] PARADO - nao tenta mais hoje")

            except subprocess.TimeoutExpired:
                bloqueou = registrar_falha(fundo)
                motivo = "timeout (5 min)"
                salvar_erro(data_json, fundo, motivo)
                print(f"    [{fundo}] TIMEOUT (5 min) (falha {_get_falhas().get(fundo,0)}/{MAX_FALHAS_POR_FUNDO})")
                if bloqueou:
                    print(f"    [{fundo}] PARADO - nao tenta mais hoje")
            except Exception as e:
                bloqueou = registrar_falha(fundo)
                motivo = str(e)
                salvar_erro(data_json, fundo, motivo)
                print(f"    [{fundo}] ERRO: {e} (falha {_get_falhas().get(fundo,0)}/{MAX_FALHAS_POR_FUNDO})")
                if bloqueou:
                    print(f"    [{fundo}] PARADO - nao tenta mais hoje")

    # 6. Mover emails processados para pasta COTAS
    # So move se TODOS os fundos do email foram processados
    for data_json, info in datas_fundos.items():
        processados = carregar_processados(data_json)
        for e in info['emails']:
            fundos_do_email = set(e['fundos'])
            if fundos_do_email.issubset(processados):
                try:
                    if mover_email_para_cotas(e['msg']):
                        print(f"  [MOVIDO] Email ADM={e['adm']} para pasta COTAS")
                except Exception:
                    pass

    # 7. Avaliar se precisa criar rascunho de cobranca (fundos aguardando ha > ALERTA_COBRANCA_MINUTOS)
    try:
        avaliar_alerta_cobranca()
    except Exception as e:
        print(f"  Aviso: avaliar_alerta_cobranca falhou ({e})")


################################## MAIN ##################################

def main():
    print()
    print("=" * 60)
    print("  ROBO DE MAILERS - COTAS DIARIAS")
    print(f"  Intervalo: {INTERVALO_MINUTOS} minutos")
    print(f"  Max falhas/fundo: {MAX_FALHAS_POR_FUNDO}")
    print(f"  Watchdog: {TIMEOUT_CICLO_SEG}s")
    print(f"  Mailer: {SCRIPT_MAILER}")
    print(f"  Diretorio: {DIRETORIO}")
    print(f"  PID: {os.getpid()}")
    print("=" * 60)
    print()
    print("  REGRA: todos os emails sao RASCUNHO (Display).")
    print("  NUNCA envia direto. Voce revisa e clica Enviar.")
    print()
    print("  ADMs monitorados: Bradesco, BTG, BNYM, Itau, XP")
    print()
    print("  Para parar: Ctrl+C")
    print("=" * 60)

    while True:
        try:
            # Watchdog: se o ciclo travar > TIMEOUT_CICLO_SEG, mata o processo
            # O auto_dash.bat ou Task Scheduler reinicia automaticamente
            watchdog_iniciar()
            processar_ciclo()
            watchdog_cancelar()
        except KeyboardInterrupt:
            watchdog_cancelar()
            print("\n\nRobo encerrado pelo usuario.")
            break
        except Exception as e:
            watchdog_cancelar()
            print(f"\n  ERRO NO CICLO: {e}")
            traceback.print_exc()

        # Atualizar timestamp do lock para o watchdog de instancias externas
        try:
            os.utime(LOCK_FILE, None)
        except:
            pass

        agora = datetime.now().strftime('%H:%M:%S')
        print(f"\n  [{agora}] Proximo ciclo em {INTERVALO_MINUTOS} minutos...")

        try:
            time.sleep(INTERVALO_MINUTOS * 60)
        except KeyboardInterrupt:
            print("\n\nRobo encerrado pelo usuario.")
            break


if __name__ == '__main__':
    main()
