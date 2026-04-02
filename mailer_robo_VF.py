"""
mailer_robo.py - Robo automatico de mailers de cotas diarias
============================================================
Monitora a caixa de entrada do Outlook a cada 5 minutos.
Quando encontra emails de aprovacao ("Carteiras Aprovadas - Fundos [ADM] - Sistema Backoffice"),
extrai a data e os fundos aprovados, e executa o mailer automaticamente.

Os emails sao abertos como rascunho (Display) para revisao antes do envio manual.

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


INTERVALO_MINUTOS = 2
DIRETORIO = r"Z:\Relações com Investidores - NOVO\codigos\cotas"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_MAILER = os.path.join(SCRIPT_DIR, "mailer_v_oficial_IPCA.py")


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
    """Move um email da Caixa de Entrada para a pasta COTAS no Outlook."""
    try:
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        pasta_cotas = None

        # Tentar como subpasta da Caixa de Entrada
        try:
            pasta_cotas = inbox.Folders("COTAS")
        except:
            pass

        # Se nao encontrou, tentar no nivel da conta
        if pasta_cotas is None:
            try:
                root_folder = inbox.Parent
                pasta_cotas = root_folder.Folders("COTAS")
            except:
                pass

        if pasta_cotas is None:
            print("    Pasta COTAS nao encontrada no Outlook.")
            return False

        msg.Move(pasta_cotas)
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


################################## CICLO PRINCIPAL ##################################

def processar_ciclo():
    """Um ciclo completo: ler emails, processar fundos novos, registrar."""

    agora = datetime.now().strftime('%H:%M:%S')
    print(f"\n{'='*60}")
    print(f"[{agora}] VERIFICANDO EMAILS DE APROVACAO...")
    print(f"{'='*60}")

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
            else:
                print(f"    [NOVO]     {f}")

        # 4. Chamar o mailer via subprocess
        fundos_str = ','.join(fundos_novos)
        resultado_path = os.path.join(DIRETORIO, "json", f"resultado_{data_json}_{datetime.now().strftime('%H%M%S')}.json")

        cmd = [
            sys.executable,
            SCRIPT_MAILER,
            '--data', info['data_completa'],
            '--fundos', fundos_str,
            '--resultado', resultado_path
        ]

        print(f"\n  Executando mailer para {len(fundos_novos)} fundo(s)...")
        print(f"  Comando: python mailer_v_auto.py --data {info['data_completa']} --fundos \"{fundos_str}\"")
        print()

        try:
            resultado = subprocess.run(cmd, timeout=600)  # 10 min timeout

            # 5. Ler resultado (quais fundos foram processados com sucesso)
            fundos_ok = []
            if os.path.exists(resultado_path):
                with open(resultado_path, 'r', encoding='utf-8') as f:
                    fundos_ok = json.load(f)
                # Limpar arquivo temporario
                os.remove(resultado_path)

            if fundos_ok:
                # Registrar como processados
                salvar_processados(data_json, fundos_ok)
                print(f"\n  {len(fundos_ok)} fundo(s) processado(s) com sucesso:")
                for f in fundos_ok:
                    print(f"    [OK] {f}")

                # Fundos que falharam
                fundos_falha = [f for f in fundos_novos if f not in fundos_ok]
                if fundos_falha:
                    print(f"\n  {len(fundos_falha)} fundo(s) com erro (serao retentados no proximo ciclo):")
                    for f in fundos_falha:
                        print(f"    [ERRO] {f}")
            else:
                print(f"\n  Nenhum fundo processado com sucesso neste ciclo.")

        except subprocess.TimeoutExpired:
            print(f"\n  TIMEOUT: mailer excedeu 10 minutos. Sera retentado no proximo ciclo.")
        except Exception as e:
            print(f"\n  ERRO ao executar mailer: {e}")

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


################################## MAIN ##################################

def main():
    print()
    print("=" * 60)
    print("  ROBO DE MAILERS - COTAS DIARIAS")
    print(f"  Intervalo: {INTERVALO_MINUTOS} minutos")
    print(f"  Mailer: {SCRIPT_MAILER}")
    print(f"  Diretorio: {DIRETORIO}")
    print("=" * 60)
    print()
    print("  O robo ira monitorar sua caixa de entrada do Outlook.")
    print("  Quando encontrar emails de aprovacao de carteiras,")
    print("  ira gerar os mailers automaticamente e abrir como")
    print("  rascunho para voce revisar e clicar ENVIAR.")
    print()
    print("  ADMs monitorados: Bradesco, BTG, BNYM, Itau, XP")
    print()
    print("  Para parar: Ctrl+C")
    print("=" * 60)

    while True:
        try:
            processar_ciclo()
        except KeyboardInterrupt:
            print("\n\nRobo encerrado pelo usuario.")
            break
        except Exception as e:
            print(f"\n  ERRO NO CICLO: {e}")
            traceback.print_exc()

        agora = datetime.now().strftime('%H:%M:%S')
        print(f"\n  [{agora}] Proximo ciclo em {INTERVALO_MINUTOS} minutos...")

        try:
            time.sleep(INTERVALO_MINUTOS * 60)
        except KeyboardInterrupt:
            print("\n\nRobo encerrado pelo usuario.")
            break


if __name__ == '__main__':
    main()
