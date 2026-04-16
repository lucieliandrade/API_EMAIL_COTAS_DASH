"""
scan_outlook.py - Escaneia caixa de entrada do Outlook e gera JSON de aprovacoes
=====================================================================================
Gera aprovados_{YYYYMMDD}.json na pasta json/ com:
- site: fundos Site aprovados (email "Carteiras Aprovadas")
- manual: fundos Manual enviados (email "COTA DIARIA")

Executar periodicamente ou antes de atualizar o dash.
NAO envia nada. Apenas leitura.
"""

import win32com.client as win32
from datetime import datetime, timedelta
import re
import os
import json
import tempfile
import holidays
import pdfplumber

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_DIR = r"Z:\Relações com Investidores - NOVO\codigos\cotas\json"

MANUAIS_MAPA = {
    "FCopel":          {"pdf": "FCOPEL FIF MULTIMERCADO - CP I RL"},
    "FCopel_Imob":     {"pdf": "FCOPEL FIF MULTIMERCADO IMOB I RL"},
    "Sabesprev":       {"pdf": "SABESPREV CAPITÂNIA MERCADO IMOB. FIF MULT. CP RL"},
    "CAPITANIA REIT":  {"pdf": "CAPITÂNIA REIT MASTER FIC FIF MM RL"},
    "PETROS RFCP":     {"pdf": "FP FOF CAPITÂNIA FIF CI RF CP RL"},
    "OPOR IMOB FII":   {"pdf": "OPORTUNIDADES IMOBILIÁRIAS CONSOLIDADO"},
    "OPOR IMOB SUBCLA":{"pdf": "OPORTUNIDADES IMOBILIÁRIAS CL A"},
    "OPOR IMOB SUBCLB":{"pdf": "OPORTUNIDADES IMOBILIÁRIAS SUB CL B"},
    "OPOR IMOB SUBCLC":{"pdf": "OPORTUNIDADES IMOBILIÁRIAS SUB CL C"},
}

FUNDOS_SITE = {
    "BNYCL12879", "CSHG MAGIS II", "BNY12748", "BNYCL12975",
    "CAPIT D INC FIC", "PORTFOLIO FIDC", "CAPITANIA PREV BP",
    "CAPITANIA YIELD 120", "INFRA ADV CLA", "XP INFRA90",
    "CAPIT REIT FI", "CAPIT MULTIPREV", "CAPIT PREMIUM",
    "CAPIT PREV FDR", "CAPITANIA TOP",
}


def ref_de_hoje():
    """Retorna a data de referencia (D-1 util)."""
    hoje = datetime.today()
    feriados = holidays.country_holidays('BR', subdiv='SP', years=[hoje.year - 1, hoje.year, hoje.year + 1])
    prev = hoje - timedelta(days=1)
    while prev.weekday() >= 5 or prev.date() in feriados:
        prev -= timedelta(days=1)
    return prev.strftime("%Y%m%d")


MESES_PT = {
    'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04',
    'mai': '05', 'jun': '06', 'jul': '07', 'ago': '08',
    'set': '09', 'out': '10', 'nov': '11', 'dez': '12',
}


def validar_pdf_manual(msg, d_ref):
    """Valida se a data no nome do anexo e a primeira data dentro do PDF batem com D-1.
    Retorna (True, '') se ok, (False, motivo) se não."""
    if msg.Attachments.Count == 0:
        return False, "sem anexo"

    att = msg.Attachments.Item(1)
    nome = att.FileName

    # 1. Extrair data do nome do arquivo (YYYYMMDD)
    match_nome = re.search(r'(\d{8})\.pdf', nome, re.IGNORECASE)
    if not match_nome:
        return False, f"nome sem data: {nome}"
    data_nome = match_nome.group(1)

    if data_nome != d_ref:
        return False, f"nome={data_nome} esperado={d_ref}"

    # 2. Abrir PDF e extrair primeira data da tabela
    tmp = os.path.join(tempfile.gettempdir(), nome)
    try:
        att.SaveAsFile(tmp)
        with pdfplumber.open(tmp) as pdf:
            text = pdf.pages[0].extract_text() or ''
        # Procurar padrao DD-mmm-AA (ex: 15-abr-26)
        match_pdf = re.search(r'(\d{1,2})-(\w{3})-(\d{2})', text)
        if not match_pdf:
            return False, "data nao encontrada no PDF"

        dia_pdf = match_pdf.group(1).zfill(2)
        mes_pdf = MESES_PT.get(match_pdf.group(2).lower(), '00')
        ano_pdf = '20' + match_pdf.group(3)
        data_pdf = f"{ano_pdf}{mes_pdf}{dia_pdf}"

        if data_pdf != d_ref:
            return False, f"PDF={data_pdf} esperado={d_ref}"

        return True, ''
    except Exception as e:
        return False, f"erro leitura: {e}"
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def scan():
    d_ref = ref_de_hoje()
    hoje = datetime.today().date()

    outlook = win32.Dispatch('Outlook.Application')
    ns = outlook.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)
    msgs = inbox.Items
    msgs.Sort("[ReceivedTime]", True)

    site_aprovados = set()
    manual_enviados = set()
    manual_erros = {}

    for msg in msgs:
        try:
            if msg.ReceivedTime.date() < hoje:
                break
            if msg.ReceivedTime.date() != hoje:
                continue
            subj = str(msg.Subject)

            # Site: email "Carteiras Aprovadas"
            if 'Carteiras Aprovadas' in subj and 'Sistema Backoffice' in subj:
                body = str(msg.Body)
                match = re.search(r'Carteiras Aprovadas\s*(.*?)Atenciosamente', body, re.DOTALL)
                if match:
                    linhas = match.group(1).strip().split('\n')
                    for l in linhas:
                        f = l.strip()
                        if f and f in FUNDOS_SITE:
                            site_aprovados.add(f)

            # Manual: email "COTA DIARIA" + validação do PDF anexo
            if 'COTA' in subj.upper() and msg.Attachments.Count > 0:
                for i in range(1, msg.Attachments.Count + 1):
                    att_nome = msg.Attachments.Item(i).FileName.upper()
                    for nome_curto, mapa in MANUAIS_MAPA.items():
                        if mapa["pdf"].upper() in att_nome:
                            ok, motivo = validar_pdf_manual(msg, d_ref)
                            if ok:
                                manual_enviados.add(nome_curto)
                            else:
                                manual_erros[nome_curto] = motivo
                                print(f"  BLOQUEADO: {nome_curto} - {motivo}")

        except:
            continue

    resultado = {
        "data_ref": d_ref,
        "scan_hora": datetime.now().strftime("%H:%M:%S"),
        "site": sorted(site_aprovados),
        "manual": sorted(manual_enviados),
        "manual_erros": manual_erros,
    }

    os.makedirs(JSON_DIR, exist_ok=True)
    path = os.path.join(JSON_DIR, f"aprovados_{d_ref}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print(f"Scan concluido - ref {d_ref}")
    print(f"  Site aprovados: {len(site_aprovados)} - {sorted(site_aprovados)}")
    print(f"  Manual enviados: {len(manual_enviados)} - {sorted(manual_enviados)}")
    print(f"  Salvo em: {path}")

    return resultado


if __name__ == "__main__":
    scan()
