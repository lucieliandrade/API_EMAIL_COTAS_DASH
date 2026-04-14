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
import holidays

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

            # Manual: email "COTA DIARIA"
            if 'COTA' in subj.upper():
                for nome_curto, mapa in MANUAIS_MAPA.items():
                    if mapa["pdf"].upper() in subj.upper():
                        manual_enviados.add(nome_curto)

        except:
            continue

    resultado = {
        "data_ref": d_ref,
        "scan_hora": datetime.now().strftime("%H:%M:%S"),
        "site": sorted(site_aprovados),
        "manual": sorted(manual_enviados),
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
