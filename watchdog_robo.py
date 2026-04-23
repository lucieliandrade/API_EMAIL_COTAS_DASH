"""
watchdog_robo.py - Monitor separado do robo de mailers.

Se o robo_log.txt ficar sem atualizacao por mais de LIMITE_INATIVO_MIN minutos,
cria um rascunho de alerta no Outlook (Display, nunca Send) para avisar que
o robo parou. Re-alerta a cada REALERTA_MIN minutos enquanto continuar parado.

Roda em loop continuo (time.sleep entre checks). Subir via watchdog_robo.bat
na pasta Startup do Windows.

REGRA: processo SEPARADO do mailer_robo.py. Se o robo morrer, esse script
continua vivo e detecta.
"""

import time
import os
import re
import json
import sys
from datetime import datetime, timedelta
import win32com.client as win32

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(SCRIPT_DIR, "robo_log.txt")
ESTADO_PATH = os.path.join(SCRIPT_DIR, "watchdog_estado.json")

LIMITE_INATIVO_MIN = 10          # apos quantos minutos sem log do robo eh considerado parado
INTERVALO_CHECK_SEG = 120        # roda verificacao a cada 2 min
REALERTA_MIN = 15                # se continuar parado, novo rascunho a cada X min
EMAIL_ALERTA = "lucieli.andrade@capitaniainvestimentos.com.br"


def ler_ultimo_timestamp_log():
    """Retorna datetime do ultimo [HH:MM:SS] VERIFICANDO no log, ou None."""
    if not os.path.exists(LOG_PATH):
        return None
    try:
        with open(LOG_PATH, "r", encoding="utf-8", errors="ignore") as f:
            f.seek(0, 2)
            tamanho = f.tell()
            offset = max(0, tamanho - 10240)  # ultimos 10KB
            f.seek(offset)
            conteudo = f.read()
        ts_list = re.findall(r'\[(\d{2}:\d{2}:\d{2})\] VERIFICANDO', conteudo)
        if not ts_list:
            return None
        h, m, s = map(int, ts_list[-1].split(":"))
        ref = datetime.now().replace(hour=h, minute=m, second=s, microsecond=0)
        if ref > datetime.now():
            ref -= timedelta(days=1)
        return ref
    except Exception:
        return None


def carregar_estado():
    if not os.path.exists(ESTADO_PATH):
        return {"ultimo_alerta": None, "alertas_enviados": 0}
    try:
        with open(ESTADO_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"ultimo_alerta": None, "alertas_enviados": 0}


def salvar_estado(estado):
    with open(ESTADO_PATH, "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)


def criar_alerta_robo_morto(minutos_parado, tentativa=1):
    """Cria rascunho no Outlook avisando que o robo esta parado.
    REGRA: apenas Display() - nunca Send."""
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        if tentativa == 1:
            mail.Subject = "ALERTA: Robo de mailers PARADO"
            cabecalho = f"<p><b>O robo de mailers de cotas diarias nao atualiza o log ha {minutos_parado} minutos.</b></p>"
        else:
            mail.Subject = f"REITERACAO #{tentativa} - Robo de mailers continua PARADO"
            cabecalho = f"<p><b>Reiteracao ({tentativa} alertas hoje):</b> o robo continua sem atividade ha {minutos_parado} minutos.</p>"
        mail.To = EMAIL_ALERTA
        mail.HTMLBody = f"""
        {cabecalho}
        <p>Provaveis causas:</p>
        <ul>
            <li>Processo Python travado ou morto (checar Gerenciador de Tarefas)</li>
            <li>Outlook fechado ou sem permissao</li>
            <li>Drive de rede Z: ou X: desconectado</li>
            <li>Maquina reiniciou sem o .bat Startup subir</li>
        </ul>
        <p>Acao: abrir o Dash de Cotas (http://192.168.3.78:8502) ou reiniciar manualmente
        via mailer_robo.bat na pasta Startup.</p>
        """
        mail.Display()
        return True
    except Exception as e:
        print(f"  ERRO ao criar alerta: {e}")
        return False


def ciclo():
    ultimo = ler_ultimo_timestamp_log()
    agora = datetime.now()
    estado = carregar_estado()

    if ultimo is None:
        return  # sem log nenhum - nada a fazer

    inativo_min = (agora - ultimo).total_seconds() / 60

    if inativo_min > LIMITE_INATIVO_MIN:
        # Robo parado
        ultimo_alerta = estado.get("ultimo_alerta")
        if ultimo_alerta is None:
            alertar = True
            tentativa = 1
        else:
            dt_ultimo = datetime.fromisoformat(ultimo_alerta)
            alertar = (agora - dt_ultimo).total_seconds() / 60 >= REALERTA_MIN
            tentativa = estado.get("alertas_enviados", 0) + 1

        if alertar:
            if criar_alerta_robo_morto(int(inativo_min), tentativa):
                estado["ultimo_alerta"] = agora.isoformat(timespec='seconds')
                estado["alertas_enviados"] = tentativa
                salvar_estado(estado)
                print(f"[{agora:%H:%M:%S}] Alerta #{tentativa}: robo parado ha {int(inativo_min)} min")
    else:
        # Robo vivo. Se havia alerta anterior, reseta.
        if estado.get("ultimo_alerta") is not None:
            print(f"[{agora:%H:%M:%S}] Robo voltou (ultimo ciclo ha {int(inativo_min)} min). Resetando alertas.")
            salvar_estado({"ultimo_alerta": None, "alertas_enviados": 0})


def main():
    print(f"[{datetime.now():%H:%M:%S}] Watchdog do robo iniciado.")
    print(f"  Log monitorado: {LOG_PATH}")
    print(f"  Limite inatividade: {LIMITE_INATIVO_MIN} min")
    print(f"  Check a cada: {INTERVALO_CHECK_SEG}s")
    print(f"  Re-alerta a cada: {REALERTA_MIN} min")
    while True:
        try:
            ciclo()
        except Exception as e:
            print(f"[{datetime.now():%H:%M:%S}] Erro no ciclo: {e}")
        time.sleep(INTERVALO_CHECK_SEG)


if __name__ == "__main__":
    main()
