# LEIA-ME — Mapa da pasta API_EMAIL_COTAS_DASH

> Guia rápido de **o que é cada arquivo** e **o que está no ar**.
> Última atualização: 2026-07-09.

## Como saber o que está rodando agora
No PowerShell:
```powershell
Get-CimInstance Win32_Process -Filter "name='python.exe' OR name='streamlit.exe'" |
  Select-Object ProcessId, @{n='Cmd';e={$_.CommandLine}} | Format-List
```

---

## 🟦 DASH DE COTAS (o que você abre no 8502)
| Arquivo | Serve para | Status |
|---|---|---|
| **status_mailers_v2.py** | O dash de cotas que está no ar (porta 8502) | ✅ **NO AR** |
| status_mailers_v3.py | Versão NOVA, em teste (sobe só no 8504) | ⚠️ teste — não está no ar |
| watchdog_dash.bat | Mantém o dash no ar (reinicia se cair). Vai no Startup do Windows | ✅ ligador oficial (roda o v2/8502) |
| reiniciar_dash.bat | Subir o dash na mão | ⛏️ manual (v2/8502) |
| auto_dash.bat | Jeito antigo de subir | 🕸️ desativado |
| dash_v2_nova.bat | Subir a versão nova (v3) no 8504 para testar | 🧪 teste |
| dash_watchdog_log.txt | Log do watchdog do dash | 🗑️ log |

> ⚠️ **Pegadinha dos nomes:** o arquivo se chama `v3`, mas o `.bat` o chama de "versão nova (v2)". O que roda no 8502 é o **v2**. O `v3` é a versão em desenvolvimento.

---

## 🟩 ROBÔ DE MAILERS (gera os rascunhos das cotas)
| Arquivo | Serve para | Status |
|---|---|---|
| **mailer_robo.py** | O robô (orquestra: lê e-mails de aprovação, chama o mailer, controla estado) | ✅ rodando |
| **mailer_v_auto.py** | O mailer de verdade (calcula e cria o rascunho). **É do Matheus** | ⛔ NUNCA mexer |
| watchdog_robo.py + watchdog_robo.bat | Mantêm o robô vivo (reinicia se cair) | ✅ rodando |
| scan_outlook.py | Detecta cotas aprovadas/enviadas manualmente (site/e-mail) | ✅ usado |
| robo_log.txt | Log do robô | 🗑️ log (costuma ficar ENORME) |
| mailer_robo.lock | Trava para não rodar dois robôs ao mesmo tempo | ⚙️ estado |

> Regra de ouro: o robô **nunca envia** — só cria RASCUNHO (Display). Você revisa e clica Enviar.

---

## 🟨 DASH DE ROTINAS (checklist diário, porta 8503)
| Arquivo | Serve para | Status |
|---|---|---|
| dash_rotinas.py | O app do checklist diário | 💤 sobe via watchdog_rotinas.bat |
| rotinas_checklist.py | A lista de itens do checklist (dados) | ✅ usado pelo app |
| dash_rotinas_preview.py | Preview de layouts (para escolher o visual) | 🧪 auxiliar |
| watchdog_rotinas.bat | Mantém o dash de rotinas no ar | ⚙️ ligador |
| rotinas_estado/ | Estado diário do checklist (marcações) | ⚙️ estado |
| rotinas_watchdog_log.txt | Log | 🗑️ log |

---

## 🟥 REDE / INFRA (executar como ADMINISTRADOR)
| Arquivo | Serve para |
|---|---|
| fixar_ip_dash.bat | Fixa o IP da máquina em 192.168.3.83 (o time sempre acha o dash) |
| voltar_ip_dhcp.bat | Reverte para IP automático (DHCP) |

---

## 📄 DOCUMENTAÇÃO / CONFIG
Os textos de leitura ficam na subpasta **`documentacao/`**:
| Arquivo (em documentacao/) | Serve para |
|---|---|
| MANUAL_DASH_COTAS.md | Manual passo a passo do dash de cotas |
| EXPLICACAO_MUDANCAS_2026-07-09.txt | O que mudou no robô/dash e por quê (histórico) |
| PLANO_MELHORIAS_SEGUNDA.md | Plano de melhorias (histórico) |
| homologacao processo completo.txt | Anotação de teste de homologação |

Na raiz ficam só os de config e este mapa:
| Arquivo | Serve para |
|---|---|
| requirements.txt | Dependências Python (o que instalar) |
| .gitignore | O que o Git deve ignorar |
| LEIA-ME.md | Este mapa |

---

## 🗑️ DESCARTÁVEL / ESTADO (não é "código")
- **Lixo/** — versões antigas e backups (já ignorado pelo Git).
- **__pycache__/** — cache do Python (recriado sozinho).
- **watchdog_estado.json** — estado interno do watchdog do dash.

---

## Endereços úteis
- Dash de cotas (time): **http://RI011:8502** — nome fixo da máquina, NÃO muda quando o IP troca (DHCP). Use sempre este.
- Dash de cotas (na própria máquina): http://localhost:8502
- Fallback se o nome não resolver: http://RI011.CAPITANIA.LOCAL:8502 ou o IP do dia
- Dash de rotinas: http://localhost:8503
- Dash de cotas NOVO (teste): http://localhost:8504
