# Plano de Melhorias - Mailers de Cotas (Segunda, 27/04/2026)

**Contexto:** após uma semana operando o robô com idempotência, watchdog e banner de órfãs,
os erros recorrentes agora são quase 100% upstream (fontes de dados ou arquivos externos).
Este plano ataca os 3 pontos de maior retrabalho.

---

## 🟥 Melhoria 1 — Pre-check de sanidade (7h, antes do expediente)

### Problema atual
Problemas de infraestrutura só são descobertos quando o robô tenta processar o 1º fundo
do dia. Ex: drive Z: desconectado, Outlook fechado, banco ODBC off, template xlsx
corrompido. Resultado: erros em cascata, atraso, retrabalho.

### O que o pre-check faz
Novo script `precheck_robo.py` agendado via Task Scheduler para rodar **toda manhã às 7h**
(antes do expediente começar). Executa:

1. **Conexão ODBC ao banco:** SELECT TOP 1 de COTAS_CAP. Se falhar, reporta.
2. **Drive Z:** lista pasta `Z:\Relações com Investidores - NOVO\codigos\cotas\json`.
3. **Drive X:** lista pasta `X:\BDM\Novo Modelo de Carteiras\Tipo_Fundos.xlsx`.
4. **Outlook acessível:** tenta `win32.Dispatch('Outlook.Application').GetNamespace("MAPI")`.
5. **Templates:** tenta carregar cada `templates/*.xlsx` com openpyxl. Lista os que dão erro.
6. **Cotas do dia:** verifica se a cota de D-1 já foi lançada em COTAS_CAP (pelo menos parcialmente).

### Saída
- Se tudo OK: log silencioso em `precheck_log.txt`.
- Se algo falhar: **rascunho no Outlook** endereçado a você com
  `ALERTA MATINAL: 3 problemas detectados antes do expediente`,
  listando cada item e ação sugerida.

### Como instalar
Adicionar `precheck_robo.bat` no Task Scheduler do Windows com trigger diário às 7h
(segunda a sexta, exceto feriados).

### Esforço estimado
2h de implementação.

### Arquivos a criar
- `precheck_robo.py`
- `precheck_robo.bat`
- Tarefa no Task Scheduler

---

## 🟥 Melhoria 2 — Detectar e reparar arquivo corrompido automaticamente

### Problema atual
Quando o BNYM envia carteira `.xlsx` com CRC inválido (como BNY11585 em 23/04), o mailer
crasha com `BadZipFile`. Fica órfão, exige intervenção manual (Claude reparar + re-rodar).

### O que a mitigação faz
Modificar `mailer_v_auto.py` para:

1. Antes de ler qualquer `.xlsx` via openpyxl/pandas, testar com `zipfile.ZipFile.read()`
   em todos os membros internos.
2. Se detectar `BadZipFile`:
   - Tentar reparar automaticamente copiando o elemento corrompido (ex: `xl/theme/theme1.xml`)
     de outro arquivo similar que abra com sucesso (ex: outra carteira BNYM do mesmo dia).
   - Salvar backup do original com sufixo `.backup.YYYYMMDD-HHMMSS`.
   - Logar: "Reparei arquivo X substituindo componente Y de arquivo Z".
   - Continuar processamento normalmente.
3. Se reparo falhar: reportar motivo específico no JSON de resultado
   (ex: `"carteira corrompida e sem fonte alternativa para reparo"`).

### Onde aplicar
- Função `cota_carteira()` (linha ~230) antes do `pd.read_excel`
- Função `mailer()` (linha ~2170) antes do `openpyxl.load_workbook`

### Esforço estimado
3h. Adiciona ~80 linhas de código.

### Risco
Baixo. Se reparo falhar, cai no fluxo de erro atual. Se reparo der certo, log mostra
exatamente o que foi feito.

### Arquivos a editar
- `mailer_v_auto.py` (nova função `_xlsx_seguro()` que envolve leitura)

---

## 🟥 Melhoria 3 — Mensagens de erro específicas por fonte de dados

### Problema atual
Hoje o mailer reporta `"Data X não consta no COTAS_CAP"` mesmo quando o problema real é:
- Cota ausente em **Cotas_Ret_Ajus** (tabela de cota ajustada — caso CAPITANIA CORP FIDC
  de 17/04 que investigamos ontem)
- Cota **igual a 0** em COTAS_CAP (dado parcial)
- **Benchmark** IMAB/IMAB5/IPCA sem dado no dia
- **Carteira** xlsx não existe no drive

Você perde tempo investigando qual tabela cobrar.

### O que a mitigação faz
Refatorar a função de validação do mailer para reportar o nome **exato** da tabela/fonte
que falhou:

| Situação atual (mensagem) | Nova mensagem |
|---|---|
| `Data X não consta no COTAS_CAP` (para CAPITANIA CORP FIDC) | `Cota Ajustada não consta em Cotas_Ret_Ajus para 17/04` |
| `Data X não consta no COTAS_CAP` (cota normal) | `Cota não consta em COTAS_CAP para 23/04` |
| `Tabela do imab sem dados para o dia` | `Benchmark IMAB sem dado em 23/04` |
| `Cota de X igual a 0 ou NaN no COTAS_CAP` | `Cota zerada/nula em COTAS_CAP para 23/04 (dado parcial)` |
| `Carteira não bate com COTAS_CAP` | `Carteira não bate: PL calculado 1.234.567, COTAS_CAP 1.234.500 (dif 67)` |

### Onde aplicar
- Função `check_cotas()` (linha ~1680)
- Função `check_bench()` (linha ~1550 aprox.)
- Função `batimento()` (linha ~1530)
- Função `mailer()` (linha ~2090) — passa a mensagem específica para `_erros[f]`

### Impacto
- Você saberá **exatamente** qual time cobrar (backoffice vs indicadores vs BNYM).
- O dash vai exibir a mensagem nova na célula de erro/aguardando.

### Esforço estimado
2h. Mudança controlada (só troca texto de mensagens + desambiguação de branch).

---

## 🟩 Resumo executivo

| # | Melhoria | Esforço | Impacto |
|---|---|---|---|
| 1 | Pre-check 7h | 2h | Evita erros em cascata no expediente |
| 2 | Auto-reparo xlsx corrompido | 3h | Elimina intervenção manual do Claude |
| 3 | Mensagens específicas por fonte | 2h | Cobra o time certo direto |
| **Total** | | **~7h** | |

Recomendação: implementar na ordem (1 → 3 → 2). Fazer em commits separados com
testes isolados. Reiniciar o robô uma vez ao final.

---

## Próximos ciclos (depois dessas 3)

- **Classificação automática de fundos novos** — sugerir via rascunho quando um fundo
  aparece pela 1ª vez sem estar em nenhuma lista.
- **Botão "destravar órfã" no dash** — evita edição manual de JSON.
- **Relatório semanal automático toda segunda 9h** — já tem script pronto (`_relatorio.py`),
  só adicionar em Task Scheduler.
- **Notificação push (toast Windows)** — além do rascunho no Outlook.
- **Log estruturado pesquisável** — hoje o log é texto plano.

---

_Plano gerado em 24/04/2026. Revisar na manhã de segunda (27/04) antes de implementar._
