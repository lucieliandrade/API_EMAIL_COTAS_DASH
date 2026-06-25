# 📬 Manual do Dash de Cotas Diárias

> **Para quem é este manual:** qualquer pessoa da equipe de RI, mesmo sem saber
> programar. Ele explica, em linguagem simples e passo a passo, **o que o dash
> faz, de onde ele lê as informações e o que cada parte da tela significa**.
>
> Arquivo do código: [`status_mailers_v2.py`](status_mailers_v2.py) (versão no ar)
> e [`status_mailers_v3.py`](status_mailers_v3.py) (versão nova, em testes).

---

## 1. O que é este dash?

É um **painel web** que mostra, em tempo quase real, o status do envio diário das
**cotas dos fundos** para os investidores. Em vez de ficar conferindo e-mail por
e-mail, a pessoa abre uma página no navegador e vê de relance:

- ✅ Quais fundos **já tiveram a cota enviada** hoje (e a que horas).
- ❌ Quais fundos **ainda estão pendentes**.
- ⏳ Quais estão **esperando a cota ser lançada no banco** (COTAS_CAP).
- 🚨 Quais deram **problema** e precisam de revisão humana.

Além das cotas, o painel também acompanha duas rotinas extras:
- **Envio Diário de XMLs** (ICATU / Aquila / BASF).
- **Esteira INTRAG** (boletas do Itaú Vida, em 7 etapas).

### Quem faz o trabalho de verdade?

O dash **não envia e-mail nenhum** e **não decide nada sozinho**. Ele é apenas um
"painel de instrumentos" — como o painel de um carro. Quem dirige são **outros
programas (os "robôs")** que rodam em segundo plano:

| Programa | O que faz |
|---|---|
| `scan_outlook.py` | Lê o Outlook procurando aprovações e gera arquivos `.json` |
| `mailer_robo.py` | Processa as cotas e gera os PDFs/e-mails |
| `watchdog_robo.py` | Vigia se o robô principal travou |
| **o dash** (este código) | **Só lê** os arquivos que os robôs deixaram e mostra na tela |

> 🔑 **Conceito central:** o dash funciona como um **leitor de arquivos**. Os robôs
> escrevem o que aconteceu em arquivos (`.json`, `.pdf`); o dash abre esses arquivos
> e desenha a tela. Se um robô está parado, o dash mostra "parado", mas ele mesmo
> continua funcionando normalmente.

---

## 2. Como abrir o dash

- **Endereço no navegador:** `http://192.168.3.81:8502/` (ou `192.168.3.83:8502`)
- O dash roda numa **máquina servidora** que fica sempre ligada. Ele é iniciado
  automaticamente pelo arquivo [`auto_dash.bat`](auto_dash.bat).
- A tecnologia por trás chama-se **Streamlit** (uma ferramenta que transforma
  código Python em página web).

### Login

Ao abrir, aparece uma tela de login:
- **Usuário:** `RI`
- **Senha:** `Capitania2025!`

> A janela dedicada que fica aberta na máquina servidora entra **sozinha**, sem
> digitar senha, porque usa um "crachá" especial na URL (`?k=ri-dash-local`).
> Qualquer outro computador precisa digitar usuário e senha normalmente.

### A página se atualiza sozinha

A cada **2 minutos** o dash recarrega automaticamente para mostrar dados novos.
Você também pode forçar a atualização clicando no botão **🔄 Atualizar**.

---

## 3. A tela, de cima para baixo

A tela é montada nesta ordem. Vamos por partes.

### 3.1. Cabeçalho azul

Mostra o título "📬 Mailers · Cotas Diárias" e, à direita, a **semana que está
sendo exibida** (ex.: `23/06 — 27/06/2026`).

### 3.2. Status do robô (a "bolinha")

Logo abaixo do cabeçalho há uma **bolinha colorida** indicando se o robô principal
está trabalhando:

| Cor | Significado | Quando aparece |
|---|---|---|
| 🟢 Verde "Ativo" | Tudo certo | Robô deu sinal de vida há menos de 5 min |
| 🟡 Amarelo "Lento" | Atenção | Última atividade entre 5 e 10 min atrás |
| 🔴 Vermelho "Parado" | **Problema!** | Mais de 10 min sem atividade |

> Como ele sabe disso? O robô escreve a hora de cada verificação no arquivo
> `robo_log.txt`. O dash pega a **última hora registrada** e compara com a hora
> atual. Se a diferença for grande, mostra "Parado".

**Se aparecer "Parado"**, surge um **banner vermelho grande** explicando o que fazer:
verificar se o programa Python está ativo no Gerenciador de Tarefas, ou reiniciar
pelo `mailer_robo.bat`. O watchdog também cria um rascunho de alerta no Outlook.

### 3.3. Log do robô (expansível)

Um item "📋 Log do robô" que, ao clicar, mostra as **últimas 30 linhas** do que o
robô registrou. Útil para entender o que aconteceu sem abrir o arquivo na mão.

### 3.4. Botões de navegação

- **◀ Semana anterior** / **Semana seguinte ▶** — navega entre semanas.
- **🏠 Semana atual** — volta para a semana de hoje (só aparece se você saiu dela).
- **🔄 Atualizar** — recarrega os dados na hora (limpa o cache).

### 3.5. Banners de alerta (só aparecem quando há problema)

Aparecem em vermelho/laranja **apenas se houver algo a resolver hoje**:

- 🚨 **Tentativas órfãs** — o robô começou a processar um fundo mas foi
  interrompido antes de terminar. Precisa de revisão humana (conferir o Outlook).
- ⏳ **Aguardando cota no banco COTAS_CAP** — a cota de um ou mais fundos ainda
  não foi lançada no banco. Mostra há quantos minutos está esperando. Depois de
  **25 minutos**, o robô cria sozinho um rascunho de cobrança no Outlook.

### 3.6. Cartões dos 5 dias da semana

Cinco cartões (Segunda a Sexta) com um **resumo visual de cada dia**:

| Cor do cartão | Significado |
|---|---|
| 🟢 Verde "Completo" | Todos os fundos enviados |
| 🔵 Azul "X pendentes" | Hoje, ainda faltam alguns |
| 🟡 Amarelo "X pendentes" | Outro dia, faltaram alguns |
| 🔴 Vermelho "Nenhum enviado" | Nenhum fundo saiu nesse dia |
| 🟣 Roxo "🏖️ Feriado" | Dia sem mercado (sem cota) |
| Cinza "Aguardando" | Dia futuro (ainda não chegou) |

Cada cartão mostra a contagem (ex.: `45 / 60`) e uma barrinha de progresso.

### 3.7. Seção "📤 Envio Diário · XMLs Mellon" (expansível)

Acompanha o envio dos XMLs para **ICATU, Aquila 6 e 7, e BASF**. Para cada cliente
mostra uma "luz" de status:

| Status | Significado |
|---|---|
| 📧 ENVIADO | Já saiu (confirmado na caixa de entrada) |
| ✅ PRONTO | Todos os arquivos chegaram — pode enviar |
| ⏳ AGUARDANDO | Ainda faltam arquivos (mostra quais) |
| ⚠️ ERRO | A pasta de rede está offline |

Quando está **PRONTO**, aparece o botão **"📧 Abrir rascunho"**. Ele abre o e-mail
já montado (com anexos) no Outlook — **mas nunca envia sozinho**. Você revisa e
envia. Assim que enviar, a cópia cai na caixa do `invest@` e o card vira
**ENVIADO** automaticamente.

> Você pode mudar a **data de referência** no seletor de data dentro dessa seção.
> Por padrão ele usa o último dia útil (D-1), pulando fins de semana e feriados.

### 3.8. Seção "🏦 Esteira INTRAG · Boletas Itaú Vida" (expansível)

Acompanha as boletas do Itaú Vida em **7 etapas (steps)**, da chegada do e-mail do
Itaú até o arquivo final aparecer na pasta de rede:

1. **Email Itaú** (automático) — chegou o e-mail?
2. **TXTs gerados** (automático) — os 3 arquivos TXT foram criados?
3. **Passivo Itaú → FIE** (manual — marcar com ✔)
4. **Ativo FIE → FIFE** (manual)
5. **Passivo FIE → FIFE** (manual)
6. **Liquidação** (manual)
7. **Arquivo na pasta net** (automático) — o arquivo final apareceu?

As etapas 3 a 6 têm uma **caixinha "feito"** que a pessoa marca conforme executa.
As demais o dash detecta sozinho. O título da seção mostra o resumo (ex.:
`⏳ 4/7 etapas concluídas`).

### 3.9. Filtros

Uma barra cinza com quatro filtros para a tabela grande abaixo:
- **ADM** — filtra por administrador (Itaú, BNYM, XP...).
- **Tipo** — Auto, Site ou Manual (explicado na seção 5).
- **Status hoje** — Todos / Enviado hoje / Pendente hoje.
- **Buscar fundo** — digite parte do nome para encontrar.

### 3.10. Tabela grande (fundo × dia)

O coração do dash. Cada **linha é um fundo** e cada **coluna é um dia da semana**.
A célula mostra o que aconteceu com aquele fundo naquele dia:

| O que aparece | Significado |
|---|---|
| ✅ 14:30 | Enviado, com a hora |
| ⚠️ 24/06 09:00 | Enviado **com atraso** (gerado depois do dia certo) |
| ❌ | Pendente / não enviado |
| ❌ Data errada (...) | Problema: o PDF tem data diferente da esperada |
| ⏳ aguardando COTAS_CAP 12min | Esperando a cota no banco |
| 🚨 ORFA - revisar Outlook | Travou no meio — revisar |
| ENVIAR | Fundo manual aprovado, pronto para você enviar |
| 🏖️ | Feriado |
| · | Dia futuro |

A coluna do **dia de hoje fica destacada** com uma borda azul. O número de fundos
exibidos aparece logo acima da tabela.

### 3.11. Pendentes de hoje

Um item expansível que lista, em "etiquetas", **todos os fundos que ainda não
foram enviados hoje**. Se estiver tudo enviado, mostra ✅ "Todos os fundos foram
enviados hoje."

### 3.12. Rodapé

Mostra a data/hora da última atualização.

---

## 4. De onde vêm os dados? (as "fontes da verdade")

O dash lê de várias pastas de rede. **Se uma dessas pastas estiver fora do ar, a
parte correspondente do dash fica vazia ou mostra erro** — não é bug do dash, é a
fonte que sumiu.

| O que | Onde fica | Quem escreve |
|---|---|---|
| Lista de fundos | `X:\...\Tipo_Fundos.xlsx` | Equipe (planilha) |
| Aprovações/erros do dia | `Z:\...\cotas\json\*.json` | `scan_outlook.py` |
| PDFs das cotas enviadas | `Z:\...\cotas\PDFs\` | `mailer_robo.py` |
| Log do robô | `robo_log.txt` (junto do código) | `mailer_robo.py` |
| XMLs Mellon | `X:\RI + BACK - PILOTO XML\...` | Sistema Mellon |
| Esteira INTRAG | `Z:\...\Boletas Fundos\INTRAG\` | Robô da esteira |

### Como o dash sabe a "data de referência" de cada coluna

Detalhe importante e que costuma confundir: **a coluna é o dia do envio, mas a cota
é sempre do dia útil anterior (D-1)**.

> Exemplo: na coluna **Quarta 01/04**, o dash procura o arquivo do dia **31/03**
> (D-1), porque a cota enviada na quarta é a do fechamento de terça.
> Na **Segunda**, o D-1 é a **sexta** anterior (pula o fim de semana).

Para acertar isso, o dash usa o **calendário oficial da B3/BVMF** (feriados de
mercado), assim ele pula corretamente feriados e fins de semana ao calcular o D-1.

---

## 5. Os três tipos de fundo

Cada fundo tem um **Tipo**, que muda a forma como o dash confirma o envio:

- **Auto** — o robô gera e o dash confirma sozinho pelo PDF que aparece na pasta.
- **Site** — a cota vem de um sistema/site externo. O dash confirma via aprovação
  registrada pelo `scan_outlook`. *(Atenção: aprovação no site **não** significa
  envio se o fundo ainda estiver "aguardando" cota no banco.)*
- **Manual** — alguém da equipe envia na mão. Quando o `scan_outlook` detecta o
  e-mail "COTA DIÁRIA" aprovado, o dash marca como pronto e mostra **ENVIAR**.

---

## 6. Perguntas frequentes / o que fazer quando...

**A tela está vazia ou dando erro de pasta.**
Provavelmente uma pasta de rede (Z:, X:) está fora do ar ou você está sem acesso a
ela. Confira se consegue abrir a pasta normalmente pelo Explorer.

**A bolinha do robô está vermelha ("Parado").**
O robô principal travou. Siga o banner vermelho: verifique o processo Python no
Gerenciador de Tarefas ou reinicie pelo `mailer_robo.bat`. Veja também o rascunho
de alerta no Outlook.

**Um fundo está "⏳ aguardando COTAS_CAP" há muito tempo.**
A cota ainda não foi lançada no banco. Depois de 25 min o robô cria um rascunho de
cobrança sozinho. Se ficar preso indevidamente, edite o arquivo
`aguardando_<data>.json` na pasta de JSONs (**nunca** o código do dash).

**Apareceu "🚨 ORFA".**
O robô começou e parou no meio. Confira no Outlook se o rascunho foi aberto. Para
destravar, apague a entrada correspondente no arquivo `tentativas_<data>.json`.

**Quero conferir uma semana passada.**
Use os botões ◀ / ▶ de navegação de semana.

**Mudei a planilha de fundos e o dash não atualizou.**
Clique em **🔄 Atualizar** (a lista de fundos fica em cache para o dash ficar rápido).

---

## 7. Resumo em uma frase

> O dash é um **painel que lê arquivos** deixados pelos robôs e pinta a tela de
> verde/amarelo/vermelho para você ver, num relance, **o que já saiu, o que falta
> e o que deu problema** no envio das cotas — sem nunca enviar nada sozinho.

---

*Documento mantido pela equipe de Relações com Investidores · Capitânia
Investimentos. Para detalhes técnicos linha a linha, veja os comentários dentro do
próprio código.*
