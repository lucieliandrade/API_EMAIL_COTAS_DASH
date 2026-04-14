# Sistema de Mailers de Cotas Diarias
## Capitania Investimentos - Relacoes com Investidores
## Atualizado em 14/04/2026

---

## Pasta do projeto
C:\Users\lucieli.andrade\OneDrive - Capitania S.A\DASH_2026\API_EMAIL_COTAS_DASH\

## GitHub
https://github.com/lucieliandrade/API_EMAIL_COTAS_DASH

---

## Codigos

### 1. mailer_robo.py - Robo principal
- Roda a cada 2 minutos
- Le emails de aprovacao na caixa de entrada
- Chama o mailer_v_auto.py para gerar rascunhos
- Move emails processados para RI_MIDDLE > COTAS
- Chama scan_outlook.py para atualizar o dash

### 2. mailer_v_auto.py - Gerador de mailers
- Recebe --data e --fundos como argumentos
- Le cotas do banco SQL (rds01.capitania.net)
- Le benchmarks (CDI, IPCA, IFIX, IMA-B)
- Gera DataFrame com tabela de rentabilidade
- Valida: NaN, zeros, batimento, PL
- Escreve no template Excel, converte para PDF
- Abre rascunho no Outlook (NUNCA envia)

### 3. scan_outlook.py - Scanner de aprovacoes
- Le caixa de entrada do Outlook
- Detecta emails "Carteiras Aprovadas" (Site) e "COTA DIARIA" (Manual)
- Gera JSON em Z:\...\cotas\json\aprovados_YYYYMMDD.json
- So leitura, nunca envia nada

### 4. status_mailers.py - Dashboard Streamlit
- Porta 8502 (ou a que estiver ativa)
- Mostra status de todos os fundos (Auto/Manual/Site)
- Le PDFs de Z:\...\cotas\PDFs\ (Auto)
- Le PDFs de X:\...\NOVO Mailer\ (Manual)
- Le JSON de aprovacoes (Site)
- Botao Exportar Excel
- Auto-refresh a cada 2 minutos

---

## Fluxo a cada 2 minutos

1. Robo escaneia caixa de entrada procurando emails "Carteiras Aprovadas - Fundos [ADM] - Sistema Backoffice"
2. Atualiza o JSON do dash (scan_outlook) - so leitura
3. Para cada fundo aprovado:
   - Verifica se ja foi processado (controle por JSON de duplicidade)
   - Verifica cotas no banco (COTAS_CAP)
   - Verifica batimento
   - Verifica PL
   - Bloqueia se algum valor for 0.00%
   - Bloqueia se tiver NaN
   - Gera o PDF a partir do template
   - Abre rascunho no Outlook (NUNCA envia)
4. Quando TODOS os fundos de um email foram processados, move o email para RI_MIDDLE > COTAS
5. Fundos que falharam ficam na caixa de entrada e sao retentados no proximo ciclo

---

## 3 Tipos de fundos no dash

- Auto = robo gera rascunho. Check quando PDF existe na pasta cotas/PDFs
- Manual = enviado por fora. Check quando PDF existe na pasta NOVO Mailer ou email COTA DIARIA na inbox
- Site = publicado no site. Com template: robo gera rascunho. Sem template: check via email aprovacao (JSON)

---

## Fundos Manual (9)

| Nome no dash | Nome formal (pasta/PDF) | ADM |
|---|---|---|
| FCopel | FCopel FIM CP / FCOPEL FIF MULTIMERCADO - CP I RL | Itau |
| FCopel_Imob | FCOPEL FIF MULTIMERCADO IMOB I RL | Itau |
| Sabesprev | SABESPREV CAPITANIA MERCADO IMOB. FIF MULT. CP RL | Itau |
| CAPITANIA REIT | CAPITANIA REIT MASTER FIC FIF MM RL | BNYM |
| PETROS RFCP | FP FOF CAPITANIA FIF CI RF CP RL | Bradesco |
| OPOR IMOB FII | OPORTUNIDADES IMOBILIARIAS CONSOLIDADO | XP |
| OPOR IMOB SUBCLA | OPORTUNIDADES IMOBILIARIAS SUB CL A / CL A | XP |
| OPOR IMOB SUBCLB | OPORTUNIDADES IMOBILIARIAS SUB CL B | XP |
| OPOR IMOB SUBCLC | OPORTUNIDADES IMOBILIARIAS SUB CL C | XP |

## Fundos Site com template (10)
BNYCL12879, CSHG MAGIS II, BNY12748, BNYCL12975, CAPIT D INC FIC, PORTFOLIO FIDC, CAPITANIA PREV BP, CAPITANIA YIELD 120, INFRA ADV CLA, XP INFRA90

## Fundos Site sem template (5)
CAPIT MULTIPREV, CAPIT PREMIUM, CAPIT PREV FDR, CAPIT REIT FI, CAPITANIA TOP

## Fundos que NAO sao do time RI
CAPITANIA INFRAFIC, CPDI F, INFRA Y FIC, CAPITANIA INFRA4/5/6, CPDI_1

---

## Pastas de dados

| Pasta | Conteudo |
|---|---|
| Z:\Relacoes com Investidores - NOVO\codigos\cotas\PDFs\ | PDFs automaticos |
| Z:\Relacoes com Investidores - NOVO\codigos\cotas\templates\ | Templates Excel por fundo |
| Z:\Relacoes com Investidores - NOVO\codigos\cotas\json\ | JSONs de controle e aprovacoes |
| X:\#CapitaniaRFE\Operational\CapitaniaMailer\NOVO Mailer\ | PDFs manuais (9 subpastas) |
| X:\BDM\Novo Modelo de Carteiras\Tipo_Fundos.xlsx | Lista de fundos, ADM, modelo_mailer |
| X:\Leonardo Salles\Gestao\Relatorio\D_Uteis.xlsx | Dias uteis |

---

## Regras de seguranca

- NUNCA email.Send() - sempre Display() (rascunho)
- NUNCA gerar email com valor 0.00% (dado faltando)
- NUNCA gerar email com NaN
- Batimento errado = bloqueado
- PL zerado = bloqueado
- Labels duplicados = bloqueado

---

## Como iniciar

1. Abrir terminal na pasta do projeto
2. Rodar: python mailer_robo.py
3. O robo fica ativo a cada 2 minutos
4. O dash roda separado: streamlit run status_mailers.py --server.port 8502
