# Dados compartilhados do dash de rotinas (checklist diario).
# Mantido separado para o app principal e o preview de layout usarem a MESMA lista.

OPCOES = ["Pendente", "Feito", "N/A"]
ICONE_ESTADO = {"Pendente": "⬜", "Feito": "✅", "N/A": "➖"}
DIAS_PT = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira",
           "Sexta-feira", "Sábado", "Domingo"]

# Ordem dos itens e dos ids e FIXA: os ids viram chave no JSON de estado.
CHECKLIST = [
    {"secao": "DIARIAMENTE", "icone": "🗓️", "itens": [
        {"id": "d1", "txt": "Liquidações e-mail"},
        {"id": "d2", "txt": "Boletar no administrador (resgates ref. e-mail liquidações)"},
        {"id": "d3", "txt": "PL fundos XP"},
        {"id": "d4", "txt": "Cotas fundos listados"},
        {"id": "d5", "txt": "Checar saldos nos FICs para pagamento dos resgates do dia"},
    ]},
    {"secao": "MOVIMENTAÇÕES", "icone": "🔄", "itens": [
        {"id": "m1",  "txt": "ICATU — mandamos tela CETIP?"},
        {"id": "m2",  "txt": "ICATU — BRE boletou?"},
        {"id": "m3",  "txt": "Zerou CREDPREV?"},
        {"id": "m4",  "txt": "PREV XP"},
        {"id": "m5",  "txt": "PREV XP — BRE boletou?"},
        {"id": "m6",  "txt": "REIT PREV XP?"},
        {"id": "m7",  "txt": "Zerou REIT PREV XP?"},
        {"id": "m8",  "txt": "INFLAÇÃO XP?"},
        {"id": "m9",  "txt": "Boletou INFLAÇÃO XP?"},
        {"id": "m10", "txt": "PREVIDENCE XP — mandou para o Zuniga?"},
        {"id": "m11", "txt": "Avisar o Zuniga (movimentos Prev Itaú)"},
        {"id": "m12", "txt": "Outros cotistas na Mellon?"},
        {"id": "m13", "txt": "BTG CAP"},
        {"id": "m14", "txt": "BTG PREV"},
        {"id": "m15", "txt": "BTG SA"},
        {"id": "m16", "txt": "BTG ALTERNATIVES"},
        {"id": "m17", "txt": "BRADESCO?"},
        {"id": "m18", "txt": "PREVs ITAÚ (passivo — Itaú Vida e Prev)", "obs": "Checar se tem R$ na conta"},
        {"id": "m19", "txt": "Zerou ITAÚ FIES × FIFES (Ativo e Passivo)", "obs": "Se aplicação: boletar 1º no ativo e só depois no passivo"},
        {"id": "m20", "txt": "Outros fundos no Itaú?"},
        {"id": "m21", "txt": "Respondeu todas as movimentações liquidadas?"},
        {"id": "m22", "txt": "MELLON — zerou?"},
    ]},
    {"secao": "PRÉVIA", "icone": "📝", "itens": [
        {"id": "p1",  "txt": "MELLON txt", "obs": "Trocou qtd de cotas por financeiro? · Atenção com movimentações entre fundos"},
        {"id": "p2",  "txt": "Zeragens dos FICs nos Masters — MELLON"},
        {"id": "p3",  "txt": "Zeragens dos FICs nos Masters — BRADESCO"},
        {"id": "p4",  "txt": "Zeragens dos FICs nos Masters — ITAÚ (Previs)"},
        {"id": "p5",  "txt": "Zeragens dos FIES nos FIFES — ITAÚ"},
        {"id": "p6",  "txt": "ITAÚ Previs — liquidou"},
        {"id": "p7",  "txt": "Zeragens dos FIES nos FIFES — ITAÚ (liquidou)"},
        {"id": "p8",  "txt": "ITAÚ — Sabesprev"},
        {"id": "p9",  "txt": "ITAÚ — FAPES"},
        {"id": "p10", "txt": "XP Renda 90 (EG)"},
        {"id": "p11", "txt": "Yield 120 (EG)"},
        {"id": "p12", "txt": "Infra Renda Adv (EG)"},
        {"id": "p13", "txt": "Infra Geral Advisory (Feeder)"},
        {"id": "p14", "txt": "BTG Login CAP"},
        {"id": "p15", "txt": "BTG Login PREV"},
        {"id": "p16", "txt": "BTG Login SA"},
        {"id": "p17", "txt": "BTG Login ALTERNATIVES"},
        {"id": "p18", "txt": "BTG/ITAÚ — FCOPEL", "obs": "Atenção"},
    ]},
    {"secao": "NET", "icone": "📊", "itens": [
        {"id": "n1",  "txt": "MELLON txt", "obs": "Trocou qtd de cotas por financeiro? · Atenção com movimentações entre fundos"},
        {"id": "n2",  "txt": "Manual — ITAÚ (todos os fundos)"},
        {"id": "n3",  "txt": "Manual — BTG Login CAP"},
        {"id": "n4",  "txt": "Manual — BTG Login PREV"},
        {"id": "n5",  "txt": "Manual — BTG Login SA"},
        {"id": "n6",  "txt": "Manual — BTG Login ALTERNATIVES"},
        {"id": "n7",  "txt": "Manual — BRADESCO"},
        {"id": "n8",  "txt": "Manual — XP Renda 90 (EG)"},
        {"id": "n9",  "txt": "Manual — Yield 120 (EG)"},
        {"id": "n10", "txt": "Manual — Infra Renda Adv (EG)"},
    ]},
    {"secao": "CONFERÊNCIA FINAL", "icone": "✅", "itens": [
        {"id": "f1", "txt": "Checou TODAS as datas de liquidação? Alguma divergência?", "obs": "ATENÇÃO TOTAL"},
        {"id": "f2", "txt": "Verificar se as movimentações fazem sentido"},
        {"id": "f3", "txt": "Tudo foi liquidado? Se sim, pode enviar por e-mail"},
        {"id": "f4", "txt": "Não deixar campo de data da coluna A em branco"},
    ]},
]

TODOS_IDS = [it["id"] for sec in CHECKLIST for it in sec["itens"]]
