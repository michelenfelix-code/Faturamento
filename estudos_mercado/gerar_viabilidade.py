"""
Gera a planilha de viabilidade das 10 oportunidades do levantamento de mercado
(ranking de out/2026), com foco em São Paulo e Alto Tietê.

Como usar:
  python gerar_viabilidade.py
  -> cria "viabilidade_oportunidades.xlsx" na mesma pasta deste script.

Para testar outras premissas, altere OPORTUNIDADES abaixo e rode de novo,
ou edite direto as células amarelas da aba "Premissas" no Excel (tudo recalcula sozinho).
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SAIDA = Path(__file__).with_name("viabilidade_oportunidades.xlsx")
MESES = 36

# ---------------------------------------------------------------------------
# Premissas (todas são ESTIMATIVAS para validar com cotações locais)
# ---------------------------------------------------------------------------
# Ordem dos valores em cada oportunidade:
#   investimento, ticket mensal, receita de implantação, clientes no mês 1, novos/mês,
#   capacidade, churn, custo variável %, impostos %, custos fixos, pró-labore
OPORTUNIDADES = [
    {"nome": "1. BPO financeiro / controladoria", "aba": "01 BPO",
     "cliente": "1 PME com contrato mensal de BPO + relatórios gerenciais",
     "valores": [25000, 1500, 1000, 2, 2, 40, 0.02, 0.30, 0.06, 3000, 5000]},
    {"nome": "2. Consultoria NR-1 psicossocial", "aba": "02 NR-1",
     "cliente": "1 empresa com acompanhamento mensal; implantação = diagnóstico + revisão do PGR",
     "valores": [40000, 1200, 8000, 1, 1.5, 30, 0.03, 0.35, 0.06, 4000, 6000]},
    {"nome": "3. Agência de cuidadores", "aba": "03 Cuidadores",
     "cliente": "1 família com 1 cuidador mensal (ex.: 12x36 diurno); custo variável = cuidador",
     "valores": [50000, 5500, 0, 2, 1.5, 40, 0.05, 0.78, 0.11, 6000, 5000]},
    {"nome": "4. Oficina de eletrificados", "aba": "04 Eletrificados",
     "cliente": "1 veículo na carteira; ticket = gasto médio mensal (revisões + reparos)",
     "valores": [250000, 350, 0, 20, 15, 300, 0.02, 0.45, 0.08, 25000, 6000]},
    {"nome": "5. Climatização / PMOC", "aba": "05 Climatizacao",
     "cliente": "1 contrato PMOC mensal; implantação = instalação/adequação inicial",
     "valores": [80000, 600, 1500, 3, 3, 80, 0.02, 0.30, 0.06, 8000, 5000]},
    {"nome": "6. Carregadores VE", "aba": "06 Carregadores",
     "cliente": "1 vaga/carregador instalado; implantação = projeto + instalação; ticket = manutenção",
     "valores": [50000, 80, 6000, 2, 4, 300, 0.01, 0.55, 0.06, 4000, 5000]},
    {"nome": "7. Manutenção predial", "aba": "07 Predial",
     "cliente": "1 condomínio/empresa com contrato preventivo mensal",
     "valores": [100000, 2500, 0, 2, 1.5, 35, 0.03, 0.45, 0.06, 8000, 5000]},
    {"nome": "8. O&M solar", "aba": "08 Solar",
     "cliente": "1 sistema fotovoltaico com plano (limpeza + inspeção + monitoramento)",
     "valores": [70000, 150, 400, 10, 15, 400, 0.02, 0.30, 0.06, 5000, 5000]},
    {"nome": "9. Síndico profissional", "aba": "09 Sindico",
     "cliente": "1 condomínio administrado",
     "valores": [20000, 3500, 0, 1, 0.5, 12, 0.015, 0.10, 0.06, 1500, 5000]},
    {"nome": "10. Clínica TEA", "aba": "10 TEA",
     "cliente": "1 paciente em terapia mensal; implantação = avaliação inicial",
     "valores": [600000, 4000, 1500, 5, 4, 80, 0.02, 0.50, 0.11, 35000, 8000]},
]
NOMES = [o["nome"] for o in OPORTUNIDADES]
ABAS_FLUXO = [f"Fluxo {o['aba']}" for o in OPORTUNIDADES]
COLS = [get_column_letter(2 + i) for i in range(len(OPORTUNIDADES))]   # B..K
COL_UNID = get_column_letter(2 + len(OPORTUNIDADES))                  # L
COL_OBS = get_column_letter(3 + len(OPORTUNIDADES))                   # M

# (rótulo, unidade, formato, observação)
PREMISSAS = [
    ("Investimento inicial", "R$", "moeda",
     "Estimativa: equipamentos, veículo, estrutura, software, marketing, CNPJ."),
    ("Ticket mensal recorrente por cliente", "R$/mês", "moeda",
     "Estimativa a validar com cotação oculta (aba Validação de campo). Ver linha 'O que é 1 cliente'."),
    ("Receita de implantação por novo cliente", "R$", "moeda",
     "Cobrança única na entrada (diagnóstico, onboarding, instalação, avaliação)."),
    ("Clientes no 1º mês", "clientes", "num", "Estimativa conservadora (rede de contatos)."),
    ("Novos clientes por mês", "clientes/mês", "num", "Estimativa. Ajustada pelo fator de cenário."),
    ("Capacidade máxima de clientes", "clientes", "num",
     "Limite da estrutura prevista nos custos fixos. Acima disso, rever custos e investimento."),
    ("Cancelamento mensal (churn)", "% ao mês", "pct", "Estimativa de perda de clientes por mês."),
    ("Custo variável", "% da receita", "pct",
     "Mão de obra direta, peças/materiais, parceiros. Cuidadores: salário + encargos do cuidador."),
    ("Impostos sobre a receita", "% da receita", "pct",
     "Referência: Simples Nacional. Confirme CNAE, anexo e Fator R. Cessão de mão de obra "
     "(cuidadores) e clínicas exigem análise de enquadramento."),
    ("Custos fixos mensais", "R$/mês", "moeda",
     "Aluguel, veículo, sistemas, contador, equipe fixa, marketing."),
    ("Pró-labore mensal", "R$/mês", "moeda", "Retirada do(a) sócio(a). Ajuste para sua necessidade."),
]
LINHA_CLIENTE = 5
LINHA = {rotulo: 6 + i for i, (rotulo, *_) in enumerate(PREMISSAS)}   # linhas 6..16

# Bloco de cenários (linhas abaixo das premissas)
L_CEN_TIT = 18
L_CEN_ESC = 19            # célula do cenário escolhido: B19
L_CEN_CAB = 21
L_CEN_INI = 22            # Pessimista, Base, Otimista: 22..24
L_FATOR_NOVOS = 26
L_FATOR_TICKET = 27
L_TMA = 29
CENARIOS = [("Pessimista", 0.7, 0.9), ("Base", 1.0, 1.0), ("Otimista", 1.3, 1.1)]
TMA_MENSAL = 0.012

# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------
FONTE = "Arial"
AZUL = Font(name=FONTE, color="0000FF")
PRETO = Font(name=FONTE, color="000000")
VERDE = Font(name=FONTE, color="008000")
NEGRITO = Font(name=FONTE, bold=True)
TITULO = Font(name=FONTE, bold=True, size=14)
CAB = Font(name=FONTE, bold=True, color="FFFFFF")
FUNDO_CAB = PatternFill("solid", fgColor="1F4E78")
AMARELO = PatternFill("solid", fgColor="FFFF00")
CINZA = PatternFill("solid", fgColor="F2F2F2")
BORDA = Border(bottom=Side(style="thin", color="BFBFBF"))

FMT = {
    "moeda": 'R$ #,##0;(R$ #,##0);"-"',
    "num": '#,##0.0;(#,##0.0);"-"',
    "pct": '0.0%;(0.0%);"-"',
    "int": '0',
}


def cabecalho(ws, linha, textos):
    for c, t in enumerate(textos, start=1):
        cel = ws.cell(row=linha, column=c, value=t)
        cel.font = CAB
        cel.fill = FUNDO_CAB
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def larguras(ws, valores):
    for i, w in enumerate(valores, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------------------
# Aba Leia-me
# ---------------------------------------------------------------------------
FONTES = [
    ("BPO / mercado", "https://www.imarcgroup.com/report/pt-br/brazil-business-process-outsourcing-market"),
    ("Reforma tributária e contratações", "https://fenacon.org.br/reforma-tributaria/"
     "reforma-tributaria-impulsiona-contratacao-de-contadores-e-tributaristas/"),
    ("NR-1 em vigor desde 26/05/2026", "https://www.rsdata.com.br/?p=36737"),
    ("Fiscalização NR-1", "https://www.contabeis.com.br/noticias/76472/"
     "nr-1-entra-em-fase-de-fiscalizacao-e-pressiona-empresas/"),
    ("Afastamentos por transtornos mentais 2025", "https://outraspalavras.net/outrasaude/"
     "quais-as-doencas-que-mais-afastaram-trabalhadores-em-2025/"),
    ("Cuidadores +74% (Sebrae)", "https://fenacon.org.br/noticias/"
     "em-cinco-anos-abertura-de-pequenos-negocios-de-servico-de-cuidadores-cresceu-74-no-brasil/"),
    ("Envelhecimento (Censo 2022)", "https://www.em.com.br/gerais/2026/07/"
     "7467827-com-o-pais-envelhecendo-quem-cuidara-dos-nossos-idosos-e-de-nos-mesmos.html"),
    ("Falta de técnicos em eletrificados (IQA)", "https://agenciainfra.com/blog/?p=37576"),
    ("Mercado de ar-condicionado 2025 (Abrava)", "https://iifiir.org/en/news/"
     "brazil-s-ac-market-continues-to-expand-in-2025-with-slower-but-sustained-growth"),
    ("Carregadores em condomínios SP", "https://portas.com.br/noticias/carregamento-carros-eletricos-condominio-sp/"),
    ("Falta de mão de obra na construção (FGV IBRE)", "https://en.clickpetroleoegas.com.br/"
     "industria-da-construcao-vive-apagao-de-mao-de-obra-falta-de-pedreiros-mestres-de-obras-e-eletricistas-vml97/"),
    ("4,3 milhões de sistemas solares", "https://www.pv-magazine-brasil.com/author/alessandra-neris/page/2"),
    ("Síndico profissional", "https://diariodocomercio.com.br/opiniao/artigo/"
     "sindico-profissional-carreira-alta-mercado-condominial/"),
    ("TEA: filas e falta de profissionais", "https://futurodasaude.com.br/?p=44484"),
]


def aba_leiame(wb):
    ws = wb.active
    ws.title = "Leia-me"
    linhas = [
        ("Viabilidade das 10 oportunidades do levantamento de mercado", TITULO),
        ("Brasil, São Paulo e Alto Tietê - ranking de out/2026", NEGRITO),
        ("", None),
        ("COMO USAR", NEGRITO),
        ("1. Edite só as células AMARELAS com texto AZUL na aba 'Premissas'. Todo o resto é fórmula.", PRETO),
        (f"2. Escolha o cenário (Pessimista / Base / Otimista) na célula B{L_CEN_ESC} da aba 'Premissas'.", PRETO),
        ("3. Compare as 10 oportunidades na aba 'Resumo' (payback, VPL, TIR, ponto de equilíbrio e ranking por VPL).", PRETO),
        ("4. As abas 'Fluxo ...' mostram mês a mês (36 meses) clientes, receita, custos e caixa acumulado.", PRETO),
        ("5. Registre concorrentes e cotações do Alto Tietê na aba 'Validação de campo' e ajuste as premissas.", PRETO),
        ("", None),
        ("LEGENDA DE CORES", NEGRITO),
        ("Texto azul + fundo amarelo = premissa digitada (pode alterar)", AZUL),
        ("Texto preto = fórmula calculada (não alterar)", PRETO),
        ("Texto verde = vínculo com outra aba", VERDE),
        ("", None),
        ("ATENÇÃO", NEGRITO),
        ("Todas as premissas financeiras são ESTIMATIVAS iniciais, não dados de mercado. "
         "Valide preços, custos e impostos antes de qualquer decisão.", PRETO),
        ("Cada negócio tem um significado próprio de 'cliente' (linha 5 da aba Premissas): família, "
         "empresa, veículo, sistema solar, condomínio, paciente.", PRETO),
        ("Os impostos usam alíquotas de referência do Simples Nacional. Confirme CNAE, anexo e Fator R.", PRETO),
        ("O modelo trata clientes como média (pode haver frações) e não inclui reajuste de preços, "
         "sazonalidade nem capital de giro: é uma visão gerencial simplificada.", PRETO),
        ("", None),
        ("FONTES DOS DADOS DE DEMANDA (levantamento de out/2026)", NEGRITO),
    ] + [(f"{t}: {u}", PRETO) for t, u in FONTES]
    for i, (texto, fonte) in enumerate(linhas, start=1):
        cel = ws.cell(row=i, column=1, value=texto)
        if fonte:
            cel.font = fonte
        if fonte is AZUL:
            cel.fill = AMARELO
    ws.column_dimensions["A"].width = 130


# ---------------------------------------------------------------------------
# Aba Premissas
# ---------------------------------------------------------------------------
def entrada(cel, valor, fmt):
    cel.value = valor
    cel.font = AZUL
    cel.fill = AMARELO
    cel.number_format = fmt


def aba_premissas(wb):
    ws = wb.create_sheet("Premissas")
    ws["A1"] = "Premissas do modelo"
    ws["A1"].font = TITULO
    ws["A2"] = "Células amarelas = estimativas editáveis. Valores iniciais são hipóteses a validar."
    ws["A2"].font = Font(name=FONTE, italic=True)
    cabecalho(ws, 4, ["Premissa"] + NOMES + ["Unidade", "Observação"])

    ws.cell(row=LINHA_CLIENTE, column=1, value="O que é 1 cliente").font = NEGRITO
    for j, o in enumerate(OPORTUNIDADES, start=2):
        cel = ws.cell(row=LINHA_CLIENTE, column=j, value=o["cliente"])
        cel.font = Font(name=FONTE, italic=True, size=9)
        cel.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[LINHA_CLIENTE].height = 75

    for k, (rotulo, unidade, fmt, obs) in enumerate(PREMISSAS):
        r = LINHA[rotulo]
        ws.cell(row=r, column=1, value=rotulo).font = PRETO
        for j, o in enumerate(OPORTUNIDADES, start=2):
            entrada(ws.cell(row=r, column=j), o["valores"][k], FMT[fmt])
        ws[f"{COL_UNID}{r}"] = unidade
        ws[f"{COL_OBS}{r}"] = obs
        ws[f"{COL_OBS}{r}"].alignment = Alignment(wrap_text=True, vertical="top")

    ws[f"A{L_CEN_TIT}"] = "Cenário"
    ws[f"A{L_CEN_TIT}"].font = NEGRITO
    ws[f"A{L_CEN_ESC}"] = "Cenário escolhido"
    entrada(ws[f"B{L_CEN_ESC}"], "Base", "General")
    dv = DataValidation(type="list", formula1='"Pessimista,Base,Otimista"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"B{L_CEN_ESC}")
    ws[f"C{L_CEN_ESC}"] = "Lista suspensa: Pessimista, Base ou Otimista (vale para todas as oportunidades)."

    cabecalho(ws, L_CEN_CAB, ["Cenário", "Fator novos clientes", "Fator ticket"])
    for i, (nome, f_novos, f_ticket) in enumerate(CENARIOS, start=L_CEN_INI):
        ws.cell(row=i, column=1, value=nome)
        entrada(ws.cell(row=i, column=2), f_novos, "0.00x")
        entrada(ws.cell(row=i, column=3), f_ticket, "0.00x")
    ws[f"D{L_CEN_INI}"] = "Fatores multiplicam 'Novos clientes por mês' e os tickets (recorrente e implantação)."

    faixa_nome = f"$A${L_CEN_INI}:$A${L_CEN_INI + 2}"
    ws[f"A{L_FATOR_NOVOS}"] = "Fator novos clientes (ativo)"
    ws[f"B{L_FATOR_NOVOS}"] = (f"=INDEX($B${L_CEN_INI}:$B${L_CEN_INI + 2},"
                               f"MATCH($B${L_CEN_ESC},{faixa_nome},0))")
    ws[f"A{L_FATOR_TICKET}"] = "Fator ticket (ativo)"
    ws[f"B{L_FATOR_TICKET}"] = (f"=INDEX($C${L_CEN_INI}:$C${L_CEN_INI + 2},"
                                f"MATCH($B${L_CEN_ESC},{faixa_nome},0))")
    for r in (L_FATOR_NOVOS, L_FATOR_TICKET):
        ws[f"B{r}"].number_format = "0.00x"

    ws[f"A{L_TMA}"] = "Taxa mínima de atratividade (TMA)"
    entrada(ws[f"B{L_TMA}"], TMA_MENSAL, "0.00%")
    ws[f"C{L_TMA}"] = ("% ao mês. Custo de oportunidade do dinheiro (ex.: aplicação conservadora). "
                       "Usada no VPL. Ajuste para sua realidade.")

    ultima_col = 3 + len(OPORTUNIDADES)
    for row in ws.iter_rows(min_row=1, max_row=L_TMA, max_col=ultima_col):
        for cel in row:
            if cel.font.name != FONTE:
                cel.font = Font(name=FONTE, bold=cel.font.bold, italic=cel.font.italic,
                                color=cel.font.color, size=cel.font.size)
            if 4 <= cel.row <= LINHA[PREMISSAS[-1][0]]:
                cel.border = BORDA
    larguras(ws, [36] + [17] * len(OPORTUNIDADES) + [13, 70])
    ws.row_dimensions[4].height = 45
    ws.freeze_panes = "B6"


# ---------------------------------------------------------------------------
# Abas de fluxo mensal
# ---------------------------------------------------------------------------
COLUNAS_FLUXO = [
    "Mês", "Clientes ativos", "Novos clientes (efetivos)", "Receita recorrente",
    "Receita de implantação", "Receita total", "Impostos", "Custos variáveis",
    "Margem de contribuição", "Custos fixos + pró-labore", "Resultado operacional",
    "Fluxo de caixa", "Caixa acumulado", "Payback atingido? (1=sim)",
]
L0 = 5                # linha do mês 0
LFIM = L0 + MESES     # linha do mês 36
F_NOVOS = f"Premissas!$B${L_FATOR_NOVOS}"
F_TICKET = f"Premissas!$B${L_FATOR_TICKET}"
TMA = f"Premissas!$B${L_TMA}"


def p(col, rotulo):
    """Referência absoluta a uma premissa na aba Premissas."""
    return f"Premissas!${col}${LINHA[rotulo]}"


def aba_fluxo(wb, nome_aba, nome_op, col):
    ws = wb.create_sheet(nome_aba)
    ws["A1"] = f"Fluxo mensal - {nome_op}"
    ws["A1"].font = TITULO
    ws["A2"] = "Cenário ativo:"
    ws["B2"] = f"=Premissas!$B${L_CEN_ESC}"
    ws["B2"].font = VERDE
    cabecalho(ws, 4, COLUNAS_FLUXO)

    cap = p(col, "Capacidade máxima de clientes")
    churn = p(col, "Cancelamento mensal (churn)")
    inicial = p(col, "Clientes no 1º mês")
    novos = p(col, "Novos clientes por mês")
    ticket = p(col, "Ticket mensal recorrente por cliente")
    impl = p(col, "Receita de implantação por novo cliente")
    imp = p(col, "Impostos sobre a receita")
    cv = p(col, "Custo variável")
    fixos = p(col, "Custos fixos mensais")
    pro = p(col, "Pró-labore mensal")
    inv = p(col, "Investimento inicial")

    # Mês 0: só o investimento
    ws.cell(row=L0, column=1, value=0)
    for c in range(2, 12):
        ws.cell(row=L0, column=c, value=0)
    ws.cell(row=L0, column=12, value=f"=-{inv}")
    ws.cell(row=L0, column=13, value=f"=L{L0}")
    ws.cell(row=L0, column=14, value=f"=IF(M{L0}>=0,1,0)")

    for r in range(L0 + 1, LFIM + 1):
        a = r - 1  # linha anterior
        f = {
            1: f"=A{a}+1",
            2: f"=MIN({cap},B{a}*(1-{churn})+IF(A{r}=1,{inicial},{novos}*{F_NOVOS}))",
            3: f"=MAX(0,B{r}-B{a}*(1-{churn}))",
            4: f"=B{r}*{ticket}*{F_TICKET}",
            5: f"=C{r}*{impl}*{F_TICKET}",
            6: f"=D{r}+E{r}",
            7: f"=F{r}*{imp}",
            8: f"=F{r}*{cv}",
            9: f"=F{r}-G{r}-H{r}",
            10: f"={fixos}+{pro}",
            11: f"=I{r}-J{r}",
            12: f"=K{r}",
            13: f"=M{a}+L{r}",
            14: f"=IF(M{r}>=0,1,0)",
        }
        for c, formula in f.items():
            ws.cell(row=r, column=c, value=formula)

    for r in range(L0, LFIM + 1):
        for c in range(1, 15):
            cel = ws.cell(row=r, column=c)
            if c in (2, 3):
                cel.number_format = FMT["num"]
            elif c in (1, 14):
                cel.number_format = FMT["int"]
            else:
                cel.number_format = FMT["moeda"]
            cel.font = PRETO
            if (r - L0) % 12 == 0 and r != L0:
                cel.fill = CINZA
    ws.cell(row=L0, column=12).font = VERDE

    ws.cell(row=LFIM + 2, column=1,
            value="Linhas cinza = fim de cada ano (meses 12, 24 e 36).").font = Font(name=FONTE, italic=True)
    larguras(ws, [7, 12, 14, 16, 16, 16, 14, 16, 16, 16, 16, 16, 16, 12])
    ws.row_dimensions[4].height = 45
    ws.freeze_panes = "B5"


# ---------------------------------------------------------------------------
# Aba Resumo
# ---------------------------------------------------------------------------
def aba_resumo(wb):
    ws = wb.create_sheet("Resumo", 1)
    ws["A1"] = "Resumo comparativo - 36 meses"
    ws["A1"].font = TITULO
    ws["A2"] = "Cenário ativo:"
    ws["B2"] = f"=Premissas!$B${L_CEN_ESC}"
    ws["B2"].font = VERDE
    cabecalho(ws, 4, ["Indicador"] + NOMES + ["Como ler"])

    l12, l36 = L0 + 12, LFIM
    indicadores = [
        ("Investimento inicial", "moeda",
         lambda a, c: f"={p(c, 'Investimento inicial')}", "Valor aplicado no mês 0."),
        ("Clientes ativos - mês 12", "num", lambda a, c: f"='{a}'!B{l12}", "Ver 'O que é 1 cliente' em Premissas."),
        ("Clientes ativos - mês 36", "num", lambda a, c: f"='{a}'!B{l36}", ""),
        ("Receita mensal - mês 12", "moeda", lambda a, c: f"='{a}'!F{l12}", ""),
        ("Receita mensal - mês 36", "moeda", lambda a, c: f"='{a}'!F{l36}", ""),
        ("Resultado mensal - mês 12", "moeda", lambda a, c: f"='{a}'!K{l12}",
         "Depois de impostos, custos e pró-labore."),
        ("Resultado mensal - mês 36", "moeda", lambda a, c: f"='{a}'!K{l36}", ""),
        ("Margem líquida - mês 12", "pct",
         lambda a, c: f"=IFERROR('{a}'!K{l12}/'{a}'!F{l12},0)", "Resultado / receita."),
        ("Margem líquida - mês 36", "pct",
         lambda a, c: f"=IFERROR('{a}'!K{l36}/'{a}'!F{l36},0)", ""),
        ("Receita acumulada - 36 meses", "moeda",
         lambda a, c: f"=SUM('{a}'!F{L0}:F{LFIM})", ""),
        ("Caixa acumulado - mês 36", "moeda", lambda a, c: f"='{a}'!M{l36}",
         "Já descontado o investimento inicial."),
        ("Ponto de equilíbrio (clientes)", "num",
         lambda a, c: (f"=IFERROR(({p(c, 'Custos fixos mensais')}+{p(c, 'Pró-labore mensal')})/"
                       f"({p(c, 'Ticket mensal recorrente por cliente')}*{F_TICKET}*"
                       f"(1-{p(c, 'Impostos sobre a receita')}-{p(c, 'Custo variável')})),\"n/d\")"),
         "Clientes recorrentes para cobrir custos fixos e pró-labore (sem implantação)."),
        ("Capacidade cobre o equilíbrio?", "txt",
         lambda a, c: (f"=IF(ISNUMBER(B{{eq}}),IF({p(c, 'Capacidade máxima de clientes')}>=B{{eq}},"
                       f"\"Sim\",\"Não - rever estrutura\"),\"n/d\")"),
         "Se 'Não', o negócio não se paga nem lotado com a estrutura prevista."),
        ("Payback (meses)", "int",
         lambda a, c: f"=IFERROR(MATCH(1,'{a}'!N{L0}:N{LFIM},0)-1,\"Acima de 36\")",
         "Mês em que o caixa acumulado volta a ficar positivo."),
        ("VPL - 36 meses", "moeda",
         lambda a, c: f"=NPV({TMA},'{a}'!L{L0 + 1}:L{LFIM})+'{a}'!L{L0}",
         "Positivo = rende mais que a TMA."),
        ("TIR mensal", "pct", lambda a, c: f"=IFERROR(IRR('{a}'!L{L0}:L{LFIM}),\"n/d\")",
         "Comparar com a TMA."),
        ("TIR anual equivalente", "pct",
         lambda a, c: f"=IFERROR((1+IRR('{a}'!L{L0}:L{LFIM}))^12-1,\"n/d\")", ""),
        ("Ranking por VPL", "int", None, "1 = maior VPL no cenário ativo."),
    ]
    linha_de = {rotulo: 5 + i for i, (rotulo, *_) in enumerate(indicadores)}
    l_eq = linha_de["Ponto de equilíbrio (clientes)"]
    l_vpl = linha_de["VPL - 36 meses"]
    ultima = get_column_letter(1 + len(OPORTUNIDADES))
    col_como = 2 + len(OPORTUNIDADES)

    for rotulo, fmt, formula, como_ler in indicadores:
        i = linha_de[rotulo]
        ws.cell(row=i, column=1, value=rotulo).font = PRETO
        for j, (aba, colp) in enumerate(zip(ABAS_FLUXO, COLS), start=2):
            letra = get_column_letter(j)
            if formula is None:
                valor = f"=RANK({letra}{l_vpl},$B${l_vpl}:${ultima}${l_vpl},0)"
            else:
                valor = formula(aba, colp).replace("B{eq}", f"{letra}{l_eq}")
            cel = ws.cell(row=i, column=j, value=valor)
            cel.font = PRETO if formula is None or rotulo.startswith("Capacidade") else VERDE
            if fmt != "txt":
                cel.number_format = FMT[fmt]
            cel.alignment = Alignment(horizontal="right")
        ws.cell(row=i, column=col_como, value=como_ler).font = PRETO
        for c in range(1, col_como + 1):
            ws.cell(row=i, column=c).border = BORDA
    for c in range(1, col_como):
        ws.cell(row=linha_de["Ranking por VPL"], column=c).font = NEGRITO

    nota = 5 + len(indicadores) + 1
    ws.cell(row=nota, column=1,
            value=f"Todos os valores dependem das premissas estimadas; troque o cenário em "
                  f"Premissas!B{L_CEN_ESC} para testar sensibilidade.").font = Font(name=FONTE, italic=True)
    larguras(ws, [32] + [17] * len(OPORTUNIDADES) + [60])
    ws.row_dimensions[4].height = 45
    ws.freeze_panes = "B5"


# ---------------------------------------------------------------------------
# Aba Validação de campo
# ---------------------------------------------------------------------------
def aba_validacao(wb):
    ws = wb.create_sheet("Validação de campo")
    ws["A1"] = "Validação de campo - concorrentes e cotações no Alto Tietê"
    ws["A1"].font = TITULO
    ws["A2"] = ("Preencha uma linha por concorrente encontrado (Google Maps, indicações, cotação oculta). "
                "A linha 5 é um EXEMPLO fictício de formato: substitua pelos seus dados.")
    ws["A2"].font = Font(name=FONTE, italic=True)
    cols = ["Oportunidade", "Cidade", "Empresa", "Nota Google", "Nº avaliações",
            "Preço cotado (R$)", "Prazo para atender (dias)", "Pontos fracos / reclamações", "Data"]
    cabecalho(ws, 4, cols)
    exemplo = [NOMES[0], "Mogi das Cruzes", "Exemplo Ltda (fictício)", 4.3, 27, 1800, 10,
               "Só faz lançamentos, sem relatório gerencial", "2026-10-20"]
    for c, v in enumerate(exemplo, start=1):
        cel = ws.cell(row=5, column=c, value=v)
        cel.font = Font(name=FONTE, color="0000FF", italic=True)
        cel.fill = AMARELO
    ws["F5"].number_format = FMT["moeda"]

    # Lista suspensa a partir da aba Premissas (lista literal passaria do limite de 255 caracteres)
    ultima = get_column_letter(1 + len(OPORTUNIDADES))
    ws["Q1"] = "Lista de oportunidades (apoio à lista suspensa)"
    for k, nome in enumerate(NOMES, start=2):
        ws.cell(row=k, column=17, value=nome)
    dv = DataValidation(type="list", formula1=f"$Q$2:$Q${1 + len(NOMES)}", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("A5:A300")

    ws.cell(row=3, column=11, value="Resumo por oportunidade").font = NEGRITO
    cab_res = ["Oportunidade", "Concorrentes", "Nota média", "Preço médio (R$)", "Prazo médio (dias)"]
    for c, t in enumerate(cab_res, start=11):
        cel = ws.cell(row=4, column=c, value=t)
        cel.font = CAB
        cel.fill = FUNDO_CAB
        cel.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, nome in enumerate(NOMES, start=5):
        ws.cell(row=i, column=11, value=nome)
        ws.cell(row=i, column=12, value=f"=COUNTIF($A$5:$A$300,K{i})")
        ws.cell(row=i, column=13, value=f'=IFERROR(AVERAGEIF($A$5:$A$300,K{i},$D$5:$D$300),"-")')
        ws.cell(row=i, column=14, value=f'=IFERROR(AVERAGEIF($A$5:$A$300,K{i},$F$5:$F$300),"-")')
        ws.cell(row=i, column=15, value=f'=IFERROR(AVERAGEIF($A$5:$A$300,K{i},$G$5:$G$300),"-")')
        ws.cell(row=i, column=13).number_format = "0.0"
        ws.cell(row=i, column=14).number_format = FMT["moeda"]
        ws.cell(row=i, column=15).number_format = "0.0"
    ws.cell(row=6 + len(NOMES), column=11,
            value="Use o preço médio para revisar o ticket na aba Premissas.").font = Font(name=FONTE, italic=True)
    ws["F4"].comment = Comment("Valor mensal cobrado pelo concorrente para o serviço equivalente.", "Modelo")

    for row in ws.iter_rows(min_row=1, max_row=6 + len(NOMES), max_col=17):
        for cel in row:
            if cel.font.name != FONTE:
                cel.font = Font(name=FONTE, bold=cel.font.bold, italic=cel.font.italic, color=cel.font.color)
    larguras(ws, [30, 18, 26, 11, 12, 16, 14, 42, 12, 3, 30, 13, 11, 16, 14, 3, 30])
    ws.row_dimensions[4].height = 32
    ws.freeze_panes = "A5"


def main():
    wb = Workbook()
    aba_leiame(wb)
    aba_premissas(wb)
    for aba, o, col in zip(ABAS_FLUXO, OPORTUNIDADES, COLS):
        aba_fluxo(wb, aba, o["nome"], col)
    aba_resumo(wb)
    aba_validacao(wb)
    wb.save(SAIDA)
    print(f"Planilha gerada: {SAIDA}")


if __name__ == "__main__":
    main()
