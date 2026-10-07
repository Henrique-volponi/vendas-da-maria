# -*- coding: utf-8 -*-
"""
maria.py — o "cérebro" da solução.

Entende o que a Maria escreve (uma frase solta ou uma linha da planilha velha)
e transforma em uma venda organizada. O MESMO interpretador é usado para:
  1) migrar a planilha antiga dela (migrar.py)
  2) registrar vendas novas por conversa (app.py)

Regra de ouro: quando não dá pra ter certeza, NÃO chuta. Marca "Confirmar?".
"""
import re
import unicodedata
from datetime import date, timedelta

# ----------------------------------------------------------------------------
# Colunas da planilha organizada (a ordem importa: é a ordem das colunas)
# ----------------------------------------------------------------------------
COLUNAS = [
    "Data", "Produto", "Cliente", "Valor (R$)", "Frete (R$)", "Pagamento",
    "Cartão", "Parcelas", "Entrega", "Rastreio", "Recebido?", "Obs",
    "Confirmar?", "Nota do sistema", "Nº original",
]

PAGAMENTOS_VALIDOS = ["Pix", "Dinheiro", "Débito", "Crédito", "Fiado", "Cartão", "Não informado"]
CARTOES_VALIDOS = ["Visa", "Master", "Elo", "Hipercard", "Amex", ""]
ENTREGAS_VALIDAS = ["Retirou na loja", "Correios", "Motoboy", "Não informado"]

# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
def norm(s):
    """minúsculo, sem acento, espaços colapsados"""
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def cap(s):
    s = (s or "").strip()
    return s[:1].upper() + s[1:] if s else s


PARTICULAS = {"da", "do", "de", "das", "dos", "e"}


def titulo(s):
    """'tia da ana' -> 'Tia da Ana'"""
    palavras = (s or "").strip().split()
    out = []
    for i, p in enumerate(palavras):
        out.append(p.lower() if (p.lower() in PARTICULAS and i > 0) else p[:1].upper() + p[1:].lower())
    return " ".join(out)


def vazio(v):
    return v is None or (isinstance(v, float) and v != v) or str(v).strip() == ""


RE_RASTREIO = re.compile(r"\b([A-Za-z]{2}\d{9}[A-Za-z]{2})\b")
RE_TELEFONE = re.compile(r"(?<!\d)(\(?\d{2}\)?\s?9?\d{4}-?\d{4})(?!\d)")

MESES = {
    "janeiro": 1, "jan": 1, "fevereiro": 2, "fev": 2, "marco": 3, "mar": 3,
    "abril": 4, "abr": 4, "maio": 5, "mai": 5, "junho": 6, "jun": 6,
    "julho": 7, "jul": 7, "agosto": 8, "ago": 8, "setembro": 9, "set": 9,
    "outubro": 10, "out": 10, "novembro": 11, "nov": 11, "dezembro": 12, "dez": 12,
}
DIAS_SEMANA = {"segunda": 0, "terca": 1, "quarta": 2, "quinta": 3, "sexta": 4, "sabado": 5, "domingo": 6}

# ----------------------------------------------------------------------------
# DATA
# ----------------------------------------------------------------------------
def _data_valida(a, m, d):
    try:
        return date(a, m, d)
    except ValueError:
        return None


def parse_data(texto, hoje, ano_padrao=None, relativo_ok=True):
    """
    Devolve (data|None, nota|None, confirmar:bool, trecho_consumido|None).
    - relativo_ok=False (migração): "ontem", "hj", "segunda" não têm referência -> None + confirmar.
    - Formato a/b/aa (ano com 2 dígitos) é lido como MÊS/DIA/ANO (americano), porque
      na planilha da Maria existem 8/16/26, 8/20/26 etc. que só fazem sentido assim.
    """
    t = norm(texto)
    ano_padrao = ano_padrao or hoje.year
    if not t:
        return None, "sem data", True, None

    # ISO 2026-08-06
    m = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", t)
    if m:
        d = _data_valida(int(m[1]), int(m[2]), int(m[3]))
        return (d, None, False, m[0]) if d else (None, f"data inválida '{m[0]}'", True, m[0])

    # dd/mm/aaaa
    m = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", t)
    if m:
        d = _data_valida(int(m[3]), int(m[2]), int(m[1]))
        return (d, None, False, m[0]) if d else (None, f"data inválida '{m[0]}'", True, m[0])

    # a/b/aa -> americano (mês/dia/ano)
    m = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{2})\b", t)
    if m:
        a, b, aa = int(m[1]), int(m[2]), 2000 + int(m[3])
        if a <= 12:
            d = _data_valida(aa, a, b)
            if d:
                return d, f"formato americano, lida como {d.strftime('%d/%m/%Y')}", False, m[0]
        d = _data_valida(aa, b, a)
        return (d, None, False, m[0]) if d else (None, f"data inválida '{m[0]}'", True, m[0])

    # dd/mm (sem ano)
    m = re.search(r"\b(\d{1,2})/(\d{1,2})\b", t)
    if m:
        dd, mm = int(m[1]), int(m[2])
        if mm > 12 and dd <= 12:
            dd, mm = mm, dd
        d = _data_valida(ano_padrao, mm, dd)
        if not d:
            return None, f"data inválida '{m[0]}'", True, m[0]
        if d > hoje:  # ano errado? tenta ano anterior
            d2 = _data_valida(ano_padrao - 1, mm, dd)
            return d2, f"'{m[0]}' sem ano, assumido {d2.year}", True, m[0]
        return d, None, False, m[0]

    # "12 de marco"
    m = re.search(r"\b(\d{1,2})\s*(?:de\s*)?(" + "|".join(sorted(MESES, key=len, reverse=True)) + r")\b", t)
    if m:
        d = _data_valida(ano_padrao, MESES[m[2]], int(m[1]))
        return (d, None, False, m[0]) if d else (None, f"data inválida '{m[0]}'", True, m[0])

    # relativas
    rel = None
    m = re.search(r"\b(hoje|hj)\b", t)
    if m:
        rel, trecho = hoje, m[0]
    elif re.search(r"\banteontem\b", t):
        rel, trecho = hoje - timedelta(days=2), "anteontem"
    elif re.search(r"\bontem\b", t):
        rel, trecho = hoje - timedelta(days=1), "ontem"
    else:
        for nome, wd in DIAS_SEMANA.items():
            m = re.search(r"\b" + nome + r"(-feira)?\b", t)
            if m:
                delta = (hoje.weekday() - wd) % 7
                rel, trecho = hoje - timedelta(days=delta), m[0]
                break
    if rel is not None:
        if relativo_ok:
            return rel, None, False, trecho
        return None, f"data relativa ('{trecho}') sem referência: qual foi o dia?", True, trecho

    return None, None, False, None  # nada parecido com data


# ----------------------------------------------------------------------------
# VALOR
# ----------------------------------------------------------------------------
def _num(s):
    s = s.strip().replace("r$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_valor(texto):
    """Devolve (valor|None, frete|None, nota|None, confirmar:bool). Texto já sem data/parcelas/telefone/rastreio."""
    t = norm(texto)
    if not t:
        return None, None, "sem valor", True

    # "35+12 frete", "35 mais 12 de frete"
    m = re.search(r"(-?\d+(?:[.,]\d+)?)\s*(?:\+|mais)\s*(\d+(?:[.,]\d+)?)\s*(?:de\s*)?(?:frete|envio|correio)", t)
    if m:
        return _num(m[1]), _num(m[2]), f"'{m[0]}': confirmar se R$ {_num(m[1]):.0f} já inclui o frete", True

    frete = None
    m = re.search(r"(?:frete|envio)\s*(?:de\s*)?r?\$?\s*(\d+(?:[.,]\d+)?)", t)
    if m:
        frete = _num(m[1])
        t = t.replace(m[0], " ")

    # R$ 35 / 35,00 / -40
    m = re.search(r"(-?)\s*r?\$?\s*(\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d+(?:[.,]\d{1,2})?)(?!\s*x\b)(?!\s*vezes)(?!\s*/)", t)
    if m:
        v = _num(m[2])
        if v is None:
            return None, frete, "valor ilegível", True
        if m[1] == "-":
            v = -v
        return v, frete, None, False
    return None, frete, "sem valor", True


# ----------------------------------------------------------------------------
# PAGAMENTO / CARTÃO / PARCELAS / ENTREGA
# ----------------------------------------------------------------------------
def parse_parcelas(texto):
    t = norm(texto)
    m = re.search(r"\b(\d{1,2})\s*x\b", t) or re.search(r"\b(?:em\s*)?(\d{1,2})\s*(?:vezes|parcelas)\b", t)
    if m:
        return int(m[1]), m[0]
    if re.fullmatch(r"\d{1,2}", t):
        return int(t), t
    return None, None


def parse_cartao(texto):
    t = norm(texto)
    for k, v in [("visa", "Visa"), ("master", "Master"), ("mastercard", "Master"), ("elo", "Elo"),
                 ("hipercard", "Hipercard"), ("amex", "Amex")]:
        if re.search(r"\b" + k + r"\w*", t):
            return v
    return None


def parse_pagamento(texto):
    """Devolve (pagamento|None, nota|None). Fiado = ainda não recebeu."""
    t = norm(texto)
    if not t:
        return None, None
    if re.search(r"\b(fiado|depois paga|fica devendo|ficou devendo|vai pagar|paga semana|paga depois|deve|devendo|pendurad)", t):
        return "Fiado", None
    if re.search(r"\bcredito\b", t):
        return "Crédito", None
    if re.search(r"\bdebito\b", t):
        return "Débito", None
    if re.search(r"\bpix\b", t):
        return "Pix", None
    if re.search(r"\b(dinheiro|especie|em especie|cash)\b", t):
        return "Dinheiro", None
    if re.search(r"\b(maquinha|maquininha|cartao|maquina)\b", t):
        return "Cartão", "pagou na maquininha: débito ou crédito?"
    return None, None


def parse_entrega(texto):
    t = norm(texto)
    if not t:
        return None
    if re.search(r"\b(correio|correios|sedex|pac|postei|enviei pelo correio|mandei pelo correio)\b", t):
        return "Correios"
    if re.search(r"\b(motoboy|moto|entregador|uber|lalamove|loggi)\b", t):
        return "Motoboy"
    if re.search(r"\b(loja|presencial|veio buscar|buscou|levou|retirou|pegou na loja|na mao|em maos|veio pegar)\b", t):
        return "Retirou na loja"
    return None


# ----------------------------------------------------------------------------
# PRODUTO / CLIENTE
# ----------------------------------------------------------------------------
PRODUTOS_CANONICOS = {
    "brinc argola": "brinco argola",
    "brinco de argola": "brinco argola",
    "argola": "brinco argola",
}


def canon_produto(p):
    n = norm(p)
    n = PRODUTOS_CANONICOS.get(n, n)
    return cap(n)


def classificar_tipo(texto, valor):
    t = norm(texto)
    if re.search(r"\b(comprei|compra de estoque|fornecedor|reposicao|repus)\b", t):
        return "Compra de estoque"
    if re.search(r"\b(devolveu|devolucao|devolvi|estorn|troca)\b", t) or (valor is not None and valor < 0):
        return "Devolução"
    return "Venda"


CLIENTE_NAO_IDENTIFICADO = re.compile(r"^(cliente|a moca|o moco|uma moca|um moco|nao lembro|nao sei|esqueci|alguem|uma pessoa|um cara|uma mulher|um homem|nao anotei)$")

# ----------------------------------------------------------------------------
# 1) LINHA DA PLANILHA VELHA -> venda organizada
# ----------------------------------------------------------------------------
def interpretar_linha(raw, n_original, hoje, ano_padrao=2026):
    """
    raw: dict com as colunas originais da Maria:
      data, oq vendeu, pra quem, valor, como pagou, cartao, vezes, entrega ou loja, rastreio, obs
    Devolve dict com COLUNAS + 'Tipo' (Venda / Devolução / Compra de estoque).
    """
    g = lambda k: "" if vazio(raw.get(k)) else str(raw.get(k)).strip()
    notas, confirmar = [], False
    todas = " | ".join(g(k) for k in raw)

    # rastreio: na coluna certa ou perdido em qualquer outra
    rastreio = None
    for m in RE_RASTREIO.finditer(todas):
        rastreio = m[1].upper()
    if g("rastreio") and not RE_RASTREIO.search(g("rastreio")):
        notas.append(f"rastreio '{g('rastreio')}' em formato estranho")
        confirmar = True

    # data
    data, nota, conf, _ = parse_data(g("data"), hoje, ano_padrao, relativo_ok=False)
    if nota:
        notas.append(nota)
    confirmar |= conf
    if data and (data < date(ano_padrao, 7, 1) or data > hoje):
        notas.append(f"data {data.strftime('%d/%m/%Y')} fora do período das outras vendas: confere?")
        confirmar = True

    # valor
    valor, frete, nota, conf = parse_valor(g("valor"))
    if nota:
        notas.append(nota)
    confirmar |= conf

    # tipo (venda / devolução / compra)
    tipo = classificar_tipo(g("oq vendeu") + " " + g("obs") + " " + g("pra quem"), valor)

    # produto
    produto = canon_produto(g("oq vendeu")) if g("oq vendeu") else ""
    if not produto:
        notas.append("sem produto")
        confirmar = True

    # cliente
    cliente_raw = g("pra quem")
    cliente = cliente_raw
    obs = g("obs")
    if RE_RASTREIO.fullmatch(cliente_raw.strip()):
        cliente = ""
        m = re.search(r"\bera (?:a|o)\s+(.+)$", norm(obs))
        if m:
            cliente = cap(m[1])
            notas.append(f"código de rastreio estava no lugar do nome; cliente vem da obs ('{obs}')")
        else:
            notas.append("código de rastreio no lugar do nome: quem foi?")
            confirmar = True
    elif RE_TELEFONE.fullmatch(cliente_raw.replace(" ", "")):
        cliente = f"Tel. {cliente_raw}"
        notas.append("só telefone, sem nome: quem é?")
        confirmar = True
    elif CLIENTE_NAO_IDENTIFICADO.match(norm(cliente_raw)):
        cliente = "Não identificado" if norm(cliente_raw) in ("nao lembro", "nao sei", "esqueci") else cap(cliente_raw)
        notas.append(f"cliente não identificado ('{cliente_raw}')")
        confirmar = True
    else:
        cliente = cap(cliente_raw)

    # pagamento (coluna, depois obs), cartão, parcelas
    pag, nota = parse_pagamento(g("como pagou"))
    if not pag:
        pag, nota = parse_pagamento(obs)
    if not pag and tipo == "Venda":
        notas.append("sem forma de pagamento")
        confirmar = True
    if nota:
        notas.append(nota)
        confirmar = True
    # "fica devendo" na obs vira fiado mesmo se a coluna diz outra coisa
    pag_obs, _ = parse_pagamento(obs)
    if pag_obs == "Fiado":
        pag = "Fiado"

    cartao = parse_cartao(g("cartao")) or parse_cartao(g("como pagou"))
    parcelas, _ = parse_parcelas(g("vezes"))
    if parcelas is None:
        parcelas, _ = parse_parcelas(g("como pagou"))
    if parcelas and parcelas > 1 and pag in (None, "Cartão", "Débito"):
        pag = "Crédito"
        notas = [n for n in notas if "maquininha" not in n]
        confirmar = any(n for n in notas if n) and confirmar
    if cartao and not pag:
        pag = "Cartão"

    # entrega
    entrega = parse_entrega(g("entrega ou loja")) or parse_entrega(obs)
    if not entrega and rastreio:
        entrega = "Correios"
    entrega = entrega or "Não informado"

    # recebido?
    if tipo != "Venda":
        recebido = ""
    else:
        recebido = "Não" if pag == "Fiado" else "Sim"

    if valor is None and tipo == "Venda":
        confirmar = True

    return {
        "Tipo": tipo,
        "Data": data,
        "Produto": produto,
        "Cliente": cliente,
        "Valor (R$)": abs(valor) if (valor is not None and tipo == "Devolução") else valor,
        "Frete (R$)": frete,
        "Pagamento": pag or "Não informado",
        "Cartão": cartao or "",
        "Parcelas": parcelas if parcelas else (1 if pag in ("Crédito", "Débito") else None),
        "Entrega": entrega,
        "Rastreio": rastreio or "",
        "Recebido?": recebido,
        "Obs": obs,
        "Confirmar?": "SIM" if confirmar else "",
        "Nota do sistema": "; ".join(n for n in notas if n),
        "Nº original": n_original,
    }


# ----------------------------------------------------------------------------
# 2) FRASE SOLTA -> venda organizada  ("vendi um colar pra Ana, 89 no pix, mandei pelo correio")
# ----------------------------------------------------------------------------
STOP_NOME = {"por", "no", "na", "em", "pagou", "pagando", "pago", "que", "e", "com", "via", "pelo", "pela",
             "mandei", "enviei", "levou", "veio", "buscou", "hoje", "hj", "ontem", "foi", "paga", "vai"}
ARTIGOS = {"um", "uma", "o", "a", "os", "as", "1", "2", "3", "dois", "duas", "tres", "três", "umas", "uns"}


def _limpa_nome(palavras):
    palavras = [p for p in palavras if p]
    # permite "a vizinha", "tia da Ana", "Seu Jorge", "menina do insta"
    out = []
    for i, p in enumerate(palavras):
        if norm(p) in STOP_NOME or re.match(r"^-?r?\$?\d", norm(p)):
            break
        out.append(p)
        if len(out) >= 5:
            break
    nome = " ".join(out).strip(" ,.")
    return nome


def interpretar_frase(texto, hoje):
    """
    Devolve dict com COLUNAS + 'Tipo' + 'Faltando' (lista de campos essenciais ausentes)
    + 'Pergunta' (a UMA pergunta a fazer de volta, ou None).
    """
    original = texto.strip()
    t = original
    notas = []

    # rastreio
    rastreio = None
    for m in RE_RASTREIO.finditer(t):
        rastreio = m[1].upper()
    t = RE_RASTREIO.sub(" ", t)

    # telefone
    telefone = None
    m = RE_TELEFONE.search(t)
    if m:
        telefone = m[1]
        t = t.replace(m[0], " ")

    # data
    data, nota, conf, trecho = parse_data(t, hoje, relativo_ok=True)
    if trecho:
        t = re.sub(re.escape(trecho), " ", t, flags=re.I)
    if nota:
        notas.append(nota)
    if data is None:
        data = hoje

    # parcelas
    parcelas, trecho = parse_parcelas(t)
    if trecho:
        t = re.sub(re.escape(trecho), " ", t, flags=re.I)

    # quantidade ("vendi 2 presilhas", "2 presilhas pra Ana") NÃO é valor
    t = re.sub(r"\b(vendi|vendeu|vendemos|comprei|saiu|sairam|mandei|enviei)\s+(\d{1,3})\s+(?=[a-záéíóúãõç])", r"\1 ", t, flags=re.I)
    t = re.sub(r"^\s*(\d{1,3})\s+(?=[a-záéíóúãõç])", "", t, flags=re.I)

    # valor (frete etc.)
    valor, frete, nota, _ = parse_valor(t)
    if nota and valor is not None:
        notas.append(nota)

    tipo = classificar_tipo(t, valor)
    pag, nota = parse_pagamento(t)
    if nota:
        notas.append(nota)
    cartao = parse_cartao(t)
    if parcelas and parcelas > 1 and pag in (None, "Cartão", "Débito"):
        pag = "Crédito"
        notas = [n for n in notas if "maquininha" not in n]
    if cartao and not pag:
        pag = "Cartão"
    entrega = parse_entrega(t) or ("Correios" if rastreio else None)

    # produto e cliente: trabalha por pedaços separados por vírgula / ponto / " e "
    pedacos = [p.strip() for p in re.split(r"[,.;]|\be\b(?=\s+(?:mandei|enviei|levou|pagou|ela|ele|pediu|veio))", t) if p.strip()]
    produto, cliente = "", ""
    usado = set()
    for i, ped in enumerate(pedacos):
        n = norm(ped)
        # "a Carla devolveu a pulseira" -> cliente antes do verbo
        mpre = re.match(r"^(?:a|o)\s+([a-z]+)\s+(devolveu|comprou|levou|pegou|pagou|encomendou)\b", n)
        if mpre and not cliente:
            cliente = mpre[1]
        mv = re.search(r"\b(vendi|vendeu|vendemos|saiu|sairam|foi vendid[oa]|devolveu|comprei|anota|anotar|registra|mandei|enviei|postei|despachei|entreguei)\b\s*(.*)", n)
        corpo = mv[2] if mv else n
        # cliente depois de "pra/para/pro"
        mc = re.search(r"\b(?:pra|para|pro|pros|pras|p/)\s+(.+)$", corpo)
        if mc and not cliente:
            cliente = _limpa_nome(mc[1].split())
            corpo = corpo[:mc.start()].strip()
            usado.add(i)
        # "o colar da Juliana" -> produto + cliente (corta antes de pagamento/entrega)
        corte = []
        for p in corpo.split():
            if parse_pagamento(p)[0] or parse_entrega(p) or norm(p) in STOP_NOME:
                break
            corte.append(p)
        corpo = " ".join(corte)
        md = re.match(r"^(.+?)\s+d[ao]\s+([a-z]+(?:\s+[a-z]+)?)\s*$", corpo)
        if md and not cliente and mv and mv[1] in ("mandei", "enviei", "postei", "despachei", "entreguei", "vendi", "vendeu"):
            cliente = md[2]
            corpo = md[1]
            usado.add(i)
        if (mv or i == 0) and not produto:
            palavras = [p for p in corpo.split() if norm(p) not in ARTIGOS]
            palavras = [p for p in palavras if not re.match(r"^-?r?\$?\d", p)]
            # corta em palavras de pagamento/entrega
            prod = []
            for p in palavras:
                if parse_pagamento(p)[0] or parse_entrega(p) or norm(p) in STOP_NOME:
                    break
                prod.append(p)
            produto = " ".join(prod).strip()
            if produto:
                usado.add(i)

    # obs = pedaços que não serviram pra nada estruturado
    obs = []
    for i, ped in enumerate(pedacos):
        if i in usado:
            continue
        n = norm(ped).strip()
        reconhecido = (parse_pagamento(n)[0] or parse_entrega(n) or parse_cartao(n)
                       or re.search(r"\d", n)
                       or re.fullmatch(r"(codigo( de rastreio)?|rastreio|rastreamento|cod)", n))
        if not reconhecido and len(n) > 2:
            obs.append(ped.strip())
    if telefone and not cliente:
        cliente = f"Tel. {telefone}"
    elif telefone:
        obs.append(f"tel. {telefone}")

    produto = canon_produto(produto) if produto else ""
    cliente = titulo(cliente)
    confirmar = False
    if cliente and CLIENTE_NAO_IDENTIFICADO.match(norm(cliente)) or norm(cliente) in ("moca", "moco", "mulher", "homem", "menina", "menino", "pessoa"):
        notas.append(f"cliente não identificado ('{cliente}')")
        cliente = "Não identificado"
        confirmar = True
    if not cliente and tipo == "Venda":
        cliente = "Não informado"
        notas.append("sem o nome de quem comprou")
        confirmar = True

    # essenciais: valor, produto, pagamento (uma pergunta de cada vez)
    faltando = []
    if valor is None and tipo != "Compra de estoque":
        faltando.append("valor")
    if not produto:
        faltando.append("produto")
    if not pag and tipo == "Venda":
        faltando.append("pagamento")

    # UMA pergunta só, a mais importante
    perguntas = {
        "valor": "Quanto foi?",
        "produto": "O que você vendeu?",
        "pagamento": "Como pagou? (pix, dinheiro, débito, crédito ou fiado)",
        "cliente": "Pra quem foi?",
    }
    pergunta = perguntas[faltando[0]] if faltando else None

    return {
        "Tipo": tipo,
        "Data": data,
        "Produto": produto,
        "Cliente": cliente,
        "Valor (R$)": abs(valor) if (valor is not None and tipo == "Devolução") else valor,
        "Frete (R$)": frete,
        "Pagamento": pag or "Não informado",
        "Cartão": cartao or "",
        "Parcelas": parcelas if parcelas else (1 if pag in ("Crédito", "Débito") else None),
        "Entrega": entrega or "Não informado",
        "Rastreio": rastreio or "",
        "Recebido?": ("" if tipo != "Venda" else ("Não" if pag == "Fiado" else "Sim")),
        "Obs": "; ".join(obs),
        "Confirmar?": "SIM" if confirmar else "",
        "Nota do sistema": "; ".join(notas),
        "Nº original": "",
        "Faltando": faltando,
        "Pergunta": pergunta,
        "Frase": original,
    }


def completar(venda, resposta, hoje):
    """Preenche o campo que estava faltando com a resposta da Maria. Devolve a venda atualizada."""
    if not venda.get("Faltando"):
        return venda
    campo = venda["Faltando"][0]
    r = resposta.strip()
    if campo == "valor":
        v, frete, _, _ = parse_valor(r)
        venda["Valor (R$)"] = v
        if frete:
            venda["Frete (R$)"] = frete
    elif campo == "produto":
        venda["Produto"] = canon_produto(r)
    elif campo == "pagamento":
        pag, _ = parse_pagamento(r)
        parcelas, _ = parse_parcelas(r)
        cartao = parse_cartao(r)
        if parcelas and parcelas > 1:
            pag = "Crédito"
            venda["Parcelas"] = parcelas
        venda["Pagamento"] = pag or "Não informado"
        if cartao:
            venda["Cartão"] = cartao
        venda["Recebido?"] = "Não" if pag == "Fiado" else "Sim"
    elif campo == "cliente":
        venda["Cliente"] = titulo(_limpa_nome(r.split()) or r)
    venda["Faltando"] = [f for f in venda["Faltando"] if f != campo]
    # ainda falta algo? pergunta a próxima (uma de cada vez)
    perguntas = {"valor": "Quanto foi?", "produto": "O que você vendeu?",
                 "pagamento": "Como pagou? (pix, dinheiro, débito, crédito ou fiado)", "cliente": "Pra quem foi?"}
    venda["Pergunta"] = perguntas[venda["Faltando"][0]] if venda["Faltando"] else None
    return venda


def resumo_venda(v):
    """Frase curta de confirmação: 'Colar, Ana, R$ 89, Pix, Correios.'"""
    partes = []
    if v.get("Tipo") and v["Tipo"] != "Venda":
        partes.append(v["Tipo"].upper())
    if v.get("Produto"):
        partes.append(v["Produto"])
    if v.get("Cliente"):
        partes.append(v["Cliente"])
    if v.get("Valor (R$)") is not None:
        s = f"R$ {v['Valor (R$)']:.2f}".replace(".", ",")
        if v.get("Frete (R$)"):
            s += f" + frete R$ {v['Frete (R$)']:.2f}".replace(".", ",")
        partes.append(s)
    pag = v.get("Pagamento")
    if pag and pag != "Não informado":
        p = pag
        if v.get("Parcelas") and v["Parcelas"] > 1:
            p += f" {v['Parcelas']}x"
        if v.get("Cartão"):
            p += f" {v['Cartão']}"
        partes.append(p)
    if v.get("Entrega") and v["Entrega"] != "Não informado":
        partes.append(v["Entrega"])
    if v.get("Rastreio"):
        partes.append(f"rastreio {v['Rastreio']}")
    if v.get("Data"):
        partes.append(v["Data"].strftime("%d/%m"))
    if v.get("Obs"):
        partes.append(f"obs: {v['Obs']}")
    return ", ".join(partes) + "."
