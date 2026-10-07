# -*- coding: utf-8 -*-
"""
ia.py — camada OPCIONAL de inteligência artificial.

Se existir uma chave (variável de ambiente ANTHROPIC_API_KEY ou OPENAI_API_KEY),
o assistente usa o modelo pra entender frases mais soltas e responder perguntas
livres sobre a tabela. Sem chave, tudo continua funcionando com as regras de
maria.py e consultas.py — a IA é um reforço, não uma dependência.
"""
import os
import json
import re
from datetime import date

import requests

from maria import COLUNAS, canon_produto, titulo

ANTHROPIC = os.environ.get("ANTHROPIC_API_KEY")
OPENAI = os.environ.get("OPENAI_API_KEY")
MODELO_ANTHROPIC = os.environ.get("MODELO_IA", "claude-sonnet-4-6")
MODELO_OPENAI = os.environ.get("MODELO_IA", "gpt-4o-mini")


def disponivel():
    return bool(ANTHROPIC or OPENAI)


def _chamar(system, user, max_tokens=700):
    if ANTHROPIC:
        r = requests.post("https://api.anthropic.com/v1/messages",
                          headers={"x-api-key": ANTHROPIC, "anthropic-version": "2023-06-01",
                                   "content-type": "application/json"},
                          json={"model": MODELO_ANTHROPIC, "max_tokens": max_tokens, "system": system,
                                "messages": [{"role": "user", "content": user}]}, timeout=30)
        r.raise_for_status()
        return "".join(b.get("text", "") for b in r.json()["content"])
    r = requests.post("https://api.openai.com/v1/chat/completions",
                      headers={"Authorization": f"Bearer {OPENAI}", "content-type": "application/json"},
                      json={"model": MODELO_OPENAI, "max_tokens": max_tokens,
                            "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}]},
                      timeout=30)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]


SYSTEM_EXTRAIR = """Você organiza as vendas de uma loja de acessórios. A dona, Maria, escreve do jeito dela.
Transforme a frase em JSON com EXATAMENTE estas chaves:
tipo ("Venda" | "Devolução" | "Compra de estoque"), data ("AAAA-MM-DD" ou null), produto, cliente,
valor (número ou null), frete (número ou null), pagamento ("Pix"|"Dinheiro"|"Débito"|"Crédito"|"Fiado"|"Cartão"|null),
cartao ("Visa"|"Master"|"Elo"|"Hipercard"|"Amex"|null), parcelas (inteiro ou null),
entrega ("Retirou na loja"|"Correios"|"Motoboy"|null), rastreio (string ou null), obs (string ou null),
duvida (string curta se algo ficou ambíguo, senão null).
Regras: "fiado", "depois paga", "fica devendo" = Fiado. "loja", "presencial", "veio buscar", "levou" = Retirou na loja.
"maquininha" sem dizer débito/crédito = "Cartão". Parcelas > 1 implica Crédito. Nunca invente valor: se não disse, null.
Responda SÓ o JSON, sem texto em volta."""


def extrair(texto, hoje, clientes_conhecidos=()):
    """Devolve dict no formato de maria.interpretar_frase ou None se a IA falhar."""
    try:
        user = f"Hoje é {hoje.isoformat()} ({hoje.strftime('%A')}). Clientes já conhecidas: {', '.join(list(clientes_conhecidos)[:60])}.\nFrase: {texto}"
        bruto = _chamar(SYSTEM_EXTRAIR, user, 400)
        bruto = re.sub(r"```json|```", "", bruto).strip()
        j = json.loads(bruto)
    except Exception:
        return None
    d = None
    if j.get("data"):
        try:
            d = date.fromisoformat(j["data"])
        except ValueError:
            d = None
    pag = j.get("pagamento")
    parcelas = j.get("parcelas")
    if parcelas and parcelas > 1:
        pag = "Crédito"
    tipo = j.get("tipo") or "Venda"
    faltando = []
    if j.get("valor") is None and tipo != "Compra de estoque":
        faltando.append("valor")
    if not j.get("produto"):
        faltando.append("produto")
    if not pag and tipo == "Venda":
        faltando.append("pagamento")
    perguntas = {"valor": "Quanto foi?", "produto": "O que você vendeu?",
                 "pagamento": "Como pagou? (pix, dinheiro, débito, crédito ou fiado)"}
    notas = [j["duvida"]] if j.get("duvida") else []
    cliente = titulo(j.get("cliente") or "")
    confirmar = bool(notas)
    if not cliente and tipo == "Venda":
        cliente, confirmar = "Não informado", True
        notas.append("sem o nome de quem comprou")
    return {
        "Tipo": tipo, "Data": d or hoje, "Produto": canon_produto(j.get("produto") or ""),
        "Cliente": cliente, "Valor (R$)": j.get("valor"), "Frete (R$)": j.get("frete"),
        "Pagamento": pag or "Não informado", "Cartão": j.get("cartao") or "",
        "Parcelas": parcelas or (1 if pag in ("Crédito", "Débito") else None),
        "Entrega": j.get("entrega") or ("Correios" if j.get("rastreio") else "Não informado"),
        "Rastreio": (j.get("rastreio") or "").upper(),
        "Recebido?": "" if tipo != "Venda" else ("Não" if pag == "Fiado" else "Sim"),
        "Obs": j.get("obs") or "", "Confirmar?": "SIM" if confirmar else "",
        "Nota do sistema": "; ".join(notas), "Nº original": "",
        "Faltando": faltando, "Pergunta": perguntas[faltando[0]] if faltando else None, "Frase": texto,
    }


SYSTEM_RESPONDER = """Você é o assistente da Maria, dona de uma loja de acessórios. Ela não entende de tecnologia.
Responda a pergunta dela usando SÓ a tabela de vendas em CSV abaixo. Português simples, curto, direto, sem jargão.
Valores em R$ com vírgula. Se a tabela não tem a resposta, diga que não encontrou. Não invente nada."""


def responder(pergunta, df, hoje):
    try:
        csv = df[[c for c in COLUNAS if c not in ("Nota do sistema", "Nº original", "Confirmar?")]].to_csv(index=False)
        return _chamar(SYSTEM_RESPONDER, f"Hoje é {hoje.isoformat()}.\n\nTABELA:\n{csv}\n\nPERGUNTA: {pergunta}", 600)
    except Exception:
        return None
