# -*- coding: utf-8 -*-
"""
app.py — o assistente da Maria.

Rodar:  streamlit run app.py
Ela conta a venda do jeito dela, o assistente anota na planilha e responde perguntas.
"""
import os
from datetime import date

import streamlit as st

import maria
import consultas
import ia
import migrar

ARQ = consultas.ARQUIVO_PADRAO
ORIG = "Vendas da maria.xlsx"

AZUL, AZUL_ESCURO, AZUL_CLARO, CINZA = "#3A9AD9", "#125496", "#E3F2FB", "#8C8C8C"

st.set_page_config(page_title="Vendas da Maria", page_icon="🛍️", layout="centered",
                   initial_sidebar_state="collapsed")

st.markdown(f"""
<style>
  html, body, [data-testid="stAppViewContainer"] {{ font-size: 18px; }}
  [data-testid="stChatMessage"] p, [data-testid="stChatMessage"] li {{ font-size: 1.12rem; line-height: 1.55; }}
  [data-testid="stChatMessage"] {{ border-radius: 16px; padding: .4rem .8rem; }}
  .stButton > button {{ font-size: 1.02rem; padding: .55rem .9rem; border-radius: 14px;
      border: 2px solid {AZUL}; color: {AZUL_ESCURO}; background: {AZUL_CLARO}; font-weight: 600; width: 100%; }}
  .stButton > button:hover {{ background: {AZUL}; color: white; border-color: {AZUL}; }}
  .stChatInput textarea {{ font-size: 1.1rem !important; }}
  h1 {{ color: {AZUL_ESCURO}; margin-bottom: 0; }}
  .sub {{ color: {CINZA}; font-size: 1.02rem; margin-top: -.2rem; margin-bottom: 1rem; }}
  .card {{ border: 2px solid {AZUL}; border-radius: 16px; padding: 1rem 1.2rem; background: {AZUL_CLARO};
      font-size: 1.15rem; color: {AZUL_ESCURO}; margin: .5rem 0 .8rem 0; }}
  .rodape {{ color: {CINZA}; font-size: .85rem; text-align: center; margin-top: 1.5rem; }}
</style>
""", unsafe_allow_html=True)

hoje = date.today()

# ---------------------------------------------------------------- primeira vez: organiza a planilha dela
if not os.path.exists(ARQ):
    if os.path.exists(ORIG):
        with st.spinner("Organizando a sua planilha pela primeira vez. Só um instante..."):
            migrar.migrar(ORIG, ARQ, hoje)
    else:
        st.error(f"Não achei a planilha. Coloque o arquivo '{ORIG}' na mesma pasta deste aplicativo.")
        st.stop()

df = consultas.carregar(ARQ)

# ---------------------------------------------------------------- estado da conversa
ss = st.session_state
BOAS_VINDAS = ("assistant",
    "Oi, Maria! Me conta o que você vendeu, do seu jeito. Por exemplo: "
    "*vendi um colar pra Ana, 89 no pix, mandei pelo correio*. "
    "Ou me pergunta alguma coisa, tipo *quem me deve?*")
ss.setdefault("msgs", [BOAS_VINDAS])
ss.setdefault("pendente", None)     # venda esperando resposta de UMA pergunta
ss.setdefault("confirmar", None)    # venda esperando o "Certo"
ss.setdefault("corrigindo", False)


def diz(texto):
    ss.msgs.append(("assistant", texto))


def fala(texto):
    ss.msgs.append(("user", texto))


def pedir_confirmacao(v):
    ss.confirmar = v
    aviso = ""
    if v.get("Nota do sistema"):
        aviso = f"\n\n⚠️ {v['Nota do sistema']}."
    diz(f"Anotei assim: **{maria.resumo_venda(v)}**{aviso}\n\nEstá certo?")


def tratar(texto):
    global df
    # 1) ela está respondendo a pergunta que eu fiz
    if ss.pendente:
        v = maria.completar(ss.pendente, texto, hoje)
        if v["Pergunta"]:
            ss.pendente = v
            diz(v["Pergunta"])
            return
        ss.pendente = None
        pedir_confirmacao(v)
        return

    # 2) "a Rita me pagou"
    if "me pagou" in maria.norm(texto):
        r = consultas.responder(texto, df, hoje)
        if isinstance(r, dict) and r.get("acao") == "receber":
            n = consultas.marcar_recebido(r["cliente"], ARQ)
            df = consultas.carregar(ARQ)
            diz(f"Boa! Marquei {n} venda(s) de **{r['cliente']}** como recebida(s). ✅" if n
                else f"**{r['cliente']}** não tinha nada pendente.")
        else:
            diz(r)
        return

    # 3) pergunta
    if consultas.parece_pergunta(texto):
        r = consultas.responder(texto, df, hoje)
        if isinstance(r, str) and r.startswith("Não entendi") and ia.disponivel():
            r = ia.responder(texto, df, hoje) or r
        diz(r)
        return

    # 4) venda: regras primeiro; IA (se tiver) só quando as regras ficaram em dúvida
    v = maria.interpretar_frase(texto, hoje)
    if ia.disponivel() and (v["Faltando"] or v["Confirmar?"] == "SIM"):
        v2 = ia.extrair(texto, hoje, df["Cliente"].unique())
        if v2 and len(v2["Faltando"]) <= len(v["Faltando"]):
            v = v2
    if v["Pergunta"]:
        ss.pendente = v
        diz(v["Pergunta"])
        return
    pedir_confirmacao(v)


def gravar_confirmada(v):
    global df
    linha = consultas.gravar(v, ARQ)
    df = consultas.carregar(ARQ)
    ss.confirmar, ss.corrigindo = None, False
    extra = ""
    if v.get("Pagamento") == "Fiado":
        extra = f" Fica anotado que **{v['Cliente']}** está devendo."
    if v.get("Entrega") == "Correios" and not v.get("Rastreio"):
        extra += " Quando tiver o código de rastreio, me manda que eu guardo."
    if v.get("Tipo", "Venda") == "Venda":
        diz(f"Pronto, anotado! ✅ Já são {len(df)} vendas na planilha.{extra}")
    else:
        diz(f"Pronto, anotado na aba **Não é venda**, pra não entrar na soma das vendas. ✅{extra}")


# ---------------------------------------------------------------- cabeçalho
st.markdown("# 🛍️ Vendas da Maria")
st.markdown('<div class="sub">Conta o que vendeu. Eu anoto e acho depois.</div>', unsafe_allow_html=True)

# atalhos: ela nem precisa digitar as perguntas de sempre
c1, c2, c3, c4 = st.columns(4)
atalhos = [(c1, "Quem me deve?"), (c2, "No correio?"), (c3, "Vendi hoje?"), (c4, "Vendi no mês?")]
for col, rotulo in atalhos:
    if col.button(rotulo, key=f"atalho_{rotulo}"):
        fala(rotulo)
        tratar(rotulo)
        st.rerun()

# ---------------------------------------------------------------- conversa
for papel, texto in ss.msgs:
    with st.chat_message(papel, avatar="👩" if papel == "user" else "🛍️"):
        st.markdown(texto)

# ---------------------------------------------------------------- venda esperando o "Certo"
if ss.confirmar and not ss.corrigindo:
    b1, b2, b3 = st.columns([1.3, 1, 1])
    if b1.button("✅ Certo, anota", key="ok"):
        fala("Certo")
        gravar_confirmada(ss.confirmar)
        st.rerun()
    if b2.button("✏️ Corrigir", key="corrigir"):
        ss.corrigindo = True
        st.rerun()
    if b3.button("🗑️ Deixa pra lá", key="cancelar"):
        fala("Deixa pra lá")
        ss.confirmar = None
        diz("Tá bom, não anotei nada.")
        st.rerun()

if ss.confirmar and ss.corrigindo:
    v = ss.confirmar
    with st.form("corrigir_form"):
        st.markdown("**Corrige o que estiver errado e aperta Salvar:**")
        a, b = st.columns(2)
        data = a.date_input("Data", value=v.get("Data") or hoje, format="DD/MM/YYYY")
        produto = b.text_input("O que vendeu", value=v.get("Produto") or "")
        cliente = a.text_input("Pra quem", value="" if v.get("Cliente") in ("Não informado", "Não identificado") else v.get("Cliente", ""))
        valor = b.number_input("Valor (R$)", value=float(v["Valor (R$)"] or 0), min_value=0.0, step=1.0, format="%.2f")
        pag_ops = ["Pix", "Dinheiro", "Débito", "Crédito", "Fiado", "Cartão", "Não informado"]
        pagamento = a.selectbox("Como pagou", pag_ops, index=pag_ops.index(v.get("Pagamento") or "Não informado"))
        cart_ops = ["", "Visa", "Master", "Elo", "Hipercard", "Amex"]
        cartao = b.selectbox("Cartão", cart_ops, index=cart_ops.index(v.get("Cartão") or ""))
        parcelas = a.number_input("Em quantas vezes", value=int(v.get("Parcelas") or 1), min_value=1, max_value=12, step=1)
        ent_ops = ["Retirou na loja", "Correios", "Motoboy", "Não informado"]
        entrega = b.selectbox("Entrega", ent_ops, index=ent_ops.index(v.get("Entrega") or "Não informado"))
        rastreio = a.text_input("Código de rastreio", value=v.get("Rastreio") or "")
        obs = b.text_input("Observação", value=v.get("Obs") or "")
        s1, s2 = st.columns(2)
        salvar = s1.form_submit_button("💾 Salvar")
        voltar = s2.form_submit_button("↩️ Voltar")
    if salvar:
        v.update({"Data": data, "Produto": maria.canon_produto(produto), "Cliente": maria.titulo(cliente) or "Não informado",
                  "Valor (R$)": valor, "Pagamento": pagamento, "Cartão": cartao, "Parcelas": int(parcelas),
                  "Entrega": entrega, "Rastreio": rastreio.strip().upper(), "Obs": obs,
                  "Recebido?": "Não" if pagamento == "Fiado" else "Sim",
                  "Confirmar?": "SIM" if (maria.titulo(cliente) or "Não informado") == "Não informado" else "",
                  "Nota do sistema": "" if cliente else "sem o nome de quem comprou"})
        fala("Corrigi: " + maria.resumo_venda(v))
        gravar_confirmada(v)
        st.rerun()
    if voltar:
        ss.corrigindo = False
        st.rerun()

# ---------------------------------------------------------------- limpar conversa
if len(ss.msgs) > 1 and not ss.confirmar and not ss.pendente:
    _, cm, _ = st.columns([2, 1.3, 2])
    if cm.button("🧹 Limpar conversa", key="limpar"):
        ss.msgs = [BOAS_VINDAS]
        ss.pendente, ss.confirmar, ss.corrigindo = None, None, False
        st.rerun()

# ---------------------------------------------------------------- entrada
texto = st.chat_input("Conta a venda ou faz uma pergunta…")
if texto:
    fala(texto)
    tratar(texto)
    st.rerun()

# ---------------------------------------------------------------- ver a planilha
with st.expander("📋 Ver as últimas vendas na planilha"):
    mostrar = df[["Data", "Produto", "Cliente", "Valor (R$)", "Pagamento", "Entrega", "Rastreio", "Recebido?"]].tail(8).iloc[::-1].copy()
    mostrar["Data"] = [d.strftime("%d/%m/%Y") if consultas.eh_data(d) else "" for d in mostrar["Data"]]
    mostrar["Valor (R$)"] = [consultas.dinheiro(x) for x in mostrar["Valor (R$)"]]
    st.dataframe(mostrar, use_container_width=True, hide_index=True)
    st.caption(f"Arquivo: {ARQ} — é a sua planilha de sempre, só organizada. Pode abrir no Excel quando quiser.")

modo = "regras + IA" if ia.disponivel() else "regras (funciona sem internet)"
st.markdown(f'<div class="rodape">{len(df)} vendas registradas · modo: {modo}</div>', unsafe_allow_html=True)
