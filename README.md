# Vendas da Maria — assistente de vendas por conversa

A Maria conta a venda do jeito dela ("vendi um colar pra Ana, 89 no pix, mandei pelo correio"),
o assistente anota na planilha dela e responde perguntas ("quem me deve?", "o que está no correio?").

## Rodar
    pip install -r requirements.txt
    streamlit run app.py

Na primeira vez, se não existir "Vendas da Maria - organizada.xlsx", o app lê "Vendas da maria.xlsx"
(a planilha original) e gera a versão organizada sozinho. Para refazer a migração na mão:
    python migrar.py "Vendas da maria.xlsx" "Vendas da Maria - organizada.xlsx"

IA opcional: defina ANTHROPIC_API_KEY ou OPENAI_API_KEY no ambiente. Sem chave, tudo funciona por regras.

## Arquivos
- maria.py      — o interpretador: lê linha velha da planilha E frase nova. Não chuta: marca "Confirmar?".
- migrar.py     — gera a planilha organizada (abas Resumo, Vendas, Não é venda, Original, Como usar).
- consultas.py  — grava na planilha e responde as perguntas da Maria por regras.
- ia.py         — camada opcional de IA (extração em JSON + perguntas livres).
- app.py        — a tela de conversa (Streamlit), cores do Grupo Fácil.
- Vendas da Maria - organizada.xlsx — a planilha dela já migrada (69 vendas, 2 não-vendas, 14 pra confirmar).

## Frases pra demo
1. vendi um colar pra Ana, 89 no pix, mandei pelo correio
2. brinco argola pra vizinha 35 dinheiro, ela veio buscar, pediu embrulho
3. vendi um anel pro Marcos            -> ele pergunta "Quanto foi?" e depois "Como pagou?"
4. 2 presilhas pra menina do insta, 30, ela paga semana que vem   -> vira fiado
5. a Carla devolveu a pulseira, 40 no dinheiro                    -> vai pra "Não é venda"
6. quem me deve?  /  o que está no correio?  /  cadê o código da encomenda da Ana?  /  quanto vendi em setembro?
7. a Rita me pagou                                                -> marca recebido
