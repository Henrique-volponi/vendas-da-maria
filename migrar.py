# -*- coding: utf-8 -*-
"""
migrar.py — lê a planilha velha da Maria e gera a planilha organizada.

Uso:  python migrar.py "Vendas da maria.xlsx" "Vendas da Maria - organizada.xlsx"

Nada é apagado: a aba "Original" guarda a planilha dela do jeito que estava.
Nada é chutado: o que o sistema não teve certeza vai com "Confirmar? = SIM".
"""
import sys
from datetime import date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

from maria import (interpretar_linha, COLUNAS, PAGAMENTOS_VALIDOS, CARTOES_VALIDOS,
                   ENTREGAS_VALIDAS)

# ---------------- cores do Grupo Fácil (tiradas do logo) ----------------
AZUL = "3A9AD9"        # azul "Fácil"
AZUL_ESCURO = "125496"
AZUL_CLARO = "E3F2FB"
CINZA = "8C8C8C"       # cinza "Grupo"
CINZA_CLARO = "F4F6F8"
ATENCAO = "FFF4CC"     # linha pra confirmar
BRANCO = "FFFFFF"

FONTE = "Arial"
f_titulo = Font(name=FONTE, size=16, bold=True, color=AZUL_ESCURO)
f_sub = Font(name=FONTE, size=10, color=CINZA, italic=True)
f_cab = Font(name=FONTE, size=10, bold=True, color=BRANCO)
f_normal = Font(name=FONTE, size=10)
f_nota = Font(name=FONTE, size=9, color=CINZA)
f_destaque = Font(name=FONTE, size=12, bold=True, color=AZUL_ESCURO)
fill_cab = PatternFill("solid", fgColor=AZUL_ESCURO)
fill_cab2 = PatternFill("solid", fgColor=AZUL)
fill_alt = PatternFill("solid", fgColor=AZUL_CLARO)
fill_at = PatternFill("solid", fgColor=ATENCAO)
fill_cinza = PatternFill("solid", fgColor=CINZA_CLARO)
borda = Border(bottom=Side(style="thin", color="D9E2EC"))
centro = Alignment(horizontal="center", vertical="center")
esq = Alignment(horizontal="left", vertical="center", wrap_text=False)

LARGURAS = {"Data": 12, "Produto": 20, "Cliente": 18, "Valor (R$)": 12, "Frete (R$)": 11,
            "Pagamento": 13, "Cartão": 9, "Parcelas": 9, "Entrega": 17, "Rastreio": 16,
            "Recebido?": 11, "Obs": 30, "Confirmar?": 11, "Nota do sistema": 58, "Nº original": 11}


def cabecalho(ws, linha, colunas, fill=fill_cab):
    for j, nome in enumerate(colunas, 1):
        c = ws.cell(row=linha, column=j, value=nome)
        c.font, c.fill, c.alignment = f_cab, fill, centro
        ws.column_dimensions[get_column_letter(j)].width = LARGURAS.get(nome, 14)
    ws.row_dimensions[linha].height = 22


def escreve_linha(ws, linha, venda, colunas, alt=False):
    atencao = venda.get("Confirmar?") == "SIM"
    for j, nome in enumerate(colunas, 1):
        v = venda.get(nome)
        if isinstance(v, float) and v != v:
            v = None
        c = ws.cell(row=linha, column=j, value=v)
        c.font = f_nota if nome == "Nota do sistema" else f_normal
        c.border = borda
        c.alignment = centro if nome in ("Data", "Valor (R$)", "Frete (R$)", "Parcelas", "Recebido?",
                                          "Confirmar?", "Nº original", "Cartão") else esq
        if nome == "Data":
            c.number_format = "dd/mm/yyyy"
        if nome in ("Valor (R$)", "Frete (R$)"):
            c.number_format = '"R$" #,##0.00'
        if atencao:
            c.fill = fill_at
        elif alt:
            c.fill = fill_alt
    if atencao:
        ws.cell(row=linha, column=colunas.index("Confirmar?") + 1).font = Font(name=FONTE, size=10, bold=True, color="9C5700")


def validacoes(ws, colunas, ate=1000):
    """listas suspensas pra Maria não digitar errado quando usar a planilha direto"""
    def dv(nome, opcoes):
        col = get_column_letter(colunas.index(nome) + 1)
        d = DataValidation(type="list", formula1='"' + ",".join(o for o in opcoes if o) + '"',
                           allow_blank=True, showErrorMessage=True,
                           errorTitle="Escolha na lista", error="Escolhe uma das opções da lista, por favor.")
        ws.add_data_validation(d)
        d.add(f"{col}2:{col}{ate}")
    dv("Pagamento", PAGAMENTOS_VALIDOS)
    dv("Cartão", CARTOES_VALIDOS)
    dv("Entrega", ENTREGAS_VALIDAS)
    dv("Recebido?", ["Sim", "Não"])
    col = get_column_letter(colunas.index("Data") + 1)
    d = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2100,1,1)",
                       allow_blank=True, showErrorMessage=True, errorTitle="Data",
                       error="Escreve a data assim: 07/10/2026")
    ws.add_data_validation(d)
    d.add(f"{col}2:{col}{ate}")
    col = get_column_letter(colunas.index("Valor (R$)") + 1)
    d = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                       showErrorMessage=True, errorTitle="Valor", error="Só o número, ex.: 35 ou 35,50")
    ws.add_data_validation(d)
    d.add(f"{col}2:{col}{ate}")
    # linha nova marcada em amarelo se Confirmar? = SIM
    colM = get_column_letter(colunas.index("Confirmar?") + 1)
    ws.conditional_formatting.add(f"A2:{get_column_letter(len(colunas))}{ate}",
                                  FormulaRule(formula=[f'${colM}2="SIM"'], fill=fill_at))


def migrar(entrada, saida, hoje=None):
    hoje = hoje or date.today()
    raw = pd.read_excel(entrada, dtype=str)
    vendas, nao_vendas = [], []
    chaves = {}
    for i, r in raw.iterrows():
        if r.isna().all():
            continue
        v = interpretar_linha(r.to_dict(), i + 2, hoje)
        if v["Tipo"] == "Venda":
            # duplicata exata?
            k = (str(v["Data"]), v["Produto"], v["Cliente"], v["Valor (R$)"], v["Pagamento"])
            if k in chaves:
                v["Confirmar?"] = "SIM"
                v["Nota do sistema"] = (v["Nota do sistema"] + "; " if v["Nota do sistema"] else "") + \
                    f"igual à linha original {chaves[k]}: lançou duas vezes?"
            else:
                chaves[k] = v["Nº original"]
            vendas.append(v)
        else:
            nao_vendas.append(v)

    wb = Workbook()

    # ===================== Vendas =====================
    ws = wb.active
    ws.title = "Vendas"
    cols = [c for c in COLUNAS]
    cabecalho(ws, 1, cols)
    for n, v in enumerate(vendas, start=2):
        escreve_linha(ws, n, v, cols, alt=(n % 2 == 0))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(len(vendas) + 1, 2)}"
    validacoes(ws, cols)

    # ===================== Não é venda =====================
    ws2 = wb.create_sheet("Não é venda")
    cols2 = ["Data", "Tipo", "Produto", "Cliente", "Valor (R$)", "Pagamento", "Obs", "Nota do sistema", "Nº original"]
    cabecalho(ws2, 1, cols2, fill=fill_cab2)
    for n, v in enumerate(nao_vendas, start=2):
        escreve_linha(ws2, n, v, cols2, alt=(n % 2 == 0))
    ws2.freeze_panes = "A2"
    ws2.cell(row=len(nao_vendas) + 3, column=1,
             value="Compras de estoque e devoluções ficam aqui pra não entrarem na soma das vendas.").font = f_nota

    # ===================== Resumo =====================
    ws3 = wb.create_sheet("Resumo", 0)
    ws3.sheet_view.showGridLines = False
    ws3["A1"] = "Vendas da Maria"
    ws3["A1"].font = f_titulo
    ws3["A2"] = "Os números abaixo se atualizam sozinhos quando você lança uma venda nova na aba Vendas."
    ws3["A2"].font = f_sub
    ws3.column_dimensions["A"].width = 34
    ws3.column_dimensions["B"].width = 18
    ws3.column_dimensions["C"].width = 3
    ws3.column_dimensions["D"].width = 26
    ws3.column_dimensions["E"].width = 16
    ws3.column_dimensions["F"].width = 18

    V = "Vendas!"
    itens = [
        ("Total vendido", f"=SUM({V}$D$2:$D$10000)", '"R$" #,##0.00'),
        ("Quantas vendas", f"=COUNTA({V}$B:$B)-1", "0"),
        ("Vendido hoje", f"=SUMIFS({V}$D:$D,{V}$A:$A,TODAY())", '"R$" #,##0.00'),
        ("Vendido neste mês", f"=SUMIFS({V}$D:$D,{V}$A:$A,\">=\"&DATE(YEAR(TODAY()),MONTH(TODAY()),1),{V}$A:$A,\"<=\"&EOMONTH(TODAY(),0))", '"R$" #,##0.00'),
        ("Vendido em setembro/2026", f"=SUMIFS({V}$D:$D,{V}$A:$A,\">=\"&DATE(2026,9,1),{V}$A:$A,\"<=\"&DATE(2026,9,30))", '"R$" #,##0.00'),
        ("Vendido em agosto/2026", f"=SUMIFS({V}$D:$D,{V}$A:$A,\">=\"&DATE(2026,8,1),{V}$A:$A,\"<=\"&DATE(2026,8,31))", '"R$" #,##0.00'),
        ("A receber (fiado)", f"=SUMIFS({V}$D:$D,{V}$K:$K,\"Não\")", '"R$" #,##0.00'),
        ("Encomendas com código de rastreio", f"=COUNTIF({V}$J$2:$J$10000,\"??*\")", "0"),
        ("Linhas pra você confirmar", f"=COUNTIF({V}$M:$M,\"SIM\")", "0"),
    ]
    ws3["A4"], ws3["B4"] = "O que você quer saber", "Resposta"
    for c in ("A4", "B4"):
        ws3[c].font, ws3[c].fill, ws3[c].alignment = f_cab, fill_cab, centro
    for i, (rotulo, formula, fmt) in enumerate(itens, start=5):
        ws3.cell(row=i, column=1, value=rotulo).font = f_normal
        c = ws3.cell(row=i, column=2, value=formula)
        c.font, c.number_format, c.alignment = f_destaque, fmt, centro
        if i % 2 == 0:
            ws3.cell(row=i, column=1).fill = fill_alt
            c.fill = fill_alt
    ws3.cell(row=5 + len(itens), column=1,
             value="O total exclui compras de estoque e devoluções (aba 'Não é venda').").font = f_nota

    # quem me deve (fórmula por cliente)
    devedores = sorted({v["Cliente"] for v in vendas if v["Recebido?"] == "Não"})
    ws3["D4"], ws3["E4"], ws3["F4"] = "Quem está me devendo", "Quanto", "Desde"
    for c in ("D4", "E4", "F4"):
        ws3[c].font, ws3[c].fill, ws3[c].alignment = f_cab, fill_cab2, centro
    for i, nome in enumerate(devedores, start=5):
        ws3.cell(row=i, column=4, value=nome).font = f_normal
        c = ws3.cell(row=i, column=5, value=f"=SUMIFS({V}$D:$D,{V}$C:$C,D{i},{V}$K:$K,\"Não\")")
        c.font, c.number_format, c.alignment = f_normal, '"R$" #,##0.00', centro
        c = ws3.cell(row=i, column=6, value=f"=IFERROR(_xlfn.MINIFS({V}$A:$A,{V}$C:$C,D{i},{V}$K:$K,\"Não\"),\"\")")
        c.font, c.number_format, c.alignment = f_normal, "dd/mm/yyyy", centro
    fim = 5 + len(devedores)
    ws3.cell(row=fim, column=4, value="Total").font = Font(name=FONTE, size=10, bold=True)
    c = ws3.cell(row=fim, column=5, value=f"=SUM(E5:E{fim - 1})" if devedores else 0)
    c.font, c.number_format, c.alignment = f_destaque, '"R$" #,##0.00', centro
    ws3.cell(row=fim + 1, column=4,
             value="Quando receber, troca 'Recebido?' pra Sim na aba Vendas e a pessoa some daqui.").font = f_nota

    # o que está no correio
    ini = fim + 3
    ws3.cell(row=ini, column=4, value="O que está no correio"), ws3.cell(row=ini, column=5, value="Código"), ws3.cell(row=ini, column=6, value="Enviado em")
    for j in (4, 5, 6):
        ws3.cell(row=ini, column=j).font, ws3.cell(row=ini, column=j).fill, ws3.cell(row=ini, column=j).alignment = f_cab, fill_cab2, centro
    env = [v for v in vendas if v["Rastreio"]]
    for i, v in enumerate(env, start=ini + 1):
        ws3.cell(row=i, column=4, value=f"{v['Produto']} — {v['Cliente']}").font = f_normal
        ws3.cell(row=i, column=5, value=v["Rastreio"]).font = f_normal
        c = ws3.cell(row=i, column=6, value=v["Data"])
        c.number_format, c.font, c.alignment = "dd/mm/yyyy", f_normal, centro
    ws3.cell(row=ini + len(env) + 1, column=4,
             value=f"Lista gerada na migração em {hoje.strftime('%d/%m/%Y')}; o assistente atualiza quando você lança envios novos.").font = f_nota

    # ===================== Original =====================
    ws4 = wb.create_sheet("Original")
    cabecalho(ws4, 1, ["Nº original"] + list(raw.columns), fill=PatternFill("solid", fgColor=CINZA))
    for i, r in raw.iterrows():
        ws4.cell(row=i + 2, column=1, value=i + 2).font = f_nota
        for j, col in enumerate(raw.columns, start=2):
            val = r[col]
            ws4.cell(row=i + 2, column=j, value=None if (isinstance(val, float) and val != val) else val).font = f_normal
    ws4.freeze_panes = "A2"
    ws4.cell(row=len(raw) + 3, column=1,
             value="Sua planilha do jeito que estava, sem mexer. 'Nº original' aparece na aba Vendas pra você conferir.").font = f_nota

    # ===================== Como usar =====================
    ws5 = wb.create_sheet("Como usar")
    ws5.sheet_view.showGridLines = False
    ws5.column_dimensions["A"].width = 110
    textos = [
        ("Como usar esta planilha", f_titulo),
        ("", f_normal),
        ("1. Pra registrar uma venda, use o assistente (o aplicativo de conversa): você só conta o que vendeu e ele anota aqui.", f_normal),
        ("2. Se preferir anotar direto, use a aba Vendas: vai escrevendo na próxima linha vazia. Pagamento, cartão, entrega e 'Recebido?' têm uma setinha com a lista pra escolher.", f_normal),
        ("3. A aba Resumo responde sozinha: quanto vendi hoje, no mês, quem me deve e o que está no correio.", f_normal),
        ("4. Linha amarela com 'Confirmar? = SIM' é uma venda antiga que o sistema não teve certeza. A coluna 'Nota do sistema' diz o que conferir. Depois de conferir, apaga o SIM.", f_normal),
        ("5. Compra de estoque e devolução ficam na aba 'Não é venda', pra não bagunçar a soma.", f_normal),
        ("6. A aba Original é a sua planilha de antes, intocada. Nada se perdeu.", f_normal),
        ("", f_normal),
        ("Cores: cabeçalho azul = colunas; linha amarela = conferir; aba Original em cinza = histórico.", f_nota),
    ]
    for i, (t, f) in enumerate(textos, start=1):
        c = ws5.cell(row=i, column=1, value=t)
        c.font = f
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws5.row_dimensions[i].height = 30 if len(t) > 90 else 18

    # aba de cores do Grupo Fácil nas abas
    for w, cor in ((ws3, AZUL_ESCURO), (ws, AZUL), (ws2, AZUL), (ws4, CINZA), (ws5, CINZA)):
        w.sheet_properties.tabColor = cor

    wb.save(saida)
    return len(vendas), len(nao_vendas), sum(1 for v in vendas if v["Confirmar?"] == "SIM")


if __name__ == "__main__":
    entrada = sys.argv[1] if len(sys.argv) > 1 else "Vendas da maria.xlsx"
    saida = sys.argv[2] if len(sys.argv) > 2 else "Vendas da Maria - organizada.xlsx"
    nv, nn, nc = migrar(entrada, saida)
    print(f"ok: {nv} vendas, {nn} não-vendas, {nc} pra confirmar -> {saida}")
