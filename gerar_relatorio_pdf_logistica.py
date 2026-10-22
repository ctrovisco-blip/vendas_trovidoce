"""
Secção de Custos Logísticos por Vendedor — Trovidoce.

Lê o mapa de vendas (para as vendas líquidas por vendedor) e o mapa
"Kms frota.xlsx" (kms por vendedor e por ano) da mesma pasta, e produz
um PDF com a estimativa de custo de combustível por vendedor.

Pressupostos (editáveis abaixo):
  - Consumo médio: 12 L / 100 km
  - Preço do gasóleo por ano (€/litro)
  - Os kms do carro de distribuição (nº 14) são somados ao vendedor 31.
"""

import os
import sys
import glob
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ── Configuração ─────────────────────────────────────────────────────────────
PASTA_FONTE   = r"I:\O meu disco\Trovidoce\Mapas Mensais"
PASTA_DESTINO = r"I:\O meu disco\Claude"
NOME_SAIDA    = "Analise_Logistica.pdf"

# Pressupostos de custo
CONSUMO_L_100  = 12.0            # litros por 100 km
PRECO_GASOLEO  = {               # €/litro, média anual (ajustar conforme necessário)
    2025: 1.55,
    2026: 2.05,
}
PRECO_GASOLEO_DEFAULT = 1.80     # usado se o ano não constar do dicionário

# Carro de distribuição -> vendedor a quem os kms são imputados
CARRO_DISTRIBUICAO = 14
VENDEDOR_DISTRIBUICAO = 31

# Cores
AZUL_ESCURO      = colors.HexColor("#1F4E79")
AZUL_MEDIO       = colors.HexColor("#2E75B6")
AZUL_CLARO       = colors.HexColor("#D6E4F0")
AZUL_MUITO_CLARO = colors.HexColor("#EBF3FB")
VERMELHO         = colors.HexColor("#C00000")
VERDE            = colors.HexColor("#375623")
BRANCO           = colors.white
# ─────────────────────────────────────────────────────────────────────────────


def _encontrar(nome_contains, nome_exacto=None):
    if nome_exacto:
        p = os.path.join(PASTA_FONTE, nome_exacto)
        if os.path.exists(p):
            return p
    padrao = os.path.join(PASTA_FONTE, f"*{nome_contains}*.xlsx")
    fich = glob.glob(padrao)
    if fich:
        return max(fich, key=os.path.getmtime)
    raise FileNotFoundError(f"Ficheiro com '{nome_contains}' não encontrado em:\n  {PASTA_FONTE}")


def encontrar_ficheiro_vendas():
    return _encontrar("COLAR AQUI", "COLAR AQUI para Analise Familias.xlsx")


def encontrar_ficheiro_kms():
    return _encontrar("Kms", "Kms frota.xlsx")


def ler_kms(ficheiro_kms):
    """Devolve {ano: {num_vendedor: kms}} com o carro 14 somado ao vendedor 31."""
    xls = pd.ExcelFile(ficheiro_kms)
    out = {}
    for sh in xls.sheet_names:
        try:
            ano = int(str(sh).strip())
        except ValueError:
            continue
        # cabeçalho dos vendedores está na 4ª linha (índice 3)
        d = pd.read_excel(ficheiro_kms, sheet_name=sh, skiprows=3)
        d = d.rename(columns={d.columns[0]: "Mes"})
        d = d[d["Mes"].notna() & (d["Mes"].astype(str).str.strip().str.lower() != "totais")]
        kms = {}
        for col in d.columns[1:]:
            try:
                vend = int(col)
            except (ValueError, TypeError):
                continue
            kms[vend] = float(pd.to_numeric(d[col], errors="coerce").sum())
        # imputar o carro de distribuição ao vendedor respectivo
        if CARRO_DISTRIBUICAO in kms:
            kms[VENDEDOR_DISTRIBUICAO] = kms.get(VENDEDOR_DISTRIBUICAO, 0) + kms.pop(CARRO_DISTRIBUICAO)
        out[ano] = kms
    return out


def ler_vendas(ficheiro_vendas):
    """Devolve (mapa_nome, {ano: {num_vendedor: vendas_liquidas}})."""
    df = pd.read_excel(ficheiro_vendas, sheet_name="COLAR AQUI", header=1)
    df["Dt_ Emissao"] = pd.to_datetime(df["Dt_ Emissao"])
    df["Ano"] = df["Dt_ Emissao"].dt.year
    mapa_nome = dict(df[["N_ Vend_", "Vendedor"]].drop_duplicates().values)
    vendas = {}
    g = df.groupby(["Ano", "N_ Vend_"])["Total Liquido"].sum()
    for (ano, vend), val in g.items():
        vendas.setdefault(int(ano), {})[int(vend)] = float(val)
    return mapa_nome, vendas


def preco_gasoleo(ano):
    return PRECO_GASOLEO.get(ano, PRECO_GASOLEO_DEFAULT)


# ── Construção da secção (para uso isolado ou integrado) ─────────────────────
def build_logistica(story, H, kms_por_ano, vendas_por_ano, mapa_nome):
    s_secao = H["s_secao"]; s_corpo = H["s_corpo"]; s_nota = H["s_nota"]
    tabela = H["tabela"]

    def fmt(n):   return f"{int(round(n)):,}".replace(",", ".")
    def fmt_eur(n): return f"{int(round(n)):,}".replace(",", ".") + " €"

    anos = sorted(set(kms_por_ano) | set(vendas_por_ano))
    # vendedores com kms (exclui o carro 14, já imputado)
    vendedores = sorted({v for a in kms_por_ano.values() for v in a})

    precos_txt = "; ".join(f"{a}: {preco_gasoleo(a):.2f} €/l" for a in anos)
    intro = (
        f"Estimativa do custo de combustível por vendedor, a partir dos quilómetros percorridos. "
        f"Pressupostos: consumo médio de <b>{CONSUMO_L_100:.0f} L/100 km</b> e preço médio do gasóleo "
        f"(<b>{precos_txt}</b>). Os quilómetros do carro de distribuição (nº {CARRO_DISTRIBUICAO}) "
        f"são imputados ao vendedor {VENDEDOR_DISTRIBUICAO}, por ser a rota em que a venda e a entrega "
        f"ocorrem em dias consecutivos."
    )

    blocos = [Paragraph("Custos Logísticos por Vendedor", s_secao), Paragraph(intro, s_corpo)]

    def nome_vend(num):
        nome = mapa_nome.get(num, str(num))
        if num == VENDEDOR_DISTRIBUICAO:
            nome = f"{nome} (c/ carro {CARRO_DISTRIBUICAO})"
        return nome

    totais_custo = {v: 0.0 for v in vendedores}

    for ano in anos:
        kms_ano = kms_por_ano.get(ano, {})
        vendas_ano = vendas_por_ano.get(ano, {})
        preco = preco_gasoleo(ano)

        linhas = [["Vendedor", "Kms", "Litros", f"Custo Gasóleo", "Vendas Líq.", "Custo/Vendas"]]
        tot_km = tot_lit = tot_cst = tot_vnd = 0.0
        for v in vendedores:
            km = kms_ano.get(v, 0.0)
            if km == 0:
                continue
            litros = km * CONSUMO_L_100 / 100.0
            custo  = litros * preco
            vnd    = vendas_ano.get(v, 0.0)
            ratio  = (custo / vnd * 100) if vnd else 0
            totais_custo[v] += custo
            tot_km += km; tot_lit += litros; tot_cst += custo; tot_vnd += vnd
            linhas.append([
                nome_vend(v), fmt(km), fmt(litros), fmt_eur(custo),
                fmt_eur(vnd), f"{ratio:.2f}%",
            ])
        ratio_tot = (tot_cst / tot_vnd * 100) if tot_vnd else 0
        linhas.append([
            "TOTAL", fmt(tot_km), fmt(tot_lit), fmt_eur(tot_cst),
            fmt_eur(tot_vnd), f"{ratio_tot:.2f}%",
        ])
        t = tabela(linhas, [5*cm, 2.2*cm, 2.2*cm, 2.6*cm, 2.6*cm, 2.4*cm])
        # negrito na linha de total
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, len(linhas)-1), (-1, len(linhas)-1), AZUL_CLARO),
            ("FONTNAME",   (0, len(linhas)-1), (-1, len(linhas)-1), "Helvetica-Bold"),
        ]))
        blocos.append(Spacer(1, 0.25*cm))
        blocos.append(Paragraph(f"<b>{ano}</b> — preço do gasóleo: {preco:.2f} €/l", s_corpo))
        blocos.append(t)

    story.append(KeepTogether(blocos[:3]))   # título + intro + 1ª tabela juntos
    for b in blocos[3:]:
        story.append(b)

    # Nota final
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Nota: o custo apresentado corresponde apenas a combustível (não inclui portagens, "
        "manutenção, amortização da viatura nem custos salariais). A coluna <b>Custo/Vendas</b> "
        "mede o peso do combustível no volume de vendas líquidas de cada vendedor.",
        s_nota
    ))


def _helpers(doc):
    def estilo(nome, **kw):
        return ParagraphStyle(nome, **kw)
    s_secao = estilo("secao", fontSize=12, textColor=BRANCO, fontName="Helvetica-Bold",
                     alignment=TA_LEFT, backColor=AZUL_ESCURO, leftIndent=6,
                     spaceAfter=8, spaceBefore=14)
    s_corpo = estilo("corpo", fontSize=9, fontName="Helvetica", spaceAfter=4, leading=13)
    s_nota  = estilo("nota", fontSize=8, textColor=colors.grey, fontName="Helvetica-Oblique")
    s_cell      = ParagraphStyle("cell", fontName="Helvetica", fontSize=8, leading=10)
    s_cell_bold = ParagraphStyle("cellb", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=BRANCO)

    def wrap(t, bold=False):
        return Paragraph(str(t), s_cell_bold if bold else s_cell)

    def tabela(dados, col_widths, header_bg=AZUL_ESCURO, alt=True):
        dados_p = [[wrap(c, bold=(ri == 0)) for c in row] for ri, row in enumerate(dados)]
        t = Table(dados_p, colWidths=col_widths, repeatRows=1)
        cmd = [
            ("BACKGROUND", (0, 0), (-1, 0), header_bg),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("ALIGN", (0, 1), (0, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
        if alt:
            for i in range(2, len(dados), 2):
                cmd.append(("BACKGROUND", (0, i), (-1, i), AZUL_MUITO_CLARO))
        t.setStyle(TableStyle(cmd))
        return t

    return {"s_secao": s_secao, "s_corpo": s_corpo, "s_nota": s_nota, "tabela": tabela}


def gerar(ficheiro_vendas=None, ficheiro_kms=None, saida=None):
    if ficheiro_vendas is None:
        ficheiro_vendas = encontrar_ficheiro_vendas()
    if ficheiro_kms is None:
        ficheiro_kms = encontrar_ficheiro_kms()

    print(f"[{datetime.now():%H:%M:%S}] Logística — vendas: {os.path.basename(ficheiro_vendas)} | kms: {os.path.basename(ficheiro_kms)}")
    kms_por_ano = ler_kms(ficheiro_kms)
    mapa_nome, vendas_por_ano = ler_vendas(ficheiro_vendas)

    os.makedirs(PASTA_DESTINO, exist_ok=True)
    if saida is None:
        saida = os.path.join(PASTA_DESTINO, NOME_SAIDA)

    doc = SimpleDocTemplate(saida, pagesize=A4, leftMargin=1.8*cm, rightMargin=1.8*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    H = _helpers(doc)
    story = [Spacer(1, 0.3*cm)]
    build_logistica(story, H, kms_por_ano, vendas_por_ano, mapa_nome)
    doc.build(story)
    print(f"[{datetime.now():%H:%M:%S}] PDF logística guardado em:\n  {saida}")
    return saida


if __name__ == "__main__":
    gerar()
