"""
Gera relatório de análise em PDF — Todas as famílias excepto café em grão.
Script independente de gerar_relatorio_pdf.py

Uso manual:
    python gerar_relatorio_pdf_outras.py
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
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ── Configuração ─────────────────────────────────────────────────────────────
PASTA_FONTE   = r"I:\O meu disco\Trovidoce\Mapas Mensais"
PASTA_DESTINO = r"I:\O meu disco\Claude"
NOME_SAIDA    = "Analise_Outras_Familias.pdf"

EXCLUIR = [
    "Buondi Grão Kg", "Buondi Decafe Grão Kg",
    "Sical Grão Kg",  "Sical Decafe Grão Kg",
    "Christina Grão Kg", "Christina Decafe Grão Kg",
]
DIAS_MES = 30.44

# Cores
AZUL_ESCURO       = colors.HexColor("#1F4E79")
AZUL_MEDIO        = colors.HexColor("#2E75B6")
AZUL_CLARO        = colors.HexColor("#D6E4F0")
AZUL_MUITO_CLARO  = colors.HexColor("#EBF3FB")
VERMELHO          = colors.HexColor("#C00000")
VERDE             = colors.HexColor("#375623")
BRANCO            = colors.white
# ─────────────────────────────────────────────────────────────────────────────


def encontrar_ficheiro_fonte():
    nome_esperado = "COLAR AQUI para Analise Familias.xlsx"
    caminho = os.path.join(PASTA_FONTE, nome_esperado)
    if os.path.exists(caminho):
        return caminho
    padrao = os.path.join(PASTA_FONTE, "*COLAR AQUI*.xlsx")
    ficheiros = glob.glob(padrao)
    if ficheiros:
        return max(ficheiros, key=os.path.getmtime)
    raise FileNotFoundError(
        f"Ficheiro '{nome_esperado}' não encontrado em:\n  {PASTA_FONTE}"
    )


def preparar_dados(ficheiro_fonte):
    df = pd.read_excel(ficheiro_fonte, sheet_name="COLAR AQUI", header=1)

    padrao_excluir = "|".join(EXCLUIR)
    df_f = df[~df["Familia Completa"].str.contains(padrao_excluir, na=False)].copy()
    df_f = df_f[df_f["Familia Completa"].notna()]

    df_f["Dt_ Emissao"] = pd.to_datetime(df_f["Dt_ Emissao"])
    df_f["Ano"] = df_f["Dt_ Emissao"].dt.year

    nome_atual = (
        df_f.sort_values("Dt_ Emissao")
        .groupby("N_ Clie_")["Cliente"].last()
        .reset_index().rename(columns={"Cliente": "Nome"})
    )
    df_f = df_f.merge(nome_atual, on="N_ Clie_")
    df_f["Cliente"] = df_f["Nome"]

    df_f["Familia_Det"] = df_f["Familia Completa"].apply(
        lambda x: x.split("\\")[-1] if "\\" in str(x) else str(x)
    )

    chave = ["N_ Vend_", "Vendedor", "N_ Clie_", "Cliente", "Familia Completa", "Familia_Det", "Ano"]
    qtd  = df_f.groupby(chave)["Quantidade"].sum().reset_index()
    span = df_f.groupby(chave)["Dt_ Emissao"].agg(
        primeira_compra="min", ultima_compra="max"
    ).reset_index()
    span["N_Meses"] = span.apply(
        lambda r: max((r["ultima_compra"] - r["primeira_compra"]).days / DIAS_MES, 1.0),
        axis=1
    )
    rel = qtd.merge(span[chave + ["N_Meses"]], on=chave)
    rel["Media"] = (rel["Quantidade"] / rel["N_Meses"]).round(1)

    return df_f, rel


def gerar(ficheiro_fonte=None):
    if ficheiro_fonte is None:
        ficheiro_fonte = encontrar_ficheiro_fonte()

    print(f"[{datetime.now():%H:%M:%S}] A ler: {os.path.basename(ficheiro_fonte)}")
    df_f, rel = preparar_dados(ficheiro_fonte)

    os.makedirs(PASTA_DESTINO, exist_ok=True)
    saida = os.path.join(PASTA_DESTINO, NOME_SAIDA)

    doc = SimpleDocTemplate(
        saida, pagesize=A4,
        leftMargin=1.8*cm, rightMargin=1.8*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )

    def estilo(nome, **kw):
        return ParagraphStyle(nome, **kw)

    s_titulo = estilo("titulo", fontSize=18, textColor=AZUL_ESCURO,
                      fontName="Helvetica-Bold", alignment=TA_CENTER, spaceAfter=4)
    s_sub    = estilo("sub", fontSize=10, textColor=AZUL_MEDIO,
                      fontName="Helvetica", alignment=TA_CENTER, spaceAfter=2)
    s_data   = estilo("data", fontSize=8, textColor=colors.grey,
                      fontName="Helvetica", alignment=TA_CENTER, spaceAfter=12)
    s_secao  = estilo("secao", fontSize=12, textColor=BRANCO,
                      fontName="Helvetica-Bold", alignment=TA_LEFT,
                      backColor=AZUL_ESCURO, leftIndent=6, spaceAfter=8, spaceBefore=14)
    s_corpo  = estilo("corpo", fontSize=9, fontName="Helvetica",
                      spaceAfter=4, leading=13)
    s_nota   = estilo("nota", fontSize=8, textColor=colors.grey,
                      fontName="Helvetica-Oblique")
    s_aval   = estilo("aval", fontName="Helvetica", fontSize=9,
                      leading=14, spaceAfter=8, leftIndent=4, rightIndent=4)

    s_cell      = ParagraphStyle("cell",      fontName="Helvetica",      fontSize=8, leading=10)
    s_cell_bold = ParagraphStyle("cell_bold", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=BRANCO)
    center      = TA_CENTER

    def wrap(texto, bold=False):
        return Paragraph(str(texto), s_cell_bold if bold else s_cell)

    def tabela(dados, col_widths, header_bg=AZUL_ESCURO, alt=True, alt_color=AZUL_MUITO_CLARO):
        dados_p = [[wrap(c, bold=(ri == 0)) for c in row] for ri, row in enumerate(dados)]
        t = Table(dados_p, colWidths=col_widths, repeatRows=1)
        cmd = [
            ("BACKGROUND", (0, 0), (-1, 0), header_bg),
            ("FONTSIZE",   (0, 0), (-1, -1), 8),
            ("ALIGN",      (0, 0), (-1, -1), "CENTER"),
            ("ALIGN",      (0, 1), (1, -1),  "LEFT"),
            ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
            ("GRID",       (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 4),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ]
        if alt:
            for i in range(2, len(dados), 2):
                cmd.append(("BACKGROUND", (0, i), (-1, i), alt_color))
        t.setStyle(TableStyle(cmd))
        return t

    story = []

    total_un   = rel["Quantidade"].sum()
    ano25      = rel[rel["Ano"] == 2025]["Quantidade"].sum()
    ano26      = rel[rel["Ano"] == 2026]["Quantidade"].sum()
    proj26     = ano26 * 2
    var        = (proj26 - ano25) / ano25 * 100 if ano25 > 0 else 0
    var_abs    = proj26 - ano25
    data_min   = df_f["Dt_ Emissao"].min().strftime("%d/%m/%Y")
    data_max   = df_f["Dt_ Emissao"].max().strftime("%d/%m/%Y")

    clientes_25 = set(rel[rel["Ano"] == 2025]["N_ Clie_"].unique())
    clientes_26 = set(rel[rel["Ano"] == 2026]["N_ Clie_"].unique())
    perdidos    = clientes_25 - clientes_26
    novos       = clientes_26 - clientes_25

    perdidos_df = (
        rel[rel["N_ Clie_"].isin(perdidos) & (rel["Ano"] == 2025)]
        .groupby(["Vendedor", "N_ Clie_", "Cliente"])["Quantidade"].sum()
        .sort_values(ascending=False).reset_index()
    )
    perdidos_rel = perdidos_df[perdidos_df["Quantidade"] >= 5]

    novos_df = (
        rel[rel["N_ Clie_"].isin(novos) & (rel["Ano"] == 2026)]
        .groupby(["Vendedor", "N_ Clie_", "Cliente"])["Quantidade"].sum()
        .sort_values(ascending=False).reset_index()
    )

    # ── Cabeçalho ──
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("TROVIDOCE", s_titulo))
    story.append(Paragraph("Análise de Vendas — Outras Famílias", s_sub))
    story.append(Paragraph(
        f"Período: {data_min} a {data_max}  |  Gerado em: {datetime.now():%d/%m/%Y %H:%M}",
        s_data
    ))
    story.append(HRFlowable(width="100%", thickness=1.5, color=AZUL_ESCURO, spaceAfter=12))

    # ── KPIs ──
    kpi_data = [
        ["Total Unidades", "Nº Clientes", "Nº Vendedores", "Nº Famílias"],
        [
            f"{total_un:,}".replace(",", "."),
            str(rel["N_ Clie_"].nunique()),
            str(rel[~rel["Vendedor"].isin(["Vendas Internas", "14/Armazém"])]["Vendedor"].nunique()),
            str(rel["Familia_Det"].nunique()),
        ]
    ]
    t_kpi = Table(kpi_data, colWidths=[4*cm]*4)
    t_kpi.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL_ESCURO),
        ("BACKGROUND", (0, 1), (-1, 1), AZUL_CLARO),
        ("TEXTCOLOR",  (0, 0), (-1, 0), BRANCO),
        ("TEXTCOLOR",  (0, 1), (-1, 1), AZUL_ESCURO),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME",   (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, 0), 8),
        ("FONTSIZE",   (0, 1), (-1, 1), 14),
        ("ALIGN",      (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("ROWHEIGHT",  (0, 0), (0, 0), 18),
        ("ROWHEIGHT",  (0, 1), (0, 1), 28),
        ("GRID",       (0, 0), (-1, -1), 0.5, BRANCO),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 0.4*cm))

    # ── 1. Evolução anual ──
    story.append(Paragraph("1. Evolução Anual", s_secao))
    var_cor = "red" if var < 0 else "green"
    story.append(Paragraph(
        f"Em 2025 foram vendidas <b>{ano25:,}</b> unidades. Nos primeiros 6 meses de 2026 foram vendidas "
        f"<b>{ano26:,}</b> unidades, o que projectado para o ano completo representa <b>{proj26:,}</b> unidades — "
        f"uma variação estimada de <font color='{var_cor}'><b>{var:+.1f}%</b></font> face a 2025.".replace(",", "."),
        s_corpo
    ))
    evo_data = [["Ano", "Unidades (real)", "Projecção anual (un.)", "Variação"]]
    evo_data.append(["2025", f"{ano25:,}".replace(",","."), "—", "—"])
    evo_data.append(["2026", f"{ano26:,}".replace(",","."),
                     f"{proj26:,}".replace(",","."), f"{var:+.1f}%"])
    t_evo = tabela(evo_data, [2.5*cm, 4*cm, 5*cm, 3.5*cm])
    t_evo.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL_ESCURO),
        ("FONTSIZE",   (0, 0), (-1, -1), 9),
        ("ALIGN",      (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",       (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
        ("BACKGROUND", (0, 2), (-1, 2), AZUL_MUITO_CLARO),
        ("TEXTCOLOR",  (3, 2), (3, 2), VERMELHO if var < 0 else VERDE),
        ("FONTNAME",   (3, 2), (3, 2), "Helvetica-Bold"),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_evo)
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("* Projecção 2026 calculada com base nos primeiros 6 meses (×2).", s_nota))

    # ── 2. Top famílias ──
    story.append(Paragraph("2. Top 20 Famílias — Período Total", s_secao))
    pf = rel.groupby("Familia_Det")["Quantidade"].sum().sort_values(ascending=False).head(20)
    fam_data = [["#", "Família", "Unidades", "% Total"]]
    for i, (f, q) in enumerate(pf.items(), 1):
        fam_data.append([str(i), f, f"{q:,}".replace(",","."), f"{q/total_un*100:.1f}%"])
    story.append(tabela(fam_data, [0.8*cm, 9.5*cm, 3*cm, 2.7*cm]))

    # ── 3. Por vendedor ──
    story.append(Paragraph("3. Desempenho por Vendedor", s_secao))
    pv = rel.groupby("Vendedor")["Quantidade"].sum().sort_values(ascending=False)
    vend_data = [["Vendedor", "Unidades", "% Total", "Nº Clientes"]]
    for v, q in pv.items():
        nc = rel[rel["Vendedor"] == v]["N_ Clie_"].nunique()
        vend_data.append([v, f"{q:,}".replace(",","."),
                          f"{q/total_un*100:.1f}%", str(nc)])
    story.append(tabela(vend_data, [6.5*cm, 3*cm, 3*cm, 3.5*cm]))

    # ── 4. Top 15 clientes ──
    story.append(Paragraph("4. Top 15 Clientes — Período Total", s_secao))
    pc = rel.groupby(["N_ Clie_", "Cliente", "Vendedor"])["Quantidade"].sum().sort_values(ascending=False).head(15)
    top_data = [["#", "Cliente", "Vendedor", "Unidades", "% Total"]]
    for i, ((cod, nome, vend), q) in enumerate(pc.items(), 1):
        v = vend.split("/")[-1] if "/" in vend else vend
        top_data.append([str(i), nome, v, f"{q:,}".replace(",","."), f"{q/total_un*100:.1f}%"])
    story.append(tabela(top_data, [0.8*cm, 7*cm, 3.5*cm, 2.5*cm, 2.2*cm]))

    # ── 5. Alertas ──
    story.append(Paragraph("5. Alertas — Clientes em Risco", s_secao))
    story.append(Paragraph(
        f"<b>{len(perdidos)}</b> clientes compraram em 2025 mas <b>não registam qualquer compra em 2026</b>. "
        f"Os {len(perdidos_rel)} com volume igual ou superior a 5 caixas em 2025 são:",
        s_corpo
    ))
    if not perdidos_rel.empty:
        perd_data = [["Vendedor", "Cliente", "Caixas 2025"]]
        for _, row in perdidos_rel.iterrows():
            v = row["Vendedor"].split("/")[-1] if "/" in row["Vendedor"] else row["Vendedor"]
            perd_data.append([v, row["Cliente"], f"{int(row['Quantidade']):,}".replace(",",".")])
        t_perd = tabela(perd_data, [4*cm, 9.5*cm, 2.5*cm], header_bg=VERMELHO,
                        alt_color=colors.HexColor("#FFF0F0"))
        story.append(t_perd)

    # ── 6. Clientes novos ──
    story.append(Paragraph("6. Clientes Novos em 2026", s_secao))
    story.append(Paragraph(
        f"<b>{len(novos)}</b> clientes iniciaram compras nestas famílias em 2026 "
        f"(sem historial em 2025), totalizando <b>{novos_df['Quantidade'].sum()}</b> caixas.",
        s_corpo
    ))
    if not novos_df.empty:
        nov_data = [["Vendedor", "Cliente", "Caixas 2026"]]
        for _, row in novos_df.iterrows():
            v = row["Vendedor"].split("/")[-1] if "/" in row["Vendedor"] else row["Vendedor"]
            nov_data.append([v, row["Cliente"], f"{int(row['Quantidade']):,}".replace(",",".")])
        story.append(tabela(nov_data, [4*cm, 9.5*cm, 2.5*cm], header_bg=VERDE,
                            alt_color=colors.HexColor("#F0FFF0")))

    # ── 7. Avaliação ──
    story.append(Paragraph("7. Avaliação de Performance", s_secao))

    un_perdidas = int(perdidos_df["Quantidade"].sum())
    un_novas    = int(novos_df["Quantidade"].sum())
    saldo_un    = un_novas - un_perdidas
    tx_retencao = (len(clientes_26) / len(clientes_25) * 100) if clientes_25 else 0
    saldo_cor   = "green" if saldo_un >= 0 else "red"
    var_cor2    = "green" if var >= 0 else "red"

    paragrafos = [
        f"<b>Tendência geral:</b> Com base nos primeiros 6 meses de 2026, a projecção anual aponta para "
        f"<b>{proj26:,} unidades</b>, face às <b>{ano25:,} unidades</b> registadas em 2025. "
        f"Isso representa uma variação estimada de "
        f"<font color='{var_cor2}'><b>{var:+.1f}%</b></font> "
        f"({'queda' if var < 0 else 'crescimento'} de {abs(int(var_abs)):,} unidades).".replace(",", "."),

        f"<b>Retenção de clientes:</b> Dos <b>{len(clientes_25)}</b> clientes activos em 2025, "
        f"<b>{len(clientes_25) - len(perdidos)}</b> mantiveram compras em 2026 "
        f"(taxa de retenção de <b>{tx_retencao:.1f}%</b>). "
        f"Os <b>{len(perdidos)}</b> clientes que não regressaram em 2026 representavam "
        f"<b>{un_perdidas:,} unidades</b> em 2025.".replace(",", "."),

        f"<b>Captação vs. perda:</b> Os <b>{len(novos)}</b> clientes novos captados em 2026 "
        f"trouxeram <b>{un_novas:,} unidades</b>. "
        f"O saldo líquido é <font color='{saldo_cor}'><b>{saldo_un:+,} unidades</b></font> — "
        f"{'os novos clientes <b>não compensam</b> o volume perdido' if saldo_un < 0 else 'os novos clientes <b>superam</b> o volume perdido'}.".replace(",", "."),

        f"<b>Nota:</b> Os dados de 2026 correspondem apenas a 6 meses (Janeiro–Junho). "
        f"Alguns dos {len(perdidos)} clientes classificados como \"perdidos\" podem retomar "
        f"compras no segundo semestre, pelo que esta análise deve ser revisitada no final do ano.",
    ]
    for p in paragrafos:
        story.append(Paragraph(p, s_aval))

    # ── Rodapé ──
    story.append(Spacer(1, 0.8*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    story.append(Paragraph(
        f"Relatório gerado automaticamente em {datetime.now():%d/%m/%Y às %H:%M} | Trovidoce",
        estilo("rodape", fontSize=7, textColor=colors.grey, alignment=TA_CENTER)
    ))

    doc.build(story)
    print(f"[{datetime.now():%H:%M:%S}] PDF guardado em:\n  {saida}")
    return saida


if __name__ == "__main__":
    ficheiro = sys.argv[1] if len(sys.argv) > 1 else None
    gerar(ficheiro)
