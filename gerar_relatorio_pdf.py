"""
Gera relatório de análise em PDF — Café em Grão (Buondi / Christina / Sical)
"""

import os
import sys
import glob
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# ── Configuração ─────────────────────────────────────────────────────────────
PASTA_FONTE   = r"I:\O meu disco\Trovidoce\Mapas Mensais"
PASTA_DESTINO = r"I:\O meu disco\Claude"
NOME_SAIDA    = "Analise_Grao_Cafe.pdf"
FAMILIAS      = ["Buondi Grão Kg", "Christina Grão Kg", "Sical Grão Kg"]
KG_POR_CAIXA  = 6

# Cores
AZUL_ESCURO  = colors.HexColor("#1F4E79")
AZUL_MEDIO   = colors.HexColor("#2E75B6")
AZUL_CLARO   = colors.HexColor("#D6E4F0")
AZUL_MUITO_CLARO = colors.HexColor("#EBF3FB")
VERMELHO     = colors.HexColor("#C00000")
VERDE        = colors.HexColor("#375623")
CINZA        = colors.HexColor("#F2F2F2")
BRANCO       = colors.white
# ─────────────────────────────────────────────────────────────────────────────


def encontrar_ficheiro_fonte():
    """Procura o ficheiro de dados na pasta fonte pelo nome esperado."""
    nome_esperado = "COLAR AQUI para Analise Familias.xlsx"
    caminho = os.path.join(PASTA_FONTE, nome_esperado)
    if os.path.exists(caminho):
        return caminho
    # fallback: qualquer ficheiro com "COLAR AQUI" no nome
    padrao = os.path.join(PASTA_FONTE, "*COLAR AQUI*.xlsx")
    ficheiros = glob.glob(padrao)
    if ficheiros:
        return max(ficheiros, key=os.path.getmtime)
    raise FileNotFoundError(
        f"Ficheiro '{nome_esperado}' não encontrado em:\n  {PASTA_FONTE}\n"
        f"Certifica-te que o ficheiro está na pasta correcta."
    )


def preparar_dados(ficheiro_fonte):
    df = pd.read_excel(ficheiro_fonte, sheet_name="COLAR AQUI", header=1)
    df_f = df[df["Familia"].isin(FAMILIAS)].copy()
    df_f["Dt_ Emissao"] = pd.to_datetime(df_f["Dt_ Emissao"])
    df_f["Ano"] = df_f["Dt_ Emissao"].dt.year
    df_f["Mes"] = df_f["Dt_ Emissao"].dt.to_period("M")

    nome_atual = (
        df_f.sort_values("Dt_ Emissao")
        .groupby("N_ Clie_")["Cliente"].last()
        .reset_index().rename(columns={"Cliente": "Nome"})
    )
    df_f = df_f.merge(nome_atual, on="N_ Clie_")
    df_f["Cliente"] = df_f["Nome"]

    DIAS_MES = 30.44
    chave = ["N_ Vend_", "Vendedor", "N_ Clie_", "Cliente", "Familia", "Ano"]
    qtd = df_f.groupby(chave)["Quantidade"].sum().reset_index()
    span = df_f.groupby(chave)["Dt_ Emissao"].agg(primeira_compra="min", ultima_compra="max").reset_index()
    span["N_Meses"] = span.apply(
        lambda r: max(((r["ultima_compra"] - r["primeira_compra"]).days / DIAS_MES), 1/DIAS_MES),
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

    estilos = getSampleStyleSheet()
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
    s_nota   = estilo("nota", fontSize=8, textColor=colors.grey,
                      fontName="Helvetica-Oblique", alignment=TA_LEFT)
    s_alerta = estilo("alerta", fontSize=9, textColor=VERMELHO,
                      fontName="Helvetica-Bold")
    s_corpo  = estilo("corpo", fontSize=9, fontName="Helvetica",
                      spaceAfter=4, leading=13)

    s_cell      = ParagraphStyle("cell",      fontName="Helvetica",      fontSize=8, leading=10)
    s_cell_bold = ParagraphStyle("cell_bold", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=BRANCO)

    def wrap(texto, bold=False):
        """Converte texto em Paragraph para permitir quebra de linha na célula."""
        return Paragraph(str(texto), s_cell_bold if bold else s_cell)

    def tabela(dados, col_widths, header_bg=AZUL_ESCURO, alt=True):
        # Converter header em Paragraphs bold brancos, restantes em Paragraphs normais
        dados_p = []
        for ri, row in enumerate(dados):
            dados_p.append([wrap(c, bold=(ri==0)) for c in row])
        t = Table(dados_p, colWidths=col_widths, repeatRows=1)
        cmd = [
            ("BACKGROUND", (0,0), (-1,0), header_bg),
            ("TEXTCOLOR",  (0,0), (-1,0), BRANCO),
            ("FONTSIZE",   (0,0), (-1,-1), 8),
            ("ALIGN",      (0,0), (-1,-1), "CENTER"),
            ("ALIGN",      (0,1), (1,-1), "LEFT"),
            ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
            ("GRID",       (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ("LEFTPADDING",   (0,0), (-1,-1), 4),
            ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ]
        if alt:
            for i in range(2, len(dados), 2):
                cmd.append(("BACKGROUND", (0,i), (-1,i), AZUL_MUITO_CLARO))
        t.setStyle(TableStyle(cmd))
        return t

    story = []
    total_cx = rel["Quantidade"].sum()
    total_kg = total_cx * KG_POR_CAIXA
    ano25 = rel[rel["Ano"]==2025]["Quantidade"].sum()
    ano26 = rel[rel["Ano"]==2026]["Quantidade"].sum()
    proj26 = ano26 * 2
    var = (proj26 - ano25) / ano25 * 100
    data_min = df_f["Dt_ Emissao"].min().strftime("%d/%m/%Y")
    data_max = df_f["Dt_ Emissao"].max().strftime("%d/%m/%Y")

    # ── Cabeçalho ──
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("TROVIDOCE", s_titulo))
    story.append(Paragraph("Análise de Vendas — Café em Grão", s_sub))
    story.append(Paragraph(f"Período: {data_min} a {data_max}  |  Gerado em: {datetime.now():%d/%m/%Y %H:%M}", s_data))
    story.append(HRFlowable(width="100%", thickness=1.5, color=AZUL_ESCURO, spaceAfter=12))

    # ── KPIs resumo ──
    kpi_data = [
        ["Total Caixas", "Total Kg", "Nº Clientes", "Nº Vendedores", "Famílias"],
        [
            f"{total_cx:,}".replace(",","."),
            f"{total_kg:,}".replace(",","."),
            str(rel["N_ Clie_"].nunique()),
            str(rel[rel["Vendedor"].str.contains("Vend|Arm", na=False)==False]["Vendedor"].nunique()),
            str(len(FAMILIAS))
        ]
    ]
    t_kpi = Table(kpi_data, colWidths=[3.2*cm]*5)
    t_kpi.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), AZUL_ESCURO),
        ("BACKGROUND", (0,1), (-1,1), AZUL_CLARO),
        ("TEXTCOLOR",  (0,0), (-1,0), BRANCO),
        ("TEXTCOLOR",  (0,1), (-1,1), AZUL_ESCURO),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME",   (0,1), (-1,1), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,0), 8),
        ("FONTSIZE",   (0,1), (-1,1), 14),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("ROWHEIGHT",  (0,0), (0,0), 18),
        ("ROWHEIGHT",  (0,1), (0,1), 28),
        ("GRID",       (0,0), (-1,-1), 0.5, BRANCO),
        ("ROUNDEDCORNERS", [4]),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 0.4*cm))

    # ── Evolução anual ──
    story.append(Paragraph("1. Evolução Anual", s_secao))
    cor_var = VERMELHO if var < 0 else VERDE
    story.append(Paragraph(
        f"Em 2025 foram vendidas <b>{ano25:,}</b> caixas. Nos primeiros 6 meses de 2026 foram vendidas "
        f"<b>{ano26:,}</b> caixas, o que projectado para o ano completo representa <b>{proj26:,}</b> caixas — "
        f"uma variação estimada de <font color='{'red' if var<0 else 'green'}'><b>{var:+.1f}%</b></font> face a 2025.".replace(",","."),
        s_corpo
    ))

    evo_data = [["Ano", "Caixas (real)", "Kg (real)", "Projecção anual (cx)", "Variação"]]
    evo_data.append(["2025", f"{ano25:,}".replace(",","."), f"{ano25*6:,}".replace(",","."), "—", "—"])
    evo_data.append(["2026", f"{ano26:,}".replace(",","."), f"{ano26*6:,}".replace(",","."),
                     f"{proj26:,}".replace(",","."), f"{var:+.1f}%"])
    t_evo = tabela(evo_data, [2*cm, 3*cm, 3*cm, 4*cm, 3*cm])
    t_evo.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), AZUL_ESCURO),
        ("TEXTCOLOR",  (0,0), (-1,0), BRANCO),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME",   (0,1), (-1,-1), "Helvetica"),
        ("FONTSIZE",   (0,0), (-1,-1), 9),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("ROWHEIGHT",  (0,0), (-1,-1), 18),
        ("GRID",       (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
        ("BACKGROUND", (0,2), (-1,2), AZUL_MUITO_CLARO),
        ("TEXTCOLOR",  (4,2), (4,2), VERMELHO if var < 0 else VERDE),
        ("FONTNAME",   (4,2), (4,2), "Helvetica-Bold"),
    ]))
    story.append(t_evo)
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("* Projecção 2026 calculada com base nos primeiros 6 meses (×2).", s_nota))

    # ── Por família ──
    story.append(Paragraph("2. Distribuição por Família", s_secao))
    pf = rel.groupby("Familia")["Quantidade"].sum().sort_values(ascending=False)
    fam_data = [["Família", "Caixas", "Kg", "% Total"]]
    for f, q in pf.items():
        fam_data.append([f, f"{q:,}".replace(",","."), f"{q*6:,}".replace(",","."), f"{q/total_cx*100:.1f}%"])
    story.append(tabela(fam_data, [7*cm, 3*cm, 3*cm, 3*cm]))

    # ── Por vendedor ──
    story.append(Paragraph("3. Desempenho por Vendedor", s_secao))
    pv = rel.groupby("Vendedor")["Quantidade"].sum().sort_values(ascending=False)
    vend_data = [["Vendedor", "Caixas", "Kg", "% Total", "Nº Clientes"]]
    for v, q in pv.items():
        nc = rel[rel["Vendedor"]==v]["N_ Clie_"].nunique()
        vend_data.append([v, f"{q:,}".replace(",","."), f"{q*6:,}".replace(",","."),
                          f"{q/total_cx*100:.1f}%", str(nc)])
    story.append(tabela(vend_data, [5.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm]))

    # ── Top 15 clientes ──
    story.append(Paragraph("4. Top 15 Clientes — Período Total", s_secao))
    pc = rel.groupby(["N_ Clie_","Cliente","Vendedor"])["Quantidade"].sum().sort_values(ascending=False).head(15)
    top_data = [["#", "Cliente", "Vendedor", "Caixas", "Kg", "% Total"]]
    for i, ((cod,nome,vend), q) in enumerate(pc.items(), 1):
        top_data.append([str(i), nome, vend.split("/")[-1] if "/" in vend else vend,
                         f"{q:,}".replace(",","."), f"{q*6:,}".replace(",","."),
                         f"{q/total_cx*100:.1f}%"])
    story.append(tabela(top_data, [0.8*cm, 6.5*cm, 3*cm, 2*cm, 2*cm, 1.7*cm]))

    # ── Alertas ──
    story.append(Paragraph("5. Alertas — Clientes em Risco", s_secao))

    clientes_25 = set(rel[rel["Ano"]==2025]["N_ Clie_"].unique())
    clientes_26 = set(rel[rel["Ano"]==2026]["N_ Clie_"].unique())
    perdidos = clientes_25 - clientes_26
    perdidos_df = (
        rel[rel["N_ Clie_"].isin(perdidos) & (rel["Ano"]==2025)]
        .groupby(["Vendedor","N_ Clie_","Cliente"])["Quantidade"].sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    # filtrar só os com volume relevante (>= 5 cx)
    perdidos_rel = perdidos_df[perdidos_df["Quantidade"] >= 5]

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
        t_perd = tabela(perd_data, [4*cm, 9.5*cm, 2.5*cm], header_bg=VERMELHO)
        # sobrepor cor alternada vermelha clara
        extra = [("BACKGROUND", (0,i), (-1,i), colors.HexColor("#FFF0F0")) for i in range(2, len(perd_data), 2)]
        t_perd.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), VERMELHO),
            ("GRID",       (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
            ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            *extra,
        ]))
        story.append(t_perd)

    # ── Clientes novos ──
    story.append(Spacer(1, 0.5*cm))
    novos = clientes_26 - clientes_25
    novos_df = (
        rel[rel["N_ Clie_"].isin(novos) & (rel["Ano"]==2026)]
        .groupby(["Vendedor","N_ Clie_","Cliente"])["Quantidade"].sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    story.append(Paragraph("6. Clientes Novos em 2026", s_secao))
    story.append(Paragraph(
        f"<b>{len(novos)}</b> clientes iniciaram compras nestas famílias em 2026 (sem historial em 2025), "
        f"totalizando <b>{novos_df['Quantidade'].sum()}</b> caixas.",
        s_corpo
    ))
    if not novos_df.empty:
        nov_data = [["Vendedor", "Cliente", "Caixas 2026"]]
        for _, row in novos_df.iterrows():
            v = row["Vendedor"].split("/")[-1] if "/" in row["Vendedor"] else row["Vendedor"]
            nov_data.append([v, row["Cliente"], f"{int(row['Quantidade']):,}".replace(",",".")])
        t_nov = tabela(nov_data, [4*cm, 9.5*cm, 2.5*cm], header_bg=VERDE)
        extra = [("BACKGROUND", (0,i), (-1,i), colors.HexColor("#F0FFF0")) for i in range(2, len(nov_data), 2)]
        t_nov.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), VERDE),
            ("GRID",       (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
            ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            *extra,
        ]))
        story.append(t_nov)

    # ── Avaliação de performance ──
    story.append(Paragraph("7. Avaliação de Performance", s_secao))

    # Cálculos
    kg_perdidos    = int(perdidos_df["Quantidade"].sum() * KG_POR_CAIXA)
    cx_perdidos    = int(perdidos_df["Quantidade"].sum())
    kg_novos       = int(novos_df["Quantidade"].sum() * KG_POR_CAIXA)
    cx_novos       = int(novos_df["Quantidade"].sum())
    saldo_kg       = kg_novos - kg_perdidos
    saldo_cx       = cx_novos - cx_perdidos
    taxa_retencao  = (len(clientes_26) / len(clientes_25) * 100) if clientes_25 else 0
    taxa_churn     = (len(perdidos) / len(clientes_25) * 100) if clientes_25 else 0

    # Variação geral 2025 vs 2026 anualizado
    var_abs = proj26 - ano25

    # Texto de avaliação
    saldo_cor   = "green" if saldo_kg >= 0 else "red"
    saldo_sinal = "positivo" if saldo_kg >= 0 else "negativo"
    var_cor     = "green" if var >= 0 else "red"

    paragrafos = []

    # § 1 — Tendência geral
    paragrafos.append(
        f"<b>Tendência geral:</b> Com base nos primeiros 6 meses de 2026, a projecção anual aponta para "
        f"<b>{proj26:,} caixas ({proj26*KG_POR_CAIXA:,} kg)</b>, face às <b>{ano25:,} caixas ({ano25*KG_POR_CAIXA:,} kg)</b> "
        f"registadas em 2025. Isso representa uma variação estimada de "
        f"<font color='{var_cor}'><b>{var:+.1f}%</b></font> "
        f"({'queda' if var < 0 else 'crescimento'} de {abs(int(var_abs)):,} caixas / {abs(int(var_abs))*KG_POR_CAIXA:,} kg).".replace(",",".")
    )

    # § 2 — Retenção de clientes
    paragrafos.append(
        f"<b>Retenção de clientes:</b> Dos <b>{len(clientes_25)}</b> clientes activos em 2025, "
        f"<b>{len(clientes_25) - len(perdidos)}</b> mantiveram compras em 2026 "
        f"(taxa de retenção de <b>{taxa_retencao:.1f}%</b>). "
        f"Os <b>{len(perdidos)}</b> clientes que não regressaram em 2026 representavam "
        f"<b>{cx_perdidos} caixas / {kg_perdidos:,} kg</b> em 2025.".replace(",",".")
    )

    # § 3 — Captação vs. perda
    paragrafos.append(
        f"<b>Captação vs. perda:</b> Os <b>{len(novos)}</b> clientes novos captados em 2026 "
        f"trouxeram <b>{cx_novos} caixas / {kg_novos:,} kg</b>. "
        f"O saldo líquido entre captação e perda é "
        f"<font color='{saldo_cor}'><b>{saldo_cx:+d} caixas / {saldo_kg:+,} kg</b></font> — "
        f"{'os novos clientes <b>não compensam</b> o volume perdido pelos que saíram' if saldo_kg < 0 else 'os novos clientes <b>superam</b> o volume perdido pelos que saíram'}.".replace(",",".")
    )

    # § 4 — Nota sobre 2026 (apenas meio ano)
    paragrafos.append(
        f"<b>Nota:</b> Os dados de 2026 correspondem apenas a 6 meses (Janeiro–Junho). "
        f"É possível que alguns dos {len(perdidos)} clientes classificados como \"perdidos\" retomem "
        f"compras no segundo semestre, pelo que esta análise deve ser revisitada no final do ano."
    )

    s_aval = ParagraphStyle("aval", fontName="Helvetica", fontSize=9, leading=14,
                             spaceAfter=8, leftIndent=4, rightIndent=4)
    s_aval_primeiro = ParagraphStyle("aval0", parent=s_aval, spaceBefore=4)

    for i, p in enumerate(paragrafos):
        story.append(Paragraph(p.replace(",","."), s_aval_primeiro if i == 0 else s_aval))


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
