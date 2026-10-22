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
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# ── Configuração ─────────────────────────────────────────────────────────────
PASTA_FONTE   = r"I:\O meu disco\Trovidoce\Mapas Mensais"
PASTA_DESTINO = r"I:\O meu disco\Claude"
NOME_SAIDA    = "Analise_Grao_Cafe.pdf"

# Famílias incluídas e respectivo factor de conversão caixa -> kg
FATORES_KG = {
    "Buondi Grão Kg":          6.0,
    "Christina Grão Kg":       6.0,
    "Sical Grão Kg":           6.0,
    "Buondi Decafe Grão Kg":   1.0,
    "Sical Decafe Grão Kg":    1.0,
    "Christina Decafe Grão Kg":1.0,
    "Buondi Decafe Doses":     0.325,
    "Sical Decafe Doses":      0.325,
    "Christina Decafe Doses":  0.325,
    "Buondi Office":           1.0,
    "Café 200g":               2.4,
}
FAMILIAS = list(FATORES_KG)

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

    # Conversão caixa -> kg específica de cada família
    df_f["Kg"] = df_f["Quantidade"] * df_f["Familia"].map(FATORES_KG)

    nome_atual = (
        df_f.sort_values("Dt_ Emissao")
        .groupby("N_ Clie_")["Cliente"].last()
        .reset_index().rename(columns={"Cliente": "Nome"})
    )
    df_f = df_f.merge(nome_atual, on="N_ Clie_")
    df_f["Cliente"] = df_f["Nome"]

    DIAS_MES = 30.44
    chave = ["N_ Vend_", "Vendedor", "N_ Clie_", "Cliente", "Familia", "Ano"]
    qtd = df_f.groupby(chave)[["Quantidade", "Kg"]].sum().reset_index()
    span = df_f.groupby(chave)["Dt_ Emissao"].agg(primeira_compra="min", ultima_compra="max").reset_index()
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
    total_kg = rel["Kg"].sum()
    ano25 = rel[rel["Ano"]==2025]["Quantidade"].sum()
    ano26 = rel[rel["Ano"]==2026]["Quantidade"].sum()
    ano25_kg = rel[rel["Ano"]==2025]["Kg"].sum()
    ano26_kg = rel[rel["Ano"]==2026]["Kg"].sum()

    # Projecção 2026: anualização pelo run-rate (dias efectivamente decorridos até à última venda)
    dados26 = df_f[df_f["Ano"] == 2026]
    if len(dados26):
        ultima_dt_26 = dados26["Dt_ Emissao"].max()
        dias_decorridos_26 = (ultima_dt_26 - pd.Timestamp(year=2026, month=1, day=1)).days + 1
        fator_anual = 365.0 / dias_decorridos_26 if dias_decorridos_26 > 0 else 0
        ultima_26_txt = ultima_dt_26.strftime("%d/%m/%Y")
    else:
        dias_decorridos_26 = 0
        fator_anual = 0
        ultima_26_txt = "—"
    meses_decorridos_26 = dias_decorridos_26 / 30.44
    proj26 = int(round(ano26 * fator_anual))
    proj26_kg = int(round(ano26_kg * fator_anual))
    var = (proj26 - ano25) / ano25 * 100 if ano25 else 0
    data_min = df_f["Dt_ Emissao"].min().strftime("%d/%m/%Y")
    data_max = df_f["Dt_ Emissao"].max().strftime("%d/%m/%Y")

    # ── Cabeçalho ──
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("TROVIDOCE", s_titulo))
    story.append(Paragraph("Análise de Vendas — Café", s_sub))
    story.append(Paragraph(f"Período: {data_min} a {data_max}  |  Gerado em: {datetime.now():%d/%m/%Y %H:%M}", s_data))
    story.append(HRFlowable(width="100%", thickness=1.5, color=AZUL_ESCURO, spaceAfter=12))

    # ── KPIs resumo ──
    kpi_data = [
        ["Total Caixas", "Total Kg", "Nº Clientes", "Nº Vendedores", "Famílias"],
        [
            f"{total_cx:,}".replace(",","."),
            f"{int(round(total_kg)):,}".replace(",","."),
            str(rel["N_ Clie_"].nunique()),
            str(rel[rel["Vendedor"].str.contains("Vend|Arm", na=False)==False]["Vendedor"].nunique()),
            str(rel["Familia"].nunique())
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
    cor_var = VERMELHO if var < 0 else VERDE
    evo_intro = Paragraph(
        f"Em 2025 foram vendidas <b>{ano25:,}</b> caixas. Até {ultima_26_txt} (≈ <b>{meses_decorridos_26:.1f}</b> meses) foram "
        f"vendidas <b>{ano26:,}</b> caixas em 2026, o que anualizado (run-rate) representa <b>{proj26:,}</b> caixas — "
        f"uma variação estimada de <font color='{'red' if var<0 else 'green'}'><b>{var:+.1f}%</b></font> face a 2025.".replace(",","."),
        s_corpo
    )
    evo_data = [["Ano", "Caixas (real)", "Kg (real)", "Projecção anual (cx)", "Variação"]]
    evo_data.append(["2025", f"{ano25:,}".replace(",","."), f"{int(round(ano25_kg)):,}".replace(",","."), "—", "—"])
    _var_hex = "#C00000" if var < 0 else "#375623"
    evo_data.append(["2026", f"{ano26:,}".replace(",","."), f"{int(round(ano26_kg)):,}".replace(",","."),
                     f"{proj26:,}".replace(",","."), f"<font color='{_var_hex}'><b>{var:+.1f}%</b></font>"])
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
    ]))
    story.append(KeepTogether([
        Paragraph("1. Evolução Anual", s_secao),
        evo_intro,
        t_evo,
        Spacer(1, 0.3*cm),
        Paragraph(
            f"* Projecção 2026 = volume até {ultima_26_txt} anualizado pelo run-rate "
            f"(× 365 / {dias_decorridos_26} dias decorridos).", s_nota
        ),
    ]))

    # ── Por família ──
    pf = rel.groupby("Familia")[["Quantidade", "Kg"]].sum().sort_values("Quantidade", ascending=False)
    fam_data = [["Família", "Caixas", "Kg", "% Total (cx)"]]
    for f, row in pf.iterrows():
        q = row["Quantidade"]; kg = row["Kg"]
        fam_data.append([f, f"{int(q):,}".replace(",","."), f"{int(round(kg)):,}".replace(",","."),
                         f"{q/total_cx*100:.1f}%"])
    story.append(KeepTogether([
        Paragraph("2. Distribuição por Família", s_secao),
        tabela(fam_data, [7*cm, 3*cm, 3*cm, 3*cm]),
    ]))

    # ── Por vendedor ──
    pv = rel.groupby("Vendedor")[["Quantidade", "Kg"]].sum().sort_values("Quantidade", ascending=False)
    vend_data = [["Vendedor", "Caixas", "Kg", "% Total (cx)", "Nº Clientes"]]
    for v, row in pv.iterrows():
        q = row["Quantidade"]; kg = row["Kg"]
        nc = rel[rel["Vendedor"]==v]["N_ Clie_"].nunique()
        vend_data.append([v, f"{int(q):,}".replace(",","."), f"{int(round(kg)):,}".replace(",","."),
                          f"{q/total_cx*100:.1f}%", str(nc)])
    story.append(KeepTogether([
        Paragraph("3. Desempenho por Vendedor", s_secao),
        tabela(vend_data, [5.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm]),
    ]))

    # ── Top 15 clientes ──
    pc = rel.groupby(["N_ Clie_","Cliente","Vendedor"])[["Quantidade", "Kg"]].sum().sort_values("Quantidade", ascending=False).head(15)
    top_data = [["#", "Cliente", "Vendedor", "Caixas", "Kg", "% Total (cx)"]]
    for i, ((cod,nome,vend), row) in enumerate(pc.iterrows(), 1):
        q = row["Quantidade"]; kg = row["Kg"]
        top_data.append([str(i), nome, vend.split("/")[-1] if "/" in vend else vend,
                         f"{int(q):,}".replace(",","."), f"{int(round(kg)):,}".replace(",","."),
                         f"{q/total_cx*100:.1f}%"])
    story.append(KeepTogether([
        Paragraph("4. Top 15 Clientes — Período Total", s_secao),
        tabela(top_data, [0.8*cm, 6.5*cm, 3*cm, 2*cm, 2*cm, 1.7*cm]),
    ]))

    # ── Alertas ──
    clientes_25 = set(rel[rel["Ano"]==2025]["N_ Clie_"].unique())
    clientes_26 = set(rel[rel["Ano"]==2026]["N_ Clie_"].unique())
    perdidos = clientes_25 - clientes_26
    perdidos_df = (
        rel[rel["N_ Clie_"].isin(perdidos) & (rel["Ano"]==2025)]
        .groupby(["Vendedor","N_ Clie_","Cliente"])[["Quantidade", "Kg"]].sum()
        .sort_values("Quantidade", ascending=False)
        .reset_index()
    )
    # filtrar só os com volume relevante (>= 5 cx)
    perdidos_rel = perdidos_df[perdidos_df["Quantidade"] >= 5]

    bloco5 = [
        Paragraph("5. Alertas — Clientes em Risco", s_secao),
        Paragraph(
            f"<b>{len(perdidos)}</b> clientes compraram em 2025 mas <b>não registam qualquer compra em 2026</b>. "
            f"Os {len(perdidos_rel)} com volume igual ou superior a 5 caixas em 2025 são:",
            s_corpo
        ),
    ]
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
        bloco5.append(t_perd)
    story.append(KeepTogether(bloco5))

    # ── Clientes novos ──
    story.append(Spacer(1, 0.5*cm))
    novos = clientes_26 - clientes_25
    novos_df = (
        rel[rel["N_ Clie_"].isin(novos) & (rel["Ano"]==2026)]
        .groupby(["Vendedor","N_ Clie_","Cliente"])[["Quantidade", "Kg"]].sum()
        .sort_values("Quantidade", ascending=False)
        .reset_index()
    )
    bloco6 = [
        Paragraph("6. Clientes Novos em 2026", s_secao),
        Paragraph(
            f"<b>{len(novos)}</b> clientes iniciaram compras nestas famílias em 2026 (sem historial em 2025), "
            f"totalizando <b>{novos_df['Quantidade'].sum()}</b> caixas.",
            s_corpo
        ),
    ]
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
        bloco6.append(t_nov)
    story.append(KeepTogether(bloco6))

    # ── Análise por vendedor ──
    sec7_head = Paragraph("7. Análise por Vendedor", s_secao)
    sec7_intro = Paragraph(
        "Evolução da carteira de cada vendedor entre 2025 e 2026 (ano em curso): clientes mantidos, "
        "perdidos e captados, taxa de retenção e projecção de volume anualizada para 2026.",
        s_corpo
    )

    # Cálculos por vendedor (excluindo estruturas não comerciais)
    excluir_vend = ["Vendas Internas", "14/Armazém"]
    vendedores_ord = [v for v in rel.sort_values("N_ Vend_")["Vendedor"].unique() if v not in excluir_vend]

    vend_rows = []
    for v in vendedores_ord:
        dv = rel[rel["Vendedor"] == v]
        c25 = set(dv[dv["Ano"] == 2025]["N_ Clie_"])
        c26 = set(dv[dv["Ano"] == 2026]["N_ Clie_"])
        mantidos = len(c25 & c26)
        perd_v   = len(c25 - c26)
        novos_v  = len(c26 - c25)
        retencao = (len(c25 & c26) / len(c25) * 100) if c25 else 0
        q25 = int(dv[dv["Ano"] == 2025]["Quantidade"].sum())
        q26 = int(dv[dv["Ano"] == 2026]["Quantidade"].sum())
        proj = int(round(q26 * fator_anual))
        var_v = ((proj - q25) / q25 * 100) if q25 else 0
        vend_rows.append({
            "vendedor": v, "c25": len(c25), "c26": len(c26),
            "mantidos": mantidos, "perdidos": perd_v, "novos": novos_v,
            "retencao": retencao, "q25": q25, "q26": q26, "proj": proj, "var": var_v,
        })

    # Tabela: clientes por vendedor
    HEX_VERDE = "#375623"
    HEX_VERM  = "#C00000"
    vend_cli = [["Vendedor", "Clientes<br/>2025", "Clientes<br/>2026", "Mantidos", "Perdidos", "Novos", "Retenção"]]
    for r in vend_rows:
        # cor embutida no texto (TEXTCOLOR não afecta Paragraphs)
        ret_txt = f"{r['retencao']:.0f}%"
        if r["retencao"] >= 90:
            ret_txt = f"<font color='{HEX_VERDE}'>{ret_txt}</font>"
        elif r["retencao"] < 80:
            ret_txt = f"<font color='{HEX_VERM}'>{ret_txt}</font>"
        perd_txt = str(r["perdidos"])
        if r["perdidos"] > r["novos"]:
            perd_txt = f"<font color='{HEX_VERM}'>{perd_txt}</font>"
        vend_cli.append([
            r["vendedor"], str(r["c25"]), str(r["c26"]),
            str(r["mantidos"]), perd_txt, str(r["novos"]),
            ret_txt,
        ])
    t_vc = tabela(vend_cli, [4.3*cm, 1.9*cm, 1.9*cm, 2*cm, 2*cm, 1.6*cm, 2*cm])
    story.append(KeepTogether([
        sec7_head, sec7_intro, t_vc,
        Spacer(1, 0.2*cm),
        Paragraph(
            "Retenção: % de clientes de 2025 que voltaram a comprar em 2026 "
            "(<font color='green'>verde ≥ 90%</font>, <font color='red'>vermelho &lt; 80%</font>).",
            s_nota
        ),
    ]))
    story.append(Spacer(1, 0.35*cm))

    # Tabela: volume e projecção por vendedor
    vend_vol = [["Vendedor", "Qtd 2025<br/>(cx)", "Qtd 2026<br/>decorrido (cx)", "Projecção<br/>2026 (cx)", "Variação"]]
    for r in vend_rows:
        var_txt = f"<font color='{HEX_VERDE if r['var'] >= 0 else HEX_VERM}'>{r['var']:+.1f}%</font>"
        vend_vol.append([
            r["vendedor"],
            f"{r['q25']:,}".replace(",","."),
            f"{r['q26']:,}".replace(",","."),
            f"{r['proj']:,}".replace(",","."),
            var_txt,
        ])
    t_vv = tabela(vend_vol, [4.3*cm, 2.8*cm, 2.8*cm, 2.8*cm, 2.8*cm])
    story.append(KeepTogether([
        t_vv,
        Spacer(1, 0.15*cm),
        Paragraph("* Projecção 2026 = volume decorrido anualizado pelo run-rate (× 365 / dias decorridos).", s_nota),
    ]))
    story.append(Spacer(1, 0.35*cm))

    # Comentário automático: melhores e piores
    if vend_rows:
        top_ret  = max(vend_rows, key=lambda r: r["retencao"])
        low_ret  = min(vend_rows, key=lambda r: r["retencao"])
        top_cres = max(vend_rows, key=lambda r: r["var"])
        low_cres = min(vend_rows, key=lambda r: r["var"])
        top_capt = max(vend_rows, key=lambda r: r["novos"])

        coment = (
            f"<b>Destaques:</b> Em retenção de clientes, <b>{top_ret['vendedor']}</b> lidera "
            f"(<font color='green'>{top_ret['retencao']:.0f}%</font>), enquanto "
            f"<b>{low_ret['vendedor']}</b> apresenta a taxa mais baixa "
            f"(<font color='red'>{low_ret['retencao']:.0f}%</font>) e merece acompanhamento. "
            f"Em tendência de volume, <b>{top_cres['vendedor']}</b> tem a melhor projecção "
            f"(<font color='{'green' if top_cres['var']>=0 else 'red'}'>{top_cres['var']:+.1f}%</font>) "
            f"e <b>{low_cres['vendedor']}</b> a mais desfavorável "
            f"(<font color='{'green' if low_cres['var']>=0 else 'red'}'>{low_cres['var']:+.1f}%</font>). "
            f"Na captação de novos clientes destaca-se <b>{top_capt['vendedor']}</b> "
            f"com {top_capt['novos']} novos clientes em 2026."
        )
        story.append(Paragraph(coment, s_corpo))

    # ── Avaliação de performance ──
    # Formatação de números com separador de milhares "." (sem mexer nas vírgulas de texto)
    def fmt(n):
        return f"{int(round(n)):,}".replace(",", ".")
    def fmts(n):   # com sinal
        return f"{int(round(n)):+,}".replace(",", ".")
    def cor(v):
        return "green" if v >= 0 else "red"

    # Grupos de famílias para o desdobramento
    GRUPO_CAFE = [
        "Buondi Grão Kg", "Sical Grão Kg", "Christina Grão Kg",
        "Buondi Decafe Doses", "Sical Decafe Doses", "Christina Decafe Doses",
        "Buondi Decafe Grão Kg", "Sical Decafe Grão Kg", "Christina Decafe Grão Kg",
    ]

    def metricas_grupo(familias):
        sub = rel[rel["Familia"].isin(familias)]
        a25  = sub[sub["Ano"] == 2025]["Quantidade"].sum()
        a25k = sub[sub["Ano"] == 2025]["Kg"].sum()
        a26  = sub[sub["Ano"] == 2026]["Quantidade"].sum()
        a26k = sub[sub["Ano"] == 2026]["Kg"].sum()
        proj  = int(round(a26 * fator_anual))
        projk = int(round(a26k * fator_anual))
        varg  = ((proj - a25) / a25 * 100) if a25 else 0
        c25 = set(sub[sub["Ano"] == 2025]["N_ Clie_"])
        c26 = set(sub[sub["Ano"] == 2026]["N_ Clie_"])
        perd = c25 - c26
        nov  = c26 - c25
        reten = (len(c25 & c26) / len(c25) * 100) if c25 else 0
        pdf = sub[sub["N_ Clie_"].isin(perd) & (sub["Ano"] == 2025)]
        ndf = sub[sub["N_ Clie_"].isin(nov) & (sub["Ano"] == 2026)]
        cx_perd = int(pdf["Quantidade"].sum()); kg_perd = int(round(pdf["Kg"].sum()))
        cx_nov  = int(ndf["Quantidade"].sum()); kg_nov  = int(round(ndf["Kg"].sum()))
        return {
            "a25": int(a25), "a25k": int(round(a25k)), "proj": proj, "projk": projk, "var": varg,
            "c25": len(c25), "mant": len(c25 & c26), "perd": len(perd), "nov": len(nov),
            "reten": reten, "cx_perd": cx_perd, "kg_perd": kg_perd,
            "cx_nov": cx_nov, "kg_nov": kg_nov,
            "saldo_cx": cx_nov - cx_perd, "saldo_kg": kg_nov - kg_perd,
        }

    grupos = [
        ("Café (grão, decaf. grão e decaf. doses)", metricas_grupo(GRUPO_CAFE)),
        ("Buondi Office", metricas_grupo(["Buondi Office"])),
        ("Café 200g",    metricas_grupo(["Café 200g"])),
    ]

    # Totais gerais (todas as famílias)
    kg_perdidos    = int(round(perdidos_df["Kg"].sum())) if not perdidos_df.empty else 0
    cx_perdidos    = int(perdidos_df["Quantidade"].sum()) if not perdidos_df.empty else 0
    kg_novos       = int(round(novos_df["Kg"].sum())) if not novos_df.empty else 0
    cx_novos       = int(novos_df["Quantidade"].sum()) if not novos_df.empty else 0
    saldo_kg       = kg_novos - kg_perdidos
    saldo_cx       = cx_novos - cx_perdidos
    taxa_retencao  = (len(clientes_26) / len(clientes_25) * 100) if clientes_25 else 0

    s_aval = ParagraphStyle("aval", fontName="Helvetica", fontSize=9, leading=14,
                             spaceAfter=8, leftIndent=4, rightIndent=4)
    s_aval_primeiro = ParagraphStyle("aval0", parent=s_aval, spaceBefore=4)

    # § Tendência geral (por grupo)
    lin_tend = []
    for nome, m in grupos:
        lin_tend.append(
            f"• <b>{nome}:</b> projecção <b>{fmt(m['proj'])} cx ({fmt(m['projk'])} kg)</b> "
            f"vs {fmt(m['a25'])} cx ({fmt(m['a25k'])} kg) em 2025 → "
            f"<font color='{cor(m['var'])}'><b>{m['var']:+.1f}%</b></font>"
        )
    para_tend = (
        f"<b>Tendência geral</b> — projecção anualizada para 2026 "
        f"(volume até {ultima_26_txt}, ≈ {meses_decorridos_26:.1f} meses):<br/>"
        + "<br/>".join(lin_tend)
    )

    # § Retenção de clientes (por grupo)
    lin_ret = []
    for nome, m in grupos:
        lin_ret.append(
            f"• <b>{nome}:</b> {m['c25']} clientes em 2025; {m['mant']} mantidos "
            f"(retenção <b>{m['reten']:.0f}%</b>); {m['perd']} perdidos "
            f"= {fmt(m['cx_perd'])} cx / {fmt(m['kg_perd'])} kg"
        )
    para_ret = "<b>Retenção de clientes</b> (por grupo, clientes que compraram cada grupo):<br/>" + "<br/>".join(lin_ret)

    # § Captação vs. perda (por grupo)
    lin_capt = []
    for nome, m in grupos:
        lin_capt.append(
            f"• <b>{nome}:</b> {m['nov']} novos = {fmt(m['cx_nov'])} cx / {fmt(m['kg_nov'])} kg; "
            f"saldo líquido <font color='{cor(m['saldo_kg'])}'><b>{fmts(m['saldo_cx'])} cx / {fmts(m['saldo_kg'])} kg</b></font>"
        )
    para_capt = "<b>Captação vs. perda</b> (novos menos perdidos, por grupo):<br/>" + "<br/>".join(lin_capt)

    # § Resumo geral (todas as famílias)
    para_geral = (
        f"<b>Resumo geral (todas as famílias):</b> projecção de "
        f"<b>{fmt(proj26)} cx ({fmt(proj26_kg)} kg)</b> vs {fmt(ano25)} cx ({fmt(ano25_kg)} kg) em 2025 "
        f"(<font color='{cor(var)}'><b>{var:+.1f}%</b></font>); taxa de retenção de <b>{taxa_retencao:.1f}%</b> "
        f"({len(perdidos)} perdidos / {len(novos)} novos); saldo líquido "
        f"<font color='{cor(saldo_kg)}'><b>{fmts(saldo_cx)} cx / {fmts(saldo_kg)} kg</b></font> — "
        f"{'os novos clientes <b>não compensam</b> o volume perdido' if saldo_kg < 0 else 'os novos clientes <b>superam</b> o volume perdido'}."
    )

    # § Nota
    para_nota = (
        f"<b>Nota:</b> Os dados de 2026 vão até {ultima_26_txt} (ano em curso). "
        f"A projecção assume que o ritmo de vendas decorrido se mantém até ao final do ano; "
        f"é possível que alguns dos {len(perdidos)} clientes classificados como \"perdidos\" retomem "
        f"compras no período restante, pelo que esta análise deve ser revisitada no final do ano."
    )

    # Manter o título com o primeiro bloco
    story.append(KeepTogether([
        Paragraph("8. Avaliação de Performance", s_secao),
        Paragraph(para_tend, s_aval_primeiro),
    ]))
    for p in [para_ret, para_capt, para_geral, para_nota]:
        story.append(Paragraph(p, s_aval))


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
