"""
Secção de Custos Logísticos + Geografia/Rotas por Vendedor — Trovidoce.

Lê, da pasta PASTA_FONTE:
  - mapa de vendas ("COLAR AQUI para Analise Familias.xlsx")
  - mapa de kms    ("Kms frota.xlsx")
  - mapa de clientes ("Clientes.XLS")  -> concelho / distrito por cliente

Produz:
  - Custo de combustível por vendedor (por ano)
  - Vendas por distrito e por concelho (com vendedor dominante)
  - Sobreposição de rotas (concelhos servidos por vários vendedores)
  - Custo logístico imputado por concelho (proxy) e rácio custo/vendas
  - Clientes a rever (baixo volume em concelho dominado por outro vendedor)

Pressupostos (editáveis abaixo):
  - Consumo médio: 12 L / 100 km
  - Preço do gasóleo CEPSA/Moeve por ano (€/litro)
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
MARCA_COMBUSTIVEL = "CEPSA / Moeve"
PRECO_GASOLEO  = {               # €/litro, média anual da marca
    2025: 1.70,
    2026: 2.10,
}
PRECO_GASOLEO_DEFAULT = 2.00

# Carro de distribuição -> vendedor a quem os kms são imputados
CARRO_DISTRIBUICAO = 14
VENDEDOR_DISTRIBUICAO = 31

# Limiar para "clientes a rever" (vendas líquidas do cliente abaixo deste valor)
LIMIAR_BAIXO_VOLUME = 2582      # ≈ mediana das vendas por cliente

# Cores
AZUL_ESCURO      = colors.HexColor("#1F4E79")
AZUL_MEDIO       = colors.HexColor("#2E75B6")
AZUL_CLARO       = colors.HexColor("#D6E4F0")
AZUL_MUITO_CLARO = colors.HexColor("#EBF3FB")
VERMELHO         = colors.HexColor("#C00000")
VERDE            = colors.HexColor("#375623")
BRANCO           = colors.white
# ─────────────────────────────────────────────────────────────────────────────


def _encontrar(nome_contains, nome_exacto=None, exts=("xlsx",)):
    if nome_exacto:
        p = os.path.join(PASTA_FONTE, nome_exacto)
        if os.path.exists(p):
            return p
    for ext in exts:
        fich = glob.glob(os.path.join(PASTA_FONTE, f"*{nome_contains}*.{ext}"))
        if fich:
            return max(fich, key=os.path.getmtime)
    raise FileNotFoundError(f"Ficheiro com '{nome_contains}' não encontrado em:\n  {PASTA_FONTE}")


def encontrar_ficheiro_vendas():
    return _encontrar("COLAR AQUI", "COLAR AQUI para Analise Familias.xlsx")

def encontrar_ficheiro_kms():
    return _encontrar("Kms", "Kms frota.xlsx")

def encontrar_ficheiro_clientes():
    return _encontrar("Clientes", "Clientes.XLS", exts=("XLS", "xls", "xlsx"))


def ler_kms(ficheiro_kms):
    """{ano: {num_vendedor: kms}} com o carro 14 somado ao vendedor 31."""
    xls = pd.ExcelFile(ficheiro_kms)
    out = {}
    for sh in xls.sheet_names:
        try:
            ano = int(str(sh).strip())
        except ValueError:
            continue
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
        if CARRO_DISTRIBUICAO in kms:
            kms[VENDEDOR_DISTRIBUICAO] = kms.get(VENDEDOR_DISTRIBUICAO, 0) + kms.pop(CARRO_DISTRIBUICAO)
        out[ano] = kms
    return out


def ler_clientes(ficheiro_clientes):
    """Devolve DataFrame com N_, Empresa, Concelho, Distrito (por código de cliente)."""
    cli = pd.read_excel(ficheiro_clientes, sheet_name=0)
    cli.columns = [str(c).strip() for c in cli.columns]
    cols = {c.lower(): c for c in cli.columns}
    def col(nome, alt=None):
        return cols.get(nome.lower(), alt)
    out = pd.DataFrame({
        "N_Clie": pd.to_numeric(cli[col("N_")], errors="coerce"),
        "Empresa": cli[col("Empresa")] if col("Empresa") else None,
        "Concelho": cli[col("Concelho")] if col("Concelho") else None,
        "Distrito": cli[col("Distrito")] if col("Distrito") else None,
    })
    out = out.dropna(subset=["N_Clie"]).drop_duplicates("N_Clie")
    out["N_Clie"] = out["N_Clie"].astype(int)
    return out


def preparar(ficheiro_vendas, ficheiro_clientes):
    """Carrega vendas detalhadas por cliente, com geografia e vendedor."""
    df = pd.read_excel(ficheiro_vendas, sheet_name="COLAR AQUI", header=1)
    df["Dt_ Emissao"] = pd.to_datetime(df["Dt_ Emissao"])
    df["Ano"] = df["Dt_ Emissao"].dt.year
    mapa_nome = dict(df[["N_ Vend_", "Vendedor"]].drop_duplicates().values)

    geo = ler_clientes(ficheiro_clientes)
    df = df.merge(geo, left_on="N_ Clie_", right_on="N_Clie", how="left")
    df["Concelho"] = df["Concelho"].fillna("(sem concelho)")
    df["Distrito"] = df["Distrito"].fillna("(sem distrito)")

    # vendas líquidas por vendedor/ano (para a secção de combustível)
    vendas_por_ano = {}
    for (ano, vend), val in df.groupby(["Ano", "N_ Vend_"])["Total Liquido"].sum().items():
        vendas_por_ano.setdefault(int(ano), {})[int(vend)] = float(val)

    return df, mapa_nome, vendas_por_ano


def preco_gasoleo(ano):
    return PRECO_GASOLEO.get(ano, PRECO_GASOLEO_DEFAULT)


def custo_total_por_vendedor(kms_por_ano):
    """{num_vendedor: custo_combustivel_total}"""
    out = {}
    for ano, kms in kms_por_ano.items():
        p = preco_gasoleo(ano)
        for v, km in kms.items():
            out[v] = out.get(v, 0.0) + km * CONSUMO_L_100 / 100.0 * p
    return out


# ── Helpers de estilo/tabela ─────────────────────────────────────────────────
def _helpers():
    def estilo(nome, **kw):
        return ParagraphStyle(nome, **kw)
    s_secao = estilo("secao", fontSize=12, textColor=BRANCO, fontName="Helvetica-Bold",
                     alignment=TA_LEFT, backColor=AZUL_ESCURO, leftIndent=6,
                     spaceAfter=8, spaceBefore=14)
    s_sub   = estilo("sub", fontSize=10.5, textColor=AZUL_ESCURO, fontName="Helvetica-Bold",
                     spaceBefore=8, spaceAfter=4)
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

    return {"s_secao": s_secao, "s_sub": s_sub, "s_corpo": s_corpo, "s_nota": s_nota, "tabela": tabela}


def _eur(n):   return f"{int(round(n)):,}".replace(",", ".") + " €"
def _num(n):   return f"{int(round(n)):,}".replace(",", ".")
def _vshort(v): return v.split("/")[-1] if isinstance(v, str) and "/" in v else str(v)


# ── Construção das secções ───────────────────────────────────────────────────
def build(story, H, df, mapa_nome, vendas_por_ano, kms_por_ano):
    s_secao = H["s_secao"]; s_sub = H["s_sub"]; s_corpo = H["s_corpo"]; s_nota = H["s_nota"]
    tabela = H["tabela"]

    anos = sorted(set(kms_por_ano) | set(vendas_por_ano))
    vendedores_km = sorted({v for a in kms_por_ano.values() for v in a})
    custo_total_v = custo_total_por_vendedor(kms_por_ano)

    def nome_vend(num):
        nome = mapa_nome.get(num, str(num))
        if num == VENDEDOR_DISTRIBUICAO:
            nome = f"{nome} (c/ carro {CARRO_DISTRIBUICAO})"
        return nome

    # ======================= 1. Custo de combustível por vendedor =======================
    precos_txt = "; ".join(f"{a}: {preco_gasoleo(a):.2f} €/l" for a in anos)
    intro = (
        f"Estimativa do custo de combustível por vendedor, a partir dos quilómetros percorridos. "
        f"Pressupostos: consumo médio de <b>{CONSUMO_L_100:.0f} L/100 km</b> e preço médio do gasóleo "
        f"<b>{MARCA_COMBUSTIVEL}</b> (<b>{precos_txt}</b>). Os quilómetros do carro de distribuição "
        f"(nº {CARRO_DISTRIBUICAO}) são imputados ao vendedor {VENDEDOR_DISTRIBUICAO}."
    )
    blocos = [Paragraph("Custos Logísticos e Geografia de Rotas", s_secao),
              Paragraph(intro, s_corpo)]
    for ano in anos:
        kms_ano = kms_por_ano.get(ano, {}); vendas_ano = vendas_por_ano.get(ano, {}); p = preco_gasoleo(ano)
        linhas = [["Vendedor", "Kms", "Litros", "Custo Gasóleo", "Vendas Líq.", "Custo/Vendas"]]
        tk = tl = tc = tv = 0.0
        for v in vendedores_km:
            km = kms_ano.get(v, 0.0)
            if km == 0: continue
            litros = km * CONSUMO_L_100 / 100.0; custo = litros * p; vnd = vendas_ano.get(v, 0.0)
            ratio = (custo / vnd * 100) if vnd else 0
            tk += km; tl += litros; tc += custo; tv += vnd
            linhas.append([nome_vend(v), _num(km), _num(litros), _eur(custo), _eur(vnd), f"{ratio:.2f}%"])
        linhas.append(["TOTAL", _num(tk), _num(tl), _eur(tc), _eur(tv), f"{(tc/tv*100) if tv else 0:.2f}%"])
        t = tabela(linhas, [5*cm, 2.2*cm, 2.2*cm, 2.6*cm, 2.6*cm, 2.4*cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, len(linhas)-1), (-1, len(linhas)-1), AZUL_CLARO),
            ("FONTNAME", (0, len(linhas)-1), (-1, len(linhas)-1), "Helvetica-Bold"),
        ]))
        blocos.append(Paragraph(f"<b>{ano}</b> — preço do gasóleo: {p:.2f} €/l", s_corpo))
        blocos.append(t)
    story.append(KeepTogether(blocos[:3]))
    for b in blocos[3:]:
        story.append(b)

    df_c = df[df["Vendedor"].str.contains("/", na=False)].copy()   # só vendedores de rota

    # ======================= 2. Vendas por distrito =======================
    gd = df.groupby("Distrito").agg(
        Vendas=("Total Liquido", "sum"), Clientes=("N_ Clie_", "nunique")
    ).sort_values("Vendas", ascending=False)
    tot_vd = gd["Vendas"].sum()
    linhas = [["Distrito", "Vendas Líq.", "% Total", "Nº Clientes"]]
    for dist, r in gd.iterrows():
        linhas.append([dist, _eur(r["Vendas"]), f"{r['Vendas']/tot_vd*100:.1f}%", str(int(r["Clientes"]))])
    story.append(KeepTogether([
        Paragraph("Vendas por Distrito", s_sub),
        tabela(linhas, [6*cm, 3.2*cm, 2.3*cm, 2.5*cm]),
    ]))

    # ======================= 3. Vendas por concelho + vendedor dominante =======================
    gc = df.groupby("Concelho").agg(
        Vendas=("Total Liquido", "sum"), Clientes=("N_ Clie_", "nunique")
    ).sort_values("Vendas", ascending=False)
    # vendedor dominante (por vendas) de cada concelho
    dom = (df_c.groupby(["Concelho", "Vendedor"])["Total Liquido"].sum()
           .reset_index().sort_values("Total Liquido", ascending=False)
           .drop_duplicates("Concelho").set_index("Concelho")["Vendedor"])
    nv = df_c.groupby("Concelho")["Vendedor"].nunique()
    linhas = [["Concelho", "Vendas Líq.", "% Total", "Clientes", "Nº Vend.", "Vendedor dominante"]]
    for conc, r in gc.head(18).iterrows():
        linhas.append([
            conc, _eur(r["Vendas"]), f"{r['Vendas']/tot_vd*100:.1f}%", str(int(r["Clientes"])),
            str(int(nv.get(conc, 0))), _vshort(dom.get(conc, "—")),
        ])
    story.append(KeepTogether([
        Paragraph("Vendas por Concelho (top 18) e vendedor dominante", s_sub),
        tabela(linhas, [3.6*cm, 2.8*cm, 1.7*cm, 1.7*cm, 1.6*cm, 3.6*cm]),
    ]))

    # ======================= 4. Sobreposição de rotas =======================
    conc_multi = nv[nv > 1].sort_values(ascending=False)
    linhas = [["Concelho", "Nº Vend.", "Vendedores (vendas líq.)"]]
    for conc in conc_multi.index:
        sub = (df_c[df_c["Concelho"] == conc].groupby("Vendedor")["Total Liquido"].sum()
               .sort_values(ascending=False))
        detalhe = ", ".join(f"{_vshort(v)} ({_eur(val)})" for v, val in sub.items())
        linhas.append([conc, str(int(conc_multi[conc])), detalhe])
    intro4 = (
        f"<b>{len(conc_multi)}</b> concelhos são visitados por mais do que um vendedor. "
        f"Concelhos com rotas mais fragmentadas são candidatos a reorganização/consolidação."
    )
    story.append(KeepTogether([
        Paragraph("Sobreposição de Rotas (concelhos com &gt;1 vendedor)", s_sub),
        Paragraph(intro4, s_corpo),
        tabela(linhas, [3.2*cm, 1.5*cm, 11*cm]),
    ]))

    # ======================= 5. Custo logístico imputado por concelho =======================
    # custo do vendedor repartido pelos concelhos proporcionalmente às suas vendas
    vendas_v_total = df_c.groupby("Vendedor")["Total Liquido"].sum()
    vendas_vc = df_c.groupby(["Vendedor", "Concelho"])["Total Liquido"].sum()
    # mapear nº vendedor -> label para custo
    label_custo = {}
    for num, c in custo_total_v.items():
        label_custo[mapa_nome.get(num, str(num))] = label_custo.get(mapa_nome.get(num, str(num)), 0) + c
    custo_conc = {}
    for (vlabel, conc), vnd in vendas_vc.items():
        vt = vendas_v_total.get(vlabel, 0)
        cvend = label_custo.get(vlabel, 0)
        if vt > 0 and cvend > 0:
            custo_conc[conc] = custo_conc.get(conc, 0) + cvend * (vnd / vt)
    linhas = [["Concelho", "Vendas Líq.", "Custo Log. (imp.)", "Custo/Vendas"]]
    rows = []
    for conc, r in gc.iterrows():
        cc = custo_conc.get(conc, 0)
        if cc == 0: continue
        rows.append((conc, r["Vendas"], cc, cc / r["Vendas"] * 100 if r["Vendas"] else 0))
    rows.sort(key=lambda x: x[3], reverse=True)   # pior rácio primeiro
    for conc, vnd, cc, ratio in rows[:18]:
        linhas.append([conc, _eur(vnd), _eur(cc), f"{ratio:.2f}%"])
    intro5 = (
        "Custo de combustível de cada vendedor repartido pelos concelhos que serve, proporcionalmente "
        "às vendas (aproximação — os kms reais por concelho não são conhecidos). Ordenado pelo pior "
        "rácio custo/vendas. Rácios altos sinalizam zonas caras de servir face ao que rendem."
    )
    story.append(KeepTogether([
        Paragraph("Custo Logístico Imputado por Concelho (proxy)", s_sub),
        Paragraph(intro5, s_corpo),
        tabela(linhas, [4.5*cm, 3*cm, 3.2*cm, 2.8*cm]),
    ]))

    # ======================= 6. Clientes a rever =======================
    vc = df_c.groupby(["N_ Clie_", "Vendedor", "Concelho"]).agg(
        Vendas=("Total Liquido", "sum")
    ).reset_index()
    nome_cli = (df.sort_values("Dt_ Emissao").groupby("N_ Clie_")["Cliente"].last())
    cand = []
    for _, r in vc.iterrows():
        conc = r["Concelho"]
        if nv.get(conc, 0) <= 1:           # só concelhos com sobreposição
            continue
        dominante = dom.get(conc)
        if dominante == r["Vendedor"]:     # já é o dominante
            continue
        if r["Vendas"] >= LIMIAR_BAIXO_VOLUME:
            continue
        cand.append((r["Vendas"], nome_cli.get(r["N_ Clie_"], r["N_ Clie_"]),
                     conc, r["Vendedor"], dominante))
    cand.sort(key=lambda x: x[0])          # menor volume primeiro
    linhas = [["Cliente", "Concelho", "Vendedor atual", "Vendas Líq.", "Sugestão (dominante)"]]
    for vnd, nome, conc, vatual, dmn in cand[:20]:
        linhas.append([str(nome), conc, _vshort(vatual), _eur(vnd), _vshort(dmn)])
    intro6 = (
        f"Clientes de baixo volume (&lt; {_eur(LIMIAR_BAIXO_VOLUME)}/período) situados em concelhos já "
        f"cobertos por outro vendedor, e cujo fornecedor actual não é o vendedor dominante da zona. "
        f"São candidatos naturais a reafectação, reduzindo deslocações. "
        f"Total identificado: <b>{len(cand)}</b> (mostrados os {min(20, len(cand))} de menor volume)."
    )
    story.append(KeepTogether([
        Paragraph("Clientes a Rever (otimização de rotas)", s_sub),
        Paragraph(intro6, s_corpo),
        tabela(linhas, [5.5*cm, 3*cm, 2.6*cm, 2.2*cm, 3.2*cm]),
    ]))

    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Nota: o custo refere-se apenas a combustível (não inclui portagens, manutenção, amortização "
        "nem salários). A geografia provém do mapa de clientes (concelho/distrito). A imputação de "
        "custo por concelho é uma aproximação proporcional às vendas, não um custo medido por rota.",
        s_nota
    ))


def gerar(ficheiro_vendas=None, ficheiro_kms=None, ficheiro_clientes=None, saida=None):
    if ficheiro_vendas is None:   ficheiro_vendas = encontrar_ficheiro_vendas()
    if ficheiro_kms is None:      ficheiro_kms = encontrar_ficheiro_kms()
    if ficheiro_clientes is None: ficheiro_clientes = encontrar_ficheiro_clientes()

    print(f"[{datetime.now():%H:%M:%S}] Logística/Geografia — vendas: {os.path.basename(ficheiro_vendas)} | "
          f"kms: {os.path.basename(ficheiro_kms)} | clientes: {os.path.basename(ficheiro_clientes)}")

    kms_por_ano = ler_kms(ficheiro_kms)
    df, mapa_nome, vendas_por_ano = preparar(ficheiro_vendas, ficheiro_clientes)

    os.makedirs(PASTA_DESTINO, exist_ok=True)
    if saida is None:
        saida = os.path.join(PASTA_DESTINO, NOME_SAIDA)

    doc = SimpleDocTemplate(saida, pagesize=A4, leftMargin=1.8*cm, rightMargin=1.8*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    H = _helpers()
    story = [Spacer(1, 0.3*cm)]
    build(story, H, df, mapa_nome, vendas_por_ano, kms_por_ano)
    doc.build(story)
    print(f"[{datetime.now():%H:%M:%S}] PDF logística/geografia guardado em:\n  {saida}")
    return saida


if __name__ == "__main__":
    gerar()
