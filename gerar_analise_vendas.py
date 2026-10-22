"""
Análise de Vendas — relatório único Trovidoce.

Funde num só PDF ("Análise_Vendas.pdf"):
  I   — Análise de Café (famílias de café, com conversão para kg)
  II  — Análise de Outras Famílias (em unidades)
  III — Custos Logísticos por Vendedor (a partir do mapa de kms)

Lê automaticamente, da pasta PASTA_FONTE:
  - "COLAR AQUI para Analise Familias.xlsx"  (vendas)
  - "Kms frota.xlsx"                          (quilómetros)

Uso:
    python gerar_analise_vendas.py
"""

import os
import sys
import glob
import tempfile
from datetime import datetime

try:
    import pymupdf
except ImportError:
    import fitz as pymupdf

import gerar_relatorio_pdf          as cafe
import gerar_relatorio_pdf_outras   as outras
import gerar_relatorio_pdf_logistica as logistica

# ── Configuração ─────────────────────────────────────────────────────────────
PASTA_FONTE   = r"I:\O meu disco\Trovidoce\Mapas Mensais"
PASTA_DESTINO = r"I:\O meu disco\Claude"
NOME_SAIDA    = "Análise_Vendas.pdf"
# ─────────────────────────────────────────────────────────────────────────────


def gerar(ficheiro_vendas=None, ficheiro_kms=None):
    # Garantir que os sub-módulos procuram na mesma pasta
    for mod in (cafe, outras, logistica):
        mod.PASTA_FONTE = PASTA_FONTE

    if ficheiro_vendas is None:
        ficheiro_vendas = cafe.encontrar_ficheiro_fonte()
    if ficheiro_kms is None:
        ficheiro_kms = logistica.encontrar_ficheiro_kms()

    os.makedirs(PASTA_DESTINO, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="analise_vendas_")

    p_cafe = os.path.join(tmp, "parte_cafe.pdf")
    p_out  = os.path.join(tmp, "parte_outras.pdf")
    p_log  = os.path.join(tmp, "parte_logistica.pdf")

    # Parte I — Café
    cafe.PASTA_DESTINO, cafe.NOME_SAIDA = tmp, "parte_cafe.pdf"
    cafe.gerar(ficheiro_vendas)

    # Parte II — Outras famílias
    outras.PASTA_DESTINO, outras.NOME_SAIDA = tmp, "parte_outras.pdf"
    outras.gerar(ficheiro_vendas)

    # Parte III — Logística
    logistica.gerar(ficheiro_vendas, ficheiro_kms, saida=p_log)

    # Fundir
    saida = os.path.join(PASTA_DESTINO, NOME_SAIDA)
    doc = pymupdf.open()
    for parte in (p_cafe, p_out, p_log):
        with pymupdf.open(parte) as sub:
            doc.insert_pdf(sub)
    doc.save(saida)
    doc.close()

    print(f"[{datetime.now():%H:%M:%S}] Relatório único guardado em:\n  {saida}")
    return saida


if __name__ == "__main__":
    gerar()
