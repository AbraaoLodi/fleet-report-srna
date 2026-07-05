"""Página 2 — 01. Disponibilidade Fabricante.

Conteúdo:
  - 4 KPI cards: Disp S{atual}, S{-1}, S{-2}, Pior semana
  - Gráfico de linha com meta 98% tracejada
  - Tabela histórica (Semana, Disp, Meta, Gap, Situação)
"""
from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
from docx.document import Document as DocumentT

from utils.charts import grafico_disponibilidade
from utils.docx_helpers import (
    HEX_CORAL, HEX_GRAFITE, HEX_VERDE, HEX_CINZA_CLARO,
    RGB_CORAL, RGB_GRAFITE, RGB_VERDE,
    add_capitulo, add_kpi_cards, add_imagem, add_tabela,
    quebra_pagina,
)
from utils.formatters import fmt_pct
from utils.logger import get_logger

log = get_logger(__name__)

META = 98.00


def _fmt_gap(gap: float) -> str:
    seta = "↑" if gap >= 0 else "↓"
    n = f"{abs(gap):.2f}".replace(".", ",")
    return f"{seta} {n} p.p."


def _situacao_semana(disp: float, disp_min: float, disp_max: float,
                       meta: float) -> str:
    """Classifica a semana de forma relativa ao período + à meta."""
    if disp >= meta:
        return "Acima da meta"
    # dentro do período (sem atingir meta), destaca melhor e pior
    if abs(disp - disp_max) < 0.005:
        return "Melhor semana do período"
    if abs(disp - disp_min) < 0.005:
        return "Pior semana do período"
    return "Abaixo da meta"


def construir_pg2(doc: DocumentT, df: pd.DataFrame) -> None:
    """df vem da query 'disponibilidade_semanal' com colunas:
    SEM, INI, FIM, HORAS_PARADA, DISP_PCT, GAP_PP."""
    add_capitulo(
        doc,
        numero="01",
        titulo="Disponibilidade Fabricante",
        subtitulo=("Responsabilidade SG · descontando eventos sob responsabilidade Serena · "
                   "Meta contratual: 98,00%"),
    )

    if df is None or df.empty:
        from utils.docx_helpers import add_paragrafo
        add_paragrafo(doc, "Dados indisponíveis para esta semana.",
                       size=10, color=HEX_CINZA_CLARO)
        quebra_pagina(doc)
        return

    df = df.sort_values("SEM").reset_index(drop=True)
    # Identifica pior semana
    pior_idx = df["DISP_PCT"].idxmin()
    pior_sem = int(df.loc[pior_idx, "SEM"])
    pior_v   = float(df.loc[pior_idx, "DISP_PCT"])

    # 4 KPI cards: as 3 últimas semanas + pior semana
    cards = []
    for i in range(-3, 0):  # -3, -2, -1 (mais recente)
        if abs(i) <= len(df):
            row = df.iloc[i]
            sem = int(row["SEM"])
            disp = float(row["DISP_PCT"])
            gap = float(row["GAP_PP"])
            sub = _fmt_gap(gap) + (" abaixo da meta" if gap < 0 else " acima da meta")
            cards.append({
                "label": f"Disponibilidade S{sem:02d}",
                "valor": fmt_pct(disp),
                "sub":   sub,
                "cor_rgb": RGB_CORAL,
            })

    pior_gap_abs = abs(float(df.loc[pior_idx, "GAP_PP"]))
    pior_gap_str = f"{pior_gap_abs:.2f}".replace(".", ",")
    cards.append({
        "label": f"Pior semana S{int(df.iloc[0]['SEM']):02d}-S{int(df.iloc[-1]['SEM']):02d}",
        "valor": fmt_pct(pior_v),
        "sub":   f"Gap máx: -{pior_gap_str} p.p. (S{pior_sem:02d})",
        "cor_rgb": RGB_CORAL,
    })
    add_kpi_cards(doc, cards)

    # Gráfico de linha
    img_path = Path(tempfile.gettempdir()) / "g_disp.png"
    grafico_disponibilidade(
        semanas=df["SEM"].astype(int).tolist(),
        valores=df["DISP_PCT"].astype(float).tolist(),
        meta=META,
        out_path=img_path,
    )
    add_imagem(doc, str(img_path), largura_cm=16.5)

    # Tabela histórica
    disp_vals = [float(r["DISP_PCT"]) for _, r in df.iterrows()]
    disp_min, disp_max = min(disp_vals), max(disp_vals)
    rows = []
    for _, r in df.iterrows():
        sem = int(r["SEM"])
        disp = float(r["DISP_PCT"])
        gap = float(r["GAP_PP"])
        rows.append([
            f"S{sem:02d}",
            fmt_pct(disp),
            fmt_pct(META),
            _fmt_gap(gap),
            _situacao_semana(disp, disp_min, disp_max, META),
        ])

    add_tabela(
        doc,
        headers=["Semana", "Disponibilidade", "Meta", "Gap vs. Meta", "Situação"],
        rows=rows,
        col_widths_pct=[0.13, 0.20, 0.15, 0.18, 0.34],
        col_status=4,
        first_col_bold=True,
        titulo=f"Evolução histórica — S{int(df.iloc[0]['SEM']):02d} a S{int(df.iloc[-1]['SEM']):02d}",
    )

    quebra_pagina(doc)