"""Pagina 2 - 01. Disponibilidade Fabricante.

A coluna SEM agora e TEXTO (ex.: "S38" ou "Ultimos 7d"), vinda do periodo.
A ordem das colunas vem da coluna ORD (1..4). A ultima (maior ORD) e o foco.
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


def _situacao_semana(disp: float, disp_min: float, disp_max: float, meta: float) -> str:
    if disp >= meta:
        return "Acima da meta"
    if abs(disp - disp_max) < 0.005:
        return "Melhor período"
    if abs(disp - disp_min) < 0.005:
        return "Pior período"
    return "Abaixo da meta"


def construir_pg2(doc: DocumentT, df: pd.DataFrame) -> None:
    """df da query 'disponibilidade_semanal': ORD, SEM (texto), INI, FIM,
    TX_*, DISP_PCT, GAP_PP, HORAS_PARADA."""
    add_capitulo(
        doc, numero="01", titulo="Disponibilidade Fabricante",
        subtitulo=("Responsabilidade OEM · descontando eventos sob responsabilidade Serena · "
                   "Meta contratual: 98,00%"),
    )

    if df is None or df.empty:
        from utils.docx_helpers import add_paragrafo
        add_paragrafo(doc, "Dados indisponíveis para este período.", size=10, color=HEX_CINZA_CLARO)
        quebra_pagina(doc)
        return

    # ordena pela coluna ORD (fallback: mantém ordem recebida)
    if "ORD" in df.columns:
        df = df.sort_values("ORD").reset_index(drop=True)
    else:
        df = df.reset_index(drop=True)

    def _rot(row) -> str:
        return str(row["SEM"])   # já é "S38" ou "Últimos 7d"

    disp_vals = [float(r["DISP_PCT"]) for _, r in df.iterrows()]
    disp_min, disp_max = min(disp_vals), max(disp_vals)

    pior_idx = df["DISP_PCT"].astype(float).idxmin()
    pior_rot = _rot(df.loc[pior_idx])
    pior_v = float(df.loc[pior_idx, "DISP_PCT"])

    # 3 primeiros cards = as 3 primeiras colunas; 4º card = a coluna em foco (última)
    cards = []
    for i in range(min(3, len(df))):
        row = df.iloc[i]
        disp = float(row["DISP_PCT"]); gap = float(row["GAP_PP"])
        sub = _fmt_gap(gap) + (" abaixo da meta" if gap < 0 else " acima da meta")
        cards.append({"label": f"Disponibilidade {_rot(row)}", "valor": fmt_pct(disp),
                      "sub": sub, "cor_rgb": RGB_CORAL})
    # 4º card: a coluna em foco (a última do df)
    foco = df.iloc[-1]
    disp_f = float(foco["DISP_PCT"]); gap_f = float(foco["GAP_PP"])
    cards.append({
        "label": f"Disponibilidade {_rot(foco)}",
        "valor": fmt_pct(disp_f),
        "sub": _fmt_gap(gap_f) + (" abaixo da meta" if gap_f < 0 else " acima da meta"),
        "cor_rgb": RGB_CORAL,
    })
    add_kpi_cards(doc, cards)

    # Gráfico de linha — passa rótulos-texto
    labels = [_rot(r) for _, r in df.iterrows()]
    img_path = Path(tempfile.gettempdir()) / "g_disp.png"
    grafico_disponibilidade(
        semanas=labels,
        valores=df["DISP_PCT"].astype(float).tolist(),
        meta=META, out_path=img_path,
    )
    add_imagem(doc, str(img_path), largura_cm=16.5)

    # Tabela histórica
    rows = []
    for _, r in df.iterrows():
        disp = float(r["DISP_PCT"]); gap = float(r["GAP_PP"])
        rows.append([_rot(r), fmt_pct(disp), fmt_pct(META), _fmt_gap(gap),
                     _situacao_semana(disp, disp_min, disp_max, META)])
    add_tabela(
        doc,
        headers=["Período", "Disponibilidade", "Meta", "Gap vs. Meta", "Situação"],
        rows=rows, col_widths_pct=[0.16, 0.20, 0.15, 0.18, 0.31],
        col_status=4, first_col_bold=True,
        titulo=f"Evolução — {labels[0]} a {labels[-1]}",
    )
    quebra_pagina(doc)