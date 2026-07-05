"""Página 3 — 02. Taxa de Falhas e Corretivas.

Conteúdo:
  - 4 KPI cards: Falha · Corr Não Planejada · Corr Planejada · Grande Corretiva
  - Gráfico de 4 séries (4 semanas), com porcentagem anotada em cada ponto
  - Tabela Top 10 Alarmes por horas de falha (semana atual)
"""
from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
from docx.document import Document as DocumentT

from utils.charts import grafico_taxas_3series
from utils.docx_helpers import (
    HEX_CORAL, HEX_VERDE, HEX_GRAFITE, HEX_CINZA_CLARO, HEX_AMARELO,
    RGB_CORAL, RGB_VERDE, RGB_GRAFITE_ESC, RGB_AMARELO,
    add_capitulo, add_kpi_cards, add_imagem, add_tabela, quebra_pagina,
    add_paragrafo,
)
from utils.formatters import fmt_pct, abrev_subsistema
from utils.logger import get_logger

log = get_logger(__name__)


def _risco_alarme(ocorrencias: int) -> str:
    if ocorrencias >= 50:
        return "Alto"
    if ocorrencias >= 10:
        return "Médio"
    return "Baixo"


def construir_pg3(doc: DocumentT, df_disp: pd.DataFrame,
                    df_alarmes: pd.DataFrame, sem_atual: int) -> None:
    """df_disp: SEM, INI, FIM, TX_FALHA, TX_CORR_NP, TX_CORR_P,
               TX_GRANDE_CORR, DISP_PCT, GAP_PP, HORAS_PARADA (4 semanas).
    df_alarmes: ALARME, SUBSISTEMA, OCORRENCIAS, WTGS_AFETADAS, HORAS
                (semana atual)."""
    add_capitulo(
        doc,
        numero="02",
        titulo="Taxa de Falhas e Corretivas",
        subtitulo=("Decomposição da indisponibilidade: Falha · Corretiva Não Planejada · "
                   "Corretiva Planejada · Grande Corretiva"),
    )

    if df_disp is None or df_disp.empty:
        add_paragrafo(doc, "Dados indisponíveis.", size=10, color=HEX_CINZA_CLARO)
        quebra_pagina(doc)
        return

    df_disp = df_disp.sort_values("SEM").reset_index(drop=True)
    last = df_disp.iloc[-1]
    prev = df_disp.iloc[-2] if len(df_disp) >= 2 else None

    tx_falha = float(last["TX_FALHA"])
    tx_np    = float(last["TX_CORR_NP"])
    tx_p     = float(last["TX_CORR_P"])
    tx_gc    = float(last["TX_GRANDE_CORR"])

    def _delta_sub(atual: float, prev_v: float | None) -> str:
        if prev_v is None:
            return "—"
        diff = atual - prev_v
        seta = "↑" if diff > 0 else "↓"
        prev_str = f"{prev_v:.2f}".replace(".", ",")
        return (f"{seta} {'redução' if diff < 0 else 'aumento'} vs "
                f"S{int(prev['SEM']):02d} ({prev_str}%)")

    prev_falha = float(prev["TX_FALHA"]) if prev is not None else None
    prev_np    = float(prev["TX_CORR_NP"]) if prev is not None else None
    prev_p     = float(prev["TX_CORR_P"]) if prev is not None else None
    prev_gc    = float(prev["TX_GRANDE_CORR"]) if prev is not None else None

    cards = [
        {"label": f"Taxa de Falha — S{sem_atual:02d}",
         "valor": fmt_pct(tx_falha),
         "sub":   _delta_sub(tx_falha, prev_falha),
         "cor_rgb": RGB_GRAFITE_ESC},
        {"label": f"Corretiva Não Planejada S{sem_atual:02d}",
         "valor": fmt_pct(tx_np),
         "sub":   _delta_sub(tx_np, prev_np),
         "cor_rgb": RGB_CORAL},
        {"label": f"Corretiva Planejada S{sem_atual:02d}",
         "valor": fmt_pct(tx_p),
         "sub":   _delta_sub(tx_p, prev_p),
         "cor_rgb": RGB_VERDE},
        {"label": f"Grande Corretiva S{sem_atual:02d}",
         "valor": fmt_pct(tx_gc),
         "sub":   _delta_sub(tx_gc, prev_gc),
         "cor_rgb": RGB_AMARELO},
    ]
    add_kpi_cards(doc, cards)

    # Gráfico 4 séries
    img_path = Path(tempfile.gettempdir()) / "g_taxas.png"
    semanas_lbl = df_disp["SEM"].astype(int).tolist()
    grafico_taxas_3series(
        semanas=semanas_lbl,
        falha=df_disp["TX_FALHA"].astype(float).tolist(),
        corr_np=df_disp["TX_CORR_NP"].astype(float).tolist(),
        corr_p=df_disp["TX_CORR_P"].astype(float).tolist(),
        grande_corr=df_disp["TX_GRANDE_CORR"].astype(float).tolist(),
        out_path=img_path,
        titulo=f"TAXA DE FALHA X CORRETIVA — S{semanas_lbl[0]:02d} A S{semanas_lbl[-1]:02d}",
    )
    add_imagem(doc, str(img_path), largura_cm=16.5)

    # Tabela Top 10 Alarmes
    if df_alarmes is None or df_alarmes.empty:
        add_paragrafo(doc, "Sem alarmes de falha registrados na semana.",
                       size=9, color=HEX_CINZA_CLARO)
    else:
        df_alarmes = df_alarmes.sort_values("HORAS", ascending=False).reset_index(drop=True)
        rows = []
        for _, r in df_alarmes.iterrows():
            alarme = str(r["ALARME"])
            sub = abrev_subsistema(r.get("SUBSISTEMA"))
            ocor = int(r["OCORRENCIAS"])
            wtgs_afetadas = int(r["WTGS_AFETADAS"])
            horas = float(r["HORAS"])
            # Quebrar alarme em código + descrição
            if "_" in alarme and alarme.split("_")[0].isdigit():
                cod = alarme.split("_")[0]
                desc = alarme.split("_", 1)[1].replace("_", " ")
            else:
                cod, desc = "—", alarme
            horas_str = f"{horas:.2f}h".replace(".", ",")
            rows.append([
                cod, desc, sub,
                str(ocor), str(wtgs_afetadas), horas_str,
            ])
        add_tabela(
            doc,
            headers=["Código", "Descrição", "Subsistema",
                     "Ocor.", "WTGs", "Horas"],
            rows=rows,
            col_widths_pct=[0.09, 0.37, 0.22, 0.09, 0.08, 0.15],
            first_col_bold=True,
            titulo=f"Falhas com maiores tempos de parada — S{sem_atual:02d}",
        )

    quebra_pagina(doc)