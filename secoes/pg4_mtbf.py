"""Página 4 — 03. MTBF e WTGs Críticas.

Conteúdo:
  - Banner amarelo de nota metodológica (fonte oficial F_REAL_WTG_MTBF_PERIOD)
  - Tabela MTBF por WTG (top 7, snapshot mais recente, fonte oficial)
  - Tabela WTGs Críticas (semana atual SG)
  - Tabela hierárquica Maiores Tempos por FALHA (top 20 WTGs com alarmes)
  - Tabela hierárquica Maiores Tempos por CORRETIVA (top 10 WTGs com alarmes)

MTTR foi removido por decisão do usuário.
"""
from __future__ import annotations

import pandas as pd
from docx.document import Document as DocumentT

from utils.docx_helpers import (
    HEX_AMARELO, HEX_CINZA_CLARO,
    add_capitulo, add_banner, add_tabela, add_paragrafo, quebra_pagina,
)
from utils.formatters import abrev_subsistema
from utils.logger import get_logger

log = get_logger(__name__)


def _status_mtbf(mtbf_h: float) -> str:
    """MTBF baixo = falha frequente."""
    if mtbf_h < 24:
        return "Crítico"
    if mtbf_h < 72:
        return "Atenção"
    return "Normal"


def _status_wtg_critica(downtime_h: float, n_eventos: int) -> str:
    if downtime_h >= 20 or n_eventos >= 20:
        return "Crítico"
    if downtime_h >= 10 or n_eventos >= 10:
        return "Atenção"
    return "Monitorar"


def _fmt_h(v) -> str:
    if v is None or pd.isna(v):
        return "—"
    return f"{float(v):.2f}h".replace(".", ",")


def _fmt_h1(v) -> str:
    if v is None or pd.isna(v):
        return "—"
    return f"{float(v):.1f}h".replace(".", ",")


def _hhmmss_from_hours(h: float) -> str:
    """Converte float horas em 'HH:MM:SS'."""
    if h is None or pd.isna(h):
        return "—"
    total_s = int(round(float(h) * 3600))
    hh = total_s // 3600
    mm = (total_s % 3600) // 60
    ss = total_s % 60
    return f"{hh:02d}:{mm:02d}:{ss:02d}"


def construir_pg4(doc: DocumentT,
                    df_mtbf_wtg: pd.DataFrame,
                    df_wtgs: pd.DataFrame,
                    df_maiores_falha: pd.DataFrame,
                    df_maiores_corretiva: pd.DataFrame,
                    semana_ini: int, semana_fim: int) -> None:
    """Constrói a página 4 com MTBF + WTGs Críticas + Maiores Tempos hierárquicas."""
    add_capitulo(
        doc,
        numero="03",
        titulo="MTBF e WTGs com Maiores Tempos em Falha",
        subtitulo=(f"MTBF por máquina (snapshot oficial) · WTGs com Maiores "
                   f"Tempos de Parada — S{semana_fim:02d}"),
    )

    snapshot_str = "snapshot mais recente disponível"
    if df_mtbf_wtg is not None and not df_mtbf_wtg.empty and "SNAPSHOT_DATE" in df_mtbf_wtg.columns:
        try:
            snapshot_str = f"snapshot de {df_mtbf_wtg.iloc[0]['SNAPSHOT_DATE']}"
        except Exception:
            pass

    add_banner(
        doc,
        f"MTBF calculado pelo time de dados Serena (fonte: F_REAL_WTG_MTBF_PERIOD, "
        f"{snapshot_str}). Top 7 WTGs com menor MTBF 7d — atualizado diariamente.",
        cor=HEX_AMARELO,
    )

    # ---- Tabela MTBF por WTG ----
    if df_mtbf_wtg is None or df_mtbf_wtg.empty:
        add_paragrafo(doc, "Sem dados de MTBF disponíveis no snapshot.",
                       size=9, color=HEX_CINZA_CLARO)
    else:
        df_mtbf_wtg = df_mtbf_wtg.sort_values("MTBF_7D", ascending=True).reset_index(drop=True)
        rows = []
        for _, r in df_mtbf_wtg.iterrows():
            wtg = str(r["WTG"])
            mtbf_7d = float(r["MTBF_7D"]) if pd.notna(r.get("MTBF_7D")) else 0
            n_fal_7d = int(r["N_FALHAS_7D"]) if pd.notna(r.get("N_FALHAS_7D")) else 0
            h_fal_7d = float(r["H_FALHAS_7D"]) if pd.notna(r.get("H_FALHAS_7D")) else 0
            rows.append([
                wtg,
                _fmt_h1(mtbf_7d),
                str(n_fal_7d),
                _fmt_h1(h_fal_7d),
                _status_mtbf(mtbf_7d),
            ])
        add_tabela(
            doc,
            headers=["WTG", "MTBF 7d", "Falhas 7d", "Horas Falha 7d", "Status"],
            rows=rows,
            col_widths_pct=[0.16, 0.18, 0.18, 0.24, 0.24],
            col_status=4,
            first_col_bold=True,
            titulo="MTBF por WTG — Top 7 com menor MTBF (maior frequência de falhas)",
        )

    # ---- Maiores Tempos por FALHA (hierárquica) ----
    if df_maiores_falha is None or df_maiores_falha.empty:
        add_paragrafo(doc, "Sem registros de falha na semana.",
                       size=9, color=HEX_CINZA_CLARO)
    else:
        rows = _hierarquica_rows(df_maiores_falha, com_motivo=False)
        add_tabela(
            doc,
            headers=["WTG", "Subsistema", "Alarme", "Ocor.",
                     "Início", "Fim", "Tempo (h)", "Tempo Total"],
            rows=rows,
            col_widths_pct=[0.09, 0.15, 0.31, 0.07, 0.08, 0.08, 0.11, 0.11],
            first_col_bold=True,
            titulo=f"WTGs com Maiores Tempos em Falha — S{semana_fim:02d} (top 20 WTGs)",
        )

    # ---- Maiores Tempos por CORRETIVA (hierárquica) ----
    if df_maiores_corretiva is None or df_maiores_corretiva.empty:
        add_paragrafo(doc, "Sem registros de corretiva na semana.",
                       size=9, color=HEX_CINZA_CLARO)
    else:
        rows = _hierarquica_rows(df_maiores_corretiva, com_motivo=True)
        add_tabela(
            doc,
            headers=["WTG", "Subsistema", "Motivo", "Alarme",
                     "Ocor.", "Início", "Fim", "Tempo (h)", "Tempo Total"],
            rows=rows,
            col_widths_pct=[0.10, 0.12, 0.12, 0.24, 0.07, 0.07, 0.07, 0.10, 0.11],
            first_col_bold=True,
            titulo=f"WTGs com Maiores Tempos de Parada — S{semana_fim:02d} (top 10 WTGs)",
        )

    quebra_pagina(doc)


def _hierarquica_rows(df: pd.DataFrame, com_motivo: bool) -> list[list[str]]:
    """Constrói linhas mostrando WTG só na primeira ocorrência (hierárquica)."""
    rows = []
    df_sorted = df.sort_values(
        ["HORAS_TOTAL_WTG", "AEROGERADOR", "HORAS_ALARME"],
        ascending=[False, True, False],
    ).reset_index(drop=True)

    wtg_anterior = None
    for _, r in df_sorted.iterrows():
        wtg = str(r["AEROGERADOR"])
        sub = abrev_subsistema(r.get("SUBSISTEMA"))
        alarme = str(r["ALARME"]) if pd.notna(r.get("ALARME")) else "—"
        ocor = int(r["OCORRENCIAS"])
        data_ini = str(r["DATA_INICIO"]) if pd.notna(r.get("DATA_INICIO")) else "—"
        data_fim = str(r["DATA_FIM"]) if pd.notna(r.get("DATA_FIM")) else "—"
        horas_a = float(r["HORAS_ALARME"]) if pd.notna(r.get("HORAS_ALARME")) else 0
        horas_t = float(r["HORAS_TOTAL_WTG"]) if pd.notna(r.get("HORAS_TOTAL_WTG")) else 0

        # Mostra WTG e total só na 1ª linha; demais linhas ficam em branco
        wtg_show = wtg if wtg != wtg_anterior else ""
        total_show = _hhmmss_from_hours(horas_t) if wtg != wtg_anterior else ""

        # Formato de data dd/mm
        def _fmt_data(s: str) -> str:
            try:
                return pd.to_datetime(s).strftime("%d/%m")
            except Exception:
                return s

        if com_motivo:
            motivo = str(r["MOTIVO"]) if pd.notna(r.get("MOTIVO")) else "—"
            # Normaliza nomes longos
            motivo_short = (
                motivo.replace("MANUTENCAO ", "M. ")
                      .replace("CORRETIVA NAO PLANEJADA", "Corr NP")
                      .replace("CORRETIVA PLANEJADA", "Corr P")
                      .replace("GRANDE CORRETIVA", "Grande Corr")
                      .title()
                      .replace("M. ", "")
            )
            rows.append([
                wtg_show, sub if wtg != wtg_anterior else "", motivo_short,
                alarme, str(ocor),
                _fmt_data(data_ini), _fmt_data(data_fim),
                _hhmmss_from_hours(horas_a),
                total_show,
            ])
        else:
            rows.append([
                wtg_show, sub if wtg != wtg_anterior else "",
                alarme, str(ocor),
                _fmt_data(data_ini), _fmt_data(data_fim),
                _hhmmss_from_hours(horas_a),
                total_show,
            ])
        wtg_anterior = wtg

    return rows