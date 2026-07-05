"""Cálculos auxiliares — padrão Serena.

Para o Fleet Report SRNA, a maioria das agregações é feita NO Snowflake
(eventlog tem milhões de linhas). Este módulo contém apenas utilitários
de conversão/cálculo em pandas, usados em pós-processamento dos DataFrames.
"""
from __future__ import annotations

import pandas as pd


def disp_a_partir_de_taxas(falha: float | None,
                            corr_np: float | None,
                            corr_p: float | None,
                            grande_corr: float | None) -> float | None:
    """Disp Possível = 100 - (Falha + CorrNP + CorrP + GC).

    Fórmula da engenheira O&M, validada à vírgula em S22-S24.
    """
    valores = [falha, corr_np, corr_p, grande_corr]
    if any(v is None for v in valores):
        return None
    return 100.0 - sum(valores)


def soma_segura(df: pd.DataFrame, col: str) -> float:
    """Soma uma coluna ignorando NaN. Devolve 0.0 se vazio."""
    if df.empty or col not in df.columns:
        return 0.0
    return float(df[col].fillna(0).sum())
