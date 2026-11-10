"""Pg 3 — 02. Principais Alarmes de Falha + evolução de taxas."""
from __future__ import annotations
import tempfile
from pathlib import Path
from utils.charts import grafico_taxas_3series
from utils.docx_helpers import (add_capitulo, add_kpi_cards, add_imagem, add_tabela,
    quebra_pagina, add_paragrafo, RGB_CORAL, HEX_CINZA_CLARO)
from utils.formatters import fmt_float, abrev_subsistema


def _ordena(df):
    """Ordena pelas colunas na ordem certa (ORD se existir, senão mantém)."""
    if df is None or df.empty:
        return df
    if "ORD" in df.columns:
        return df.sort_values("ORD").reset_index(drop=True)
    return df.reset_index(drop=True)


def construir_pg3(doc, df_taxas, df_alarmes):
    add_capitulo(doc, "02", "Perfil de Manutenção",
        "Composição das taxas de indisponibilidade e principais alarmes de falha do período")

    # KPI cards das 4 taxas do período em foco (última coluna)
    if df_taxas is not None and not df_taxas.empty:
        d = _ordena(df_taxas)
        atual = d.iloc[-1]
        rot = str(atual["SEM"])   # já é "S38" ou "Últimos 7d"
        cards = [
            {"label": f"Taxa de Falha {rot}", "valor": fmt_float(float(atual["TX_FALHA"]))+"%", "sub": "% horas-frota", "cor_rgb": RGB_CORAL},
            {"label": f"Corretiva Não Planejada {rot}", "valor": fmt_float(float(atual["TX_CORR_NP"]))+"%", "sub": "% horas-frota", "cor_rgb": RGB_CORAL},
            {"label": f"Corretiva Planejada {rot}", "valor": fmt_float(float(atual["TX_CORR_P"]))+"%", "sub": "% horas-frota", "cor_rgb": RGB_CORAL},
            {"label": f"Grande Corretiva {rot}", "valor": fmt_float(float(atual["TX_GRANDE_CORR"]))+"%", "sub": "% horas-frota", "cor_rgb": RGB_CORAL},
        ]
        add_kpi_cards(doc, cards)

    # Gráfico de taxas (4 colunas) — passa rótulos-texto direto
    if df_taxas is not None and not df_taxas.empty:
        d = _ordena(df_taxas)
        img = Path(tempfile.gettempdir())/"g_taxas.png"
        grafico_taxas_3series(
            d["SEM"].tolist(), d["TX_FALHA"].astype(float).tolist(),
            d["TX_CORR_NP"].astype(float).tolist(), d["TX_CORR_P"].astype(float).tolist(),
            d["TX_GRANDE_CORR"].astype(float).tolist(), str(img))
        add_imagem(doc, str(img), 16.8)

    # Top alarmes de falha
    if df_alarmes is not None and not df_alarmes.empty:
        rows = []
        for _, r in df_alarmes.iterrows():
            rows.append([abrev_subsistema(r["SUBSISTEMA"]), str(r["ALARME"]),
                str(int(r["OCORRENCIAS"])), str(int(r["WTGS_AFETADAS"])),
                fmt_float(float(r["HORAS"]))+" h"])
        add_tabela(doc, ["Subsistema", "Alarme", "Ocor.", "WTGs", "Horas"], rows,
            col_widths_pct=[0.17, 0.47, 0.09, 0.09, 0.18], col_destaque_coral=4, first_col_bold=False,
            titulo="Top 10 alarmes por horas de falha — período em foco")
    else:
        add_paragrafo(doc, "Nenhum alarme de falha registrado no período.", size=10, color=HEX_CINZA_CLARO)
    quebra_pagina(doc)