"""Pg 3 — 02. Principais Alarmes de Falha + evolução de taxas."""
from __future__ import annotations
import tempfile
from pathlib import Path
from utils.charts import grafico_taxas_3series
from utils.docx_helpers import (add_capitulo, add_kpi_cards, add_imagem, add_tabela,
    quebra_pagina, add_paragrafo, RGB_CORAL, HEX_CINZA_CLARO)
from utils.formatters import fmt_float, abrev_subsistema

def construir_pg3(doc, df_taxas, df_alarmes):
    add_capitulo(doc, "02", "Perfil de Manutenção",
        "Composição das taxas de indisponibilidade e principais alarmes de falha da semana")
    # KPI cards das 4 taxas da semana atual (última linha)
    if df_taxas is not None and not df_taxas.empty:
        d = df_taxas.sort_values("SEM")
        atual = d.iloc[-1]
        sem_atual = int(atual["SEM"])
        cards = [
            {"label":f"Taxa de Falha S{sem_atual:02d}","valor":fmt_float(float(atual["TX_FALHA"]))+"%","sub":"% horas-frota","cor_rgb":RGB_CORAL},
            {"label":f"Corretiva Não Planejada S{sem_atual:02d}","valor":fmt_float(float(atual["TX_CORR_NP"]))+"%","sub":"% horas-frota","cor_rgb":RGB_CORAL},
            {"label":f"Corretiva Planejada S{sem_atual:02d}","valor":fmt_float(float(atual["TX_CORR_P"]))+"%","sub":"% horas-frota","cor_rgb":RGB_CORAL},
            {"label":f"Grande Corretiva S{sem_atual:02d}","valor":fmt_float(float(atual["TX_GRANDE_CORR"]))+"%","sub":"% horas-frota","cor_rgb":RGB_CORAL},
        ]
        add_kpi_cards(doc, cards)
    # Gráfico de taxas (4 semanas)
    if df_taxas is not None and not df_taxas.empty:
        d = df_taxas.sort_values("SEM")
        img = Path(tempfile.gettempdir())/"g_taxas.png"
        grafico_taxas_3series(
            d["SEM"].astype(int).tolist(), d["TX_FALHA"].astype(float).tolist(),
            d["TX_CORR_NP"].astype(float).tolist(), d["TX_CORR_P"].astype(float).tolist(),
            d["TX_GRANDE_CORR"].astype(float).tolist(), str(img))
        add_imagem(doc, str(img), 16.8)
    # Top alarmes de falha
    if df_alarmes is not None and not df_alarmes.empty:
        rows = []
        for _,r in df_alarmes.iterrows():
            rows.append([abrev_subsistema(r["SUBSISTEMA"]), str(r["ALARME"]),
                str(int(r["OCORRENCIAS"])), str(int(r["WTGS_AFETADAS"])),
                fmt_float(float(r["HORAS"]))+" h"])
        add_tabela(doc, ["Subsistema","Alarme","Ocor.","WTGs","Horas"], rows,
            col_widths_pct=[0.17,0.47,0.09,0.09,0.18], col_destaque_coral=4, first_col_bold=False,
            titulo="Top 10 alarmes por horas de falha — semana atual")
    else:
        add_paragrafo(doc, "Nenhum alarme de falha registrado na semana.", size=10, color=HEX_CINZA_CLARO)
    quebra_pagina(doc)
