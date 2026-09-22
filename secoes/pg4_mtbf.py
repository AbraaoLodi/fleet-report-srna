"""Pg 4 — 03. Confiabilidade: MTBF, WTGs críticas e maiores paradas."""
from __future__ import annotations
from utils.docx_helpers import (add_capitulo, add_kpi_cards, add_tabela, quebra_pagina,
    add_paragrafo, RGB_CORAL, HEX_CINZA_CLARO, add_tabela_hierarquica)
from utils.formatters import fmt_float, fmt_date_curta, abrev_subsistema

def _tabela_hierarquica(doc, df, titulo, com_motivo=False):
    """Agrupa por WTG (primeira aparição), mostra alarmes indentados."""
    if df is None or df.empty:
        add_paragrafo(doc, "Sem registros no período.", size=9, color=HEX_CINZA_CLARO); return
    grupos = []
    for wtg in df["AEROGERADOR"].drop_duplicates().tolist():
        sub = df[df["AEROGERADOR"]==wtg]
        h_total = float(sub.iloc[0]["HORAS_TOTAL_WTG"])
        linhas = []
        for _,r in sub.iterrows():
            alarme = str(r["ALARME"])
            if com_motivo:
                mot = str(r.get("MOTIVO","")); tag={"FALHA":"F","MANUTENCAO CORRETIVA NAO PLANEJADA":"CNP",
                    "MANUTENCAO CORRETIVA PLANEJADA":"CP","GRANDE CORRETIVA":"GC",
                    "MANUTENCAO PREVENTIVA":"PV","MANUTENCAO PREDITIVA":"PD"}.get(mot,"")
                alarme = f"[{tag}] {alarme}" if tag else alarme
            linhas.append([alarme, abrev_subsistema(r["SUBSISTEMA"]), str(int(r["OCORRENCIAS"])),
                fmt_date_curta(r["DATA_INICIO"]), fmt_date_curta(r["DATA_FIM"]), fmt_float(float(r["HORAS_ALARME"]))])
        grupos.append({"wtg":wtg,"h_total":h_total,"linhas":linhas})
    add_tabela_hierarquica(doc, grupos,
        headers=["Alarme","Subsistema","Ocor.","Início","Fim","Horas"],
        col_widths_pct=[0.38,0.20,0.09,0.11,0.11,0.11], titulo=titulo)

def construir_pg4(doc, df_mtbf, df_falha, df_corretiva):
    add_capitulo(doc, "03", "Confiabilidade & WTGs Críticas",
        "MTBF por aerogerador (snapshot oficial) e maiores tempos de parada por falha da semana")
    # Tabela de MTBF
    if df_mtbf is not None and not df_mtbf.empty:
        d = df_mtbf.sort_values("MTBF_7D")
        rows = [[str(r["WTG"]), fmt_float(float(r["MTBF_7D"]))+" h", str(int(r["N_FALHAS_7D"])),
                 fmt_float(float(r["H_FALHAS_7D"]))+" h"] for _,r in d.iterrows()]
        add_tabela(doc, ["Aerogerador","MTBF (7 dias)","Falhas (7d)","Horas paradas"], rows,
            col_widths_pct=[0.30,0.26,0.22,0.22], col_destaque_coral=1, first_col_bold=True,
            titulo="WTGs críticas por MTBF — menor tempo entre falhas")

    # Maiores tempos por falha (hierárquica) — logo abaixo do MTBF, mesma página
    _tabela_hierarquica(doc, df_falha,
        "Maiores tempos de parada por FALHA — detalhamento por aerogerador → alarme (Top 10)")
    quebra_pagina(doc)
    # Maiores tempos por corretiva (hierárquica)
    add_capitulo(doc, "04", "Maiores Tempos de Parada",
        "Aerogeradores com maior tempo de parada na semana — todos os motivos de manutenção (Top 10)")
    _tabela_hierarquica(doc, df_corretiva, "Detalhamento por aerogerador → alarme", com_motivo=True)
    quebra_pagina(doc)
