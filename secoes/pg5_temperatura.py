"""Pg 5 — 05. Monitoramento de Temperatura (ML + SCADA).

Classifica cada sensor pelo mapa da frota (data/mapa_temperatura.json):
CRITICO (componente de desgaste) vai para a tabela hierárquica principal;
AUX (refrigeração/ambiente) vai para uma nota compacta abaixo.
Os limites vêm dinamicamente da query (por frota), não hardcoded.
"""
from __future__ import annotations
import json
from pathlib import Path
import pandas as pd
from utils.docx_helpers import (add_capitulo, add_tabela_hierarquica_temp, quebra_pagina,
    add_paragrafo, add_banner, HEX_CINZA_CLARO)
from utils.formatters import fmt_float

_MAPA_PATH = Path(__file__).resolve().parent.parent / "data" / "mapa_temperatura.json"
try:
    _MAPA_TEMP = json.load(open(_MAPA_PATH, encoding="utf-8"))
except Exception:
    _MAPA_TEMP = {}

def _ordem(rotulo):
    r = rotulo.lower()
    if "enrol" in r and "gerador" in r: return 0
    if "rolam" in r and "gerador" in r: return 1
    if "slipring" in r: return 2
    if "rolamento principal" in r: return 3
    if "gearbox" in r: return 4
    if "hidráulic" in r: return 5
    if "trafo" in r: return 6
    if "pitch" in r: return 7
    return 9

def _sev_txt(sev):
    return {2:"Crítico",1:"Atenção"}.get(int(sev),"Normal")

def _classificar(frota_key, variable):
    mapa = _MAPA_TEMP.get(frota_key, {})
    if variable in mapa:
        rot, cat = mapa[variable]
        return rot, cat
    return variable.replace("_TEMP","").replace("_"," ").title(), "AUX"

def construir_pg5(doc, df_temp, periodo_str="", frota_key="SG"):
    add_capitulo(doc, "05", "Monitoramento de Temperatura",
        f"Sensores com pico acima dos limites do fabricante · janela {periodo_str} · \u26a0 alerta (derate) · \u26d4 crítico (downtime)")
    if df_temp is None or df_temp.empty:
        add_paragrafo(doc, "Nenhum sensor acima dos limites no período.", size=10, color=HEX_CINZA_CLARO)
        quebra_pagina(doc); return

    df = df_temp.copy()
    df.columns = [c.upper() for c in df.columns]
    rotulos, cats = [], []
    for _, r in df.iterrows():
        rot, cat = _classificar(frota_key, str(r["VARIABLE"]))
        rotulos.append(rot); cats.append(cat)
    df["ROTULO"] = rotulos
    df["CATEGORIA"] = cats
    df["ORD"] = df["ROTULO"].map(_ordem)

    criticos = df[df["CATEGORIA"] == "CRITICO"].copy()
    auxiliares = df[df["CATEGORIA"] == "AUX"].copy()

    if not criticos.empty:
        rank = (criticos.assign(_c=(criticos["SEV"]==2).astype(int))
                .groupby("WTG").agg(n_crit=("_c","sum"), pico_max=("PICO","max"))
                .reset_index().sort_values(["n_crit","pico_max"], ascending=[False,False]))
        rank = rank.head(7)
        grupos = []
        for wtg in rank["WTG"]:
            sub = criticos[criticos["WTG"]==wtg].sort_values(["ORD","PICO"], ascending=[True,False])
            pico_max = float(sub["PICO"].max())
            linhas = []
            for _,r in sub.iterrows():
                media = fmt_float(float(r["MEDIA"]))+" °C" if pd.notna(r["MEDIA"]) else "—"
                der = int(r["DERATE"]) if pd.notna(r["DERATE"]) else "·"
                dow = int(r["DOWNTIME"]) if pd.notna(r["DOWNTIME"]) else "·"
                linhas.append([
                    str(r["ROTULO"]),
                    fmt_float(float(r["PICO"]))+" °C",
                    media,
                    f"{der}/{dow} °C",
                    _sev_txt(r["SEV"]),
                ])
            grupos.append({"wtg":wtg, "h_total":pico_max, "linhas":linhas})
        add_tabela_hierarquica_temp(
            doc, grupos,
            headers=["Componente","Pico","Média (7d)","Limites \u26a0/\u26d4","Severidade"],
            col_widths_pct=[0.30,0.15,0.16,0.20,0.19],
            col_status=4,
            titulo=f"Top {len(rank)} aerogeradores por severidade térmica — componentes críticos")
    else:
        add_paragrafo(doc, "Nenhum componente crítico acima dos limites no período.",
            size=10, color=HEX_CINZA_CLARO)

    if not auxiliares.empty:
        n_aux = len(auxiliares)
        n_wtg_aux = auxiliares["WTG"].nunique()
        distintos = sorted(auxiliares["ROTULO"].unique())
        resumo = ", ".join(distintos[:5]) + (f" +{len(distintos)-5}" if len(distintos)>5 else "")
        add_paragrafo(doc,
            f"Sensores auxiliares (refrigeração/ambiente) acima do limite: {n_aux} leitura(s) "
            f"em {n_wtg_aux} WTG(s) — {resumo}. Oscilação típica desses sensores; monitorar sem criticidade.",
            size=8, color=HEX_CINZA_CLARO, before=6)

    add_banner(doc, "Picos acima do limite de downtime (\u26d4) nos componentes de desgaste indicam risco de "
        "derate/parada. Recomenda-se inspeção do sistema de refrigeração e dos componentes nas WTGs críticas.",
        cor="F59E0B")
    quebra_pagina(doc)
