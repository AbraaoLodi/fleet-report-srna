"""Pg 5 — 04. Análise de Temperatura.

Lista as Top 5 WTGs que cruzaram algum limite de temperatura na semana
(últimos 7 dias), detalhando TODOS os componentes quentes de cada uma
(gerador enrolamento/rolamento/slipring, gearbox, conversor, trafo, nacele).

Fonte dos limites (oficial, mesma do cockpit):
    OEM_PROD.MACHINE_LEARNING.F_REAL_EOL_WTG_TEMPERATURE_ANOMALY_EVENTS_DAILY
      · DERATE_LIMIT   -> limite amarelo (alerta / derate)
      · DOWNTIME_LIMIT -> limite vermelho (crítico / parada)
      · MAX_TEMP       -> pico diário

Pico (MAX) e limites vêm da tabela ML; a média (AVG) vem do SCADA_RAW.

Classificação por sensor:
    pico >= downtime            -> CRÍTICO   (vermelho)
    derate <= pico < downtime   -> ALERTA    (amarelo)
    pico < derate               -> NORMAL    (não listado)
    sensor inválido (>50% ruim) -> SENSOR C/ FALHA

Regras de dado:
    - leitura inválida: VALUE <= 0 ou VALUE >= 200
    - sempre o UPLOADED_AT mais recente (CTE last_upload)
"""
from __future__ import annotations

import pandas as pd

from utils.docx_helpers import (
    HEX_CORAL, HEX_AMARELO, HEX_VERDE, HEX_CINZA_CLARO, HEX_GRAFITE,
    add_capitulo, add_kpi_cards, add_tabela, add_banner, add_paragrafo,
)
from utils.logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Rótulos de status (batem com STATUS_COLORS do docx_helpers)
# ---------------------------------------------------------------------------
ST_CRITICO = "Crítico"     # -> coral bold
ST_ALERTA  = "Atenção"     # -> amarelo bold
ST_FALHA   = "Sensor c/ falha"

# ordem de exibição dos componentes dentro de cada WTG
ORDEM_COMPONENTE = {
    "Gerador enrol.":    0,
    "Gerador rolam.":    1,
    "Gerador slipring":  2,
    "Gearbox":           3,
    "Conversor":         4,
    "Trafo":             5,
    "Nacele":            6,
    "Outro":             9,
}


def _fmt_temp(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{float(v):.1f}".replace(".", ",") + "°C"


def _fmt_limite(derate, downtime) -> str:
    """Ex.: '130 / 135' (amarelo / vermelho)."""
    def _n(x):
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return "—"
        return f"{int(round(float(x)))}"
    return f"{_n(derate)} / {_n(downtime)}"


def _status_label(sev: int) -> str:
    if sev >= 2:
        return ST_CRITICO
    if sev == 1:
        return ST_ALERTA
    return "Normal"


def construir_pg5(doc, df_temp: pd.DataFrame, periodo_str: str) -> None:
    """Constrói a página de Análise de Temperatura.

    df_temp — uma linha por WTG × componente que cruzou limite, colunas:
        WTG, COMPONENTE, PICO, MEDIA, DERATE, DOWNTIME, SEV, N_CRIT, N_ALERTA
      (SEV: 2=crítico, 1=alerta; N_CRIT/N_ALERTA são por WTG, p/ ranking)
    """
    add_capitulo(
        doc, "04", "Análise de Temperatura",
        f"WTGs com sensores acima do limite · {periodo_str}",
    )

    if df_temp is None or df_temp.empty:
        add_banner(
            doc,
            "Nenhuma WTG cruzou limite de temperatura no período. "
            "Todos os sensores dentro da faixa normal (abaixo do derate).",
            cor=HEX_VERDE,
        )
        _rodape_metodologia(doc)
        return

    df = df_temp.copy()
    df.columns = [c.upper() for c in df.columns]

    # ----- KPIs do topo -----
    n_wtgs = df["WTG"].nunique()
    n_criticos = int((df["SEV"] >= 2).sum())
    n_alertas = int((df["SEV"] == 1).sum())

    # componente que mais aparece como crítico (curto p/ caber no card)
    crit = df[df["SEV"] >= 2]
    if not crit.empty:
        comp_full = crit["COMPONENTE"].value_counts().idxmax()
        # primeira palavra ("Gerador enrol." -> "Gerador") p/ caber na fonte grande
        comp_dom = comp_full.split()[0]
    else:
        comp_dom = "—"

    # pior WTG (maior pico crítico)
    if not crit.empty:
        pior = crit.loc[crit["PICO"].idxmax()]
        pior_valor = _fmt_temp(pior["PICO"])
        pior_sub = f"{pior['WTG']} · maior pico crítico"
    else:
        pior_valor = "—"
        pior_sub = "sem pico crítico"

    add_kpi_cards(doc, [
        {"label": "WTGs com alerta/crítico", "valor": str(n_wtgs),
         "sub": f"no período (top {n_wtgs} exibidas)"},
        {"label": "Sensores críticos", "valor": str(n_criticos),
         "sub": f"{n_alertas} em alerta (amarelo)"},
        {"label": "Componente dominante", "valor": comp_dom,
         "sub": "mais recorrente nos críticos"},
        {"label": "Pior pico", "valor": pior_valor,
         "sub": pior_sub},
    ])

    # ----- Ordena: WTG por ranking (n_crit desc), componente por ordem física -----
    df["ORD_COMP"] = df["COMPONENTE"].map(ORDEM_COMPONENTE).fillna(9)
    # chave de ranking da WTG: usa N_CRIT e maior pico
    rank = (df.groupby("WTG")
              .agg(n_crit=("N_CRIT", "max"),
                   n_alerta=("N_ALERTA", "max"),
                   pico_max=("PICO", "max"))
              .reset_index()
              .sort_values(["n_crit", "pico_max"], ascending=[False, False]))
    ordem_wtg = {w: i for i, w in enumerate(rank["WTG"])}
    df["ORD_WTG"] = df["WTG"].map(ordem_wtg)
    df = df.sort_values(["ORD_WTG", "ORD_COMP"]).reset_index(drop=True)

    # ----- Monta linhas da tabela (WTG só aparece na 1ª linha do grupo) -----
    headers = ["WTG", "Componente", "Pico", "Média", "Limite (⚠/⛔)", "Status"]
    col_widths = [0.13, 0.26, 0.13, 0.13, 0.18, 0.17]
    rows: list[list[str]] = []
    wtg_ant = None
    for _, r in df.iterrows():
        wtg = str(r["WTG"])
        mostra_wtg = wtg if wtg != wtg_ant else ""
        rows.append([
            mostra_wtg,
            str(r["COMPONENTE"]),
            _fmt_temp(r["PICO"]),
            _fmt_temp(r["MEDIA"]),
            _fmt_limite(r.get("DERATE"), r.get("DOWNTIME")),
            _status_label(int(r["SEV"])),
        ])
        wtg_ant = wtg

    add_tabela(
        doc,
        headers=headers,
        rows=rows,
        col_widths_pct=col_widths,
        col_status=5,          # coluna Status recebe cor automática
        first_col_bold=True,
        titulo=f"Top {n_wtgs} WTGs com maiores temperaturas — {periodo_str}",
    )

    # ----- Nota metodológica -----
    _rodape_metodologia(doc)


def _rodape_metodologia(doc) -> None:
    add_banner(
        doc,
        "Limites oficiais por sensor (⚠ derate / ⛔ downtime): Gerador enrol. "
        "130/135 · Gerador rolam. 90/100 · Gerador slipring 75/80 · "
        "Gearbox rolam. 85/90 · Gearbox óleo 75/80 · Trafo enrol. 140/145 · "
        "Nacele —/45 °C. Pico via agregação MAX; média via AVG. "
        "Fonte: cockpit de anomalia térmica (SG).",
        cor=HEX_AMARELO,
    )