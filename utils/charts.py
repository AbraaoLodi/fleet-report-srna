"""Gráficos matplotlib estilo Serena — PNGs inseridos no docx.

Paleta CORRETA (#FF5246 coral, #32CAA0 verde, #4B4B4B grafite,
#F59E0B amarelo, #9CA3AF cinza).

Tipografia: títulos em fonte **Serena**; rótulos/eixos em **Poppins**
(com fallback silencioso para DejaVu Sans se as fontes não estiverem
instaladas no sistema — em produção, rodar utils/fonts.py uma vez).

Funções:
    grafico_disponibilidade  — linha coral + meta 98% tracejada (eixo dinâmico)
    grafico_taxas_3series    — falha grafite + corr-NP coral + corr-P verde + GC amarelo
    grafico_barras_temp      — barras horizontais por WTG, cor por subsistema
"""
from __future__ import annotations

import tempfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # backend sem display
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator

from utils.fonts import FONTE_TITULO, FONTE_CORPO, registrar_fontes

registrar_fontes()


# Paleta como hex
CORAL          = "#FF5246"
CORAL_ESCURO   = "#D23532"
GRAFITE        = "#4B4B4B"
GRAFITE_ESC    = "#181818"
VERDE          = "#32CAA0"
AMARELO        = "#F59E0B"
CINZA          = "#9CA3AF"
CINZA_BORDA    = "#E5E3E0"
AREIA          = "#EDE9E5"
BRANCO         = "#FFFFFF"

# Mapa subsistema → cor para o gráfico de temperatura
COR_SUBSISTEMA = {
    "GENERATOR":          CORAL,
    "HYDRAULICS":         GRAFITE,
    "CONVERTER":          AMARELO,
    "FREQUENCY CONVERTER": AMARELO,
    "BRAKE":              CINZA,
    "GEARBOX":            CORAL_ESCURO,
    "PITCH SYSTEM":       VERDE,
    "POWER ELECTRICAL":   GRAFITE_ESC,
    "POWER ELECTRICAL SYSTEM": GRAFITE_ESC,
    "CONTROL & COMMUNICATION SYSTEM": "#6B7280",
    "YAW SYSTEM":         "#6B7280",
}


def _titulo(fig, texto: str, *, x: float = 0.04, y: float = 0.99) -> None:
    """Título padronizado em fonte Serena, alinhado à esquerda."""
    fig.suptitle(texto, fontsize=12, color=GRAFITE_ESC, fontweight="bold",
                 fontfamily=FONTE_TITULO, x=x, ha="left", y=y)


def _setup_axes(ax, *, eixo_y_titulo: str = ""):
    """Aplica estilo Serena nos eixos (grid discreto, sem molduras)."""
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color(CINZA_BORDA)
    ax.tick_params(colors=GRAFITE, labelsize=9, length=0)
    for lbl in ax.get_xticklabels() + ax.get_yticklabels():
        lbl.set_fontfamily(FONTE_CORPO)
    if eixo_y_titulo:
        ax.set_ylabel(eixo_y_titulo, color=GRAFITE, fontsize=9,
                      fontfamily=FONTE_CORPO)
    ax.set_facecolor(BRANCO)
    ax.grid(axis="y", color=CINZA_BORDA, linewidth=0.8, alpha=0.6)
    ax.set_axisbelow(True)


# =============================================================================
# Pg 2 — Disponibilidade ao longo das semanas
# =============================================================================

def grafico_disponibilidade(semanas: list[int], valores: list[float],
                              meta: float = 98.0,
                              out_path: str | Path = str(Path(tempfile.gettempdir()) / "g_disp.png"),
                              titulo: str = "Evolução semanal — disponibilidade vs. meta") -> str:
    """Linha coral + meta tracejada cinza (rótulo da meta fora da série)."""
    fig, ax = plt.subplots(figsize=(8.5, 3.4), dpi=170)
    fig.patch.set_facecolor(BRANCO)

    labels = [f"S{s:02d}" for s in semanas]
    x = list(range(len(semanas)))

    # Linha da meta primeiro (fica no fundo)
    ax.axhline(meta, color=CINZA, linewidth=1.3, linestyle=(0, (6, 4)),
               zorder=1)

    ax.plot(x, valores, color=CORAL, linewidth=2.6, marker="o",
            markersize=7.5, markerfacecolor=CORAL, markeredgecolor=BRANCO,
            markeredgewidth=1.4, zorder=3)

    # Área sutil sob a linha (limitada ao piso do eixo, sem vazar)
    y_piso = max(0, min([v for v in valores if v is not None]) - 5)
    ax.fill_between(x, valores, y_piso, color=CORAL, alpha=0.06, zorder=0)

    # Anotações dos valores — acima do ponto, em Poppins
    for xi, vi in zip(x, valores):
        if vi is None:
            continue
        ax.annotate(f"{vi:.2f}".replace(".", ",") + "%",
                    xy=(xi, vi), xytext=(0, 11), textcoords="offset points",
                    ha="center", fontsize=9.5, color=CORAL_ESCURO,
                    fontweight="bold", fontfamily=FONTE_CORPO, zorder=4)

    # Eixo Y dinâmico — engloba valores E meta, com folga p/ respiro
    todos_v = [v for v in valores if v is not None]
    v_min = min(min(todos_v), meta)
    v_max = max(max(todos_v), meta)
    span = max(v_max - v_min, 1.0)
    y_min = v_min - span * 0.28          # mais respiro abaixo do menor ponto
    y_max = v_max + span * 0.22          # folga no topo p/ a linha da meta
    y_min = max(0, y_min)
    y_max = min(100.6, y_max)
    ax.set_ylim(y_min, y_max)
    passo = 1 if (y_max - y_min) <= 8 else (2 if (y_max - y_min) <= 20 else 5)
    ax.yaxis.set_major_locator(MultipleLocator(passo))

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10.5, color=GRAFITE)
    # folga lateral mínima (só o necessário para os marcadores não cortarem)
    ax.set_xlim(-0.12, len(x) - 1 + 0.12)

    # Rótulo da meta no canto superior direito, sobre a linha tracejada
    ax.annotate(f"Meta {meta:.0f}%".replace(".0", ""),
                xy=(x[-1], meta), xytext=(0, 6), textcoords="offset points",
                ha="right", va="bottom", fontsize=9, color=CINZA,
                fontweight="bold", fontfamily=FONTE_CORPO,
                annotation_clip=False)

    _setup_axes(ax)
    _titulo(fig, titulo.upper())

    plt.tight_layout(rect=(0, 0, 1, 0.90))
    fig.savefig(out_path, bbox_inches="tight", pad_inches=0.05, facecolor=BRANCO)
    plt.close(fig)
    return str(out_path)


# =============================================================================
# Pg 3 — Taxa de Falha vs Corretivas (4 séries)
# =============================================================================

def grafico_taxas_3series(semanas: list[int], falha: list[float],
                            corr_np: list[float], corr_p: list[float],
                            grande_corr: list[float] | None = None,
                            out_path: str | Path = str(Path(tempfile.gettempdir()) / "g_taxas.png"),
                            titulo: str = "Taxa de falha x corretiva") -> str:
    """4 séries: Falha=grafite, CorrNP=coral, CorrP=verde, GC=amarelo.

    Layout limpo: eixo Y começa em 0 com teto justo; rótulos apenas no
    ÚLTIMO ponto de cada série (a legenda cuida da identificação). Sem
    caixinhas nem linhas-guia — muito mais legível."""
    fig, ax = plt.subplots(figsize=(9.0, 3.9), dpi=170)
    fig.patch.set_facecolor(BRANCO)

    x = list(range(len(semanas)))
    n = len(x)

    series = [
        ("Falha", falha, GRAFITE_ESC),
        ("Corretiva Não Planejada", corr_np, CORAL),
        ("Corretiva Planejada", corr_p, VERDE),
    ]
    if grande_corr is not None:
        series.append(("Grande Corretiva", grande_corr, AMARELO))

    # Linhas
    for label, vals, cor in series:
        if vals is None:
            continue
        ax.plot(x, vals, color=cor, linewidth=2.4, marker="o",
                markersize=6.5, markerfacecolor=cor, markeredgecolor=BRANCO,
                markeredgewidth=1.2, label=label, zorder=3)

    # Teto justo do eixo Y
    todos: list[float] = []
    for _, vals, _ in series:
        if vals:
            todos.extend([v for v in vals if v is not None])
    v_max = max(todos) if todos else 8
    y_max = v_max * 1.15 + 0.5
    ax.set_ylim(-y_max * 0.14, y_max)   # folga inferior maior p/ rótulos verdes
    passo = 1 if y_max <= 9 else 2
    ax.yaxis.set_major_locator(MultipleLocator(passo))

    # Rótulos em TODAS as semanas. Direção fixa por série para evitar colisão
    # entre as que ficam próximas do zero (Falha, CorrP) e com o eixo X.
    #   CorrNP (coral)   → acima
    #   GC (amarelo)     → acima
    #   Falha (grafite)  → acima
    #   CorrP (verde)    → abaixo (única embaixo, longe das outras)
    offsets = {
        CORAL:       (0,  10, "bottom"),
        AMARELO:     (0,  10, "bottom"),
        GRAFITE_ESC: (0,  10, "bottom"),
        VERDE:       (0, -12, "top"),
    }
    for label, vals, cor in series:
        if not vals:
            continue
        dx, dy, va = offsets.get(cor, (0, 10, "bottom"))
        for xi, vi in zip(x, vals):
            if vi is None:
                continue
            ddy, dva = dy, va
            # Falha: se estiver alta (perto da linha GC ~2%), rotula ABAIXO
            # do ponto para não colar na série amarela.
            if cor == GRAFITE_ESC and vi >= 1.3:
                ddy, dva = -12, "top"
            ax.annotate(f"{vi:.2f}".replace(".", ","),
                        xy=(xi, vi), xytext=(dx, ddy), textcoords="offset points",
                        ha="center", va=dva, fontsize=7.5, color=cor,
                        fontweight="bold", fontfamily=FONTE_CORPO, zorder=4)

    ax.set_xticks(x)
    ax.set_xticklabels([f"S{s:02d}" for s in semanas], fontsize=10.5, color=GRAFITE)
    ax.tick_params(axis="x", pad=8)
    ax.set_xlim(-0.3, n - 1 + 0.6)
    _setup_axes(ax, eixo_y_titulo="% de horas-frota")
    _titulo(fig, titulo.upper())

    leg = ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.26),
                    ncol=len(series), frameon=False, fontsize=9,
                    handlelength=1.4, columnspacing=1.6)
    for t in leg.get_texts():
        t.set_fontfamily(FONTE_CORPO)
        t.set_color(GRAFITE)

    plt.tight_layout(rect=(0, 0.06, 0.98, 0.90))
    fig.savefig(out_path, bbox_inches="tight", facecolor=BRANCO)
    plt.close(fig)
    return str(out_path)


# =============================================================================
# Pg 5 — Barras horizontais (downtime térmico por WTG)
# =============================================================================

def grafico_barras_temp(rows: list[dict],
                          out_path: str | Path = str(Path(tempfile.gettempdir()) / "g_temp.png"),
                          titulo: str = "Análise de temperatura — downtime por evento térmico",
                          subtitulo: str = "") -> str:
    """Barras horizontais por WTG, cor por subsistema."""
    if not rows:
        fig, ax = plt.subplots(figsize=(8.5, 2.5), dpi=170)
        ax.text(0.5, 0.5, "Sem eventos térmicos na janela",
                ha="center", va="center", color=CINZA, fontsize=11,
                fontfamily=FONTE_CORPO)
        ax.axis("off")
        fig.savefig(out_path, bbox_inches="tight", facecolor=BRANCO)
        plt.close(fig)
        return str(out_path)

    n = len(rows)
    altura = max(2.2, 0.5 * n + 1.4)
    fig, ax = plt.subplots(figsize=(8.5, altura), dpi=170)
    fig.patch.set_facecolor(BRANCO)

    y_pos = list(range(n))
    labels_y = [str(r["wtg"]) for r in rows]
    valores = [r["horas"] for r in rows]
    cores = [COR_SUBSISTEMA.get(r["subsistema"].upper(), CORAL) for r in rows]

    for yi in y_pos:
        if yi % 2 == 1:
            ax.axhspan(yi - 0.45, yi + 0.45, color="#FAF9F8", zorder=0)

    ax.barh(y_pos, valores, color=cores, height=0.6, zorder=2)

    for yi, vi in zip(y_pos, valores):
        ax.text(vi + max(valores) * 0.015, yi,
                f"{vi:.2f}".replace(".", ",") + "h",
                va="center", fontsize=8.5, color=GRAFITE, fontweight="bold",
                fontfamily=FONTE_CORPO)

    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels_y, fontsize=9, color=GRAFITE_ESC)
    ax.invert_yaxis()

    ax.set_xlim(0, max(valores) * 1.30 if valores else 1)
    passo_x = max(1, round(max(valores) / 6)) if valores else 1
    ax.xaxis.set_major_locator(MultipleLocator(passo_x))
    ax.set_xlabel("Horas de parada por temperatura (últimos 7 dias)",
                  fontsize=8.5, color=CINZA, fontfamily=FONTE_CORPO)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color(CINZA_BORDA)
    ax.tick_params(colors=GRAFITE, labelsize=9, length=0)
    ax.grid(axis="x", color=CINZA_BORDA, linewidth=0.5, alpha=0.6)
    ax.set_axisbelow(True)

    _titulo(fig, titulo.upper(), x=0.02)

    subs_unicos, cores_unicas = [], []
    for r in rows:
        chave = r["subsistema"].upper()
        if chave not in subs_unicos:
            subs_unicos.append(chave)
            cores_unicas.append(COR_SUBSISTEMA.get(chave, CORAL))
    handles = [plt.Rectangle((0, 0), 1, 1, color=c) for c in cores_unicas]
    nomes = [_titulo_subsistema(s) for s in subs_unicos]
    leg = ax.legend(handles, nomes, loc="lower right", bbox_to_anchor=(1.0, 1.0),
                    ncol=min(len(subs_unicos), 5), frameon=False, fontsize=8)
    for t in leg.get_texts():
        t.set_fontfamily(FONTE_CORPO)

    plt.tight_layout(rect=(0, 0, 1, 0.90))
    fig.savefig(out_path, bbox_inches="tight", facecolor=BRANCO)
    plt.close(fig)
    return str(out_path)


def _titulo_subsistema(subsistema: str) -> str:
    mapa = {
        "GENERATOR":          "Gerador",
        "HYDRAULICS":         "Hidráulica",
        "CONVERTER":          "Conversor",
        "FREQUENCY CONVERTER": "Conversor",
        "BRAKE":              "Freio",
        "GEARBOX":            "Gearbox",
        "PITCH SYSTEM":       "Pitch",
        "POWER ELECTRICAL":   "Elétrico",
        "POWER ELECTRICAL SYSTEM": "Elétrico",
        "YAW SYSTEM":         "Yaw",
    }
    return mapa.get(subsistema.upper(), subsistema.title())