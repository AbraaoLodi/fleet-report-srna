"""Gráficos matplotlib estilo Serena — PNGs inseridos no docx.

Paleta CORRETA (#FF5246 coral, #32CAA0 verde, #4B4B4B grafite,
#F59E0B amarelo, #9CA3AF cinza).

Funções:
    grafico_disponibilidade  — linha coral + meta 98% tracejada (eixo 94-100)
    grafico_taxas_3series    — falha grafite + corr-NP coral + corr-P verde
    grafico_barras_temp      — barras horizontais por WTG, cor por subsistema
"""
from __future__ import annotations

import os
import tempfile
from pathlib import Path
from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # backend sem display
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator


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


def _setup_axes(ax, *, eixo_y_titulo: str = ""):
    """Aplica estilo Serena nos eixos."""
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(CINZA_BORDA)
    ax.spines["bottom"].set_color(CINZA_BORDA)
    ax.tick_params(colors=GRAFITE, labelsize=9)
    if eixo_y_titulo:
        ax.set_ylabel(eixo_y_titulo, color=GRAFITE, fontsize=9)
    ax.set_facecolor(BRANCO)
    ax.grid(axis="y", color=CINZA_BORDA, linewidth=0.6, alpha=0.7)
    ax.set_axisbelow(True)


# =============================================================================
# Pg 2 — Disponibilidade ao longo das semanas
# =============================================================================

def grafico_disponibilidade(semanas: list[int], valores: list[float],
                              meta: float = 98.0, out_path: str | Path = str(Path(tempfile.gettempdir()) / "g_disp.png"),
                              titulo: str = "EVOLUÇÃO SEMANAL — DISPONIBILIDADE VS. META (98%)") -> str:
    """Linha coral + meta tracejada cinza."""
    fig, ax = plt.subplots(figsize=(8.5, 3.3), dpi=160)
    fig.patch.set_facecolor(BRANCO)

    labels = [f"S{s}" for s in semanas]
    x = list(range(len(semanas)))

    ax.plot(x, valores, color=CORAL, linewidth=2.4, marker="o",
            markersize=7, markerfacecolor=CORAL, markeredgewidth=0,
            label="Disponibilidade (%)")
    ax.axhline(meta, color=CINZA, linewidth=1.4, linestyle="--",
               label=f"Meta {int(meta)}%")

    # Anotações em cima dos pontos (valor real)
    for xi, vi in zip(x, valores):
        if vi is None:
            continue
        ax.annotate(f"{vi:.2f}".replace(".", ",") + "%",
                    xy=(xi, vi), xytext=(0, 9), textcoords="offset points",
                    ha="center", fontsize=9, color=CORAL, weight="bold")

    # Rótulo único da meta (à direita da linha tracejada)
    ax.annotate(f"Meta {int(meta)}%",
                xy=(x[-1], meta), xytext=(8, 0), textcoords="offset points",
                ha="left", va="center", fontsize=8, color=CINZA)

    # Eixo Y dinâmico — engloba valores reais E a meta, com folga
    todos_v = [v for v in valores if v is not None]
    if todos_v:
        v_min = min(min(todos_v), meta)
        v_max = max(max(todos_v), meta)
        margem = max(1.0, (v_max - v_min) * 0.25)
        y_min = max(0, v_min - margem)
        y_max = min(100.5, v_max + margem)
    else:
        y_min, y_max = 80, 100
    ax.set_ylim(y_min, y_max)
    span = y_max - y_min
    passo = 1 if span <= 8 else (2 if span <= 20 else 5)
    ax.yaxis.set_major_locator(MultipleLocator(passo))

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10, color=GRAFITE)
    ax.set_xlim(-0.4, len(x) - 1 + 0.7)  # folga p/ rótulo da meta à direita
    _setup_axes(ax)

    # Título
    fig.suptitle(titulo, fontsize=10, color=GRAFITE_ESC, weight="bold",
                 x=0.07, ha="left", y=0.99)

    plt.tight_layout(rect=(0, 0, 1, 0.92))
    fig.savefig(out_path, bbox_inches="tight", facecolor=BRANCO)
    plt.close(fig)
    return str(out_path)


# =============================================================================
# Pg 3 — Taxa de Falha vs Corretivas (3 séries)
# =============================================================================

def grafico_taxas_3series(semanas: list[int], falha: list[float],
                            corr_np: list[float], corr_p: list[float],
                            grande_corr: list[float] | None = None,
                            out_path: str | Path = str(Path(tempfile.gettempdir()) / "g_taxas.png"),
                            titulo: str = "TAXA DE FALHA X CORRETIVA — S18 A S21") -> str:
    """4 séries: Falha=grafite, CorrNP=coral, CorrP=verde, GC=amarelo.
    Cada valor é anotado numa caixinha branca com borda da cor da série,
    posicionada à esquerda/direita do ponto e empilhada em slots verticais
    fixos, garantindo que nunca se sobreponham."""
    fig, ax = plt.subplots(figsize=(9.2, 4.6), dpi=160)
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

    # Desenha as linhas
    for label, vals, cor in series:
        if vals is None:
            continue
        ax.plot(x, vals, color=cor, linewidth=2.2, marker="o",
                markersize=6, markerfacecolor=cor, markeredgewidth=0,
                label=label)

    todos: list[float] = []
    for _, vals, _ in series:
        if vals:
            todos.extend([v for v in vals if v is not None])
    v_max = max(todos) if todos else 8
    import math as _math
    # teto com folga para os rótulos empilhados (até 4 por coluna)
    y_max = _math.ceil(v_max) + 3
    ax.set_ylim(-1.5, y_max)
    ax.yaxis.set_major_locator(MultipleLocator(1))  # grid de 1 em 1 (zoom)

    # Para cada coluna (semana), ordena os 4 valores e empilha os rótulos
    # em slots verticais fixos, à direita (1ª metade) ou esquerda (2ª metade)
    # do ponto, em caixinhas brancas — assim nunca colidem.
    for xi in x:
        pontos = []
        for _, vals, cor in series:
            if vals is None:
                continue
            vi = vals[xi]
            if vi is not None:
                pontos.append((vi, cor))
        if not pontos:
            continue
        pontos.sort(key=lambda t: -t[0])  # maior em cima

        # lado: colunas da metade esquerda jogam rótulo à direita; da direita, à esquerda
        if xi <= (n - 1) / 2:
            lado, ha = 1, "left"
        else:
            lado, ha = -1, "right"

        # slots verticais fixos centrados na média dos valores da coluna
        m = len(pontos)
        centro = sum(v for v, _ in pontos) / m
        passo_slot = max(0.85, y_max * 0.13)   # distância vertical fixa entre rótulos
        # topo do empilhamento
        y0 = centro + (m - 1) / 2 * passo_slot
        for idx, (vi, cor) in enumerate(pontos):
            y_label = y0 - idx * passo_slot
            # caixinha branca com borda da cor da série, em slot vertical fixo
            ax.text(xi + lado * 0.16, y_label, f"{vi:.2f}".replace(".", ","),
                    ha=ha, va="center", fontsize=8, color=cor, weight="bold",
                    bbox=dict(boxstyle="round,pad=0.18", facecolor="white",
                              edgecolor=cor, linewidth=0.8))
            # linha-guia fina ligando o ponto ao rótulo
            ax.plot([xi, xi + lado * 0.16], [vi, y_label],
                    color=cor, linewidth=0.5, alpha=0.4, zorder=1)

    ax.set_xticks(x)
    ax.set_xticklabels([f"S{s}" for s in semanas], fontsize=10, color=GRAFITE)
    ax.set_xlim(-0.7, n - 1 + 0.7)
    _setup_axes(ax)

    fig.suptitle(titulo, fontsize=10, color=GRAFITE_ESC, weight="bold",
                 x=0.07, ha="left", y=0.99)

    ncol = len(series)
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.22),
              ncol=ncol, frameon=False, fontsize=8.5)

    plt.tight_layout(rect=(0, 0.04, 1, 0.92))
    fig.savefig(out_path, bbox_inches="tight", facecolor=BRANCO)
    plt.close(fig)
    return str(out_path)


# =============================================================================
# Pg 5 — Barras horizontais (downtime térmico por WTG)
# =============================================================================

def grafico_barras_temp(rows: list[dict],
                          out_path: str | Path = str(Path(tempfile.gettempdir()) / "g_temp.png"),
                          titulo: str = "ANÁLISE DE TEMPERATURA — DOWNTIME POR EVENTO TÉRMICO",
                          subtitulo: str = "") -> str:
    """Barras horizontais por WTG, cor por subsistema.

    rows: [{"wtg": "A-II-15", "subsistema": "HYDRAULICS",
            "horas": 10.69, "label": "Hidráulica"}]
    Já ordenado descendente pelo orquestrador.
    """
    if not rows:
        # gráfico vazio
        fig, ax = plt.subplots(figsize=(8.5, 2.5), dpi=160)
        ax.text(0.5, 0.5, "Sem eventos térmicos na janela",
                ha="center", va="center", color=CINZA, fontsize=11)
        ax.axis("off")
        fig.savefig(out_path, bbox_inches="tight", facecolor=BRANCO)
        plt.close(fig)
        return str(out_path)

    n = len(rows)
    altura = max(2.2, 0.5 * n + 1.4)
    fig, ax = plt.subplots(figsize=(8.5, altura), dpi=160)
    fig.patch.set_facecolor(BRANCO)

    y_pos = list(range(n))
    # Eixo Y mostra só a WTG (o subsistema vai na cor + legenda)
    labels_y = [str(r["wtg"]) for r in rows]
    valores = [r["horas"] for r in rows]
    cores = [COR_SUBSISTEMA.get(r["subsistema"].upper(), CORAL) for r in rows]

    for yi in y_pos:
        if yi % 2 == 1:
            ax.axhspan(yi - 0.45, yi + 0.45, color="#FAF9F8", zorder=0)

    ax.barh(y_pos, valores, color=cores, height=0.6, zorder=2)

    # Anotação ao lado da barra: só as horas (subsistema já está na cor)
    for yi, vi in zip(y_pos, valores):
        ax.text(vi + max(valores) * 0.015, yi,
                f"{vi:.2f}".replace(".", ",") + "h",
                va="center", fontsize=8.5, color=GRAFITE, weight="bold")

    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels_y, fontsize=9, color=GRAFITE_ESC)
    ax.invert_yaxis()

    ax.set_xlim(0, max(valores) * 1.30 if valores else 1)
    passo_x = max(1, round(max(valores) / 6)) if valores else 1
    ax.xaxis.set_major_locator(MultipleLocator(passo_x))
    ax.set_xlabel("Horas de parada por temperatura (últimos 7 dias)",
                  fontsize=8.5, color=CINZA)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(CINZA_BORDA)
    ax.spines["bottom"].set_color(CINZA_BORDA)
    ax.tick_params(colors=GRAFITE, labelsize=9)
    ax.grid(axis="x", color=CINZA_BORDA, linewidth=0.5, alpha=0.6)
    ax.set_axisbelow(True)

    fig.suptitle(titulo, fontsize=10, color=GRAFITE_ESC, weight="bold",
                 x=0.02, ha="left", y=0.99)

    # Legenda dos subsistemas presentes — acima do gráfico, fora da área de dados
    subs_unicos = []
    cores_unicas = []
    for r in rows:
        chave = r["subsistema"].upper()
        if chave not in subs_unicos:
            subs_unicos.append(chave)
            cores_unicas.append(COR_SUBSISTEMA.get(chave, CORAL))

    handles = [plt.Rectangle((0, 0), 1, 1, color=c) for c in cores_unicas]
    nomes = [_titulo_subsistema(s) for s in subs_unicos]
    ax.legend(handles, nomes,
              loc="lower right", bbox_to_anchor=(1.0, 1.0),
              ncol=min(len(subs_unicos), 5),
              frameon=False, fontsize=8)

    plt.tight_layout(rect=(0, 0, 1, 0.90))
    fig.savefig(out_path, bbox_inches="tight", facecolor=BRANCO)
    plt.close(fig)
    return str(out_path)


def _titulo_subsistema(subsistema: str) -> str:
    """HYDRAULICS → Hidráulica; GENERATOR → Gerador; etc."""
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