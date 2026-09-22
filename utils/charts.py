"""Gráficos matplotlib estilo Serena — PNGs inseridos no docx."""
from __future__ import annotations
import tempfile
from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
from utils.fonts import FONTE_TITULO, FONTE_CORPO, registrar_fontes
registrar_fontes()

CORAL="#FF5246"; CORAL_ESCURO="#D23532"; GRAFITE="#4B4B4B"; GRAFITE_ESC="#181818"
VERDE="#32CAA0"; AMARELO="#F59E0B"; CINZA="#9CA3AF"; CINZA_BORDA="#E5E3E0"
AREIA="#EDE9E5"; BRANCO="#FFFFFF"

COR_SUBSISTEMA = {
    "GENERATOR":CORAL,"HYDRAULICS":GRAFITE,"CONVERTER":AMARELO,"FREQUENCY CONVERTER":AMARELO,
    "BRAKE":CINZA,"GEARBOX":CORAL_ESCURO,"PITCH SYSTEM":VERDE,"POWER ELECTRICAL":GRAFITE_ESC,
    "POWER ELECTRICAL SYSTEM":GRAFITE_ESC,"CONTROL & COMMUNICATION SYSTEM":"#6B7280","YAW SYSTEM":"#6B7280",
}

def _titulo(fig, texto, x=0.04, y=0.99):
    fig.suptitle(texto, fontsize=12, color=GRAFITE_ESC, fontweight="bold",
                 fontfamily=FONTE_TITULO, x=x, ha="left", y=y)

def _setup_axes(ax, eixo_y_titulo=""):
    for s in ("top","right","left"): ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(CINZA_BORDA)
    ax.tick_params(colors=GRAFITE, labelsize=9, length=0)
    for lbl in ax.get_xticklabels()+ax.get_yticklabels(): lbl.set_fontfamily(FONTE_CORPO)
    if eixo_y_titulo: ax.set_ylabel(eixo_y_titulo, color=GRAFITE, fontsize=9, fontfamily=FONTE_CORPO)
    ax.set_facecolor(BRANCO)
    ax.grid(axis="y", color=CINZA_BORDA, linewidth=0.8, alpha=0.6)
    ax.set_axisbelow(True)

def grafico_disponibilidade(semanas, valores, meta=98.0,
        out_path=str(Path(tempfile.gettempdir())/"g_disp.png"),
        titulo="Evolução semanal — disponibilidade vs. meta"):
    fig, ax = plt.subplots(figsize=(8.5,3.4), dpi=170); fig.patch.set_facecolor(BRANCO)
    labels=[f"S{s:02d}" for s in semanas]; x=list(range(len(semanas)))
    ax.axhline(meta, color=CINZA, linewidth=1.3, linestyle=(0,(6,4)), zorder=1)
    ax.plot(x, valores, color=CORAL, linewidth=2.6, marker="o", markersize=7.5,
            markerfacecolor=CORAL, markeredgecolor=BRANCO, markeredgewidth=1.4, zorder=3)
    y_piso = max(0, min([v for v in valores if v is not None])-5)
    ax.fill_between(x, valores, y_piso, color=CORAL, alpha=0.06, zorder=0)
    for xi,vi in zip(x,valores):
        if vi is None: continue
        ax.annotate(f"{vi:.2f}".replace(".",",")+"%", xy=(xi,vi), xytext=(0,11),
                    textcoords="offset points", ha="center", fontsize=9.5, color=CORAL_ESCURO,
                    fontweight="bold", fontfamily=FONTE_CORPO, zorder=4)
    todos=[v for v in valores if v is not None]
    v_min=min(min(todos),meta); v_max=max(max(todos),meta); span=max(v_max-v_min,1.0)
    y_min=max(0, v_min-span*0.28); y_max=min(100.6, v_max+span*0.22)
    ax.set_ylim(y_min,y_max)
    passo = 1 if (y_max-y_min)<=8 else (2 if (y_max-y_min)<=20 else 5)
    ax.yaxis.set_major_locator(MultipleLocator(passo))
    ax.set_xticks(x); ax.set_xticklabels(labels, fontsize=10.5, color=GRAFITE)
    ax.set_xlim(-0.12, len(x)-1+0.12)
    ax.annotate(f"Meta {meta:.0f}%".replace(".0",""), xy=(x[-1],meta), xytext=(0,6),
                textcoords="offset points", ha="right", va="bottom", fontsize=9, color=CINZA,
                fontweight="bold", fontfamily=FONTE_CORPO, annotation_clip=False)
    _setup_axes(ax); _titulo(fig, titulo.upper())
    plt.tight_layout(rect=(0,0,1,0.90))
    fig.savefig(out_path, bbox_inches="tight", pad_inches=0.05, facecolor=BRANCO); plt.close(fig)
    return str(out_path)

def grafico_taxas_3series(semanas, falha, corr_np, corr_p, grande_corr=None,
        out_path=str(Path(tempfile.gettempdir())/"g_taxas.png"), titulo="Taxa de falha x corretiva"):
    """Barras agrupadas: 4 motivos lado a lado por semana. Escala e rótulos
    dimensionam-se automaticamente ao valor máximo (funciona tanto em frotas de
    taxa baixa, ex. Vestas, quanto alta, ex. SG)."""
    fig, ax = plt.subplots(figsize=(9.0,3.9), dpi=170); fig.patch.set_facecolor(BRANCO)
    x=list(range(len(semanas)))
    series=[("Falha",falha,GRAFITE_ESC),("Corretiva Não Planejada",corr_np,CORAL),
            ("Corretiva Planejada",corr_p,VERDE)]
    if grande_corr is not None: series.append(("Grande Corretiva",grande_corr,AMARELO))
    w=0.20
    todos=[v for _,vals,_ in series if vals for v in vals if v is not None]
    vmax=max(todos) if todos else 1
    n=len(series)
    base=-(n-1)/2.0
    for i,(label,vals,cor) in enumerate(series):
        if vals is None: continue
        offs=[xi+(base+i)*w for xi in x]
        ax.bar(offs, vals, color=cor, label=label, width=w, zorder=3)
        for xi,vi in zip(offs,vals):
            if vi is not None and vi>=vmax*0.04:
                ax.text(xi, vi+vmax*0.02, f"{vi:.2f}".replace(".",","), ha="center", va="bottom",
                        fontsize=6.5, color=cor, fontweight="bold", fontfamily=FONTE_CORPO, zorder=4)
    for s in ("top","right","left"): ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(CINZA_BORDA)
    ax.tick_params(colors=GRAFITE, labelsize=9, length=0)
    ax.set_xticks(x); ax.set_xticklabels([f"S{s:02d}" for s in semanas], fontsize=10.5, color=GRAFITE)
    ax.grid(axis="y", color=CINZA_BORDA, linewidth=0.8, alpha=0.6); ax.set_axisbelow(True)
    ax.set_ylabel("% de horas-frota", color=GRAFITE, fontsize=9, fontfamily=FONTE_CORPO)
    ax.set_ylim(0, vmax*1.20)
    _titulo(fig, titulo.upper())
    leg = ax.legend(loc="lower center", bbox_to_anchor=(0.5,-0.28), ncol=len(series),
                    frameon=False, fontsize=8.5, handlelength=1.4, columnspacing=1.6)
    for t in leg.get_texts(): t.set_fontfamily(FONTE_CORPO); t.set_color(GRAFITE)
    plt.tight_layout(rect=(0,0.05,1,0.90))
    fig.savefig(out_path, bbox_inches="tight", facecolor=BRANCO); plt.close(fig)
    return str(out_path)

def grafico_barras_warnings(labels, valores, cores=None,
        out_path=str(Path(tempfile.gettempdir())/"g_warn.png"),
        titulo="Warnings mais recorrentes", eixo_x="Ocorrências"):
    """Barras horizontais para os warnings/alarmes mais recorrentes."""
    if not labels:
        fig, ax = plt.subplots(figsize=(8.5,2.2), dpi=170)
        ax.text(0.5,0.5,"Sem warnings no período", ha="center", va="center",
                color=CINZA, fontsize=11, fontfamily=FONTE_CORPO); ax.axis("off")
        fig.savefig(out_path, bbox_inches="tight", facecolor=BRANCO); plt.close(fig)
        return str(out_path)
    n=len(labels); altura=max(2.2, 0.46*n+1.3)
    fig, ax = plt.subplots(figsize=(9.6,altura), dpi=170); fig.patch.set_facecolor(BRANCO)
    y=list(range(n))
    if cores is None: cores=[CORAL]*n
    for yi in y:
        if yi%2==1: ax.axhspan(yi-0.45,yi+0.45,color="#FAF9F8",zorder=0)
    ax.barh(y, valores, color=cores, height=0.62, zorder=2)
    vmax=max(valores)
    for yi,vi in zip(y,valores):
        ax.text(vi+vmax*0.015, yi, f"{int(vi):,}".replace(",","."), va="center",
                fontsize=8.5, color=GRAFITE, fontweight="bold", fontfamily=FONTE_CORPO)
    ax.set_yticks(y); ax.set_yticklabels(labels, fontsize=8, color=GRAFITE_ESC)
    ax.invert_yaxis(); ax.set_xlim(0, vmax*1.16)
    ax.set_xlabel(eixo_x, fontsize=8.5, color=CINZA, fontfamily=FONTE_CORPO)
    for s in ("top","right","left"): ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(CINZA_BORDA)
    ax.tick_params(colors=GRAFITE, labelsize=9, length=0)
    ax.grid(axis="x", color=CINZA_BORDA, linewidth=0.5, alpha=0.6); ax.set_axisbelow(True)
    _titulo(fig, titulo.upper(), x=0.02)
    plt.tight_layout(rect=(0,0,1,0.91))
    fig.savefig(out_path, bbox_inches="tight", facecolor=BRANCO); plt.close(fig)
    return str(out_path)
