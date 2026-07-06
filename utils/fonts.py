"""Registro das fontes da marca Serena no matplotlib.

A marca usa duas famílias:
  - **Serena** (proprietária) → títulos e destaques
  - **Poppins** (Google Fonts) → textos longos e funcionais

Este módulo procura os arquivos de fonte em locais conhecidos e os
registra no matplotlib. Se não encontrar, faz fallback silencioso para
"DejaVu Sans" (a fonte padrão do matplotlib), garantindo que a geração
do relatório nunca quebre por causa de fonte ausente.

Coloque os .otf/.ttf em UMA destas pastas (a primeira que existir vence):
  - $SERENA_FONTS_DIR (variável de ambiente, se definida)
  - <raiz_do_projeto>/assets/fonts/
  - ~/.fonts/
"""
from __future__ import annotations

import os
import glob
from pathlib import Path

import matplotlib.font_manager as fm

# Nomes das famílias como o matplotlib as reconhece após o registro.
FONTE_TITULO = "Serena"
FONTE_CORPO = "Poppins"

_FALLBACK = "DejaVu Sans"
_ja_registrou = False


def _pastas_candidatas() -> list[Path]:
    pastas: list[Path] = []
    env = os.environ.get("SERENA_FONTS_DIR")
    if env:
        pastas.append(Path(env))
    raiz = Path(__file__).resolve().parent.parent
    pastas.append(raiz / "assets" / "fonts")
    pastas.append(Path.home() / ".fonts")
    return pastas


def registrar_fontes() -> None:
    """Registra Serena + Poppins no matplotlib (idempotente)."""
    global FONTE_TITULO, FONTE_CORPO, _ja_registrou
    if _ja_registrou:
        return
    _ja_registrou = True

    familias = {f.name for f in fm.fontManager.ttflist}

    for pasta in _pastas_candidatas():
        if not pasta.is_dir():
            continue
        for padrao in ("*.otf", "*.ttf"):
            for arq in glob.glob(str(pasta / padrao)):
                try:
                    fm.fontManager.addfont(arq)
                except Exception:
                    pass
        familias = {f.name for f in fm.fontManager.ttflist}

    # Fallback silencioso se alguma família não estiver disponível.
    if FONTE_TITULO not in familias:
        FONTE_TITULO = _FALLBACK
    if FONTE_CORPO not in familias:
        FONTE_CORPO = _FALLBACK