"""Gera o símbolo Serena em branco/transparente para a capa.

Uso (rodar UMA vez, na máquina onde a skill serena-brand está disponível):

    python utils/assets_gen.py

Lê o símbolo primário oficial (coral sobre branco) e produz uma versão
branca com fundo transparente, recortada no bounding box do conteúdo,
salva em assets/serena_symbol_white_trim.png — que é o arquivo consumido
por docx_helpers.add_capa().

Requer Pillow e numpy (pip install pillow numpy).
"""
from __future__ import annotations

from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
ASSETS = RAIZ / "assets"

# Caminhos candidatos do símbolo primário na skill de marca.
CANDIDATOS = [
    Path("/mnt/skills/organization/serena-brand/assets/simbolo/01_SERENA_SYMBOL_PRIMARY_RGB.png"),
    RAIZ / "assets" / "01_SERENA_SYMBOL_PRIMARY_RGB.png",
]


def gerar() -> Path:
    from PIL import Image
    import numpy as np

    origem = next((c for c in CANDIDATOS if c.exists()), None)
    if origem is None:
        raise FileNotFoundError(
            "Símbolo primário não encontrado. Copie "
            "01_SERENA_SYMBOL_PRIMARY_RGB.png para a pasta assets/ do projeto."
        )

    img = Image.open(origem).convert("RGBA")
    arr = np.array(img)
    r, g, b = arr[:, :, 0].astype(int), arr[:, :, 1].astype(int), arr[:, :, 2].astype(int)
    # pixels coral (não brancos) → branco opaco; resto → transparente
    is_coral = (r > 200) & (g < 160) & (b < 160)
    out = np.zeros_like(arr)
    out[:, :, 0:3] = 255
    out[:, :, 3] = np.where(is_coral, 255, 0).astype("uint8")

    branco = Image.fromarray(out, "RGBA")
    bbox = branco.getbbox()
    if bbox:
        branco = branco.crop(bbox)

    ASSETS.mkdir(parents=True, exist_ok=True)
    destino = ASSETS / "serena_symbol_white_trim.png"
    branco.save(destino)
    return destino


if __name__ == "__main__":
    caminho = gerar()
    print(f"✓ Símbolo branco gerado em: {caminho}")