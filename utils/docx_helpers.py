"""Helpers docx — identidade Serena para o Fleet Reliability Report SRNA.

Paleta (validada contra serena-brand SKILL):
  Coral        #FF5246  (destaque, crítico, alarme)
  Coral escuro #D23532  (negativo)
  Grafite      #4B4B4B  (textos principais, header)
  Grafite esc  #181818  (títulos)
  Verde        #32CAA0  (concluído, bom — uso funcional/UI)
  Amarelo      #F59E0B  (em andamento, atenção)
  Cinza claro  #9CA3AF  (secundário, agendado)
  Areia        #EDE9E5  (fundo alternado tabela)
  Branco       #FFFFFF

Tipografia (marca):
  - FONT_TITULO = "Serena"  → títulos, capa, badge, cabeçalhos de seção
  - FONT_BODY   = "Poppins" → corpo, tabelas, rótulos
  (a fonte precisa estar instalada no Word do usuário; caso contrário o
   Word faz substituição — o docx referencia a família pelo nome.)

Layout:
  - Capa fullpage com SÍMBOLO OFICIAL Serena (PNG) + título grande + linha coral
  - Header grafite fixo + badge coral com a semana
  - Rodapé institucional
  - Cabeçalho de seção: número coral + título Serena + subtítulo cinza
  - KPI cards com altura uniforme
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.document import Document as DocumentT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from docx.shared import Cm, Mm, Pt, RGBColor, Emu, Inches


# =============================================================================
# ASSETS
# =============================================================================
# Símbolo oficial (três barras) em branco com fundo transparente, para uso
# sobre a faixa grafite da capa. Gere-o uma vez a partir do símbolo primário
# da skill serena-brand (ver utils/assets_gen.py) e salve em assets/.
_ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
SIMBOLO_BRANCO = _ASSETS_DIR / "serena_symbol_white_trim.png"


# =============================================================================
# PALETA SERENA — fonte única de verdade
# =============================================================================
HEX_CORAL          = "FF5246"
HEX_CORAL_ESCURO   = "D23532"
HEX_GRAFITE        = "4B4B4B"
HEX_GRAFITE_ESC    = "181818"
HEX_VERDE          = "32CAA0"
HEX_AMARELO        = "F59E0B"
HEX_CINZA_CLARO    = "9CA3AF"
HEX_AREIA          = "EDE9E5"
HEX_AREIA_HEAD     = "F5F2EF"   # mais claro p/ header de tabela
HEX_BRANCO         = "FFFFFF"
HEX_BORDA          = "E5E3E0"

RGB_CORAL          = RGBColor(0xFF, 0x52, 0x46)
RGB_CORAL_ESCURO   = RGBColor(0xD2, 0x35, 0x32)
RGB_GRAFITE        = RGBColor(0x4B, 0x4B, 0x4B)
RGB_GRAFITE_ESC    = RGBColor(0x18, 0x18, 0x18)
RGB_VERDE          = RGBColor(0x32, 0xCA, 0xA0)
RGB_AMARELO        = RGBColor(0xF5, 0x9E, 0x0B)
RGB_CINZA_CLARO    = RGBColor(0x9C, 0xA3, 0xAF)
RGB_BRANCO         = RGBColor(0xFF, 0xFF, 0xFF)

# Tipografia da marca
FONT_TITULO = "Serena"
FONT_BODY   = "Poppins"

# Mapeamento de status → cor + bold
STATUS_COLORS: dict[str, tuple[str, bool]] = {
    "em andamento":  (HEX_AMARELO, False),
    "aberto":        (HEX_CORAL, False),
    "concluído":     (HEX_VERDE, False),
    "concluido":     (HEX_VERDE, False),
    "agendado":      (HEX_CINZA_CLARO, False),
    "crítico":       (HEX_CORAL, True),
    "critico":       (HEX_CORAL, True),
    "atenção":       (HEX_AMARELO, True),
    "atencao":       (HEX_AMARELO, True),
    "monitorar":     (HEX_VERDE, True),
    "bom":           (HEX_VERDE, True),
    "melhor tempo":  (HEX_VERDE, True),
    "médio":         (HEX_AMARELO, False),
    "medio":         (HEX_AMARELO, False),
    "normal":        (HEX_CINZA_CLARO, False),
    "alto":          (HEX_CORAL, True),
    "baixo":         (HEX_CINZA_CLARO, False),
    "volume alto":   (HEX_AMARELO, True),
    "pior semana do período": (HEX_CORAL, True),
    "melhor semana do período": (HEX_VERDE, True),
    "acima da meta": (HEX_VERDE, True),
    "próximo da meta": (HEX_AMARELO, False),
    "abaixo da meta": (HEX_CINZA_CLARO, False),
}


# =============================================================================
# XML helpers
# =============================================================================

def _shade(cell, hex_color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def _set_borders(cell, color: str = HEX_BORDA, size: int = 4,
                  top: bool = True, bottom: bool = True,
                  left: bool = False, right: bool = False) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for side, enabled in [("top", top), ("bottom", bottom),
                          ("left", left), ("right", right)]:
        elem = OxmlElement(f"w:{side}")
        if enabled:
            elem.set(qn("w:val"), "single")
            elem.set(qn("w:sz"), str(size))
            elem.set(qn("w:color"), color)
        else:
            elem.set(qn("w:val"), "nil")
        tc_borders.append(elem)
    tc_pr.append(tc_borders)


def _no_borders(cell) -> None:
    _set_borders(cell, top=False, bottom=False, left=False, right=False)


def _row_cant_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    cant.set(qn("w:val"), "true")
    tr_pr.append(cant)


def _row_as_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    tr_pr.append(th)
    _row_cant_split(row)


def _row_min_height(row, twips: int) -> None:
    """Altura mínima da linha (hRule=atLeast). twips = 1/20 pt."""
    tr_pr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(twips))
    h.set(qn("w:hRule"), "atLeast")
    tr_pr.append(h)


def _set_cell_width_pct(cell, pct: int) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("w:tcW"))
    if existing is not None:
        tcPr.remove(existing)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(pct))
    tcW.set(qn("w:type"), "pct")
    tcPr.append(tcW)


def _fill_table_width_pct(table, pct: int = 5000) -> None:
    tblPr = table._tbl.tblPr
    existing = tblPr.find(qn("w:tblW"))
    if existing is not None:
        tblPr.remove(existing)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(pct))
    tblW.set(qn("w:type"), "pct")
    tblPr.append(tblW)
    existing_ind = tblPr.find(qn("w:tblInd"))
    if existing_ind is not None:
        tblPr.remove(existing_ind)
    tblInd = OxmlElement("w:tblInd")
    tblInd.set(qn("w:w"), "0")
    tblInd.set(qn("w:type"), "dxa")
    tblPr.append(tblInd)


def _fix_table_width(table, largura_emu: int) -> None:
    twips = int(largura_emu / 635)
    tblPr = table._tbl.tblPr
    existing = tblPr.find(qn("w:tblW"))
    if existing is not None:
        tblPr.remove(existing)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(twips))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)
    existing_ind = tblPr.find(qn("w:tblInd"))
    if existing_ind is not None:
        tblPr.remove(existing_ind)
    tblInd = OxmlElement("w:tblInd")
    tblInd.set(qn("w:w"), "0")
    tblInd.set(qn("w:type"), "dxa")
    tblPr.append(tblInd)
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")


def _set_cell_margins(cell, top: int = 80, bottom: int = 80,
                       left: int = 120, right: int = 120) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom),
                       ("left", left), ("right", right)]:
        elem = OxmlElement(f"w:{side}")
        elem.set(qn("w:w"), str(val))
        elem.set(qn("w:type"), "dxa")
        mar.append(elem)
    tc_pr.append(mar)


def _page_break(paragraph) -> None:
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)


def _set_paragraph_spacing(p, before: int = 0, after: int = 0,
                            line: float | None = None) -> None:
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line


def _set_run_letter_spacing(run, val: int) -> None:
    r_pr = run._element.get_or_add_rPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:val"), str(val))
    r_pr.append(spacing)


def _add_bottom_border(paragraph, color: str = HEX_CINZA_CLARO,
                        size: int = 4) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _hex_to_rgb(hex_color: str) -> RGBColor:
    return RGBColor(int(hex_color[0:2], 16),
                    int(hex_color[2:4], 16),
                    int(hex_color[4:6], 16))


# =============================================================================
# Configuração de margens e seções
# =============================================================================

def configurar_pagina(doc: DocumentT,
                       top: float = 0.0, bottom: float = 0.7,
                       left: float = 1.5, right: float = 1.5,
                       header: float = 0.0, footer: float = 0.5) -> None:
    section = doc.sections[-1]
    section.top_margin    = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.left_margin   = Cm(left)
    section.right_margin  = Cm(right)
    section.header_distance = Cm(header)
    section.footer_distance = Cm(footer)


# =============================================================================
# HEADER E RODAPÉ FIXOS
# =============================================================================

def configurar_header_rodape(doc: DocumentT, semana: int, ano: int) -> None:
    for section in doc.sections:
        if getattr(section, "_eh_capa", False):
            continue
        _aplicar_header_section(section, semana, ano)
        _aplicar_footer_section(section)


def _aplicar_header_section(section, semana: int, ano: int) -> None:
    section.header_distance = Cm(0.5)
    section.top_margin = Cm(2.0)
    header = section.header
    header.is_linked_to_previous = False
    for p in list(header.paragraphs):
        if p.text:
            p.text = ""

    largura_total = section.page_width - section.left_margin - section.right_margin
    tbl = header.add_table(rows=1, cols=2, width=largura_total)
    tbl.autofit = False
    tbl.allow_autofit = False
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    col_esq_w = int(largura_total * 0.80)
    col_dir_w = largura_total - col_esq_w
    tbl.columns[0].width = col_esq_w
    tbl.columns[1].width = col_dir_w
    _fill_table_width_pct(tbl, 5000)

    # --- ESQUERDA: fundo grafite com título + subtítulo ---
    cel_esq = tbl.rows[0].cells[0]
    cel_esq.width = col_esq_w
    _set_cell_width_pct(cel_esq, 4000)   # 80%
    _shade(cel_esq, HEX_GRAFITE_ESC)
    _no_borders(cel_esq)
    _set_cell_margins(cel_esq, top=120, bottom=120, left=320, right=160)

    p1 = cel_esq.paragraphs[0]
    _set_paragraph_spacing(p1, before=0, after=0)
    # símbolo textual compacto (o "S" da marca) — mantido pequeno no header
    r_s = p1.add_run("≡  ")
    r_s.font.name = FONT_TITULO
    r_s.font.size = Pt(11)
    r_s.font.bold = True
    r_s.font.color.rgb = RGB_CORAL
    r_t = p1.add_run("Fleet Reliability Report · Assuruá SG | G97")
    r_t.font.name = FONT_TITULO
    r_t.font.size = Pt(10)
    r_t.font.bold = True
    r_t.font.color.rgb = RGB_BRANCO

    p2 = cel_esq.add_paragraph()
    _set_paragraph_spacing(p2, before=2, after=0)
    r_sub = p2.add_run(f"Complexo Eólico Assuruá · 49 WTGs · S{semana:02d} / {ano}")
    r_sub.font.name = FONT_BODY
    r_sub.font.size = Pt(8)
    r_sub.font.color.rgb = RGBColor(0x9C, 0xA3, 0xAF)

    # --- DIREITA: badge coral com semana ---
    cel_dir = tbl.rows[0].cells[1]
    cel_dir.width = col_dir_w
    _set_cell_width_pct(cel_dir, 1000)   # 20%
    cel_dir.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _shade(cel_dir, HEX_CORAL)
    _no_borders(cel_dir)
    _set_cell_margins(cel_dir, top=120, bottom=120, left=160, right=300)

    p_badge = cel_dir.paragraphs[0]
    _set_paragraph_spacing(p_badge, before=0, after=0)
    p_badge.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_badge = p_badge.add_run(f"S{semana:02d} / {ano}")
    r_badge.font.name = FONT_TITULO
    r_badge.font.size = Pt(12)
    r_badge.font.bold = True
    r_badge.font.color.rgb = RGB_BRANCO


def _aplicar_footer_section(section) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    for p in list(footer.paragraphs):
        if p.text:
            p.text = ""
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_paragraph_spacing(p, before=0, after=0)
    r = p.add_run("Serena Energia — Gestão de Ativos SRNA · Documento interno")
    r.font.name = FONT_BODY
    r.font.size = Pt(8)
    r.font.color.rgb = RGB_CINZA_CLARO


# =============================================================================
# CAPA FULLPAGE
# =============================================================================

def add_capa(doc: DocumentT, semana: int, ano: int, periodo_str: str) -> None:
    """Capa com identidade Serena:
    - Faixa grafite no topo (altura fixa) com o SÍMBOLO OFICIAL (PNG) + wordmark
    - Corpo branco: linha coral + título grande em fonte Serena + metadados
    - Espaçamentos generosos (nada colado nas bordas)
    """
    section = doc.sections[0]
    section._eh_capa = True
    section.top_margin    = Cm(0)
    section.bottom_margin = Cm(0)
    section.left_margin   = Cm(0)
    section.right_margin  = Cm(0)
    section.header_distance = Cm(0)
    section.footer_distance = Cm(0)

    # Margem esquerda ÚNICA da capa (faixa grafite E corpo usam a mesma).
    MARGEM_CAPA = Cm(2.5)
    MARGEM_CAPA_TWIPS = int(2.5 * 567)   # 1 cm = 567 twips → 1418

    largura_total = section.page_width

    # ============================================================
    # Faixa grafite no topo (altura FIXA)
    # ============================================================
    tbl_top = doc.add_table(rows=1, cols=1)
    tbl_top.autofit = False
    tbl_top.allow_autofit = False
    tbl_top.columns[0].width = largura_total
    cel_top = tbl_top.rows[0].cells[0]
    cel_top.width = largura_total
    cel_top.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _shade(cel_top, HEX_GRAFITE_ESC)
    _no_borders(cel_top)
    # margens verticais internas (o recuo esquerdo é feito por left_indent
    # nos parágrafos abaixo — o Word 365 ignora w:tcMar em tabela pct, então
    # não dependemos dele para o recuo horizontal).
    _set_cell_margins(cel_top, top=200, bottom=200, left=120, right=600)
    _fill_table_width_pct(tbl_top, 5000)

    tr = tbl_top.rows[0]._tr
    tr_pr = tr.get_or_add_trPr()
    tr_height = OxmlElement("w:trHeight")
    # "exact": altura fixa que o Word NÃO infla. Com "atLeast" o Word 365
    # estica a faixa muito além do mínimo (LibreOffice não). ~8,5 cm.
    tr_height.set(qn("w:val"), "4800")
    tr_height.set(qn("w:hRule"), "exact")
    tr_pr.append(tr_height)
    _row_cant_split(tbl_top.rows[0])

    # Recuo esquerdo do conteúdo da faixa = mesma margem do corpo.
    IND = Cm(2.5)

    # Conteúdo centralizado verticalmente na faixa (vertical_alignment=CENTER).
    # SÍMBOLO OFICIAL (PNG branco). Fallback: "≡" tipográfico coral.
    p_sim = cel_top.paragraphs[0]
    p_sim.paragraph_format.left_indent = IND
    _set_paragraph_spacing(p_sim, before=0, after=0, line=1.0)
    if SIMBOLO_BRANCO.exists():
        p_sim.add_run().add_picture(str(SIMBOLO_BRANCO), height=Cm(2.6))
    else:
        r_fb = p_sim.add_run("≡")
        r_fb.font.name = FONT_TITULO
        r_fb.font.size = Pt(72)
        r_fb.font.bold = True
        r_fb.font.color.rgb = RGB_CORAL

    # Wordmark branco
    p_assoc = cel_top.add_paragraph()
    p_assoc.paragraph_format.left_indent = IND
    _set_paragraph_spacing(p_assoc, before=14, after=0)
    r_assoc = p_assoc.add_run("SERENA  ENERGIA")
    r_assoc.font.name = FONT_TITULO
    r_assoc.font.size = Pt(12)
    r_assoc.font.bold = True
    r_assoc.font.color.rgb = RGB_BRANCO
    _set_run_letter_spacing(r_assoc, 90)

    # ============================================================
    # Corpo branco — título + meta (mesma margem esquerda da faixa)
    # ============================================================
    IND = MARGEM_CAPA

    for _ in range(2):
        p_blank = doc.add_paragraph()
        _set_paragraph_spacing(p_blank, before=14, after=0)
        p_blank.add_run(" ").font.size = Pt(14)

    # Linha coral curta
    p_linha = doc.add_paragraph()
    p_linha.paragraph_format.left_indent = IND
    _set_paragraph_spacing(p_linha, before=22, after=10)
    r_linha = p_linha.add_run("▬▬")
    r_linha.font.name = FONT_TITULO
    r_linha.font.size = Pt(20)
    r_linha.font.bold = True
    r_linha.font.color.rgb = RGB_CORAL

    # Título — 2 linhas, fonte Serena
    for linha in ["FLEET RELIABILITY", "REPORT"]:
        p_titulo = doc.add_paragraph()
        p_titulo.paragraph_format.left_indent = IND
        _set_paragraph_spacing(p_titulo, before=0, after=0, line=1.04)
        r_titulo = p_titulo.add_run(linha)
        r_titulo.font.name = FONT_TITULO
        r_titulo.font.size = Pt(44)
        r_titulo.font.bold = True
        r_titulo.font.color.rgb = RGB_GRAFITE_ESC

    # Subtítulo
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.left_indent = IND
    _set_paragraph_spacing(p_sub, before=18, after=0)
    r_sub = p_sub.add_run("Assuruá SG  |  G97 · 49 WTGs")
    r_sub.font.name = FONT_BODY
    r_sub.font.size = Pt(16)
    r_sub.font.color.rgb = RGB_GRAFITE

    # Metadados
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.left_indent = IND
    _set_paragraph_spacing(p_meta, before=14, after=0)
    r_meta_sem = p_meta.add_run(f"S{semana:02d} · {periodo_str}")
    r_meta_sem.font.name = FONT_BODY
    r_meta_sem.font.size = Pt(10)
    r_meta_sem.font.bold = True
    r_meta_sem.font.color.rgb = RGB_CORAL
    r_meta_sep = p_meta.add_run("    ·    ")
    r_meta_sep.font.name = FONT_BODY
    r_meta_sep.font.size = Pt(10)
    r_meta_sep.font.color.rgb = RGB_CINZA_CLARO
    r_meta_resp = p_meta.add_run("Responsabilidade SG    ·    SRNA — Gestão de Ativos Serena")
    r_meta_resp.font.name = FONT_BODY
    r_meta_resp.font.size = Pt(10)
    r_meta_resp.font.color.rgb = RGB_CINZA_CLARO


def add_quebra_para_conteudo(doc: DocumentT, semana: int, ano: int) -> None:
    from docx.enum.section import WD_SECTION
    nova = doc.add_section(WD_SECTION.NEW_PAGE)
    nova.top_margin    = Cm(2.0)
    nova.bottom_margin = Cm(1.4)
    nova.left_margin   = Cm(1.4)
    nova.right_margin  = Cm(1.4)
    nova.header_distance = Cm(0.5)
    nova.footer_distance = Cm(0.4)
    _aplicar_header_section(nova, semana, ano)
    _aplicar_footer_section(nova)


# =============================================================================
# CABEÇALHO DE SEÇÃO
# =============================================================================

def add_capitulo(doc: DocumentT, numero: str, titulo: str,
                  subtitulo: str = "") -> None:
    p_h = doc.add_paragraph()
    _set_paragraph_spacing(p_h, before=4, after=0, line=1.15)
    p_h.paragraph_format.left_indent = Cm(0)
    r_num = p_h.add_run(f"{numero}.")
    r_num.font.name = FONT_TITULO
    r_num.font.size = Pt(15)
    r_num.font.bold = True
    r_num.font.color.rgb = RGB_CORAL
    # separador de largura fixa (evita o "buraco" que o Word abre entre runs
    # com letter-spacing diferente)
    r_gap = p_h.add_run("\u2002")   # en-space
    r_gap.font.name = FONT_TITULO
    r_gap.font.size = Pt(15)
    r_tit = p_h.add_run(titulo.upper())
    r_tit.font.name = FONT_TITULO
    r_tit.font.size = Pt(15)
    r_tit.font.bold = True
    r_tit.font.color.rgb = RGB_GRAFITE_ESC
    _set_run_letter_spacing(r_tit, 20)

    if subtitulo:
        p_sub = doc.add_paragraph()
        _set_paragraph_spacing(p_sub, before=4, after=6, line=1.2)
        p_sub.paragraph_format.left_indent = Cm(0)
        r_sub = p_sub.add_run(subtitulo)
        r_sub.font.name = FONT_BODY
        r_sub.font.size = Pt(9)
        r_sub.font.color.rgb = RGB_GRAFITE

    p_linha = doc.add_paragraph()
    _set_paragraph_spacing(p_linha, before=0, after=12)
    p_linha.paragraph_format.left_indent = Cm(0)
    _add_bottom_border(p_linha, color=HEX_CINZA_CLARO, size=4)


# =============================================================================
# KPI CARDS
# =============================================================================

def _rotulo_tem_2_linhas(texto: str, col_w_emu: int) -> bool:
    """Estima se o rótulo (7,5pt bold, CAPS) quebra em 2 linhas na largura
    útil da coluna. Aproximação por contagem de caracteres: em Poppins bold
    7,5pt caps, cabem ~ (largura_cm × 4,4) caracteres por linha.

    Não precisa ser exato — só evita adicionar uma 2ª linha em labels que
    já têm 2. O padding lateral é descontado.
    """
    largura_cm = (col_w_emu / 360000) - 0.5   # EMU→cm menos padding aprox.
    cap = max(6, int(largura_cm * 4.4))
    return len(texto) > cap


def add_kpi_cards(doc: DocumentT, cards: list[dict]) -> None:
    """Linha de KPI cards com ALTURA UNIFORME e valor em fonte Serena.

    Cada card: {"label","valor","sub","cor_rgb"?}
    Os rótulos são normalizados para 2 linhas de altura, garantindo que os
    VALORES fiquem alinhados horizontalmente entre todos os cards.
    """
    if not cards:
        return
    n = len(cards)
    section = doc.sections[-1]
    largura_total = section.page_width - section.left_margin - section.right_margin
    col_w = int(largura_total / n)

    tbl = doc.add_table(rows=1, cols=n)
    tbl.autofit = False
    tbl.allow_autofit = False
    _fix_table_width(tbl, col_w * n)
    for i, col in enumerate(tbl.columns):
        col.width = col_w
    # altura mínima uniforme evita cards "desproporcionais" (Pg 5)
    _row_min_height(tbl.rows[0], 1250)   # ~2,2 cm
    _row_cant_split(tbl.rows[0])

    for i, card in enumerate(cards):
        cel = tbl.rows[0].cells[i]
        cel.width = col_w
        _set_cell_width_pct(cel, int(5000 / n))
        cel.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        # 1º card sem padding esquerdo → alinha exatamente com a margem da
        # página (e com o cabeçalho de seção acima). Demais cards têm recuo
        # pequeno para respeitar a borda divisória à esquerda.
        pad_esq = 0 if i == 0 else 200
        _set_cell_margins(cel, top=140, bottom=140, left=pad_esq, right=150)
        _set_borders(cel, color=HEX_BORDA, size=4,
                     top=False, bottom=False, left=(i > 0), right=False)

        # --- Rótulo com ALTURA FIXA (sempre 2 linhas) ---
        # Assim o valor abaixo começa SEMPRE na mesma altura em todos os
        # cards, independentemente de o label ocupar 1 ou 2 linhas. É isso
        # que mantém os valores ("5", "Gerador", "157,3°C") alinhados.
        p_label = cel.paragraphs[0]
        _set_paragraph_spacing(p_label, before=0, after=3, line=1.05)
        label_txt = card["label"].upper()
        r_label = p_label.add_run(label_txt)
        r_label.font.name = FONT_BODY
        r_label.font.size = Pt(7.5)
        r_label.font.bold = True
        r_label.font.color.rgb = RGB_GRAFITE
        _set_run_letter_spacing(r_label, 20)
        # completa para 2 linhas quando o label é curto (quebra + linha vazia)
        if not _rotulo_tem_2_linhas(label_txt, col_w):
            r_br = p_label.add_run()
            r_br.add_break()
            r_fill = p_label.add_run("\u00a0")   # nbsp p/ dar corpo à 2ª linha
            r_fill.font.name = FONT_BODY
            r_fill.font.size = Pt(7.5)

        # Valor (fonte Serena, destaque)
        p_val = cel.add_paragraph()
        _set_paragraph_spacing(p_val, before=2, after=3, line=1.0)
        r_val = p_val.add_run(card["valor"])
        r_val.font.name = FONT_TITULO
        r_val.font.size = Pt(25)
        r_val.font.bold = True
        r_val.font.color.rgb = card.get("cor_rgb", RGB_CORAL)

        # Sub
        if card.get("sub"):
            p_sub = cel.add_paragraph()
            _set_paragraph_spacing(p_sub, before=2, after=0, line=1.05)
            r_sub = p_sub.add_run(card["sub"])
            r_sub.font.name = FONT_BODY
            r_sub.font.size = Pt(8)
            r_sub.font.color.rgb = RGB_CINZA_CLARO


# =============================================================================
# TABELAS
# =============================================================================

def add_tabela(doc: DocumentT, headers: list[str], rows: list[list[str]],
                col_widths_pct: list[float] | None = None,
                col_status: int | None = None,
                col_destaque_coral: int | None = None,
                first_col_bold: bool = False,
                titulo: str | None = None) -> None:
    if titulo:
        p = doc.add_paragraph()
        _set_paragraph_spacing(p, before=10, after=6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(titulo)
        r.font.name = FONT_TITULO
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = RGB_GRAFITE_ESC

    section = doc.sections[-1]
    largura_total = section.page_width - section.left_margin - section.right_margin

    n_cols = len(headers)
    if col_widths_pct is None:
        col_widths_pct = [1.0 / n_cols] * n_cols
    col_widths_emu = [int(largura_total * p) for p in col_widths_pct]

    tbl = doc.add_table(rows=1, cols=n_cols)
    tbl.autofit = False
    tbl.allow_autofit = False
    for i, col in enumerate(tbl.columns):
        col.width = col_widths_emu[i]
    _fix_table_width(tbl, sum(col_widths_emu))

    # Header
    hdr_row = tbl.rows[0]
    _row_as_header(hdr_row)
    for i, h in enumerate(headers):
        cel = hdr_row.cells[i]
        cel.width = col_widths_emu[i]
        cel.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _shade(cel, HEX_AREIA_HEAD)
        _set_borders(cel, color=HEX_BORDA, size=4,
                     top=False, bottom=True, left=False, right=False)
        _set_cell_margins(cel, top=60, bottom=60, left=110, right=110)
        p = cel.paragraphs[0]
        _set_paragraph_spacing(p, before=0, after=0)
        r = p.add_run(h.upper())
        r.font.name = FONT_BODY
        r.font.size = Pt(7)
        r.font.bold = True
        r.font.color.rgb = RGB_GRAFITE
        _set_run_letter_spacing(r, 6)

    # Linhas
    for ri, row in enumerate(rows):
        tr = tbl.add_row()
        _row_cant_split(tr)
        bg = HEX_AREIA if ri % 2 == 1 else HEX_BRANCO
        for ci, val in enumerate(row):
            cel = tr.cells[ci]
            cel.width = col_widths_emu[ci]
            cel.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _shade(cel, bg)
            _set_borders(cel, color=HEX_BORDA, size=2,
                         top=False, bottom=True, left=False, right=False)
            _set_cell_margins(cel, top=54, bottom=54, left=110, right=110)
            p = cel.paragraphs[0]
            _set_paragraph_spacing(p, before=0, after=0, line=1.0)
            r = p.add_run(str(val) if val is not None else "—")
            r.font.name = FONT_BODY
            r.font.size = Pt(8.5)

            if col_status is not None and ci == col_status:
                status_key = str(val).strip().lower()
                cor_hex, bold = None, False
                for chave, (hex_c, b) in STATUS_COLORS.items():
                    if chave in status_key:
                        cor_hex, bold = hex_c, b
                        break
                if cor_hex:
                    r.font.color.rgb = _hex_to_rgb(cor_hex)
                    r.font.bold = bold
                else:
                    r.font.color.rgb = RGB_GRAFITE
            elif col_destaque_coral is not None and ci == col_destaque_coral:
                r.font.color.rgb = RGB_CORAL
                r.font.bold = True
            elif first_col_bold and ci == 0:
                r.font.color.rgb = RGB_GRAFITE_ESC
                r.font.bold = True
            else:
                r.font.color.rgb = RGB_GRAFITE


# =============================================================================
# BANNER — nota metodológica / alerta
# =============================================================================

def add_banner(doc: DocumentT, texto: str, cor: str = HEX_AMARELO) -> None:
    section = doc.sections[-1]
    largura_total = section.page_width - section.left_margin - section.right_margin

    tbl = doc.add_table(rows=1, cols=2)
    tbl.autofit = False
    tbl.allow_autofit = False
    barra_w = int(Cm(0.14))
    texto_w = largura_total - barra_w
    tbl.columns[0].width = barra_w
    tbl.columns[1].width = texto_w
    _fix_table_width(tbl, largura_total)

    cel_barra = tbl.rows[0].cells[0]
    cel_barra.width = barra_w
    _shade(cel_barra, cor)
    _no_borders(cel_barra)
    _set_cell_margins(cel_barra, top=0, bottom=0, left=0, right=0)
    _set_paragraph_spacing(cel_barra.paragraphs[0], before=0, after=0)
    cel_barra.paragraphs[0].add_run("")

    cel_txt = tbl.rows[0].cells[1]
    cel_txt.width = texto_w
    cel_txt.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _shade(cel_txt, HEX_AREIA_HEAD)   # leve fundo areia p/ separar do corpo
    _no_borders(cel_txt)
    _set_cell_margins(cel_txt, top=90, bottom=90, left=180, right=140)
    p_txt = cel_txt.paragraphs[0]
    _set_paragraph_spacing(p_txt, before=0, after=0, line=1.15)
    r = p_txt.add_run(texto)
    r.font.name = FONT_BODY
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGB_GRAFITE


def add_paragrafo(doc: DocumentT, texto: str, *, size: float = 9,
                   bold: bool = False, color: str = HEX_GRAFITE,
                   before: int = 0, after: int = 4) -> None:
    p = doc.add_paragraph()
    _set_paragraph_spacing(p, before=before, after=after)
    r = p.add_run(texto)
    r.font.name = FONT_BODY
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = _hex_to_rgb(color)


def add_bullet(doc: DocumentT, texto: str, *, size: float = 9) -> None:
    try:
        p = doc.add_paragraph(style="List Bullet")
    except KeyError:
        p = doc.add_paragraph()
        texto = "• " + texto
    _set_paragraph_spacing(p, before=0, after=2)
    for r in list(p.runs):
        r.text = ""
    r = p.add_run(texto)
    r.font.name = FONT_BODY
    r.font.size = Pt(size)
    r.font.color.rgb = RGB_GRAFITE


def add_imagem(doc: DocumentT, path: str, largura_cm: float = 16.5) -> None:
    p = doc.add_paragraph()
    _set_paragraph_spacing(p, before=6, after=10)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=Cm(largura_cm))


def quebra_pagina(doc: DocumentT) -> None:
    p = doc.add_paragraph()
    _page_break(p)