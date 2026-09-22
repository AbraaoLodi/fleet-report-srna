"""Helpers docx — identidade Serena para o Fleet Reliability Report SRNA."""
from __future__ import annotations
from pathlib import Path
from docx import Document
from docx.document import Document as DocumentT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor

_ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
SIMBOLO_BRANCO = _ASSETS_DIR / "serena_symbol_white_trim.png"
LOGO_BRANCO = _ASSETS_DIR / "logo_branco.png"
LOGO_CORAL = _ASSETS_DIR / "logo_coral.png"

# PALETA SERENA
HEX_CORAL="FF5246"; HEX_CORAL_ESCURO="D23532"; HEX_GRAFITE="4B4B4B"
HEX_GRAFITE_ESC="181818"; HEX_VERDE="32CAA0"; HEX_AMARELO="F59E0B"
HEX_CINZA_CLARO="9CA3AF"; HEX_AREIA="EDE9E5"; HEX_AREIA_HEAD="F5F2EF"
HEX_BRANCO="FFFFFF"; HEX_BORDA="E5E3E0"

RGB_CORAL=RGBColor(0xFF,0x52,0x46); RGB_CORAL_ESCURO=RGBColor(0xD2,0x35,0x32)
RGB_GRAFITE=RGBColor(0x4B,0x4B,0x4B); RGB_GRAFITE_ESC=RGBColor(0x18,0x18,0x18)
RGB_VERDE=RGBColor(0x32,0xCA,0xA0); RGB_AMARELO=RGBColor(0xF5,0x9E,0x0B)
RGB_CINZA_CLARO=RGBColor(0x9C,0xA3,0xAF); RGB_BRANCO=RGBColor(0xFF,0xFF,0xFF)

FONT_TITULO="Serena"; FONT_BODY="Poppins"

STATUS_COLORS = {
    "em andamento":(HEX_AMARELO,False),"aberto":(HEX_CORAL,False),
    "concluído":(HEX_VERDE,False),"concluido":(HEX_VERDE,False),
    "agendado":(HEX_CINZA_CLARO,False),"crítico":(HEX_CORAL,True),
    "critico":(HEX_CORAL,True),"atenção":(HEX_AMARELO,True),
    "atencao":(HEX_AMARELO,True),"monitorar":(HEX_VERDE,True),
    "bom":(HEX_VERDE,True),"médio":(HEX_AMARELO,False),"medio":(HEX_AMARELO,False),
    "normal":(HEX_CINZA_CLARO,False),"alto":(HEX_CORAL,True),"baixo":(HEX_CINZA_CLARO,False),
    "pior semana do período":(HEX_CORAL,True),"melhor semana do período":(HEX_VERDE,True),
    "acima da meta":(HEX_VERDE,True),"próximo da meta":(HEX_AMARELO,False),
    "abaixo da meta":(HEX_CINZA_CLARO,False),"sensor c/ falha":(HEX_CINZA_CLARO,True),
}

def _shade(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),"clear"); shd.set(qn("w:color"),"auto"); shd.set(qn("w:fill"),hex_color)
    tc_pr.append(shd)

def _set_borders(cell,color=HEX_BORDA,size=4,top=True,bottom=True,left=False,right=False):
    tc_pr = cell._tc.get_or_add_tcPr()
    tb = OxmlElement("w:tcBorders")
    for side,en in [("top",top),("bottom",bottom),("left",left),("right",right)]:
        el = OxmlElement(f"w:{side}")
        if en:
            el.set(qn("w:val"),"single"); el.set(qn("w:sz"),str(size)); el.set(qn("w:color"),color)
        else: el.set(qn("w:val"),"nil")
        tb.append(el)
    tc_pr.append(tb)

def _no_borders(cell): _set_borders(cell,top=False,bottom=False,left=False,right=False)

def _row_cant_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    c = OxmlElement("w:cantSplit"); c.set(qn("w:val"),"true"); tr_pr.append(c)

def _row_as_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader"); th.set(qn("w:val"),"true"); tr_pr.append(th)
    _row_cant_split(row)

def _row_height(row, twips, rule="atLeast"):
    tr_pr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight"); h.set(qn("w:val"),str(twips)); h.set(qn("w:hRule"),rule)
    tr_pr.append(h)

def _set_cell_width_pct(cell, pct):
    tcPr = cell._tc.get_or_add_tcPr()
    ex = tcPr.find(qn("w:tcW"))
    if ex is not None: tcPr.remove(ex)
    w = OxmlElement("w:tcW"); w.set(qn("w:w"),str(pct)); w.set(qn("w:type"),"pct"); tcPr.append(w)

def _fill_table_width_pct(table, pct=5000):
    tblPr = table._tbl.tblPr
    for tag in ("w:tblW","w:tblInd"):
        ex = tblPr.find(qn(tag))
        if ex is not None: tblPr.remove(ex)
    w = OxmlElement("w:tblW"); w.set(qn("w:w"),str(pct)); w.set(qn("w:type"),"pct"); tblPr.append(w)
    ind = OxmlElement("w:tblInd"); ind.set(qn("w:w"),"0"); ind.set(qn("w:type"),"dxa"); tblPr.append(ind)

def _fix_table_width(table, largura_emu):
    twips = int(largura_emu/635)
    tblPr = table._tbl.tblPr
    for tag in ("w:tblW","w:tblInd"):
        ex = tblPr.find(qn(tag))
        if ex is not None: tblPr.remove(ex)
    w = OxmlElement("w:tblW"); w.set(qn("w:w"),str(twips)); w.set(qn("w:type"),"dxa"); tblPr.append(w)
    ind = OxmlElement("w:tblInd"); ind.set(qn("w:w"),"0"); ind.set(qn("w:type"),"dxa"); tblPr.append(ind)
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout"); tblPr.append(layout)
    layout.set(qn("w:type"),"fixed")

def _set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc_pr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for side,val in [("top",top),("bottom",bottom),("left",left),("right",right)]:
        el = OxmlElement(f"w:{side}"); el.set(qn("w:w"),str(val)); el.set(qn("w:type"),"dxa"); mar.append(el)
    tc_pr.append(mar)

def _page_break(paragraph):
    paragraph.add_run().add_break(WD_BREAK.PAGE)

def _sp(p, before=0, after=0, line=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after)
    if line is not None: pf.line_spacing = line

def _letter_spacing(run, val):
    r_pr = run._element.get_or_add_rPr()
    sp = OxmlElement("w:spacing"); sp.set(qn("w:val"),str(val)); r_pr.append(sp)

def _bottom_border(paragraph, color=HEX_CINZA_CLARO, size=4):
    p_pr = paragraph._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr"); b = OxmlElement("w:bottom")
    b.set(qn("w:val"),"single"); b.set(qn("w:sz"),str(size)); b.set(qn("w:space"),"1"); b.set(qn("w:color"),color)
    bdr.append(b); p_pr.append(bdr)

def _hex_to_rgb(h): return RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))

def configurar_pagina(doc, top=0.0, bottom=0.7, left=1.5, right=1.5, header=0.0, footer=0.5):
    s = doc.sections[-1]
    s.top_margin=Cm(top); s.bottom_margin=Cm(bottom); s.left_margin=Cm(left)
    s.right_margin=Cm(right); s.header_distance=Cm(header); s.footer_distance=Cm(footer)

# =============================================================================
# HEADER E RODAPÉ
# =============================================================================
def _aplicar_header_section(section, semana, ano, frota_label, rotulo_semana=None):
    _rot = rotulo_semana if rotulo_semana else f"S{semana:02d}"
    section.header_distance = Cm(0.5)
    section.top_margin = Cm(2.2)
    header = section.header
    header.is_linked_to_previous = False
    for p in list(header.paragraphs):
        if p.text: p.text = ""
    largura = section.page_width - section.left_margin - section.right_margin
    tbl = header.add_table(rows=1, cols=2, width=largura)
    tbl.autofit=False; tbl.allow_autofit=False; tbl.alignment=WD_TABLE_ALIGNMENT.LEFT
    col_esq = int(largura*0.78); col_dir = largura-col_esq
    tbl.columns[0].width=col_esq; tbl.columns[1].width=col_dir
    _fill_table_width_pct(tbl,5000)

    ce = tbl.rows[0].cells[0]; ce.width=col_esq
    _set_cell_width_pct(ce,3900); _shade(ce,HEX_GRAFITE_ESC); _no_borders(ce)
    _set_cell_margins(ce,top=110,bottom=110,left=300,right=140)
    ce.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p1 = ce.paragraphs[0]; _sp(p1,0,0)
    rs = p1.add_run("≡  "); rs.font.name=FONT_TITULO; rs.font.size=Pt(11); rs.font.bold=True; rs.font.color.rgb=RGB_CORAL
    rt = p1.add_run(f"Fleet Reliability Report · {frota_label}")
    rt.font.name=FONT_TITULO; rt.font.size=Pt(10); rt.font.bold=True; rt.font.color.rgb=RGB_BRANCO
    p2 = ce.add_paragraph(); _sp(p2,2,0)
    rsub = p2.add_run("Complexo Eólico Assuruá · Gestão de Ativos SRNA")
    rsub.font.name=FONT_BODY; rsub.font.size=Pt(8); rsub.font.color.rgb=RGBColor(0x9C,0xA3,0xAF)

    cd = tbl.rows[0].cells[1]; cd.width=col_dir
    _set_cell_width_pct(cd,1100); cd.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
    _shade(cd,HEX_CORAL); _no_borders(cd); _set_cell_margins(cd,top=110,bottom=110,left=140,right=280)
    pb = cd.paragraphs[0]; _sp(pb,0,0); pb.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    rb = pb.add_run(f"{_rot} / {ano}")
    rb.font.name=FONT_TITULO; rb.font.size=Pt(12); rb.font.bold=True; rb.font.color.rgb=RGB_BRANCO

def _aplicar_footer_section(section):
    footer = section.footer
    footer.is_linked_to_previous = False
    for p in list(footer.paragraphs):
        if p.text: p.text = ""
    p = footer.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.LEFT; _sp(p,0,0)
    r = p.add_run("Serena Energia — Gestão de Ativos SRNA · Documento interno")
    r.font.name=FONT_BODY; r.font.size=Pt(8); r.font.color.rgb=RGB_CINZA_CLARO

# =============================================================================
# CAPA FULLPAGE
# =============================================================================
def add_capa(doc, semana, ano, periodo_str, frota_titulo, frota_sub, rotulo_semana=None):
    _rot_capa = rotulo_semana if rotulo_semana else f"S{semana:02d}"
    section = doc.sections[0]
    section._eh_capa = True
    section.top_margin=Cm(0); section.bottom_margin=Cm(0)
    section.left_margin=Cm(0); section.right_margin=Cm(0)
    section.header_distance=Cm(0); section.footer_distance=Cm(0)
    largura = section.page_width

    tbl_top = doc.add_table(rows=1, cols=1)
    tbl_top.autofit=False; tbl_top.allow_autofit=False
    tbl_top.columns[0].width=largura
    ct = tbl_top.rows[0].cells[0]; ct.width=largura
    ct.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
    _shade(ct,HEX_GRAFITE_ESC); _no_borders(ct)
    _set_cell_margins(ct,top=200,bottom=200,left=120,right=600)
    _fill_table_width_pct(tbl_top,5000)
    tr = tbl_top.rows[0]._tr; tr_pr = tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight"); h.set(qn("w:val"),"5000"); h.set(qn("w:hRule"),"exact"); tr_pr.append(h)
    _row_cant_split(tbl_top.rows[0])

    IND = Cm(2.5)
    # Logo oficial branco (fallback: símbolo)
    p_logo = ct.paragraphs[0]; p_logo.paragraph_format.left_indent=IND; _sp(p_logo,0,0,1.0)
    if LOGO_BRANCO.exists():
        p_logo.add_run().add_picture(str(LOGO_BRANCO), height=Cm(1.9))
    elif SIMBOLO_BRANCO.exists():
        p_logo.add_run().add_picture(str(SIMBOLO_BRANCO), height=Cm(2.4))
    else:
        rfb = p_logo.add_run("≡"); rfb.font.name=FONT_TITULO; rfb.font.size=Pt(64); rfb.font.bold=True; rfb.font.color.rgb=RGB_CORAL

    IND = Cm(2.5)
    for _ in range(2):
        pb = doc.add_paragraph(); _sp(pb,14,0); pb.add_run(" ").font.size=Pt(14)
    p_linha = doc.add_paragraph(); p_linha.paragraph_format.left_indent=IND; _sp(p_linha,22,10)
    rl = p_linha.add_run("▬▬"); rl.font.name=FONT_TITULO; rl.font.size=Pt(20); rl.font.bold=True; rl.font.color.rgb=RGB_CORAL

    for linha in ["FLEET RELIABILITY","REPORT"]:
        pt = doc.add_paragraph(); pt.paragraph_format.left_indent=IND; _sp(pt,0,0,1.04)
        rt = pt.add_run(linha); rt.font.name=FONT_TITULO; rt.font.size=Pt(44); rt.font.bold=True; rt.font.color.rgb=RGB_GRAFITE_ESC

    ps = doc.add_paragraph(); ps.paragraph_format.left_indent=IND; _sp(ps,18,0)
    rsub = ps.add_run(frota_titulo); rsub.font.name=FONT_BODY; rsub.font.size=Pt(16); rsub.font.color.rgb=RGB_GRAFITE

    pm = doc.add_paragraph(); pm.paragraph_format.left_indent=IND; _sp(pm,14,0)
    rms = pm.add_run(f"{_rot_capa} · {periodo_str}")
    rms.font.name=FONT_BODY; rms.font.size=Pt(10); rms.font.bold=True; rms.font.color.rgb=RGB_CORAL
    rsep = pm.add_run("    ·    "); rsep.font.name=FONT_BODY; rsep.font.size=Pt(10); rsep.font.color.rgb=RGB_CINZA_CLARO
    rmr = pm.add_run(frota_sub); rmr.font.name=FONT_BODY; rmr.font.size=Pt(10); rmr.font.color.rgb=RGB_CINZA_CLARO

def add_quebra_para_conteudo(doc, semana, ano, frota_label, rotulo_semana=None):
    from docx.enum.section import WD_SECTION
    nova = doc.add_section(WD_SECTION.NEW_PAGE)
    nova.top_margin=Cm(2.2); nova.bottom_margin=Cm(1.4)
    nova.left_margin=Cm(1.4); nova.right_margin=Cm(1.4)
    nova.header_distance=Cm(0.5); nova.footer_distance=Cm(0.4)
    _aplicar_header_section(nova, semana, ano, frota_label, rotulo_semana)
    _aplicar_footer_section(nova)

# =============================================================================
# CABEÇALHO DE SEÇÃO
# =============================================================================
def add_capitulo(doc, numero, titulo, subtitulo="", quebra_antes=False):
    ph = doc.add_paragraph(); _sp(ph,4,0,1.15); ph.paragraph_format.left_indent=Cm(0)
    if quebra_antes:
        ph.paragraph_format.page_break_before = True
    rn = ph.add_run(f"{numero}."); rn.font.name=FONT_TITULO; rn.font.size=Pt(15); rn.font.bold=True; rn.font.color.rgb=RGB_CORAL
    rg = ph.add_run("\u2002"); rg.font.name=FONT_TITULO; rg.font.size=Pt(15)
    rt = ph.add_run(titulo.upper()); rt.font.name=FONT_TITULO; rt.font.size=Pt(15); rt.font.bold=True; rt.font.color.rgb=RGB_GRAFITE_ESC
    _letter_spacing(rt,20)
    if subtitulo:
        psb = doc.add_paragraph(); _sp(psb,4,6,1.2); psb.paragraph_format.left_indent=Cm(0)
        rsb = psb.add_run(subtitulo); rsb.font.name=FONT_BODY; rsb.font.size=Pt(9); rsb.font.color.rgb=RGB_GRAFITE
    pl = doc.add_paragraph(); _sp(pl,0,12); pl.paragraph_format.left_indent=Cm(0)
    _bottom_border(pl, color=HEX_CINZA_CLARO, size=4)

# =============================================================================
# KPI CARDS
# =============================================================================
def _rotulo_2_linhas(texto, col_w_emu):
    largura_cm = (col_w_emu/360000) - 0.5
    cap = max(6, int(largura_cm*4.4))
    return len(texto) > cap

def add_kpi_cards(doc, cards):
    if not cards: return
    n = len(cards)
    section = doc.sections[-1]
    largura = section.page_width - section.left_margin - section.right_margin
    col_w = int(largura/n)
    tbl = doc.add_table(rows=1, cols=n)
    tbl.autofit=False; tbl.allow_autofit=False
    _fix_table_width(tbl, col_w*n)
    for col in tbl.columns: col.width=col_w
    _row_height(tbl.rows[0], 1300, "atLeast")
    _row_cant_split(tbl.rows[0])
    for i,card in enumerate(cards):
        cel = tbl.rows[0].cells[i]; cel.width=col_w
        _set_cell_width_pct(cel, int(5000/n))
        cel.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        pad = 0 if i==0 else 220
        _set_cell_margins(cel,top=150,bottom=150,left=pad,right=160)
        _set_borders(cel,color=HEX_BORDA,size=4,top=False,bottom=False,left=(i>0),right=False)
        pl = cel.paragraphs[0]; _sp(pl,0,3,1.05)
        lbl = card["label"].upper()
        rl = pl.add_run(lbl); rl.font.name=FONT_BODY; rl.font.size=Pt(7.5); rl.font.bold=True; rl.font.color.rgb=RGB_GRAFITE
        _letter_spacing(rl,20)
        if not _rotulo_2_linhas(lbl, col_w):
            pl.add_run().add_break()
            rf = pl.add_run("\u00a0"); rf.font.name=FONT_BODY; rf.font.size=Pt(7.5)
        pv = cel.add_paragraph(); _sp(pv,2,3,1.0)
        rv = pv.add_run(card["valor"]); rv.font.name=FONT_TITULO; rv.font.size=Pt(25); rv.font.bold=True
        rv.font.color.rgb = card.get("cor_rgb", RGB_CORAL)
        if card.get("sub"):
            psb = cel.add_paragraph(); _sp(psb,2,0,1.05)
            rsb = psb.add_run(card["sub"]); rsb.font.name=FONT_BODY; rsb.font.size=Pt(8); rsb.font.color.rgb=RGB_CINZA_CLARO

# =============================================================================
# TABELAS
# =============================================================================
def add_tabela(doc, headers, rows, col_widths_pct=None, col_status=None,
               col_destaque_coral=None, first_col_bold=False, titulo=None):
    if titulo:
        p = doc.add_paragraph(); _sp(p,10,6); p.paragraph_format.keep_with_next=True
        r = p.add_run(titulo); r.font.name=FONT_TITULO; r.font.size=Pt(10.5); r.font.bold=True; r.font.color.rgb=RGB_GRAFITE_ESC
    section = doc.sections[-1]
    largura = section.page_width - section.left_margin - section.right_margin
    n_cols = len(headers)
    if col_widths_pct is None: col_widths_pct=[1.0/n_cols]*n_cols
    widths = [int(largura*p) for p in col_widths_pct]
    tbl = doc.add_table(rows=1, cols=n_cols)
    tbl.autofit=False; tbl.allow_autofit=False
    for i,col in enumerate(tbl.columns): col.width=widths[i]
    _fix_table_width(tbl, sum(widths))
    hdr = tbl.rows[0]; _row_as_header(hdr)
    for i,h in enumerate(headers):
        cel = hdr.cells[i]; cel.width=widths[i]; cel.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
        _shade(cel,HEX_AREIA_HEAD)
        _set_borders(cel,color=HEX_BORDA,size=4,top=False,bottom=True,left=False,right=False)
        _set_cell_margins(cel,top=60,bottom=60,left=110,right=110)
        p = cel.paragraphs[0]; _sp(p,0,0)
        r = p.add_run(h.upper()); r.font.name=FONT_BODY; r.font.size=Pt(7); r.font.bold=True; r.font.color.rgb=RGB_GRAFITE
        _letter_spacing(r,6)
    for ri,row in enumerate(rows):
        tr = tbl.add_row(); _row_cant_split(tr)
        bg = HEX_AREIA if ri%2==1 else HEX_BRANCO
        for ci,val in enumerate(row):
            cel = tr.cells[ci]; cel.width=widths[ci]; cel.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
            _shade(cel,bg)
            _set_borders(cel,color=HEX_BORDA,size=2,top=False,bottom=True,left=False,right=False)
            _set_cell_margins(cel,top=54,bottom=54,left=110,right=110)
            p = cel.paragraphs[0]; _sp(p,0,0,1.0)
            r = p.add_run(str(val) if val is not None else "—"); r.font.name=FONT_BODY; r.font.size=Pt(8.5)
            if col_status is not None and ci==col_status:
                key = str(val).strip().lower(); cor,bold=None,False
                for chave,(hc,b) in STATUS_COLORS.items():
                    if chave in key: cor,bold=hc,b; break
                r.font.color.rgb = _hex_to_rgb(cor) if cor else RGB_GRAFITE
                r.font.bold = bold
            elif col_destaque_coral is not None and ci==col_destaque_coral:
                r.font.color.rgb=RGB_CORAL; r.font.bold=True
            elif first_col_bold and ci==0:
                r.font.color.rgb=RGB_GRAFITE_ESC; r.font.bold=True
            else:
                r.font.color.rgb=RGB_GRAFITE

# =============================================================================
# BANNER / PARÁGRAFO / IMAGEM
# =============================================================================
def add_banner(doc, texto, cor=HEX_AMARELO):
    section = doc.sections[-1]
    largura = section.page_width - section.left_margin - section.right_margin
    tbl = doc.add_table(rows=1, cols=2)
    tbl.autofit=False; tbl.allow_autofit=False
    barra_w = int(Cm(0.14)); texto_w = largura-barra_w
    tbl.columns[0].width=barra_w; tbl.columns[1].width=texto_w
    _fix_table_width(tbl, largura)
    cb = tbl.rows[0].cells[0]; cb.width=barra_w
    _shade(cb,cor); _no_borders(cb); _set_cell_margins(cb,0,0,0,0)
    _sp(cb.paragraphs[0],0,0); cb.paragraphs[0].add_run("")
    ct = tbl.rows[0].cells[1]; ct.width=texto_w; ct.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
    _shade(ct,HEX_AREIA_HEAD); _no_borders(ct); _set_cell_margins(ct,top=90,bottom=90,left=180,right=140)
    p = ct.paragraphs[0]; _sp(p,0,0,1.15)
    r = p.add_run(texto); r.font.name=FONT_BODY; r.font.size=Pt(8.5); r.font.color.rgb=RGB_GRAFITE

def add_paragrafo(doc, texto, size=9, bold=False, color=HEX_GRAFITE, before=0, after=4):
    p = doc.add_paragraph(); _sp(p,before,after)
    r = p.add_run(texto); r.font.name=FONT_BODY; r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=_hex_to_rgb(color)

def add_imagem(doc, path, largura_cm=16.5):
    p = doc.add_paragraph(); _sp(p,6,10); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=Cm(largura_cm))

def quebra_pagina(doc):
    _page_break(doc.add_paragraph())

# =============================================================================
# TABELA HIERÁRQUICA (WTG → alarmes)
# =============================================================================
def add_tabela_hierarquica(doc, grupos, headers, col_widths_pct, titulo=None):
    """grupos: [{'wtg':str,'h_total':float,'linhas':[[c1..cn],...]}]"""
    if titulo:
        p = doc.add_paragraph(); _sp(p,10,6); p.paragraph_format.keep_with_next=True
        r = p.add_run(titulo); r.font.name=FONT_TITULO; r.font.size=Pt(10.5); r.font.bold=True; r.font.color.rgb=RGB_GRAFITE_ESC
    section = doc.sections[-1]
    largura = section.page_width - section.left_margin - section.right_margin
    n_cols = len(headers)
    widths = [int(largura*p) for p in col_widths_pct]
    tbl = doc.add_table(rows=1, cols=n_cols)
    tbl.autofit=False; tbl.allow_autofit=False
    for i,col in enumerate(tbl.columns): col.width=widths[i]
    _fix_table_width(tbl, sum(widths))
    hdr = tbl.rows[0]; _row_as_header(hdr)
    for i,h in enumerate(headers):
        cel = hdr.cells[i]; cel.width=widths[i]; cel.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
        _shade(cel,HEX_AREIA_HEAD); _set_borders(cel,color=HEX_BORDA,size=4,bottom=True)
        _set_cell_margins(cel,top=60,bottom=60,left=110,right=110)
        p = cel.paragraphs[0]; _sp(p,0,0)
        r = p.add_run(h.upper()); r.font.name=FONT_BODY; r.font.size=Pt(7); r.font.bold=True; r.font.color.rgb=RGB_GRAFITE
        _letter_spacing(r,6)
    from utils.formatters import fmt_float
    for gi,grupo in enumerate(grupos):
        # Linha de grupo (WTG)
        tr = tbl.add_row(); _row_cant_split(tr)
        c0 = tr.cells[0]
        # mescla todas as colunas menos a última (que mostra o total)
        c_merge = c0
        for ci in range(1, n_cols-1):
            c_merge = c_merge.merge(tr.cells[ci])
        for cel in (c_merge, tr.cells[n_cols-1]):
            _shade(cel,HEX_GRAFITE_ESC); _set_borders(cel,color=HEX_GRAFITE_ESC,size=4,top=True,bottom=True)
            _set_cell_margins(cel,top=52,bottom=52,left=110,right=110)
            cel.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
        pg = c_merge.paragraphs[0]; _sp(pg,0,0)
        rb = pg.add_run("▸ "); rb.font.name=FONT_BODY; rb.font.size=Pt(9); rb.font.bold=True; rb.font.color.rgb=RGB_CORAL
        rw = pg.add_run(str(grupo["wtg"])); rw.font.name=FONT_BODY; rw.font.size=Pt(9); rw.font.bold=True; rw.font.color.rgb=RGB_BRANCO
        ph = tr.cells[n_cols-1].paragraphs[0]; _sp(ph,0,0); ph.alignment=WD_ALIGN_PARAGRAPH.LEFT
        rh = ph.add_run(fmt_float(grupo["h_total"])+" h"); rh.font.name=FONT_BODY; rh.font.size=Pt(9); rh.font.bold=True; rh.font.color.rgb=RGB_CORAL
        # Linhas de alarme
        for li,linha in enumerate(grupo["linhas"]):
            trl = tbl.add_row(); _row_cant_split(trl)
            bg = HEX_AREIA if li%2==1 else HEX_BRANCO
            for ci,val in enumerate(linha):
                cel = trl.cells[ci]; cel.width=widths[ci]; cel.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
                _shade(cel,bg); _set_borders(cel,color=HEX_BORDA,size=2,bottom=True)
                pad_l = 260 if ci==0 else 110
                _set_cell_margins(cel,top=48,bottom=48,left=pad_l,right=110)
                p = cel.paragraphs[0]; _sp(p,0,0,1.0)
                r = p.add_run(str(val)); r.font.name=FONT_BODY; r.font.size=Pt(8)
                if ci==0: r.font.color.rgb=RGB_GRAFITE_ESC
                elif ci==n_cols-1: r.font.color.rgb=RGB_CORAL; r.font.bold=True
                else: r.font.color.rgb=RGB_GRAFITE

# =============================================================================
# TABELA HIERÁRQUICA COM STATUS (temperatura: WTG → componentes + severidade)
# =============================================================================
def add_tabela_hierarquica_temp(doc, grupos, headers, col_widths_pct, col_status=None, titulo=None):
    """Igual à hierárquica, mas a última coluna pode receber cor de status.
    grupos: [{'wtg':str,'h_total':float(pico),'linhas':[[c1..cn],...]}]
    O cabeçalho da WTG mostra o pico máximo em coral na última coluna."""
    if titulo:
        p = doc.add_paragraph(); _sp(p,10,6); p.paragraph_format.keep_with_next=True
        r = p.add_run(titulo); r.font.name=FONT_TITULO; r.font.size=Pt(10.5); r.font.bold=True; r.font.color.rgb=RGB_GRAFITE_ESC
    section = doc.sections[-1]
    largura = section.page_width - section.left_margin - section.right_margin
    n_cols = len(headers)
    widths = [int(largura*p) for p in col_widths_pct]
    tbl = doc.add_table(rows=1, cols=n_cols)
    tbl.autofit=False; tbl.allow_autofit=False
    for i,col in enumerate(tbl.columns): col.width=widths[i]
    _fix_table_width(tbl, sum(widths))
    hdr = tbl.rows[0]; _row_as_header(hdr)
    for i,h in enumerate(headers):
        cel = hdr.cells[i]; cel.width=widths[i]; cel.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
        _shade(cel,HEX_AREIA_HEAD); _set_borders(cel,color=HEX_BORDA,size=4,bottom=True)
        _set_cell_margins(cel,top=60,bottom=60,left=110,right=110)
        p = cel.paragraphs[0]; _sp(p,0,0)
        r = p.add_run(h.upper()); r.font.name=FONT_BODY; r.font.size=Pt(7); r.font.bold=True; r.font.color.rgb=RGB_GRAFITE
        _letter_spacing(r,6)
    for grupo in grupos:
        # Linha de grupo (WTG) — mescla col 0 até penúltima; última = pico máx
        tr = tbl.add_row(); _row_cant_split(tr)
        c_merge = tr.cells[0]
        for ci in range(1, n_cols-1):
            c_merge = c_merge.merge(tr.cells[ci])
        for cel in (c_merge, tr.cells[n_cols-1]):
            _shade(cel,HEX_GRAFITE_ESC); _set_borders(cel,color=HEX_GRAFITE_ESC,size=4,top=True,bottom=True)
            _set_cell_margins(cel,top=52,bottom=52,left=110,right=110)
            cel.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
        pg = c_merge.paragraphs[0]; _sp(pg,0,0)
        rb = pg.add_run("▸ "); rb.font.name=FONT_BODY; rb.font.size=Pt(9); rb.font.bold=True; rb.font.color.rgb=RGB_CORAL
        rw = pg.add_run(str(grupo["wtg"])); rw.font.name=FONT_BODY; rw.font.size=Pt(9); rw.font.bold=True; rw.font.color.rgb=RGB_BRANCO
        ph = tr.cells[n_cols-1].paragraphs[0]; _sp(ph,0,0); ph.alignment=WD_ALIGN_PARAGRAPH.LEFT
        from utils.formatters import fmt_float
        rh = ph.add_run("pico "+fmt_float(grupo["h_total"])+" °C")
        rh.font.name=FONT_BODY; rh.font.size=Pt(8.5); rh.font.bold=True; rh.font.color.rgb=RGB_CORAL
        # Linhas de componente
        for li,linha in enumerate(grupo["linhas"]):
            trl = tbl.add_row(); _row_cant_split(trl)
            bg = HEX_AREIA if li%2==1 else HEX_BRANCO
            for ci,val in enumerate(linha):
                cel = trl.cells[ci]; cel.width=widths[ci]; cel.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
                _shade(cel,bg); _set_borders(cel,color=HEX_BORDA,size=2,bottom=True)
                pad_l = 260 if ci==0 else 110
                _set_cell_margins(cel,top=48,bottom=48,left=pad_l,right=110)
                p = cel.paragraphs[0]; _sp(p,0,0,1.0)
                r = p.add_run(str(val)); r.font.name=FONT_BODY; r.font.size=Pt(8)
                if col_status is not None and ci==col_status:
                    key = str(val).strip().lower(); cor,bold=None,False
                    for chave,(hc,b) in STATUS_COLORS.items():
                        if chave in key: cor,bold=hc,b; break
                    r.font.color.rgb = _hex_to_rgb(cor) if cor else RGB_GRAFITE
                    r.font.bold = bold
                elif ci==0:
                    r.font.color.rgb=RGB_GRAFITE_ESC
                else:
                    r.font.color.rgb=RGB_GRAFITE
