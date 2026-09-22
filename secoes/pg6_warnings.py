"""Pg 6 — 07. Alarmes & Avisos do SCADA (fabricante) — janela de 2 semanas."""
from __future__ import annotations
import tempfile
from pathlib import Path
from utils.charts import grafico_barras_warnings, CORAL, GRAFITE, AMARELO
from utils.docx_helpers import (add_capitulo, add_kpi_cards, add_tabela, add_imagem,
    quebra_pagina, add_paragrafo, add_banner, RGB_CORAL, HEX_CINZA_CLARO)

def _wtgs_str(lista, maxn=6):
    if not lista: return "—"
    s = ", ".join(lista[:maxn])
    if len(lista) > maxn: s += f" +{len(lista)-maxn}"
    return s

def construir_pg6(doc, warn, periodo_str=""):
    if not warn:
        add_capitulo(doc, "06", "Alarmes & Avisos SCADA",
            f"Eventos do sistema do fabricante · janela {periodo_str} · alarmes críticos e avisos recorrentes")
        add_paragrafo(doc, "Arquivo de warnings não fornecido para esta semana.", size=10, color=HEX_CINZA_CLARO)
        quebra_pagina(doc); return
    ce = warn["col_evt"]; cd = warn["col_disp"]
    tem_alarmes = bool(warn["rank_alarmes"])

    # --- Alarmes: só monta a seção 06 se houver eventos de categoria Alarme.
    # Alguns fabricantes (ex.: este export da Vestas) não distinguem
    # Alarme/Aviso -- nesse caso tudo entra como Aviso e a seção 06 é pulada,
    # a de Avisos assume o número 06 para não deixar uma seção vazia no meio.
    num_avisos = "07"
    if tem_alarmes:
        add_capitulo(doc, "06", "Alarmes & Avisos SCADA",
            f"Eventos do sistema do fabricante · janela {periodo_str} · alarmes críticos e avisos recorrentes")
        labels = [f"{a['cod']} {a['desc'][:34]}" for a in warn["rank_alarmes"][:8]]
        vals = [a["ocor"] for a in warn["rank_alarmes"][:8]]
        img = Path(tempfile.gettempdir())/"g_alarmes.png"
        grafico_barras_warnings(labels, vals, cores=[CORAL]*len(labels), out_path=str(img),
            titulo="Alarmes críticos por ocorrência", eixo_x="Ocorrências")
        add_imagem(doc, str(img), 16.2)
        rows = []
        for a in warn["rank_alarmes"]:
            wl = warn["wtgs_por_alarme"].get(a[ce], [])
            rows.append([a["cod"], a["desc"], str(a["ocor"]), str(a["wtgs"]), _wtgs_str(wl, 5)])
        add_tabela(doc, ["Cód.","Alarme","Ocor.","WTGs","Máquinas afetadas"], rows,
            col_widths_pct=[0.08,0.40,0.08,0.08,0.36], col_destaque_coral=2, first_col_bold=True,
            titulo="Alarmes críticos mais frequentes")
    else:
        num_avisos = "06"

    # --- Avisos: gráfico de WTGs no topo, tabela de avisos, banner ---
    add_capitulo(doc, num_avisos, "Avisos Recorrentes",
        "Warnings de maior recorrência — indicadores de degradação e alvos de manutenção preventiva",
        quebra_antes=tem_alarmes)
    if warn["wtg_alarmes"]:
        labels = [str(w[cd]) for w in warn["wtg_alarmes"][:8]]
        vals = [w["n_alarmes"] for w in warn["wtg_alarmes"][:8]]
        img2 = Path(tempfile.gettempdir())/"g_wtg_alarmes.png"
        grafico_barras_warnings(labels, vals, cores=[GRAFITE]*len(labels), out_path=str(img2),
            titulo="WTGs com mais avisos" if not tem_alarmes else "WTGs com mais alarmes críticos",
            eixo_x="Nº de eventos" if not tem_alarmes else "Nº de alarmes")
        add_imagem(doc, str(img2), 14.5)
    if warn["rank_avisos"]:
        rows = []
        for a in warn["rank_avisos"]:
            wl = warn["wtgs_por_aviso"].get(a[ce], [])
            rows.append([a["cod"], a["desc"], f"{a['ocor']:,}".replace(",","."), str(a["wtgs"]), _wtgs_str(wl, 5)])
        add_tabela(doc, ["Cód.","Aviso","Ocor.","WTGs","Máquinas afetadas"], rows,
            col_widths_pct=[0.08,0.40,0.09,0.07,0.36], col_destaque_coral=2, first_col_bold=True,
            titulo="Avisos mais recorrentes")
        # banner dinâmico: usa o aviso que de fato lidera este relatório,
        # não um código fixo (evita afirmar algo que não é verdade para a frota/janela)
        top = warn["rank_avisos"][0]
        n_wtgs_top = warn["wtg_alarmes"][0][cd] if warn["wtg_alarmes"] else None
        add_banner(doc, f"O aviso {top['cod']} ({top['desc'][:70]}) é o mais recorrente no período, com "
            f"{top['ocor']:,} ocorrência(s) em {top['wtgs']} WTG(s). Recomenda-se investigar a causa raiz "
            f"nas máquinas mais afetadas antes que evolua para parada não planejada.".replace(",","."), cor="F59E0B")