"""Orquestrador — Fleet Reliability Report SRNA.

Uso: python gerar_fleet_report.py <semana> <ano> <frota> [csv_warnings]
Ex:  python gerar_fleet_report.py 28 2026 SG C:\\caminho\\AWS.csv
     python gerar_fleet_report.py 29-30 2026 VE   (janela quinzenal)

Executa as 7 queries direto no Snowflake via SSO (data/connection.py) e monta
o .docx com a identidade Serena.
"""
from __future__ import annotations
import sys
from datetime import date, timedelta
from pathlib import Path
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parent))
from config.settings import FROTAS, resolver_frota, META_DISP
from utils.semana_sg import janela_da_semana
from utils.fonts import registrar_fontes
from utils.docx_helpers import configurar_pagina, add_capa, add_quebra_para_conteudo
from utils.logger import get_logger
from data.connection import run_query          # <<< SUA conexão SSO (não sobe ao GitHub)
from data.queries import QUERIES
from data.warnings_parser import processar_warnings
from secoes.pg2_disponibilidade import construir_pg2
from secoes.pg3_alarmes import construir_pg3
from secoes.pg4_mtbf import construir_pg4
from secoes.pg5_temperatura import construir_pg5
from secoes.pg6_warnings import construir_pg6

log = get_logger("fleet.orquestrador")

# Ajuste este caminho para a pasta de saída da sua máquina:
OUT_DIR_PADRAO = r"C:\Users\abraão.lodi\OneDrive - Serena\Área de Trabalho\Downloads\dev\fleet_report_srna_v3\output"


def _filtro_exclusao(frota) -> str:
    """Monta a cláusula AND ... NOT IN (...) de exclusão de alarme (só GE tem)."""
    if not frota.exclusao_alarmes:
        return ""
    lista = ",".join("'" + a.replace("'", "''") + "'" for a in frota.exclusao_alarmes)
    return f"AND n.EVENT_NAME NOT IN ({lista})"


def gerar(semana:int, ano:int, frota_key:str, csv_warnings:str|None=None,
          out_dir=OUT_DIR_PADRAO, semana_fim:int|None=None) -> str:
    """Gera o relatório. Para janela de uma semana, passe só `semana`.
    Para janela de múltiplas semanas (ex.: Vestas quinzenal S29-S30), passe
    `semana` (início) e `semana_fim` (fim)."""
    registrar_fontes()
    frota = FROTAS[frota_key]
    jan = janela_da_semana(semana, ano)
    jan_fim = janela_da_semana(semana_fim, ano) if semana_fim else jan
    periodo = f"{jan.inicio.strftime('%d/%m')} a {jan_fim.fim.strftime('%d/%m/%Y')}"
    if semana_fim and semana_fim != semana:
        rotulo_sem = f"S{semana:02d}-S{semana_fim:02d}"
    else:
        rotulo_sem = f"S{semana:02d}"
    log.info(f"Gerando relatório {rotulo_sem}/{ano} — {frota.cluster_oem} ({periodo})")

    # janela de temperatura: últimos 7 dias
    temp_fim = date.today() - timedelta(days=1)
    temp_ini = temp_fim - timedelta(days=6)
    periodo_temp = f"{temp_ini.strftime('%d/%m')}–{temp_fim.strftime('%d/%m')}"
    # janela de warnings: 2 semanas (semana atual + anterior)
    jan_ant = janela_da_semana(semana-1, ano)
    periodo_warn = f"S{semana-1:02d}–S{semana:02d} ({jan_ant.inicio.strftime('%d/%m')}–{jan.fim.strftime('%d/%m')})"

    # ------------------------------------------------------------------
    # EXECUÇÃO DAS QUERIES — direto no Snowflake via SSO (run_query)
    # ------------------------------------------------------------------
    # 1) nº de WTGs nominal da frota (denominador da disponibilidade)
    df_wtgs = run_query(QUERIES["contagem_wtgs"], {"cluster": frota.cluster_oem})
    n_wtgs = int(df_wtgs["N_WTGS"].iloc[0])

    # parâmetros comuns a todas as queries (janela final = jan_fim para intervalos)
    params = {
        "cluster":   frota.cluster_oem,
        "n_wtgs":    n_wtgs,
        "sem_atual": semana,
        "sem_fim":   semana_fim if semana_fim else semana,
        "sab_atual": jan.inicio.isoformat(),
        "sab_fim":   jan_fim.inicio.isoformat(),
        "data_ini":  temp_ini.isoformat(),
        "data_fim":  temp_fim.isoformat(),
    }

    excl = _filtro_exclusao(frota)

    # helper: injeta {FILTRO_EXCLUSAO} (não é %(nome)s) e roda a query
    def q(chave):
        frota = FROTAS[frota_key]
        # cláusula por EVENT_NAME (mtbf) — a que você já tinha
        excl = ""
        if frota.exclusao_alarmes:
            lista = ",".join("'" + a.replace("'", "''") + "'" for a in frota.exclusao_alarmes)
            excl = f"AND n.EVENT_NAME NOT IN ({lista})"
        # cláusula por evento_parada (taxas de falha) — a nova
        excl_evt = ""
        if frota.exclusao_alarmes:
            lista = ",".join("'" + a.replace("'", "''") + "'" for a in frota.exclusao_alarmes)
            excl_evt = f"AND v.evento_parada NOT IN ({lista})"

        sql = QUERIES[chave].replace("{FILTRO_EXCLUSAO}", excl)
        sql = sql.replace("{FILTRO_EXCLUSAO_EVT}", excl_evt)
        return run_query(sql, params)

    df_disp   = q("disponibilidade_semanal")
    df_alarm  = q("top_alarmes_falha")
    df_mtbf   = q("mtbf_wtg")
    df_falha  = q("maiores_tempos_falha")
    df_corr   = q("maiores_tempos_corretiva")
    df_temp   = q("temperatura_top5")

    # garante nomes de coluna em MAIÚSCULAS (as seções esperam assim)
    for _df in (df_disp, df_alarm, df_mtbf, df_falha, df_corr, df_temp):
        if _df is not None and not _df.empty:
            _df.columns = [str(c).upper() for c in _df.columns]

    warn = None
    if csv_warnings and Path(csv_warnings).exists():
        warn = processar_warnings(csv_warnings, jan_ant.inicio, jan.fim, top_n=10)

    # ------------------------------------------------------------------
    # MONTAGEM DO DOCUMENTO
    # ------------------------------------------------------------------
    doc = Document()
    configurar_pagina(doc)
    add_capa(doc, semana, ano, periodo, frota.titulo_capa, frota.sub_capa, rotulo_semana=rotulo_sem)
    add_quebra_para_conteudo(doc, semana, ano, frota.label_header, rotulo_semana=rotulo_sem)

    construir_pg2(doc, df_disp)
    construir_pg3(doc, df_disp, df_alarm)
    construir_pg4(doc, df_mtbf, df_falha, df_corr)
    construir_pg5(doc, df_temp, periodo_temp, frota_key=frota_key)
    if warn:
        construir_pg6(doc, warn, periodo_warn)

    # Remove quebra de página órfã no fim (última seção deixa um parágrafo vazio)
    while len(doc.paragraphs) > 0:
        ult = doc.paragraphs[-1]
        tem_break = ult._p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br')
        if not ult.text.strip() and (tem_break or not ult.runs):
            ult._p.getparent().remove(ult._p)
        else:
            break

    Path(out_dir).mkdir(parents=True, exist_ok=True)
    _rot_arq = f"S{semana:02d}-S{semana_fim:02d}" if (semana_fim and semana_fim != semana) else f"S{semana:02d}"
    nome = f"Fleet_Report_{_rot_arq}_{ano}_{frota_key}.docx"
    caminho = str(Path(out_dir)/nome)
    doc.save(caminho)
    log.info(f"Salvo: {caminho}")
    return caminho


if __name__ == "__main__":
    # semana pode ser "29" (uma semana) ou "29-30" (intervalo)
    arg_sem = sys.argv[1]
    if "-" in arg_sem:
        semana, semana_fim = (int(p) for p in arg_sem.split("-", 1))
    else:
        semana, semana_fim = int(arg_sem), None
    ano = int(sys.argv[2])
    frota_key = resolver_frota(sys.argv[3]) if len(sys.argv) > 3 else "SG"
    csv = sys.argv[4] if len(sys.argv) > 4 else None
    print(gerar(semana, ano, frota_key, csv, semana_fim=semana_fim))