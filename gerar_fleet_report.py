"""Orquestrador — Fleet Reliability Report SRNA.

Uso: python gerar_fleet_report.py <modo> <ano> <frota> [csv_warnings]

  <modo>:
    7d       -> janela movel: ultimos 7 dias (terminando ONTEM)
    15d      -> janela movel: ultimos 15 dias (terminando ONTEM)
    sNN / NN -> semana SG fixa NN (sabado-sexta), comportamento classico
    NN-MM    -> intervalo de semanas SG (ex.: quinzenal Vestas)

Ex:  python gerar_fleet_report.py 7d 2026 SG C:\\caminho\\AWS.csv
     python gerar_fleet_report.py s38 2026 SG
     python gerar_fleet_report.py 29-30 2026 VE

Nos modos moveis (7d/15d), a pagina 1 mostra as 3 ultimas semanas SG fechadas
+ a janela movel; as secoes 2-5 usam a janela movel. No modo semana, tudo
segue a semana SG informada (comportamento anterior).
"""
from __future__ import annotations
import sys
from datetime import date, timedelta
from pathlib import Path
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parent))
from config.settings import FROTAS, resolver_frota, META_DISP
from utils.periodo import resolver_periodo
from utils.fonts import registrar_fontes
from utils.docx_helpers import configurar_pagina, add_capa, add_quebra_para_conteudo
from utils.logger import get_logger
from data.connection import run_query
from data.queries import QUERIES
from data.warnings_parser import processar_warnings
from secoes.pg2_disponibilidade import construir_pg2
from secoes.pg3_alarmes import construir_pg3
from secoes.pg4_mtbf import construir_pg4
from secoes.pg5_temperatura import construir_pg5
from secoes.pg6_warnings import construir_pg6

log = get_logger("fleet.orquestrador")

OUT_DIR_PADRAO = r"C:\Users\abraão.lodi\OneDrive - Serena\Área de Trabalho\Downloads\dev\fleet_report_srna_v3\output"


def _colunas_sql(periodo) -> str:
    """Monta as linhas de VALUES para a query de disponibilidade a partir das
    4 colunas do periodo (3 SG + movel, ou 4 SG). fim_excl = dia seguinte ao
    ultimo dia contado (00:00), pois o clip usa fim exclusivo."""
    linhas = []
    for i, col in enumerate(periodo.historico, start=1):
        rot = col.rotulo.replace("'", "''")
        fim_excl = (col.fim + timedelta(days=1)).isoformat()
        linhas.append(
            f"({i}, '{rot}', '{col.ini.isoformat()}'::TIMESTAMP_NTZ, "
            f"'{fim_excl}'::TIMESTAMP_NTZ)")
    return ",\n        ".join(linhas)


def gerar(modo: str, ano: int, frota_key: str, csv_warnings: str | None = None,
          out_dir=OUT_DIR_PADRAO) -> str:
    registrar_fontes()
    frota = FROTAS[frota_key]
    per = resolver_periodo(modo, ano)

    periodo_txt = f"{per.foco_ini:%d/%m} a {per.foco_fim:%d/%m/%Y}"
    rotulo = per.rotulo_foco
    log.info(f"Gerando relatório [{rotulo}] {ano} — {frota.cluster_oem} ({periodo_txt})")

    # janela das secoes 2-5 (foco). fim exclusivo = dia seguinte 00:00.
    janela_ini = per.foco_ini.isoformat()
    janela_fim = (per.foco_fim + timedelta(days=1)).isoformat()

    # temperatura: sempre ultimos 7 dias terminando ontem (janela propria)
    temp_fim = date.today() - timedelta(days=1)
    temp_ini = temp_fim - timedelta(days=6)
    periodo_temp = f"{temp_ini:%d/%m}–{temp_fim:%d/%m}"

    # warnings: 2 semanas ate o fim do foco
    warn_ini = per.foco_fim - timedelta(days=13)
    periodo_warn = f"{warn_ini:%d/%m}–{per.foco_fim:%d/%m}"

    # nº de WTGs (denominador)
    df_wtgs = run_query(QUERIES["contagem_wtgs"], {"cluster": frota.cluster_oem})
    n_wtgs = int(df_wtgs["N_WTGS"].iloc[0])

    params = {
        "cluster":     frota.cluster_oem,
        "n_wtgs":      n_wtgs,
        "colunas_sql": _colunas_sql(per),   # <<< novo: 4 colunas da pg 1
        "janela_ini":  janela_ini,          # <<< novo: foco das secoes 2-5
        "janela_fim":  janela_fim,
        "data_ini":    temp_ini.isoformat(),
        "data_fim":    temp_fim.isoformat(),
    }

    def q(chave):
        excl = ""
        excl_evt = ""
        if frota.exclusao_alarmes:
            lista = ",".join("'" + a.replace("'", "''") + "'" for a in frota.exclusao_alarmes)
            excl = f"AND n.EVENT_NAME NOT IN ({lista})"
            excl_evt = f"AND v.evento_parada NOT IN ({lista})"
        sql = QUERIES[chave]
        # colunas_sql precisa ir SEM aspas (é fragmento SQL), então substitui à parte
        if "%(colunas_sql)s" in sql:
            sql = sql.replace("%(colunas_sql)s", params["colunas_sql"])
        sql = sql.replace("{FILTRO_EXCLUSAO}", excl)
        sql = sql.replace("{FILTRO_EXCLUSAO_EVT}", excl_evt)
        return run_query(sql, params)

    df_disp   = q("disponibilidade_semanal")
    df_alarm  = q("top_alarmes_falha")
    df_mtbf   = q("mtbf_wtg")
    df_falha  = q("maiores_tempos_falha")
    df_corr   = q("maiores_tempos_corretiva")
    df_temp   = q("temperatura_top5")

    for _df in (df_disp, df_alarm, df_mtbf, df_falha, df_corr, df_temp):
        if _df is not None and not _df.empty:
            _df.columns = [str(c).upper() for c in _df.columns]

    warn = None
    if csv_warnings and Path(csv_warnings).exists():
        warn = processar_warnings(csv_warnings, warn_ini, per.foco_fim, top_n=10)

    doc = Document()
    configurar_pagina(doc)
    add_capa(doc, rotulo, ano, periodo_txt, frota.titulo_capa, frota.sub_capa, rotulo_semana=rotulo)
    add_quebra_para_conteudo(doc, rotulo, ano, frota.label_header, rotulo_semana=rotulo)

    construir_pg2(doc, df_disp)
    construir_pg3(doc, df_disp, df_alarm)
    construir_pg4(doc, df_mtbf, df_falha, df_corr)
    construir_pg5(doc, df_temp, periodo_temp, frota_key=frota_key)
    if warn:
        construir_pg6(doc, warn, periodo_warn)

    while len(doc.paragraphs) > 0:
        ult = doc.paragraphs[-1]
        tem_break = ult._p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br')
        if not ult.text.strip() and (tem_break or not ult.runs):
            ult._p.getparent().remove(ult._p)
        else:
            break

    Path(out_dir).mkdir(parents=True, exist_ok=True)
    nome = f"Fleet_Report_{per.rotulo_arquivo}_{ano}_{frota_key}.docx"
    caminho = str(Path(out_dir) / nome)
    doc.save(caminho)
    log.info(f"Salvo: {caminho}")
    return caminho


if __name__ == "__main__":
    modo = sys.argv[1]                       # "7d" | "15d" | "s38" | "38" | "29-30"
    ano = int(sys.argv[2])
    frota_key = resolver_frota(sys.argv[3]) if len(sys.argv) > 3 else "SG"
    csv = sys.argv[4] if len(sys.argv) > 4 else None
    print(gerar(modo, ano, frota_key, csv))