"""Fleet Reliability Report SRNA — Complexo Eólico Assuruá SG | G97

Uso:
    python gerar_fleet_report.py --semana 21
    python gerar_fleet_report.py --semana 21 --ano 2026
    python gerar_fleet_report.py --semana 25 --mock     # roda com dados fake

Estrutura de páginas geradas:
    Capa   Fullpage Serena (grafite + linha coral + título branco)
    Pg 2   01. Disponibilidade Fabricante
    Pg 3   02. Taxa de Falhas e Corretivas
    Pg 4   03. MTBF e WTGs com Maiores Tempos em Falha
    Pg 5   04. Análise de Temperatura
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from docx import Document

from config.settings import APP_CONFIG
from data.queries import QUERIES
from secoes.pg2_disponibilidade import construir_pg2
from secoes.pg3_taxas import construir_pg3
from secoes.pg4_mtbf import construir_pg4
from secoes.pg5_temperatura import construir_pg5
from utils.docx_helpers import (
    add_capa, add_quebra_para_conteudo, configurar_pagina,
)
from utils.logger import get_logger
from utils.semana_sg import janela_da_semana

log = get_logger("gerar_fleet_report")


# =============================================================================
# Janela de temperatura (últimos 7 dias, até ontem)
# =============================================================================

def _janela_temperatura() -> tuple[str, str, str]:
    """Devolve (data_ini_iso, data_fim_iso, rotulo) da janela de 7 dias
    que termina ONTEM (dado do dia corrente ainda é parcial)."""
    hoje = date.today()
    fim = hoje - timedelta(days=1)          # até ontem
    ini = fim - timedelta(days=6)           # 7 dias corridos
    rotulo = (f"{ini.strftime('%d/%m')} a {fim.strftime('%d/%m/%Y')} · "
              "últimos 7 dias")
    return ini.isoformat(), fim.isoformat(), rotulo


# =============================================================================
# Mocks — para testar layout sem Snowflake
# =============================================================================

def _mock_dados(semana: int, ano: int) -> dict:
    """Devolve DataFrames mockados (mesma estrutura das queries reais)."""
    semanas = [semana - 3, semana - 2, semana - 1, semana]

    df_disp = pd.DataFrame({
        "SEM":          semanas,
        "INI":          [janela_da_semana(s, ano).inicio for s in semanas],
        "FIM":          [janela_da_semana(s, ano).fim for s in semanas],
        "TX_FALHA":       [1.00, 0.87, 0.88, 0.76],
        "TX_CORR_NP":     [4.91, 7.45, 7.78, 5.81],
        "TX_CORR_P":      [2.10, 0.35, 0.49, 0.00],
        "TX_GRANDE_CORR": [1.65, 2.04, 2.04, 2.04],
        "DISP_PCT":       [90.34, 89.29, 88.81, 91.39],
        "GAP_PP":         [-7.66, -8.71, -9.19, -6.61],
        "HORAS_PARADA":   [795.4, 880.0, 920.0, 707.0],
    })

    df_alarmes = pd.DataFrame({
        "ALARME":        ["205_LOW LEVEL HYDRAULIC GROUP REFRIGERATION OIL",
                          "2118_BUS ERROR - INTERBUS FAILURE",
                          "5253_EXCESS TEMPERATURE IN THE RADIATOR'S HARDWARE",
                          "203_LOW PRESSURE ON HYDRAULIC GROUP",
                          "500_HIGH GENERATOR WINDING TEMPERATURE",
                          "813_LOW PRESSURE OF EMERGENCY ACCUMULATORS",
                          "5201_MAGNETOTHERMIC TRIGGER OF THE STATOR PROTECTION"],
        "SUBSISTEMA":    ["HYDRAULICS", "CONTROL & COMMUNICATION SYSTEM",
                          "FREQUENCY CONVERTER", "HYDRAULICS", "GENERATOR",
                          "PITCH SYSTEM", "FREQUENCY CONVERTER"],
        "OCORRENCIAS":   [3, 4, 5, 2, 7, 1, 1],
        "WTGS_AFETADAS": [3, 3, 1, 2, 1, 1, 1],
        "HORAS":         [19.61, 9.81, 8.09, 7.29, 6.47, 5.19, 4.65],
    })

    df_mtbf_wtg = pd.DataFrame({
        "WTG":           ["A-IV-15", "A-V-09", "A-VII-05", "A-II-02",
                          "A-V-07", "A-VII-07", "A-IV-02"],
        "SNAPSHOT_DATE": ["2026-06-28"] * 7,
        "MTBF_7D":       [5.7, 9.6, 12.8, 20.2, 23.5, 39.0, 53.2],
        "MTBF_15D":      [360.0] * 7,
        "MTBF_30D":      [720.0] * 7,
        "N_FALHAS_7D":   [27, 19, 13, 9, 7, 5, 2],
        "N_FALHAS_15D":  [30, 22, 15, 10, 8, 6, 3],
        "N_FALHAS_30D":  [35, 25, 18, 12, 9, 7, 4],
        "H_FALHAS_7D":   [13.1, 5.7, 1.5, 6.6, 3.3, 19.8, 0.4],
        "GRUPO_RI":      ["Assurua"] * 7,
        "OEM":           ["Assurua SG"] * 7,
        "FABRICANTE":    ["GAMESA"] * 7,
    })

    df_wtgs = pd.DataFrame({
        "WTG":              ["A-V-03", "A-V-10"],
        "SUBSISTEMA":       ["HYDRAULICS", "PITCH SYSTEM"],
        "N_EVENTOS":        [1, 2],
        "DOWNTIME_H":       [11.32, 9.84],
        "MTTR_H":           [11.32, 4.92],
        "ALARME_PRINCIPAL": ["205_LOW LEVEL HYDRAULIC", "5201_MAGNETOTHERMIC"],
        "STATUS":           ["Atenção", "Monitorar"],
    })

    df_maiores_falha = pd.DataFrame({
        "AEROGERADOR":     ["A-V-03", "A-V-10", "A-V-10", "A-IV-15", "A-V-06"],
        "SUBSISTEMA":      ["HYDRAULICS", "PITCH SYSTEM", "PITCH SYSTEM",
                            "FREQUENCY CONVERTER", "CONTROL & COMMUNICATION SYSTEM"],
        "ALARME":          ["205_LOW LEVEL HYDRAULIC GROUP REFRIGERATION OIL",
                            "813_LOW PRESSURE OF EMERGENCY ACCUMULATORS",
                            "5201_MAGNETOTHERMIC TRIGGER OF THE STATOR PROTECTION",
                            "5253_EXCESS TEMPERATURE IN THE RADIATOR'S HARDWARE",
                            "2118_BUS ERROR - INTERBUS FAILURE"],
        "OCORRENCIAS":     [1, 1, 1, 5, 1],
        "DATA_INICIO":     ["2026-06-27", "2026-06-28", "2026-06-27",
                            "2026-06-27", "2026-06-28"],
        "DATA_FIM":        ["2026-06-27", "2026-06-28", "2026-06-27",
                            "2026-06-27", "2026-06-28"],
        "HORAS_ALARME":    [11.3242, 5.1936, 4.6422, 8.0917, 7.9278],
        "HORAS_TOTAL_WTG": [11.3242, 9.8358, 9.8358, 8.0917, 7.9278],
    })

    df_maiores_corretiva = pd.DataFrame({
        "AEROGERADOR":     ["A-II-15", "A-IV-09", "A-IV-12", "A-IV-13"],
        "SUBSISTEMA":      ["GEARBOX", "BLADE", "BLADE", "BLADE"],
        "MOTIVO":          ["GRANDE CORRETIVA", "MANUTENCAO CORRETIVA NAO PLANEJADA",
                            "MANUTENCAO CORRETIVA NAO PLANEJADA",
                            "MANUTENCAO CORRETIVA NAO PLANEJADA"],
        "ALARME":          ["1004_WINDTURBINE PAUSE"] * 4,
        "OCORRENCIAS":     [7, 7, 7, 6],
        "DATA_INICIO":     ["2026-06-27"] * 4,
        "DATA_FIM":        ["2026-07-03"] * 4,
        "HORAS_ALARME":    [167.9981, 167.9981, 167.9981, 134.05],
        "HORAS_TOTAL_WTG": [167.9981, 167.9981, 167.9981, 142.59],
    })

    # Temperatura — mock com os dados reais validados (Top 5 WTGs S27)
    df_temp = pd.DataFrame({
        "WTG":        ["A-V-07", "A-V-07", "A-II-02", "A-IV-03", "A-IV-03",
                       "A-IV-03", "A-V-03", "A-VII-09"],
        "COMPONENTE": ["Gerador enrol.", "Gearbox", "Gerador enrol.",
                       "Gerador enrol.", "Gerador rolam.", "Gearbox",
                       "Gerador enrol.", "Gerador enrol."],
        "PICO":       [149.3, 85.2, 153.1, 144.6, 92.5, 85.6, 145.0, 147.5],
        "MEDIA":      [95.6, 77.9, 106.8, 102.9, 74.2, 75.8, 102.1, 98.6],
        "DERATE":     [130, 85, 130, 130, 90, 85, 130, 130],
        "DOWNTIME":   [135, 90, 135, 135, 100, 90, 135, 135],
        "SEV":        [2, 2, 2, 2, 1, 1, 2, 2],
        "N_CRIT":     [4, 4, 3, 3, 3, 3, 3, 3],
        "N_ALERTA":   [1, 1, 0, 3, 3, 3, 0, 0],
    })

    return {
        "df_disp":              df_disp,
        "df_alarmes":           df_alarmes,
        "df_mtbf_wtg":          df_mtbf_wtg,
        "df_wtgs":              df_wtgs,
        "df_maiores_falha":     df_maiores_falha,
        "df_maiores_corretiva": df_maiores_corretiva,
        "df_temp":              df_temp,
    }


def _consultar_snowflake(semana: int, ano: int) -> dict:
    """Executa todas as queries reais contra Snowflake."""
    from data.connection import run_query, close_connection, SnowflakeUnavailable

    j = janela_da_semana(semana, ano)
    params_sem = {"sem_atual": semana,
                  "sab_atual": j.inicio.isoformat()}

    data_ini, data_fim, _ = _janela_temperatura()
    params_temp = {"data_ini": data_ini, "data_fim": data_fim}

    try:
        log.info("Pg 2 — Disponibilidade + 4 taxas (4 semanas)...")
        df_disp = run_query(QUERIES["disponibilidade_semanal"], params_sem)
        log.info("Pg 3 — Top 10 alarmes por falha...")
        df_alarmes = run_query(QUERIES["top_alarmes_falha"], params_sem)
        log.info("Pg 4 — WTGs Críticas...")
        df_wtgs = run_query(QUERIES["wtgs_criticas"], params_sem)
        log.info("Pg 4 — MTBF por WTG (snapshot oficial)...")
        df_mtbf_wtg = run_query(QUERIES["mtbf_wtg"], {})
        log.info("Pg 4 — Maiores Tempos por Falha (top 20)...")
        df_maiores_falha = run_query(QUERIES["maiores_tempos_falha"], params_sem)
        log.info("Pg 4 — Maiores Tempos por Corretiva (top 10)...")
        df_maiores_corretiva = run_query(QUERIES["maiores_tempos_corretiva"], params_sem)
        log.info("Pg 5 — Temperatura Top 5 (7 dias)...")
        df_temp = run_query(QUERIES["temperatura_top5"], params_temp)
    except SnowflakeUnavailable as exc:
        log.error("Falha Snowflake: %s", exc)
        close_connection()
        raise
    finally:
        close_connection()

    return {
        "df_disp":              df_disp,
        "df_alarmes":           df_alarmes,
        "df_wtgs":              df_wtgs,
        "df_mtbf_wtg":          df_mtbf_wtg,
        "df_maiores_falha":     df_maiores_falha,
        "df_maiores_corretiva": df_maiores_corretiva,
        "df_temp":              df_temp,
    }


def main(semana: int, ano: int, mock: bool = False) -> int:
    j = janela_da_semana(semana, ano)
    log.info("=" * 60)
    log.info("Fleet Report SRNA · S%02d/%d (%s a %s)",
             semana, ano, j.inicio.strftime("%d/%m"), j.fim.strftime("%d/%m"))
    log.info("=" * 60)

    if mock:
        log.info("Modo MOCK — usando dados de exemplo.")
        dados = _mock_dados(semana, ano)
    else:
        dados = _consultar_snowflake(semana, ano)

    # ----- Montar o documento -----
    doc = Document()

    # Capa
    add_capa(doc, semana=semana, ano=ano,
             periodo_str=f"{j.inicio.strftime('%d')}–{j.fim.strftime('%d %b %Y')}".lower())

    # Quebra de seção pra começar as páginas de conteúdo (com header/rodapé)
    add_quebra_para_conteudo(doc, semana=semana, ano=ano)

    # Pg 2
    log.info("Construindo Pg 2 — Disponibilidade...")
    construir_pg2(doc, dados["df_disp"])

    # Pg 3
    log.info("Construindo Pg 3 — Taxa de Falhas e Corretivas...")
    construir_pg3(doc, dados["df_disp"], dados["df_alarmes"], sem_atual=semana)

    # Pg 4
    log.info("Construindo Pg 4 — MTBF e WTGs com Maiores Tempos em Falha...")
    sem_ini = semana - 3
    construir_pg4(doc, dados["df_mtbf_wtg"], dados["df_wtgs"],
                  dados["df_maiores_falha"], dados["df_maiores_corretiva"],
                  semana_ini=sem_ini, semana_fim=semana)

    # Pg 5 — Temperatura
    log.info("Construindo Pg 5 — Análise de Temperatura...")
    _, _, periodo_temp = _janela_temperatura()
    construir_pg5(doc, dados["df_temp"], periodo_str=periodo_temp)

    # ----- Salvar -----
    APP_CONFIG.output_dir.mkdir(parents=True, exist_ok=True)
    out_path = APP_CONFIG.output_dir / f"Fleet_Report_SRNA_S{semana:02d}_{ano}.docx"
    doc.save(out_path)
    log.info("✓ Relatório salvo: %s", out_path)
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--semana", type=int, required=True)
    ap.add_argument("--ano", type=int, default=date.today().year)
    ap.add_argument("--mock", action="store_true",
                    help="Roda com dados mock (sem Snowflake)")
    args = ap.parse_args()
    sys.exit(main(args.semana, args.ano, mock=args.mock))