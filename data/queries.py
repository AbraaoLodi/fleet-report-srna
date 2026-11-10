"""Catalogo de queries SQL parametrizadas - Fleet Reliability Report SRNA.

FONTE: OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS

MODOS DE JANELA (resolvidos em utils/periodo.py):
  - As secoes 2-5 usam uma janela [janela_ini, janela_fim] (datas explicitas),
    que pode ser uma semana SG fechada OU os ultimos 7/15 dias.
  - A disponibilidade (pg 1) recebe as 4 colunas ja montadas como linhas
    (rotulo + ini + fim) via o parametro %(colunas_sql)s.

TRATAMENTO DE EVENTO ABERTO:
  - Evento com data_fim conhecido (mesmo futuro) -> conta ate o fim real,
    limitado ao fim da coluna/janela.
  - Evento com data_fim = NULL (orfao) -> conta ate CURRENT_TIMESTAMP,
    limitado ao fim da coluna/janela.

Parametros:
  %(cluster)s, %(n_wtgs)s, %(colunas_sql)s (disponibilidade),
  %(janela_ini)s, %(janela_fim)s (secoes 2-5),
  %(data_ini)s, %(data_fim)s (temperatura),
  {FILTRO_EXCLUSAO}, {FILTRO_EXCLUSAO_EVT}
"""
from __future__ import annotations

Q_CONTAGEM_WTGS = """
SELECT COUNT(DISTINCT wtg) AS n_wtgs
FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS
WHERE cluster_oem = %(cluster)s;
"""

Q_DISPONIBILIDADE_SEMANAL = """
WITH colunas AS (
    SELECT * FROM ( VALUES
        %(colunas_sql)s
    ) AS c(ord, rotulo, ini, fim_excl)
),
base AS (
    SELECT c.ord, c.rotulo, c.ini, c.fim_excl, v.motivo_parada,
        DATEDIFF('second',
            GREATEST(v.data_inicio, c.ini),
            LEAST(COALESCE(v.data_fim, CURRENT_TIMESTAMP()), c.fim_excl)
        ) AS seg
    FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS v
    JOIN colunas c
      ON v.data_inicio < c.fim_excl
     AND COALESCE(v.data_fim, CURRENT_TIMESTAMP()) > c.ini
    WHERE v.cluster_oem = %(cluster)s AND v.validado = 'True'
      AND v.responsabilidade IN ('Fabricante', 'Nenhum dos dois')
      AND (v.sistema_parada IS NULL OR v.sistema_parada NOT IN ('ENVIRONMENT', 'GRID CONNECTION'))
      AND (v.motivo_parada  IS NULL OR v.motivo_parada  NOT IN ('AMBIENTAL', 'CURTAILMENT'))
      {FILTRO_EXCLUSAO_EVT}
)
SELECT c.ord AS ord, c.rotulo AS sem, c.ini::DATE AS ini, DATEADD(day,-1,c.fim_excl)::DATE AS fim,
    ROUND(100.0*COALESCE(SUM(CASE WHEN motivo_parada='FALHA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600),2) AS tx_falha,
    ROUND(100.0*COALESCE(SUM(CASE WHEN motivo_parada='MANUTENCAO CORRETIVA NAO PLANEJADA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600),2) AS tx_corr_np,
    ROUND(100.0*COALESCE(SUM(CASE WHEN motivo_parada='MANUTENCAO CORRETIVA PLANEJADA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600),2) AS tx_corr_p,
    ROUND(100.0*COALESCE(SUM(CASE WHEN motivo_parada='GRANDE CORRETIVA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600),2) AS tx_grande_corr,
    ROUND(100
        - 100.0*COALESCE(SUM(CASE WHEN motivo_parada='FALHA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600)
        - 100.0*COALESCE(SUM(CASE WHEN motivo_parada='MANUTENCAO CORRETIVA NAO PLANEJADA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600)
        - 100.0*COALESCE(SUM(CASE WHEN motivo_parada='MANUTENCAO CORRETIVA PLANEJADA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600)
        - 100.0*COALESCE(SUM(CASE WHEN motivo_parada='GRANDE CORRETIVA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600),2) AS disp_pct,
    ROUND((100
        - 100.0*COALESCE(SUM(CASE WHEN motivo_parada='FALHA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600)
        - 100.0*COALESCE(SUM(CASE WHEN motivo_parada='MANUTENCAO CORRETIVA NAO PLANEJADA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600)
        - 100.0*COALESCE(SUM(CASE WHEN motivo_parada='MANUTENCAO CORRETIVA PLANEJADA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600)
        - 100.0*COALESCE(SUM(CASE WHEN motivo_parada='GRANDE CORRETIVA' THEN GREATEST(seg,0) END),0)/(%(n_wtgs)s*7*24*3600)) - 98.0, 2) AS gap_pp,
    ROUND(COALESCE(SUM(GREATEST(seg,0)),0)/3600.0, 2) AS horas_parada
FROM colunas c
LEFT JOIN base e ON e.ord = c.ord
GROUP BY c.ord, c.rotulo, c.ini, c.fim_excl ORDER BY c.ord;
"""

Q_TOP_ALARMES_FALHA = """
WITH janela AS (
    SELECT %(janela_ini)s::TIMESTAMP_NTZ AS ini, %(janela_fim)s::TIMESTAMP_NTZ AS fim
)
SELECT v.evento_parada AS alarme, v.sistema_parada AS subsistema,
    COUNT(*) AS ocorrencias, COUNT(DISTINCT v.wtg) AS wtgs_afetadas,
    ROUND(COALESCE(SUM(GREATEST(DATEDIFF('second',
        GREATEST(v.data_inicio, j.ini),
        LEAST(COALESCE(v.data_fim, CURRENT_TIMESTAMP()), j.fim)
    ),0)),0)/3600.0, 2) AS horas
FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS v CROSS JOIN janela j
WHERE v.cluster_oem = %(cluster)s AND v.validado = 'True'
  AND v.data_inicio < j.fim
  AND COALESCE(v.data_fim, CURRENT_TIMESTAMP()) > j.ini
  AND v.motivo_parada = 'FALHA'
  AND v.responsabilidade IN ('Fabricante','Nenhum dos dois')
  AND v.sistema_parada NOT IN ('ENVIRONMENT','GRID CONNECTION')
  {FILTRO_EXCLUSAO_EVT}
GROUP BY v.evento_parada, v.sistema_parada ORDER BY horas DESC LIMIT 10;
"""

Q_MTBF_WTG = """
WITH d_spe_dedup AS (
    SELECT DISTINCT grupo_oem, grupo_ri, fabricante FROM OEM_PROD.DEVICES.D_ALL_SPE
),
mtbf_wtgs AS (
    SELECT DISTINCT f.wtg, f.grupo_oem FROM OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD f
    WHERE f.date = (SELECT MAX(date) FROM OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD)
      AND f.grupo_oem = %(cluster)s
),
falhas_por_periodo AS (
    SELECT n.WTG,
        COUNT(CASE WHEN n.TIME_ON >= %(janela_ini)s::TIMESTAMP_NTZ THEN 1 END) AS n_falhas_janela,
        ROUND(SUM(CASE WHEN n.TIME_ON >= %(janela_ini)s::TIMESTAMP_NTZ
            THEN DATEDIFF('minute', n.TIME_ON, LEAST(COALESCE(n.TIME_OFF,CURRENT_TIMESTAMP()),CURRENT_TIMESTAMP()))/60.0 END),1) AS h_falhas_janela
    FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_NORMALIZED n
    JOIN mtbf_wtgs m ON m.wtg = n.WTG
    WHERE n.TIME_ON >= DATEADD(day,-30,CURRENT_DATE)::TIMESTAMP_NTZ
      AND n.MOTIVO_PARADA = 'FALHA' {FILTRO_EXCLUSAO}
    GROUP BY n.WTG
)
SELECT d.grupo_ri, f.grupo_oem AS oem, d.fabricante, f.wtg, f.date AS snapshot_date,
    TRY_CAST(f.mtbf_7days AS FLOAT) AS mtbf_7d,
    COALESCE(fp.n_falhas_janela,0) AS n_falhas_7d,
    COALESCE(fp.h_falhas_janela,0) AS h_falhas_7d
FROM OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD f
LEFT JOIN d_spe_dedup d USING (grupo_oem)
LEFT JOIN falhas_por_periodo fp ON fp.WTG = f.wtg
WHERE f.date = (SELECT MAX(date) FROM OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD)
  AND f.grupo_oem = %(cluster)s AND d.grupo_ri IS NOT NULL
ORDER BY TRY_CAST(f.mtbf_7days AS FLOAT) ASC NULLS LAST LIMIT 7;
"""

Q_MAIORES_TEMPOS_FALHA = """
WITH janela AS (
    SELECT %(janela_ini)s::TIMESTAMP_NTZ AS ini, %(janela_fim)s::TIMESTAMP_NTZ AS fim
),
eventos AS (
    SELECT v.wtg, v.evento_parada, v.sistema_parada, v.data_inicio,
        v.data_fim AS data_fim,
        GREATEST(DATEDIFF('second',
            GREATEST(v.data_inicio, j.ini),
            LEAST(COALESCE(v.data_fim, CURRENT_TIMESTAMP()), j.fim)
        ),0)/3600.0 AS horas
    FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS v CROSS JOIN janela j
    WHERE v.cluster_oem = %(cluster)s AND v.validado = 'True'
      AND v.data_inicio < j.fim
      AND COALESCE(v.data_fim, CURRENT_TIMESTAMP()) > j.ini
      AND v.motivo_parada = 'FALHA'
      AND v.responsabilidade IN ('Fabricante','Nenhum dos dois')
      AND v.sistema_parada NOT IN ('ENVIRONMENT','GRID CONNECTION')
      {FILTRO_EXCLUSAO_EVT}
),
wtg_total AS (SELECT wtg, SUM(horas) AS h_total FROM eventos GROUP BY wtg),
top_wtgs AS (SELECT wtg, h_total FROM wtg_total ORDER BY h_total DESC LIMIT 10)
SELECT t.wtg AS aerogerador, e.sistema_parada AS subsistema, e.evento_parada AS alarme,
    COUNT(*) AS ocorrencias, MIN(e.data_inicio)::DATE AS data_inicio,
    MAX(e.data_fim)::DATE AS data_fim,
    BOOLOR_AGG(e.data_fim IS NULL) AS em_aberto,
    ROUND(SUM(e.horas),4) AS horas_alarme, ROUND(t.h_total,4) AS horas_total_wtg
FROM eventos e JOIN top_wtgs t ON t.wtg = e.wtg
GROUP BY t.wtg, t.h_total, e.sistema_parada, e.evento_parada
ORDER BY t.h_total DESC, horas_alarme DESC;
"""

Q_MAIORES_TEMPOS_CORRETIVA = """
WITH janela AS (
    SELECT %(janela_ini)s::TIMESTAMP_NTZ AS ini, %(janela_fim)s::TIMESTAMP_NTZ AS fim
),
eventos AS (
    SELECT v.wtg, v.evento_parada, v.sistema_parada, v.motivo_parada, v.data_inicio,
        v.data_fim AS data_fim,
        GREATEST(DATEDIFF('second',
            GREATEST(v.data_inicio, j.ini),
            LEAST(COALESCE(v.data_fim, CURRENT_TIMESTAMP()), j.fim)
        ),0)/3600.0 AS horas
    FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS v CROSS JOIN janela j
    WHERE v.cluster_oem = %(cluster)s AND v.validado = 'True'
      AND v.data_inicio < j.fim
      AND COALESCE(v.data_fim, CURRENT_TIMESTAMP()) > j.ini
      AND v.responsabilidade IN ('Fabricante','Nenhum dos dois')
      AND v.sistema_parada NOT IN ('ENVIRONMENT','GRID CONNECTION')
      AND v.motivo_parada IN ('FALHA','MANUTENCAO CORRETIVA NAO PLANEJADA',
                              'MANUTENCAO CORRETIVA PLANEJADA','GRANDE CORRETIVA',
                              'MANUTENCAO PREVENTIVA','MANUTENCAO PREDITIVA')
      {FILTRO_EXCLUSAO_EVT}
),
wtg_total AS (SELECT wtg, SUM(horas) AS h_total FROM eventos GROUP BY wtg),
top_wtgs AS (SELECT wtg, h_total FROM wtg_total ORDER BY h_total DESC LIMIT 10)
SELECT t.wtg AS aerogerador, e.sistema_parada AS subsistema, e.motivo_parada AS motivo,
    e.evento_parada AS alarme, COUNT(*) AS ocorrencias,
    MIN(e.data_inicio)::DATE AS data_inicio, MAX(e.data_fim)::DATE AS data_fim,
    BOOLOR_AGG(e.data_fim IS NULL) AS em_aberto,
    ROUND(SUM(e.horas),4) AS horas_alarme, ROUND(t.h_total,4) AS horas_total_wtg
FROM eventos e JOIN top_wtgs t ON t.wtg = e.wtg
GROUP BY t.wtg, t.h_total, e.sistema_parada, e.motivo_parada, e.evento_parada
ORDER BY t.h_total DESC, horas_alarme DESC;
"""

Q_TEMPERATURA_TOP5 = """
WITH periodo AS (SELECT %(data_ini)s::DATE AS ini, %(data_fim)s::DATE AS fim),
limites AS (
    SELECT DISTINCT VARIABLE,
        TRY_CAST(NULLIF(TRIM(DERATE_LIMIT::VARCHAR),'') AS FLOAT) AS derate,
        TRY_CAST(NULLIF(TRIM(DOWNTIME_LIMIT::VARCHAR),'') AS FLOAT) AS downtime
    FROM OEM_PROD.MACHINE_LEARNING.F_REAL_EOL_WTG_TEMPERATURE_ANOMALY_EVENTS_DAILY
    WHERE GRUPO_OEM = %(cluster)s
      AND DATE = (SELECT MAX(DATE) FROM OEM_PROD.MACHINE_LEARNING.F_REAL_EOL_WTG_TEMPERATURE_ANOMALY_EVENTS_DAILY WHERE GRUPO_OEM = %(cluster)s)
),
pico_ml AS (
    SELECT m.WTG, m.VARIABLE, ROUND(MAX(m.MAX_TEMP),1) AS pico
    FROM OEM_PROD.MACHINE_LEARNING.F_REAL_EOL_WTG_TEMPERATURE_ANOMALY_EVENTS_DAILY m, periodo p
    WHERE m.GRUPO_OEM = %(cluster)s AND m.DATE BETWEEN p.ini AND p.fim
      AND m.MAX_TEMP > 0 AND m.MAX_TEMP < 200
    GROUP BY m.WTG, m.VARIABLE
),
last_upload AS (
    SELECT r.WTG, r.TIMESTAMP_LOCAL, r.VARIABLE, MAX(r.UPLOADED_AT) AS lu
    FROM OEM_PROD.SCADA.F_REAL_EOL_WTG_SCADA_RAW r, periodo p
    WHERE r.DATE_LOCAL BETWEEN p.ini AND p.fim AND r.GRUPO_OEM = %(cluster)s AND r.VARIABLE LIKE '%%TEMP%%'
    GROUP BY ALL
),
media_scada AS (
    SELECT r.WTG, r.VARIABLE, ROUND(AVG(CASE WHEN r.VALUE>0 AND r.VALUE<200 THEN r.VALUE END),1) AS media
    FROM OEM_PROD.SCADA.F_REAL_EOL_WTG_SCADA_RAW r
    INNER JOIN last_upload lu USING (WTG, TIMESTAMP_LOCAL, VARIABLE), periodo p
    WHERE r.DATE_LOCAL BETWEEN p.ini AND p.fim AND r.GRUPO_OEM = %(cluster)s
      AND r.VARIABLE LIKE '%%TEMP%%' AND r.AGGREGATED = 'AVG' AND r.UPLOADED_AT = lu.lu
    GROUP BY r.WTG, r.VARIABLE
),
sensores AS (
    SELECT p.WTG, p.VARIABLE, p.pico, m.media, l.derate, l.downtime,
        CASE WHEN l.downtime IS NOT NULL AND p.pico >= l.downtime THEN 2
             WHEN l.derate   IS NOT NULL AND p.pico >= l.derate   THEN 1 ELSE 0 END AS sev
    FROM pico_ml p LEFT JOIN limites l ON l.VARIABLE = p.VARIABLE
    LEFT JOIN media_scada m ON m.WTG = p.WTG AND m.VARIABLE = p.VARIABLE
)
SELECT s.WTG AS wtg, s.VARIABLE AS variable, s.pico AS pico, s.media AS media,
    s.derate AS derate, s.downtime AS downtime, s.sev AS sev
FROM sensores s
WHERE s.sev >= 1
ORDER BY s.sev DESC, s.pico DESC;
"""

QUERIES = {
    "contagem_wtgs":            Q_CONTAGEM_WTGS,
    "disponibilidade_semanal":  Q_DISPONIBILIDADE_SEMANAL,
    "top_alarmes_falha":        Q_TOP_ALARMES_FALHA,
    "mtbf_wtg":                 Q_MTBF_WTG,
    "maiores_tempos_falha":     Q_MAIORES_TEMPOS_FALHA,
    "maiores_tempos_corretiva": Q_MAIORES_TEMPOS_CORRETIVA,
    "temperatura_top5":         Q_TEMPERATURA_TOP5,
}