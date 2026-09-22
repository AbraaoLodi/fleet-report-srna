"""Catálogo de queries SQL parametrizadas — Fleet Reliability Report SRNA.

FONTE: OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS

DIFERENÇAS CRÍTICAS vs. a _DAY (motivo destas correções):
  1. Eventos em aberto têm DATA_FIM = NULL → todo LEAST(data_fim, ...) precisa de
     COALESCE(data_fim, <limite>) senão o DATEDIFF vira NULL e o Python quebra
     (TypeError: float() argument ... not 'NoneType').
  2. A coluna `data` é uma linha por evento (não linha-por-dia). O filtro de janela
     deve ser por OVERLAP de timestamp (data_inicio < fim_janela AND
     COALESCE(data_fim, limite) > ini_janela), não por `data BETWEEN`.

Tratamento de evento aberto (Opção B): conta horas até o FIM DA JANELA da semana.
  → limite = DATEADD(day,1,s.fim) na disponibilidade; j.fim nas demais.
  Para contar até AGORA (Opção A), troque o limite por CURRENT_TIMESTAMP().

Parâmetros:
  %(cluster)s, %(n_wtgs)s, %(sem_fim)s, %(sab_fim)s, %(sab_atual)s,
  %(data_ini)s, %(data_fim)s, {FILTRO_EXCLUSAO}
"""
from __future__ import annotations

# nº de WTGs da frota, direto do banco
Q_CONTAGEM_WTGS = """
SELECT COUNT(DISTINCT wtg) AS n_wtgs
FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS
WHERE cluster_oem = %(cluster)s;
"""

# Pg 2 — Disponibilidade + 4 taxas (4 semanas SG)
Q_DISPONIBILIDADE_SEMANAL = """
WITH semanas AS (
    SELECT sem_num AS sem,
        DATEADD(week, sem_num - %(sem_fim)s, %(sab_fim)s::DATE) AS ini,
        DATEADD(day, 6, DATEADD(week, sem_num - %(sem_fim)s, %(sab_fim)s::DATE)) AS fim
    FROM (
        SELECT %(sem_fim)s - 3 AS sem_num UNION ALL SELECT %(sem_fim)s - 2
        UNION ALL SELECT %(sem_fim)s - 1 UNION ALL SELECT %(sem_fim)s
    )
),
base AS (
    SELECT s.sem, s.ini, s.fim, v.motivo_parada,
        DATEDIFF('second',
            GREATEST(v.data_inicio, s.ini::TIMESTAMP_NTZ),
            LEAST(COALESCE(v.data_fim, DATEADD(day,1,s.fim)::TIMESTAMP_NTZ),
                  DATEADD(day, 1, s.fim)::TIMESTAMP_NTZ)
        ) AS seg
    FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS v
    JOIN semanas s
      ON v.data_inicio < DATEADD(day,1,s.fim)::TIMESTAMP_NTZ
     AND COALESCE(v.data_fim, DATEADD(day,1,s.fim)::TIMESTAMP_NTZ) > s.ini::TIMESTAMP_NTZ
    WHERE v.cluster_oem = %(cluster)s AND v.validado = 'True'
      AND v.responsabilidade IN ('Fabricante', 'Nenhum dos dois')
      AND (v.sistema_parada IS NULL OR v.sistema_parada NOT IN ('ENVIRONMENT', 'GRID CONNECTION'))
      AND (v.motivo_parada  IS NULL OR v.motivo_parada  NOT IN ('AMBIENTAL', 'CURTAILMENT'))
      {FILTRO_EXCLUSAO_EVT}
)
SELECT s.sem AS sem, s.ini::DATE AS ini, s.fim::DATE AS fim,
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
FROM semanas s
LEFT JOIN base e ON e.sem = s.sem
GROUP BY s.sem, s.ini, s.fim ORDER BY s.sem;
"""

# Pg 3 — Top 10 alarmes por horas de falha (semana atual)
Q_TOP_ALARMES_FALHA = """
WITH janela AS (
    SELECT %(sab_atual)s::TIMESTAMP_NTZ AS ini, DATEADD(day,7,%(sab_fim)s)::TIMESTAMP_NTZ AS fim
)
SELECT v.evento_parada AS alarme, v.sistema_parada AS subsistema,
    COUNT(*) AS ocorrencias, COUNT(DISTINCT v.wtg) AS wtgs_afetadas,
    ROUND(COALESCE(SUM(GREATEST(DATEDIFF('second',
        GREATEST(v.data_inicio, j.ini),
        LEAST(COALESCE(v.data_fim, j.fim), j.fim)
    ),0)),0)/3600.0, 2) AS horas
FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS v CROSS JOIN janela j
WHERE v.cluster_oem = %(cluster)s AND v.validado = 'True'
  AND v.data_inicio < j.fim
  AND COALESCE(v.data_fim, j.fim) > j.ini
  AND v.motivo_parada = 'FALHA'
  AND v.responsabilidade IN ('Fabricante','Nenhum dos dois')
  AND v.sistema_parada NOT IN ('ENVIRONMENT','GRID CONNECTION')
  {FILTRO_EXCLUSAO_EVT}
GROUP BY v.evento_parada, v.sistema_parada ORDER BY horas DESC LIMIT 10;
"""

# Pg 4 — MTBF por WTG (inalterada: usa F_REAL_WTG_MTBF_PERIOD + NORMALIZED)
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
        COUNT(CASE WHEN n.TIME_ON >= DATEADD(day,-7,CURRENT_DATE)::TIMESTAMP_NTZ THEN 1 END) AS n_falhas_7d,
        ROUND(SUM(CASE WHEN n.TIME_ON >= DATEADD(day,-7,CURRENT_DATE)::TIMESTAMP_NTZ
            THEN DATEDIFF('minute', n.TIME_ON, LEAST(COALESCE(n.TIME_OFF,CURRENT_TIMESTAMP()),CURRENT_TIMESTAMP()))/60.0 END),1) AS h_falhas_7d
    FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_NORMALIZED n
    JOIN mtbf_wtgs m ON m.wtg = n.WTG
    WHERE n.TIME_ON >= DATEADD(day,-30,CURRENT_DATE)::TIMESTAMP_NTZ
      AND n.MOTIVO_PARADA = 'FALHA' {FILTRO_EXCLUSAO}
    GROUP BY n.WTG
)
SELECT d.grupo_ri, f.grupo_oem AS oem, d.fabricante, f.wtg, f.date AS snapshot_date,
    TRY_CAST(f.mtbf_7days AS FLOAT) AS mtbf_7d,
    COALESCE(fp.n_falhas_7d,0) AS n_falhas_7d,
    COALESCE(fp.h_falhas_7d,0) AS h_falhas_7d
FROM OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD f
LEFT JOIN d_spe_dedup d USING (grupo_oem)
LEFT JOIN falhas_por_periodo fp ON fp.WTG = f.wtg
WHERE f.date = (SELECT MAX(date) FROM OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD)
  AND f.grupo_oem = %(cluster)s AND d.grupo_ri IS NOT NULL
ORDER BY TRY_CAST(f.mtbf_7days AS FLOAT) ASC NULLS LAST LIMIT 7;
"""

# Pg 4 — Maiores tempos por FALHA (top, hierárquica WTG→Alarme)
Q_MAIORES_TEMPOS_FALHA = """
WITH janela AS (
    SELECT %(sab_atual)s::TIMESTAMP_NTZ AS ini, DATEADD(day,7,%(sab_fim)s)::TIMESTAMP_NTZ AS fim
),
eventos AS (
    SELECT v.wtg, v.evento_parada, v.sistema_parada, v.data_inicio,
        COALESCE(v.data_fim, j.fim) AS data_fim,
        GREATEST(DATEDIFF('second',
            GREATEST(v.data_inicio, j.ini),
            LEAST(COALESCE(v.data_fim, j.fim), j.fim)
        ),0)/3600.0 AS horas
    FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS v CROSS JOIN janela j
    WHERE v.cluster_oem = %(cluster)s AND v.validado = 'True'
      AND v.data_inicio < j.fim
      AND COALESCE(v.data_fim, j.fim) > j.ini
      AND v.motivo_parada = 'FALHA'
      AND v.responsabilidade IN ('Fabricante','Nenhum dos dois')
      AND v.sistema_parada NOT IN ('ENVIRONMENT','GRID CONNECTION')
      {FILTRO_EXCLUSAO_EVT}
),
wtg_total AS (SELECT wtg, SUM(horas) AS h_total FROM eventos GROUP BY wtg),
top_wtgs AS (SELECT wtg, h_total FROM wtg_total ORDER BY h_total DESC LIMIT 10)
SELECT t.wtg AS aerogerador, e.sistema_parada AS subsistema, e.evento_parada AS alarme,
    COUNT(*) AS ocorrencias, MIN(e.data_inicio)::DATE AS data_inicio, MAX(e.data_fim)::DATE AS data_fim,
    ROUND(SUM(e.horas),4) AS horas_alarme, ROUND(t.h_total,4) AS horas_total_wtg
FROM eventos e JOIN top_wtgs t ON t.wtg = e.wtg
GROUP BY t.wtg, t.h_total, e.sistema_parada, e.evento_parada
ORDER BY t.h_total DESC, horas_alarme DESC;
"""

# Pg 4 — Maiores tempos por CORRETIVA (top 10)
Q_MAIORES_TEMPOS_CORRETIVA = """
WITH janela AS (
    SELECT %(sab_atual)s::TIMESTAMP_NTZ AS ini, DATEADD(day,7,%(sab_fim)s)::TIMESTAMP_NTZ AS fim
),
eventos AS (
    SELECT v.wtg, v.evento_parada, v.sistema_parada, v.motivo_parada, v.data_inicio,
        COALESCE(v.data_fim, j.fim) AS data_fim,
        GREATEST(DATEDIFF('second',
            GREATEST(v.data_inicio, j.ini),
            LEAST(COALESCE(v.data_fim, j.fim), j.fim)
        ),0)/3600.0 AS horas
    FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS v CROSS JOIN janela j
    WHERE v.cluster_oem = %(cluster)s AND v.validado = 'True'
      AND v.data_inicio < j.fim
      AND COALESCE(v.data_fim, j.fim) > j.ini
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
    ROUND(SUM(e.horas),4) AS horas_alarme, ROUND(t.h_total,4) AS horas_total_wtg
FROM eventos e JOIN top_wtgs t ON t.wtg = e.wtg
GROUP BY t.wtg, t.h_total, e.sistema_parada, e.motivo_parada, e.evento_parada
ORDER BY t.h_total DESC, horas_alarme DESC;
"""

# Pg 5 — Temperatura Top N (inalterada)
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