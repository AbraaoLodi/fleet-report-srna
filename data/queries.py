"""Catálogo de queries SQL — Fleet Reliability Report SRNA.

Validado contra cockpit Serena em 23/06/2026:
  - Pg 2 (Disp + 4 Taxas)          → bate à vírgula em S23, S24, S25
  - Pg 3 (Top Alarmes Falha)       → bate (410_OGS 12:45 / 1836_MV SWITCH TRIP 08:55)
  - Pg 4 (WTGs Críticas)           → validado em S25 fechada
  - Pg 4 (MTBF por WTG)            → bate à vírgula com cockpit (snapshot 21/06)
  - Pg 4/5 (Maiores Tempos)        → validado em S25 fechada
  - Pg 5 (Temperatura Top 5)       → validado contra site oficial e planilha
                                      Temperature Analysis by WTG (A-II-02 / A-II-03)

Filtros base (governança Serena, alinhado à engenheira O&M):
    CLUSTER_OEM = 'Assurua SG'
    VALIDADO = 'True'
    RESPONSABILIDADE IN ('Fabricante', 'Nenhum dos dois')
    SISTEMA_PARADA NOT IN ('ENVIRONMENT', 'GRID CONNECTION')
    MOTIVO_PARADA NOT IN ('AMBIENTAL', 'CURTAILMENT')
"""
from __future__ import annotations


# =============================================================================
# Pg 2 — DISPONIBILIDADE FABRICANTE + 4 TAXAS POR SEMANA (4 semanas SG)
# Fórmula da engenheira: Disp Possível = 100 - (Falha + CorrNP + CorrP + GC)
# =============================================================================
Q_DISPONIBILIDADE_SEMANAL = """
WITH semanas AS (
    SELECT
        sem_num AS sem,
        DATEADD(week, sem_num - %(sem_atual)s, %(sab_atual)s::DATE) AS ini,
        DATEADD(day, 6,
            DATEADD(week, sem_num - %(sem_atual)s, %(sab_atual)s::DATE)) AS fim
    FROM (
        SELECT %(sem_atual)s - 3 AS sem_num UNION ALL
        SELECT %(sem_atual)s - 2 UNION ALL
        SELECT %(sem_atual)s - 1 UNION ALL
        SELECT %(sem_atual)s
    )
),
base AS (
    SELECT
        s.sem, s.ini, s.fim,
        v.motivo_parada,
        DATEDIFF('second',
            GREATEST(v.data_inicio, s.ini::TIMESTAMP_NTZ),
            LEAST(v.data_fim, DATEADD(day, 1, s.fim)::TIMESTAMP_NTZ)
        ) AS seg
    FROM OEM_PROD.AVAILABILITY.V_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS_DAY v
    JOIN semanas s ON v.data BETWEEN s.ini AND s.fim
    WHERE v.cluster_oem      = 'Assurua SG'
      AND v.validado         = 'True'
      AND v.responsabilidade IN ('Fabricante', 'Nenhum dos dois')
      AND (v.sistema_parada IS NULL OR v.sistema_parada NOT IN ('ENVIRONMENT', 'GRID CONNECTION'))
      AND (v.motivo_parada  IS NULL OR v.motivo_parada  NOT IN ('AMBIENTAL', 'CURTAILMENT'))
)
SELECT
    s.sem                              AS sem,
    s.ini::DATE                        AS ini,
    s.fim::DATE                        AS fim,
    ROUND(100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'FALHA' THEN seg END), 0)
                  / (49 * 7 * 24 * 3600), 2)                                                   AS tx_falha,
    ROUND(100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'MANUTENCAO CORRETIVA NAO PLANEJADA' THEN seg END), 0)
                  / (49 * 7 * 24 * 3600), 2)                                                   AS tx_corr_np,
    ROUND(100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'MANUTENCAO CORRETIVA PLANEJADA' THEN seg END), 0)
                  / (49 * 7 * 24 * 3600), 2)                                                   AS tx_corr_p,
    ROUND(100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'GRANDE CORRETIVA' THEN seg END), 0)
                  / (49 * 7 * 24 * 3600), 2)                                                   AS tx_grande_corr,
    ROUND(100
        - 100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'FALHA' THEN seg END), 0)
              / (49 * 7 * 24 * 3600)
        - 100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'MANUTENCAO CORRETIVA NAO PLANEJADA' THEN seg END), 0)
              / (49 * 7 * 24 * 3600)
        - 100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'MANUTENCAO CORRETIVA PLANEJADA' THEN seg END), 0)
              / (49 * 7 * 24 * 3600)
        - 100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'GRANDE CORRETIVA' THEN seg END), 0)
              / (49 * 7 * 24 * 3600), 2)                                                       AS disp_pct,
    ROUND((100
        - 100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'FALHA' THEN seg END), 0)
              / (49 * 7 * 24 * 3600)
        - 100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'MANUTENCAO CORRETIVA NAO PLANEJADA' THEN seg END), 0)
              / (49 * 7 * 24 * 3600)
        - 100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'MANUTENCAO CORRETIVA PLANEJADA' THEN seg END), 0)
              / (49 * 7 * 24 * 3600)
        - 100.0 * COALESCE(SUM(CASE WHEN motivo_parada = 'GRANDE CORRETIVA' THEN seg END), 0)
              / (49 * 7 * 24 * 3600)) - 98.0, 2)                                               AS gap_pp,
    ROUND(COALESCE(SUM(seg), 0) / 3600.0, 2)                                                   AS horas_parada
FROM semanas s
LEFT JOIN base e ON e.sem = s.sem
GROUP BY s.sem, s.ini, s.fim
ORDER BY s.sem;
"""


# =============================================================================
# Pg 3 — TOP 10 ALARMES POR HORAS DE FALHA (semana atual SG, mesmo parcial)
# Equivale à "Maiores Tempos de Parada por Falha" da Luisa.
# Validado: 410_OGS 12:45 / 1836_MV SWITCH TRIP 08:55 em S25 fechada.
# =============================================================================
Q_TOP_ALARMES_FALHA = """
WITH janela AS (
    SELECT %(sab_atual)s::TIMESTAMP_NTZ                  AS ini,
           DATEADD(day, 7, %(sab_atual)s)::TIMESTAMP_NTZ AS fim
)
SELECT
    v.evento_parada                                                        AS alarme,
    v.sistema_parada                                                       AS subsistema,
    COUNT(*)                                                               AS ocorrencias,
    COUNT(DISTINCT v.wtg)                                                  AS wtgs_afetadas,
    ROUND(
        SUM(DATEDIFF('second',
                     GREATEST(v.data_inicio, j.ini),
                     LEAST(v.data_fim, j.fim)
            )) / 3600.0
    , 2)                                                                   AS horas
FROM OEM_PROD.AVAILABILITY.V_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS_DAY v
CROSS JOIN janela j
WHERE v.cluster_oem      = 'Assurua SG'
  AND v.validado         = 'True'
  AND v.data             BETWEEN j.ini::DATE AND DATEADD(day, -1, j.fim)::DATE
  AND v.motivo_parada    = 'FALHA'
  AND v.responsabilidade IN ('Fabricante', 'Nenhum dos dois')
  AND v.sistema_parada NOT IN ('ENVIRONMENT', 'GRID CONNECTION')
GROUP BY v.evento_parada, v.sistema_parada
ORDER BY horas DESC
LIMIT 10;
"""


# =============================================================================
# Pg 4 — WTGs CRÍTICAS POR FALHA (semana atual SG)
# Status: Crítico (≥20h) / Atenção (10-20h) / Monitorar (<10h)
# =============================================================================
Q_WTGS_CRITICAS = """
WITH janela AS (
    SELECT %(sab_atual)s::TIMESTAMP_NTZ                  AS ini,
           DATEADD(day, 7, %(sab_atual)s)::TIMESTAMP_NTZ AS fim
),
eventos AS (
    SELECT v.wtg, v.sistema_parada, v.evento_parada,
        DATEDIFF('second',
            GREATEST(v.data_inicio, j.ini),
            LEAST(v.data_fim, j.fim)
        ) / 3600.0 AS horas
    FROM OEM_PROD.AVAILABILITY.V_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS_DAY v
    CROSS JOIN janela j
    WHERE v.cluster_oem      = 'Assurua SG'
      AND v.validado         = 'True'
      AND v.data             BETWEEN j.ini::DATE AND DATEADD(day, -1, j.fim)::DATE
      AND v.motivo_parada    = 'FALHA'
      AND v.responsabilidade IN ('Fabricante', 'Nenhum dos dois')
      AND v.sistema_parada NOT IN ('ENVIRONMENT', 'GRID CONNECTION')
),
wtg_agg AS (
    SELECT wtg, COUNT(*) AS eventos, SUM(horas) AS downtime, AVG(horas) AS mttr
    FROM eventos GROUP BY wtg
),
alarme_top AS (
    SELECT wtg, evento_parada AS alarme_principal FROM (
        SELECT wtg, evento_parada,
               ROW_NUMBER() OVER (PARTITION BY wtg ORDER BY COUNT(*) DESC, MIN(evento_parada)) AS rn
        FROM eventos GROUP BY wtg, evento_parada
    ) WHERE rn = 1
),
subsistema_top AS (
    SELECT wtg, sistema_parada AS subsistema_principal FROM (
        SELECT wtg, sistema_parada,
               ROW_NUMBER() OVER (PARTITION BY wtg ORDER BY SUM(horas) DESC, MIN(sistema_parada)) AS rn
        FROM eventos GROUP BY wtg, sistema_parada
    ) WHERE rn = 1
)
SELECT
    w.wtg                          AS wtg,
    s.subsistema_principal         AS subsistema,
    w.eventos                      AS n_eventos,
    ROUND(w.downtime, 2)           AS downtime_h,
    ROUND(w.mttr, 2)               AS mttr_h,
    a.alarme_principal             AS alarme_principal,
    CASE
        WHEN w.downtime >= 20 THEN 'Crítico'
        WHEN w.downtime >= 10 THEN 'Atenção'
        ELSE 'Monitorar'
    END                            AS status
FROM wtg_agg w
LEFT JOIN alarme_top      a ON a.wtg = w.wtg
LEFT JOIN subsistema_top  s ON s.wtg = w.wtg
ORDER BY w.downtime DESC
LIMIT 10;
"""


# =============================================================================
# Pg 4 — MTBF POR WTG (top 7, snapshot mais recente)
# Fonte oficial: OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD (mantido pelo time de dados)
# Validado: bate à vírgula com cockpit Serena.
# =============================================================================
Q_MTBF_WTG = """
WITH d_spe_dedup AS (
    SELECT DISTINCT grupo_oem, grupo_ri, fabricante
    FROM OEM_PROD.DEVICES.D_ALL_SPE
),
mtbf_wtgs AS (
    SELECT DISTINCT f.wtg, f.grupo_oem
    FROM OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD f
    WHERE f.date = (SELECT MAX(date) FROM OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD)
      AND f.grupo_oem = 'Assurua SG'
),
falhas_por_periodo AS (
    SELECT
        n.WTG,
        COUNT(CASE WHEN n.TIME_ON >= DATEADD(day, -7,  CURRENT_DATE)::TIMESTAMP_NTZ THEN 1 END) AS n_falhas_7d,
        COUNT(CASE WHEN n.TIME_ON >= DATEADD(day, -15, CURRENT_DATE)::TIMESTAMP_NTZ THEN 1 END) AS n_falhas_15d,
        COUNT(CASE WHEN n.TIME_ON >= DATEADD(day, -30, CURRENT_DATE)::TIMESTAMP_NTZ THEN 1 END) AS n_falhas_30d,
        ROUND(SUM(CASE WHEN n.TIME_ON >= DATEADD(day, -7, CURRENT_DATE)::TIMESTAMP_NTZ
            THEN DATEDIFF('minute', n.TIME_ON,
                     LEAST(COALESCE(n.TIME_OFF, CURRENT_TIMESTAMP()), CURRENT_TIMESTAMP())
                 ) / 60.0 END), 1) AS h_falhas_7d
    FROM OEM_PROD.AVAILABILITY.F_REAL_EOL_WTG_EVENTLOG_NORMALIZED n
    JOIN mtbf_wtgs m ON m.wtg = n.WTG
    WHERE n.TIME_ON >= DATEADD(day, -30, CURRENT_DATE)::TIMESTAMP_NTZ
      AND n.GRUPO_RI != 'Good Night'
      AND n.MOTIVO_PARADA = 'FALHA'
    GROUP BY n.WTG
)
SELECT
    d.grupo_ri,
    f.grupo_oem                          AS oem,
    d.fabricante,
    f.wtg,
    f.date                               AS snapshot_date,
    TRY_CAST(f.mtbf_7days  AS FLOAT)     AS mtbf_7d,
    TRY_CAST(f.mtbf_15days AS FLOAT)     AS mtbf_15d,
    TRY_CAST(f.mtbf_30days AS FLOAT)     AS mtbf_30d,
    COALESCE(fp.n_falhas_7d,  0)         AS n_falhas_7d,
    COALESCE(fp.n_falhas_15d, 0)         AS n_falhas_15d,
    COALESCE(fp.n_falhas_30d, 0)         AS n_falhas_30d,
    COALESCE(fp.h_falhas_7d,  0)         AS h_falhas_7d
FROM OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD f
LEFT JOIN d_spe_dedup d USING (grupo_oem)
LEFT JOIN falhas_por_periodo fp ON fp.WTG = f.wtg
WHERE f.date = (SELECT MAX(date) FROM OEM_PROD.SDGM.F_REAL_WTG_MTBF_PERIOD)
  AND f.grupo_oem = 'Assurua SG'
  AND d.grupo_ri IS NOT NULL
ORDER BY TRY_CAST(f.mtbf_7days AS FLOAT) ASC NULLS LAST
LIMIT 7;
"""


# =============================================================================
# Pg 4/5 — MAIORES TEMPOS DE PARADA POR FALHA (top 20, hierárquica WTG→Alarme)
# =============================================================================
Q_MAIORES_TEMPOS_FALHA = """
WITH janela AS (
    SELECT %(sab_atual)s::TIMESTAMP_NTZ                  AS ini,
           DATEADD(day, 7, %(sab_atual)s)::TIMESTAMP_NTZ AS fim
),
eventos AS (
    SELECT
        v.wtg, v.evento_parada, v.sistema_parada, v.data_inicio, v.data_fim,
        DATEDIFF('second',
            GREATEST(v.data_inicio, j.ini),
            LEAST(v.data_fim, j.fim)
        ) / 3600.0 AS horas
    FROM OEM_PROD.AVAILABILITY.V_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS_DAY v
    CROSS JOIN janela j
    WHERE v.cluster_oem      = 'Assurua SG'
      AND v.validado         = 'True'
      AND v.data             BETWEEN j.ini::DATE AND DATEADD(day, -1, j.fim)::DATE
      AND v.motivo_parada    = 'FALHA'
      AND v.responsabilidade IN ('Fabricante', 'Nenhum dos dois')
      AND v.sistema_parada NOT IN ('ENVIRONMENT', 'GRID CONNECTION')
),
wtg_total AS (
    SELECT wtg, SUM(horas) AS h_total FROM eventos GROUP BY wtg
),
top_wtgs AS (
    SELECT wtg, h_total FROM wtg_total ORDER BY h_total DESC LIMIT 20
)
SELECT
    t.wtg                              AS aerogerador,
    e.sistema_parada                   AS subsistema,
    e.evento_parada                    AS alarme,
    COUNT(*)                           AS ocorrencias,
    MIN(e.data_inicio)::DATE           AS data_inicio,
    MAX(e.data_fim)::DATE              AS data_fim,
    ROUND(SUM(e.horas), 4)             AS horas_alarme,
    ROUND(t.h_total, 4)                AS horas_total_wtg
FROM eventos e
JOIN top_wtgs t ON t.wtg = e.wtg
GROUP BY t.wtg, t.h_total, e.sistema_parada, e.evento_parada
ORDER BY t.h_total DESC, horas_alarme DESC;
"""


# =============================================================================
# Pg 4/5 — WTGs COM MAIORES TEMPOS DE PARADA (top 10)
# Motivos: Falha + Corretiva Não Planejada + Grande Corretiva
# (exclui Corretiva Planejada e Preventiva/Preditiva)
# =============================================================================
Q_MAIORES_TEMPOS_CORRETIVA = """
WITH janela AS (
    SELECT %(sab_atual)s::TIMESTAMP_NTZ                  AS ini,
           DATEADD(day, 7, %(sab_atual)s)::TIMESTAMP_NTZ AS fim
),
eventos AS (
    SELECT
        v.wtg, v.evento_parada, v.sistema_parada, v.motivo_parada,
        v.data_inicio, v.data_fim,
        DATEDIFF('second',
            GREATEST(v.data_inicio, j.ini),
            LEAST(v.data_fim, j.fim)
        ) / 3600.0 AS horas
    FROM OEM_PROD.AVAILABILITY.V_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS_DAY v
    CROSS JOIN janela j
    WHERE v.cluster_oem      = 'Assurua SG'
      AND v.validado         = 'True'
      AND v.data             BETWEEN j.ini::DATE AND DATEADD(day, -1, j.fim)::DATE
      AND v.responsabilidade IN ('Fabricante', 'Nenhum dos dois')
      AND v.sistema_parada NOT IN ('ENVIRONMENT', 'GRID CONNECTION')
      AND v.motivo_parada  IN ('FALHA',
                               'MANUTENCAO CORRETIVA NAO PLANEJADA',
                               'GRANDE CORRETIVA')
),
wtg_total AS (
    SELECT wtg, SUM(horas) AS h_total FROM eventos GROUP BY wtg
),
top_wtgs AS (
    SELECT wtg, h_total FROM wtg_total ORDER BY h_total DESC LIMIT 10
)
SELECT
    t.wtg                              AS aerogerador,
    e.sistema_parada                   AS subsistema,
    e.motivo_parada                    AS motivo,
    e.evento_parada                    AS alarme,
    COUNT(*)                           AS ocorrencias,
    MIN(e.data_inicio)::DATE           AS data_inicio,
    MAX(e.data_fim)::DATE              AS data_fim,
    ROUND(SUM(e.horas), 4)             AS horas_alarme,
    ROUND(t.h_total, 4)                AS horas_total_wtg
FROM eventos e
JOIN top_wtgs t ON t.wtg = e.wtg
GROUP BY t.wtg, t.h_total, e.sistema_parada, e.motivo_parada, e.evento_parada
ORDER BY t.h_total DESC, horas_alarme DESC;
"""


# =============================================================================
# Pg 5 — TEMPERATURA: TOP 5 WTGs QUE CRUZARAM LIMITE (últimos 7 dias)
# -----------------------------------------------------------------------------
# Limites oficiais (mesma fonte do cockpit):
#   OEM_PROD.MACHINE_LEARNING.F_REAL_EOL_WTG_TEMPERATURE_ANOMALY_EVENTS_DAILY
#     DERATE_LIMIT   = limite amarelo (derate/alerta)
#     DOWNTIME_LIMIT = limite vermelho (downtime/crítico)
#     MAX_TEMP       = pico diário
# Pico via ML.MAX_TEMP; média via SCADA AGGREGATED='AVG'.
# Inválido: VALUE<=0 ou VALUE>=200. Sempre UPLOADED_AT mais recente.
# Ranking Top 5: mais sensores críticos, depois maior pico.
# Validado contra site oficial e planilha Temperature Analysis by WTG
# (A-II-02 pico Winding3=153,1 / A-II-03 trafo e nacele).
# Saída: 1 linha por WTG×componente em ALERTA ou CRÍTICO.
#   Colunas: WTG, COMPONENTE, PICO, MEDIA, DERATE, DOWNTIME, SEV, N_CRIT, N_ALERTA
#   SEV: 2=crítico, 1=alerta
# =============================================================================
Q_TEMPERATURA_TOP5 = """
WITH periodo AS (
    SELECT %(data_ini)s::DATE AS ini, %(data_fim)s::DATE AS fim
),
limites AS (
    SELECT DISTINCT VARIABLE, DERATE_LIMIT, DOWNTIME_LIMIT
    FROM OEM_PROD.MACHINE_LEARNING.F_REAL_EOL_WTG_TEMPERATURE_ANOMALY_EVENTS_DAILY
    WHERE GRUPO_OEM = 'Assurua SG'
      AND DATE = (
          SELECT MAX(DATE)
          FROM OEM_PROD.MACHINE_LEARNING.F_REAL_EOL_WTG_TEMPERATURE_ANOMALY_EVENTS_DAILY
          WHERE GRUPO_OEM = 'Assurua SG'
      )
),
pico_ml AS (
    SELECT m.WTG, m.VARIABLE, ROUND(MAX(m.MAX_TEMP), 1) AS pico
    FROM OEM_PROD.MACHINE_LEARNING.F_REAL_EOL_WTG_TEMPERATURE_ANOMALY_EVENTS_DAILY m,
         periodo p
    WHERE m.GRUPO_OEM = 'Assurua SG'
      AND m.DATE BETWEEN p.ini AND p.fim
    GROUP BY m.WTG, m.VARIABLE
),
last_upload AS (
    SELECT r.WTG, r.TIMESTAMP_LOCAL, r.VARIABLE, MAX(r.UPLOADED_AT) AS lu
    FROM OEM_PROD.SCADA.F_REAL_EOL_WTG_SCADA_RAW r, periodo p
    WHERE r.DATE_LOCAL BETWEEN p.ini AND p.fim
      AND r.GRUPO_OEM = 'Assurua SG'
      AND r.VARIABLE LIKE '%%TEMP%%'
    GROUP BY ALL
),
media_scada AS (
    SELECT r.WTG, r.VARIABLE,
           ROUND(AVG(CASE WHEN r.VALUE > 0 AND r.VALUE < 200 THEN r.VALUE END), 1) AS media
    FROM OEM_PROD.SCADA.F_REAL_EOL_WTG_SCADA_RAW r
    INNER JOIN last_upload lu USING (WTG, TIMESTAMP_LOCAL, VARIABLE),
         periodo p
    WHERE r.DATE_LOCAL BETWEEN p.ini AND p.fim
      AND r.GRUPO_OEM = 'Assurua SG'
      AND r.VARIABLE LIKE '%%TEMP%%'
      AND r.AGGREGATED = 'AVG'
      AND r.UPLOADED_AT = lu.lu
    GROUP BY r.WTG, r.VARIABLE
),
sensores AS (
    SELECT
        p.WTG, p.VARIABLE, p.pico, m.media,
        l.DERATE_LIMIT AS derate, l.DOWNTIME_LIMIT AS downtime,
        CASE
            WHEN p.VARIABLE LIKE 'GENERATOR_WINDING%%'   THEN 'Gerador enrol.'
            WHEN p.VARIABLE LIKE 'GENERATOR_BEARING%%'   THEN 'Gerador rolam.'
            WHEN p.VARIABLE = 'GENERATOR_SLIPRING_TEMP'  THEN 'Gerador slipring'
            WHEN p.VARIABLE LIKE 'GEARBOX%%'             THEN 'Gearbox'
            WHEN p.VARIABLE LIKE 'TRAFO%%'               THEN 'Trafo'
            WHEN p.VARIABLE = 'NACELLE_TEMP'             THEN 'Nacele'
            WHEN p.VARIABLE LIKE 'CONVERTER%%'           THEN 'Conversor'
            ELSE 'Outro'
        END AS componente,
        CASE
            WHEN l.DOWNTIME_LIMIT IS NOT NULL AND p.pico >= l.DOWNTIME_LIMIT THEN 2
            WHEN l.DERATE_LIMIT   IS NOT NULL AND p.pico >= l.DERATE_LIMIT   THEN 1
            ELSE 0
        END AS sev
    FROM pico_ml p
    LEFT JOIN limites l     ON l.VARIABLE = p.VARIABLE
    LEFT JOIN media_scada m ON m.WTG = p.WTG AND m.VARIABLE = p.VARIABLE
),
rank_wtg AS (
    SELECT WTG,
           SUM(CASE WHEN sev = 2 THEN 1 ELSE 0 END) AS n_crit,
           SUM(CASE WHEN sev = 1 THEN 1 ELSE 0 END) AS n_alerta,
           MAX(pico) AS pico_geral
    FROM sensores
    GROUP BY WTG
    HAVING n_crit > 0 OR n_alerta > 0
    ORDER BY n_crit DESC, pico_geral DESC
    LIMIT 5
)
SELECT
    s.WTG                          AS wtg,
    s.componente                   AS componente,
    MAX(s.pico)                    AS pico,
    MAX(s.media)                   AS media,
    MAX(s.derate)                  AS derate,
    MAX(s.downtime)                AS downtime,
    MAX(s.sev)                     AS sev,
    r.n_crit                       AS n_crit,
    r.n_alerta                     AS n_alerta
FROM sensores s
JOIN rank_wtg r ON r.WTG = s.WTG
WHERE s.sev >= 1
GROUP BY s.WTG, s.componente, r.n_crit, r.n_alerta
ORDER BY r.n_crit DESC, MAX(s.pico) DESC;
"""


# =============================================================================
# Catálogo
# =============================================================================
QUERIES = {
    "disponibilidade_semanal":   Q_DISPONIBILIDADE_SEMANAL,
    "top_alarmes_falha":         Q_TOP_ALARMES_FALHA,
    "wtgs_criticas":             Q_WTGS_CRITICAS,
    "mtbf_wtg":                  Q_MTBF_WTG,
    "maiores_tempos_falha":      Q_MAIORES_TEMPOS_FALHA,
    "maiores_tempos_corretiva":  Q_MAIORES_TEMPOS_CORRETIVA,
    "temperatura_top5":          Q_TEMPERATURA_TOP5,
}