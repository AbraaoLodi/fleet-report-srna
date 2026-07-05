"""Conexão Snowflake via SSO (externalbrowser).

Replica o padrão do gerar_reuniao.py — mesmo módulo, mesmo comportamento.
Mantém conexão compartilhada thread-safe.
"""
from __future__ import annotations

import threading

import pandas as pd
import snowflake.connector
from snowflake.connector import DictCursor

from config.settings import SNOWFLAKE_CONFIG
from utils.logger import get_logger

log = get_logger(__name__)


class SnowflakeUnavailable(RuntimeError):
    """Erro amigável quando Snowflake não está acessível."""


_lock = threading.Lock()
_shared_conn: snowflake.connector.SnowflakeConnection | None = None


def _build_kwargs() -> dict:
    cfg = SNOWFLAKE_CONFIG
    if not cfg.account or not cfg.user:
        raise SnowflakeUnavailable(
            "SNOWFLAKE_ACCOUNT ou SNOWFLAKE_USER não definidos no .env."
        )
    kwargs: dict = {
        "account": cfg.account,
        "user": cfg.user,
        "authenticator": cfg.authenticator or "externalbrowser",
        "client_session_keep_alive": True,
        "login_timeout": 60,
        "network_timeout": 600,   # 10 min — warehouse dormindo pode demorar
    }
    if cfg.role:
        kwargs["role"] = cfg.role
    if cfg.warehouse:
        kwargs["warehouse"] = cfg.warehouse
    if cfg.database:
        kwargs["database"] = cfg.database
    if cfg.schema:
        kwargs["schema"] = cfg.schema
    return kwargs


def _connect() -> snowflake.connector.SnowflakeConnection:
    log.info("Abrindo conexão Snowflake (SSO)…")
    conn = snowflake.connector.connect(**_build_kwargs())
    log.info("Conexão Snowflake estabelecida.")
    return conn


def get_connection() -> snowflake.connector.SnowflakeConnection:
    """Retorna conexão compartilhada, reabrindo se necessário."""
    global _shared_conn
    with _lock:
        if _shared_conn is None or _shared_conn.is_closed():
            _shared_conn = _connect()
        return _shared_conn


def close_connection() -> None:
    global _shared_conn
    with _lock:
        if _shared_conn is not None and not _shared_conn.is_closed():
            try:
                _shared_conn.close()
            except Exception:
                pass
        _shared_conn = None


def run_query(sql: str, params: dict | None = None) -> pd.DataFrame:
    """Executa SQL e devolve DataFrame.

    Faz substituição manual de %(nome)s -> 'valor', mantendo compatibilidade
    com queries que usam esse formato (padrão Serena).
    """
    try:
        if params:
            for key, value in params.items():
                if isinstance(value, list):
                    safe_items = [str(v).replace("'", "''") for v in value]
                    formatted = ",".join(f"'{s}'" for s in safe_items)
                    sql = sql.replace(f"%({key})s", formatted)
                elif isinstance(value, str):
                    safe = value.replace("'", "''")
                    sql = sql.replace(f"%({key})s", f"'{safe}'")
                else:
                    sql = sql.replace(f"%({key})s", str(value))

        conn = get_connection()
        log.debug("SQL preview: %s", sql[:140].replace("\n", " "))

        cursor = conn.cursor(DictCursor)
        try:
            cursor.execute(sql)
            rows = cursor.fetchall()
            df = pd.DataFrame(rows)
            if not df.empty:
                df.columns = [str(c).upper() for c in df.columns]
            log.info("Query OK — %d linhas.", len(df))
            return df
        finally:
            cursor.close()

    except SnowflakeUnavailable:
        raise
    except Exception as exc:
        log.exception("Erro na execução da query.")
        raise SnowflakeUnavailable(f"Falha na query: {exc}") from exc
