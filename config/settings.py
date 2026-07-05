"""Configurações globais — lê do .env e expõe dataclasses tipadas."""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

# Carrega .env se python-dotenv estiver instalado (opcional)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


PROJECT_ROOT = Path(__file__).resolve().parent.parent


@dataclass(frozen=True)
class SnowflakeConfig:
    account: str | None
    user: str | None
    authenticator: str | None = "externalbrowser"
    role: str | None = None
    warehouse: str | None = None
    database: str | None = None
    schema: str | None = None


@dataclass(frozen=True)
class AppConfig:
    log_level: str
    output_dir: Path
    inputs_dir: Path
    # Específicos do Fleet Report SRNA
    cluster_oem: str          # 'Assurua SG' (validado contra Luisa + engenheira)
    n_wtgs_frota: int         # 49 (cadastro D_EOL_WTG)


SNOWFLAKE_CONFIG = SnowflakeConfig(
    account=os.environ.get("SNOWFLAKE_ACCOUNT"),
    user=os.environ.get("SNOWFLAKE_USER"),
    authenticator=os.environ.get("SNOWFLAKE_AUTHENTICATOR", "externalbrowser"),
    role=os.environ.get("SNOWFLAKE_ROLE"),
    warehouse=os.environ.get("SNOWFLAKE_WAREHOUSE"),
    database=os.environ.get("SNOWFLAKE_DATABASE", "OEM_PROD"),
    schema=os.environ.get("SNOWFLAKE_SCHEMA"),
)

APP_CONFIG = AppConfig(
    log_level=os.environ.get("LOG_LEVEL", "INFO"),
    output_dir=PROJECT_ROOT / "output",
    inputs_dir=PROJECT_ROOT / "inputs",
    cluster_oem=os.environ.get("CLUSTER_OEM", "Assurua SG"),
    n_wtgs_frota=int(os.environ.get("N_WTGS_FROTA", "49")),
)
