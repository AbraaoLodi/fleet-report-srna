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

# ============================================================
# Catálogo de frotas do Complexo Assuruá
# ============================================================
from dataclasses import dataclass

@dataclass(frozen=True)
class FrotaConfig:
    cluster_oem: str
    titulo_capa: str
    sub_capa: str
    label_header: str
    exclusao_alarmes: tuple = ()

FROTAS = {
    "SG":    FrotaConfig("Assurua SG",    "Assuruá SG  |  G97",      "Siemens Gamesa · G97",   "Assuruá SG | G97",
                         exclusao_alarmes=("908_GROUND EMERGENCY BUTTON",)),
    "GE":    FrotaConfig("Assurua GE",    "Assuruá GE  |  2.X",      "GE · 2.X",               "Assuruá GE | 2.X",
                         exclusao_alarmes=("140_MOTOR PROTECTION PITCH MOTOR",)),
    "GE158": FrotaConfig("Assurua GE158", "Assuruá GE158  |  Cypress","GE · Cypress",          "Assuruá GE158 | Cypress"),
    "VE":    FrotaConfig("Assurua VE",    "Assuruá VE  |  Vestas",   "Vestas",                 "Assuruá VE | Vestas"),
}

ALIAS = {
    "sg":"SG","siemens":"SG","gamesa":"SG","sgre":"SG","g97":"SG",
    "ge":"GE","ge2x":"GE","ge 2x":"GE","ge2.x":"GE",
    "ge158":"GE158","cypress":"GE158","ge cypress":"GE158",
    "ve":"VE","vestas":"VE",
}

def resolver_frota(texto: str) -> str:
    t = texto.strip().lower().replace("assuruá","").replace("assurua","").strip()
    if t.upper() in FROTAS: return t.upper()
    if t in ALIAS: return ALIAS[t]
    for k in FROTAS:
        if k.lower() in t: return k
    raise ValueError(f"Frota não reconhecida: '{texto}'. Opções: {list(FROTAS)}")

META_DISP = 98.00