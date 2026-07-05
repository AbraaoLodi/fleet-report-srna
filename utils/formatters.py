"""Formatadores numéricos no padrão brasileiro."""
from __future__ import annotations
 
import math
from datetime import date
 
 
def _is_nan(v) -> bool:
    return v is None or (isinstance(v, float) and math.isnan(v))
 
 
def fmt_float(v: float | None, decimals: int = 2) -> str:
    if _is_nan(v):
        return "—"
    s = f"{v:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")
 
 
def fmt_pct(v: float | None, decimals: int = 2) -> str:
    """Formata um número como percentual brasileiro: 94.84 → '94,84%'."""
    if _is_nan(v):
        return "—"
    return f"{fmt_float(v, decimals)}%"
 
 
def fmt_delta_pp(v: float | None, decimals: int = 2) -> str:
    """Variação em pontos percentuais: -1.65 → '▼ 1,65 p.p.'."""
    if _is_nan(v):
        return "—"
    arrow = "▲" if v >= 0 else "▼"
    return f"{arrow} {fmt_float(abs(v), decimals)} p.p."
 
 
def fmt_date_br(d: date | None) -> str:
    """Data → 'DD/MM/AAAA'."""
    if d is None:
        return "—"
    return d.strftime("%d/%m/%Y")
 
 
def fmt_date_curta(d: date | None) -> str:
    """Data → 'DD/MM'."""
    if d is None:
        return "—"
    return d.strftime("%d/%m")
 
 
def fmt_horas_hhmmss(horas: float | None) -> str:
    """Horas decimal → 'HH:MM:SS'.
 
    Usado nas tabelas detalhadas do Fleet Report (formato igual ao do
    relatório original da engenheira, ex: 17:17:10).
    """
    if _is_nan(horas):
        return "—"
    total_segundos = int(round(horas * 3600))
    h, resto = divmod(total_segundos, 3600)
    m, s = divmod(resto, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"
 
 
def fmt_horas_hhmm(horas: float | None) -> str:
    """Horas decimal → 'HH:MM' (sem segundos)."""
    if _is_nan(horas):
        return "—"
    total_min = int(round(horas * 60))
    h, m = divmod(total_min, 60)
    return f"{h:02d}:{m:02d}"
 
 
def abrev_subsistema(s) -> str:
    """Abrevia nomes de subsistema longos para manter linhas de tabela
    uniformes (evita quebra em 3 linhas de 'CONTROL & COMMUNICATION SYSTEM')."""
    if s is None or str(s).strip() in ("", "—", "nan", "None"):
        return "—"
    mapa = {
        "CONTROL & COMMUNICATION SYSTEM": "Control & Comm.",
        "FREQUENCY CONVERTER": "Freq. Converter",
        "POWER ELECTRICAL SYSTEM": "Power Electrical",
        "HYDRAULICS": "Hydraulics",
        "GENERATOR": "Generator",
        "PITCH SYSTEM": "Pitch System",
        "YAW SYSTEM": "Yaw System",
        "GEARBOX": "Gearbox",
        "BLADE": "Blade",
        "ENVIRONMENT": "Environment",
    }
    chave = str(s).strip().upper()
    if chave in mapa:
        return mapa[chave]
    t = str(s).strip().title()
    return t if len(t) <= 18 else t[:17] + "."