"""Convenção de semana SG da Serena Energia.

ATENÇÃO: esta é a convenção CORRETA para o Fleet Report SRNA, validada
contra o Weekly Report da Luisa (S17-S19) e o Fleet Report da engenheira
de O&M (S22-S24) à vírgula.

Regra: a semana 1 é a primeira semana SG (sábado-sexta) que CONTÉM o dia
01/01 do ano. Em 2026, S1 começa em sábado 27/12/2025 (porque 01/01/2026
caiu numa quinta-feira).

NOTA: o script gerar_reuniao.py (reuniões semanais de time) usa uma
convenção DIFERENTE — onde S1 = primeira semana que começa no ano. Não
copiar de lá!
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta


# ---------------------------------------------------------------------------
# Cálculo do início da semana 1 do ano
# ---------------------------------------------------------------------------

def _primeiro_sabado_sg(ano: int) -> date:
    """Sábado em que começa a S1 do ano (pode estar no ano anterior).

    Regra: o sábado mais próximo de 01/01 (anterior ou no dia).
    """
    primeiro_jan = date(ano, 1, 1)
    weekday = primeiro_jan.weekday()  # seg=0 … sáb=5 … dom=6
    if weekday == 5:                  # 01/01 já é sábado
        return primeiro_jan
    # Quantos dias voltar para chegar no sábado anterior
    dias_voltar = (weekday + 2) % 7   # qui=3 -> 5 ; dom=6 -> 1 ; etc.
    if dias_voltar == 0:
        dias_voltar = 7
    return primeiro_jan - timedelta(days=dias_voltar)


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class JanelaSemanal:
    """Uma semana SG numerada."""
    semana: int
    ano: int                  # ano "SG", não calendário (pode diferir em jan)
    inicio: date              # sábado
    fim: date                 # sexta

    def __str__(self) -> str:
        return (f"S{self.semana:02d}/{self.ano} "
                f"({self.inicio.strftime('%d/%m')}-{self.fim.strftime('%d/%m')})")


def janela_da_semana(semana: int, ano: int) -> JanelaSemanal:
    """Devolve (sábado, sexta) da semana SG numerada."""
    primeiro_sab = _primeiro_sabado_sg(ano)
    inicio = primeiro_sab + timedelta(days=7 * (semana - 1))
    fim = inicio + timedelta(days=6)
    return JanelaSemanal(semana=semana, ano=ano, inicio=inicio, fim=fim)


def semana_de(d: date) -> JanelaSemanal:
    """Devolve a JanelaSemanal que contém a data d.

    Considera os limites de ano (S1 de 2026 começa em 27/12/2025).
    """
    # Sábado da semana de d
    weekday = d.weekday()
    if weekday == 5:
        sabado = d
    else:
        dias_voltar = (weekday + 2) % 7
        if dias_voltar == 0:
            dias_voltar = 7
        sabado = d - timedelta(days=dias_voltar)

    # Tentar enquadrar em diferentes anos SG
    for ano_tentativa in (d.year, d.year + 1, d.year - 1):
        primeiro_sab = _primeiro_sabado_sg(ano_tentativa)
        proximo_primeiro_sab = _primeiro_sabado_sg(ano_tentativa + 1)
        if primeiro_sab <= sabado < proximo_primeiro_sab:
            delta_semanas = (sabado - primeiro_sab).days // 7
            num_semana = delta_semanas + 1
            return JanelaSemanal(
                semana=num_semana,
                ano=ano_tentativa,
                inicio=sabado,
                fim=sabado + timedelta(days=6),
            )

    raise ValueError(f"Não foi possível calcular a semana SG de {d}")


def semana_anterior(j: JanelaSemanal) -> JanelaSemanal:
    """Devolve a semana imediatamente anterior à janela j."""
    return semana_de(j.inicio - timedelta(days=1))


def janelas_relatorio(hoje: date,
                       n_anteriores: int = 3) -> dict:
    """Calcula janelas usadas no Fleet Report.

    Devolve:
        {
            'corrente': JanelaSemanal,        # semana onde 'hoje' está
            'corrente_ate': date,             # 'hoje' (limite da parcial)
            'corrente_completa': bool,        # True se hoje é a sexta da semana
            'anteriores': [JanelaSemanal, ...]  # n últimas semanas fechadas
        }

    Exemplo: hoje = segunda 22/06/2026
        corrente   = S26 (20/06-26/06), ate=22/06, completa=False
        anteriores = [S25, S24, S23]
    """
    corrente = semana_de(hoje)
    corrente_completa = (hoje >= corrente.fim)

    anteriores: list[JanelaSemanal] = []
    cursor = corrente
    for _ in range(n_anteriores):
        cursor = semana_anterior(cursor)
        anteriores.append(cursor)

    return {
        "corrente": corrente,
        "corrente_ate": hoje,
        "corrente_completa": corrente_completa,
        "anteriores": anteriores,
    }
