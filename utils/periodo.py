"""Resolução do modo de operação do relatório.

O primeiro argumento da linha de comando define o modo:
  - "7d"  → janela móvel dos últimos 7 dias (terminando ONTEM)
  - "15d" → janela móvel dos últimos 15 dias (terminando ONTEM)
  - "sNN" → semana SG fixa NN (sábado-sexta), comportamento clássico
            (também aceita "NN" puro e intervalo "NN-MM" para quinzenal SG)

Em todos os modos, o objeto Periodo entrega:
  - foco_ini / foco_fim : janela principal das seções 2-5 (timestamps)
  - historico           : lista de (rotulo, ini, fim) para o card da pg 1,
                          da mais antiga para a mais recente; a última é a
                          coluna em foco.
  - rotulo_foco         : rótulo curto da coluna em foco ("Últimos 7d", "S38")
  - rotulo_arquivo      : sufixo do nome do arquivo ("7d", "S38")
  - eh_movel            : True para 7d/15d, False para semana fixa
"""
from __future__ import annotations
from dataclasses import dataclass
from datetime import date, timedelta

from utils.semana_sg import janela_da_semana, semana_de


@dataclass(frozen=True)
class ColunaHistorico:
    rotulo: str          # "S38" ou "Últimos 7d"
    sub: str             # "12/09 – 18/09 · semana SG"
    ini: date
    fim: date            # inclusivo (último dia contado)
    eh_foco: bool = False


@dataclass(frozen=True)
class Periodo:
    modo: str                    # "movel" | "semana"
    n_dias: int | None           # 7 ou 15 no modo móvel; None na semana
    foco_ini: date
    foco_fim: date               # inclusivo
    historico: list[ColunaHistorico]
    rotulo_foco: str
    rotulo_arquivo: str
    ano: int

    @property
    def eh_movel(self) -> bool:
        return self.modo == "movel"


def _hoje() -> date:
    return date.today()


def resolver_periodo(arg: str, ano: int) -> Periodo:
    """Interpreta o 1º argumento e devolve o Periodo correspondente.

    Aceita: "7d", "15d", "sNN", "NN", "NN-MM".
    Lança ValueError com mensagem clara para qualquer outra coisa.
    """
    a = arg.strip().lower()

    # ---- Modo móvel: 7d ou 15d ----
    if a in ("7d", "15d"):
        n = int(a[:-1])
        fim = _hoje() - timedelta(days=1)          # ontem, inclusivo
        ini = fim - timedelta(days=n - 1)          # janela de n dias
        # 3 últimas semanas SG JÁ fechadas (a corrente, aberta, fica de fora)
        sem_corrente = semana_de(_hoje())
        historico = []
        for k in (3, 2, 1):
            j = janela_da_semana(sem_corrente.semana - k, ano)
            historico.append(ColunaHistorico(
                rotulo=f"S{j.semana:02d}", sub=f"{j.inicio:%d/%m} – {j.fim:%d/%m} · semana SG",
                ini=j.inicio, fim=j.fim, eh_foco=False))
        # coluna em foco: a janela móvel
        historico.append(ColunaHistorico(
            rotulo=f"Últimos {n}d", sub=f"{ini:%d/%m} – {fim:%d/%m} · janela móvel",
            ini=ini, fim=fim, eh_foco=True))
        return Periodo(modo="movel", n_dias=n, foco_ini=ini, foco_fim=fim,
                       historico=historico, rotulo_foco=f"Últimos {n}d",
                       rotulo_arquivo=f"{n}d", ano=ano)

    # ---- Modo semana fixa: sNN, NN, NN-MM (ou sNN-sMM) ----
    def _num_semana(txt: str) -> int:
        t = txt.strip().lstrip("s")
        if not t.isdigit():
            raise ValueError(
                f"Modo inválido: '{arg}'. Use '7d', '15d' ou 'sNN' (ex.: 's38').")
        return int(t)

    if "-" in a:
        p_ini, p_fim = a.split("-", 1)
        ini_sem, fim_sem = _num_semana(p_ini), _num_semana(p_fim)
    else:
        ini_sem = fim_sem = _num_semana(a)

    j_ini = janela_da_semana(ini_sem, ano)
    j_fim = janela_da_semana(fim_sem, ano)
    # 4 semanas SG (as 3 anteriores + a própria), como no comportamento clássico
    historico = []
    for k in (3, 2, 1, 0):
        j = janela_da_semana(fim_sem - k, ano)
        historico.append(ColunaHistorico(
            rotulo=f"S{j.semana:02d}", sub=f"{j.inicio:%d/%m} – {j.fim:%d/%m} · semana SG",
            ini=j.inicio, fim=j.fim, eh_foco=(k == 0)))
    rot = f"S{ini_sem:02d}" if ini_sem == fim_sem else f"S{ini_sem:02d}-S{fim_sem:02d}"
    return Periodo(modo="semana", n_dias=None, foco_ini=j_ini.inicio, foco_fim=j_fim.fim,
                   historico=historico, rotulo_foco=rot, rotulo_arquivo=rot, ano=ano)