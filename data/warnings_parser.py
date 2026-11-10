"""Parser do CSV/XLS de warnings do SCADA do fabricante.

Reconhece automaticamente os formatos (pelas colunas presentes) e normaliza
para o mesmo schema interno antes do ranking:

- Siemens Gamesa (SG): Usina eletrica, Dispositivo, Categoria, Evento,
  Subevento/Categorizacao, Data de inicio, Data Final, Duracao, Descricao.
  Categorias: 'Estado' (descartado), 'Aviso', 'Alarme'.
  ** O export pode vir em INGLES ** (Power Plant, Device, Category, Event,
  Start Date...). Nesse caso, _traduzir_cabecalhos() converte os nomes e as
  categorias (State/Warning/Alarm -> Estado/Aviso/Alarme) antes de tudo.

- Vestas (VE): Unit, Serial no., Code, Description, Detected, Device ack., ...
"""
from __future__ import annotations
import re
import json
import pandas as pd
from pathlib import Path

_MAPA_PATH = Path(__file__).resolve().parent / "mapa_alarmes_sg.json"
try:
    _MAPA_EN = json.load(open(_MAPA_PATH, encoding="utf-8"))
except Exception:
    _MAPA_EN = {}

_ROMANOS = {1:"I",2:"II",3:"III",4:"IV",5:"V",6:"VI",7:"VII",8:"VIII",9:"IX",10:"X",
            11:"XI",12:"XII",13:"XIII",14:"XIV",15:"XV"}

# --- Tradução de cabeçalhos EN -> PT (formato SG exportado em inglês) --------
# Mapeia pelo nome exato (após strip). Só afeta o export SG em inglês; o
# formato Vestas e o SG já em português passam intactos (nenhuma chave bate,
# ou já estão em PT).
_CABECALHOS_EN_PT = {
    "Power Plant": "Usina elétrica",
    "Device": "Dispositivo",
    "Category": "Categoria",
    "Event": "Evento",
    "Subevent / Categorization": "Subevento / Categorização",
    "Start Date": "Data de início",
    "End Date": "Data Final",
    "Duration": "Duração",
    "Categorization description": "Descrição",
}
_CATEGORIAS_EN_PT = {
    "State": "Estado", "Warning": "Aviso", "Alarm": "Alarme",
    # já em PT (idempotente)
    "Estado": "Estado", "Aviso": "Aviso", "Alarme": "Alarme",
}


def _traduzir_cabecalhos(df: pd.DataFrame) -> pd.DataFrame:
    """Se o CSV veio em inglês (formato SG), converte nomes de coluna e a
    coluna Categoria para português. Não toca no formato Vestas (colunas
    diferentes) nem no SG que já veio em português."""
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    # traduz nomes de coluna que casam com o mapa EN->PT
    renomear = {c: _CABECALHOS_EN_PT[c] for c in df.columns if c in _CABECALHOS_EN_PT}
    if renomear:
        df = df.rename(columns=renomear)
    # traduz valores da coluna Categoria (State/Warning/Alarm -> PT)
    if "Categoria" in df.columns:
        df["Categoria"] = df["Categoria"].map(
            lambda x: _CATEGORIAS_EN_PT.get(str(x).strip(), str(x).strip()))
    return df


def _ler(path: str) -> pd.DataFrame:
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(p)
    for sep in (";", ","):
        try:
            df = pd.read_csv(p, sep=sep, encoding="utf-8-sig")
            if df.shape[1] >= 6:
                return df
        except Exception:
            pass
    return pd.read_csv(p, sep=";", encoding="utf-8-sig")

def _split_codigo(evento: str):
    s = str(evento).strip()
    m = re.match(r"^(\d+)\s+(.*)$", s)
    if not m:
        return "-", s
    cod, resto = m.group(1), m.group(2)
    if cod in _MAPA_EN:
        return cod, _MAPA_EN[cod]
    mbr = re.match(r"^\[Br\]\s*(.+)$", resto)
    if mbr:
        return cod, mbr.group(1).strip()
    return cod, resto.strip()

def _unit_vestas_para_exibicao(unit: str) -> str:
    m = re.match(r"^AS4(\d{2})_(\d{2})$", str(unit).strip())
    if not m:
        return str(unit).strip()
    circuito, pos = int(m.group(1)), int(m.group(2))
    romano = _ROMANOS.get(circuito, str(circuito))
    return f"AS4-{romano}-{pos:02d}"

def _detectar_formato(df: pd.DataFrame) -> str:
    cols = set(c.strip() for c in df.columns)
    if {"Unit", "Code", "Description", "Detected"}.issubset(cols):
        return "VE"
    if any("Dispositivo" in c for c in cols) or any("Usina" in c for c in cols):
        return "SG"
    raise ValueError(
        "Formato de warnings nao reconhecido -- colunas nao batem com SG "
        f"nem Vestas. Colunas encontradas: {sorted(cols)}"
    )

def _normalizar_sg(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    col_disp = next(c for c in df.columns if "Dispositivo" in c or "WTG" in c)
    col_cat  = next(c for c in df.columns if "Categoria" in c)
    col_evt  = next(c for c in df.columns if c == "Evento" or ("Evento" in c and "Sub" not in c))
    col_ini  = next(c for c in df.columns if "início" in c.lower() or "inicio" in c.lower())

    out = pd.DataFrame()
    out["wtg"] = df[col_disp].astype(str)
    out["categoria"] = df[col_cat].map({"Aviso":"Aviso","Alarme":"Alarme"}).fillna(df[col_cat])
    out["evento_bruto"] = df[col_evt].astype(str)
    out["data_ini"] = pd.to_datetime(df[col_ini], errors="coerce")
    out = out[out["categoria"].isin(["Aviso","Alarme"])]
    out["usina"] = df["Usina elétrica"].iloc[0] if "Usina elétrica" in df.columns and len(df) else ""
    return out

def _normalizar_vestas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    out = pd.DataFrame()
    out["wtg"] = df["Unit"].map(_unit_vestas_para_exibicao)
    out["categoria"] = "Aviso"
    cod = df["Code"].astype("Int64").astype(str)
    desc = df["Description"].astype(str).str.strip()
    out["evento_bruto"] = cod + " " + desc
    out["data_ini"] = pd.to_datetime(df["Detected"], format="%d/%m/%Y %H:%M:%S.%f", errors="coerce")
    if out["data_ini"].isna().all():
        out["data_ini"] = pd.to_datetime(df["Detected"], dayfirst=True, errors="coerce")
    data_fim_evt = pd.to_datetime(df["Device ack."], format="%d/%m/%Y %H:%M:%S.%f", errors="coerce")
    if data_fim_evt.isna().all():
        data_fim_evt = pd.to_datetime(df["Device ack."], dayfirst=True, errors="coerce")
    out["duracao_h"] = (data_fim_evt - out["data_ini"]).dt.total_seconds() / 3600.0
    out["usina"] = "Assuruá VE"
    return out

def processar_warnings(path: str, data_ini=None, data_fim=None, top_n: int = 10) -> dict:
    df_bruto = _ler(path)
    df_bruto = _traduzir_cabecalhos(df_bruto)   # <<< EN->PT se necessário
    formato = _detectar_formato(df_bruto)
    df = _normalizar_sg(df_bruto) if formato == "SG" else _normalizar_vestas(df_bruto)

    if data_ini and data_fim:
        di, dfim = pd.to_datetime(data_ini), pd.to_datetime(data_fim) + pd.Timedelta(days=1)
        m = (df["data_ini"] >= di) & (df["data_ini"] < dfim)
        if m.sum() > 0:
            df = df[m]

    usina = str(df["usina"].iloc[0]) if "usina" in df.columns and len(df) else ""
    df["_cod"], df["_desc"] = zip(*df["evento_bruto"].map(_split_codigo))

    def _rank(sub, n):
        if sub.empty:
            return pd.DataFrame(columns=["evento_bruto","ocor","wtgs","cod","desc"])
        g = (sub.groupby(["evento_bruto"]).agg(ocor=("evento_bruto","size"), wtgs=("wtg","nunique")).reset_index())
        g[["cod","desc"]] = g["evento_bruto"].apply(lambda e: pd.Series(_split_codigo(e)))
        g = g.sort_values(["ocor","wtgs"], ascending=[False,False]).head(n)
        return g

    alarmes = df[df["categoria"]=="Alarme"]
    avisos  = df[df["categoria"]=="Aviso"]

    rank_alarmes = _rank(alarmes, top_n)
    rank_avisos  = _rank(avisos, top_n)

    wtg_alarmes = (alarmes.groupby("wtg").size().reset_index(name="n_alarmes")
                   .sort_values("n_alarmes", ascending=False).head(top_n))

    def _wtgs_por_evento(sub, eventos):
        out = {}
        for e in eventos:
            wl = sorted(sub[sub["evento_bruto"]==e]["wtg"].unique().tolist())
            out[e] = wl
        return out

    return {
        "formato": formato,
        "usina": usina,
        "n_avisos": int(len(avisos)),
        "n_alarmes": int(len(alarmes)),
        "n_wtgs": int(df["wtg"].nunique()),
        "rank_alarmes": rank_alarmes.to_dict("records"),
        "rank_avisos": rank_avisos.to_dict("records"),
        "wtg_alarmes": wtg_alarmes.to_dict("records"),
        "wtgs_por_alarme": _wtgs_por_evento(alarmes, rank_alarmes["evento_bruto"].tolist()),
        "wtgs_por_aviso": _wtgs_por_evento(avisos, rank_avisos["evento_bruto"].tolist()),
        "col_disp": "wtg", "col_evt": "evento_bruto",
    }