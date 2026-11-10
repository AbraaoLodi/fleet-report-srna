# Fleet Reliability Report — SRNA

Geração automática do relatório semanal de confiabilidade do **Complexo Eólico
Assuruá SG** (49 WTGs Siemens Gamesa G97), a partir de dados do Snowflake.

O script gera um documento Word (`.docx`) com as seções de disponibilidade,
taxas de falha/corretiva, MTBF e análise de temperatura, seguindo a identidade
visual da Serena.

---
python gerar_fleet_report.py 7d 2026 SG            # últimos 7 dias (termina ontem)
python gerar_fleet_report.py 15d 2026 SG           # últimos 15 dias (termina ontem)
python gerar_fleet_report.py s38 2026 SG           # semana SG fixa (clássico)

## Estrutura das páginas

| Página | Conteúdo |
|--------|----------|
| Capa   | Capa institucional Serena |
| Pg 2   | 01. Disponibilidade Fabricante + 4 taxas |
| Pg 3   | 02. Taxa de Falhas e Corretivas |
| Pg 4   | 03. MTBF e WTGs com Maiores Tempos de Parada |
| Pg 5   | 04. Análise de Temperatura (limites oficiais por sensor) |

---

## Requisitos

- Python 3.11+
- Acesso ao Snowflake da Serena (SSO / `externalbrowser`)
- Dependências em `requirements.txt`

---

## Instalação

```powershell
# 1. Clonar o repositório
git clone https://github.com/SEU_USUARIO/fleet-report-srna.git
cd fleet-report-srna

# 2. Criar e ativar o ambiente virtual
python -m venv .venv
.\.venv\Scripts\Activate.ps1        # Windows PowerShell

# 3. Instalar dependências
pip install -r requirements.txt

# 4. Configurar credenciais
copy .env.example .env               # depois edite o .env com seus dados
```

---

## Configuração (`.env`)

Copie `.env.example` para `.env` e preencha. As principais variáveis:

| Variável | Descrição |
|----------|-----------|
| `SNOWFLAKE_ACCOUNT` | Identificador da conta Snowflake |
| `SNOWFLAKE_USER` | Usuário (e-mail Serena) |
| `SNOWFLAKE_AUTHENTICATOR` | `externalbrowser` (SSO) |
| `CLUSTER_OEM` | Cluster OEM alvo (padrão `Assurua SG`) |
| `N_WTGS_FROTA` | Nº de WTGs nominais (padrão `49`) |

> O `.env` contém credenciais e **não é versionado** (está no `.gitignore`).

---

## Uso

```powershell
# Gerar o relatório de uma semana (conecta no Snowflake)
python gerar_fleet_report.py --semana 27 --ano 2026

# Testar o layout sem Snowflake (dados de exemplo)
python gerar_fleet_report.py --semana 27 --mock
```

O `.docx` é salvo em `output/Fleet_Report_SRNA_S{semana}_{ano}.docx`.

> Feche o arquivo no Word antes de rodar novamente (senão dá erro de permissão).

---

## Convenções

- **Semana SG:** sábado a sexta-feira (não ISO). Ex: S27/2026 = 27/06 a 03/07.
- **Filtros de governança** (relatório contratual OEM): fonte validada
  `V_REAL_EOL_WTG_EVENTLOG_VALIDATED_ALERTS_DAY`, com
  `RESPONSABILIDADE IN ('Fabricante', 'Nenhum dos dois')`.
- **Temperatura:** limites oficiais por sensor (mesma fonte do cockpit),
  pico via agregação `MAX`, média via `AVG`.

---

## Organização do código

```
config/     Configurações e leitura do .env
data/        Conexão Snowflake + catálogo de queries SQL
secoes/      Um módulo por seção do relatório (pg2..pg5)
utils/       Helpers: docx, gráficos, formatadores, logger, semana SG
gerar_fleet_report.py   Orquestrador principal
```

---

## Status

Projeto interno de Gestão de Ativos (SRNA). Números validados à vírgula contra
o cockpit Serena e o relatório da engenharia de O&M.
