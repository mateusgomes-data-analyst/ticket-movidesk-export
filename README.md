# ticket-movidesk-export
Script de exportacao de dados do movidesk

# Movidesk Ticket Export

Script genérico e autocontido para consultar a [API pública de tickets do Movidesk](https://atendimento.movidesk.com/kb/pt-br/article/256/api-tickets) e exportar o resultado para arquivo.

Por ser um software utilizado em produção, minimizei o máximo de códigos para que pudesse ser disponibilizador. Portanto, essa versão:
1. Não aplica regras de negócio.
2. Não filtra por origem/status (além do período).
3. Não calcula indicadores.
4. Apenas busca os tickets no intervalo informado e grava uma tabela.

---

## O que este script faz

1. Autentica com o token da API Movidesk
2. Busca tickets pelo campo `createdDate` entre `--start-date` e `--end-date`
3. Pagina automaticamente (`$top` / `$skip`)
4. Achata campos aninhados (`owner`, `clients`, `serviceFull`, etc.) em colunas
5. Exporta para **Excel**, **CSV** ou **JSON**

## O que este script **não** faz

- Não usa nomes de empresas, produtos ou pessoas
- Não inclui lógica de “clientes em adoção” ou KPIs internos
- Não interpreta campos customizados (eles vão em JSON bruto na coluna `customFieldValues`)
- Não altera datas, tempos, status ou valores numéricos da API

---

## Requisitos

- Python 3.10+
- Dependências:

```bash
pip install requests pandas openpyxl
```

---

## Como obter o token

No Movidesk:

1. Acesse as configurações da conta
2. Vá em **Conta > Parâmetros > Ambiente > Chaves de acesso / Tokens** (o caminho pode variar conforme a versão)
3. Gere ou copie o token da API pública

Guarde o token com segurança.

---

## Uso rápido

```bash
cd movidesk-api-export

python movidesk_ticket_export.py \
  --token SEU_TOKEN_AQUI \
  --start-date 2025-01-01 \
  --end-date 2025-01-31
```

Por padrão gera um arquivo Excel:

```text
movidesk_tickets_YYYYMMDD_HHMMSS.xlsx
```

### Token por variável de ambiente

```bash
# Windows (PowerShell)
$env:MOVIDESK_API_TOKEN = "SEU_TOKEN_AQUI"
python movidesk_ticket_export.py --start-date 2025-01-01 --end-date 2025-01-31

# Linux / macOS
export MOVIDESK_API_TOKEN="SEU_TOKEN_AQUI"
python movidesk_ticket_export.py --start-date 2025-01-01 --end-date 2025-01-31
```

### Formatos de saída

```bash
# Excel
python movidesk_ticket_export.py --token TOKEN --start-date 2025-01-01 --end-date 2025-01-31 --format xlsx

# CSV (separador ; e UTF-8 com BOM)
python movidesk_ticket_export.py --token TOKEN --start-date 2025-01-01 --end-date 2025-01-31 --format csv

# JSON tabular
python movidesk_ticket_export.py --token TOKEN --start-date 2025-01-01 --end-date 2025-01-31 --format json
```

### Nome do arquivo de saída

```bash
python movidesk_ticket_export.py \
  --token TOKEN \
  --start-date 2025-01-01 \
  --end-date 2025-01-31 \
  --format csv \
  --output tickets_janeiro
```

### Também salvar o JSON bruto da API

```bash
python movidesk_ticket_export.py \
  --token TOKEN \
  --start-date 2025-01-01 \
  --end-date 2025-01-31 \
  --raw-json tickets_raw.json
```

### Fuso horário do filtro

O filtro OData usa offset de timezone (padrão `-03:00`, horário de Brasília):

```bash
python movidesk_ticket_export.py \
  --token TOKEN \
  --start-date 2025-01-01 \
  --end-date 2025-01-31 \
  --timezone-offset -03:00
```

### Tickets antigos (`/tickets/past`)

A API Movidesk separa:

| Endpoint | Uso |
|----------|-----|
| `/tickets` | tickets recentes (em geral até ~90 dias) |
| `/tickets/past` | histórico mais antigo |

O script **escolhe automaticamente** `/tickets/past` quando `--start-date` tem mais de 90 dias. Sem isso, um período antigo pode retornar só alguns tickets (ou nenhum).

```bash
# Forçar histórico
python movidesk_ticket_export.py --token TOKEN --start-date 2025-01-01 --end-date 2025-01-31 --past

# Forçar endpoint recente
python movidesk_ticket_export.py --token TOKEN --start-date 2026-06-01 --end-date 2026-06-30 --no-past
```

---

## Colunas exportadas

| Coluna | Origem na API |
|--------|----------------|
| `id` | `id` |
| `protocol` | `protocol` |
| `subject` | `subject` |
| `status` | `status` |
| `baseStatus` | `baseStatus` |
| `category` | `category` |
| `justification` | `justification` |
| `origin` | código numérico de origem |
| `originLabel` | rótulo oficial do código de origem |
| `createdDate` | `createdDate` |
| `resolvedIn` | `resolvedIn` |
| `lastUpdate` | `lastUpdate` |
| `slaResponseDate` | `slaResponseDate` |
| `slaSolutionDate` | `slaSolutionDate` |
| `ownerTeam` | `ownerTeam` |
| `owner` | `owner.businessName` |
| `createdBy` | `createdBy.businessName` |
| `clientId` | primeiro item de `clients.id` |
| `clientName` | primeiro cliente (`businessName`) |
| `clientPersonType` | `clients[0].personType` |
| `serviceFirstLevel` | `serviceFirstLevel` |
| `serviceSecondLevel` | `serviceSecondLevel` |
| `serviceThirdLevel` | `serviceThirdLevel` |
| `serviceFull` | `serviceFull` (lista unida por ` > `) |
| `lifetimeWorkingTime` | `lifetimeWorkingTime` (segundos) |
| `resolvedInFirstCall` | `resolvedInFirstCall` |
| `chatWaitingTime` | `chatWaitingTime` |
| `chatTalkTime` | `chatTalkTime` |
| `tags` | `tags` (lista unida por vírgula) |
| `cc` | `cc` |
| `customFieldValues` | JSON bruto dos campos customizados |

Os nomes das colunas seguem o vocabulário da própria API, para facilitar o cruzamento com a documentação oficial.

---

## Parâmetros CLI

| Parâmetro | Obrigatório | Descrição |
|-----------|-------------|-----------|
| `--token` | sim* | Token da API (`MOVIDESK_API_TOKEN` também vale) |
| `--start-date` | sim | Data inicial `YYYY-MM-DD` |
| `--end-date` | sim | Data final `YYYY-MM-DD` |
| `--timezone-offset` | não | Offset do filtro (padrão `-03:00`) |
| `--format` | não | `xlsx`, `csv` ou `json` (padrão `xlsx`) |
| `--output` | não | Caminho/nome do arquivo de saída |
| `--raw-json` | não | Salva também o payload bruto da API |
| `--past` | não | Força endpoint `/tickets/past` |
| `--no-past` | não | Força endpoint `/tickets` |

\* Obrigatório via argumento ou variável de ambiente.

---

## Limites e boas práticas

- A API Movidesk possui limite de requisições. O script espera **0,5s** entre chamadas e pagina de **1000** em **1000**.
- Intervalos muito grandes podem demorar. Prefira exportar por mês ou por semana.
- Se receber erro HTTP 429/403, reduza a frequência ou verifique permissões do token.
- Não compartilhe o token publicamente.

---

## Estrutura desta pasta

```text
movidesk-api-export/
├── movidesk_ticket_export.py   # script único
└── README.md                   # este manual
```

Você pode copiar apenas esta pasta para outro repositório Git. Não há dependência do restante do projeto.

---

## Licença sugerida

Se for publicar, recomenda-se MIT (ou a licença do repositório pai).

---

## Referência

- Documentação oficial da API de tickets Movidesk  
  https://atendimento.movidesk.com/kb/pt-br/article/256/api-tickets
