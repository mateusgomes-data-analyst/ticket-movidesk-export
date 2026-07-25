"""
Movidesk Ticket Export (generico)

Essa não é a ferramenta utilizada em produção por mim. Eu diminui muito o código para poder não trazer
dados sensíveis e apenas testar e estruturar a retirada de tickets.
Para retirar os tickets, vai precisar de um token de API. Para obter o token, voce deve solicitar ao suporte da Movidesk.


Para mais informações, entre em contato comigo no linkedin: www.linkedin.com/in/mateus-gomes-279349218

Requirements:
    pip install requests pandas openpyxl

exemplos de uso:
    python movidesk_ticket_export.py --token SEU_TOKEN --start-date 2025-01-01 --end-date 2025-01-31
    python movidesk_ticket_export.py --token SEU_TOKEN --start-date 2025-01-01 --end-date 2025-01-31 --format csv
    set MOVIDESK_API_TOKEN=YOUR_TOKEN
    python movidesk_ticket_export.py --start-date 2025-01-01 --end-date 2025-01-31
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------
API_BASE_URL = "https://api.movidesk.com/public/v1"
TICKETS_URL = f"{API_BASE_URL}/tickets"
TICKETS_PAST_URL = f"{API_BASE_URL}/tickets/past"
# Movidesk keeps recent tickets on /tickets; older ones require /tickets/past
PAST_THRESHOLD_DAYS = 90
PAGE_SIZE = 1000
REQUEST_DELAY_SECONDS = 0.5
REQUEST_TIMEOUT_SECONDS = 45
MAX_RETRIES = 3

# Official Movidesk origin codes (public API documentation)
ORIGIN_LABELS: Dict[str, str] = {
    "0": "First Action",
    "1": "Web - Client",
    "2": "Web - Agent",
    "3": "Email received",
    "4": "System trigger",
    "5": "Chat (online)",
    "6": "Chat (offline)",
    "7": "Email sent by system",
    "8": "Contact form",
    "9": "Web API",
    "10": "Automatic ticket opening",
    "11": "Jira integration",
    "12": "Redmine integration",
    "13": "Telephony - inbound",
    "14": "Telephony - outbound",
    "15": "Telephony - missed",
    "16": "Telephony - queue abandonment",
    "17": "Remote access",
    "18": "WhatsApp",
    "19": "Movidesk integration",
    "20": "Zenvia Chat integration",
    "21": "Telephony - unanswered",
    "22": "Facebook Messenger",
    "23": "WhatsApp Business Movidesk",
    "24": "Altu",
    "25": "WhatsApp Active",
}

# Fields requested from the API (standard ticket schema)
SELECT_FIELDS = (
    "id,protocol,subject,status,category,createdDate,resolvedIn,baseStatus,"
    "justification,origin,ownerTeam,serviceFirstLevel,serviceSecondLevel,"
    "serviceThirdLevel,serviceFull,lifetimeWorkingTime,resolvedInFirstCall,"
    "chatWaitingTime,chatTalkTime,clients,createdBy,owner,tags,cc,slaSolutionDate,"
    "slaResponseDate,lastUpdate"
)

# TicketClientApiDto does not expose personName — only businessName / personType.
EXPAND_FIELDS = (
    "clients($select=id,businessName,personType),"
    "owner($select=id,businessName),"
    "createdBy($select=id,businessName),"
    "customFieldValues($expand=items)"
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _safe_get(obj: Any, *keys: Any, default: Any = "") -> Any:
    current = obj
    for key in keys:
        if isinstance(current, dict):
            current = current.get(key, default)
        elif isinstance(current, list) and isinstance(key, int):
            current = current[key] if key < len(current) else default
        else:
            return default
        if current is None:
            return default
    return current


def _person_name(person: Any) -> str:
    if not isinstance(person, dict):
        return ""
    return (
        person.get("businessName")
        or person.get("personName")
        or person.get("id")
        or ""
    )


def _join_service_full(value: Any) -> str:
    if isinstance(value, list):
        return " > ".join(str(x) for x in value if x is not None and str(x).strip())
    return str(value or "")


def _tags_to_str(value: Any) -> str:
    if isinstance(value, list):
        return ", ".join(str(x) for x in value)
    return str(value or "")


def _custom_fields_to_json(ticket: Dict[str, Any]) -> str:
    """Keep custom fields as raw JSON (IDs/values as returned by the API)."""
    raw = ticket.get("customFieldValues")
    if not raw:
        return ""
    try:
        return json.dumps(raw, ensure_ascii=False)
    except (TypeError, ValueError):
        return str(raw)


def flatten_ticket(ticket: Dict[str, Any]) -> Dict[str, Any]:
    """
    Flatten one ticket into a flat row.

    Column names mirror Movidesk concepts. Nested objects are expanded into
    simple text fields. No business transformation is applied.
    """
    clients = ticket.get("clients") or []
    first_client = clients[0] if isinstance(clients, list) and clients else {}

    origin_raw = ticket.get("origin")
    origin_key = "" if origin_raw is None else str(origin_raw).strip()

    return {
        "id": ticket.get("id", ""),
        "protocol": ticket.get("protocol", ""),
        "subject": ticket.get("subject", ""),
        "status": ticket.get("status", ""),
        "baseStatus": ticket.get("baseStatus", ""),
        "category": ticket.get("category", ""),
        "justification": ticket.get("justification", ""),
        "origin": origin_raw if origin_raw is not None else "",
        "originLabel": ORIGIN_LABELS.get(origin_key, origin_key),
        "createdDate": ticket.get("createdDate", ""),
        "resolvedIn": ticket.get("resolvedIn", ""),
        "lastUpdate": ticket.get("lastUpdate", ""),
        "slaResponseDate": ticket.get("slaResponseDate", ""),
        "slaSolutionDate": ticket.get("slaSolutionDate", ""),
        "ownerTeam": ticket.get("ownerTeam", ""),
        "owner": _person_name(ticket.get("owner")),
        "createdBy": _person_name(ticket.get("createdBy")),
        "clientId": _safe_get(first_client, "id", default=""),
        "clientName": _person_name(first_client),
        "clientPersonType": _safe_get(first_client, "personType", default=""),
        "serviceFirstLevel": ticket.get("serviceFirstLevel", ""),
        "serviceSecondLevel": ticket.get("serviceSecondLevel", ""),
        "serviceThirdLevel": ticket.get("serviceThirdLevel", ""),
        "serviceFull": _join_service_full(ticket.get("serviceFull")),
        "lifetimeWorkingTime": ticket.get("lifetimeWorkingTime", ""),
        "resolvedInFirstCall": ticket.get("resolvedInFirstCall", ""),
        "chatWaitingTime": ticket.get("chatWaitingTime", ""),
        "chatTalkTime": ticket.get("chatTalkTime", ""),
        "tags": _tags_to_str(ticket.get("tags")),
        "cc": ticket.get("cc", ""),
        "customFieldValues": _custom_fields_to_json(ticket),
    }


# ---------------------------------------------------------------------------
# API client
# ---------------------------------------------------------------------------
class MovideskClient:
    def __init__(self, token: str) -> None:
        if not token or not str(token).strip():
            raise ValueError("API token is required.")
        self.token = str(token).strip()
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/json",
                "User-Agent": "MovideskGenericExport/1.0",
            }
        )
        self._last_request_at = 0.0

    def _throttle(self) -> None:
        elapsed = time.time() - self._last_request_at
        if elapsed < REQUEST_DELAY_SECONDS:
            time.sleep(REQUEST_DELAY_SECONDS - elapsed)

    def get(self, params: Dict[str, Any], base_url: str = TICKETS_URL) -> Any:
        self._throttle()
        query = dict(params)
        query["token"] = self.token
        qs = "&".join(
            f"{k}={quote(str(v))}" if str(k).startswith("$") else f"{k}={v}"
            for k, v in query.items()
        )
        url = f"{base_url}?{qs}"

        last_error: Optional[Exception] = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                response = self.session.get(url, timeout=REQUEST_TIMEOUT_SECONDS)
                self._last_request_at = time.time()
                response.raise_for_status()
                return response.json()
            except requests.exceptions.Timeout as exc:
                last_error = exc
                time.sleep(2 ** (attempt - 1))
            except requests.exceptions.HTTPError as exc:
                detail = ""
                try:
                    detail = exc.response.json().get("message", exc.response.text[:200])
                except Exception:
                    detail = (exc.response.text or "")[:200]
                raise RuntimeError(
                    f"HTTP {exc.response.status_code}: {detail}"
                ) from exc
            except requests.exceptions.RequestException as exc:
                last_error = exc
                time.sleep(2 ** (attempt - 1))

        raise RuntimeError(f"Request failed after {MAX_RETRIES} attempts: {last_error}")


def _should_use_past_endpoint(start_date: str, force_past: Optional[bool] = None) -> bool:
    """
    Movidesk serves recent tickets on /tickets and older ones on /tickets/past.

    When start_date is older than PAST_THRESHOLD_DAYS, the regular endpoint often
    returns only a tiny subset (or none). The main desktop extractor uses the same rule.
    """
    if force_past is True:
        return True
    if force_past is False:
        return False
    start = datetime.strptime(start_date, "%Y-%m-%d")
    return start < datetime.now() - timedelta(days=PAST_THRESHOLD_DAYS)


def fetch_tickets(
    client: MovideskClient,
    start_date: str,
    end_date: str,
    timezone_offset: str = "-03:00",
    force_past: Optional[bool] = None,
) -> List[Dict[str, Any]]:
    """
    Fetch all tickets created between start_date and end_date (inclusive).

    Dates must be YYYY-MM-DD. No status/origin filters are applied.
    """
    _validate_date(start_date)
    _validate_date(end_date)

    use_past = _should_use_past_endpoint(start_date, force_past=force_past)
    base_url = TICKETS_PAST_URL if use_past else TICKETS_URL

    all_tickets: List[Dict[str, Any]] = []
    skip = 0

    date_filter = (
        f"createdDate ge {start_date}T00:00:00{timezone_offset} "
        f"and createdDate le {end_date}T23:59:59{timezone_offset}"
    )

    print(f"Fetching tickets from {start_date} to {end_date} ...")
    print(f"  endpoint: {base_url}" + (" (historical)" if use_past else " (recent)"))

    while True:
        params = {
            "$select": SELECT_FIELDS,
            "$expand": EXPAND_FIELDS,
            "$orderby": "createdDate desc",
            "$top": PAGE_SIZE,
            "$skip": skip,
            "$filter": date_filter,
        }
        payload = client.get(params, base_url=base_url)
        # API may return a bare list or an object with items/value
        if isinstance(payload, list):
            page = payload
        elif isinstance(payload, dict):
            page = payload.get("items") or payload.get("value") or []
        else:
            page = []
        if not page:
            break

        all_tickets.extend(page)
        print(f"  page skip={skip}: +{len(page)} (total {len(all_tickets)})")

        if len(page) < PAGE_SIZE:
            break
        skip += PAGE_SIZE

    print(f"Done. {len(all_tickets)} ticket(s) returned by the API.")
    return all_tickets


def _validate_date(value: str) -> None:
    try:
        datetime.strptime(value, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError(f"Invalid date '{value}'. Use YYYY-MM-DD.") from exc


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
def tickets_to_dataframe(tickets: List[Dict[str, Any]]) -> pd.DataFrame:
    rows = [flatten_ticket(t) for t in tickets]
    return pd.DataFrame(rows)


def save_dataframe(df: pd.DataFrame, output_path: str, fmt: str) -> str:
    fmt = fmt.lower().strip()
    if fmt == "csv":
        if not output_path.lower().endswith(".csv"):
            output_path += ".csv"
        df.to_csv(output_path, index=False, encoding="utf-8-sig", sep=";")
    elif fmt in {"xlsx", "excel"}:
        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"
        df.to_excel(output_path, index=False, engine="openpyxl")
    elif fmt == "json":
        if not output_path.lower().endswith(".json"):
            output_path += ".json"
        df.to_json(output_path, orient="records", force_ascii=False, indent=2)
    else:
        raise ValueError("format must be one of: csv, xlsx, json")
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Generic Movidesk ticket exporter. "
            "Fetches tickets from the public API and saves them as a table."
        )
    )
    parser.add_argument(
        "--token",
        default=os.getenv("MOVIDESK_API_TOKEN", ""),
        help="Movidesk API token (or set MOVIDESK_API_TOKEN).",
    )
    parser.add_argument(
        "--start-date",
        required=True,
        help="Start date (YYYY-MM-DD), filter on createdDate.",
    )
    parser.add_argument(
        "--end-date",
        required=True,
        help="End date (YYYY-MM-DD), filter on createdDate.",
    )
    parser.add_argument(
        "--timezone-offset",
        default="-03:00",
        help="Timezone offset used in the OData filter (default: -03:00).",
    )
    parser.add_argument(
        "--format",
        default="xlsx",
        choices=["csv", "xlsx", "json"],
        help="Output format (default: xlsx).",
    )
    parser.add_argument(
        "--output",
        default="",
        help="Output file path (default: movidesk_tickets_YYYYMMDD_HHMMSS.<ext>).",
    )
    parser.add_argument(
        "--raw-json",
        default="",
        help="Optional path to also save the raw API tickets as JSON.",
    )
    past_group = parser.add_mutually_exclusive_group()
    past_group.add_argument(
        "--past",
        action="store_true",
        help="Force /tickets/past (historical endpoint).",
    )
    past_group.add_argument(
        "--no-past",
        action="store_true",
        help="Force /tickets (recent endpoint), even for old dates.",
    )
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    if not args.token:
        print(
            "ERROR: missing API token. Use --token or MOVIDESK_API_TOKEN.",
            file=sys.stderr,
        )
        return 1

    try:
        force_past: Optional[bool] = None
        if args.past:
            force_past = True
        elif args.no_past:
            force_past = False

        client = MovideskClient(args.token)
        tickets = fetch_tickets(
            client,
            start_date=args.start_date,
            end_date=args.end_date,
            timezone_offset=args.timezone_offset,
            force_past=force_past,
        )

        if args.raw_json:
            with open(args.raw_json, "w", encoding="utf-8") as fh:
                json.dump(tickets, fh, ensure_ascii=False, indent=2)
            print(f"Raw API JSON saved to: {args.raw_json}")

        df = tickets_to_dataframe(tickets)

        output = args.output
        if not output:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output = f"movidesk_tickets_{stamp}"

        saved = save_dataframe(df, output, args.format)
        print(f"Exported {len(df)} row(s) to: {saved}")
        print(f"Columns: {', '.join(df.columns)}")
        return 0

    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
