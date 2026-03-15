#!/usr/bin/env python3
"""
Tiller bookkeeping categorization agent.

Reads transactions from a Tiller Google Sheet, learns from past categorizations,
and fills in blank Category cells using Claude AI.

Uncertain categorizations are highlighted yellow so you can review them.
"""

import json
import os
import sys
from collections import Counter
from pathlib import Path

import anthropic
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

load_dotenv()

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

CREDENTIALS_DIR = Path(__file__).parent / "credentials"
TOKEN_PATH = CREDENTIALS_DIR / "token.json"
CLIENT_SECRET_PATH = CREDENTIALS_DIR / "client_secret.json"

CHUNK_SIZE = 50  # transactions per Claude request

# Yellow highlight for uncertain categorizations
YELLOW = {"red": 1.0, "green": 0.95, "blue": 0.4}


# ── Google Sheets helpers ──────────────────────────────────────────────────────

def get_sheets_service():
    """Authenticate with Google and return a Sheets API service."""
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CLIENT_SECRET_PATH.exists():
                print(f"\nError: {CLIENT_SECRET_PATH} not found.")
                print("Download your OAuth 2.0 credentials from Google Cloud Console")
                print("and save them as:  credentials/client_secret.json")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CLIENT_SECRET_PATH), SCOPES
            )
            creds = flow.run_local_server(port=0)

        CREDENTIALS_DIR.mkdir(exist_ok=True)
        TOKEN_PATH.write_text(creds.to_json())

    return build("sheets", "v4", credentials=creds)


def get_sheet_id(service, spreadsheet_id: str, sheet_name: str) -> int:
    """Return the numeric sheetId for a given tab name."""
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in meta["sheets"]:
        if sheet["properties"]["title"] == sheet_name:
            return sheet["properties"]["sheetId"]
    raise ValueError(f"Tab '{sheet_name}' not found in spreadsheet.")


def read_sheet(service, spreadsheet_id: str, sheet_name: str) -> list[list]:
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=sheet_name)
        .execute()
    )
    return result.get("values", [])


def write_categories(
    service,
    spreadsheet_id: str,
    sheet_name: str,
    updates: dict[int, str],  # row_index (1-based) -> category
    cat_col: int,              # 0-based column index
):
    """Batch-write categories back to the sheet."""
    if not updates:
        return
    col_letter = chr(ord("A") + cat_col)
    data = [
        {"range": f"{sheet_name}!{col_letter}{row}", "values": [[cat]]}
        for row, cat in sorted(updates.items())
    ]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": data},
    ).execute()


def highlight_cells(
    service,
    spreadsheet_id: str,
    sheet_id: int,
    rows: list[int],   # 1-based row indices
    col: int,          # 0-based column index
    color: dict,
):
    """Apply a background color to specific cells."""
    if not rows:
        return
    requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row - 1,  # 0-based
                    "endRowIndex": row,
                    "startColumnIndex": col,
                    "endColumnIndex": col + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": color
                    }
                },
                "fields": "userEnteredFormat.backgroundColor",
            }
        }
        for row in rows
    ]
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()


# ── Parsing helpers ────────────────────────────────────────────────────────────

def _col(headers: list[str], *names: str) -> int:
    """Return the index of the first matching header (case-insensitive), or -1."""
    lower = [h.lower().strip() for h in headers]
    for name in names:
        if name.lower() in lower:
            return lower.index(name.lower())
    return -1


def parse_transactions(rows: list[list]) -> tuple[list[dict], int]:
    """
    Convert raw sheet rows into a list of transaction dicts.
    Returns (transactions, category_column_index).
    """
    if not rows:
        return [], -1

    headers = rows[0]
    cat_col  = _col(headers, "Category", "category")
    desc_col = _col(headers, "Description", "description", "Full Description")
    amt_col  = _col(headers, "Amount", "amount")
    date_col = _col(headers, "Date", "date")
    acct_col = _col(headers, "Account", "account")

    if cat_col == -1:
        print(f"Error: 'Category' column not found. Headers: {headers}")
        sys.exit(1)
    if desc_col == -1:
        print(f"Error: 'Description' column not found. Headers: {headers}")
        sys.exit(1)

    max_col = max(c for c in [cat_col, desc_col, amt_col, date_col, acct_col] if c != -1)

    def cell(row, idx):
        if idx == -1 or idx >= len(row):
            return ""
        return row[idx].strip()

    transactions = []
    for i, row in enumerate(rows[1:], start=2):  # 1-indexed; row 1 is header
        padded = row + [""] * max(0, max_col + 1 - len(row))
        transactions.append({
            "row_index":   i,
            "description": cell(padded, desc_col),
            "category":    cell(padded, cat_col),
            "amount":      cell(padded, amt_col),
            "date":        cell(padded, date_col),
            "account":     cell(padded, acct_col),
        })

    return transactions, cat_col


# ── Claude categorization ──────────────────────────────────────────────────────

def build_known_patterns(transactions: list[dict]) -> dict[str, str]:
    """
    Return a dict of description.lower() -> most-common category
    built from already-categorized rows.
    """
    counts: dict[str, Counter] = {}
    for t in transactions:
        if t["category"] and t["description"]:
            key = t["description"].lower()
            counts.setdefault(key, Counter())[t["category"]] += 1
    return {desc: ctr.most_common(1)[0][0] for desc, ctr in counts.items()}


def categorize_chunk(
    chunk: list[dict],
    known_patterns: dict[str, str],
    all_categories: list[str],
    client: anthropic.Anthropic,
) -> tuple[dict[int, str], list[int]]:
    """
    Ask Claude to categorize one chunk of uncategorized transactions.
    Returns (row -> category, list of uncertain row indices).
    """

    categories_list = "\n".join(f"  - {c}" for c in sorted(set(all_categories)))
    examples = "\n".join(
        f'  "{desc}" → "{cat}"'
        for desc, cat in sorted(known_patterns.items())[:150]
    )
    txns_text = "\n".join(
        f'  Row {t["row_index"]}: description="{t["description"]}", '
        f'amount={t["amount"]}, date={t["date"]}, account="{t["account"]}"'
        for t in chunk
    )

    system_prompt = f"""You are a bookkeeping assistant that categorizes financial transactions \
for a Tiller Money Google Sheet.

Categories this person uses:
{categories_list}

Examples of how they have categorized past transactions:
{examples}

Rules:
- Prefer categories from the list above.
- If a transaction clearly belongs to a new category, create one matching \
the user's naming style (capitalization, word choice).
- Be consistent: similar merchants should get the same category.
- Never leave a category blank.
- Set "confident" to false if you are guessing — e.g. the merchant is ambiguous, \
the description is cryptic, or it could reasonably belong to multiple categories. \
Set it to true when you are certain."""

    user_message = f"""Categorize these transactions. \
Return a JSON object with a "categorizations" array.

{txns_text}"""

    response = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        thinking={"type": "adaptive"},
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
        output_config={
            "format": {
                "type": "json_schema",
                "schema": {
                    "type": "object",
                    "properties": {
                        "categorizations": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "row_index": {"type": "integer"},
                                    "category":  {"type": "string"},
                                    "confident": {"type": "boolean"},
                                },
                                "required": ["row_index", "category", "confident"],
                                "additionalProperties": False,
                            },
                        }
                    },
                    "required": ["categorizations"],
                    "additionalProperties": False,
                },
            }
        },
    )

    for block in response.content:
        if block.type == "text":
            data = json.loads(block.text)
            categories = {item["row_index"]: item["category"] for item in data["categorizations"]}
            uncertain  = [item["row_index"] for item in data["categorizations"] if not item["confident"]]
            return categories, uncertain

    return {}, []


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    spreadsheet_id = os.environ.get("TILLER_SPREADSHEET_ID")
    if not spreadsheet_id:
        print("Error: TILLER_SPREADSHEET_ID is not set in your .env file.")
        sys.exit(1)

    sheet_name = os.environ.get("TILLER_SHEET_NAME", "Transactions")

    print("Connecting to Google Sheets...")
    service = get_sheets_service()

    print(f"Reading '{sheet_name}' tab...")
    rows = read_sheet(service, spreadsheet_id, sheet_name)
    if not rows:
        print("Sheet is empty.")
        return

    transactions, cat_col = parse_transactions(rows)

    categorized   = [t for t in transactions if t["category"]]
    uncategorized = [t for t in transactions if not t["category"] and t["description"]]

    print(f"Total rows:           {len(transactions)}")
    print(f"Already categorized:  {len(categorized)}")
    print(f"Needing categories:   {len(uncategorized)}")

    if not uncategorized:
        print("\nNothing to do — all transactions are already categorized!")
        return

    known_patterns = build_known_patterns(categorized)
    all_categories = sorted({t["category"] for t in categorized if t["category"]})

    print(f"\nLearned {len(known_patterns)} description patterns.")
    print(f"Found {len(all_categories)} categories in use.")

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    chunks = [uncategorized[i:i + CHUNK_SIZE] for i in range(0, len(uncategorized), CHUNK_SIZE)]
    all_updates:   dict[int, str] = {}
    all_uncertain: list[int]      = []

    for i, chunk in enumerate(chunks, start=1):
        print(f"\nCategorizing batch {i}/{len(chunks)} ({len(chunk)} transactions)...")
        results, uncertain = categorize_chunk(chunk, known_patterns, all_categories, client)
        all_updates.update(results)
        all_uncertain.extend(uncertain)
        print(f"  → {len(results)} categories assigned, {len(uncertain)} flagged as uncertain.")

    print(f"\nWriting {len(all_updates)} categories back to the sheet...")
    write_categories(service, spreadsheet_id, sheet_name, all_updates, cat_col)

    if all_uncertain:
        print(f"Highlighting {len(all_uncertain)} uncertain cells yellow...")
        sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)
        highlight_cells(service, spreadsheet_id, sheet_id, all_uncertain, cat_col, YELLOW)

    print("\nDone!")
    sample_map = {t["row_index"]: t["description"] for t in uncategorized}

    confident_rows   = [r for r in all_updates if r not in all_uncertain]
    uncertain_sample = [r for r in all_uncertain][:10]

    if uncertain_sample:
        print(f"\n⚠ Uncertain (highlighted yellow — please review):")
        for row in uncertain_sample:
            desc = sample_map.get(row, "?")
            cat  = all_updates.get(row, "?")
            print(f"  Row {row:>4}: \"{desc}\"  →  {cat}")

    print(f"\n✓ Confident ({len(confident_rows)} total), sample:")
    for row in list(all_updates.keys())[:5]:
        if row not in all_uncertain:
            desc = sample_map.get(row, "?")
            cat  = all_updates.get(row, "?")
            print(f"  Row {row:>4}: \"{desc}\"  →  {cat}")


if __name__ == "__main__":
    main()
