"""
Daily Amazon BSR updater for Google Sheets.

Reads ASINs from column A of the BSR Tracking sheet, fetches the current
Best Seller Rank from Amazon, and writes results to column E (Current BSR).

Required environment variables / GitHub Secrets:
  GOOGLE_SHEETS_CREDENTIALS  - Service account JSON (as a string)
  SPREADSHEET_ID             - Google Sheets document ID
  SHEET_NAME                 - Sheet tab name (default: "BSR Tracking")
  AMAZON_MARKETPLACE         - e.g. "com", "co.uk", "de" (default: "com")
"""

import json
import os
import random
import time
import re
import logging
from datetime import datetime

import gspread
import requests
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

# ── Configuration ────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
SHEET_NAME = os.environ.get("SHEET_NAME", "BSR Tracking")
MARKETPLACE = os.environ.get("AMAZON_MARKETPLACE", "com")

ASIN_COL = 0       # Column A (0-indexed)
CURRENT_BSR_COL = 4  # Column E (0-indexed)
HEADER_ROW = 0     # Row 1 is the header (0-indexed)

# Delay range between requests to avoid rate-limiting (seconds)
MIN_DELAY = 3
MAX_DELAY = 7

USER_AGENTS = [
    (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/605.1.15 (KHTML, like Gecko) "
        "Version/17.4.1 Safari/605.1.15"
    ),
    (
        "Mozilla/5.0 (X11; Linux x86_64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/123.0.0.0 Safari/537.36"
    ),
]

# ── Google Sheets helpers ────────────────────────────────────────────────────

def get_sheet() -> gspread.Worksheet:
    raw = os.environ["GOOGLE_SHEETS_CREDENTIALS"]
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    return spreadsheet.worksheet(SHEET_NAME)


def read_asins(sheet: gspread.Worksheet) -> list[tuple[int, str]]:
    """Return list of (row_index_1based, asin) skipping the header."""
    all_values = sheet.col_values(ASIN_COL + 1)  # gspread is 1-indexed
    results = []
    for i, val in enumerate(all_values):
        if i == HEADER_ROW:
            continue
        asin = val.strip()
        if asin:
            results.append((i + 1, asin))  # 1-based row number
    return results


def write_bsr(sheet: gspread.Worksheet, row: int, value: str) -> None:
    col_letter = chr(ord("A") + CURRENT_BSR_COL)  # 'E'
    cell = f"{col_letter}{row}"
    sheet.update_acell(cell, value)


# ── Amazon scraping helpers ──────────────────────────────────────────────────

def build_url(asin: str) -> str:
    return f"https://www.amazon.{MARKETPLACE}/dp/{asin}"


def fetch_page(asin: str, session: requests.Session) -> BeautifulSoup | None:
    url = build_url(asin)
    headers = {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": (
            "text/html,application/xhtml+xml,application/xml;"
            "q=0.9,image/webp,*/*;q=0.8"
        ),
        "Accept-Encoding": "gzip, deflate, br",
        "DNT": "1",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    }
    try:
        resp = session.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        return BeautifulSoup(resp.text, "html.parser")
    except requests.RequestException as exc:
        log.warning("Failed to fetch %s: %s", url, exc)
        return None


def parse_bsr(soup: BeautifulSoup) -> str | None:
    """
    Try multiple Amazon HTML patterns to extract BSR text.
    Returns a cleaned string like "#9,679 in Health & Personal Care" or None.
    """
    # Pattern 1: detail bullets list (most common current layout)
    for li in soup.select("li.a-list-item"):
        text = li.get_text(" ", strip=True)
        if "Best Sellers Rank" in text:
            return _clean_bsr_text(text)

    # Pattern 2: product details table rows
    for th in soup.find_all("th"):
        if "Best Sellers Rank" in th.get_text():
            td = th.find_next_sibling("td")
            if td:
                return _clean_bsr_text(td.get_text(" ", strip=True))

    # Pattern 3: legacy <li id="SalesRank">
    sales_rank = soup.find("li", id="SalesRank")
    if sales_rank:
        return _clean_bsr_text(sales_rank.get_text(" ", strip=True))

    # Pattern 4: span with rank text inside detail sections
    for span in soup.select("#detailBulletsWrapper_feature_div span.a-list-item"):
        text = span.get_text(" ", strip=True)
        if "Best Sellers Rank" in text:
            return _clean_bsr_text(text)

    return None


def _clean_bsr_text(raw: str) -> str:
    # Remove the label prefix
    text = re.sub(r"Best Sellers? Rank[:\s]*", "", raw, flags=re.IGNORECASE)
    # Collapse whitespace / newlines
    text = re.sub(r"\s+", " ", text).strip()
    # Truncate to the first meaningful rank entry (before any parenthetical)
    # Keep up to ~100 chars to include sub-category if present
    return text[:200]


# ── Main loop ────────────────────────────────────────────────────────────────

def main() -> None:
    log.info("Starting BSR update — %s", datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"))

    sheet = get_sheet()
    asins = read_asins(sheet)
    log.info("Found %d ASINs to process", len(asins))

    session = requests.Session()
    updated = 0
    failed = 0

    for row, asin in asins:
        log.info("Row %d | ASIN %s", row, asin)
        soup = fetch_page(asin, session)

        if soup is None:
            log.warning("  Skipping — could not fetch page")
            write_bsr(sheet, row, "ERROR: fetch failed")
            failed += 1
        else:
            bsr = parse_bsr(soup)
            if bsr:
                log.info("  BSR: %s", bsr)
                write_bsr(sheet, row, bsr)
                updated += 1
            else:
                log.warning("  BSR not found on page")
                write_bsr(sheet, row, "ERROR: BSR not found")
                failed += 1

        # Polite delay between requests
        delay = random.uniform(MIN_DELAY, MAX_DELAY)
        log.debug("  Sleeping %.1fs", delay)
        time.sleep(delay)

    log.info("Done. Updated: %d | Failed: %d", updated, failed)


if __name__ == "__main__":
    main()
