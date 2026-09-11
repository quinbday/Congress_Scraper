"""
Congressional Trades Scraper — Python port of the original R pipeline.

Pipeline
  1. Scrape trade listings from capitoltrades.com, page by page.
  2. Validate tickers and pull sector/industry/price data via yfinance
  3. Scrape House/Senate committee rosters and match each politician to
     the committees they sit on.
  4. Write everything to a formatted Excel workbook.
"""

import glob
import logging
import os
import random
import re
import threading
import time
from datetime import date, timedelta
from email.utils import parsedate_to_datetime
from urllib.parse import urlparse

import pandas as pd
import requests
import yfinance as yf
from lxml import html
from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# yfinance logs "No data found" / "possibly delisted" straight to stderr on
# every failed lookup, even when we've already caught the exception — this
# just quiets that, it doesn't change behavior.
logging.getLogger("yfinance").setLevel(logging.CRITICAL)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
BASE_URL = "https://www.capitoltrades.com/trades"
BASE_HOST = "https://www.capitoltrades.com"

MAX_PAGES =  3085         # Set to ~10 for testing
REQUEST_DELAY = 1.5       # seconds between capitoltrades.com page fetches
MIN_REQUEST_INTERVAL = 1.0  # minimum interval for every HTML request per host
YAHOO_REQUEST_INTERVAL = 1.0  # minimum interval between yfinance calls
CHECKPOINT_EVERY = 25     # pages between checkpoint saves
INCREMENTAL_PATH = "incremental_latest.xlsx"
OUTPUT_PATH = "C:\\Users\\quinb\\OneDrive\\Desktop\\Python Code\\Congress_Trading.xlsx"

SIZE_BUCKETS = [
    "<1K", "1K\u201315K", "15K\u201350K", "50K\u2013100K", "100K\u2013250K",
    "250K\u2013500K", "500K\u20131M", "1M\u20135M", "5M\u201325M", "25M\u201350M",
]
SIZE_COLORS = [
    "228B22", "00FF00", "00FA9A", "00FFFF", "FF7F50",
    "CD5C5C", "FF0000", "8B0000", "FFD700", "FFFF00",
]

HOUSE_COMMITTEE_CODES = [
    "AG00", "AP00", "AS00", "BA00", "BU00", "ED00", "FA00", "GO00",
    "HA00", "HI00", "HM00", "HS00", "IF00", "IG00", "II00", "IN00",
    "JU00", "PW00", "RU00", "SM00", "SY00", "TR00", "VR00", "WM00",
]
SENATE_COMMITTEE_CODES = [
    "SSAP", "SSAF", "SSAS", "SSBK", "SSBU", "SSCM", "SSEG", "SSEV",
    "SSFI", "SSFR", "SSHR", "SSGA", "SSRA", "SSJU", "SSSB", "SSVA",
    "SLET", "SLIA", "JSPE", "JSTX", "JSLC", "JSEC", "SLIN", "SPAG",
]
NAME_SUFFIXES = {"jr.", "jr", "sr.", "sr", "ii", "iii", "iv", "v"}

session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0"})
_request_times = {}
_request_lock = threading.Lock()
_last_yahoo_request = 0.0


# ---------------------------------------------------------------------------
# Networking
# ---------------------------------------------------------------------------
def safe_get_tree(url, retries=6, backoff=1.0, timeout=15):
    """Fetch a URL and return a parsed lxml tree, retrying with backoff.

    Uses the shared module-level `session` so TCP connections are pooled
    and reused instead of opened one at a time.

    A 429 (Too Many Requests) gets special handling: it's not a transient
    network blip, it's the server explicitly telling us to slow down, so it
    gets a much longer wait (honoring a Retry-After header if the server
    sends one) and more attempts than a generic error -- 3 quick retries
    with ~1-4s backoff was giving up mid-throttle-window and permanently
    losing that page's data.
    """
    last_err = None
    for attempt in range(retries):
        try:
            host = urlparse(url).netloc
            with _request_lock:
                elapsed = time.monotonic() - _request_times.get(host, 0.0)
                if elapsed < MIN_REQUEST_INTERVAL:
                    time.sleep(MIN_REQUEST_INTERVAL - elapsed)
                _request_times[host] = time.monotonic()
            resp = session.get(url, timeout=timeout)
            if resp.status_code == 429:
                retry_after = resp.headers.get("Retry-After")
                try:
                    wait = float(retry_after) if retry_after else max(backoff, 15)
                except (TypeError, ValueError):
                    try:
                        retry_at = parsedate_to_datetime(retry_after)
                        wait = max(0, retry_at.timestamp() - time.time())
                    except (TypeError, ValueError, OverflowError):
                        wait = max(backoff, 15)
                wait = min(max(wait, 1), 300)
                if attempt < retries - 1:
                    print(f"  429 from {url} -- waiting {wait:.0f}s (attempt {attempt + 1}/{retries})")
                    time.sleep(wait)
                    backoff = min(wait * 2, 300)
                    continue
                resp.raise_for_status()  # out of attempts -- raise normally
            resp.raise_for_status()
            return html.fromstring(resp.content)
        except requests.RequestException as e:
            last_err = e
            if attempt < retries - 1:
                time.sleep(min(backoff, 300))
                backoff = min(backoff * 2, 300)
    raise last_err


def wait_for_yahoo():
    """Keep sequential yfinance calls from becoming a request burst."""
    global _last_yahoo_request
    with _request_lock:
        elapsed = time.monotonic() - _last_yahoo_request
        if elapsed < YAHOO_REQUEST_INTERVAL:
            time.sleep(YAHOO_REQUEST_INTERVAL - elapsed)
        _last_yahoo_request = time.monotonic()


def yahoo_call(call):
    """Run one yfinance operation after applying the global Yahoo throttle."""
    wait_for_yahoo()
    return call()


def xpath_text(tree, xpath, trim=True):
    """Equivalent of rvest's `html_elements(xpath=...) %>% html_text()`."""
    out = [el.text_content() for el in tree.xpath(xpath)]
    return [t.strip() for t in out] if trim else out


# ---------------------------------------------------------------------------
# Yahoo Finance helpers (via yfinance — replaces getQuote / HTML scrape /
# getSymbols from the R version)
# ---------------------------------------------------------------------------
_price_cache = {}      # ticker -> last price (or None), cached for this run
_industry_cache = {}   # ticker -> (industry, sector), cached for this run
_hist_cache = {}        # (ticker, target_date, snap) -> price, cached for this run


def normalize_ticker(ticker):
    """capitoltrades.com scrapes share-class tickers with a slash (BRK/B),
    but Yahoo Finance expects a hyphen (BRK-B). Also uppercases and trims whitespace."""
    return ticker.strip().upper().replace("/", "-")


def get_last_price(ticker):
    """Live/last quote for a ticker, or None if unavailable. Cached per
    ticker for this run so a bad/delisted symbol only gets queried once
    instead of failing repeatedly across every row that references it."""
    ticker = normalize_ticker(ticker)
    if ticker in _price_cache:
        return _price_cache[ticker]
    try:
        fi = yahoo_call(lambda: yf.Ticker(ticker).fast_info)
        last = fi.get("last_price") or fi.get("lastPrice")
        price = float(last) if last and last > 0 else None
    except Exception:
        price = None
    _price_cache[ticker] = price
    return price


def get_industry_sector(ticker):
    """Returns (industry, sector), each defaulting to 'ETF' if unavailable. 
    Cached per ticker for this run."""
    ticker = normalize_ticker(ticker)
    if ticker in _industry_cache:
        return _industry_cache[ticker]
    try:
        info = yahoo_call(lambda: yf.Ticker(ticker).info)
        industry = info.get("industry") or "ETF"
        sector = info.get("sector") or "ETF"
    except Exception:
        industry, sector = "ETF", "ETF"
    _industry_cache[ticker] = (industry, sector)
    return industry, sector


def get_hist_price(ticker, target_date, snap="before", max_tries=4):
    """Historical close nearest target_date, snapping before/after it."""
    ticker = normalize_ticker(ticker)
    today = date.today()
    if target_date > today:
        return get_last_price(ticker)

    cache_key = (ticker, target_date, snap)
    if cache_key in _hist_cache:
        return _hist_cache[cache_key]

    fetch_from = target_date - timedelta(days=14)
    fetch_to = min(target_date + timedelta(days=3), today)

    result = float("nan")
    wait = 1
    for attempt in range(max_tries):
        try:
            hist = yahoo_call(
                lambda: yf.Ticker(ticker).history(
                    start=fetch_from, end=fetch_to + timedelta(days=1)
                )
            )
            failed = False
        except Exception:
            hist = None
            failed = True

        if not failed:
            # A successful call that legitimately came back empty (no
            # trading data in this window) won't change on retry with the
            # same dates. Retrying just wastes requests and helps trigger
            # rate limiting. Only network-level failures are worth retrying.
            if hist is not None and not hist.empty:
                idx_dates = [d.date() for d in hist.index]
                if snap == "before":
                    candidates = [d for d in idx_dates if d <= target_date]
                    chosen = max(candidates) if candidates else None
                else:
                    candidates = [d for d in idx_dates if d >= target_date]
                    chosen = min(candidates) if candidates else None

                if chosen is not None:
                    row = hist.loc[[d.date() == chosen for d in hist.index]]
                    val = float(row["Close"].iloc[0])
                    if val > 0:
                        result = val
            break  # successful call either way
        if attempt < max_tries - 1:
            time.sleep(wait)
            wait *= 2

    if result != result:  # NaN check without importing math
        result = get_last_price(ticker)  # last resort
        result = result if result is not None else float("nan")

    _hist_cache[cache_key] = result
    return result


# ---------------------------------------------------------------------------
# Committee scraping
# ---------------------------------------------------------------------------
def sanitize_sheet_name(name):
    name = re.sub(r"[\[\]\*\?:/\\]", "", name)
    name = name.replace(",", "")
    return name.strip()


def scrape_house_committee(code):
    url = f"https://clerk.house.gov/committees/{code}"
    try:
        tree = safe_get_tree(url)
    except Exception:
        return {"title": code, "members": []}

    title_raw = " ".join(xpath_text(tree, "//*[@class='library-h1-wrap']", trim=False))
    title_clean = re.sub(r"\s+", " ", title_raw).replace("Share", "").strip()

    members_raw = xpath_text(
        tree,
        "//*[@class='library-lists library-list_ol_v0 committee-members-list']/li",
    )

    members = []
    for x in members_raw:
        x = re.sub(r"\s+", " ", x).strip()
        x = re.sub(r"^\d+\.\s*", "", x)                 # drop leading "1. "
        x = re.sub(r",\s*[A-Z]{2}.*$", "", x)            # drop ", ST ..." tail
        parts = [p for p in x.split(" ") if p]
        while parts and re.match(r"^[A-Z]+$", parts[-1]):  # drop trailing CAPS labels
            parts.pop()
        members.append(" ".join(parts))

    return {"title": title_clean, "members": members}


def scrape_senate_committee(code):
    url = (
        "https://www.senate.gov/general/committee_membership/"
        f"committee_memberships_{code}.htm"
    )
    try:
        tree = safe_get_tree(url)
    except Exception:
        return {"title": code, "members": []}

    titles = xpath_text(tree, "//*[@class='contenttitle']")
    title = titles[0] if titles else code

    members_raw = xpath_text(tree, "//*[@class='contenttext']")
    entry = next((m for m in members_raw if "Majority Members" in m), "")
    extracted = re.findall(r"[A-Z][a-zA-Z]+, [A-Z][a-zA-Z. ]+(?=\s*\()", entry)

    members = []
    for n in extracted:
        m = re.match(r"^([^,]+),\s*(.+)$", n)
        members.append(f"{m.group(2)} {m.group(1)}".strip() if m else n.strip())

    return {"title": title, "members": members}


def scrape_committee(code):
    return (
        scrape_senate_committee(code)
        if code in SENATE_COMMITTEE_CODES
        else scrape_house_committee(code)
    )


def match_committees(politician, legislation_values, committee_cache):
    """Returns a comma-separated string of committee names this politician
    appears to sit on, based on fuzzy first/last-name matching against
    scraped committee rosters."""
    parts = politician.split(" ")
    if parts and parts[-1].lower().rstrip(".") in NAME_SUFFIXES:
        parts = parts[:-1]
    if not parts:
        return ""

    first_name, last_name = parts[0], parts[-1]
    pattern_forward = re.compile(re.escape(first_name) + ".*" + re.escape(last_name), re.I)
    pattern_reverse = re.compile(re.escape(last_name) + ",.*" + re.escape(first_name), re.I)

    if "Senate" in legislation_values:
        codes = SENATE_COMMITTEE_CODES
    elif "House" in legislation_values:
        codes = HOUSE_COMMITTEE_CODES
    else:
        codes = HOUSE_COMMITTEE_CODES + SENATE_COMMITTEE_CODES

    matched = []
    for code in codes:
        info = committee_cache.get(code)
        if info is None:
            info = scrape_committee(code)
            committee_cache[code] = info
            time.sleep(0.1)  # avoid loading/scraping the committee pages too quickly
        body_text = " ".join(info["members"])
        if pattern_forward.search(body_text) or pattern_reverse.search(body_text):
            matched.append(info["title"])

    return ", ".join(matched)


# ---------------------------------------------------------------------------
# Trade-page scraping
# ---------------------------------------------------------------------------
def parse_party_field(text):
    """Mirrors the triplet that splits one 'PartyLegislationST'
    string into Party / Legislation / State,  falls back to the original
    text on no match."""
    m = re.search(r"(Democrat|Republican)", text)
    party = m.group(1) if m else text
    m = re.search(r"(House|Senate)", text)
    legislation = m.group(1) if m else text
    m = re.search(r"(House|Senate)([A-Z]{2})$", text)
    state = m.group(2) if m else text
    return party, legislation, state


def robust_parse_date(date_str):
    # capitoltrades.com renders September as "Sept" (4 letters) instead of
    # the standard 3-letter abbreviation used everywhere else, which breaks
    # a plain %d %b %Y parse for every September trade -- this was silently
    # producing NaN Price_Purchased/Price_Current for those rows. Fall back
    # to fixing just that case before giving up.
    if pd.isna(date_str):
        return None
    try:
        parsed = pd.to_datetime(date_str, format="%d %b %Y")
    except (ValueError, TypeError):
        try:
            parsed = pd.to_datetime(str(date_str).replace("Sept", "Sep"), format="%d %b %Y")
        except (ValueError, TypeError):
            return None
    return parsed.date() if pd.notna(parsed) else None


def scrape_trade_page(page_num):
    url = BASE_URL if page_num == 1 else f"{BASE_URL}?page={page_num}"
    tree = safe_get_tree(url)

    politician_page = xpath_text(
        tree,
        "//*[@class = 'politician-name overflow-hidden overflow-ellipsis text-[13px] text-foreground']",
    )
    party_page = xpath_text(
        tree,
        "//*[@class = 'politician-info mt-1 text-size-2 font-medium leading-none text-txt-dimmer']",
    )
    links = [
        el.get("href")
        for el in tree.xpath(
            "//a[contains(@class, 'hover:no-underline') and contains(@class, 'text-txt-interactive')]"
        )
    ]

    sep_page = [parse_party_field(p) for p in party_page]  # list of (party, legis, state)

    asset_page = [
        t.split(":")[0].strip()
        for t in xpath_text(tree, "//*[@class = 'q-field issuer-ticker']")
    ]

    date_traded_page = xpath_text(tree, "//*[@class = 'text-size-3 font-medium']", trim=False)
    year_traded = xpath_text(tree, "//*[@class = 'text-size-2 text-txt-dimmer']")

    filed_after_raw = xpath_text(
        tree,
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' reporting-gap-tier--1 ')"
        " or contains(concat(' ', normalize-space(@class), ' '), ' reporting-gap-tier--2 ')"
        " or contains(concat(' ', normalize-space(@class), ' '), ' reporting-gap-tier--3 ')"
        " or contains(concat(' ', normalize-space(@class), ' '), ' reporting-gap-tier--4 ')]",
    )
    filed_after = []
    for t in filed_after_raw:
        digits = re.sub(r"[^0-9.]", "", t)
        filed_after.append(float(digits) if digits else float("nan"))

    # year_traded comes in groups of 3: [disclosure_year, trade_year, _]
    # (positions confirmed against Filed_After -- see the note further down)
    year_disclosed_raw = year_traded[0::3]
    year_traded_raw = year_traded[1::3]
    valid_year_mask = [
        not re.match(r"(?i)^(today|yesterday)$", v) for v in year_disclosed_raw
    ]

    type_of_trade_page = xpath_text(
        tree,
        "//*[@class = 'align-middle [&:has([role=checkbox])]:pr-0 p-0']",
    )
    type_of_trade_cur = [
        t for t in type_of_trade_page if t in ("buy", "sell", "receive", "exchange")
    ]

    size_page = xpath_text(
        tree,
        "//*[@class = 'mt-1 text-size-2 text-txt-dimmer hover:text-foreground']",
    )
    filtered_sizes = []
    size_counter = 0
    for val in size_page:
        if val in SIZE_BUCKETS:
            idx = size_counter
            size_counter += 1
            if idx < len(valid_year_mask) and valid_year_mask[idx]:
                filtered_sizes.append(val)

    # date_traded_page alternates: [disclosure_date, trade_date, disclosure_date, ...]
    # NOTE: verified against Filed_After (days between trade and disclosure,
    # scraped independently from a separate part of the page) that the FIRST
    # date in each pair is the DISCLOSURE date and the SECOND is the actual
    # TRADE date -- the reverse of what this code originally assumed.
    # corr(Date - Date_Disclosed, Filed_After) = 0.9998 once assigned this way,
    # vs. near-zero (and often negative) the other way around.
    filtered_dates, filtered_disclosed = [], []
    date_counter = 0
    for j, text in enumerate(date_traded_page):
        clean = text.replace("\n", " ").strip()
        if j % 2 == 0:  # disclosure date (position 0 in each pair)
            idx = date_counter
            date_counter += 1
            if idx >= len(valid_year_mask) or not valid_year_mask[idx]:
                continue
            filtered_disclosed.append(f"{clean} {year_disclosed_raw[idx]}")
        else:  # trade date (position 1 in each pair), paired with the most recent disclosure date
            if date_counter == 0:
                continue
            idx = date_counter - 1
            if idx >= len(valid_year_mask) or not valid_year_mask[idx]:
                continue
            trade_year = (
                year_traded_raw[idx] if idx < len(year_traded_raw) else year_disclosed_raw[idx]
            )
            filtered_dates.append(f"{clean} {trade_year}")

    # ---- Validate tickers via yfinance, collecting only real ones ----
    valid_assets, valid_indexes, valid_prices = [], [], []
    for k, raw_ticker in enumerate(asset_page):
        if not raw_ticker or raw_ticker in ("N/A", "NA"):
            continue
        price = get_last_price(raw_ticker)
        if price is not None:
            valid_assets.append(raw_ticker)
            valid_indexes.append(k)
            valid_prices.append(price)
        time.sleep(0.15 + random.uniform(0, 0.15))  # stay within Yahoo Finance rate limits

    final_valid_indexes, final_valid_assets, final_valid_prices = [], [], []
    industry_list, sector_list = [], []
    for asset, orig_index in zip(valid_assets, valid_indexes):
        industry, sector = get_industry_sector(asset)
        final_valid_indexes.append(orig_index)
        final_valid_assets.append(asset)
        final_valid_prices.append(valid_prices[valid_indexes.index(orig_index)])
        industry_list.append(industry)
        sector_list.append(sector)

    if not final_valid_indexes:
        return None

    def take(seq):
        return [seq[i] for i in final_valid_indexes if i < len(seq)]

    n = len(final_valid_assets)
    cand = {
        "Politician": take(politician_page),
        "Legislation": [sep_page[i][1] for i in final_valid_indexes if i < len(sep_page)],
        "Party": [sep_page[i][0] for i in final_valid_indexes if i < len(sep_page)],
        "Trade_Type": take(type_of_trade_cur),
        "Date": take(filtered_dates),
        "Date_Disclosed": take(filtered_disclosed),
        "Amount": take(filtered_sizes),
        "State_Politician": [sep_page[i][2] for i in final_valid_indexes if i < len(sep_page)],
        "Filed_After": take(filed_after),
    }
    if not all(len(v) == n for v in cand.values()):
        # Lengths disagree across the per-trade fields on this page — same
        # situation the R version handles by dropping Today/Yesterday rows
        # and, failing that, skipping the page entirely.
        return None

    # ---- Per-trade politician/company state lookup ----
    state_company = []
    for i in final_valid_indexes:
        href = links[i] if i < len(links) else None
        if not href:
            state_company.append(None)
            continue
        time.sleep(0.1)
        try:
            pol_tree = safe_get_tree(BASE_HOST + href)
        except Exception:
            state_company.append(None)
            continue
        state_vals = [
            s
            for s in xpath_text(pol_tree, "//*[@class='text-size-4 font-medium text-foreground']")
            if not re.search(r"[0-9]", s)
        ]
        if len(state_vals) >= 2 and state_vals[1] != "N/A":
            state_company.append(state_vals[1])
        elif state_vals:
            state_company.append(state_vals[0])
        else:
            state_company.append(None)

    page_df = pd.DataFrame(
        {
            "Politician": cand["Politician"],
            "Legislation": cand["Legislation"],
            "Party": cand["Party"],
            "Trade_Type": cand["Trade_Type"],
            "Stock": final_valid_assets,
            "Industry": industry_list,
            "Sector": sector_list,
            "Date": cand["Date"],
            "Date_Disclosed": cand["Date_Disclosed"],
            "Amount": cand["Amount"],
            "State_Company": state_company,
            "State_Politician": cand["State_Politician"],
            "Filed_After": cand["Filed_After"],
        }
    )

    page_df = page_df[
        page_df["Date"].notna()
        & page_df["Amount"].notna()
        & (page_df["Date"].astype(str).str.strip() != "")
        & (page_df["Amount"].astype(str).str.strip() != "")
        & (page_df["Date"].astype(str).str.strip().str.lower() != "n/a")
        & (page_df["Amount"].astype(str).str.strip().str.lower() != "n/a")
    ]
    if page_df.empty:
        return None

    # ---- Historical prices ----
    purchased_prices, current_prices = [], []
    for _, row in page_df.iterrows():
        trade_type = str(row["Trade_Type"]).lower()
        trade_date = robust_parse_date(row["Date"])
        disc_date = robust_parse_date(row["Date_Disclosed"])

        purchased = get_hist_price(row["Stock"], disc_date, "before") if disc_date else float("nan")
        time.sleep(0.2)

        if disc_date and trade_type in ("sell", "buy"):
            offset = 90 if trade_type == "sell" else 180
            target = disc_date + timedelta(days=offset)
            current = get_hist_price(row["Stock"], target, "after")
        else:
            current = get_last_price(row["Stock"]) or float("nan")
        time.sleep(0.2)

        purchased_prices.append(purchased)
        current_prices.append(current)

    page_df = page_df.copy()
    page_df["Price_Purchased"] = purchased_prices
    page_df["Price_Current"] = current_prices
    return page_df


# ---------------------------------------------------------------------------
# Output workbook
# ---------------------------------------------------------------------------
def write_final_workbook(df, path):
    df.to_excel(path, sheet_name="Trades", index=False)
    wb = load_workbook(path)
    ws = wb["Trades"]
    headers = [c.value for c in ws[1]]
    n_rows = ws.max_row

    if "Amount" in headers:
        col_letter = get_column_letter(headers.index("Amount") + 1)
        rng = f"{col_letter}2:{col_letter}{n_rows}"
        for bucket, color in zip(SIZE_BUCKETS, SIZE_COLORS):
            fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            ws.conditional_formatting.add(
                rng, CellIsRule(operator="equal", formula=[f'"{bucket}"'], fill=fill)
            )

    if "Trade_Type" in headers:
        col_letter = get_column_letter(headers.index("Trade_Type") + 1)
        rng = f"{col_letter}2:{col_letter}{n_rows}"
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=['"buy"'], font=Font(color="008000"))
        )
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=['"sell"'], font=Font(color="FF0000"))
        )

    wb.save(path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main(max_pages=MAX_PAGES):
    frames = []
    checkpoint_frames = []  # rows scraped since the last checkpoint only
    stock_data = pd.DataFrame()

    for i in range(1, max_pages + 1):
        time.sleep(REQUEST_DELAY)
        try:
            page_df = scrape_trade_page(i)
        except Exception as e:
            print(f"Page {i} failed, skipping: {e}")
            continue

        if page_df is None or page_df.empty:
            continue

        frames.append(page_df)
        checkpoint_frames.append(page_df)
        stock_data = pd.concat(frames, ignore_index=True)

        stock_data.to_excel(INCREMENTAL_PATH, sheet_name="Trades", index=False)

        if i % CHECKPOINT_EVERY == 0 and checkpoint_frames:
            checkpoint_df = pd.concat(checkpoint_frames, ignore_index=True)
            print(f"Checkpoint at page {i}. New rows this checkpoint: {len(checkpoint_df)}")
            checkpoint_df.to_excel(f"checkpoint_page_{i}.xlsx", sheet_name="Trades", index=False)
            checkpoint_frames = []  # reset — next checkpoint file starts fresh

    # ---- Recover from checkpoint files, if any are left over from a
    # crashed prior run -- done BEFORE committee matching now (previously
    # this ran after and overwrote stock_data with a version that had no
    # Committees column at all -- a real bug, not just a style choice) ----
    checkpoint_files = sorted(glob.glob("checkpoint_page_*.xlsx"))
    if checkpoint_files:
        print(f"Found {len(checkpoint_files)} checkpoint files. Combining...")
        stock_data = pd.concat(
            [pd.read_excel(f, sheet_name="Trades") for f in checkpoint_files],
            ignore_index=True,
        )
        print(f"Total rows after combining: {len(stock_data)}")
    else:
        print("No checkpoint files found, using stock_data from memory.")

    # A long-running scrape can re-encounter the same trade under a
    # different page number if the live site's pagination shifts underneath
    # it mid-run (new trades push older ones onto later pages) -- final
    # safety-net dedup before writing out.
    if not stock_data.empty:
        before = len(stock_data)
        stock_data = stock_data.drop_duplicates().reset_index(drop=True)
        if before != len(stock_data):
            print(f"Removed {before - len(stock_data)} duplicate rows (pagination drift during the run)")

    # ---- Committee mapping ----
    committee_cache = {}
    mapping = {}
    for pol in stock_data["Politician"].unique() if not stock_data.empty else []:
        leg_values = set(stock_data.loc[stock_data["Politician"] == pol, "Legislation"])
        mapping[pol] = match_committees(pol, leg_values, committee_cache)

    if not stock_data.empty:
        stock_data["Committees"] = stock_data["Politician"].map(mapping)
        stock_data["Committees"] = stock_data["Committees"].fillna("").replace("", "None")

    write_final_workbook(stock_data, OUTPUT_PATH)
    print(f"{OUTPUT_PATH} saved successfully.")

    for f in checkpoint_files:
        os.remove(f)
    if checkpoint_files:
        print("Checkpoint files removed.")


if __name__ == "__main__":
    main()