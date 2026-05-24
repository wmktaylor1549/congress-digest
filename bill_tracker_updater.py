import os
import json
import requests
import logging
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

API_KEY           = os.getenv("CONGRESS_API_KEY", "")
EXCEL_PATH        = Path("5.26.26 Climate Tracker .xlsx")
STATE_FILE        = Path("bill_tracker_state.json")
DAYS_TO_LOOK_BACK = 7
API_BASE          = "https://api.congress.gov/v3"

KEYWORDS = [
    "climate", "energy", "clean air", "clean water", "emissions",
    "renewable", "carbon", "fossil fuel", "electric", "solar", "wind",
    "nuclear", "pipeline", "greenhouse", "EPA", "environmental",
    "petroleum", "natural gas", "geothermal", "hydropower", "water",
]

HOUSE_SHEET  = "House"
SENATE_SHEET = "Senate"

BILL_TYPE_MAP = {
    "HR": "H.R.", "HRES": "H.RES.", "HJRES": "H.J.RES.", "HCONRES": "H.CON.RES.",
    "S": "S.", "SRES": "S.RES.", "SJRES": "S.J.RES.", "SCONRES": "S.CON.RES.",
}

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)


def load_state():
    if STATE_FILE.exists():
        with open(STATE_FILE) as f:
            return json.load(f)
    return {"seen_ids": [], "last_run": None}


def save_state(state):
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)


def fetch_bills(from_date, congress=119):
    url = f"{API_BASE}/bill/{congress}"
    params = {
        "api_key": API_KEY,
        "fromDateTime": f"{from_date}T00:00:00Z",
        "sort": "updateDate+desc",
        "limit": 250,
        "format": "json",
    }
    bills  = []
    offset = 0
    while True:
        params["offset"] = offset
        try:
            resp = requests.get(url, params=params, timeout=15)
            resp.raise_for_status()
            data = resp.json()
        except requests.RequestException as e:
            log.error(f"API request failed: {e}")
            break
        batch = data.get("bills", [])
        if not batch:
            break
        bills.extend(batch)
        pagination = data.get("pagination", {})
        if len(bills) >= pagination.get("count", 0):
            break
        offset += 250
    log.info(f"Fetched {len(bills)} bills from Congress.gov")
    return bills


def fetch_bill_detail(bill_type, bill_number, congress=119):
    url = f"{API_BASE}/bill/{congress}/{bill_type.lower()}/{bill_number}"
    try:
        resp = requests.get(url, params={"api_key": API_KEY, "format": "json"}, timeout=15)
        resp.raise_for_status()
        return resp.json().get("bill", {})
    except requests.RequestException as e:
        log.warning(f"Could not fetch detail for {bill_type}{bill_number}: {e}")
        return {}


def fetch_cosponsors(bill_type, bill_number, congress=119):
    url = f"{API_BASE}/bill/{congress}/{bill_type.lower()}/{bill_number}/cosponsors"
    try:
        resp = requests.get(url, params={"api_key": API_KEY, "format": "json"}, timeout=15)
        resp.raise_for_status()
        return resp.json().get("cosponsors", [])
    except requests.RequestException as e:
        log.warning(f"Could not fetch cosponsors for {bill_type}{bill_number}: {e}")
        return []


def is_climate_energy_bill(bill):
    title       = (bill.get("title") or "").lower()
    policy_area = (bill.get("policyArea", {}) or {}).get("name", "").lower()
    return any(kw in title or kw in policy_area for kw in KEYWORDS)


def format_bill_number(bill_type, number):
    prefix = BILL_TYPE_MAP.get(bill_type.upper(), bill_type + ".")
    return f"{prefix} {number}"


def format_sponsor(sponsor, is_senate):
    if not sponsor:
        return ""
    full_name = sponsor.get("fullName") or f"{sponsor.get('firstName', '')} {sponsor.get('lastName', '')}".strip()
    party     = sponsor.get("party", "")
    state     = sponsor.get("state", "")
    district  = sponsor.get("district", "")
    prefix    = "Sen." if is_senate else "Rep."
    suffix    = f"{party}-{state}"
    if district:
        suffix += f"-{district}"
    return f"{prefix} {full_name} ({suffix})"


def format_date(date_str):
    if not date_str:
        return ""
    try:
        dt = datetime.strptime(date_str[:10], "%Y-%m-%d")
        return dt.strftime("%A, %B %-d. %Y")
    except ValueError:
        return date_str


def latest_action_text(detail):
    latest = detail.get("latestAction", {})
    text   = latest.get("text", "")
    date   = latest.get("actionDate", "")
    return f"{text} {date}".strip() if text else ""


def count_cosponsors_by_party(cosponsors):
    dem = rep = ind = 0
    for c in cosponsors:
        party = (c.get("party") or "").upper()
        if party == "D":
            dem += 1
        elif party == "R":
            rep += 1
        else:
            ind += 1
    return dem, rep, ind


def determine_sheet(bill_type):
    bt = bill_type.upper()
    if bt in ("S", "SRES", "SJRES", "SCONRES"):
        return SENATE_SHEET
    return HOUSE_SHEET


def append_bill_row(ws, row_data):
    new_row = ws.max_row + 1
    for col_idx, value in enumerate(row_data, start=1):
        cell           = ws.cell(row=new_row, column=col_idx, value=value)
        cell.font      = Font(name="Arial", size=10)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    log.info(f"  -> Appended: {row_data[0]} | {row_data[2]}")


def update_excel(new_bills):
    if not EXCEL_PATH.exists():
        log.error(f"Excel file not found: {EXCEL_PATH}")
        return
    wb        = load_workbook(EXCEL_PATH)
    house_ws  = wb[HOUSE_SHEET]
    senate_ws = wb[SENATE_SHEET]
    for bill in new_bills:
        ws = house_ws if bill["sheet"] == HOUSE_SHEET else senate_ws
        row_data = [
            b
