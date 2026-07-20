"""
Standalone Daily Report app — Banksy Museum NYC.
Fully in English.

Run locally:
    streamlit run saisie_manager_ny.py

What it does:
- Daily entry form for the NYC manager (resellers, channels, extras, cash, etc.)
- Saves each day as JSON (Data_SAISIE/B_NY_YYYY-MM-DD.json)
- Fills the colored MasterSheet template (Master_Template_XLS/MasterSheet_NY_Test.xlsx)
  on demand — output = Master_Template_XLS/MasterSheet_NY.xlsx
"""

import json
import os
import sqlite3
from datetime import date, datetime

import pandas as pd
import streamlit as st

# museum_data.db path (project root)
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "museum_data.db")

# ==============================================================================
# CONFIGURATION
# ==============================================================================
SITE_ID = "B_NY"
SITE_CODE = "NY"       # short code — convention: NY (New York), MA (Madrid), BA (Barcelona)…
ACCESS_CODE = "SET_IN_SECRETS"   # real code lives ONLY in Streamlit secrets
CURRENCY = "USD"

# MasterSheet Excel — SINGLE SOURCE OF TRUTH: the colored/paged template.
# Never generate an Excel from scratch — always clone this template and fill it.
_HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(_HERE, "Master_Template_XLS",
                              f"MasterSheet_{SITE_CODE}_Test.xlsx")
OUT_PATH = os.path.join(_HERE, "Master_Template_XLS",
                        f"MasterSheet_{SITE_CODE}.xlsx")

# ---- Cloud persistence (Google Sheet "MasterSheet NY") -----------------------
# The dedicated sheet is the SOURCE OF TRUTH. On Streamlit Cloud the local disk
# is ephemeral, so every save is written here via the OAuth write token.
GSHEET_TAB = "EXPORT_MASTERSHEET"
GSHEET_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
_DEFAULT_GSHEET_ID = "1ot4nhpfyBvs6xbGKGdQrI4NcAE_PduyvGsrQeXDdM7w"
_LOCAL_TOKEN = os.path.join(_HERE, "credentials", "oauth_token_saisie_ny.json")


def _secret(key, default=None):
    """Read a Streamlit secret if configured, else fall back (local run)."""
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return default


GSHEET_ID = _secret("gsheet_id", _DEFAULT_GSHEET_ID)
# On cloud the main DB is NOT deployed → its absence marks cloud mode, which
# hides the local-only tooling (Excel-from-JSON, DB push, JSON history).
IS_CLOUD = not os.path.exists(DB_PATH)

# Login code — overridable via secret so it's never committed to GitHub.
ACCESS_CODE = _secret("access_code", ACCESS_CODE)
# Code required to modify a date that already has a record (overridable via secret).
MOD_CODE = str(_secret("mod_code", "123456"))

# Resellers — Source: Banksy Museum-NYC 2024+2025+2026 MASTER Financials
# Resellers aligned with NY MASTER Sheet — "INCOME 2026" tab (20 columns).
# Source of truth: Google Sheet 1C51BNAjC-B7ul1tHWATypTyEzj9zHxG0nw0FABvcuLc
RESELLERS = [
    ("Fever", "Local"),
    ("Ticketmaster", "OTA"),
    ("Groupon", "OTA"),
    ("Civitatis", "OTA"),
    ("GetYourGuide", "OTA"),
    ("Viator", "OTA"),
    ("Tiqets", "OTA"),
    ("Headout", "OTA"),
    ("GoCity", "OTA"),
    ("Musement", "OTA"),
    ("Rakuten", "OTA"),
    ("Bucket Listers", "Local"),
    ("Klook", "OTA"),
    ("Trip.com", "OTA"),
    ("Tripster", "OTA"),
    ("HotelBeds", "B2B"),
    ("Fun Express", "B2B"),
    ("Tamice", "B2B"),
    ("Cowboy", "B2B"),
    ("Events", "Direct"),
]

# Direct channels — how walk-in visitors found out about the museum
CHANNELS = [
    "BOX OFFICE",
    "Our Website",
    "Word of Mouth",
    "Internet / Google",
    "Google Maps",
    "Social Media",
    "Press",
    "Radio / TV",
    "Hotel Concierge",
    "Walked By",
]

# Ticket rates — source: "Museum Ticket Pricing" tab (Banksy Museum row)
TICKET_RATES = [
    ("Anytime / Enhanced", 40.00),
    ("Adult",              30.00),
    ("Child",              26.00),
    ("Senior",             26.00),
    ("Military",           26.00),
    ("Groups (per pers.)", 21.00),
]

# Group Sales — categories tracked separately for reporting to Decision Makers
GROUP_TYPES = [
    "School Groups",
    "Corporate Groups",
    "Private Tours",
    "Birthday Parties",
    "Other Groups",
]

st.set_page_config(
    page_title="Banksy NYC — Daily Report",
    page_icon="🏛️",
    layout="centered",
)

DATA_DIR = os.path.join(_HERE, "Data_SAISIE")
os.makedirs(DATA_DIR, exist_ok=True)


# ==============================================================================
# I/O
# ==============================================================================
def save_daily_report(report_date, data):
    filename = os.path.join(DATA_DIR, f"{SITE_ID}_{report_date.isoformat()}.json")
    data = dict(data)
    data["site_id"] = SITE_ID
    data["date"] = report_date.isoformat()
    data["timestamp"] = datetime.now().isoformat()
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return filename


def load_existing_report(report_date):
    filename = os.path.join(DATA_DIR, f"{SITE_ID}_{report_date.isoformat()}.json")
    if os.path.exists(filename):
        with open(filename, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def _list_json_files():
    """Return sorted list of daily JSON files for this site."""
    return sorted(
        f for f in os.listdir(DATA_DIR)
        if f.startswith(SITE_ID + "_") and f.endswith(".json")
    )


# ---- Alias : app names → template header names -------------------------------
# The colored template uses canonical French/EN headers. Some form labels differ.
_ALIAS = {
    "Get Your Guide":    "GetYourGuide",
    "Head Out":          "Headout",
    "Ticket Master":     "Ticketmaster",
    "Go City":           "GoCity",
    "Internet / Google": "Internet",
    "Social Media":      "Social Media (FB/Insta)",
    "Hotel Concierge":   "Hotel",
    "Walked By":         "From Our Street",
}
# Resellers present in the NY form but NOT in the template → agrégat
_RESELLER_AGREGE = {"Rakuten", "Bucket Listers", "Trip.com", "Tripster",
                    "HotelBeds", "Fun Express", "Tamice", "Cowboy",
                    "Events", "Company"}

# Group → template column mapping (shared by the Sheet writer)
_GROUP_TARGET = {"School Groups": "School Groups",
                 "Corporate Groups": "Groupe 1 (Entrées)",
                 "Private Tours": "Groupe 2 (Entrées)",
                 "Birthday Parties": "Groupe 3 (Entrées)"}
_GROUP_NAME = {"Corporate Groups": "Groupe 1 (Nom)",
               "Private Tours": "Groupe 2 (Nom)",
               "Birthday Parties": "Groupe 3 (Nom)"}
_GROUP_PRICE = {"Corporate Groups": "Groupe 1 (Prix)",
                "Private Tours": "Groupe 2 (Prix)",
                "Birthday Parties": "Groupe 3 (Prix)"}


# ==============================================================================
# GOOGLE SHEET WRITER — the dedicated "MasterSheet NY" is the source of truth
# ==============================================================================
def _load_token_info():
    """OAuth authorized-user info: from Streamlit secrets on cloud, else local."""
    sec = _secret("gcp_token")
    if sec is not None:
        return dict(sec)
    with open(_LOCAL_TOKEN, encoding="utf-8") as f:
        return json.load(f)


@st.cache_resource(show_spinner=False)
def _gsheet_service():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    creds = Credentials.from_authorized_user_info(_load_token_info(), GSHEET_SCOPES)
    if not creds.valid:
        creds.refresh(Request())
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _resolve_col(col_map, name):
    if name in col_map:
        return col_map[name]
    if name in _ALIAS and _ALIAS[name] in col_map:
        return col_map[_ALIAS[name]]
    return None


def _report_to_row_list(col_map, d, report_date, width):
    """Build the positional row (list of length `width`) for EXPORT_MASTERSHEET."""
    row = [""] * width

    def put(col_name, value):
        c = _resolve_col(col_map, col_name)
        if c and 1 <= c <= width:
            row[c - 1] = value

    put("Date", report_date.isoformat())
    aggr = 0
    for name, _ in RESELLERS:
        qty = d.get(f"rev_{name}") or 0
        if not qty:
            continue
        c = _resolve_col(col_map, name)
        if c:
            row[c - 1] = qty
        elif name in _RESELLER_AGREGE:
            aggr += qty
    if aggr:
        put("Resellers (Agrégé)", aggr)
    for chan in CHANNELS:
        qty = d.get(f"canal_{chan}") or 0
        if qty:
            put(chan, qty)
    for gtype in GROUP_TYPES:
        pax = d.get(f"grp_{gtype}_pax") or 0
        if not pax:
            continue
        if _GROUP_TARGET.get(gtype):
            put(_GROUP_TARGET[gtype], int(pax))
        if _GROUP_NAME.get(gtype):
            put(_GROUP_NAME[gtype], gtype)
        if _GROUP_PRICE.get(gtype):
            put(_GROUP_PRICE[gtype], float(d.get(f"grp_{gtype}_rate") or 0))
    if d.get("audio_qty"):
        put("Audioguide", int(d["audio_qty"]))
    for k, tgt in [("caisse_cash", "Recette Cash (EUR)"),
                   ("caisse_card", "Recette Card (EUR)"),
                   ("caisse_online", "Recette Online (EUR)")]:
        v = d.get(k)
        if v:
            put(tgt, float(v))
    eff_total = d.get("total_tickets") or d.get("tickets_auto")
    if eff_total:
        put("Total Entrées", int(eff_total))
    if d.get("total_revenue"):
        put("Chiffre d'Affaires (€)", float(d["total_revenue"]))
    notes = []
    if d.get("special_event"):
        notes.append(f"🎉 {d['special_event']}")
    if d.get("remarks"):
        notes.append(d["remarks"])
    if notes:
        put("Remarques", " · ".join(notes))
    return row


def save_to_gsheet(report_date, d):
    """Write the day into EXPORT_MASTERSHEET (idempotent: 1 row per date).
    Returns (row_number, values_read_back) — read-after-write proof.
    Raises on any failure so the caller can surface it (no silent loss)."""
    ss = _gsheet_service().spreadsheets()
    meta = ss.get(spreadsheetId=GSHEET_ID,
                  fields="sheets(properties(sheetId,title,"
                         "gridProperties(rowCount)))").execute()
    sheet_id = row_count = None
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == GSHEET_TAB:
            sheet_id = s["properties"]["sheetId"]
            row_count = s["properties"]["gridProperties"]["rowCount"]
            break
    if sheet_id is None:
        raise RuntimeError(f"Tab '{GSHEET_TAB}' not found in the sheet.")

    hdr = ss.values().get(spreadsheetId=GSHEET_ID,
                          range=f"'{GSHEET_TAB}'!2:2").execute().get("values", [[]])
    names = hdr[0] if hdr else []
    col_map = {str(nm).strip(): i for i, nm in enumerate(names, start=1)
               if nm and str(nm).strip()}
    width = len(names)
    if width == 0:
        raise RuntimeError("EXPORT_MASTERSHEET header row is empty.")

    # locate existing row for this date (col A, from row 4), else append
    colA = ss.values().get(spreadsheetId=GSHEET_ID,
                           range=f"'{GSHEET_TAB}'!A4:A").execute().get("values", [])
    date_iso = report_date.isoformat()
    target = None
    for i, r in enumerate(colA):
        if r and str(r[0]).strip()[:10] == date_iso:
            target = i + 4
            break
    if target is None:
        target = 4 + len(colA)

    # extend the grid if the target row is beyond the current size
    if target > row_count:
        ss.batchUpdate(spreadsheetId=GSHEET_ID, body={"requests": [
            {"appendDimension": {"sheetId": sheet_id, "dimension": "ROWS",
                                 "length": target - row_count}}]}).execute()

    row_vals = _report_to_row_list(col_map, d, report_date, width)
    rng = f"'{GSHEET_TAB}'!A{target}:{_col_letter(width)}{target}"
    ss.values().update(spreadsheetId=GSHEET_ID, range=rng,
                       valueInputOption="USER_ENTERED",
                       body={"values": [row_vals]}).execute()
    back = ss.values().get(spreadsheetId=GSHEET_ID,
                           range=rng).execute().get("values", [[]])
    return target, (back[0] if back else [])


def _whatsapp_recap(report_date, d):
    """Pre-composed WhatsApp recap link (wa.me) — museum debrief style, one-tap
    send to any group. Plain text (no astral emojis, which break in wa.me URLs)."""
    from urllib.parse import quote
    total = int(d.get("tickets_auto") or d.get("total_tickets") or 0)
    validated = int(d.get("total_tickets") or 0)
    rev = float(d.get("total_revenue") or 0)

    lines = [report_date.strftime("%d/%m/%Y"), "", f"TOTAL: {total}"]
    if validated:
        lines.append(f"Total validé: {validated}")
    # breakdown: only resellers/groups with tickets that day
    breakdown = [f"{name} : {int(d.get(f'rev_{name}') or 0)}"
                 for name, _kind in RESELLERS if int(d.get(f"rev_{name}") or 0)]
    grp = sum(int(d.get(f"grp_{g}_pax") or 0) for g in GROUP_TYPES)
    if grp:
        breakdown.append(f"Groupe : {grp}")
    if breakdown:
        lines.append("")
        lines.extend(breakdown)
    # cash desk breakdown (Cash / Card / Online) — fall back to estimated revenue
    caisse_lines = [f"{lbl} : ${float(d.get(k) or 0):,.0f}" for k, lbl in
                    (("caisse_cash", "Cash"), ("caisse_card", "Card"),
                     ("caisse_online", "Online")) if float(d.get(k) or 0)]
    lines.append("")
    if caisse_lines:
        lines.extend(caisse_lines)
    else:
        lines.append(f"Revenue : ${rev:,.0f}")
    if d.get("special_event"):
        lines.append(f"Event : {d['special_event']}")
    return "https://wa.me/?text=" + quote("\n".join(lines))


def load_from_gsheet(report_date):
    """Read the existing row for this date from the Sheet and reverse-map it to a
    report dict (for prefilling / recall). Returns None if the date has no data.
    Note: ticket-rate breakdown is not stored in the Sheet, so it isn't recovered."""
    try:
        ss = _gsheet_service().spreadsheets()
        hdr = ss.values().get(spreadsheetId=GSHEET_ID,
                              range=f"'{GSHEET_TAB}'!2:2").execute().get("values", [[]])
        names = hdr[0] if hdr else []
        col_map = {str(nm).strip(): i for i, nm in enumerate(names, 1)
                   if nm and str(nm).strip()}
        colA = ss.values().get(spreadsheetId=GSHEET_ID,
                               range=f"'{GSHEET_TAB}'!A4:A").execute().get("values", [])
        date_iso = report_date.isoformat()
        target = None
        for i, r in enumerate(colA):
            if r and str(r[0]).strip()[:10] == date_iso:
                target = i + 4
                break
        if target is None:
            return None
        row = ss.values().get(spreadsheetId=GSHEET_ID,
                              range=f"'{GSHEET_TAB}'!{target}:{target}").execute().get("values", [[]])
        row = row[0] if row else []

        def gv(colname):
            c = _resolve_col(col_map, colname)
            if c and c - 1 < len(row):
                v = row[c - 1]
                return v if v not in ("", None) else None
            return None

        rep, has = {}, False
        for name, _kind in RESELLERS:
            v = gv(name)
            if v is not None:
                try:
                    rep[f"rev_{name}"] = float(v); has = True
                except ValueError:
                    pass
        for chan in CHANNELS:
            v = gv(chan)
            if v is not None:
                try:
                    rep[f"canal_{chan}"] = int(float(v)); has = True
                except ValueError:
                    pass
        for gtype in GROUP_TYPES:
            tcol = _GROUP_TARGET.get(gtype)
            v = gv(tcol) if tcol else None
            if v is not None:
                try:
                    rep[f"grp_{gtype}_pax"] = int(float(v)); has = True
                except ValueError:
                    pass
            pc = _GROUP_PRICE.get(gtype)
            pv = gv(pc) if pc else None
            if pv is not None:
                try:
                    rep[f"grp_{gtype}_rate"] = float(pv)
                except ValueError:
                    pass
        av = gv("Audioguide")
        if av is not None:
            try:
                rep["audio_qty"] = int(float(av)); has = True
            except ValueError:
                pass
        for key, col in (("caisse_cash", "Recette Cash (EUR)"),
                         ("caisse_card", "Recette Card (EUR)"),
                         ("caisse_online", "Recette Online (EUR)")):
            v = gv(col)
            if v is not None:
                try:
                    rep[key] = float(v); has = True
                except ValueError:
                    pass
        tv = gv("Total Entrées")
        if tv is not None:
            try:
                rep["total_tickets"] = int(float(tv)); has = True
            except ValueError:
                pass
        rv = gv("Remarques")
        if rv:
            has = True
            s = str(rv)
            if s.startswith("🎉 "):
                parts = s[2:].split(" · ", 1)
                rep["special_event"] = parts[0].strip()
                if len(parts) > 1:
                    rep["remarks"] = parts[1]
            else:
                rep["remarks"] = s
        return rep if has else None
    except Exception:
        return None


def export_mastersheet_ny():
    """Clone MasterSheet_NY_Test.xlsx (colored, paged template) and fill the
    EXPORT_MASTERSHEET tab with EVERY saved daily report. Formulas and styles
    of the template are preserved. Output = Master_Template_XLS/MasterSheet_NY.xlsx.
    Returns (path, n_days) or raises.
    """
    import openpyxl
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")
    files = _list_json_files()
    if not files:
        raise ValueError("No daily report saved yet.")

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb["EXPORT_MASTERSHEET"]

    # Build header -> column index map (row 2)
    col_map = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=2, column=c).value
        if v:
            col_map[str(v).strip()] = c

    def resolve_col(name):
        if name in col_map:
            return col_map[name]
        if name in _ALIAS and _ALIAS[name] in col_map:
            return col_map[_ALIAS[name]]
        return None

    # Clear existing data rows (row 4+), preserve formulas
    for r in range(4, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if not isinstance(cell.value, str) or not cell.value.startswith("="):
                cell.value = None

    start_row = 4
    r = start_row
    for f in files:
        date_iso = f.replace(SITE_ID + "_", "").replace(".json", "")
        with open(os.path.join(DATA_DIR, f)) as fh:
            d = json.load(fh)
        try:
            date_obj = datetime.strptime(date_iso, "%Y-%m-%d")
        except ValueError:
            continue
        # Date
        if "Date" in col_map:
            cell = ws.cell(row=r, column=col_map["Date"], value=date_obj)
            cell.number_format = "DD/MM/YYYY"
        # Resellers → direct cols; aggregate fallback for those absent from template
        aggr_qty = 0
        for name, _ in RESELLERS:
            qty = d.get(f"rev_{name}") or 0
            if not qty:
                continue
            col = resolve_col(name)
            if col:
                ws.cell(row=r, column=col, value=qty)
            elif name in _RESELLER_AGREGE:
                aggr_qty += qty
        if aggr_qty and "Resellers (Agrégé)" in col_map:
            ws.cell(row=r, column=col_map["Resellers (Agrégé)"], value=aggr_qty)
        # Channels
        for chan in CHANNELS:
            qty = d.get(f"canal_{chan}") or 0
            if qty:
                col = resolve_col(chan)
                if col:
                    ws.cell(row=r, column=col, value=qty)
        # Groups (School Groups direct col ; Corporate/Private/Birthday → Groupe 1/2/3)
        group_target = {
            "School Groups":     "School Groups",
            "Corporate Groups":  "Groupe 1 (Entrées)",
            "Private Tours":     "Groupe 2 (Entrées)",
            "Birthday Parties":  "Groupe 3 (Entrées)",
        }
        group_name_col = {
            "Corporate Groups": "Groupe 1 (Nom)",
            "Private Tours":    "Groupe 2 (Nom)",
            "Birthday Parties": "Groupe 3 (Nom)",
        }
        group_price_col = {
            "Corporate Groups": "Groupe 1 (Prix)",
            "Private Tours":    "Groupe 2 (Prix)",
            "Birthday Parties": "Groupe 3 (Prix)",
        }
        for gtype in GROUP_TYPES:
            pax = d.get(f"grp_{gtype}_pax") or 0
            if not pax:
                continue
            tcol = group_target.get(gtype)
            if tcol and tcol in col_map:
                ws.cell(row=r, column=col_map[tcol], value=int(pax))
                nc = group_name_col.get(gtype)
                if nc and nc in col_map:
                    ws.cell(row=r, column=col_map[nc], value=gtype)
                pc = group_price_col.get(gtype)
                if pc and pc in col_map:
                    ws.cell(row=r, column=col_map[pc],
                            value=float(d.get(f"grp_{gtype}_rate") or 0))
        # Audioguide
        if "Audioguide" in col_map and d.get("audio_qty"):
            ws.cell(row=r, column=col_map["Audioguide"], value=int(d["audio_qty"]))
        # Cash desk (USD in EUR-labeled cols by convention)
        for k, tgt in [("caisse_cash",   "Recette Cash (EUR)"),
                       ("caisse_card",   "Recette Card (EUR)"),
                       ("caisse_online", "Recette Online (EUR)")]:
            v = d.get(k)
            if v and tgt in col_map:
                ws.cell(row=r, column=col_map[tgt], value=float(v))
        # Totals
        if "Total Entrées" in col_map and d.get("total_tickets"):
            ws.cell(row=r, column=col_map["Total Entrées"],
                    value=int(d["total_tickets"]))
        if "Chiffre d'Affaires (€)" in col_map and d.get("total_revenue"):
            ws.cell(row=r, column=col_map["Chiffre d'Affaires (€)"],
                    value=float(d["total_revenue"]))
        # Notes (concat special event + remarks)
        notes = []
        if d.get("special_event"):
            notes.append(f"🎉 {d['special_event']}")
        if d.get("remarks"):
            notes.append(d["remarks"])
        if notes and "Remarques" in col_map:
            ws.cell(row=r, column=col_map["Remarques"], value=" · ".join(notes))
        r += 1

    last_row = r - 1

    # TOTAL row with formulas
    tot_row = last_row + 2
    ws.cell(row=tot_row, column=1, value="TOTAL")
    ws.cell(row=tot_row, column=1).font = Font(bold=True, size=11, color="FFFFFF")
    ws.cell(row=tot_row, column=1).fill = PatternFill("solid", fgColor="305496")
    for c in range(2, ws.max_column + 1):
        hdr = ws.cell(row=2, column=c).value or ""
        if hdr in ("Date", "Remarques") or "Nom" in str(hdr):
            continue
        letter = get_column_letter(c)
        ws.cell(row=tot_row, column=c,
                value=f"=SUM({letter}{start_row}:{letter}{last_row})")
        ws.cell(row=tot_row, column=c).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=tot_row, column=c).fill = PatternFill("solid", fgColor="305496")

    # Update Configuration_Museum tab with SITE_ID
    if "Configuration_Museum" in wb.sheetnames:
        cfg = wb["Configuration_Museum"]
        for rr in range(2, cfg.max_row + 1):
            if str(cfg.cell(row=rr, column=1).value).upper() == "SITE_ID":
                cfg.cell(row=rr, column=2, value=SITE_ID)
                break

    ws.freeze_panes = "B4"
    wb.save(OUT_PATH)
    return OUT_PATH, last_row - start_row + 1


# ==============================================================================
# LOGIN
# ==============================================================================
if "ny_logged_in" not in st.session_state:
    st.session_state.ny_logged_in = False

if not st.session_state.ny_logged_in:
    st.markdown("# Banksy Museum NYC")
    st.markdown("### Daily Report — Manager Login")
    st.markdown("---")

    code = st.text_input("Access code", type="password", placeholder="Enter your manager code")
    if st.button("Sign in", type="primary", use_container_width=True):
        if code == ACCESS_CODE:
            st.session_state.ny_logged_in = True
            st.rerun()
        else:
            st.error("Invalid code")

    st.markdown("---")
    st.caption("Contact admin if you have forgotten your code.")
    st.stop()


# ==============================================================================
# FORM
# ==============================================================================
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown("# Banksy Museum NYC")
    st.caption("Daily activity report")
with col_h2:
    if st.button("🔓 Sign out"):
        st.session_state.ny_logged_in = False
        st.rerun()

st.markdown("---")

report_date = st.date_input(
    "📅 Report date",
    value=date.today(),
    format="DD/MM/YYYY",
    max_value=date.today(),
)

# Recall a day from the Sheet (cloud-persistent). Load ONLY when the date changes,
# then wipe widget state so all fields re-init from the recalled values.
_date_iso = report_date.isoformat()
if st.session_state.get("_loaded_date") != _date_iso:
    st.session_state["_existing"] = load_from_gsheet(report_date) \
        or load_existing_report(report_date)
    for _k in list(st.session_state.keys()):
        if _k.startswith(("rate_", "rev_", "canal_", "grp_", "audio_", "caisse_")) \
           or _k in ("total_tickets", "special_event", "remarks", "mod_code_input"):
            del st.session_state[_k]
    st.session_state["_loaded_date"] = _date_iso
    st.rerun()
existing = st.session_state.get("_existing")

# Editing a date that already has a record requires the modification code.
edit_unlocked = True
if existing:
    st.warning(f"📝 A record already exists for {report_date.strftime('%d/%m/%Y')} — "
               "fields are pre-filled below.")
    _mod = st.text_input("🔒 Modification code", type="password", key="mod_code_input",
                         placeholder="Enter the code to modify this record")
    edit_unlocked = (_mod == MOD_CODE)
    if not edit_unlocked:
        st.caption("🔒 Enter the modification code above to unlock saving.")


def get_default(key, default=0):
    if existing and key in existing:
        return existing[key]
    return default


# ==============================================================================
# 🎟️ TICKET PRICING BREAKDOWN — nb de billets par tarif
# Source : "Museum Ticket Pricing" tab of the NY MASTER Sheet.
# ==============================================================================
st.markdown("### 🎟️ Ticket rates sold today")
st.caption("Tickets sold at the door, breakdown per rate (auto-computes revenue).")

pricing_data = {}
pricing_revenue = 0.0
for i in range(0, len(TICKET_RATES), 3):
    cols = st.columns(3)
    for j, col in enumerate(cols):
        idx = i + j
        if idx >= len(TICKET_RATES):
            break
        name, price = TICKET_RATES[idx]
        with col:
            qty = st.number_input(
                name,
                min_value=0,
                value=int(get_default(f"rate_{name}", 0)),
                step=1,
                key=f"rate_{name}",
            )
            if qty > 0:
                pricing_data[name] = {"qty": qty, "price": price, "revenue": qty * price}
                pricing_revenue += qty * price

st.caption(f"💰 Ticket rates revenue: **${pricing_revenue:,.2f}**")

# ==============================================================================
# 🔵 RESELLERS & PLATFORMS — nb de tickets vendus par plateforme
# ==============================================================================
st.markdown("### 🔵 Resellers & Platforms")
st.caption(f"Number of tickets sold per platform — currency: **{CURRENCY}**")

reseller_data = {}
cols_per_row = 3
for i in range(0, len(RESELLERS), cols_per_row):
    cols = st.columns(cols_per_row)
    for j, col in enumerate(cols):
        idx = i + j
        if idx >= len(RESELLERS):
            break
        name, kind = RESELLERS[idx]
        with col:
            val = st.number_input(
                name,
                min_value=0.0,
                value=float(get_default(f"rev_{name}", 0)),
                step=1.0,
                key=f"rev_{name}",
                help=kind,
            )
            if val > 0:
                reseller_data[name] = val

# ==============================================================================
# 🟢 DIRECT CHANNELS — attribution : « comment ont-ils connu le musée ? »
# ==============================================================================
st.markdown("### 🟢 Direct channels")
st.caption("How visitors found us — declared at entrance")

channel_data = {}
for i in range(0, len(CHANNELS), cols_per_row):
    cols = st.columns(cols_per_row)
    for j, col in enumerate(cols):
        idx = i + j
        if idx >= len(CHANNELS):
            break
        name = CHANNELS[idx]
        with col:
            val = st.number_input(
                name,
                min_value=0,
                value=int(get_default(f"canal_{name}", 0)),
                step=1,
                key=f"canal_{name}",
            )
            if val > 0:
                channel_data[name] = val

# ==============================================================================
# 👥 GROUP SALES — School, Corporate, Private, Birthday…
# ==============================================================================
st.markdown("### 👥 Group Sales")
st.caption("Structured tracking of group visits (School / Corporate / Private / Birthday). "
           "Rate = per-person price (default $21 = Groups rate).")

group_data = {}
group_revenue = 0.0
for gtype in GROUP_TYPES:
    with st.expander(f"➕ {gtype}", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            g_pax = st.number_input("People", min_value=0,
                                     value=int(get_default(f"grp_{gtype}_pax", 0)),
                                     step=1, key=f"grp_{gtype}_pax")
        with c2:
            g_rate = st.number_input("Rate / person", min_value=0.0,
                                      value=float(get_default(f"grp_{gtype}_rate", 21.0)),
                                      step=1.0, key=f"grp_{gtype}_rate")
        with c3:
            g_rev = g_pax * g_rate
            st.metric("Revenue", f"${g_rev:,.2f}")
        g_notes = st.text_input("Notes (contact, org name…)",
                                 value=get_default(f"grp_{gtype}_notes", ""),
                                 key=f"grp_{gtype}_notes")
        if g_pax > 0:
            group_data[gtype] = {"pax": g_pax, "rate": g_rate,
                                  "revenue": g_rev, "notes": g_notes}
            group_revenue += g_rev

if group_revenue:
    st.caption(f"👥 Groups total revenue: **${group_revenue:,.2f}**")

# ==============================================================================
# 🎧 AUDIOGUIDE — qty + price + auto revenue
# ==============================================================================
st.markdown("### 🎧 Audioguide")
c1, c2, c3 = st.columns(3)
with c1:
    audio_qty = st.number_input("Audioguides rented", min_value=0,
                                 value=int(get_default("audio_qty", 0)),
                                 step=1, key="audio_qty")
with c2:
    audio_price = st.number_input("Price / unit ($)", min_value=0.0,
                                   value=float(get_default("audio_price", 5.0)),
                                   step=0.5, key="audio_price")
with c3:
    audio_revenue = audio_qty * audio_price
    st.metric("Audio revenue", f"${audio_revenue:,.2f}")

# ==============================================================================
# 💵 CASH DESK — Cash / Card / Online (physical takings breakdown)
# ==============================================================================
st.markdown(f"### 💵 Daily takings ({CURRENCY})")
c1, c2, c3 = st.columns(3)
with c1:
    caisse_cash = st.number_input("Cash", min_value=0.0,
                                  value=float(get_default("caisse_cash", 0)),
                                  step=0.5, key="caisse_cash")
with c2:
    caisse_card = st.number_input("Card", min_value=0.0,
                                  value=float(get_default("caisse_card", 0)),
                                  step=0.5, key="caisse_card")
with c3:
    caisse_online = st.number_input("Online", min_value=0.0,
                                    value=float(get_default("caisse_online", 0)),
                                    step=0.5, key="caisse_online")

st.caption(f"💰 Cash desk total: **${(caisse_cash+caisse_card+caisse_online):,.2f}**  ·  "
           f"Consistency check → should match Ticket rates revenue "
           f"(${pricing_revenue:,.2f}) + Audio (${audio_revenue:,.2f}) "
           f"+ Groups (${group_revenue:,.2f}).")

# ==============================================================================
# ⬛ SUMMARY & CROSS-CHECK
# ==============================================================================
st.markdown("### ⬛ Summary")
tickets_from_rates = sum(v["qty"] for v in pricing_data.values())
tickets_from_resellers = sum(reseller_data.values())
tickets_from_channels = sum(channel_data.values())
tickets_from_groups = sum(g["pax"] for g in group_data.values())
tickets_auto = tickets_from_rates + tickets_from_resellers + tickets_from_channels + tickets_from_groups
total_revenue = pricing_revenue + group_revenue + audio_revenue

c1, c2, c3 = st.columns(3)
with c1:
    total_tickets = st.number_input(
        "Total visitors (if known)",
        min_value=0,
        value=int(get_default("total_tickets", 0)),
        step=1,
        key="total_tickets",
    )
with c2:
    st.metric("Auto tickets (rates + resellers + channels + groups)", f"{tickets_auto:.0f}")
with c3:
    st.metric("Total estimated revenue", f"${total_revenue:,.2f}")

# ---- Special event & remarks
special_event = st.text_input(
    "🎉 Special event / holiday",
    value=get_default("special_event", ""),
    placeholder="e.g. Halloween Festival, Thanksgiving, Broadway premiere…",
    key="special_event",
)
remarks = st.text_area(
    "📝 Notes",
    value=get_default("remarks", ""),
    key="remarks",
    placeholder="Anything worth mentioning: staff issue, technical incident, VIP visit…",
)

st.markdown("---")

# ==============================================================================
# ✅ SAVE
# ==============================================================================
if st.button("✅ Save report", type="primary", use_container_width=True):
    if existing and not edit_unlocked:
        st.error("🔒 Modification code required to overwrite an existing record.")
        st.stop()
    report = {}
    # Ticket rates breakdown
    for name, d in pricing_data.items():
        report[f"rate_{name}_qty"] = d["qty"]
        report[f"rate_{name}_price"] = d["price"]
        report[f"rate_{name}_revenue"] = d["revenue"]
    # Resellers
    for k, v in reseller_data.items():
        report[f"rev_{k}"] = v
    # Channels
    for k, v in channel_data.items():
        report[f"canal_{k}"] = v
    # Groups (detailed)
    for gtype, d in group_data.items():
        report[f"grp_{gtype}_pax"] = d["pax"]
        report[f"grp_{gtype}_rate"] = d["rate"]
        report[f"grp_{gtype}_revenue"] = d["revenue"]
        report[f"grp_{gtype}_notes"] = d["notes"]
    # Audioguide
    report["audio_qty"] = audio_qty
    report["audio_price"] = audio_price
    report["audio_revenue"] = audio_revenue
    # Cash desk
    report["caisse_cash"] = caisse_cash
    report["caisse_card"] = caisse_card
    report["caisse_online"] = caisse_online
    # Aggregates
    report["total_tickets"] = total_tickets
    report["tickets_auto"] = tickets_auto
    report["total_revenue"] = total_revenue
    # Context
    report["special_event"] = special_event
    report["remarks"] = remarks

    # 1. Google Sheet = source of truth. Retry on transient errors (e.g. broken
    #    pipe); the write is idempotent (1 row per date) so retry is safe.
    row_n, _last_err = None, None
    for _attempt in range(3):
        try:
            row_n, _back = save_to_gsheet(report_date, report)
            break
        except Exception as e:
            _last_err = e
            try:
                _gsheet_service.clear()   # drop a possibly-stale cached connection
            except Exception:
                pass
    if row_n is None:
        st.error(f"⚠️ NOT saved to the sheet — please retry. Details: {_last_err}")
        st.stop()
    st.session_state["_loaded_date"] = None   # force a fresh re-read next time
    st.success(
        f"✅ Report for {report_date.strftime('%m/%d/%Y')} saved to "
        f"MasterSheet NY (row {row_n})."
    )
    st.link_button(
        "📲 Send WhatsApp recap",
        _whatsapp_recap(report_date, report),
        use_container_width=True,
    )

    # 2. Local JSON copy (kept for the local Excel/DB tooling; harmless on cloud).
    try:
        json_path = save_daily_report(report_date, report)
        st.caption(f"Local copy: `{os.path.basename(json_path)}`")
    except Exception:
        pass

# ==============================================================================
# 📤 PUSH TO museum_data.db
# ==============================================================================
def _get_usd_eur_rate(cur, date_iso):
    """Return USD/EUR rate for a given date (fallback : last known)."""
    row = cur.execute("""SELECT rate FROM fx_daily
        WHERE ccy_from='USD' AND ccy_to='EUR' AND date <= ?
        ORDER BY date DESC LIMIT 1""", (date_iso,)).fetchone()
    return float(row[0]) if row and row[0] else 1.08  # fallback approx


def push_to_db(date_from, date_to):
    """Import all Data_SAISIE/B_NY_YYYY-MM-DD.json between date_from..date_to
    into museum_data.db (sales + sales_details).
    USD values are converted to EUR via fx_daily (BCE).
    Idempotent: INSERT OR REPLACE on sales, DELETE + INSERT on sales_details.
    Returns dict with counters."""
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(f"DB not found: {DB_PATH}")

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    n_days = n_details = 0
    total_tickets_pushed = 0
    total_revenue_eur = 0.0
    detail_days = []

    for f in _list_json_files():
        d_str = f.replace(f"{SITE_ID}_", "").replace(".json", "")
        try:
            d_iso = datetime.strptime(d_str, "%Y-%m-%d").date().isoformat()
        except ValueError:
            continue
        if d_iso < date_from.isoformat() or d_iso > date_to.isoformat():
            continue

        with open(os.path.join(DATA_DIR, f), encoding="utf-8") as fh:
            data = json.load(fh)

        total_tk = int(data.get("total_tickets") or 0)
        total_rev_usd = float(data.get("total_revenue") or 0)

        # Convert USD -> EUR via fx_daily
        fx = _get_usd_eur_rate(cur, d_iso)
        total_rev_eur = total_rev_usd / fx if fx else 0.0
        prix_moy_eur = (total_rev_eur / total_tk) if total_tk > 0 else None

        # INSERT OR REPLACE sales (identifier = date + site)
        # First find existing sale_id if any
        row = cur.execute("SELECT id FROM sales WHERE site=? AND date=?",
                          (SITE_ID, d_iso)).fetchone()
        raw_msg = f"saisie_manager_ny @ {data.get('timestamp', '')}"
        if row:
            sale_id = row[0]
            cur.execute("""UPDATE sales SET total=?, prix_moyen_ref=?,
                raw_msg=?, is_manuel=0 WHERE id=?""",
                (total_tk, prix_moy_eur, raw_msg, sale_id))
            # Wipe old sales_details for this sale_id
            cur.execute("DELETE FROM sales_details WHERE sale_id=?", (sale_id,))
        else:
            cur.execute("""INSERT INTO sales
                (date, site, total, validated, prix_moyen_ref, raw_msg, is_manuel)
                VALUES (?, ?, ?, ?, ?, ?, 0)""",
                (d_iso, SITE_ID, total_tk, total_tk, prix_moy_eur, raw_msg))
            sale_id = cur.lastrowid

        n_details_row = 0
        # Insert sales_details : revendeurs
        for k, v in data.items():
            if k.startswith("rev_") and float(v or 0) > 0:
                rev_name = k[4:]  # strip "rev_"
                cur.execute("""INSERT INTO sales_details
                    (sale_id, revendeur, quantite, canal, classification, revendeur_standard)
                    VALUES (?, ?, ?, ?, ?, ?)""",
                    (sale_id, rev_name, float(v), "Revendeur", "Revendeur",
                     rev_name.upper()))
                n_details_row += 1
            elif k.startswith("canal_") and float(v or 0) > 0:
                canal_name = k[6:]  # strip "canal_"
                cur.execute("""INSERT INTO sales_details
                    (sale_id, revendeur, quantite, canal, classification, revendeur_standard)
                    VALUES (?, ?, ?, ?, ?, ?)""",
                    (sale_id, canal_name, float(v), "Canal", "Canal",
                     canal_name.upper()))
                n_details_row += 1
            elif k.startswith("rate_") and k.endswith("_qty") and float(v or 0) > 0:
                tarif = k[5:-4]  # strip "rate_" and "_qty"
                cur.execute("""INSERT INTO sales_details
                    (sale_id, revendeur, quantite, canal, classification, revendeur_standard)
                    VALUES (?, ?, ?, ?, ?, ?)""",
                    (sale_id, tarif, float(v), "Tarif", "Tarif",
                     tarif.upper()))
                n_details_row += 1

        n_days += 1
        n_details += n_details_row
        total_tickets_pushed += total_tk
        total_revenue_eur += total_rev_eur
        detail_days.append((d_iso, total_tk, total_rev_eur, fx, n_details_row))

    conn.commit()
    conn.close()

    return {
        "n_days": n_days,
        "n_details": n_details,
        "total_tickets": total_tickets_pushed,
        "total_revenue_eur": total_revenue_eur,
        "detail_days": detail_days,
    }


# ------------------------------------------------------------------------------
# Cloud stops here: the tooling below (Excel-from-JSON, DB push, JSON history) is
# local-only. On the cloud, the Google Sheet is already the source of truth.
# ------------------------------------------------------------------------------
if IS_CLOUD:
    st.stop()

# ==============================================================================
# 📊 EXPORT MASTERSHEET — clone the colored/paged template + fill it
# ==============================================================================
st.markdown("---")
with st.container(border=True):
    st.markdown("### 📊 Export MasterSheet Excel")
    st.caption(
        f"Clones `Master_Template_XLS/MasterSheet_{SITE_CODE}_Test.xlsx` "
        f"(colored template — {119} columns, formulas, extra tabs "
        "Configuration/Légende/Mapping/Commissions/Lexique) and fills the "
        "EXPORT_MASTERSHEET tab with every saved daily report. Output = "
        f"`Master_Template_XLS/MasterSheet_{SITE_CODE}.xlsx`."
    )

    if st.button("🛠️ Generate MasterSheet Excel", type="primary",
                  use_container_width=True, key="btn_gen_mastersheet"):
        try:
            path, n_days = export_mastersheet_ny()
            st.success(f"✅ {n_days} day(s) consolidated into `{os.path.basename(path)}`")
        except FileNotFoundError as e:
            st.error(str(e))
        except ValueError as e:
            st.warning(str(e))
        except Exception as e:
            st.error(f"Error: {e}")

    if os.path.exists(OUT_PATH):
        mtime = datetime.fromtimestamp(os.path.getmtime(OUT_PATH))
        with open(OUT_PATH, "rb") as fh:
            st.download_button(
                f"📥 Download MasterSheet_{SITE_CODE}.xlsx",
                data=fh.read(),
                file_name=f"MasterSheet_{SITE_CODE}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.caption(f"Last generated: {mtime.strftime('%m/%d/%Y %H:%M')}")
    else:
        st.caption("Not generated yet — click Generate above.")

# ==============================================================================
# 📤 PUSH TO museum_data.db (only shown if DB is present next to the app)
# ==============================================================================
_DB_PRESENT = os.path.exists(DB_PATH)
if not _DB_PRESENT:
    # Manager standalone install : hide DB section entirely.
    pass
else:
    st.markdown("---")
    with st.container(border=True):
        st.markdown("### 📤 Push to museum_data.db")
        st.caption(
            "Load the saved daily reports (JSON) into the main DB. "
            "Idempotent — running twice on the same period simply overwrites."
        )

        _all_files = _list_json_files()
        if _all_files:
            _dates = []
            for _f in _all_files:
                try:
                    _dates.append(datetime.strptime(
                        _f.replace(f"{SITE_ID}_", "").replace(".json", ""),
                        "%Y-%m-%d").date())
                except ValueError:
                    pass
            _min_d, _max_d = (min(_dates), max(_dates)) if _dates \
                else (date.today(), date.today())

            c1, c2 = st.columns(2)
            with c1:
                _from = st.date_input("From", value=_min_d, min_value=_min_d,
                                        max_value=_max_d, key="push_from")
            with c2:
                _to = st.date_input("To", value=_max_d, min_value=_min_d,
                                      max_value=_max_d, key="push_to")

            if st.button("📤 Push selected reports to museum_data.db",
                          type="primary", use_container_width=True,
                          key="btn_push_db"):
                try:
                    res = push_to_db(_from, _to)
                    total_usd = sum(rev_eur * fx for _d, _tk, rev_eur, fx, _n
                                     in res['detail_days'])
                    st.success(
                        f"✅ {res['n_days']} day(s) pushed · "
                        f"{res['total_tickets']:,} tickets · "
                        f"{total_usd:,.0f} USD · "
                        f"{res['n_details']} detail rows"
                    )
                    if res['detail_days']:
                        with st.expander("Details per day"):
                            rows_usd = [(d, tk, rev_eur * fx, n)
                                        for d, tk, rev_eur, fx, n
                                        in res['detail_days']]
                            st.dataframe(pd.DataFrame(
                                rows_usd,
                                columns=["Date", "Tickets", "Revenue (USD)",
                                         "Detail rows"]
                            ), use_container_width=True, hide_index=True)
                except Exception as e:
                    st.error(f"Error: {e}")
        else:
            st.info("No JSON report yet — save at least one daily report first.")

# ---- RECENT HISTORY ----
st.markdown("### 📋 Last 7 reports")
files = sorted(
    [f for f in os.listdir(DATA_DIR) if f.startswith(SITE_ID) and f.endswith(".json")],
    reverse=True,
)[:7]
if files:
    for f in files:
        d_str = f.replace(f"{SITE_ID}_", "").replace(".json", "")
        try:
            d = datetime.strptime(d_str, "%Y-%m-%d").strftime("%m/%d/%Y")
        except ValueError:
            d = d_str
        with open(os.path.join(DATA_DIR, f)) as fh:
            data = json.load(fh)
        revs = {k.replace("rev_", ""): v for k, v in data.items() if k.startswith("rev_") and v > 0}
        total = data.get("total_tickets", 0)
        st.caption(f"**{d}** — Total: {total} — {revs}")
else:
    st.caption("No report saved yet.")
