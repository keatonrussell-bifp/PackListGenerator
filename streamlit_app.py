import streamlit as st
import pandas as pd
import io
import pdfplumber
import re
import os
from datetime import datetime

# ---------------- REGEX AND HELPERS ----------------

TRANSPORT_RX = re.compile(r"Transport\s*:\s*([A-Z0-9]+)")
SECTION_GRADE_RX = re.compile(
    r"Sawn timber,\s*Taeda-Pine,\s*([^,]+),",
    re.IGNORECASE
)
ROW_RX_IMPERIAL = re.compile(
    r'(?P<thick>\d+)[,\.]0"?\s+'
    r'(?P<width>\d+)[,\.]0"?\s+'
    r'(?P<length>(?:\d+\s*\d/\d|\d+))"?\s+'
    r'(?P<qty>\d+)\s*Pcs\s+'
    r'(?P<mbf>\d+[,\.]\d+)\s*mbf',
    re.IGNORECASE
)
ROW_RX_METRIC = re.compile(
    r'(?P<thick_mm>\d+[,\.\d]*)\s*mm\s+'
    r'(?P<width_mm>\d+[,\.\d]*)\s*mm\s+'
    r'(?P<length_m>\d+[,\.\d]*)\s*m\b.*?'
    r'(?P<qty>\d+)\s*Pcs',
    re.IGNORECASE
)

PCS_RX  = re.compile(r'(\d[\d\.,]*)\s*pcs', re.IGNORECASE)
MBF_RX  = re.compile(r'(\d[\d\.,]*)\s*mbf', re.IGNORECASE)
PKGS_RX = re.compile(r'(\d[\d\.,]*)\s*pkg[s]?', re.IGNORECASE)

MM_TO_IN = 1.0 / 25.4
M_TO_IN  = 39.37007874015748


def frac_to_float(s: str) -> float:
    s = s.strip().replace(',', '.')
    if ' ' in s and '/' in s:
        whole, frac = s.split(' ', 1)
        num, den = frac.split('/')
        return float(whole) + float(num) / float(den)
    if '/' in s:
        num, den = s.split('/')
        return float(num) / float(den)
    return float(s)


def parse_footer_totals(full_text: str):
    pcs_vals = [m.group(1) for m in PCS_RX.finditer(full_text)]
    mbf_vals = [m.group(1) for m in MBF_RX.finditer(full_text)]
    pkgs_vals = [m.group(1) for m in PKGS_RX.finditer(full_text)]

    def to_int(x):   return int(x.replace('.', '').replace(',', ''))
    def to_float(x): return float(x.replace('.', '').replace(',', '.'))

    return {
        "pcs":  to_int(pcs_vals[-1])   if pcs_vals else None,
        "mbf":  to_float(mbf_vals[-1]) if mbf_vals else None,
        "pkgs": to_int(pkgs_vals[-1])  if pkgs_vals else None,
    }


def parse_pdf_from_bytes(pdf_bytes, name="PDF"):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pages = [p.extract_text() or "" for p in pdf.pages]

    full_text = "\n".join(pages)
    expected  = parse_footer_totals(full_text)

    transport = None
    items = []
    cur_grade = None

    for page_text in pages:
        if transport is None:
            m = TRANSPORT_RX.search(page_text)
            if m:
                transport = m.group(1)

        for ln in (ln.strip() for ln in page_text.splitlines() if ln.strip()):

            g = SECTION_GRADE_RX.search(ln)
            if g:
                cur_grade = g.group(1).strip().upper()

            mi = ROW_RX_IMPERIAL.search(ln)
            if mi:
                items.append({
                    "Transport": transport or name,
                    "Thickness (in)": float(mi.group("thick").replace(",", ".")),
                    "Width (in)":     float(mi.group("width").replace(",", ".")),
                    "Length (in)":    frac_to_float(mi.group("length")),
                    "Quantity (pcs)": int(mi.group("qty")),
                    "Volume (mbf)":   float(mi.group("mbf").replace(",", ".")),
                    "Grade":          cur_grade or "UNKNOWN",
                })
                continue

            mm = ROW_RX_METRIC.search(ln)
            if mm:
                items.append({
                    "Transport":      transport or name,
                    "Thickness (in)": float(mm.group("thick_mm").replace(",", ".").replace(" ", "")) * MM_TO_IN,
                    "Width (in)":     float(mm.group("width_mm").replace(",", ".").replace(" ", "")) * MM_TO_IN,
                    "Length (in)":    float(mm.group("length_m").replace(",", ".").replace(" ", "")) * M_TO_IN,
                    "Quantity (pcs)": int(mm.group("qty")),
                    "Volume (mbf)":   float('nan'),
                    "Grade":          cur_grade or "UNKNOWN",
                })
                continue

    if not transport:
        transport = name

    if not items:
        return transport, None, expected, None

    df = pd.DataFrame(items)

    df["Dimension"] = (
        df["Thickness (in)"].round().astype(int).astype(str) + "X" +
        df["Width (in)"].round().astype(int).astype(str) + "-" +
        df["Length (in)"].round().astype(int).astype(str)
    )

    summary = (
        df.groupby(["Transport", "Grade", "Dimension"], as_index=False)
          .agg(Count=("Transport", "size"),
               Total_Pcs=("Quantity (pcs)", "sum"),
               Total_MBF=("Volume (mbf)", "sum"))
    )

    actual = {
        "pkgs": int(df.shape[0]),
        "pcs":  int(df["Quantity (pcs)"].sum()),
        "mbf":  float(df["Volume (mbf)"].sum(skipna=True)) if df["Volume (mbf)"].notna().any() else None,
    }

    return transport, summary, expected, actual


# ---------------- MSC LOOKUP (API first, scrape fallback) ----------------

def _norm_container(x: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", (x or "").upper().strip())


def _parse_date_safe(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip()
    fmts = [
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
        "%d %b %Y",
        "%d %B %Y",
        "%b %d, %Y",
        "%B %d, %Y",
    ]
    for f in fmts:
        try:
            return datetime.strptime(s, f).date().isoformat()
        except Exception:
            pass
    return s


def fetch_msc_arrivals_by_booking(booking_number: str, log_fn=None) -> dict:
    """
    Returns: {CONTAINER_NUMBER: arrival_or_eta_string}
    - If MSC_TT_BASE_URL + MSC_TT_TOKEN are set, attempts API.
    - Else, scrapes MSC public tracking page via Playwright.
    """
    booking_number = (booking_number or "").strip()
    if not booking_number:
        return {}

    def log(msg: str):
        if log_fn:
            log_fn(msg)

    # ---- Option A: MSC API (stable if you have credentials) ----
    base_url = os.getenv("MSC_TT_BASE_URL", "").strip()
    token = os.getenv("MSC_TT_TOKEN", "").strip()

    if base_url and token:
        try:
            import requests
            headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

            candidates = [
                (f"{base_url.rstrip('/')}/track-trace/shipments", {"bookingReference": booking_number}),
                (f"{base_url.rstrip('/')}/track-trace/shipments/booking-references/{booking_number}", None),
            ]

            def walk(obj):
                if isinstance(obj, dict):
                    yield obj
                    for v in obj.values():
                        yield from walk(v)
                elif isinstance(obj, list):
                    for it in obj:
                        yield from walk(it)

            for url, params in candidates:
                r = requests.get(url, headers=headers, params=params, timeout=30)
                if r.status_code >= 400:
                    log(f"MSC API candidate failed: {url} -> {r.status_code}")
                    continue

                data = r.json()
                arrivals = {}

                containers = set()
                for node in walk(data):
                    for k in ("equipmentReference", "containerNumber", "equipmentReferenceNumber"):
                        if k in node and isinstance(node[k], str):
                            containers.add(_norm_container(node[k]))

                for node in walk(data):
                    equip = None
                    for k in ("equipmentReference", "containerNumber", "equipmentReferenceNumber"):
                        if k in node and isinstance(node[k], str):
                            equip = _norm_container(node[k])
                            break
                    if not equip:
                        continue

                    ts = None
                    for k in ("eventDateTime", "eventCreatedDateTime", "timestamp", "transportEventDateTime"):
                        if k in node and isinstance(node[k], str):
                            ts = node[k]
                            break
                    if not ts:
                        continue

                    desc = " ".join(
                        str(node.get(k, "")).lower()
                        for k in ("eventType", "eventClassifierCode", "eventTypeCode", "transportEventTypeCode", "eventDescription")
                    )

                    if any(w in desc for w in ["arriv", "discharg", "gate in", "available", "unload", "pod"]):
                        prev = arrivals.get(equip)
                        if (not prev) or (str(ts) > str(prev)):
                            arrivals[equip] = ts

                for c in containers:
                    arrivals.setdefault(c, "")

                log("MSC API lookup succeeded.")
                return {k: _parse_date_safe(v) for k, v in arrivals.items() if k}

            log("MSC API configured but no candidates worked; falling back to scraping.")

        except Exception as e:
            log(f"MSC API error; falling back to scraping. {e}")

    # ---- Option B: Playwright scrape (best-effort fallback) ----
    from playwright.sync_api import sync_playwright

    arrivals = {}
    container_rx = re.compile(r"\b([A-Z]{4}\d{7})\b")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://www.msc.com/en/track-a-shipment", wait_until="domcontentloaded", timeout=60000)

        try:
            page.get_by_text("Booking Number", exact=True).click(timeout=20000)
        except Exception:
            log("Could not click 'Booking Number' tab explicitly; continuing anyway.")

        filled = False
        for sel in ["input[type='text']", "input", "textarea"]:
            loc = page.locator(sel)
            for i in range(min(loc.count(), 10)):
                try:
                    if loc.nth(i).is_visible():
                        loc.nth(i).fill(booking_number)
                        filled = True
                        break
                except Exception:
                    continue
            if filled:
                break

        if not filled:
            browser.close()
            log("Failed to find an input box on MSC page.")
            return {}

        try:
            page.keyboard.press("Enter")
        except Exception:
            pass

        page.wait_for_timeout(2500)

        rows = page.locator("table tr")
        if rows.count() > 0:
            for i in range(rows.count()):
                txt = " ".join(rows.nth(i).all_inner_texts()).strip()
                m = container_rx.search(txt)
                if not m:
                    continue
                c = _norm_container(m.group(1))

                date_candidates = re.findall(
                    r"(\b\d{4}-\d{2}-\d{2}\b|\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b|\b\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\b)",
                    txt
                )
                arrivals[c] = _parse_date_safe(date_candidates[-1]) if date_candidates else ""

        if not arrivals:
            body_text = page.inner_text("body")
            containers = sorted(set(_norm_container(m.group(1)) for m in container_rx.finditer(body_text)))
            for c in containers:
                arrivals.setdefault(c, "")

        browser.close()

    log("MSC scrape finished.")
    return arrivals


def apply_arrival_dates_to_pivot(pivot: pd.DataFrame, booking_number: str, log_fn=None) -> pd.DataFrame:
    booking_number = (booking_number or "").strip()
    if not booking_number:
        pivot[("Arrival Date", "")] = ""
        return pivot

    arrivals_map = fetch_msc_arrivals_by_booking(booking_number, log_fn=log_fn)

    def lookup_arrival(transport_val):
        key = _norm_container(str(transport_val))
        return arrivals_map.get(key, "")

    pivot[("Arrival Date", "")] = [lookup_arrival(idx) for idx in pivot.index]
    return pivot


# ---------------- STREAMLIT UI ----------------

st.title("HSTP Packlist → Dimension Summary")
st.write("Upload packing list PDFs to generate a consolidated Excel summary.")

booking_number = st.text_input(
    "Booking Number (optional):",
    value="",
    help="Used to query MSC tracking and append container arrival/ETA dates."
)

uploaded_files = st.file_uploader(
    "Choose PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    summaries = []
    log_lines = []

    def log(msg: str):
        log_lines.append(msg)

    for f in uploaded_files:
        log(f"\nParsing: {f.name}")
        transport, summary, expected, actual = parse_pdf_from_bytes(f.read(), name=f.name)

        if summary is not None:
            summaries.append(summary)
            log(f"✓ {transport}: {len(summary)} rows")
        else:
            log(f"⚠ No data found in {f.name}")

    st.text_area("Log Output", "\n".join(log_lines), height=250)

    if summaries:
        combined = pd.concat(summaries, ignore_index=True)

        pivot = combined.pivot_table(
            index="Transport",
            columns=["Grade", "Dimension"],
            values="Count",
            aggfunc="sum",
            fill_value=0
        )

        pivot = pivot.replace(0, "")

        pivot[("Booking Number", "")] = booking_number.strip() if booking_number.strip() else ""

        try:
            pivot = apply_arrival_dates_to_pivot(pivot, booking_number, log_fn=log)
        except Exception as e:
            log(f"⚠ MSC lookup failed: {e}")
            pivot[("Arrival Date", "")] = ""

        tail_cols = [("Booking Number", ""), ("Arrival Date", "")]
        existing = [c for c in pivot.columns if c not in tail_cols]
        pivot = pivot.reindex(columns=existing + tail_cols)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            pivot.to_excel(writer, sheet_name="All Containers")
        output.seek(0)

        output_name = st.text_input(
            "Output Excel filename:",
            value="_packlist.xlsx",
            help="Enter the filename for the generated Excel summary."
        )

        st.download_button(
            "Download Excel Summary",
            data=output,
            file_name=output_name if output_name.strip() else "HSTP_All.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
