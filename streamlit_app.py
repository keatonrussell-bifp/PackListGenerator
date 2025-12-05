import streamlit as st
import pandas as pd
import io
from pathlib import Path
import pdfplumber
import re
import math

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


# ---------------- STREAMLIT UI ----------------

st.title("HSTP Packlist → Dimension Summary")
st.write("Upload packing list PDFs to generate a consolidated Excel summary.")

uploaded_files = st.file_uploader(
    "Choose PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    summaries = []
    log_output = ""

    for f in uploaded_files:
        log_output += f"\nParsing: {f.name}\n"
        transport, summary, expected, actual = parse_pdf_from_bytes(f.read(), name=f.name)

        if summary is not None:
            summaries.append(summary)
            log_output += f"✓ {transport}: {len(summary)} rows\n"
        else:
            log_output += f"⚠ No data found in {f.name}\n"

    st.text_area("Log Output", log_output, height=250)

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

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            pivot.to_excel(writer, sheet_name="All Containers")
        output.seek(0)

        # --- User-defined filename ---
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
