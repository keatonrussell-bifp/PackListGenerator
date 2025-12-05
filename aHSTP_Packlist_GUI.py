#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HSTP Packing List -> Dim Summary (GUI)
"""

import re
import math
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
try:
    import pdfplumber
except Exception:
    raise SystemExit(
        "Missing dependency 'pdfplumber'. Install with:\n"
        "  pip install pdfplumber pandas openpyxl xlsxwriter"
    )

# ---------- Regex & helpers ----------

TRANSPORT_RX = re.compile(r"Transport\s*:\s*([A-Z0-9]+)")

# Detect section headers like:
#   Sawn timber, Taeda-Pine, 3COM, dried, planed BLI
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


def inches_to_feet_label(inches: float) -> str:
    return f"{int(round(inches / 12.0))}'"


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


def parse_pdf(pdf_path: Path):
    with pdfplumber.open(pdf_path) as pdf:
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

            # --- Detect a new section grade dynamically ---
            g = SECTION_GRADE_RX.search(ln)
            if g:
                cur_grade = g.group(1).strip().upper()

            # --- Parse imperial lines ---
            mi = ROW_RX_IMPERIAL.search(ln)
            if mi:
                items.append({
                    "Transport": transport or pdf_path.stem,
                    "Thickness (in)": float(mi.group("thick").replace(",", ".")),
                    "Width (in)":     float(mi.group("width").replace(",", ".")),
                    "Length (in)":    frac_to_float(mi.group("length")),
                    "Quantity (pcs)": int(mi.group("qty")),
                    "Volume (mbf)":   float(mi.group("mbf").replace(",", ".")),
                    "Grade":          cur_grade or "UNKNOWN",
                })
                continue

            # --- Parse metric fence lines ---
            mm = ROW_RX_METRIC.search(ln)
            if mm:
                items.append({
                    "Transport":      transport or pdf_path.stem,
                    "Thickness (in)": float(mm.group("thick_mm").replace(",", ".").replace(" ", "")) * MM_TO_IN,
                    "Width (in)":     float(mm.group("width_mm").replace(",", ".").replace(" ", "")) * MM_TO_IN,
                    "Length (in)":    float(mm.group("length_m").replace(",", ".").replace(" ", "")) * M_TO_IN,
                    "Quantity (pcs)": int(mm.group("qty")),
                    "Volume (mbf)":   float('nan'),
                    "Grade":          cur_grade or "UNKNOWN",
                })
                continue

    if not transport:
        transport = pdf_path.stem

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


# ---------- GUI ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HSTP Packlist → Dim Summary")

        self.update_idletasks()
        w, h = 900, 750
        ws, hs = self.winfo_screenwidth(), self.winfo_screenheight()
        x, y = (ws // 2) - (w // 2), (hs // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(820, 600)

        self.pdf_paths = []
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm_select = ttk.LabelFrame(self, text="1) Select packing list PDFs")
        frm_select.pack(fill="x", **pad)
        ttk.Button(frm_select, text="Add PDFs…", command=self.add_pdfs).grid(row=0, column=0, sticky="w", padx=8, pady=8)
        ttk.Button(frm_select, text="Clear",    command=self.clear_pdfs).grid(row=0, column=1, sticky="w", padx=8, pady=8)

        self.lst = tk.Listbox(frm_select, height=8)
        self.lst.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=8, pady=(0,8))
        frm_select.columnconfigure(2, weight=1)
        frm_select.rowconfigure(1,  weight=1)

        frm_out = ttk.LabelFrame(self, text="2) Choose output Excel file")
        frm_out.pack(fill="x", **pad)
        self.out_var = tk.StringVar(value=str(Path.cwd() / "HSTP_All.xlsx"))
        ttk.Entry(frm_out, textvariable=self.out_var).grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        frm_out.columnconfigure(0, weight=1)
        ttk.Button(frm_out, text="Browse…", command=self.choose_out).grid(row=0, column=1, padx=8, pady=8)

        frm_run = ttk.LabelFrame(self, text="3) Run")
        frm_run.pack(fill="both", expand=True, **pad)
        self.txt = tk.Text(frm_run, height=16)
        self.txt.pack(fill="both", expand=True, padx=8, pady=8)

        self.btn_run = ttk.Button(self, text="Create Excel", command=self.run_clicked)
        self.btn_run.pack(pady=(0,12))
        self.lbl_status = ttk.Label(self, text="Ready.")
        self.lbl_status.pack(pady=(0,8))

    def add_pdfs(self):
        paths = filedialog.askopenfilenames(
            title="Select packing list PDFs",
            filetypes=[("PDF files","*.pdf"), ("All files","*.*")]
        )
        if paths:
            for p in paths:
                pp = Path(p)
                if pp not in self.pdf_paths:
                    self.pdf_paths.append(pp)
                    self.lst.insert("end", str(pp))

    def clear_pdfs(self):
        self.pdf_paths.clear()
        self.lst.delete(0, "end")

    def choose_out(self):
        p = filedialog.asksaveasfilename(
            title="Save Excel as…", defaultextension=".xlsx",
            filetypes=[("Excel workbook","*.xlsx")], initialfile="HSTP_All.xlsx"
        )
        if p:
            self.out_var.set(p)

    def log(self, msg: str):
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.update_idletasks()

    def _validate(self, pdf_name, expected, actual):
        def fmt(v):
            return "—" if v is None or (isinstance(v, float) and pd.isna(v)) else f"{v}"
        ok_pcs = (expected.get("pcs") is None) or (expected["pcs"] == actual.get("pcs"))
        ok_pkg = (expected.get("pkgs") is None) or (expected["pkgs"] == actual.get("pkgs"))
        ok_mbf = (expected.get("mbf") is None) or (abs((actual.get("mbf") or 0.0) - expected["mbf"]) < 1e-3)
        status = "PASS" if (ok_pcs and ok_pkg and ok_mbf) else "MISMATCH"
        self.log(f"  Validation [{status}] for {pdf_name}:")
        self.log(f"    Expected → pcs={fmt(expected.get('pcs'))}, mbf={fmt(expected.get('mbf'))}, pkgs={fmt(expected.get('pkgs'))}")
        self.log(f"    Actual   → pcs={fmt(actual.get('pcs'))},  mbf={fmt(actual.get('mbf'))},  pkgs={fmt(actual.get('pkgs'))}")
        if status != "PASS":
            self.log("    ⚠ Check the PDF totals text or parsing rules for this file.")

    def run_clicked(self):
        if not self.pdf_paths:
            messagebox.showwarning("No PDFs", "Please add at least one PDF.")
            return

        out_path = Path(self.out_var.get()).expanduser()
        out_path.parent.mkdir(parents=True, exist_ok=True)

        self.btn_run.config(state="disabled")
        self.lbl_status.config(text="Processing…")

        def worker():
            try:
                parts = []
                for pdf in self.pdf_paths:
                    self.log(f"Parsing: {pdf.name}")
                    try:
                        transport, summary, expected, actual = parse_pdf(pdf)
                        if summary is not None and not summary.empty:
                            parts.append(summary)
                            self.log(f"  ✓ {transport}: {len(summary)} rows")
                            self._validate(pdf.name, expected, actual)
                        else:
                            self.log(f"  ⚠ No data found in {pdf.name}")
                    except Exception as e:
                        self.log(f"  ✗ Failed: {e}")

                if not parts:
                    self.log("No valid data parsed.")
                    messagebox.showwarning("No data", "No valid data extracted from the selected PDFs.")
                    return

                combined = pd.concat(parts, ignore_index=True)

                pivot = combined.pivot_table(
                    index="Transport",
                    columns=["Grade", "Dimension"],
                    values="Count",
                    aggfunc="sum",
                    fill_value=0
                )

                # Sorting grades alphabetically except FENCE last
                if isinstance(pivot.columns, pd.MultiIndex):
                    grades = sorted({g for g, _ in pivot.columns if g != "FENCE"})
                    if "FENCE" in {g for g, _ in pivot.columns}:
                        grades.append("FENCE")  # send FENCE to end

                    def sort_key(col):
                        grade, dim = col
                        return (grades.index(grade), dim)

                    pivot = pivot.reindex(
                        sorted(pivot.columns, key=sort_key),
                        axis=1
                    )
                    pivot.columns.names = ["Grade", "Dimension"]

                pivot.index.name = None
                pivot = pivot.replace(0, "")

                with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
                    pivot.to_excel(writer, sheet_name="All Containers")
                    ws = writer.sheets["All Containers"]
                    ws.freeze_panes(2, 1)

                self.log(f"\n✅ Done. Wrote: {out_path}")
                self.lbl_status.config(text=f"Done → {out_path}")
                messagebox.showinfo("Success", f"Created:\n{out_path}")
            finally:
                self.btn_run.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
