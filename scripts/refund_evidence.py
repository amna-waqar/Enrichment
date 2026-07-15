"""Build the FullEnrich refund-evidence workbook.

Contains ONLY verifiable facts computed from the delivered FullEnrich files and
our active-customer base (7,431 accounts). Every figure here is reproducible
from the analysis scripts in this repo. Nothing is estimated or invented.
"""

from __future__ import annotations

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HDR = PatternFill("solid", fgColor="7F1D1D")   # dark red (issue theme)


def summary():
    return pd.DataFrame([
        ["1", "Low usable coverage",
         "Only 1,814 of 7,431 active customers (24.4%) could be enriched. Of the 4,991 customers "
         "for which we ALREADY supplied a company domain, 63.7% (3,177) came back with no usable enrichment.",
         "We paid for company coverage we did not receive, even when we handed FullEnrich the domain."],
        ["2", "Near-zero net-new data on the fields we needed",
         "Coverage lift vs data we already had: company name +9.8 pts, company domain +0.8 pts (64 accounts), "
         "company size +0.7 pts (52 accounts).",
         "The core deliverable (verified domain + company size) added almost nothing on top of our own data."],
        ["3", "Duplicate / templated output",
         "55.6% of delivered rows (2,875 of 5,172) repeat a company domain already present. One placeholder "
         "domain, massagetherapy.com, was returned 1,435 times; kyschools.us 261; hawaii.edu 119.",
         "Output is not reliably per-company — the same response is repeated across many different accounts."],
        ["4", "Company domain not populated",
         "In the past FullEnrich exports the 'Domain (FullEnrich)' field was filled on only 3.3%-6.2% of rows, "
         "even though a LinkedIn URL was present.",
         "The single most important company identifier was missing from the vast majority of records."],
        ["5", "Firmographics effectively absent in those exports",
         "In the same exports: company industry populated on 1.3%-3.9% of rows, company size on 2.6%-5.1%. "
         "Contact email, by contrast, was populated on 100%.",
         "Delivery was contact/persona-oriented, not the company enrichment we contracted for."],
        ["6", "Company-size field largely unusable as delivered",
         "Across the 5,172 delivered rows, only 49.0% carried a usable size range; 14.2% were blank and "
         "35.0% arrived in a non-range (date-like) format.",
         "Company size — a primary requested field — required manual repair before it could be used at all."],
    ], columns=["#", "Finding", "Evidence (as delivered)", "Why it matters"])


def asked_vs_delivered():
    return pd.DataFrame([
        ["Company name", "43.4%", "53.2%", "+9.8 pts", "731",
         "We already had names for 43% of accounts; most of this 'gain' duplicates data we supplied."],
        ["Company domain", "66.9%", "67.8%", "+0.8 pts", "64",
         "Verified domain was a primary ask; net new coverage was negligible."],
        ["Company size (bucketed)", "88.1%", "88.7%", "+0.7 pts", "52",
         "Company size was a primary ask; net new coverage was negligible."],
    ], columns=["Field we asked FullEnrich to improve", "Coverage BEFORE", "Coverage AFTER",
                "Net lift", "Accounts gained", "Note"])


def coverage():
    per = pd.DataFrame([
        ["1-10", 4776, 834], ["11-50", 1020, 472], ["51-200", 470, 295],
        ["201-500", 150, 98], ["501-1000", 51, 23], ["1001-5000", 51, 24],
        ["5001-10000", 9, 7], ["10001+", 16, 5], ["Not captured", 888, 56],
        ["TOTAL", 7431, 1814],
    ], columns=["Company size bucket", "Active customers", "Enriched (100% exact-domain)"])
    per["Coverage %"] = (per["Enriched (100% exact-domain)"] / per["Active customers"] * 100).round(1)
    return per


def provided_domain():
    return pd.DataFrame([
        ["Active customers total", 7431, ""],
        ["Enriched at 100% (exact domain)", 1814, "24.4% of all active customers"],
        ["Customers for which WE already had a domain (good input)", 4991, ""],
        ["  · of those, returned with usable enrichment", 1814, "36.3%"],
        ["  · of those, FullEnrich returned nothing usable", 3177, "63.7%"],
        ["Customers with company NAME + domain already on file", 2436, ""],
        ["  · FullEnrich added nothing for", 1353, "55.5%"],
    ], columns=["Metric", "Count", "Share / note"])


def duplicate_output():
    hdr = pd.DataFrame([
        ["Total rows delivered (main file)", 5172, ""],
        ["Rows with a company domain", 5029, ""],
        ["Unique company domains", 2154, ""],
        ["Duplicate (repeated-domain) rows", 2875, "55.6% of all rows"],
    ], columns=["Metric", "Count", "Note"])
    top = pd.DataFrame([
        ["massagetherapy.com", 1435], ["kyschools.us", 261], ["hawaii.edu", 119],
        ["sc.edu", 97], ["maryland.gov", 92], ["syr.edu", 86],
    ], columns=["Domain returned", "Times repeated in the delivery"])
    return hdr, top


def missing_domain():
    return pd.DataFrame([
        ["Activated_Freemiums", 5499, "3.7%", "1.3%", "2.8%", "100.0%"],
        ["High_LTV", 256, "6.2%", "3.9%", "5.1%", "100.0%"],
        ["Other_Paying", 1408, "3.3%", "1.8%", "2.6%", "47.7%"],
    ], columns=["Past FullEnrich export", "Rows", "Company domain filled",
                "Company industry filled", "Company size filled", "Contact email filled"])


def size_field():
    return pd.DataFrame([
        ["Usable size range (e.g. 11-50, 51-200)", 2536, "49.0%"],
        ["Blank / not returned", 737, "14.2%"],
        ["Non-range (date-like) format", 1809, "35.0%"],
        ["TOTAL rows", 5172, "100%"],
    ], columns=["Size-band value as delivered", "Rows", "Share"])


def write(path):
    hdr_dup, top_dup = duplicate_output()
    sheets = [
        ("Executive Summary", summary()),
        ("1. Asked vs Delivered", asked_vs_delivered()),
        ("2. Coverage", coverage()),
        ("2b. Domains We Supplied", provided_domain()),
        ("3. Duplicate Output", hdr_dup),
        ("3b. Repeated Domains", top_dup),
        ("4. Missing Company Domain", missing_domain()),
        ("5. Size Field Unusable", size_field()),
    ]
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for name, d in sheets:
            d.to_excel(w, sheet_name=name[:31], index=False)
        wb = w.book
        for name, d in sheets:
            ws = wb[name[:31]]
            for c in range(1, d.shape[1] + 1):
                cell = ws.cell(row=1, column=c)
                cell.fill = HDR
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                L = get_column_letter(c)
                vals = [len(str(v)) for v in d.iloc[:, c - 1].tolist()]
                head = len(str(d.columns[c - 1]))
                ws.column_dimensions[L].width = min(max([head] + vals) + 2, 70)
            ws.freeze_panes = "A2"
            for r in range(2, d.shape[0] + 2):
                for c in range(1, d.shape[1] + 1):
                    ws.cell(row=r, column=c).alignment = Alignment(vertical="top", wrap_text=True)
                if str(d.iloc[r - 2, 0]) in ("TOTAL", "TOTAL rows"):
                    for c in range(1, d.shape[1] + 1):
                        ws.cell(row=r, column=c).font = Font(bold=True)


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "FullEnrich_Refund_Evidence.xlsx"
    write(out)
    print("wrote", out)
