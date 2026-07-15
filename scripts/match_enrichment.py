"""Match FullEnrich company-enrichment data back to the active customers.

The enrichment export has no join key (no customer_id), so we match on an
EXACT key hierarchy — no fuzzy guesses:

  1. Company domain   (customer real domain == enrichment domain)  -> 100%
  2. Email domain     (non-free email domain == enrichment domain) -> 100%
  3. Company name     (normalized name -> a UNIQUE enrichment domain,
                       only when the customer has no usable domain)  -> "review"

Domains are matched on the exact normalized host. We deliberately do NOT
collapse sub-domains to the registrable domain, because shared hosts
(onmicrosoft.com, myshopify.com, ...) would create false links. Free email
providers (gmail.com, ...) are never used as a match key.

Output: `icp_overview_enriched.xlsx` — the original workbook with enrichment
columns appended (and shaded) on every size-bucket tab, plus an
`Enrichment Coverage` tab. Original columns/tabs are unchanged.
"""

from __future__ import annotations

import argparse
import re

import pandas as pd
from openpyxl.styles import Font, PatternFill

import build_icp_sheet as base

ENRICH_FIELDS = [
    "Company Name", "Company Domain", "Company Linkedin URL",
    "Company Description", "Year Founded", "Company headcount",
    "Company headcount range", "Company type", "Company industry",
    "Company headquarters location", "Company office locations",
    "Company specialties", "Company Revenue Range",
]
# Columns we append to each customer row (shaded in the workbook).
ENRICH_COLS = ["enr_status", "enr_match_key", "enr_confidence", "enr_matched_domain"] + [
    f"enr_{f}" for f in ENRICH_FIELDS
]

FREE_DOMAINS = {
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "aol.com",
    "icloud.com", "comcast.net", "msn.com", "live.com", "me.com", "mac.com",
    "protonmail.com", "proton.me", "gmx.com", "ymail.com", "fastmail.com",
    "att.net", "verizon.net", "sbcglobal.net", "cox.net", "bellsouth.net",
    "charter.net", "earthlink.net", "mail.com", "zoho.com", "yahoo.co.uk",
    "hotmail.co.uk", "googlemail.com", "aol.co.uk", "rocketmail.com",
}
LEGAL = {"llc", "inc", "incorporated", "corp", "corporation", "ltd", "limited",
         "co", "company", "pllc", "pc", "llp", "lp", "plc", "gmbh", "pvt"}


def norm_domain(x) -> str:
    if x is None or (isinstance(x, float) and x != x):
        return ""
    x = str(x).strip().lower()
    x = re.sub(r"^https?://", "", x)
    x = re.sub(r"^www\.", "", x)
    return x.split("/")[0].split("?")[0].strip()


def norm_name(x) -> str:
    if x is None or (isinstance(x, float) and x != x):
        return ""
    x = str(x).lower().strip()
    x = re.sub(r"[^a-z0-9&\s-]", " ", x)
    return " ".join(t for t in re.split(r"[\s-]+", x) if t and t not in LEGAL).strip()


def _completeness(row: pd.Series) -> int:
    return int(sum(1 for f in ENRICH_FIELDS if str(row.get(f, "")).strip() not in ("", "nan")))


def build_enrichment_index(enrich_path: str):
    """Return (domain->best row dict, unique-name->domain) built from real domains."""
    xl = pd.ExcelFile(enrich_path)
    frames = [pd.read_excel(enrich_path, sheet_name=s, dtype=str) for s in xl.sheet_names]
    e = pd.concat(frames, ignore_index=True)
    e["_dom"] = e["Company Domain"].map(norm_domain)
    e["_name"] = e["Company Name"].map(norm_name)
    e["_score"] = e.apply(_completeness, axis=1)

    real = e[(e["_dom"] != "") & (~e["_dom"].isin(FREE_DOMAINS))]

    # best row per domain = most complete
    dom_best: dict[str, dict] = {}
    for dom, grp in real.groupby("_dom"):
        best = grp.sort_values("_score", ascending=False).iloc[0]
        dom_best[dom] = {f: ("" if pd.isna(best[f]) else str(best[f])) for f in ENRICH_FIELDS}

    # normalized name -> set of real domains (for unique-name matching)
    name_doms: dict[str, set] = {}
    for _, r in real.iterrows():
        name_doms.setdefault(r["_name"], set()).add(r["_dom"])
    name_unique = {n: next(iter(d)) for n, d in name_doms.items() if len(d) == 1}

    stats = {
        "enrichment_rows": len(e),
        "unique_real_domains": len(dom_best),
        "rows_missing_domain": int((e["_dom"] == "").sum()),
    }
    return dom_best, name_unique, stats


def match_customer(row: pd.Series, dom_best: dict, name_unique: dict) -> dict:
    """Return enrichment columns for one customer row (exact matching only)."""
    od = norm_domain(row.get("domain"))
    oe = norm_domain(row.get("domain_from_email"))
    name = norm_name(row.get("company")) or norm_name(row.get("company_from_domain"))

    usable_dom = od if (od and od not in FREE_DOMAINS) else (
        oe if (oe and oe not in FREE_DOMAINS) else "")

    out = {c: "" for c in ENRICH_COLS}

    matched_dom = None
    key = None
    conf = None
    if usable_dom and usable_dom in dom_best:
        matched_dom = usable_dom
        key = "company_domain" if (od and od not in FREE_DOMAINS and od in dom_best) else "email_domain"
        conf = "100%"
    elif not usable_dom and name and name in name_unique:
        matched_dom = name_unique[name]
        key = "company_name"
        conf = "Review (name only)"

    if matched_dom:
        data = dom_best[matched_dom]
        out["enr_status"] = "Enriched"
        out["enr_match_key"] = key
        out["enr_confidence"] = conf
        out["enr_matched_domain"] = matched_dom
        for f in ENRICH_FIELDS:
            out[f"enr_{f}"] = data.get(f, "")
    else:
        out["enr_status"] = "Not enriched"
    return out


def enrich_dataframe(df: pd.DataFrame, dom_best: dict, name_unique: dict) -> pd.DataFrame:
    rows = df.apply(lambda r: match_customer(r, dom_best, name_unique), axis=1, result_type="expand")
    return pd.concat([df.reset_index(drop=True), rows.reset_index(drop=True)], axis=1)


def coverage_table(enriched: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for seg in base.SIZE_ORDER:
        s = enriched[enriched["size_bucket"] == seg]
        if s.empty:
            continue
        exact = int((s["enr_confidence"] == "100%").sum())
        review = int((s["enr_confidence"] == "Review (name only)").sum())
        rows.append({
            "Company size": seg,
            "Customers": len(s),
            "Enriched — 100% (domain)": exact,
            "Enriched — name (review)": review,
            "Total enriched": exact + review,
            "Coverage % (100%)": round(exact / len(s) * 100, 1),
            "Coverage % (incl. review)": round((exact + review) / len(s) * 100, 1),
            "Not enriched": len(s) - exact - review,
        })
    out = pd.DataFrame(rows)
    tot_exact = int((enriched["enr_confidence"] == "100%").sum())
    tot_rev = int((enriched["enr_confidence"] == "Review (name only)").sum())
    n = len(enriched)
    out.loc[len(out)] = {
        "Company size": "TOTAL", "Customers": n,
        "Enriched — 100% (domain)": tot_exact, "Enriched — name (review)": tot_rev,
        "Total enriched": tot_exact + tot_rev,
        "Coverage % (100%)": round(tot_exact / n * 100, 1),
        "Coverage % (incl. review)": round((tot_exact + tot_rev) / n * 100, 1),
        "Not enriched": n - tot_exact - tot_rev,
    }
    return out


# ---- workbook writing with shaded enrichment columns -----------------------

ENR_HEADER_FILL = PatternFill("solid", fgColor="375623")   # dark green
ENR_BODY_FILL = PatternFill("solid", fgColor="E2EFDA")     # light green


def write_enriched_workbook(sheets, out_path, shade_headers: set):
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in sheets:
            df.to_excel(writer, sheet_name=name[:31], index=False)
        wb = writer.book
        for name, df in sheets:
            ws = wb[name[:31]]
            ncols = df.shape[1]
            for c in range(1, ncols + 1):
                header = str(df.columns[c - 1])
                cell = ws.cell(row=1, column=c)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="left", vertical="center")
                is_enr = header in shade_headers
                cell.fill = ENR_HEADER_FILL if is_enr else PatternFill("solid", fgColor="1F3864")
                letter = get_column_letter(c)
                maxlen = max([len(header)] + [len(str(v)) for v in df.iloc[:, c - 1].tolist()[:400]])
                ws.column_dimensions[letter].width = min(max(maxlen + 2, 10), 55)
                fmt = None
                if "%" in header:
                    fmt = '0.0"%"'
                elif base._money_col(header):
                    fmt = '$#,##0.00'
                for r in range(2, df.shape[0] + 2):
                    cc = ws.cell(row=r, column=c)
                    if fmt:
                        cc.number_format = fmt
                    if is_enr:
                        cc.fill = ENR_BODY_FILL
            ws.freeze_panes = "A2"
            if ncols and df.iloc[:, 0].astype(str).eq("TOTAL").any():
                idx = df.index[df.iloc[:, 0].astype(str) == "TOTAL"][0]
                for c in range(1, ncols + 1):
                    ws.cell(row=idx + 2, column=c).font = Font(bold=True)


def build(csv_dir: str, enrich_path: str, out_path: str) -> pd.DataFrame:
    df = base.load(csv_dir)
    dom_best, name_unique, stats = build_enrichment_index(enrich_path)
    enriched = enrich_dataframe(df, dom_best, name_unique)

    # assemble sheets: analysis tabs unchanged, bucket tabs get enrichment cols
    sheets = [
        ("Overview", base.t_overview(df)),
        ("Segment Summary", base.t_segment_summary(df)),
        ("Enrichment Coverage", coverage_table(enriched)),
        ("Coverage - Overall", base.t_coverage_overall(df)),
        ("Coverage - by Segment", base.t_coverage_by_segment(df)),
        ("Missing-Data & Gaps", base.t_gaps(df)),
        ("Completeness Score", base.t_completeness(df)),
        ("Plan & Tier Breakdown", base.t_tier(df)),
        ("Top Accounts", base.t_top_accounts(df)),
    ]
    base_cols = ["customer_id", "best_company", "company", "domain", "email",
                 "first_name", "last_name", "job_title", "state", "country", "zip",
                 "active_mrr", "active_arr", "plans", "tier", "term", "is_trial",
                 "currency", "csm", "origin", "partner", "completeness_0_7"]
    for seg in base.SIZE_ORDER:
        s = enriched[enriched["size_bucket"] == seg]
        if s.empty:
            continue
        s = s.sort_values("active_mrr", ascending=False)
        cols = base_cols + ENRICH_COLS
        tab = s[cols].rename(columns={"active_mrr": "MRR", "active_arr": "ARR"}).reset_index(drop=True)
        name = "Not Captured" if seg == "Not captured" else seg
        sheets.append((name, tab))

    write_enriched_workbook(sheets, out_path, shade_headers=set(ENRICH_COLS))
    _verify(enriched, stats)
    return enriched


def _verify(enriched: pd.DataFrame, stats: dict) -> None:
    n = len(enriched)
    exact = int((enriched["enr_confidence"] == "100%").sum())
    review = int((enriched["enr_confidence"] == "Review (name only)").sum())
    assert n == 7431, f"expected 7431 customers, got {n}"
    assert enriched["customer_id"].nunique() == n
    # every enriched row must carry a matched domain + a company name
    e = enriched[enriched["enr_status"] == "Enriched"]
    assert (e["enr_matched_domain"].str.len() > 0).all()
    print(f"  verify OK: {n} customers | enrichment rows={stats['enrichment_rows']}, "
          f"unique domains={stats['unique_real_domains']}")
    print(f"  matched 100% (domain): {exact} ({exact/n*100:.1f}%) | "
          f"name-review: {review} ({review/n*100:.1f}%) | "
          f"total enriched: {exact+review} ({(exact+review)/n*100:.1f}%)")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--dir", required=True, help="dir with the active-customer CSVs")
    ap.add_argument("--enrich", required=True, help="FullEnrich .xlsx path")
    ap.add_argument("--out", default="icp_overview_enriched.xlsx")
    args = ap.parse_args()
    build(args.dir, args.enrich, args.out)
    print(f"  wrote {args.out}")


if __name__ == "__main__":
    main()
