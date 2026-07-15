"""v2 workbook: re-bucket enriched companies into their TRUE size segment and
report data coverage BEFORE vs AFTER enrichment.

Builds on match_enrichment.py. Two additions the user asked for:

  1. Re-bucketing — when a customer is matched to enrichment and the enriched
     LinkedIn size band is known, move the customer to that size bucket (e.g. a
     "1-10" account FullEnrich says is 51-200 moves to 51-200). The original
     bucket is preserved in a column for audit; nothing is deleted.

  2. Before/after coverage — a dedicated tab comparing field fill-rates using
     only Chargebee data (before) vs Chargebee + FullEnrich (after).

Data-quality note: the enrichment "Company headcount range" column is corrupted
by Excel (the "1-10" band was auto-converted to the date 2026-02-10). We decode
it back; the numeric "Company headcount" column is noisy and NOT used for
bucketing. Blank/ambiguous bands are never trusted — the customer keeps its
original bucket.
"""

from __future__ import annotations

import argparse
import re

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import build_icp_sheet as base
import match_enrichment as me

NUM_BUCKETS = ["1-10", "11-50", "51-200", "201-500", "501-1000",
               "1001-5000", "5001-10000", "10001+"]
CLEAN_BANDS = {"11-50", "51-200", "201-500", "501-1000", "1001-5000", "5001-10000", "10001+"}
META_COLS = ["original_bucket", "bucket_source", "moved"]


def decode_band(v) -> str | None:
    """Decode FullEnrich size band to one of our buckets; None if untrusted."""
    if v is None or (isinstance(v, float) and v != v):
        return None
    s = str(v).strip()
    if s == "" or s.lower() == "nan":
        return None
    if s in CLEAN_BANDS:
        return s
    if s in ("1", "1-10", "2-10"):
        return "1-10"
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):   # Excel-mangled "1-10"/"2-10" -> date
        return "1-10"
    return None


def build_index(enrich_path: str):
    """domain -> best enrichment row (range decoded), unique-name -> domain,
    domain -> decoded size band."""
    xl = pd.ExcelFile(enrich_path)
    e = pd.concat([pd.read_excel(enrich_path, sheet_name=s, dtype=str) for s in xl.sheet_names],
                  ignore_index=True)
    e["_dom"] = e["Company Domain"].map(me.norm_domain)
    e["_name"] = e["Company Name"].map(me.norm_name)
    e["_score"] = e.apply(me._completeness, axis=1)
    e["_band"] = e["Company headcount range"].map(decode_band)

    real = e[(e["_dom"] != "") & (~e["_dom"].isin(me.FREE_DOMAINS))]
    dom_best, dom_band = {}, {}
    for dom, grp in real.groupby("_dom"):
        best = grp.sort_values("_score", ascending=False).iloc[0]
        rec = {f: ("" if pd.isna(best[f]) else str(best[f])) for f in me.ENRICH_FIELDS}
        rec["Company headcount range"] = best["_band"] or ""   # show decoded band, not date
        dom_best[dom] = rec
        dom_band[dom] = best["_band"]

    name_doms: dict[str, set] = {}
    for _, r in real.iterrows():
        name_doms.setdefault(r["_name"], set()).add(r["_dom"])
    name_unique = {n: next(iter(d)) for n, d in name_doms.items() if len(d) == 1}

    stats = {"enrichment_rows": len(e), "unique_real_domains": len(dom_best)}
    return dom_best, name_unique, dom_band, stats


def enrich_and_rebucket(df: pd.DataFrame, dom_best, name_unique, dom_band) -> pd.DataFrame:
    recs = []
    for _, row in df.iterrows():
        out = me.match_customer(row, dom_best, name_unique)
        orig = row["size_bucket"]
        matched_dom = out.get("enr_matched_domain") or ""
        band = dom_band.get(matched_dom) if matched_dom else None
        # Only re-bucket on 100%-confidence (exact domain) matches — never on
        # lower-confidence name-only matches. No errors allowed.
        if out.get("enr_confidence") != "100%":
            band = None
        if band in NUM_BUCKETS:
            final = band
            source = "FullEnrich" if final != orig else "Chargebee (confirmed)"
        else:
            final = orig
            source = "Chargebee (original)"
        out["original_bucket"] = orig
        out["size_bucket_final"] = final
        out["bucket_source"] = source
        out["moved"] = "Yes" if final != orig else ""
        recs.append(out)
    add = pd.DataFrame(recs)
    return pd.concat([df.reset_index(drop=True), add.reset_index(drop=True)], axis=1)


# ---- coverage tables -------------------------------------------------------

def _present(s: pd.Series) -> pd.Series:
    return ~(s.isna() | (s.astype(str).str.strip().isin(["", "nan"])))


def coverage_before_after(en: pd.DataFrame) -> pd.DataFrame:
    n = len(en)
    def enr_present(col):
        return _present(en[col]) if col in en else pd.Series(False, index=en.index)

    company_b = _present(en["company"])
    company_a = company_b | enr_present("enr_Company Name")
    domain_b = _present(en["domain"])
    domain_a = domain_b | enr_present("enr_Company Domain")
    size_b = en["original_bucket"] != "Not captured"
    size_a = en["size_bucket_final"] != "Not captured"
    geo_b = _present(en["state"]) | _present(en["country"])
    geo_a = geo_b | enr_present("enr_Company headquarters location")
    zero = pd.Series(False, index=en.index)
    rows = [
        ("Company name", company_b, company_a),
        ("Company domain", domain_b, domain_a),
        ("Company size (bucketed)", size_b, size_a),
        ("Geography (HQ / state-country)", geo_b, geo_a),
        ("Industry", zero, enr_present("enr_Company industry")),
        ("Employee headcount", zero, enr_present("enr_Company headcount")),
        ("Revenue range", zero, enr_present("enr_Company Revenue Range")),
        ("LinkedIn URL", zero, enr_present("enr_Company Linkedin URL")),
        ("Year founded", zero, enr_present("enr_Year Founded")),
        ("HQ location", zero, enr_present("enr_Company headquarters location")),
    ]
    out = []
    for label, b, a in rows:
        bp, ap = int(b.sum()), int(a.sum())
        out.append({
            "Field": label,
            "Present before": bp, "Coverage before %": round(bp / n * 100, 1),
            "Present after": ap, "Coverage after %": round(ap / n * 100, 1),
            "Δ accounts": ap - bp, "Δ %": round((ap - bp) / n * 100, 1),
        })
    return pd.DataFrame(out)


def enrichment_coverage(en: pd.DataFrame) -> pd.DataFrame:
    """Match coverage per FINAL bucket, with original counts for movement."""
    orig_counts = en["original_bucket"].value_counts()
    rows = []
    for seg in base.SIZE_ORDER:
        s = en[en["size_bucket_final"] == seg]
        if s.empty and seg not in orig_counts:
            continue
        exact = int((s["enr_confidence"] == "100%").sum())
        review = int((s["enr_confidence"] == "Review (name only)").sum())
        fin = len(s)
        rows.append({
            "Company size": seg,
            "Accounts (final)": fin,
            "Accounts (original)": int(orig_counts.get(seg, 0)),
            "Net move": fin - int(orig_counts.get(seg, 0)),
            "Enriched — 100%": exact,
            "Enriched — review": review,
            "Total enriched": exact + review,
            "Coverage % (100%)": round(exact / fin * 100, 1) if fin else 0.0,
            "Coverage % (incl. review)": round((exact + review) / fin * 100, 1) if fin else 0.0,
        })
    out = pd.DataFrame(rows)
    ex = int((en["enr_confidence"] == "100%").sum())
    rv = int((en["enr_confidence"] == "Review (name only)").sum())
    n = len(en)
    out.loc[len(out)] = {
        "Company size": "TOTAL", "Accounts (final)": n, "Accounts (original)": n, "Net move": 0,
        "Enriched — 100%": ex, "Enriched — review": rv, "Total enriched": ex + rv,
        "Coverage % (100%)": round(ex / n * 100, 1),
        "Coverage % (incl. review)": round((ex + rv) / n * 100, 1),
    }
    return out


def rebucket_log(en: pd.DataFrame) -> pd.DataFrame:
    m = en[en["moved"] == "Yes"].copy()
    m = m.sort_values(["original_bucket", "size_bucket_final", "active_mrr"], ascending=[True, True, False])
    cols = {"customer_id": "customer_id", "best_company": "company",
            "enr_matched_domain": "matched_domain", "original_bucket": "from_bucket",
            "size_bucket_final": "to_bucket", "enr_match_key": "match_key",
            "active_mrr": "MRR"}
    return m[list(cols)].rename(columns=cols).reset_index(drop=True)


# ---- workbook --------------------------------------------------------------

ENR_HDR = PatternFill("solid", fgColor="375623"); ENR_BODY = PatternFill("solid", fgColor="E2EFDA")
META_HDR = PatternFill("solid", fgColor="806000"); META_BODY = PatternFill("solid", fgColor="FFF2CC")
BLUE = PatternFill("solid", fgColor="1F3864")


def write_wb(sheets, out_path, green: set, amber: set):
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in sheets:
            df.to_excel(writer, sheet_name=name[:31], index=False)
        wb = writer.book
        for name, df in sheets:
            ws = wb[name[:31]]
            for c in range(1, df.shape[1] + 1):
                header = str(df.columns[c - 1])
                cell = ws.cell(row=1, column=c)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="left", vertical="center")
                g, a = header in green, header in amber
                cell.fill = ENR_HDR if g else META_HDR if a else BLUE
                letter = get_column_letter(c)
                maxlen = max([len(header)] + [len(str(v)) for v in df.iloc[:, c - 1].tolist()[:400]])
                ws.column_dimensions[letter].width = min(max(maxlen + 2, 10), 55)
                fmt = '0.0"%"' if "%" in header else ('$#,##0.00' if base._money_col(header) else None)
                for r in range(2, df.shape[0] + 2):
                    cc = ws.cell(row=r, column=c)
                    if fmt:
                        cc.number_format = fmt
                    if g:
                        cc.fill = ENR_BODY
                    elif a:
                        cc.fill = META_BODY
            ws.freeze_panes = "A2"
            if df.shape[1] and df.iloc[:, 0].astype(str).eq("TOTAL").any():
                idx = df.index[df.iloc[:, 0].astype(str) == "TOTAL"][0]
                for c in range(1, df.shape[1] + 1):
                    ws.cell(row=idx + 2, column=c).font = Font(bold=True)


def overview(df, en, stats, moved, size_b, size_a):
    n = len(en)
    ex = int((en["enr_confidence"] == "100%").sum())
    rv = int((en["enr_confidence"] == "Review (name only)").sum())
    rows = [
        ("Report", "iFax — Active Customers by Company Size (Enriched + Re-bucketed)"),
        ("Total active accounts", f"{n:,}"),
        ("Enrichment file rows / unique domains", f"{stats['enrichment_rows']:,} / {stats['unique_real_domains']:,}"),
        ("Companies matched — 100% (exact domain)", f"{ex:,} ({ex/n*100:.1f}%)"),
        ("Companies matched — name (review)", f"{rv:,} ({rv/n*100:.1f}%)"),
        ("Total enriched", f"{ex+rv:,} ({(ex+rv)/n*100:.1f}%)"),
        ("Accounts re-bucketed (size corrected)", f"{moved:,}"),
        ("Size coverage BEFORE enrichment", f"{size_b:,} ({size_b/n*100:.1f}%)"),
        ("Size coverage AFTER enrichment", f"{size_a:,} ({size_a/n*100:.1f}%)"),
        ("", ""),
        ("NOTES", ""),
        ("Bucketing", "Enriched accounts are placed in the size bucket from FullEnrich's LinkedIn "
                      "band; unmatched accounts keep their Chargebee bucket. 'original_bucket' is retained."),
        ("Data fix", "FullEnrich 'headcount range' was Excel-corrupted (1-10 band shown as the date "
                     "2026-02-10); decoded back to 1-10. Noisy numeric headcount not used for bucketing."),
        ("Confidence", "Only exact domain/email-domain matches are 100%. Name-only matches are flagged "
                       "'review' and excluded from the 100% figure. No fuzzy guesses."),
    ]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


def build(csv_dir, enrich_path, out_path):
    df = base.load(csv_dir)
    dom_best, name_unique, dom_band, stats = build_index(enrich_path)
    en = enrich_and_rebucket(df, dom_best, name_unique, dom_band)

    # df keyed on FINAL bucket for the bucket-partitioned analysis tabs
    df_final = en.copy()
    df_final["size_bucket"] = en["size_bucket_final"]

    size_b = int((en["original_bucket"] != "Not captured").sum())
    size_a = int((en["size_bucket_final"] != "Not captured").sum())
    moved = int((en["moved"] == "Yes").sum())

    sheets = [
        ("Overview", overview(df, en, stats, moved, size_b, size_a)),
        ("Coverage Before vs After", coverage_before_after(en)),
        ("Enrichment Coverage", enrichment_coverage(en)),
        ("Re-bucketing Log", rebucket_log(en)),
        ("Segment Summary", base.t_segment_summary(df_final)),
        ("Coverage - Overall", base.t_coverage_overall(df)),
        ("Coverage - by Segment", base.t_coverage_by_segment(df_final)),
        ("Missing-Data & Gaps", base.t_gaps(df_final)),
        ("Completeness Score", base.t_completeness(df_final)),
        ("Plan & Tier Breakdown", base.t_tier(df_final)),
        ("Top Accounts", base.t_top_accounts(df_final)),
    ]
    base_cols = ["customer_id", "best_company", "company", "domain", "email",
                 "first_name", "last_name", "job_title", "state", "country", "zip",
                 "active_mrr", "active_arr", "plans", "tier", "term", "is_trial",
                 "currency", "csm", "origin", "partner", "completeness_0_7"]
    for seg in base.SIZE_ORDER:
        s = en[en["size_bucket_final"] == seg]
        if s.empty:
            continue
        s = s.sort_values("active_mrr", ascending=False)
        cols = base_cols + me.ENRICH_COLS + META_COLS
        tab = s[cols].rename(columns={"active_mrr": "MRR", "active_arr": "ARR"}).reset_index(drop=True)
        sheets.append(("Not Captured" if seg == "Not captured" else seg, tab))

    write_wb(sheets, out_path, green=set(me.ENRICH_COLS), amber=set(META_COLS))
    _verify(en, df, moved, size_b, size_a)
    return en


def _verify(en, df, moved, size_b, size_a):
    n = len(en)
    assert n == 7431 and en["customer_id"].nunique() == n, "row/id count changed"
    # no rows lost across re-bucketing
    assert int(en["size_bucket_final"].value_counts().sum()) == n
    # every re-bucketed row is enriched and lands in a numeric bucket
    mv = en[en["moved"] == "Yes"]
    assert (mv["enr_status"] == "Enriched").all(), "moved a non-enriched row"
    assert mv["size_bucket_final"].isin(NUM_BUCKETS).all(), "moved into a non-size bucket"
    assert size_a >= size_b, "size coverage decreased"
    ex = int((en["enr_confidence"] == "100%").sum())
    print(f"  verify OK: {n} accounts, no row loss")
    print(f"  matched 100%: {ex} ({ex/n*100:.1f}%) | re-bucketed: {moved}")
    print(f"  size coverage: before {size_b} ({size_b/n*100:.1f}%) -> after {size_a} ({size_a/n*100:.1f}%)")
    # show net movement per bucket
    orig = df["size_bucket"].value_counts()
    fin = en["size_bucket_final"].value_counts()
    print("  bucket   original -> final (net):")
    for b in base.SIZE_ORDER:
        o, f = int(orig.get(b, 0)), int(fin.get(b, 0))
        print(f"    {b:12s} {o:>5} -> {f:>5} ({f-o:+d})")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--dir", required=True)
    ap.add_argument("--enrich", required=True)
    ap.add_argument("--out", default="icp_overview_enriched_rebucketed.xlsx")
    a = ap.parse_args()
    build(a.dir, a.enrich, a.out)
    print(f"  wrote {a.out}")


if __name__ == "__main__":
    main()
