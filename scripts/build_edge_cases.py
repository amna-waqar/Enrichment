"""Split the 'left to enrich' leads into edge cases vs clean, both bucketed.

Edge case = a lead we should NOT auto-enrich because its domain is unreliable
as a company key:
  A) name != domain  — the Chargebee company name does not match the domain and
     is not an abbreviation/acronym of it (e.g. "charles river medical
     associates" on mgb.org = Mass General Brigham).
  B) shared/parent domain — the same domain is used by 2+ genuinely different
     companies (parent org, health system, etc.).
  C) blank company name sitting on a domain that is itself shared or carries a
     name-mismatch elsewhere.

Outputs two workbooks, each bucketed by the original Chargebee size bucket:
  - Edge_Cases.xlsx        (for manual Chargebee review before enriching)
  - Left_To_Enrich.xlsx    (clean list, edge cases removed)
"""
import re
import sys
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import build_icp_sheet as base
import build_final_workbook as bf
import match_enrichment as me

D = "/root/.claude/uploads/815bd82b-022c-5eff-8466-360a03104051/"
FE = D + "15cd01b6-Chargebee_Company_Enrichment_.xlsx"
EXTRA = [("Other_Paying", D + "69ada757-Copy_of_Other_Paying.xlsx"),
         ("High_LTV", D + "b3d7526a-Copy_of_High_LTV.xlsx"),
         ("Activated_Freemiums", D + "09a09842-Copy_of_Activated_Freemiums.xlsx")]
ORIG = ["customer_id", "company", "company_from_domain", "domain", "domain_from_email",
        "email", "first_name", "last_name", "job_title", "state", "country", "zip",
        "active_mrr", "active_arr", "plans", "currency", "csm", "origin", "partner"]

TWO_LEVEL = {"co.uk", "org.uk", "ac.uk", "gov.uk", "com.au", "org.au", "net.au", "co.nz",
             "com.mx", "co.il", "org.il", "com.br", "co.za"}
STOP = {"the", "and", "of", "for", "a", "an", "at", "in", "on", "to", "by", "dba", "inc",
        "llc", "ltd", "corp", "co", "company", "pllc", "pc", "llp", "lp", "plc", "group",
        "services", "service", "associates", "assoc", "corporation", "limited"}


def sld(domain):
    d = me.norm_domain(domain)
    if not d:
        return ""
    p = d.split(".")
    if len(p) >= 3 and ".".join(p[-2:]) in TWO_LEVEL:
        return p[-3]
    return p[-2] if len(p) >= 2 else p[0]


def toks(name):
    name = re.sub(r"[^a-z0-9\s]", " ", str(name).lower())
    return [t for t in name.split() if t and t not in STOP and len(t) > 1]


def name_matches_domain(company, domain):
    """True/False if assessable, None if company blank."""
    s = re.sub(r"[^a-z0-9]", "", sld(domain).lower())
    t = toks(company)
    if not s or not t:
        return None
    compact = "".join(t)
    acr = "".join(x[0] for x in t)
    if s in compact or compact in s:
        return True
    if len(acr) >= 2 and (s == acr or s.startswith(acr) or acr in s or acr.startswith(s)):
        return True
    # acronym-of-domain-tokens: first letters of the domain label match nothing here,
    # but company acronym as a prefix of the domain label is a common abbreviation form
    if len(acr) >= 3 and s.startswith(acr[:3]):
        return True
    for x in t:
        if len(x) >= 4 and (x in s or s in x):
            return True
        # common compression: first 4+ chars of a company word appear in the domain
        # label (e.g. "bettendorf" -> "bettpeds", "michigan" -> "mich...")
        if len(x) >= 5 and x[:4] in s:
            return True
    if len(t[0]) >= 5 and t[0][:5] in s:
        return True
    return False


def _blank(x):
    return x is None or (isinstance(x, float) and x != x) or str(x).strip().lower() in ("", "nan")


def _distinct_company_groups(names):
    """Count genuinely-different company names (collapse substring variants)."""
    ns = sorted({re.sub(r"[^a-z0-9 ]", "", n.lower()).strip() for n in names if n}, key=len)
    groups = []
    for n in ns:
        if any(n in g or g in n for g in groups):
            continue
        groups.append(n)
    return len(groups)


def classify(en):
    def ud(r):
        for c in ("domain", "domain_from_email"):
            d = me.norm_domain(r.get(c))
            if d and d not in me.FREE_DOMAINS:
                return d
        return ""
    en = en.copy()
    en["_ud"] = en.apply(ud, axis=1)

    # domain-level facts across ALL active customers
    names_by_dom, count_by_dom, mism_by_dom = {}, {}, {}
    for _, r in en.iterrows():
        d = r["_ud"]
        if not d:
            continue
        count_by_dom[d] = count_by_dom.get(d, 0) + 1
        c = r.get("company")
        if not _blank(c):
            names_by_dom.setdefault(d, set()).add(str(c).strip())
            if name_matches_domain(c, d) is False:
                mism_by_dom[d] = True
    groups_by_dom = {d: _distinct_company_groups(ns) for d, ns in names_by_dom.items()}

    def reasons(r):
        d = r["_ud"]
        if not d:
            return []
        out = []
        c = r.get("company")
        if not _blank(c) and name_matches_domain(c, d) is False:
            out.append("name doesn't match domain")
        if groups_by_dom.get(d, 0) >= 2:
            out.append("shared domain (2+ different companies)")
        if _blank(c) and count_by_dom.get(d, 0) >= 2 and (mism_by_dom.get(d) or groups_by_dom.get(d, 0) >= 2):
            out.append("blank name on a shared/mismatched domain")
        return out

    en["_reasons"] = en.apply(reasons, axis=1)
    en["edge_reason"] = en["_reasons"].map(lambda r: "; ".join(r))
    en["_edge"] = en["_reasons"].map(bool)
    return en


HDR = PatternFill("solid", fgColor="7F1D1D")
HDR2 = PatternFill("solid", fgColor="1F3864")


def _write(sheets, out_path, hdr):
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        for name, d in sheets:
            d.to_excel(w, sheet_name=name[:31], index=False)
        wb = w.book
        for name, d in sheets:
            ws = wb[name[:31]]
            for c in range(1, d.shape[1] + 1):
                cell = ws.cell(row=1, column=c)
                cell.fill = hdr
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="left", vertical="center")
                L = get_column_letter(c)
                vals = [len(str(v)) for v in d.iloc[:, c - 1].tolist()[:400]]
                ws.column_dimensions[L].width = min(max([len(str(d.columns[c - 1]))] + vals) + 2, 48)
                if "%" in str(d.columns[c - 1]):
                    fmt = '0.0"%"'
                elif base._money_col(str(d.columns[c - 1])):
                    fmt = '$#,##0.00'
                else:
                    fmt = None
                if fmt:
                    for r in range(2, d.shape[0] + 2):
                        ws.cell(row=r, column=c).number_format = fmt
            ws.freeze_panes = "A2"
            if d.iloc[:, 0].astype(str).eq("TOTAL").any():
                i = d.index[d.iloc[:, 0].astype(str) == "TOTAL"][0]
                for c in range(1, d.shape[1] + 1):
                    ws.cell(row=i + 2, column=c).font = Font(bold=True)


def build(edge_path, left_path):
    df = base.load(D)
    dom_best, name_unique, dom_band, dom_source, _ = bf.build_multi_index(FE, EXTRA)
    en = bf.enrich(df, dom_best, name_unique, dom_band, dom_source)
    un = en[en["enr_status"] == "Not enriched"].copy()
    un = classify(un)

    edge = un[un["_edge"]].copy()
    clean = un[~un["_edge"]].copy()

    # ---- summaries ----
    def summary(sub, label):
        rows = []
        for seg in base.SIZE_ORDER:
            s = sub[sub["size_bucket"] == seg]
            rows.append({"Company size bucket": seg, label: len(s),
                         "MRR": round(s["active_mrr"].sum(), 2)})
        out = pd.DataFrame(rows)
        out.loc[len(out)] = {"Company size bucket": "TOTAL", label: len(sub),
                             "MRR": round(sub["active_mrr"].sum(), 2)}
        return out

    # ---- edge workbook ----
    esheets = [("Summary — Edge Cases", summary(edge, "Edge cases"))]
    for seg in base.SIZE_ORDER:
        s = edge[edge["size_bucket"] == seg]
        if s.empty:
            continue
        tab = s[ORIG + ["edge_reason"]].sort_values("active_mrr", ascending=False).reset_index(drop=True)
        esheets.append(("Not Captured" if seg == "Not captured" else seg, tab))
    _write(esheets, edge_path, HDR)

    # ---- clean left-to-enrich workbook ----
    lsheets = [("Summary — Left to Enrich", summary(clean, "Left to enrich"))]
    for seg in base.SIZE_ORDER:
        s = clean[clean["size_bucket"] == seg]
        if s.empty:
            continue
        tab = s[ORIG].sort_values("active_mrr", ascending=False).reset_index(drop=True)
        lsheets.append(("Not Captured" if seg == "Not captured" else seg, tab))
    _write(lsheets, left_path, HDR2)

    assert len(edge) + len(clean) == len(un)
    print(f"Not-enriched: {len(un)}  ->  edge cases {len(edge)} | clean left-to-enrich {len(clean)}")
    print("\nEdge cases by bucket:")
    print(edge["size_bucket"].value_counts().reindex(base.SIZE_ORDER).fillna(0).astype(int).to_string())
    print("\nEdge reason counts:")
    from collections import Counter
    cc = Counter()
    for rs in edge["_reasons"]:
        for x in rs:
            cc[x] += 1
    for k, v in cc.most_common():
        print(f"  {k}: {v}")
    print(f"\nwrote {edge_path} and {left_path}")


if __name__ == "__main__":
    build("/home/user/Enrichment/Edge_Cases.xlsx", "/home/user/Enrichment/Left_To_Enrich.xlsx")
