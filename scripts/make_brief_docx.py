"""Render the FullEnrich refund meeting brief as a formatted .docx."""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

RED = RGBColor(0x7F, 0x1D, 0x1D)
GREY = RGBColor(0x55, 0x55, 0x55)


def h(doc, text, size=13, color=RED, space_before=10):
    p = doc.add_paragraph()
    p.space_before = Pt(space_before)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return p


def para(doc, text, italic=False, color=None, size=10.5):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.italic = italic
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return p


def bullet(doc, label, text):
    p = doc.add_paragraph(style="List Bullet")
    if label:
        r = p.add_run(label + " ")
        r.bold = True
        r.font.size = Pt(10.5)
    r2 = p.add_run(text)
    r2.font.size = Pt(10.5)


def build(path):
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10.5)
    for m in ("top", "bottom", "left", "right"):
        setattr(doc.sections[0], f"{m}_margin", Inches(0.7))

    title = doc.add_paragraph()
    r = title.add_run("Meeting Brief — FullEnrich (Floris): Refund Discussion")
    r.bold = True
    r.font.size = Pt(17)
    r.font.color.rgb = RED

    for k, v in [
        ("Prepared for", "internal alignment + the call with Floris"),
        ("Subject", "Company-enrichment engagement on our active customer base (7,431 accounts)"),
        ("Objective", "Recover the spend — target: full ~$3K refund; floor: partial credit — on the grounds that "
         "the delivery did not meet the company-enrichment need and carried material data-quality issues."),
    ]:
        p = doc.add_paragraph()
        rb = p.add_run(f"{k}: ")
        rb.bold = True
        rb.font.size = Pt(10.5)
        p.add_run(v).font.size = Pt(10.5)

    # 1
    h(doc, "1. What we asked for")
    para(doc, "Enrich our accounts with three company-level fields we could act on: verified company domain, "
              "company size, and firmographics (industry / revenue / HQ) — so we can segment and divvy up "
              "accounts by org size.")

    # 2
    h(doc, "2. What we actually got (the headline)")
    para(doc, "For the three fields that mattered, the net-new coverage on top of data we already held was negligible:")
    t = doc.add_table(rows=1, cols=4)
    t.style = "Light Grid Accent 1"
    for i, c in enumerate(["Field", "Before", "After", "Net lift"]):
        run = t.rows[0].cells[i].paragraphs[0].add_run(c)
        run.bold = True
    for row in [
        ("Company name", "43.4%", "53.2%", "+9.8 pts (we already supplied names)"),
        ("Company domain", "66.9%", "67.8%", "+0.8 pts (64 accounts)"),
        ("Company size", "88.1%", "88.7%", "+0.7 pts (52 accounts)"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].paragraphs[0].add_run(val)
    para(doc, "")
    para(doc, "Only 24.4% of the active base (1,814 / 7,431) could be enriched at all. Most damning: of the 4,991 "
              "accounts where we already handed FullEnrich a company domain, 63.7% (3,177) came back with nothing "
              "usable. When they had the identifier, they still couldn't return the company.")

    # 3
    h(doc, "3. Data-integrity issues (loss of confidence)")
    bullet(doc, "Duplicated / templated output:", "55.6% of delivered rows (2,875 / 5,172) repeat a domain. One "
           "placeholder — massagetherapy.com — was returned 1,435 times; kyschools.us 261×, hawaii.edu 119×. The "
           "same answer is repeated across many different companies.")
    bullet(doc, "Company domain not populated:", "in the past FullEnrich exports the company-domain field was filled "
           "on only 3–6% of rows despite a LinkedIn URL being present.")
    bullet(doc, "Firmographics effectively absent", "in those exports: industry 1–4%, size 3–5% populated — while "
           "contact email was 100%. Delivery was contact/persona-oriented, not company enrichment.")
    bullet(doc, "Company-size field unusable as delivered:", "only 49% of rows carried a usable size range; 14% blank, "
           "35% in a non-range format requiring manual repair.")

    # 4
    h(doc, "4. The core argument — we are not their ICP")
    para(doc, "FullEnrich is LinkedIn / persona-first. If a company isn't well-represented on LinkedIn, it isn't "
              "enriched — and much of our SMB/healthcare long-tail isn't. That is a methodology fit problem, not a "
              "one-off miss: paying for future volume (or the remaining credit) makes no sense because their database "
              "doesn't cover the companies our database is made of.")

    # 5
    h(doc, "5. The ask")
    bullet(doc, "1.", "Full refund of the ~$3K (floor: negotiate to ~$2,500 / credit-back), on the basis of "
           "(a) coverage far below a usable threshold even on domains we supplied, (b) the integrity defects above, "
           "and (c) demonstrated ICP mismatch.")
    bullet(doc, "2.", "If refused: we do not intend to spend the remaining credit — redeploying it elsewhere is a "
           "better outcome for both sides than a second failed run.")

    # internal
    doc.add_paragraph()
    h(doc, "Internal prep — anticipate Floris's counters (DO NOT hand this section over)", size=12, color=GREY)
    bullet(doc, "“I ran it and got 80%.”", "Different denominator. Our verified exact-domain match against our active "
           "customers is 24.4%; on the domains we supplied, 63.7% returned nothing usable. Ask him to define his 80% "
           "and reconcile to account-level coverage.")
    bullet(doc, "“I delivered 25% on LinkedIn / industry.”", "Agreed — when it matched, per-record firmographics were "
           "fine (industry ~91%, revenue ~85% on the matched subset). Do NOT argue per-record quality — it's a losing "
           "thread. Keep the argument on coverage + duplicates + domain-not-populated + ICP fit.")
    bullet(doc, "“It's persona-based, not company-based.”", "That is exactly our point — we contracted for company "
           "enrichment; a persona-first tool is the wrong fit, which is why we're not their ICP.")
    bullet(doc, "Handle with care:", "the “size delivered as dates (35%)” line may partly be an Excel/export artifact "
           "on our side. Lead with the duplicates and domain-not-populated defects (unambiguous); use the size-field "
           "point as support, not the spearhead.")
    bullet(doc, "Keep off the table:", "the internal re-bucketing work and the messy HubSpot/Clay numbers — they don't "
           "help this negotiation and open side-debates.")

    doc.add_paragraph()
    p = para(doc, "Supporting evidence: see FullEnrich_Refund_Evidence.xlsx — every figure above is on a tab, computed "
                  "directly from the delivered files (reproducible, no estimates).", italic=True, color=GREY, size=9.5)

    doc.save(path)


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "FullEnrich_Meeting_Brief.docx"
    build(out)
    print("wrote", out)
