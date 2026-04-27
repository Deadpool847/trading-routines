"""Generate the daily trading-intelligence .docx brief."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-04-27"
BIAS = "BEAR"
ENTRY_PERMISSION = "CAUTION"

OUT_PATH = f"trading-briefs/2026/04-April/Trading_Brief_{DATE_ISO}_{BIAS}.docx"


def shade_cell(cell, hex_fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)


def add_heading(doc, text, level=1, color=None):
    h = doc.add_heading(text, level=level)
    if color:
        for run in h.runs:
            run.font.color.rgb = color
    return h


def add_para(doc, text, bold=False, italic=False, size=11, color=None, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return p


def add_bullets(doc, items):
    for item in items:
        doc.add_paragraph(item, style="List Bullet")


def add_kv_table(doc, rows, header_fill="1F4E79", header_color=RGBColor(0xFF, 0xFF, 0xFF)):
    table = doc.add_table(rows=len(rows), cols=2)
    table.style = "Light Grid Accent 1"
    for i, (k, v) in enumerate(rows):
        c0, c1 = table.rows[i].cells
        c0.text = k
        c1.text = v
        for run in c0.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(10.5)
        for run in c1.paragraphs[0].runs:
            run.font.size = Pt(10.5)
    return table


def add_stock_card(doc, stock):
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    shade_cell(cell, "FFF4E6")
    p_title = cell.paragraphs[0]
    r = p_title.add_run(f"{stock['name']}  ({stock['ticker']})")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

    fields = [
        ("CMP", stock["cmp"]),
        ("Catalyst", stock["catalyst"]),
        ("Crude Alignment", stock["crude"]),
        ("Option Watch", stock["option"]),
        ("Entry Trigger", stock["entry"]),
        ("Target", stock["target"]),
        ("Stop Loss", stock["stop"]),
        ("Confidence", stock["confidence"]),
    ]
    for k, v in fields:
        p = cell.add_paragraph()
        rk = p.add_run(f"{k}: ")
        rk.bold = True
        rk.font.size = Pt(10.5)
        rv = p.add_run(v)
        rv.font.size = Pt(10.5)
    doc.add_paragraph()


def build_doc():
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    # ====== HEADER ======
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("ELITE DAILY TRADING INTELLIGENCE BRIEF")
    r.bold = True
    r.font.size = Pt(20)
    r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rs = sub.add_run("Indian F&O Day Trader — Nifty + Stocks")
    rs.italic = True
    rs.font.size = Pt(11)
    rs.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # Header KPI table
    hdr = doc.add_table(rows=1, cols=3)
    hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c1, c2, c3 = hdr.rows[0].cells
    c1.text = f"DATE\n{DATE_ISO} (Mon)"
    c2.text = f"BIAS\n{BIAS}"
    c3.text = f"ENTRY\n{ENTRY_PERMISSION}"
    shade_cell(c1, "1F4E79")
    shade_cell(c2, "C0392B")
    shade_cell(c3, "E67E22")
    for c in (c1, c2, c3):
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    doc.add_paragraph()

    # ====== SECTION 1: MACRO SNAPSHOT ======
    add_heading(doc, "Section 1 — Macro Snapshot", level=1,
                color=RGBColor(0x1F, 0x4E, 0x79))
    macro_rows = [
        ("MCX Crude (May)", "~₹9,191 (Fri close, -0.77%); gap-up risk on Sunday Brent +2%"),
        ("Brent / WTI", "$107.89 / $96.63 (+2% Sunday, 3-week high)"),
        ("US S&P 500", "7,165.08 (+0.80%) — record"),
        ("US Nasdaq", "24,836.60 (+1.63%) — record"),
        ("US Dow", "49,230.71 (-0.16%)"),
        ("US VIX", "18.71 (-3.11%) — normal"),
        ("CNN Fear & Greed", "67 — Greed"),
        ("Nikkei / Hang Seng / Kospi", "+0.53% / +0.24% / +1.0%"),
        ("Gift Nifty", "~23,954 (-180 / -0.75%) → modest GAP DOWN"),
        ("DXY / USD-INR", "97.62 / 94.11"),
        ("Today's Event Risk", "HIGH (FOMC Wed + Iran headlines + Q4 wave)"),
    ]
    add_kv_table(doc, macro_rows)
    doc.add_paragraph()

    # ====== SECTION 2: CRUDE RULE APPLIED ======
    add_heading(doc, "Section 2 — Crude Rule Applied",
                level=1, color=RGBColor(0xC0, 0x39, 0x2B))
    add_para(doc,
             "TODAY'S CRUDE RULE MODE: BEAR — PUTS ONLY",
             bold=True, size=13, color=RGBColor(0xC0, 0x39, 0x2B))
    add_bullets(doc, [
        "MCX Crude at ₹9,191 sits in the ₹9,000–₹10,000 BEAR band.",
        "Brent +2% Sunday to $107.89 after Trump cancelled the Witkoff/Kushner Islamabad mission.",
        "Strait of Hormuz remains effectively closed for the third straight month; Iran ruled out 'forced negotiations'.",
        "Implication: long-only crude beneficiaries (ONGC, Oil India) okay; downstream OMC, paints, aviation = avoid longs.",
        "Critical flip level: a sustained move BELOW ₹9,000 resets bias to neutral and re-allows directional calls.",
    ])
    doc.add_paragraph()

    # ====== SECTION 3: INDIAN MARKET INTERNALS ======
    add_heading(doc, "Section 3 — Indian Market Internals",
                level=1, color=RGBColor(0x1F, 0x4E, 0x79))
    internals = [
        ("Nifty Friday Close", "23,897.95 (-275 / -1.14%) — 3rd red session"),
        ("Key Support", "23,800 (S1) | 23,555–23,500 (S2)"),
        ("Key Resistance", "24,200 (R1) | 24,350–24,400 (R2)"),
        ("Trend (50 DMA)", "Just broke below — bearish short-term"),
        ("India VIX", "19.7 (+6%) — ELEVATED FEAR (half-size rule)"),
        ("FII (23-Apr cash)", "Net SELL ₹3,254.71 cr"),
        ("DII (23-Apr cash)", "Net BUY ₹941.35 cr"),
        ("FII 5-day trend", "Consistent SELLING"),
        ("Nifty PCR", "~0.67 — bearish (call writers dominant)"),
        ("Max OI Call (resistance)", "24,000 / 24,200"),
        ("Max OI Put (support)", "23,800"),
        ("Max Pain", "~23,900"),
        ("Weekly/Monthly Expiry", "Thursday 30-Apr (monthly)"),
        ("Bank Nifty", "55,956.85 (-0.62%) — slightly weaker than Nifty"),
        ("Sector Leaders", "Metals, FMCG, Upstream Oil PSU (ONGC/OIL)"),
        ("Sector Laggards", "IT (-5.29%), Aviation, Paints, Downstream OMC"),
    ]
    add_kv_table(doc, internals)
    doc.add_paragraph()

    # ====== SECTION 4: STOCK OPPORTUNITIES ======
    add_heading(doc, "Section 4 — Stock Opportunities (F&O)",
                level=1, color=RGBColor(0xC0, 0x39, 0x2B))
    add_para(doc,
             "Two clean setups; honest call — skipping a third weak setup. "
             "Trade only when full pre-conditions align.",
             italic=True, size=10.5, color=RGBColor(0x55, 0x55, 0x55))

    add_stock_card(doc, {
        "name": "INFOSYS",
        "ticker": "NSE: INFY",
        "cmp": "~₹1,395 (after 5%+ Friday slide)",
        "catalyst": "Cautious FY27 guidance dragged the entire IT pack -5.29% Friday; weekend risk-off + weak rupee extends downside.",
        "crude": "YES — BEAR-mode aligned (PUT permitted).",
        "option": "INFY 1380 PE / 1400 PE (30-Apr monthly expiry).",
        "entry": "Sustained break below ₹1,390 on 5-min close + Nifty rejecting 23,950.",
        "target": "₹1,360 (≈ -2.5%).",
        "stop": "Close above ₹1,415 OR Nifty reclaims 24,050.",
        "confidence": "MEDIUM (oversold bounce risk).",
    })

    add_stock_card(doc, {
        "name": "ONGC",
        "ticker": "NSE: ONGC",
        "cmp": "~₹290–295",
        "catalyst": "Brent $107+ and Hormuz shut → upstream realisation tailwind. ONGC already +14% YTD; Sunday's crude jump reignites the trade.",
        "crude": "YES — anti-correlation long permitted under BEAR-crude rule (only crude-beneficiary call allowed).",
        "option": "ONGC 300 CE (30-Apr expiry).",
        "entry": "ONGC opens flat/green AND MCX crude holds above ₹9,300 by 9:45 AM.",
        "target": "₹305 (≈ +3%).",
        "stop": "Close below ₹287 OR crude flips below ₹9,000.",
        "confidence": "HIGH.",
    })
    doc.add_paragraph()

    # ====== SECTION 5: DOMESTIC MACRO ======
    add_heading(doc, "Section 5 — Domestic Macro & Triggers",
                level=1, color=RGBColor(0x1F, 0x4E, 0x79))
    add_bullets(doc, [
        "RBI: Held repo at 5.25% on 8-Apr meeting; neutral stance; FY27 CPI projection 4.6%. No fresh action today.",
        "CPI (March): 3.4% YoY — below 4% target; oil-driven import inflation rising as a watch-item.",
        "Earnings done: RIL Q4 PAT -12.55% YoY to ₹16,971 cr; ₹6 dividend; stock -0.81% to ₹1,332. Mild negative drag at open.",
        "Earnings ahead this week: 28-Apr — Coal India, UltraTech Cement, Varun Beverages. Later: Maruti, Adani Ports, Kotak Bank.",
        "US FOMC: 28–29 Apr — 99.7% odds of HOLD at 3.50–3.75% (Polymarket). Fed-week chop expected.",
        "SEBI / NSE: No fresh F&O regulatory action overnight.",
    ])
    doc.add_paragraph()

    # ====== SECTION 6: FINAL VERDICT ======
    add_heading(doc, "Section 6 — Final Verdict",
                level=1, color=RGBColor(0xC0, 0x39, 0x2B))
    verdict = [
        ("Crude Rule Mode", "BEAR — PUTS ONLY (MCX ₹9,191, rising)"),
        ("Market Bias", "BEAR (3rd red session, IT collapse, Iran talks failed Sun)"),
        ("VIX Sizing Rule", "HALF SIZE — VIX 19.7 (elevated)"),
        ("Key Support", "23,800"),
        ("Key Resistance", "24,200"),
        ("Critical Crude Flip Level", "₹9,000 — below this, bias resets to neutral"),
        ("Top 2 Risk Events",
         "1) Surprise Iran-US backchannel news during cash hours; "
         "2) Fed Wednesday + RIL Q4 digestion"),
        ("ENTRY PERMISSION", "CAUTION"),
    ]
    add_kv_table(doc, verdict)

    add_para(doc, "")
    one_liner = doc.add_paragraph()
    one_liner.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = one_liner.add_run("ONE-LINE SUMMARY: ")
    r1.bold = True
    r1.font.size = Pt(11.5)
    r1.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    r2 = one_liner.add_run(
        "Today is a BEAR day. Crude at ₹9,191 = BEAR mode. "
        "Watch INFOSYS for sub-₹1,390 puts and ONGC for crude-tailwind calls. "
        "Key risk: Fed Wednesday + Iran headline whip. "
        "Size HALF based on VIX at 19.7."
    )
    r2.font.size = Pt(11.5)

    # ====== FOOTER ======
    doc.add_paragraph()
    foot = doc.add_paragraph()
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rf = foot.add_run(
        "Generated by Daily Trading Routine — for personal educational use only. "
        "Not financial advice."
    )
    rf.italic = True
    rf.font.size = Pt(9)
    rf.font.color.rgb = RGBColor(0x77, 0x77, 0x77)

    doc.save(OUT_PATH)
    print(f"Saved: {OUT_PATH}")


if __name__ == "__main__":
    build_doc()
