"""Generate the daily trading-intelligence .docx brief."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-04-23"
BIAS = "WAIT"
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


def add_kv_table(doc, rows, header=None, shade_header=True):
    ncols = len(rows[0])
    table = doc.add_table(rows=(1 if header else 0) + len(rows), cols=ncols)
    table.style = "Light Grid Accent 1"
    r_idx = 0
    if header:
        hdr = table.rows[0]
        for i, txt in enumerate(header):
            cell = hdr.cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(txt)
            run.bold = True
            run.font.size = Pt(11)
            if shade_header:
                shade_cell(cell, "1F3864")
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r_idx = 1
    for i, row in enumerate(rows):
        trow = table.rows[r_idx + i]
        for j, val in enumerate(row):
            cell = trow.cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(11)
            if j == 0:
                run.bold = True
    return table


def badge_line(doc, label, value, fill_hex, text_color=RGBColor(0xFF, 0xFF, 0xFF)):
    table = doc.add_table(rows=1, cols=1)
    table.autofit = True
    cell = table.rows[0].cells[0]
    cell.text = ""
    shade_cell(cell, fill_hex)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{label}: {value}")
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = text_color


# ============== BUILD DOCUMENT ==============
doc = Document()

# Tighten page margins
for section in doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

# ---------- HEADER ----------
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("ELITE DAILY TRADING INTELLIGENCE BRIEF")
run.bold = True
run.font.size = Pt(20)
run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_run = sub.add_run("Indian F&O Day Trader — Nifty + Stocks")
sub_run.italic = True
sub_run.font.size = Pt(12)

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
dr = date_p.add_run(f"Date: Thursday, 23 April 2026  |  Generated 08:45 IST")
dr.bold = True
dr.font.size = Pt(11)

# Bias + permission badges
badge_line(doc, "TODAY'S BIAS", BIAS, "F1C232")           # amber for WAIT
badge_line(doc, "ENTRY PERMISSION", ENTRY_PERMISSION, "E69138")  # orange for CAUTION
badge_line(doc, "CRUDE RULE MODE", "MILD BULL (caution tilt)", "6AA84F")

doc.add_paragraph()

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Reading", "Interpretation"],
    rows=[
        ["MCX Crude (₹/bbl)", "~8,390", "Mild Bull zone (7,500–8,500); fragile"],
        ["Brent / WTI", "$98.20 / $89.22", "Off highs after Iran ceasefire extended"],
        ["India VIX", "17.53 (-6.71%)", "Elevated but cooling — half-size trades"],
        ["US VIX (CBOE)", "18.92 (-2.97%)", "Calm-side of fear; calmest since March"],
        ["S&P 500", "7,137.90 (+1.05%)", "Record high — TAILWIND"],
        ["Nasdaq", "24,657.57 (+1.64%)", "Record high — TAILWIND"],
        ["Dow Jones", "49,490.03 (+0.69%)", "Strong close"],
        ["Fear & Greed (US)", "70 — Greed", "Near Extreme Greed threshold"],
        ["Nikkei 225", "59,585.86 (+0.4%)", "Record high — bullish Asia"],
        ["Hang Seng", "26,163.24 (-1.22%)", "Weak — drag on Asia sentiment"],
        ["Gift Nifty", "24,304.5 (-59.5)", "Flat-to-slight gap DOWN vs 24,378 close"],
        ["DXY", "98.57 (+0.18%)", "Below 104 — moderate rupee pressure"],
        ["USD/INR", "93.80", "Rupee weak (-9.71% YoY) — FII caution"],
    ])

# ---------- SECTION 2 — CRUDE RULE APPLIED ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule Applied", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "MODE: MILD BULL — with a CAUTION tilt.", bold=True, size=12)
add_bullets(doc, [
    "MCX crude sits at ~₹8,390, inside the ₹7,500–₹8,500 band that normally permits full-size calls.",
    "BUT crude hit ₹10,888 earlier this month during the Iran war spike; the band is statistically fragile.",
    "Trump extended the US-Iran ceasefire indefinitely on 22 Apr. Brent fell below $98, WTI below $89.",
    "Strait of Hormuz remains partially halted; a gunboat attack on a Liberia-flagged vessel earlier in week.",
    "Vitol CEO pegs total war-driven supply loss at 600–700 million barrels so far (potentially 1 billion).",
    "Critical level: ₹8,500 MCX. Break above = downgrade to half-size. Break of ₹9,000 = flip to puts-only.",
    "Tail-risk: any single Iran-related headline can re-spike crude 5–7% intraday.",
])

add_para(doc, "Crude-aligned sector bias:", bold=True, size=11)
add_bullets(doc, [
    "Bullish: ONGC, Oil India (upstream — $1/bbl adds ~₹6,180 cr to ONGC earnings annually)",
    "Bearish: BPCL, HPCL, IOC (OMCs squeezed), IndiGo & aviation (+$10 Brent = +₹3,500–4,000 cr fuel bill)",
])

# ---------- SECTION 3 — INDIAN MARKET INTERNALS ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Indian Market Internals", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Nifty Technical Levels", bold=True, size=12)
add_kv_table(doc,
    header=["Level", "Value", "Meaning"],
    rows=[
        ["Previous Close", "24,378", "Closed -199 pts (-0.81%); broke 24,400"],
        ["Expected Open", "~24,320–24,340", "Mild gap down per Gift Nifty"],
        ["Support S1", "24,322", "Intraday low zone from 22 Apr"],
        ["Support S2", "24,000", "Psychological; breach = 23,800 target"],
        ["Resistance R1", "24,450–24,500", "Strong cap for three sessions"],
        ["Resistance R2", "24,600", "Bullish reclaim level"],
        ["200 DMA", "23,527", "Well above — long-term trend BULLISH"],
        ["50 DMA", "~24,446", "Testing from below — bearish short-term"],
    ])

doc.add_paragraph()
add_para(doc, "Volatility & Flows", bold=True, size=12)
add_bullets(doc, [
    "India VIX: 17.53 (-6.71%) — ELEVATED zone (17–22). Rule: reduce size, widen stops.",
    "FII (22 Apr): Net SELL ₹2,078 cr — bearish flow signal.",
    "DII (22 Apr): Net SELL ₹1,048 cr — unusual, normally DII absorbs. Both-sell = cautious.",
    "Rupee at 93.80 reinforces FII caution; DXY at 98.57 still benign (<104).",
])

doc.add_paragraph()
add_para(doc, "Options Positioning (28 Apr expiry — no weekly expiry today; NSE moved to Tuesday)", bold=True, size=12)
add_bullets(doc, [
    "Max Pain (28 Apr): 24,400 — gravitational pull level",
    "Highest Call OI: 24,500 CE (OI ~7.8 M) — acts as ceiling/resistance",
    "Highest Put OI: concentrated 24,000 — support/floor zone",
    "PCR: reported neutral-to-slightly-bearish; call writers have upper hand near 24,500",
    "Interpretation: smart money expects 24,000–24,500 range-bound drift into Tuesday expiry",
])

doc.add_paragraph()
add_para(doc, "Sector Rotation", bold=True, size=12)
add_bullets(doc, [
    "Leaders (YTD): PSU banks +29%, Metals +27%, Autos +22%, Private banks +15%, Infra +12%",
    "Laggards (YTD): IT -12%, Pharma -4%, Energy -3% (downstream drag)",
    "Yesterday: Bank Nifty outperformed (~56,565). IT & Pharma soft (Wipro -2.83%, Sun Pharma -1.04%)",
    "Today's tilt: PSU banks + upstream oil (ONGC, Oil India) favoured. Avoid aviation, OMCs.",
])

# ---------- SECTION 4 — STOCK OPPORTUNITIES ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — Stock Opportunities", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc,
    "Two catalyst-driven setups qualify today. Everything else is noise — two strong picks beat five weak ones.",
    italic=True, size=10)

# Card helper
def stock_card(doc, rows, border_hex="1F3864"):
    t = doc.add_table(rows=len(rows), cols=2)
    t.style = "Light Grid Accent 1"
    for i, (k, v) in enumerate(rows):
        lc = t.rows[i].cells[0]
        rc = t.rows[i].cells[1]
        lc.text = ""
        rc.text = ""
        p1 = lc.paragraphs[0]
        r1 = p1.add_run(k)
        r1.bold = True
        r1.font.size = Pt(11)
        shade_cell(lc, "D9E1F2")
        p2 = rc.paragraphs[0]
        r2 = p2.add_run(v)
        r2.font.size = Pt(11)

add_para(doc, "STOCK 1 — INFOSYS (INFY)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP", "₹1,268.80 (down -3.38% on 22 Apr into results)"),
    ("Catalyst", "Q4 FY26 results TODAY, 3:45 PM IST — press 4:30 PM, call 5:30 PM"),
    ("Crude Alignment", "NEUTRAL (IT is crude-agnostic; but rupee weakness is a tailwind for exporters)"),
    ("Consensus", "PAT ~₹7,508 cr (+4% YoY); Revenue ~₹46,567 cr (+13.7% YoY); seq PAT -1.5%"),
    ("Option Watch", "INFY 1260 CE / 1260 PE — 24 Apr Fri weekly OR May monthly straddle"),
    ("Entry Trigger", "POST-RESULT play only. Do NOT pre-position. After 3:45 PM wait for 5-min confirmation: a close above ₹1,290 = long call; close below ₹1,230 = long put."),
    ("Target", "Beat: 1,320 / 1,350. Miss: 1,210 / 1,180"),
    ("Stop Loss", "Post-entry, 5-min candle reversal beyond trigger bar"),
    ("Confidence", "HIGH on volatility; MEDIUM on direction"),
])

doc.add_paragraph()
add_para(doc, "STOCK 2 — OIL INDIA (OIL)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP", "Large-cap PSU upstream (refer live quote)"),
    ("Catalyst", "Geopolitical crude premium intact despite ceasefire; Strait of Hormuz still impaired"),
    ("Crude Alignment", "YES — direct upstream beneficiary; every $1 Brent = meaningful realisation uplift"),
    ("Option Watch", "ATM CE on monthly expiry (May) — OI liquid; avoid weekly low-liquidity series"),
    ("Entry Trigger", "Brent holds above $95 intraday + stock SuperTrend GREEN on 15-min + MCX crude reclaims ₹8,500"),
    ("Target", "+4–5% on the underlying; calls can 1.5–2x"),
    ("Stop Loss", "Brent breaks $92 OR MCX crude breaks ₹8,200 OR stock breaks prior day low"),
    ("Confidence", "MEDIUM — requires crude to hold; Iran headlines can swing either way"),
])

doc.add_paragraph()
add_para(doc, "Watchlist (not actionable without confirmation)", bold=True, size=11)
add_bullets(doc, [
    "ONGC — same crude-bull thesis as Oil India; watch for same triggers",
    "IndiGo — counter-trade: if crude BREAKS ₹8,000 on fresh Iran de-escalation, long calls; not today's base case",
    "Bank Nifty — 56,565 zone; above 56,800 = reinitiate, below 56,200 = stand aside",
    "Adani Energy Solutions — Q4 today; unpredictable, skip unless you specialise",
])

# ---------- SECTION 5 — DOMESTIC MACRO ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Domestic Macro & India-Specific Triggers", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Driver", "Status"],
    rows=[
        ["RBI Repo Rate", "5.25% unchanged (MPC 6–8 Apr); stance: NEUTRAL"],
        ["CPI (Mar 2026)", "3.4% YoY — below 4% target (supportive)"],
        ["RBI CPI Projection FY27", "4.6%"],
        ["GDP Growth Projection FY27", "6.9%"],
        ["Earnings Today", "INFOSYS (marquee), Adani Energy Solutions, several mid-caps"],
        ["Earnings Yesterday", "Nestle India beat (+27% PAT); Wipro miss; HCLTech reported"],
        ["US Fed Events This Week", "None major today; watch Fed speak on tape"],
        ["US CPI/PPI", "Not releasing today"],
        ["India Macro Data Today", "None scheduled"],
        ["Regulatory", "SEBI expiry-day move: Nifty weekly now TUESDAY (not Thursday)"],
        ["Domestic Macro Verdict", "SUPPORTIVE on rates + CPI; HEADWIND from earnings uncertainty"],
    ])

# ---------- SECTION 6 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 6 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["Crude Rule Mode", "MILD BULL (caution tilt — ₹8,390 MCX)"],
        ["Market Bias", "WAIT — Nifty broke 24,400; needs to reclaim 24,450 to re-enter"],
        ["VIX Sizing Rule", "HALF-SIZE (VIX 17.53 in 17–22 band)"],
        ["Key Support", "24,322 → 24,000"],
        ["Key Resistance", "24,500 → 24,600"],
        ["Critical Crude Level", "₹8,500 MCX — flips bias to half-size if crossed"],
        ["Top Risk Event 1", "Infosys Q4 result 3:45 PM — can gap IT sector ±3%"],
        ["Top Risk Event 2", "Any fresh Iran headline — crude whiplash"],
        ["Entry Permission", "CAUTION — no fresh longs until 24,450 reclaim"],
    ])

doc.add_paragraph()
add_para(doc, "Reason for CAUTION:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "Nifty closed -0.81% below 24,400 support. Both FII (-₹2,078 cr) and DII (-₹1,048 cr) were net sellers — "
    "a rare double-sell that punishes impatient longs. US overnight was a tailwind (S&P, Nasdaq at records) "
    "but Gift Nifty is only flat-to-down. Infosys earnings after-market is the binary event that will "
    "dictate IT-sector direction into Friday.",
    size=11)

doc.add_paragraph()
add_para(doc, "ONE-LINE SUMMARY", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Today is a WAIT day. Crude at ₹8,390 = MILD BULL with fragile tilt. Watch INFY on post-result 5-min "
    "confirmation after 3:45 PM; shadow OIL INDIA if Brent holds $95. Key risk: Iran headlines and INFY "
    "guidance. Size HALF based on VIX at 17.53.\"",
    italic=True, size=12)

# ---------- FOOTER ----------
doc.add_paragraph()
doc.add_paragraph()
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run("Generated by Daily Trading Routine — for personal educational use only. Not financial advice.")
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
