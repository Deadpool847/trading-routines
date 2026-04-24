"""Generate the daily trading-intelligence .docx brief."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-04-24"
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
dr = date_p.add_run(f"Date: Friday, 24 April 2026  |  Generated 08:45 IST")
dr.bold = True
dr.font.size = Pt(11)

# Bias + permission badges
badge_line(doc, "TODAY'S BIAS", BIAS, "F1C232")           # amber for WAIT
badge_line(doc, "ENTRY PERMISSION", ENTRY_PERMISSION, "E69138")  # orange for CAUTION
badge_line(doc, "CRUDE RULE MODE", "CAUTION — HALF SIZE ONLY", "CC0000")

doc.add_paragraph()

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Reading", "Interpretation"],
    rows=[
        ["MCX Crude (₹/bbl)", "~₹8,755 (WTI $94.14 × INR 94.11)", "CAUTION zone (₹8,500–9,000) — HALF SIZE"],
        ["Brent / WTI", "$101.91 rising to $106 / $94.14", "War premium back — Iran ship seizures"],
        ["India VIX", "18.30 (+ from 17.49 intraday low)", "Elevated 17–22 band — reduce size"],
        ["US VIX (CBOE)", "18.92", "Calm-side of fear; below panic"],
        ["S&P 500", "7,108.40 (-0.41%)", "Pulled back from record — HEADWIND"],
        ["Nasdaq", "24,438.50 (-0.89%)", "Software selloff — IT drag"],
        ["Dow Jones", "49,310.32 (-0.36%)", "Modest decline"],
        ["Fear & Greed (US)", "~70 Greed (cooling)", "Off Extreme Greed — softening"],
        ["Nikkei 225", "+0.55% (core CPI 1.8%)", "Japan inflation re-accelerating"],
        ["Hang Seng", "-0.45%", "Weak China sentiment"],
        ["Kospi", "-0.22%", "Asia-Pac mostly soft"],
        ["Gift Nifty", "24,232 (+69 pts vs 24,173 close)", "Slight gap UP — but fragile"],
        ["DXY", "98.57 (-0.02%)", "Below 104 — limited rupee pressure"],
        ["USD/INR", "94.11 (rupee down 10.4% YoY)", "Weak rupee — FII caution"],
    ])

# ---------- SECTION 2 — CRUDE RULE APPLIED ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule Applied", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "MODE: CAUTION — HALF SIZE ONLY.", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_bullets(doc, [
    "MCX crude ~₹8,755 (derived: WTI $94.14 × INR 94.11) — sits squarely in ₹8,500–₹9,000 CAUTION band.",
    "Brent closed $101.91 on 22 Apr, climbed Wed on Iran ship-seizure news, quoted above $106 this morning.",
    "WTI settled $92.96 on Wed; rose to $94.14 on Thu (4th straight session of gains).",
    "Iran seized TWO vessels in Strait of Hormuz Tuesday — shortly after Trump extended ceasefire.",
    "US-Iran talks in Pakistan collapsed; ~800 vessels (incl. 426 tankers) still stranded.",
    "EIA: global supply collapsed 10.1 mb/d in March — 'largest disruption in history.' Shut-ins rise to 9.1 mb/d in April.",
    "Critical level: ₹9,000 MCX. If breached on fresh Iran escalation = flip to PUTS ONLY. Call trades die here.",
    "Downside trigger: any concrete Iran de-escalation = crude dumps fast; would upgrade to full-bull mode.",
])

add_para(doc, "Crude-aligned sector bias (HIGH crude = selective):", bold=True, size=11)
add_bullets(doc, [
    "Bullish: ONGC, Oil India — upstream EBITDA expected +16–36% QoQ on $100+ Brent (ICICI Sec)",
    "Bearish: BPCL, HPCL, IOC (OMC margin squeeze), IndiGo, SpiceJet (aviation fuel costs)",
    "Bearish: Asian Paints, Berger, tyres (MRF, Apollo), plastics (raw-material-linked)",
    "RIL O2C segment now hurt by Hormuz freight/gas cost disruption — see Section 4",
])

# ---------- SECTION 3 — INDIAN MARKET INTERNALS ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Indian Market Internals", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Nifty Technical Levels", bold=True, size=12)
add_kv_table(doc,
    header=["Level", "Value", "Meaning"],
    rows=[
        ["Previous Close (23 Apr)", "24,173.05", "Down -205 pts (-0.84%); 2nd straight selloff"],
        ["Sensex Close", "77,664 (-852 pts, -1.09%)", "Bank-heavy drag"],
        ["Expected Open", "~24,240 (+69 per Gift Nifty)", "Bounce attempt, not a reversal"],
        ["Support S1", "24,134", "Yesterday's intraday low"],
        ["Support S2", "24,000", "Psychological + Put OI zone"],
        ["Support S3", "23,800", "Weekly trend break point"],
        ["Resistance R1", "24,310", "Yesterday's intraday high"],
        ["Resistance R2", "24,400", "Max Pain + reclaim level"],
        ["Resistance R3", "24,500", "Highest Call OI ceiling"],
        ["50 DMA", "~24,122 (bearish pivot)", "Index now testing from above"],
        ["200 DMA", "~23,530 (still above)", "Long-term trend BULLISH"],
    ])

doc.add_paragraph()
add_para(doc, "Volatility & Flows", bold=True, size=12)
add_bullets(doc, [
    "India VIX: 18.30 — ELEVATED (17–22 band). Rule: half-size positions, widen stops.",
    "VIX intraday range 17.49–19.12 yesterday — fear bubbled on Iran headlines.",
    "FII trend (month to date): net SELLERS ~₹20,410 cr cumulatively on NSE cash.",
    "DII trend (month to date): net BUYERS ~₹13,490 cr — partial absorption only.",
    "Rupee 94.11 + DXY 98.57 + weak FII = tape vulnerable to any fresh sell trigger.",
])

doc.add_paragraph()
add_para(doc, "Options Positioning (28 Apr Tuesday expiry — NO weekly expiry today)", bold=True, size=12)
add_bullets(doc, [
    "Max Pain (28 Apr): 24,400 — gravitational magnet above spot (bulls want this)",
    "Highest Call OI: 24,500 CE — firm ceiling; call writers dominant here",
    "Highest Put OI: 24,000 PE — put writers defending; breach = waterfall risk",
    "PCR: 1.1577 (neutral-to-slightly-bullish); needs >1.2 for conviction long",
    "Interpretation: 24,000–24,500 range expected into Tuesday expiry; bias slightly bullish from put writers but fragile if 24,000 breaks",
    "NOTE: NSE weekly now Tuesday (not Thursday) — no expiry decay edge today",
])

doc.add_paragraph()
add_para(doc, "Sector Rotation", bold=True, size=12)
add_bullets(doc, [
    "Yesterday's leaders DOWN: Banking (lead decline), IT (soft), Metals (mixed)",
    "Leaders YTD (context): PSU banks +29%, Metals +27%, Autos +22%, Private banks +15%",
    "Laggards YTD: IT -12%, Pharma -4%, Energy downstream -3%",
    "Today's tilt: UPSTREAM oil (ONGC, Oil India) favoured; AVOID aviation, OMCs, IT",
    "Defence PSUs (HAL, BEL, BDL) — long-term strong but overbought short-term",
])

# ---------- SECTION 4 — STOCK OPPORTUNITIES ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — Stock Opportunities", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc,
    "Two catalyst-driven setups qualify today. RIL is a binary post-market event — do NOT pre-position.",
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

add_para(doc, "STOCK 1 — RELIANCE INDUSTRIES (RIL)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP", "Reference live; large-cap heavyweight, Nifty's biggest weight"),
    ("Catalyst", "Q4 FY26 results TODAY — board meets 24 Apr; release expected post-market ~7 PM"),
    ("Crude Alignment", "NEGATIVE — O2C hurt by Hormuz freight costs + higher gas costs; refining spreads weak"),
    ("Consensus", "Revenue ₹2.7–2.8 tn (+up to 10% YoY); EBITDA ₹44–45k cr (flat); PAT ₹16.2–18.5k cr (up to -17% YoY)"),
    ("Option Watch", "Monthly expiry ATM straddle (May series — 29 May). AVOID weekly — NSE moved to Tuesday."),
    ("Entry Trigger", "POST-RESULT ONLY. Do NOT pre-position. Monday's open determined by tonight's print + guidance on Jio IPO timeline."),
    ("Target", "Beat + strong Jio IPO commentary: +3–5%. Miss + weak O2C guide: -3–5%."),
    ("Stop Loss", "Post-entry only; 5-min candle reversal of trigger bar, OR ₹1,180 cap on the downside"),
    ("Confidence", "HIGH on volatility; LOW on direction (binary event — don't guess)"),
])

doc.add_paragraph()
add_para(doc, "STOCK 2 — OIL INDIA (OIL)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP", "Large-cap PSU upstream — reference live quote"),
    ("Catalyst", "Brent $101–$106 zone; every $1 Brent = meaningful realisation uplift; ICICI Sec: upstream EBITDA +16–36% QoQ"),
    ("Crude Alignment", "YES — direct upstream beneficiary; war premium = dividend for realisation"),
    ("Option Watch", "ATM CE on monthly expiry (May 29). Avoid thin weekly OI."),
    ("Entry Trigger", "Brent holds $100 intraday + stock SuperTrend GREEN on 15-min + Nifty does NOT break 24,000"),
    ("Target", "+4–6% on the underlying; calls can 1.5–2x"),
    ("Stop Loss", "Brent breaks $97 OR MCX crude breaks ₹8,500 OR stock breaks prior day low"),
    ("Confidence", "MEDIUM — requires crude to hold AND broader Nifty not to tank"),
])

doc.add_paragraph()
add_para(doc, "Watchlist (not actionable without confirmation)", bold=True, size=11)
add_bullets(doc, [
    "ONGC — same crude-bull thesis as Oil India; ₹285 zone; 2nd interim div ₹6.25 declared",
    "INFY — Q4 beat (PAT +20.87% to ₹8,501 cr) BUT ADRs fell 5.49% pre-market on weak FY27 guidance (1.5–3.5% CC growth); IT headwind today",
    "Adani Energy — Q4 out yesterday (PAT +5.66%); earnings call 11 AM today; overbought after 45–50% April rally — skip",
    "BPCL/HPCL/IndiGo — consider PUTS if Brent breaks $106 decisively",
    "Bank Nifty — lead yesterday's decline; stand aside until 55,800 holds or reclaims",
])

# ---------- SECTION 5 — DOMESTIC MACRO ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Domestic Macro & India-Specific Triggers", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Driver", "Status"],
    rows=[
        ["RBI Repo Rate", "5.25% unchanged (MPC 6–8 Apr); stance NEUTRAL"],
        ["CPI (Mar 2026)", "3.40% YoY — below 4% target (supportive)"],
        ["WPI (Mar 2026)", "3.88% — HIGHEST in 3 years (crude + manufacturing driven)"],
        ["RBI CPI Projection FY27", "4.6%"],
        ["GDP Growth Projection FY27", "6.9%"],
        ["Earnings Today", "RELIANCE (marquee, post-market); Adani Energy call 11 AM; several mid-caps"],
        ["Earnings Yesterday", "Infosys beat on profit but weak guidance; Adani Energy PAT +5.66%"],
        ["US Fed Events", "No major releases today; next CPI 12 May; next PPI 13 May"],
        ["US Jobs Data", "Next NFP 8 May; March was +178k, unemployment 4.3%"],
        ["India Macro Data Today", "None scheduled"],
        ["Regulatory", "NSE weekly expiry now TUESDAY (not Thursday) — no decay edge today"],
        ["Domestic Macro Verdict", "NEUTRAL — CPI supportive; WPI + crude + weak flows = headwind"],
    ])

# ---------- SECTION 6 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 6 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["Crude Rule Mode", "CAUTION — HALF SIZE ONLY (MCX ~₹8,755 in ₹8,500–9,000 band)"],
        ["Market Bias", "WAIT with bearish undertone — Nifty -0.84%; needs 24,310 reclaim to retry longs"],
        ["VIX Sizing Rule", "HALF-SIZE (India VIX 18.30 in 17–22 elevated band)"],
        ["Key Support", "24,134 → 24,000 → 23,800"],
        ["Key Resistance", "24,310 → 24,400 (max pain) → 24,500 (call wall)"],
        ["Critical Crude Level", "₹9,000 MCX — breach = flip to PUTS ONLY"],
        ["Top Risk Event 1", "Reliance Q4 post-market — can gap index Monday ±1–2% (RIL is Nifty's biggest weight)"],
        ["Top Risk Event 2", "Iran escalation headline — another Hormuz incident can spike crude another 5–7%"],
        ["Entry Permission", "CAUTION — no fresh longs until 24,310 reclaim AND crude doesn't break ₹9,000"],
    ])

doc.add_paragraph()
add_para(doc, "Reason for CAUTION:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "Two-day selloff has broken 24,400 support; index closed at 24,173 (near yesterday's low). Iran seized "
    "two ships Tuesday and US talks in Pakistan collapsed — crude is on a 4-session win streak with Brent above "
    "$101 (reported above $106 this morning). US markets pulled back overnight on same concerns (S&P -0.41%, "
    "Nasdaq -0.89%). RIL's Q4 after market is tonight's binary catalyst — a miss on O2C/guide could drag Nifty "
    "lower Monday given RIL's heavy index weight. With VIX at 18.30 and MCX crude in the caution band, this is "
    "a day to size HALF and wait for confirmation, not to hunt aggressive longs.",
    size=11)

doc.add_paragraph()
add_para(doc, "ONE-LINE SUMMARY", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Today is a WAIT day with bearish undertone. Crude at ₹8,755 MCX = CAUTION (half size). "
    "Watch RIL Q4 post-market (don't pre-position) and shadow OIL INDIA if Brent holds $100 and Nifty holds 24,000. "
    "Key risks: Iran escalation headlines and RIL guidance on O2C/Jio IPO. Size HALF based on VIX at 18.30.\"",
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
