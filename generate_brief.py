"""Generate the daily trading-intelligence .docx brief."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-07-06"
BIAS = "CAUTIOUS"
ENTRY_PERMISSION = "YELLOW"

OUT_PATH = f"trading-briefs/2026/07-July/Trading_Brief_{DATE_ISO}_{BIAS}.docx"


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


def stock_card(doc, rows):
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


# ============== BUILD DOCUMENT ==============
doc = Document()

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
dr = date_p.add_run("Date: Monday, 6 July 2026  |  Generated 08:45 IST")
dr.bold = True
dr.font.size = Pt(11)

badge_line(doc, "TODAY'S BIAS", "CAUTIOUS BULL (defensive tilt)", "F1C232")
badge_line(doc, "ENTRY PERMISSION", "YELLOW — wait 60 min, require confirmation", "E69138")
badge_line(doc, "CRUDE RULE MODE", "AGGRESSIVE_BULL by formula (~₹6,520 MCX) — treat as FRAGILE", "6AA84F")

doc.add_paragraph()

# ---------- SECTION 0 — POSTMORTEM / OPERATIONAL FLAG ----------
add_heading(doc, "Section 0 — Yesterday's Postmortem", level=1, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc, "AUTOMATION GAP — CANNOT SCORE YESTERDAY", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_bullets(doc, [
    "The last brief on file is dated 2026-04-23 (WAIT day, CAUTION permission). Today is 2026-07-06 — a 74-day gap.",
    "trade_log.txt has been empty for the entire gap — zero paper trades logged since the log file was created (2026-04-30).",
    "There is no 'yesterday' within a reasonable window to audit. Brief accuracy score: N/A (not computable).",
    "Per Phase 6 spirit, treating today as a fresh start under a self-imposed recovery posture: 4 stocks instead of 5-7, "
    "every grade held one notch below what the catalysts alone would justify, 60-minute wait instead of 30 before any entry, "
    "and HALF size despite VIX (11.80) technically qualifying for full size.",
    "Operator action needed: confirm whether the daily-brief routine and the paper-trading system were both simply not run "
    "for ~10 weeks, or whether something is broken in the pipeline. This is flagged, not silently absorbed.",
])

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
doc.add_paragraph()
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Reading", "Interpretation"],
    rows=[
        ["WTI Crude", "~$68.3–68.8/bbl", "Down from $89.22 in Apr — multi-month lows"],
        ["Brent Crude (context only)", "~$71.3–72.5/bbl", "Spread to WTI ~$3.5–4/bbl; NEVER used for MCX proxy"],
        ["USD/INR", "~95.24–95.34", "Rupee weaker than April's 93.80 — record-low territory"],
        ["MCX Crude Proxy (WTI×INR)", "≈ ₹6,507–6,550/bbl", "Well inside AGGRESSIVE_BULL (<₹7,500) — see fragility flag"],
        ["India VIX (3 Jul close)", "11.80 (-3.99%)", "Calm zone (<15) — full size by formula; HALF applied (see S0)"],
        ["Nifty 50 (3 Jul close)", "24,270.85 (+0.39%)", "3rd straight winning session"],
        ["Sensex (3 Jul close)", "77,763.91 (+0.34%)", "In line with Nifty"],
        ["Gift Nifty (Mon AM)", "~24,330–24,465 (unconfirmed)", "CONFLICTING sources — one bullish, one 'marginally negative'; treat directionally only"],
        ["Nikkei 225", "69,744 (+1.47%)", "Strong Asia tailwind"],
        ["Hang Seng", "23,055 (+0.8%)", "Modest positive"],
        ["Kospi", "8,088.34 (+5.76%)", "Unusually large single-day move — sanity-check before relying on it"],
        ["Shanghai Composite", "4,043.64 (+0.37%)", "Single-sourced, mild positive"],
        ["Dow Jones (Thu 2 Jul)", "52,900.07 (+1.14%)", "Record closing high; US closed Fri 3 Jul for July 4"],
        ["S&P 500 (Thu 2 Jul)", "7,483.24 (~flat)", "Near record"],
        ["Nasdaq (Thu 2 Jul)", "25,832.67 (-0.8%)", "Semiconductor selloff (SMH -4.5%)"],
        ["FII (3 Jul, cash)", "+₹1,355 cr (net buy)", "Small positive vs ~₹1.92L cr cumulative 2026 outflow — trend still negative"],
        ["DII (3 Jul, cash)", "-₹1,954 cr (net sell)", "Unusual — DII normally absorbs; both-flows diverge from norm"],
    ])

# ---------- SECTION 2 — CRUDE RULE APPLIED ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule Applied", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "MODE: AGGRESSIVE_BULL by formula — but FRAGILE, not clean.", bold=True, size=12)
add_bullets(doc, [
    "MCX proxy ≈ ₹6,507–6,550 (WTI ~$68.5 × USDINR ~95.3), comfortably under the ₹7,500 AGGRESSIVE_BULL threshold.",
    "This is a huge drop from April's ₹8,390 (MILD_BULL) despite an ACTIVE, unresolved Iran-US-Israel war.",
    "Iran-US 'Islamabad Memorandum' (17 Jun) broke down into real strikes within 10 days (26-29 Jun: vessel attack, "
    "US strikes on Iran, Iranian strikes on Kuwait/Bahrain bases). Talks are now paused 4-9 July for Khamenei's funeral; "
    "next round 11 July. Iran's ambassador says it will 'definitely' charge Hormuz fees — contradicts the MoU's toll-free terms.",
    "Strait of Hormuz transit is running at ~25-30% of normal volume. IRGC patrol boats issued a fresh warning to shipping "
    "as recently as 5 July (yesterday) — this is not a resolved situation, it is a lull.",
    "Per the desk's own rule, weekend/holiday geopolitical calm requires 3-5 sessions of confirmation before being treated "
    "as a shift — this lull does not clear that bar; treat the low crude print as fragile, not settled.",
    "Critical level unchanged conceptually: ₹7,500 MCX (AGGRESSIVE_BULL/MILD_BULL boundary) and ₹8,500 (MILD_BULL/CAUTION). "
    "Currently far below both, but a single Hormuz incident can move WTI 5-10%+ intraday and erase this cushion fast.",
])

add_para(doc, "Crude-aligned sector bias (regime has FLIPPED vs April):", bold=True, size=11)
add_bullets(doc, [
    "Bullish now: OMCs (BPCL, HPCL, IOC — cheap crude widens marketing margins), aviation (IndiGo — lower fuel bill), "
    "auto (lower running costs support demand).",
    "Bearish now: ONGC, Oil India (upstream realisations compress as WTI/Brent fall from April's levels) — inverse of the April thesis.",
    "Hedge: Indian defence names (BEL, HAL) benefit from the geopolitical backdrop independent of the crude-price direction.",
])

# ---------- SECTION 3 — INDIAN MARKET INTERNALS ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Indian Market Internals", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Nifty Technical Levels", bold=True, size=12)
add_kv_table(doc,
    header=["Level", "Value", "Meaning"],
    rows=[
        ["Previous Close", "24,270.85", "+95.15 pts (+0.39%); 3rd straight up session"],
        ["Day Range (3 Jul)", "24,252.35 – 24,378.15", "Tight range, orderly session"],
        ["Gift Nifty implied open", "Unconfirmed — sources conflict", "Treat as flat-to-mild either way until confirmed at open"],
        ["Support (Max Pain zone)", "24,200–24,271", "Very tight to spot — gravitational pull into Tue expiry"],
        ["Support (Put OI wall)", "24,000", "Psychological + heaviest put OI"],
        ["Resistance (Call OI wall)", "25,000", "Heaviest call OI — hard cap for the week"],
    ])

doc.add_paragraph()
add_para(doc, "Volatility & Flows", bold=True, size=12)
add_bullets(doc, [
    "India VIX: 11.80 (-3.99%) — CALM zone (<15). Formula says full size; HALF applied given the 74-day audit gap (Section 0).",
    "FII (3 Jul): net BUY ₹1,355 cr — small positive print, but 2026 YTD cumulative outflow (~₹1.92L cr through early May) "
    "already exceeds all of 2025 — the underlying trend is still a headwind.",
    "DII (3 Jul): net SELL ₹1,954 cr — unusual; normally DII absorbs FII selling, here both flows point the same odd way.",
    "USD/INR at 95.24-95.34 is materially weaker than April's 93.80 — continued FII caution on currency.",
])

doc.add_paragraph()
add_para(doc, "Options Positioning (7 Jul Tuesday expiry — NSE moved Nifty weekly from Thursday to Tuesday)", bold=True, size=12)
add_bullets(doc, [
    "Max Pain (as of 3 Jul): ~24,200–24,271 — right on top of spot, a genuinely tight gravitational read.",
    "Highest Call OI: 25,000 CE — resistance/ceiling for the week.",
    "Highest Put OI: 24,000 PE — support/floor.",
    "Interpretation: smart money positioning implies range-bound drift 24,000–25,000 into Tuesday's expiry.",
    "F&O ban list for 6 Jul: one aggregator reports 'no securities in ban' — COULD NOT independently verify via a live NSE "
    "fetch (403s on direct checks). Verify against nseindia.com before entering any position outside today's 4 setups.",
])

doc.add_paragraph()
add_para(doc, "Sector Rotation / Catalysts", bold=True, size=12)
add_bullets(doc, [
    "IT sector under heavy brokerage pressure into Q1 FY27 results: ICICI Securities cut TCS target -33.6% (₹2,800→₹1,860), "
    "Infosys -27%, HCL Tech -34%, Wipro -22%; Kotak also cut IT estimates. 8-9 of 10 Nifty IT stocks are 20-27% below their 200-DMA.",
    "TCS reports Q1 FY27 on Thursday 9 July (after hours) — kicks off the IT earnings season "
    "(HCL Tech 13 Jul, LTTS 14 Jul, Tech Mahindra 16 Jul, LTIMindtree 17 Jul, Infosys/Coforge/Persistent 23 Jul).",
    "Auto sector strong on June sales: Maruti +23.8% YoY (200,390 units), Tata Motors PV +69% YoY (63,083 units), "
    "Tata Motors EV +183% YoY (record monthly EV sales).",
    "RBI (June MPC): held repo at 5.25%, neutral stance, BUT cut FY27 GDP forecast to 6.6% (from 6.9%) and RAISED CPI "
    "forecast to 5.1% (from 4.6%), explicitly citing the West Asia conflict, energy prices, and monsoon uncertainty.",
    "IMD: June 2026 rainfall ~40% below normal (5th driest June since 1901); July forecast also below-normal. "
    "Negative read-through for rural demand/FMCG, and a direct input into RBI's own inflation-forecast hike.",
])

# ---------- SECTION 4 — CONTRARIAN CHECK ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — Contrarian Check (Layer 1/2/3)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Layer 1 — Consensus:", bold=True, size=11)
add_para(doc, "\"Crude at multi-month lows, US jobs miss cools Fed hike fears, gold at records, Nifty on a 3-day winning "
              "streak into expiry — risk-on, buy the dip, VIX at 11.80 means smooth sailing.\"", italic=True, size=11)

add_para(doc, "Layer 2 — Likely wrong:", bold=True, size=11)
add_bullets(doc, [
    "Consensus is pricing crude as if Hormuz is basically resolved. It isn't — transit is still only 25-30% of normal, "
    "IRGC issued a fresh shipping warning yesterday (5 Jul), and the 'calm' is an explicitly temporary funeral pause "
    "(talks resume 11 Jul) on top of a track record of MoUs collapsing within days.",
    "The India VIX at 11.80 may be underpricing this fragility rather than correctly reading a genuinely calm market.",
    "DII net-selling ₹1,954 cr on a day FII bought is an odd divergence from the usual pattern — worth watching, not ignoring.",
])

add_para(doc, "Layer 3 — Underestimated:", bold=True, size=11)
add_bullets(doc, [
    "Below-normal monsoon (5th driest June since 1901) plus RBI's own inflation-forecast hike is a slow-burning "
    "stagflation-lite risk that's being crowded out of the narrative by Iran headlines.",
    "IT's brokerage-driven de-rating into results week is being treated as sector-specific, but IT's large index weight "
    "means a bad TCS print Thursday could spill into broader Nifty sentiment, not stay contained.",
    "The FY26 FII cumulative outflow (~₹1.92L cr) dwarfs the single-day ₹1,355 cr buy print — the flow trend underneath "
    "the calm surface is still negative.",
])

add_para(doc, "Flip triggers: BEARISH.", bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc, "Entry permission: YELLOW — proceed only with confirmation and the 60-minute wait discipline.", bold=True, size=11)

# ---------- SECTION 5 — STOCK OPPORTUNITIES ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Stock Opportunities (4 setups — reduced from the usual 5-7, see Section 0)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "STOCK 1 — TATA MOTORS PASSENGER VEHICLES (TMPV) — CALL — Grade A (downgraded from A+)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot / Sector", "₹432 / Lot 800 / Auto"),
    ("Why this grade", "June PV sales +69% YoY (63,083 units); EV sales +183% YoY, a record month. Low-crude regime is a "
                        "tailwind for fuel-cost-sensitive discretionary auto demand. Downgraded one notch for the audit-gap caution."),
    ("Option Setup", "July monthly ATM/slightly-OTM CE (avoid low-liquidity weekly-style strikes — TMPV is stock options, monthly only)"),
    ("Entry Trigger", "(1) Nifty Auto index green at open, (2) TMPV 15-min close holds above ₹430, (3) Nifty overall not below "
                       "24,150, (4) WAIT 60 minutes post-open before entry"),
    ("Target / SL / R:R", "T1 +3% (40% book), T2 +5-6% (40% book) / SL: 15-min close below ₹420 / R:R ≈ 2.2:1"),
    ("Confidence", "MEDIUM-HIGH — clean catalyst + macro alignment; note TMPV is a newly demerged contract (since Oct 2025 "
                    "split) with a shorter trading history — check live liquidity before sizing up"),
])

doc.add_paragraph()
add_para(doc, "STOCK 2 — BPCL — CALL — Grade A", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot / Sector", "₹310.4 / Lot 1975 / Energy (OMC)"),
    ("Capital-fit flag", "Lot × spot × 3% ≈ ₹18,400 — ~23% over the ₹15,000 rule-of-thumb. Trade the OPTION premium only "
                          "(not futures) so real capital-at-risk stays well inside budget; flagged rather than silently ignored."),
    ("Why this grade", "WTI crashed to ~$68 from April's $89 — AGGRESSIVE_BULL crude regime directly widens OMC refining/"
                        "marketing margins. Direct, clean beneficiary of today's regime."),
    ("Option Setup", "July monthly ATM CE"),
    ("Entry Trigger", "(1) MCX crude proxy stays below ₹7,000 intraday, (2) BPCL 15-min close above ₹312, (3) Nifty Energy "
                       "index green, (4) WAIT 60 minutes post-open"),
    ("Target / SL / R:R", "T1 +3% (40% book), T2 +5% (40% book) / SL: 15-min close below ₹302 / R:R ≈ 2:1"),
    ("Confidence", "MEDIUM — clean thesis, but the exact tail risk that kills it (a Hormuz shock spiking crude) is live "
                    "and unresolved right now. Respect the stop strictly."),
])

doc.add_paragraph()
add_para(doc, "STOCK 3 — TCS — PUT — Grade A (pre-results technical short, NOT an earnings-gap bet)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot / Sector", "~₹2,100 / Lot 175 / IT"),
    ("Why this grade", "ICICI Securities cut TP -33.6% (₹2,800→₹1,860); Kotak also slashed IT estimates. TCS ~25% below its "
                        "200-DMA; 8-9 of 10 Nifty IT names similarly de-rated ahead of Q1 FY27 results Thu 9 Jul."),
    ("Option Setup", "July monthly ATM/slightly-OTM PE"),
    ("Entry Trigger", "(1) TCS 15-min close below ₹2,080, (2) Nifty IT index red, (3) no positive pre-results news reversing "
                       "sentiment, (4) WAIT 60 minutes post-open"),
    ("HARD RULE", "This is a pre-results technical short, not a results bet. FLATTEN fully by Wednesday 8 July close — "
                   "do NOT hold through Thursday's print in either direction."),
    ("Target / SL / R:R", "T1 -3% (40% book), T2 -5% (40% book) / SL: 15-min close above ₹2,130 / R:R ≈ 2:1"),
    ("Confidence", "MEDIUM — strong bearish catalyst, but the stock is already well off its highs, so some mean-reversion "
                    "bounce risk exists before Thursday."),
])

doc.add_paragraph()
add_para(doc, "STOCK 4 — BHARAT ELECTRONICS (BEL) — CALL — Grade B+ (macro-hedge)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot / Sector", "₹418 / Lot 1425 / Defence"),
    ("Capital-fit flag", "Lot × spot × 3% ≈ ₹17,900 — also over the ₹15,000 rule-of-thumb; option-premium-only sizing applies here too."),
    ("Why this grade", "Active Iran-Israel-US war, impaired Hormuz, Israel still occupying ~2,000 sq km of Lebanon — a live "
                        "geopolitical backdrop that historically supports India's defence/indigenisation theme. Single strong "
                        "catalyst + macro alignment; this is the portfolio's explicit hedge, not the highest-conviction pick."),
    ("Option Setup", "July monthly ATM CE"),
    ("Entry Trigger", "(1) any fresh escalation headline (Hormuz incident, new strikes) OR (2) BEL 15-min close above ₹422 on "
                       "its own technical strength, (3) WAIT 60 minutes, (4) size as a pure hedge — do not scale beyond 1 lot equivalent"),
    ("Target / SL / R:R", "T1 +3% (40% book), T2 +5% (40% book) / SL: 15-min close below ₹408 / R:R ≈ 2:1"),
    ("Confidence", "MEDIUM — works best precisely if the other three CALL/PUT setups struggle, i.e., if geopolitics worsens."),
])

doc.add_paragraph()
add_para(doc, "Stocks to actively avoid today", bold=True, size=11)
add_bullets(doc, [
    "ONGC / Oil India — upstream realisations directly hurt by the WTI collapse; April's bullish thesis has fully flipped.",
    "IndiGo / aviation — plausible low-crude beneficiary, but no confirmed fresh catalyst today; park for a follow-up brief "
    "once a specific trigger is confirmed rather than trading on an inferred story.",
    "Vodafone Idea — Nomura downgraded to Neutral on execution/funding risk.",
    "Anything outside these 4 names without a live, verified F&O ban-list check (today's 'clean' reading is unconfirmed).",
])

# ---------- SECTION 6 — DOMESTIC MACRO ----------
doc.add_paragraph()
add_heading(doc, "Section 6 — Domestic Macro & India-Specific Triggers", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Driver", "Status"],
    rows=[
        ["RBI Repo Rate", "5.25% held (June MPC, 6-0 unanimous); stance NEUTRAL; next MPC 3-5 Aug 2026"],
        ["RBI FY27 GDP Forecast", "Cut to 6.6% (from 6.9% in April)"],
        ["RBI FY27 CPI Forecast", "Raised to 5.1% (from 4.6%) — citing war, energy prices, monsoon"],
        ["Monsoon (IMD)", "June 2026 ~40% below normal (5th driest June since 1901); July also forecast below-normal"],
        ["India-US Trade Talks", "Unresolved; Section 122 10% blanket tariff expires 24 Jul 2026; reciprocal tariff cut to "
                                  "~18% (from 25%) under negotiation, not finalised"],
        ["Earnings This Week", "TCS Q1 FY27 — Thu 9 Jul (after hours); opens the IT results season"],
        ["June Auto Sales", "Maruti +23.8% YoY; Tata Motors PV +69% YoY, EV +183% YoY (record)"],
        ["F&O Expiry", "Nifty weekly — Tuesday 7 Jul 2026 (NSE moved weekly expiry to Tuesday, effective Sep 2025)"],
        ["Fed", "New Chair Kevin Warsh; 17 Jun FOMC held rates, hawkish dot plot; 2 Jul jobs miss (57k vs 110k) pared hike "
                "odds to ~50% for Sept; next FOMC 28-29 Jul"],
        ["China PMI (June)", "50.3, back in expansion (from 50.0)"],
        ["Gold", "Near record highs, ~$4,170-4,200/oz"],
        ["Domestic Macro Verdict", "MIXED — supportive on rates-held/auto-demand, headwind from monsoon/inflation-forecast "
                                    "and unresolved trade-tariff deadline"],
    ])

# ---------- SECTION 7 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 7 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["Crude Rule Mode", "AGGRESSIVE_BULL by formula (~₹6,520 MCX) — treat as FRAGILE given active Iran war/Hormuz status"],
        ["Market Bias", "Cautiously bullish with a defensive tilt (3 CALL / 1 PUT / 1 built-in hedge)"],
        ["Confidence", "MEDIUM — reduced specifically because of the 74-day audit gap (Section 0)"],
        ["VIX Sizing Rule", "11.80 qualifies for FULL size by formula; HALF SIZE applied as a self-imposed override"],
        ["Key Support", "24,200 (max pain) → 24,000 (put wall)"],
        ["Key Resistance", "25,000 (call wall)"],
        ["Critical Crude Level", "₹7,500 MCX (AGGRESSIVE_BULL/MILD_BULL boundary) — far below currently, but one Hormuz "
                                  "shock can close that gap in hours"],
        ["Top Risk Event 1", "Any Hormuz incident before the 11 Jul talks resume — funeral pause is temporary, not a resolution"],
        ["Top Risk Event 2", "TCS Q1 results Thu 9 Jul — IT sector already de-rated hard into the print"],
        ["Entry Permission", "YELLOW — wait 60 minutes, require confirmation on every setup"],
        ["Expected Nifty Range", "24,100 – 24,500, with a pull toward the 24,200-24,271 max-pain zone into Tuesday's expiry"],
    ])

doc.add_paragraph()
add_para(doc, "Bull case:", bold=True, size=11)
add_para(doc, "Low crude + cooling Fed-hike fears + strong auto/GST tailwinds extend the 3-day Nifty winning streak toward 24,500+.", size=11)
add_para(doc, "Bear case:", bold=True, size=11)
add_para(doc, "A Hormuz incident or a weak TCS print (or both) breaks the calm-VIX complacency fast; Nifty retests 24,000.", size=11)
add_para(doc, "Most likely scenario:", bold=True, size=11)
add_para(doc, "Range-bound drift 24,150-24,450 into Tuesday's expiry; real direction resolved by Thursday's TCS print and by "
              "whether the Iran funeral pause survives past 9 July without another strike cycle.", size=11)

doc.add_paragraph()
add_para(doc, "THE ONE NUMBER THAT MATTERS", bold=True, size=13, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc, "₹7,500 MCX crude — the AGGRESSIVE_BULL/MILD_BULL boundary. We're ₹1,000 below it with room to spare, but a "
              "single Hormuz headline can move WTI 5-10%+ intraday and erase that cushion before lunch.", italic=True, size=11)

# ---------- SECTION 8 — PAPER TRADING NOTE ----------
doc.add_paragraph()
add_heading(doc, "Section 8 — Note on Paper Trading Today", level=1, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "trade_log.txt has recorded ZERO trades across the entire 10-week gap since the last brief — verify the paper-trading "
    "system itself is actually running and logging, independent of whether the daily brief routine was being run.",
    "Watch today specifically for: does the system fire a signal on any of the 4 setups above, and does the log actually "
    "capture it? Treat today as a basic plumbing check as much as a trading day.",
    "Running zero-trade day count: effectively unknown/unbounded across the gap — restart the count from today.",
])

doc.add_paragraph()
add_para(doc, "ONE-LINE SUMMARY", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"74-day gap since the last logged brief — treat today as a fresh start, not a continuation. Crude has crashed to "
    "₹6,500 MCX (AGGRESSIVE_BULL) despite an unresolved Iran war and an impaired Hormuz Strait; Nifty (24,271) is on a "
    "3-day win streak into Tuesday's expiry with VIX calm at 11.80. Plays: TMPV & BPCL calls (low-crude tailwind), TCS "
    "put (broker-hammered IT sector, flatten before Thursday's results), BEL call as a geopolitical hedge. Half-size, "
    "wait 60 minutes, entry permission YELLOW.\"",
    italic=True, size=12)

add_para(doc,
    "OPERATOR CONFIRMATION REQUIRED — reply READY TO TRADE / QUESTIONS: / OVERRIDE: [reason] before acting on this brief.",
    bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))

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
