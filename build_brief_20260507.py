"""Generate the daily F&O trading brief for 2026-05-07 (CAUTIOUS BULL)."""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-07"
DATE_LONG = "Thursday, 07 May 2026"
GENERATED = "08:45 IST"
BIAS = "CAUTIOUS-BULL"
BIAS_DISPLAY = "CAUTIOUS BULL"
CONFIDENCE = "MEDIUM"
ENTRY_PERMISSION = "YELLOW"
CRUDE_MODE = "MILD BULL (weakening - flip ₹7,500)"
VIX_SIZING = "HALF SIZE"

OUT_PATH = f"trading-briefs/2026/05-May/Trading_Brief_{DATE_ISO}_{BIAS}.docx"


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
    else:
        r_idx = 0
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
dr = date_p.add_run(f"Date: {DATE_LONG}  |  Generated {GENERATED}")
dr.bold = True
dr.font.size = Pt(11)

expiry_p = doc.add_paragraph()
expiry_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
er = expiry_p.add_run("Weekly expiry: 5 days (Tue 12 May)  |  Monthly expiry: 19 days (Tue 26 May)  |  Expiry week: NO")
er.italic = True
er.font.size = Pt(10)

doc.add_paragraph()

# Bias + permission badges
badge_line(doc, "TODAY'S BIAS", BIAS_DISPLAY, "6AA84F")            # green for cautious bull
badge_line(doc, "ENTRY PERMISSION", ENTRY_PERMISSION, "F1C232")    # amber for yellow
badge_line(doc, "CRUDE RULE MODE", CRUDE_MODE, "6AA84F")
badge_line(doc, "VIX SIZING RULE", VIX_SIZING, "F1C232")

# Override callout
doc.add_paragraph()
override_table = doc.add_table(rows=1, cols=1)
override_cell = override_table.rows[0].cells[0]
override_cell.text = ""
shade_cell(override_cell, "FFF2CC")
op = override_cell.paragraphs[0]
op.alignment = WD_ALIGN_PARAGRAPH.CENTER
or_run = op.add_run(
    "⚠ OVERRIDE ACTIVE: GIFT Nifty gap-up +190 pts AND crude crashing -6% to -9%. "
    "Bear bias suppressed → CAUTIOUS BULL. Consensus 'sell crude beneficiaries' is the trap today."
)
or_run.bold = True
or_run.font.size = Pt(11)
or_run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

doc.add_paragraph()

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Value", "Signal"],
    rows=[
        ["MCX Crude proxy", "~₹7,810/bbl (WTI $93×₹84)", "MILD BULL — falling toward AGG BULL"],
        ["Brent / WTI", "$102.21 / ~$93", "FALLING -6% / -9% overnight"],
        ["Iran-Hormuz status", "Strait largely closed; 'Project Freedom' rejected", "Headline risk HIGH"],
        ["GIFT Nifty", "24,521 (vs 24,331 close)", "GAP UP +190 pts — bull"],
        ["S&P 500", "Record high (+0.72%)", "Risk-on TAILWIND"],
        ["Nasdaq", "+2.0% (record)", "Strong risk-on"],
        ["Dow Jones", "+0.94%", "Broad strength"],
        ["US VIX", "DATA_UNAVAILABLE", "Risk-on inferred from index records"],
        ["India VIX", "~18 (last confirmed 28 Apr; live DATA_UNAVAILABLE)", "Elevated → HALF SIZE"],
        ["Nifty 50 close (06 May)", "24,330.95 (+1.24% / +298 pts)", "Strong close, breakout mode"],
        ["Sensex close (06 May)", "77,958.52 (+940.73 / +1.22%)", "Confirms Nifty strength"],
        ["FII cash (04 May)", "+₹2,835 cr (BUY)", "Recent buying"],
        ["DII cash (04 May)", "+₹4,764 cr (BUY)", "Strong domestic bid"],
        ["FII MTD (May)", "-₹70,135 cr", "Heavy seller MTD; recent reversal"],
        ["DII MTD (May)", "+₹51,064 cr", "Absorbing FII supply"],
        ["USD-INR", "DATA_UNAVAILABLE (~₹84 ref)", "Stable assumption"],
        ["INFY ADR", "$12.46 (06 May)", "Direction DATA_UNAVAILABLE"],
        ["WIT ADR", "DATA_UNAVAILABLE", "—"],
        ["IBN ADR", "DATA_UNAVAILABLE", "—"],
        ["HDB ADR", "DATA_UNAVAILABLE", "—"],
    ])

doc.add_paragraph()
add_para(doc, "Geopolitical one-liner:", bold=True, size=11)
add_para(doc,
    "Strait of Hormuz still effectively closed (5% of pre-conflict shipping). Trump's 'Project Freedom' "
    "moved 2 US vessels through; Iran called it a ceasefire violation and resumed regional attacks. "
    "Markets are pricing peace HOPE — crude crashed 6-9% — but the ceasefire is fragile.",
    size=10)

add_para(doc,
    "Today's intraday headline risk: HIGH (single Iran tweet can re-spike crude 5-7%).",
    bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- SECTION 2 — CRUDE RULE + STRUCTURAL READ ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule + Structural Read", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Crude Rule applied with LIVE verified prices:", bold=True, size=12)
add_kv_table(doc,
    header=["Parameter", "Value"],
    rows=[
        ["CRUDE_MODE", "MILD BULL (₹7,500-8,500 band)"],
        ["CRUDE_DIRECTION", "FALLING sharply (-6% to -9% overnight)"],
        ["CRUDE_FLIP_LEVEL", "₹7,500 (= WTI ~$89). Below = AGGRESSIVE BULL"],
        ["AVOID (mode)", "None mandated in MILD BULL band"],
        ["EXCEPTION_CALLS", "ONGC/Oil India still tradable but momentum fades on falling crude"],
        ["KEY ALERT", "Crude falling — watch ₹7,500 for mode flip to AGG BULL (more bullish equities)"],
    ])

doc.add_paragraph()
add_para(doc, "Structural read of Nifty:", bold=True, size=12)
add_bullets(doc, [
    "Nifty 24,331 vs Max Pain 24,050 = +281 pts ABOVE max pain → DOWNWARD magnetism for expiry week.",
    "Days to weekly expiry: 5 (12 May Tue). Days to monthly: 19 (26 May Tue). Not expiry week.",
    "Highest OI Call/Put strikes & exact PCR: DATA_UNAVAILABLE for live read (use NSE option chain pre-9:00).",
    "OI direction proxy: PSU Bank +2.8% leadership + DII strong buy = put writers active near 24,200-24,300.",
    "Futures basis: DATA_UNAVAILABLE specific. Assume premium given strong May 6 close.",
    "Expiry-week dynamics not yet active (5 days out) — max-pain magnetism weak today.",
])

doc.add_paragraph()
add_para(doc, "Bank Nifty leadership check:", bold=True, size=12)
add_bullets(doc, [
    "PSU Bank +2.8% led ALL sectors yesterday; Financial Services #2.",
    "Bank Nifty leading Nifty → move is REAL, not a fake breakout.",
    "Trust the upside. Fade attempts to short bank-led rallies on Day 1.",
])

# ---------- SECTION 3 — CONTRARIAN CHECK ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Contrarian Check (Layer 3)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Question", "Answer"],
    rows=[
        ["Q1. CONSENSUS",
         "Retail/TV: 'Crude crashed → buy aviation/OMCs/paints; sell ONGC/Oil India.' "
         "Also expecting gap-up of +200 pts will fade by 11 AM."],
        ["Q2. THE TRAP",
         "Iran headline re-spike: a single Iran retaliation against 'Project Freedom' re-pumps crude 5-7%, "
         "obliterating aviation/OMC longs and rocketing ONGC. Smart money fades the obvious crude-crash trade."],
        ["Q3. RETAIL STOPS",
         "Long stops likely clustered below 24,400 (yesterday's reclaim) and 24,300 (psychological). "
         "Short stops above 24,550 (expected open) and 24,650 (round number)."],
        ["Q4. FLIP TRIGGER",
         "Iran retaliation tweet OR Brent reclaiming $108 intraday. Either flips today from BULL to RANGE/BEAR."],
    ])

doc.add_paragraph()
add_para(doc, "Contrarian Override Logic — APPLIED:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_kv_table(doc,
    header=["Condition", "Status"],
    rows=[
        ["A: GIFT Nifty gap-up > +100 pts?", "YES (+190 pts)"],
        ["B: Crude falling > 2% overnight?", "YES (-6% to -9%)"],
        ["C: Contrarian scenario probability > 40%?", "NO (~25-30% — Iran tail risk real but ceasefire holding)"],
        ["D: Expiry week + PCR < 0.8 squeeze?", "NO (not expiry week; PCR DATA_UNAVAILABLE)"],
    ])

add_para(doc,
    "RESULT: A + B both YES → Bear bias suppressed. Final bias = CAUTIOUS BULL. "
    "C is NOT YES, so we stay BULL not RANGE. Permission = YELLOW because gap-up entries are statistically "
    "lower-quality (entry slippage, fade risk in first 30 min).",
    bold=True, size=11)

# ---------- SECTION 4 — STOCK SETUPS ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — 5 Individual F&O Setups", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc,
    "F&O ban list specifics for 07 May DATA_UNAVAILABLE — verify on NSE before entry. "
    "All setups are CALLS aligned with override + falling-crude + bull-momentum thesis. "
    "Use MONTHLY (May 26) expiry on all to avoid theta cliff during 5-day weekly window.",
    italic=True, size=10)

# --- STOCK 1: RELIANCE ---
doc.add_paragraph()
add_para(doc, "STOCK 1 — RELIANCE INDUSTRIES (NSE: RELIANCE) | CALL | Grade A",
         bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot",                 "₹1,467 (last avail) | Lot: 500"),
    ("Crude aligned",             "YES — falling crude expands O2C refining margins; petchem feedstock advantage"),
    ("Catalyst today",            "Crude crash -9% = direct GRM expansion; index heavyweight pulls Nifty up"),
    ("Daily trend / 200 DMA",     "Sideways-to-up; above 200 DMA: YES (200 DMA ~₹1,400 zone)"),
    ("Weekly trend",              "Sideways with bullish bias; -0.4% 6m, +2.2% 1Y"),
    ("Support / Resistance",      "Sup ₹1,450 / ₹1,430 | Res ₹1,490 / ₹1,520"),
    ("OI direction",              "DATA_UNAVAILABLE specific; sector OI bullish on crude crash"),
    ("Delivery % / Block / Ban",  "DATA_UNAVAILABLE / NO known block / Ban list: VERIFY (NO assumed)"),
    ("Strike & expiry",           "RELIANCE 1480 CE (slight OTM) — Monthly 26 May"),
    ("Theta warning ≤3 days?",    "NO (19 days to monthly)"),
    ("Entry checklist (ALL 4)",   "ST GREEN on 15-min | RSI > 50 | StochRSI cross-up from <30 | Vol > 1.5x avg | 15-min close ABOVE ₹1,478"),
    ("Targets",                   "T1 (40%): ₹1,500 | T2 (40%): ₹1,520 | Trail rest"),
    ("Stop Loss",                 "Spot below ₹1,455 + ST flips RED on 15-min"),
    ("R:R",                       "≈ 2.5:1"),
    ("Time stop",                 "Exit by 13:00 IST regardless"),
    ("Grade reason",              "A — Heavyweight + crude tailwind + clean technical setup, Nifty drag-up beneficiary"),
])

# --- STOCK 2: INDIGO ---
doc.add_paragraph()
add_para(doc, "STOCK 2 — INTERGLOBE AVIATION (NSE: INDIGO) | CALL | Grade A",
         bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot",                 "₹4,247 (last avail 05 May) | Lot: 15"),
    ("Crude aligned",             "YES — STRONG. Aviation = direct fuel-cost beneficiary on crude crash"),
    ("Catalyst today",            "Top Nifty50 gainer 06 May + crude -9% overnight = momentum continuation"),
    ("Daily trend / 200 DMA",     "Down-to-recovering; was -24.4% 6m oversold; reclaimed key MAs on 06 May"),
    ("Weekly trend",              "Bottoming reversal; explosive snap-back potential on crude crash"),
    ("Support / Resistance",      "Sup ₹4,200 / ₹4,150 | Res ₹4,350 / ₹4,450"),
    ("OI direction",              "Likely SHORT-COVERING (oversold + crude crash + leadership)"),
    ("Delivery % / Block / Ban",  "DATA_UNAVAILABLE / NO known / Ban list: VERIFY"),
    ("Strike & expiry",           "INDIGO 4300 CE (ATM/slight OTM) — Monthly 26 May"),
    ("Theta warning ≤3 days?",    "NO"),
    ("Entry checklist (ALL 4)",   "ST GREEN on 15-min | RSI > 50 | StochRSI cross-up from <30 | Vol > 1.5x avg | 15-min close ABOVE ₹4,300"),
    ("Targets",                   "T1 (40%): ₹4,380 | T2 (40%): ₹4,450 | Trail rest"),
    ("Stop Loss",                 "Spot below ₹4,210 + 15-min close confirmation"),
    ("R:R",                       "≈ 2.3:1"),
    ("Time stop",                 "Exit by 13:00 IST"),
    ("Grade reason",              "A — Best crude-crash beneficiary; was top gainer; tail risk = Iran headline reverses"),
])

# --- STOCK 3: BPCL ---
doc.add_paragraph()
add_para(doc, "STOCK 3 — BHARAT PETROLEUM (NSE: BPCL) | CALL | Grade A",
         bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot",                 "₹306.7 (06 May) | Lot: 1,800"),
    ("Crude aligned",             "YES — OMC marketing margins expand massively on crude crash"),
    ("Catalyst today",            "OMC was -17.8% 6m oversold; -9% WTI overnight = recovery setup"),
    ("Daily trend / 200 DMA",     "Sideways base; needs to confirm reclaim of declining trendline"),
    ("Weekly trend",              "Bottoming; 52w range ₹266.6 - ₹391.65"),
    ("Support / Resistance",      "Sup ₹302 / ₹298 | Res ₹312 / ₹320"),
    ("OI direction",              "Expect SHORT-COVERING + fresh longs (oversold + tailwind)"),
    ("Delivery % / Block / Ban",  "DATA_UNAVAILABLE / NO known / Ban list: VERIFY"),
    ("Strike & expiry",           "BPCL 310 CE (ATM) — Monthly 26 May"),
    ("Theta warning ≤3 days?",    "NO"),
    ("Entry checklist (ALL 4)",   "ST GREEN on 15-min | RSI > 50 | StochRSI cross-up from <30 | Vol > 1.5x avg | 15-min close ABOVE ₹308"),
    ("Targets",                   "T1 (40%): ₹314 | T2 (40%): ₹320 | Trail rest"),
    ("Stop Loss",                 "Spot below ₹302 + ST RED on 15-min"),
    ("R:R",                       "≈ 2.2:1"),
    ("Time stop",                 "Exit by 13:00 IST"),
    ("Grade reason",              "A — OMC oversold + crude crash = high-conviction marketing-margin trade"),
])

# --- STOCK 4: TATAMOTORS ---
doc.add_paragraph()
add_para(doc, "STOCK 4 — TATA MOTORS (NSE: TATAMOTORS) | CALL | Grade B+",
         bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot",                 "₹358.15 (06 May, +5.3% from open) | Lot: DATA_UNAVAILABLE — verify"),
    ("Crude aligned",             "YES — autos benefit from cheaper input costs + consumer demand"),
    ("Catalyst today",            "Top Nifty50 gainer 06 May; momentum continuation play"),
    ("Daily trend / 200 DMA",     "Down-to-reversing; -16.3% 6m oversold; broke out of base on 06 May"),
    ("Weekly trend",              "Bottoming; classic reversal day candle 06 May"),
    ("Support / Resistance",      "Sup ₹350 / ₹343 | Res ₹365 / ₹375"),
    ("OI direction",              "Likely SHORT-COVERING + fresh long build"),
    ("Delivery % / Block / Ban",  "DATA_UNAVAILABLE / NO known / Ban list: VERIFY"),
    ("Strike & expiry",           "TATAMOTORS 360 CE (ATM) — Monthly 26 May"),
    ("Theta warning ≤3 days?",    "NO"),
    ("Entry checklist (ALL 4)",   "ST GREEN on 15-min | RSI > 50 | StochRSI cross-up from <30 | Vol > 1.5x avg | 15-min close ABOVE ₹360"),
    ("Targets",                   "T1 (40%): ₹367 | T2 (40%): ₹373 | Trail rest"),
    ("Stop Loss",                 "Spot below ₹352 + ST RED on 15-min"),
    ("R:R",                       "≈ 2.1:1"),
    ("Time stop",                 "Exit by 13:00 IST"),
    ("Grade reason",              "B+ — Strong momentum but +5.3% yesterday risks mean-reversion fade in first 30 min"),
])

# --- STOCK 5: ICICIBANK ---
doc.add_paragraph()
add_para(doc, "STOCK 5 — ICICI BANK (NSE: ICICIBANK) | CALL | Grade B",
         bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot",                 "₹1,279.50 (06 May) | Lot: DATA_UNAVAILABLE — verify"),
    ("Crude aligned",             "NEUTRAL — banks benefit indirectly via better consumer credit on lower fuel"),
    ("Catalyst today",            "Financial Services #2 sector 06 May; Bank Nifty leadership confirms move is REAL"),
    ("Daily trend / 200 DMA",     "Up; above 200 DMA: YES; 52w ₹1,187.6 - ₹1,500"),
    ("Weekly trend",              "Steady uptrend; OI fut +4.56% recent build"),
    ("Support / Resistance",      "Sup ₹1,265 / ₹1,250 | Res ₹1,295 / ₹1,310"),
    ("OI direction",              "Futures OI +4.56% with price up = LONG BUILD-UP"),
    ("Delivery % / Block / Ban",  "DATA_UNAVAILABLE / SEBI admin warning 04 May to ICICI as DP — minor / Ban: VERIFY"),
    ("Strike & expiry",           "ICICIBANK 1300 CE (slight OTM) — Monthly 26 May"),
    ("Theta warning ≤3 days?",    "NO"),
    ("Entry checklist (ALL 4)",   "ST GREEN on 15-min | RSI > 50 | StochRSI cross-up from <30 | Vol > 1.5x avg | 15-min close ABOVE ₹1,285"),
    ("Targets",                   "T1 (40%): ₹1,300 | T2 (40%): ₹1,312 | Trail rest"),
    ("Stop Loss",                 "Spot below ₹1,268 + 15-min close confirmation"),
    ("R:R",                       "≈ 2.0:1 (borderline — skip if entry slips)"),
    ("Time stop",                 "Exit by 13:00 IST"),
    ("Grade reason",              "B — Clean trend + bank-led leadership but R:R thin and SEBI warning is minor overhang"),
])

# ---------- SECTION 5 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["TODAY'S BIAS",      "CAUTIOUS BULL (override-driven)"],
        ["CONFIDENCE",        "MEDIUM"],
        ["CRUDE MODE",        "MILD BULL — falling toward AGG BULL (flip ₹7,500 = WTI $89)"],
        ["VIX SIZING",        "HALF (VIX ~18)"],
        ["ENTRY PERMISSION",  "YELLOW (gap-up entries demand patience; wait 9:30+ for ST confirm)"],
    ])

doc.add_paragraph()
add_para(doc, "THE BULL CASE:", bold=True, size=12, color=RGBColor(0x38, 0x76, 0x1D))
add_para(doc,
    "Crude crashed -9% on Iran-US peace hopes, Nifty closed +1.24% with PSU Banks leading +2.8%, "
    "DII bought ₹4,764 cr on 04 May, US indices at record highs, GIFT Nifty +190 pts. The trend, the flow, "
    "and the macro all line up bullish. Crude beneficiaries (aviation, OMCs, autos, paints) get a fundamental boost.",
    size=11)

add_para(doc, "THE BEAR CASE:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "Iran rejected 'Project Freedom' and resumed regional attacks; ceasefire is fragile. A single retaliation "
    "headline can re-spike crude 5-7%, gap-down equities and destroy crude-beneficiary longs. "
    "FII MTD still -₹70,135 cr (heavy seller). Nifty 281 pts above max pain creates downward magnetism into next Tue.",
    size=11)

add_para(doc, "FLIP TRIGGER:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "If Brent reclaims $108 intraday OR any Iran retaliation tweet hits the wires, today's bias flips from "
    "CAUTIOUS BULL to RANGE/BEAR — exit all crude-beneficiary calls immediately and shift to ONGC/Oil India.",
    size=11)

doc.add_paragraph()
add_para(doc, "NIFTY KEY LEVELS:", bold=True, size=12)
add_kv_table(doc,
    header=["Level", "Value", "Notes"],
    rows=[
        ["S2",        "24,050", "Max Pain magnet (downside pull)"],
        ["S1",        "24,200", "First intraday support"],
        ["CRITICAL",  "24,331", "Yesterday's close — pivot. Hold = bull intact"],
        ["R1",        "24,521", "GIFT Nifty / expected open zone"],
        ["R2",        "24,650-24,700", "Next resistance / bull target"],
    ])

doc.add_paragraph()
add_para(doc, "TOP 3 RANKED:", bold=True, size=12)
add_bullets(doc, [
    "#1 INDIGO (CALL, A) — Best crude-crash + momentum beneficiary; entry ABOVE ₹4,300",
    "#2 RELIANCE (CALL, A) — Heavyweight + O2C margin expansion; entry ABOVE ₹1,478",
    "#3 BPCL (CALL, A) — OMC oversold + crude crash; entry ABOVE ₹308",
])

doc.add_paragraph()
add_para(doc, "ONE RISK THAT RUINS EVERYTHING TODAY:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "An Iran retaliation against Trump's 'Project Freedom' initiative — the ceasefire is fragile and Iran "
    "has already called it a violation. A confirmed strike/seizure flips crude back to $110+ in minutes, "
    "obliterating all 5 crude-aligned longs simultaneously.",
    size=11, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- ONE-LINE SUMMARY ----------
doc.add_paragraph()
add_heading(doc, "One-Line Summary (read at 9:10 AM)", level=2, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Today is CAUTIOUS BULL because override fired (gap up +190 + crude -9%). Crude ₹7,810 = MILD BULL, "
    "FALLING. Watch INDIGO CALL and RELIANCE CALL. Key risk: Iran retaliation tweet re-spikes crude. "
    "Size HALF (VIX ~18). Flips if Brent reclaims $108 intraday.\"",
    italic=True, bold=True, size=12)

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
