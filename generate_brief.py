"""Generate the daily trading-intelligence .docx brief."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-12"
BIAS = "BEAR"
CONFIDENCE = "MEDIUM"
ENTRY_PERMISSION = "YELLOW"

OUT_PATH = f"trading-briefs/2026/05-May/Trading_Brief_{DATE_ISO}_{BIAS}.docx"
OUT_PATH_FLAT = f"Trading_Briefs/Trading_Brief_{DATE_ISO}_{BIAS}.docx"


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
dr = date_p.add_run("Date: Tuesday, 12 May 2026  |  Generated 08:50 IST  |  Weekly expiry: 2 days  |  Monthly expiry: 16 days  |  Expiry week: YES")
dr.bold = True
dr.font.size = Pt(10)

# Bias + permission badges
badge_line(doc, "TODAY'S BIAS", f"{BIAS} (Confidence: {CONFIDENCE})", "C00000")     # red
badge_line(doc, "ENTRY PERMISSION", ENTRY_PERMISSION, "E69138")                       # orange
badge_line(doc, "CRUDE MODE", "CAUTION (₹8,624 MCX) — Direction: RISING", "B45F06")  # dark orange
badge_line(doc, "VIX SIZING", "HALF SIZE (mandatory)", "F1C232")                      # amber
badge_line(doc, "EXPIRY ALERT", "Weekly expiry Thu — MONTHLY OPTIONS ONLY", "990000")

doc.add_paragraph()

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Value", "Signal"],
    rows=[
        ["MCX Crude (proxy)", "₹8,624/bbl", "CAUTION zone — bear tilt"],
        ["Brent / WTI", "$104.97 / $98.19", "Brent +0.73%, RISING; up ~3% prior session"],
        ["GIFT Nifty gap", "-195 pts (vs 24,259.5)", "GAP DOWN — bearish open"],
        ["S&P 500", "7,412.84 (+0.19% May 10; +0.38% May 11 intraday)", "Risk-on globally; mild divergence"],
        ["Nasdaq", "26,274 (+0.10% to +0.30%)", "Tech firm; INFY ADR diverges"],
        ["US VIX", "~17 (elevated on Iran)", "Fear — calm-side"],
        ["India VIX", "DATA_UNAVAILABLE (estimated 17–20)", "Apply HALF SIZE rule (mandatory)"],
        ["Nifty close (11 May)", "₹24,259.5", "Below 24,400 cap; weak structure"],
        ["FII cash (5–7 May)", "Net SELL ₹341 cr (May 7)", "Bearish — MTD FII –₹6,962 cr"],
        ["DII cash (May 7)", "Net BUY ₹441 cr (MTD +₹14,645 cr)", "Absorbing — supportive but lagging"],
        ["USD-INR", "~₹88 (rupee weak under Hormuz risk)", "FII caution amplified"],
        ["INFY ADR", "$12.52 (-2.42% vs prev $12.83)", "INFY likely opens ~₹1,147 (down ~2.4%)"],
        ["WIT ADR", "$1.95 (flat/marginal)", "Wipro neutral-to-soft open"],
        ["IBN ADR", "$26.46 (mild)", "ICICI Bank flat-to-soft"],
        ["HDB ADR", "$25.02 (flat)", "HDFC Bank neutral open"],
    ])

add_para(doc, "Geopolitical one-liner: Strait of Hormuz CLOSED since 4 Mar 2026 (Iran declared closure). Brent eased from $114 peak to $105 but supply loss ~880M bbl YTD. Saudi Aramco CEO: oil won't normalise before 2027 if Hormuz persists.", bold=False, size=10)
add_para(doc, "Intraday headline risk: HIGH (single tweet/strike can re-spike crude 5–8%). Direction of risk: crude spike on any escalation.", bold=True, size=10, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- SECTION 2 — CRUDE RULE + STRUCTURAL READ ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule + Structural Read", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "CRUDE RULE APPLIED", bold=True, size=12)
add_bullets(doc, [
    "MCX Crude proxy ≈ WTI $98.19 × USD-INR ₹88 ≈ ₹8,624/bbl → CAUTION band (₹8,500–9,000).",
    "Direction: RISING (Brent +0.73% overnight, +~3% prior session; Hormuz unresolved).",
    "Mode is STRENGTHENING toward bear — flag: Crude rising, bear mode deepening.",
    "CRUDE_FLIP_LEVEL: ₹9,000 MCX above = STRONG BEAR (no calls); ₹8,500 below = MILD BULL relief.",
    "AVOID sectors: OMCs (BPCL/HPCL/IOC), aviation, paints, tyres, fertilisers, OEM autos with high logistics share.",
    "EXCEPTION CALLS permitted: ONGC, Oil India, MRPL, Reliance (E&P leg), upstream PSUs.",
])

doc.add_paragraph()
add_para(doc, "STRUCTURAL READ — Nifty Options & Flows", bold=True, size=12)
add_bullets(doc, [
    "Nifty 24,259 vs Max Pain 24,250 → spot ~9 pts above max pain. Magnetism = mild DOWN pull into Thu expiry.",
    "Highest Call OI (ceiling): around ₹24,500 (call writers active).",
    "Highest Put OI (floor): around ₹24,000 (put writers defending).",
    "PCR: DATA_UNAVAILABLE for live read — historical bias in crisis weeks tilts >1.1 (put writers defend floor).",
    "Days to weekly expiry: 2 (Thu 14 May). ≤3 → MONTHLY EXPIRY MANDATORY for option buying.",
    "Expiry week dynamics: range-bound bias 24,000–24,500. Below 24,000 = panic unwind; above 24,500 = short squeeze.",
    "Bank Nifty: track open relative to Nifty — if Bank Nifty LAGS Nifty downside, treat fall as fake; if it LEADS, real breakdown.",
    "Futures basis: typically narrow-discount in pre-expiry crisis weeks (bearish skew).",
])

# ---------- SECTION 3 — CONTRARIAN CHECK ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Contrarian Check (Layer 3)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Question", "Answer"],
    rows=[
        ["Q1. CONSENSUS",
         "Every TV analyst expects Nifty to slide toward 24,000 on Hormuz fear + weak INFY ADR + FII selling. Headlines scream 'BEAR'."],
        ["Q2. THE TRAP",
         "If Iran signals back-channel talks or oil dips intraday on inventory headlines, the 24,000 longs get faced with a sharp 200-pt squeeze higher. Put writers defending 24,000 profit; consensus put buyers get fried."],
        ["Q3. RETAIL STOPS",
         "Long stops clustered BELOW ₹24,100 (psychological); short stops ABOVE ₹24,450 (last failed bounce). Magnet: 24,400 round number."],
        ["Q4. FLIP TRIGGER",
         "Any credible Iran de-escalation headline or partial Hormuz reopening = crude -5% → Nifty rips 250 pts into 24,500 by 12:00. Or, conversely, fresh strike on tanker = crude +5%, Nifty caves to 23,950 instantly."],
    ])

doc.add_paragraph()
add_para(doc, "CONTRARIAN OVERRIDE CHECK", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_bullets(doc, [
    "A. GIFT Nifty gap-up > +100 pts?  NO (gap DOWN −195)",
    "B. Crude falling > 2% overnight?  NO (crude RISING)",
    "C. Contrarian scenario probability > 40%?  NO (~30%)",
    "D. Expiry week + PCR < 0.8?  UNKNOWN (PCR data unavailable; assume >1.0 in crisis)",
    "→ No override conditions met. BEAR bias STANDS. Confidence remains MEDIUM (not HIGH) due to single-headline reversal risk.",
])

# ---------- SECTION 4 — FIVE INDIVIDUAL F&O SETUPS ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — 5 Individual F&O Setups", level=1, color=RGBColor(0x1F, 0x38, 0x64))

# ---- SETUP 1 ----
add_para(doc, "SETUP 1 — BPCL (NSE: BPCL)  |  PUT  |  Grade A+", bold=True, size=13, color=RGBColor(0xC0, 0x00, 0x00))
stock_card(doc, [
    ("CMP", "₹297.00 (-2% on 11 May; closed weak)"),
    ("Lot Size", "1,800"),
    ("Crude Aligned", "YES (OMC = direct crude victim per rule)"),
    ("Catalyst", "Q4 FY26 results TODAY 10:00 AM IST — refining margins under crude squeeze; PM Modi has called for fuel-consumption cut."),
    ("Technical — Daily", "DOWN (-18.7% over 6m); below 50 DMA ₹309.58"),
    ("Above 200 DMA", "NO — below long-term trend; institutional outflow signal"),
    ("Support / Resistance", "Support ₹294 / ₹287  |  Resistance ₹303 / ₹310"),
    ("Structural — OI", "SHORT BUILD (price down + OI rising in Apr-May futures)"),
    ("Delivery %", "STABLE-FALLING; weak hand selling"),
    ("Ban list / Block deal", "NO ban / No reported block deal"),
    ("Option Setup", "Buy 28 May 300 PE (ATM, monthly expiry) — strike near CMP for delta. AVOID weekly 14 May PE (≤3 days = theta risk)."),
    ("Theta Warning ≤3d", "Skipping weekly (only monthly per rule)"),
    ("Entry Checklist (ALL 4)",
     "ST RED on 15-min  |  RSI(14) below 50 on 15-min  |  StochRSI cross down from >70  |  Volume > 1.5× 20-period avg. Confirmation: 15-min close BELOW ₹295."),
    ("Target T1 (40%)", "₹287 (premium ~+45%)"),
    ("Target T2 (40%)", "₹278 (premium ~+90%)"),
    ("Stop Loss", "Spot above ₹303.5 AND 15-min ST flips GREEN (combined trigger)"),
    ("R:R", "~2.5:1 (acceptable)"),
    ("Time Stop", "Exit by 13:00 IST regardless (event-day vol decay risk)"),
    ("Grade Reason", "Q4 catalyst TODAY + clear crude-victim alignment + technical breakdown confirmed = A+ conviction"),
])

doc.add_paragraph()
# ---- SETUP 2 ----
add_para(doc, "SETUP 2 — ONGC (NSE: ONGC)  |  CALL  |  Grade A", bold=True, size=13, color=RGBColor(0x00, 0x66, 0x00))
stock_card(doc, [
    ("CMP", "₹282.80 (May 7 reference; 52w high ₹307.50)"),
    ("Lot Size", "2,250"),
    ("Crude Aligned", "YES (upstream = crude beneficiary per rule exception)"),
    ("Catalyst", "Elevated Brent ($105) + Hormuz supply loss + JM Financial Buy ₹340 target; Board meet 26 May (FY26 results + dividend)"),
    ("Technical — Daily", "UP — breakout above ₹270-272 resistance with volume; pullback held"),
    ("Above 200 DMA", "YES — bullish institutional structure"),
    ("Support / Resistance", "Support ₹278 / ₹272  |  Resistance ₹290 / ₹300"),
    ("Structural — OI", "LONG BUILD (price up + OI rising)"),
    ("Delivery %", "RISING — strong-hand accumulation"),
    ("Ban list / Block deal", "NO ban / no reported block"),
    ("Option Setup", "Buy 28 May 285 CE (ATM/slightly OTM monthly) — avoid weekly per ≤3d rule"),
    ("Theta Warning ≤3d", "No — monthly used"),
    ("Entry Checklist (ALL 4)",
     "ST GREEN on 15-min  |  RSI(14) above 50  |  StochRSI cross up from <30  |  Volume > 1.5× 20-period avg. Confirmation: 15-min close ABOVE ₹284."),
    ("Target T1 (40%)", "₹290 (premium ~+50%)"),
    ("Target T2 (40%)", "₹297 (premium ~+95%)"),
    ("Stop Loss", "Spot below ₹278 AND 15-min ST flips RED"),
    ("R:R", "~2.3:1"),
    ("Time Stop", "13:00 IST exit; if Iran peace headline drops crude -3%, exit IMMEDIATELY (mode flip)"),
    ("Grade Reason", "Best crude-beneficiary in bear-crude regime; aligned with macro; A (not A+ due to reversal-on-headline risk)"),
])

doc.add_paragraph()
# ---- SETUP 3 ----
add_para(doc, "SETUP 3 — InterGlobe Aviation / IndiGo (NSE: INDIGO)  |  PUT  |  Grade A", bold=True, size=13, color=RGBColor(0xC0, 0x00, 0x00))
stock_card(doc, [
    ("CMP", "₹4,323.40 (-4.4% on 11 May from ₹4,522.70)"),
    ("Lot Size", "150"),
    ("Crude Aligned", "YES (aviation = crude victim per rule)"),
    ("Catalyst", "ATF prices linked to Brent ($105) = direct margin compression; +$10 Brent ≈ +₹3,500-4,000 cr fuel bill"),
    ("Technical — Daily", "DOWN — confirmed downtrend"),
    ("Above 200 DMA", "NO (200 DMA ₹5,221.75; spot ₹4,323) — major weakness"),
    ("Support / Resistance", "Support ₹4,300 / ₹4,200  |  Resistance ₹4,435 (50 DMA) / ₹4,522"),
    ("Structural — OI", "SHORT BUILD (sustained price decline + OI rising)"),
    ("Delivery %", "FALLING — distribution"),
    ("Ban list / Block deal", "NO ban"),
    ("Option Setup", "Buy 28 May 4300 PE (ATM monthly)"),
    ("Theta Warning ≤3d", "No — monthly used"),
    ("Entry Checklist (ALL 4)",
     "ST RED on 15-min  |  RSI below 50  |  StochRSI cross down from >70  |  Volume > 1.5× avg. Confirmation: 15-min close BELOW ₹4,300."),
    ("Target T1 (40%)", "₹4,220 (premium ~+40%)"),
    ("Target T2 (40%)", "₹4,150 (premium ~+85%)"),
    ("Stop Loss", "Spot above ₹4,435 (50 DMA reclaim) AND ST flips GREEN"),
    ("R:R", "~2.4:1"),
    ("Time Stop", "13:00 IST"),
    ("Grade Reason", "Pristine crude-victim + clean technical breakdown + below both DMAs; A grade because stretched short — short-squeeze risk on Iran peace headline keeps it from A+"),
])

doc.add_paragraph()
# ---- SETUP 4 ----
add_para(doc, "SETUP 4 — Infosys (NSE: INFY)  |  PUT  |  Grade A", bold=True, size=13, color=RGBColor(0xC0, 0x00, 0x00))
stock_card(doc, [
    ("CMP", "₹1,175.70 close 11 May; expected open ~₹1,147 (ADR -2.4%)"),
    ("Lot Size", "400"),
    ("Crude Aligned", "NEUTRAL (IT sector; not crude-sensitive). Driver = ADR-led gap-down + IT-sector weakness."),
    ("Catalyst", "INFY ADR closed $12.52 (prev $12.83 = -2.42%); US tech rotation; rupee strength is mild headwind for exporters"),
    ("Technical — Daily", "DOWN (-21.57% YoY); 52w range 1,149-1,728"),
    ("Above 200 DMA", "NO — bearish structure"),
    ("Support / Resistance", "Support ₹1,150 / ₹1,128  |  Resistance ₹1,190 / ₹1,220"),
    ("Structural — OI", "SHORT BUILD on weekly futures"),
    ("Delivery %", "STABLE"),
    ("Ban list / Block deal", "NO ban"),
    ("Option Setup", "Buy 28 May 1160 PE (ATM monthly)"),
    ("Theta Warning ≤3d", "No — monthly used"),
    ("Entry Checklist (ALL 4)",
     "ST RED on 15-min  |  RSI below 50  |  StochRSI cross down from >70  |  Volume > 1.5× avg. Confirmation: 15-min close BELOW ₹1,150."),
    ("Target T1 (40%)", "₹1,135 (premium ~+45%)"),
    ("Target T2 (40%)", "₹1,118 (premium ~+85%)"),
    ("Stop Loss", "Spot above ₹1,190 AND ST flips GREEN"),
    ("R:R", "~2.2:1"),
    ("Time Stop", "13:00 IST"),
    ("Grade Reason", "Strongest ADR-confirmed direction (-2.4% overnight); A (not A+) because near 52w low — bounce-trap risk"),
])

doc.add_paragraph()
# ---- SETUP 5 ----
add_para(doc, "SETUP 5 — Bharat Electronics (NSE: BEL)  |  CALL  |  Grade B", bold=True, size=13, color=RGBColor(0x00, 0x66, 0x00))
stock_card(doc, [
    ("CMP", "₹438.15 (May 11)"),
    ("Lot Size", "475"),
    ("Crude Aligned", "POSITIVE (defence/geopolitical tailwind in Hormuz crisis)"),
    ("Catalyst", "Defence order momentum: ₹30,000 cr FY26 order intake; ₹1,251 cr Army GBMES order (5 May); Hormuz crisis = sustained defence spend tailwind"),
    ("Technical — Daily", "SIDEWAYS-UP near 52w high ₹473.45"),
    ("Above 200 DMA", "YES (200 DMA ₹413.83) — institutional bullish"),
    ("Support / Resistance", "Support ₹430 / ₹420  |  Resistance ₹450 / ₹465"),
    ("Structural — OI", "NEUTRAL-to-LONG BUILD"),
    ("Delivery %", "STABLE"),
    ("Ban list / Block deal", "NO ban"),
    ("Option Setup", "Buy 28 May 440 CE (ATM monthly)"),
    ("Theta Warning ≤3d", "No — monthly used"),
    ("Entry Checklist (ALL 4)",
     "ST GREEN on 15-min  |  RSI above 50  |  StochRSI cross up from <30  |  Volume > 1.5× avg. Confirmation: 15-min close ABOVE ₹442."),
    ("Target T1 (40%)", "₹450 (premium ~+40%)"),
    ("Target T2 (40%)", "₹460 (premium ~+75%)"),
    ("Stop Loss", "Spot below ₹430 AND ST flips RED"),
    ("R:R", "~2.0:1 (borderline)"),
    ("Time Stop", "13:00 IST"),
    ("Grade Reason", "B grade — extended near 52w high; overbought; needs trigger; thesis intact but better as add-on, not lead trade"),
])

# ---------- SECTION 5 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["Today's Bias", "BEAR"],
        ["Confidence", "MEDIUM (single-headline reversal risk)"],
        ["Crude Mode", "CAUTION (₹8,624 MCX, RISING)"],
        ["VIX Sizing", "HALF SIZE (mandatory)"],
        ["Entry Permission", "YELLOW (proceed selectively, half-size)"],
        ["Expiry Constraint", "MONTHLY EXPIRY ONLY (weekly ≤3 days)"],
        ["Max Trades", "3 max (per rulebook)"],
        ["Daily Risk Cap", "2% of capital"],
        ["Hard Exit", "13:30 IST — no exceptions"],
    ])

doc.add_paragraph()
add_para(doc, "THE BULL CASE (write even if bias = BEAR):", bold=True, size=11, color=RGBColor(0x00, 0x66, 0x00))
add_para(doc, "DII relentlessly absorbing FII selling (+₹14,645 cr MTD); US S&P at record; rupee weakness aids exporters once panic abates; any whiff of Iran de-escalation collapses the crude premium and triggers a sharp short-squeeze rip toward 24,500.", size=10)

add_para(doc, "THE BEAR CASE (primary view):", bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc, "Hormuz closed since March (no resolution path); Brent stuck >$100 with 880M bbl supply loss; FII selling (-₹6,962 cr MTD); GIFT Nifty gap DOWN -195; INFY ADR -2.4%; Bank Nifty leadership uncertain. Max pain 24,250 magnetism toward downside.", size=10)

add_para(doc, "FLIP TRIGGER:", bold=True, size=11, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc, "If Brent breaks below $100 intraday on credible Iran/Hormuz de-escalation OR Nifty reclaims ₹24,400 with strong Bank Nifty leadership, bias flips to CAUTIOUS BULL — close all puts immediately, switch to ONGC/BEL calls only after fresh confirmation.", size=10)

doc.add_paragraph()
add_para(doc, "NIFTY KEY LEVELS", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_kv_table(doc, rows=[
    ["S2", "₹23,950 (panic-unwind target)"],
    ["S1", "₹24,000 (put-writer floor)"],
    ["CRITICAL (Max Pain)", "₹24,250"],
    ["R1", "₹24,400 (failed cap)"],
    ["R2", "₹24,500 (call-writer ceiling — squeeze trigger above)"],
])

doc.add_paragraph()
add_para(doc, "TOP 3 RANKED TRADES", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "#1 BPCL — A+ — Q4 result-day put on crude victim — entry on 15-min close BELOW ₹295",
    "#2 ONGC — A — Crude-beneficiary call (lone bull idea) — entry on 15-min close ABOVE ₹284",
    "#3 INDIGO — A — Aviation put on ATF squeeze — entry on 15-min close BELOW ₹4,300",
])

doc.add_paragraph()
add_para(doc, "ONE RISK THAT RUINS EVERYTHING TODAY", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc, "A single Iran-US back-channel headline (Reuters/Bloomberg) crashing Brent -5% intraday. Bear setups detonate, ONGC/Oil India crash, INDIGO/BPCL rip higher. This is a binary tail risk — keep stops tight and exit aggressively if you see any peace-related wire.", size=10)

doc.add_paragraph()
add_para(doc, "ONE-LINE SUMMARY (read at 9:10 AM)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Today is BEAR because GIFT Nifty gap-DOWN 195 + Hormuz still closed + crude rising + INFY ADR -2.4%. "
    "Crude ₹8,624 = CAUTION mode, RISING. Watch BPCL PUT and ONGC CALL. Key risk: any Iran de-escalation headline. "
    "Size HALF. Flips if Brent breaks $100 or Nifty reclaims 24,400.\"",
    italic=True, size=11)

# ---------- DATA NOTES ----------
doc.add_paragraph()
add_heading(doc, "Data Notes & Source Caveats", level=2, color=RGBColor(0x80, 0x80, 0x80))
add_bullets(doc, [
    "Brent live: $104.97 (tradingeconomics.com, May 12). Earlier May 11 peak quoted near $114; eased.",
    "GIFT Nifty: 24,064 referenced (icicidirect/5paisa) implying -195 gap vs 24,259.5 close.",
    "ADR data: INFY $12.52 (prev $12.83); HDB $25.02 (flat); IBN $26.46; WIT $1.95 (Wise, Yahoo).",
    "India VIX live value DATA_UNAVAILABLE; treated as 17–20 (HALF SIZE) given Hormuz regime — flagged.",
    "PCR live value DATA_UNAVAILABLE; biased >1.0 in crisis week per historical pattern — flagged.",
    "FII/DII data anchored to May 7 (latest verified); MTD figures from NSDL aggregates.",
    "INFY Q4 already reported (PAT +27.75%); today's INFY weakness is ADR/sector-rotation driven, not earnings.",
])

# ---------- FOOTER ----------
doc.add_paragraph()
doc.add_paragraph()
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run("Generated by Daily Trading Routine — for personal educational use only. Not financial advice.")
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

import os
os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
os.makedirs(os.path.dirname(OUT_PATH_FLAT), exist_ok=True)
doc.save(OUT_PATH)
doc.save(OUT_PATH_FLAT)
print(f"Wrote {OUT_PATH}")
print(f"Wrote {OUT_PATH_FLAT}")
