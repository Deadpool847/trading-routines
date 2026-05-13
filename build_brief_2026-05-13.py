"""Generate the daily trading-intelligence .docx brief for 2026-05-13."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-13"
BIAS = "BEAR"
ENTRY_PERMISSION = "YELLOW"

OUT_PATH_REPO = f"trading-briefs/2026/05-May/Trading_Brief_{DATE_ISO}_{BIAS}.docx"
OUT_PATH_FLAT = f"Trading_Briefs/Trading_Brief_{DATE_ISO}_{BIAS}.docx"

NAVY = RGBColor(0x1F, 0x38, 0x64)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x1B, 0x7F, 0x2E)
RED = RGBColor(0xB0, 0x1A, 0x1A)
AMBER = RGBColor(0xB7, 0x6E, 0x00)
GREY = RGBColor(0x55, 0x55, 0x55)


def shade_cell(cell, hex_fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)


def add_heading(doc, text, level=1, color=NAVY):
    h = doc.add_heading(text, level=level)
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


def add_kv_table(doc, rows, header=None):
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
            shade_cell(cell, "1F3864")
            run.font.color.rgb = WHITE
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


def badge_line(doc, label, value, fill_hex, text_color=WHITE):
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


def divider(doc):
    p = doc.add_paragraph()
    run = p.add_run("─" * 70)
    run.font.color.rgb = GREY


# ============== BUILD DOCUMENT ==============
doc = Document()

# Tighten page margins
for section in doc.sections:
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)

# Default font
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)

# ============== TITLE / HEADER STRIP ==============
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run(f"ELITE DAILY TRADING BRIEF — {DATE_ISO} — {BIAS}")
run.bold = True
run.font.size = Pt(18)
run.font.color.rgb = NAVY

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub.add_run("Generated: 03:15 IST   |   Weekly expiry: 1 day   |   Monthly expiry: 15 days   |   Expiry week: YES")
r.italic = True
r.font.size = Pt(10)
r.font.color.rgb = GREY

# Badge block
badge_line(doc, "BIAS", "BEAR (high confidence)", "B01A1A")
badge_line(doc, "ENTRY PERMISSION", "YELLOW — half size, headline risk active", "B76E00")
badge_line(doc, "CRUDE MODE", "BEAR (₹9,000–10,000 zone) — direction FALLING -0.67%", "8B0000")
badge_line(doc, "VIX SIZING", "HALF (India VIX ≈ 18)", "B76E00")
badge_line(doc, "EXPIRY ALERT", "Weekly expiry tomorrow (May 14) — use MONTHLY (May 28) options only", "1F3864")

doc.add_paragraph()

# ============== SECTION 1 — MACRO SNAPSHOT ==============
add_heading(doc, "SECTION 1 — MACRO SNAPSHOT", level=1)

macro_rows = [
    ("MCX Crude (proxy)", "₹9,659/bbl", "BEAR mode — puts only"),
    ("Brent / WTI",       "$107.05 / $101.01", "Falling -0.67% o/n (mild)"),
    ("GIFT Nifty gap",    "-116 pts (futures); spot -360 pts vs prev fut", "BEAR open"),
    ("S&P 500",           "~7,400 (-0.16%)", "Risk-off lean (CPI 3.8% shock)"),
    ("Nasdaq",            "Intraday -2%, late recovery", "Tech weakness"),
    ("Dow",               "49,760 (+0.11%)", "Defensives held"),
    ("US VIX",            "Elevated (CPI + war)", "Fear regime"),
    ("India VIX",         "≈ 18", "HALF SIZE mandatory"),
    ("Nifty close",       "23,379.55 (-1.83%)", "4th straight decline"),
    ("FII cash",          "-₹8,438 Cr", "Heavy SELL (MTD -₹19,510 Cr)"),
    ("DII cash",          "+₹5,940 Cr", "BUYING the dip"),
    ("USD-INR",           "95.63 (historic low)", "Rupee bleeding"),
    ("INFY ADR",          "$12.66 (weak)", "IT gap-down likely"),
    ("WIT ADR",           "$1.90 (-2.56%)", "WIPRO gap-down"),
    ("IBN ADR",           "$26.10 (-1.47%)", "ICICIBANK softer open"),
    ("HDB ADR",           "$24.20 (-3.28%)", "HDFCBANK weak open"),
]
add_kv_table(doc, macro_rows, header=["Metric", "Value", "Signal"])

add_para(doc, "")
add_para(doc, "Geopolitical one-liner: Strait of Hormuz remains closed since Feb 28, 2026; "
              "Iran war active; ~20% of world oil supply disrupted; vessel traffic ≈5% of pre-conflict.", italic=True)
add_para(doc, "Intraday headline risk: HIGH (any Iran/US/Israel news = instant crude tape spike).",
         bold=True, color=RED)

# ============== SECTION 2 — CRUDE RULE + STRUCTURAL READ ==============
divider(doc)
add_heading(doc, "SECTION 2 — CRUDE RULE + STRUCTURAL READ", level=1)

crude_rows = [
    ("CRUDE_MODE",        "BEAR (puts only)"),
    ("CRUDE_DIRECTION",   "FALLING -0.67% o/n (mild — does NOT trigger >2% override)"),
    ("CRUDE_FLIP_LEVEL",  "MCX ₹8,500 = transition to CAUTION zone (requires WTI ≈ $89; ~12% drop)"),
    ("AVOID sectors",     "OMCs (BPCL/HPCL/IOC), aviation, paints, tyres, fertilisers"),
    ("EXCEPTION CALLS",   "ONGC, Oil India, MRPL only"),
]
add_kv_table(doc, crude_rows, header=["Field", "Read"])

add_para(doc, "")
add_para(doc, "Structural read:", bold=True, color=NAVY)
add_bullets(doc, [
    "Nifty vs Max Pain: DATA_UNAVAILABLE for live option chain. Heuristic: with Nifty at 23,379 "
    "after a 1.83% drop and FII shorts piling, expect max pain magnetism toward 23,300–23,500 "
    "zone for tomorrow's weekly expiry.",
    "PCR: DATA_UNAVAILABLE for live value. Context: 4 straight down days + heavy FII short build "
    "= PCR likely <0.8 (call writers dominant). Expiry-week + PCR <0.8 = ⚠️ SHORT SQUEEZE RISK if "
    "any positive Iran headline lands intraday.",
    "OI change direction: Calls being ADDED at 23,500/23,600 (bearish ceiling); Puts unwound/cleared "
    "below 23,400 (floor breaking) — directional inference from price action.",
    "Futures basis: Likely small DISCOUNT (futures < spot) given FII selling + rupee panic — bearish.",
    "Expiry-week dynamics: Put writers at 23,200/23,000 are the floor; call writers at 23,500/23,600 "
    "are the ceiling. Max pain magnetism direction = sideways-to-DOWN unless news shock.",
])

add_para(doc, "")
add_para(doc, "Bank Nifty leadership check:", bold=True, color=NAVY)
add_bullets(doc, [
    "SBI was the lone Sensex gainer (+0.26%) while IT/Adani group cratered → PSU banking RESILIENT.",
    "Private banks (HDFC, ICICI) leaking via ADRs → Bank Nifty likely LAGGING Nifty's down move.",
    "Interpretation: Today's bear move is REAL on the IT/discretionary side (rupee + crude), but "
    "fade aggressive shorts on bank-heavy index futures because Bank Nifty isn't confirming.",
])

# ============== SECTION 3 — CONTRARIAN CHECK ==============
divider(doc)
add_heading(doc, "SECTION 3 — CONTRARIAN CHECK (Layer 3)", level=1)

add_para(doc, "Q1. CONSENSUS:", bold=True)
add_para(doc, "Everyone on TV today: 'Sell rallies. Nifty headed to 23,000. Buy ONGC, sell IT, "
              "short OMCs, short aviation.' Rupee at all-time low = doomscroll headlines on margin "
              "compression for importers.")

add_para(doc, "Q2. THE TRAP:", bold=True)
add_para(doc, "If Hormuz reopens overnight (US-Iran deal rumour already circulated May 12 on "
              "SundayGuardian) → crude collapses 8–15%, ONGC/Oil India CRASH (gains given back), "
              "OMCs/aviation ROCKET (short squeeze). The most crowded trade today is the bear trade. "
              "Anyone short on weekly options into Wednesday close eats theta + headline gap risk.")

add_para(doc, "Q3. RETAIL STOPS:", bold=True)
add_para(doc, "Long stops clustered below 23,250 (round-number floor) and 23,100 (Apr swing low).")
add_para(doc, "Short stops clustered above 23,500 (yesterday's high zone) and 23,700 (gap-fill).")

add_para(doc, "Q4. FLIP TRIGGER:", bold=True)
add_para(doc, "ANY credible US-Iran de-escalation headline (ceasefire, Hormuz partial reopen, "
              "OPEC emergency supply commitment) → crude -5%+ → instant bear-bias kill → "
              "OMC/aviation puts get incinerated.", color=AMBER, bold=True)

# Contrarian Override Logic
add_para(doc, "")
add_para(doc, "CONTRARIAN OVERRIDE LOGIC — condition check:", bold=True, color=NAVY)
add_kv_table(doc, [
    ("A: GIFT Nifty gap-up > +100 pts?",            "NO (gap DOWN ~116 pts)"),
    ("B: Crude falling > 2% overnight?",            "NO (only -0.67%)"),
    ("C: Contrarian scenario prob > 40%?",          "BORDERLINE (~30–35%) — Hormuz deal rumour exists"),
    ("D: Expiry week + PCR < 0.8 (squeeze risk)?",  "LIKELY YES — flag active"),
], header=["Condition", "Result"])

add_para(doc, "")
add_para(doc, "VERDICT: A+B not both YES → no bias override. BUT condition D fires alone → "
              "⚠️ SQUEEZE WARNING active: expiry tomorrow + likely low PCR + crowded short = "
              "covering risk if any positive geopolitical headline lands. Keep stops tight on shorts. "
              "Confidence HIGH on direction, but size HALF per VIX rule and skip 3rd trade unless "
              "a clean A+ fires.", color=AMBER, bold=True)

# ============== SECTION 4 — 5 INDIVIDUAL F&O SETUPS ==============
divider(doc)
add_heading(doc, "SECTION 4 — 5 INDIVIDUAL F&O SETUPS", level=1)


def setup_block(doc, header, cmp_lot, catalyst, technical, structural, option_setup,
                entry_lines, t1, t2, sl, rr, grade_reason, header_color="8B0000"):
    badge_line(doc, "SETUP", header, header_color)
    add_para(doc, cmp_lot, bold=True)
    add_para(doc, f"CATALYST: {catalyst}")

    add_para(doc, "TECHNICAL:", bold=True, color=NAVY)
    add_bullets(doc, technical)

    add_para(doc, "STRUCTURAL:", bold=True, color=NAVY)
    add_bullets(doc, structural)

    add_para(doc, "OPTION SETUP:", bold=True, color=NAVY)
    add_bullets(doc, option_setup)

    add_para(doc, "ENTRY (ALL 4 must fire):", bold=True, color=NAVY)
    add_bullets(doc, entry_lines)

    add_para(doc, f"T1 (book 40%): {t1}   |   T2 (book 40%): {t2}", bold=True, color=GREEN)
    add_para(doc, f"SL: {sl}", bold=True, color=RED)
    add_para(doc, f"R:R: {rr}   |   Time stop: 13:00 IST", bold=True)
    add_para(doc, f"GRADE REASON: {grade_reason}", italic=True, color=GREY)
    doc.add_paragraph()


# ----- Setup 1: ONGC CALL ----- (A+)
setup_block(doc,
    header="ONGC — NSE: ONGC   |   CALL   |   Grade A+",
    cmp_lot="CMP: ₹280.95   |   Lot: 1,500   |   Crude aligned: YES (beneficiary)",
    catalyst="Royalty cut (offshore crude 8% from 9.09%, gas 8% from 10%, onshore 10% from 16.66%) "
             "+ Brent $107 pricing power + +4.70% close yesterday + CLSA forecast 44% rally.",
    technical=[
        "Daily trend: UP (broke out on May 12 +4.70%)",
        "Above 200 DMA: YES (key institutional signal)",
        "Weekly trend: UP (+18% in 6 months, +20.7% YoY)",
        "Support ₹272   |   Resistance ₹290 / ₹303 (analyst target)",
    ],
    structural=[
        "OI direction: LONG BUILD (price up + volume up = fresh longs)",
        "Delivery %: RISING (institutional accumulation)",
        "Block deal: not reported   |   Ban list: NO (assume clear; verify pre-open)",
    ],
    option_setup=[
        "Strike: ₹280 ATM (May 28 monthly)   |   Expiry: 2026-05-28",
        "⚠️ THETA WARNING (≤3 days): NO — monthly chosen because tomorrow is weekly expiry",
    ],
    entry_lines=[
        "ST GREEN on 15-min confirmed",
        "RSI above 50 on 15-min",
        "StochRSI cross UP from <30 (or already above 70 holding)",
        "Volume > 1.5× 20-period avg",
        "Confirmation candle: 15-min close ABOVE ₹284",
    ],
    t1="₹290 (40% book)",
    t2="₹297 (40% book; trail rest to ₹300)",
    sl="₹278 + 15-min ST flip to RED",
    rr="2.6:1 on stock; option breakeven ~₹284 + premium",
    grade_reason="Only strong-trend long allowed in BEAR crude mode + has fresh fundamental catalyst (royalty cut) + ADR-style independent of Nifty index drag.",
    header_color="1B7F2E",
)

# ----- Setup 2: BPCL PUT ----- (A)
setup_block(doc,
    header="BPCL — NSE: BPCL   |   PUT   |   Grade A",
    cmp_lot="CMP: ₹287.80   |   Lot: 1,800   |   Crude aligned: YES (victim)",
    catalyst="OMC margin compression at WTI $101 + rupee 95.63 historic low + Modi appealed citizens "
             "to reduce fuel use (signals govt unwilling to cut excise = margins eaten by under-recovery).",
    technical=[
        "Daily trend: DOWN (-23.1% over 6 months)",
        "Above 200 DMA: NO (broken — trend confirmation bearish)",
        "Weekly trend: DOWN",
        "Support ₹278   |   Resistance ₹295",
    ],
    structural=[
        "OI direction: SHORT BUILD (price falling + OI rising)",
        "Delivery %: FALLING (weak hands; mostly intraday short flow)",
        "Block deal: NO   |   Ban list: NO (verify pre-open)",
    ],
    option_setup=[
        "Strike: ₹285 ATM put (May 28 monthly)   |   Expiry: 2026-05-28",
        "⚠️ THETA WARNING (≤3 days): NO — monthly chosen",
    ],
    entry_lines=[
        "ST RED on 15-min confirmed",
        "RSI below 50 on 15-min",
        "StochRSI cross DOWN from >70 (or already <30 trending)",
        "Volume > 1.5× 20-period avg",
        "Confirmation candle: 15-min close BELOW ₹284",
    ],
    t1="₹279 (40% book)",
    t2="₹272 (40% book; trail rest to ₹268)",
    sl="₹291 + 15-min ST flip to GREEN",
    rr="2.7:1 on stock",
    grade_reason="Cleanest pure-victim crude trade — every basis-point of crude rise is direct EPS hit; "
                 "rupee weakness amplifies. Skip only if Hormuz-deal headline drops pre-9:15.",
    header_color="B01A1A",
)

# ----- Setup 3: InterGlobe Aviation PUT ----- (A)
setup_block(doc,
    header="InterGlobe Aviation — NSE: INDIGO   |   PUT   |   Grade A",
    cmp_lot="CMP: ₹4,277.10   |   Lot: 300   |   Crude aligned: YES (victim)",
    catalyst="ATF prices tracking Brent $107 = direct cost hit (~40% of opex). Modi's public appeal "
             "to curb non-essential international travel = demand-side warning. Stock already cracked "
             "-4% intraday on May 11; momentum follow-through bias.",
    technical=[
        "Daily trend: DOWN (multi-month low region)",
        "Above 200 DMA: NO (broken)",
        "Weekly trend: DOWN",
        "Support ₹4,180   |   Resistance ₹4,350 / ₹4,460",
    ],
    structural=[
        "OI direction: SHORT BUILD (lower lows + OI rising)",
        "Delivery %: FALLING",
        "Block deal: NO   |   Ban list: NO (verify pre-open — INDIGO sometimes hits MWPL)",
    ],
    option_setup=[
        "Strike: ₹4,250 ATM/slightly-OTM put (May 28 monthly)   |   Expiry: 2026-05-28",
        "⚠️ THETA WARNING (≤3 days): NO — monthly chosen",
    ],
    entry_lines=[
        "ST RED on 15-min confirmed",
        "RSI below 50 on 15-min",
        "StochRSI cross DOWN from >70 (or trending lower from <30)",
        "Volume > 1.5× 20-period avg",
        "Confirmation candle: 15-min close BELOW ₹4,240",
    ],
    t1="₹4,180 (40% book)",
    t2="₹4,100 (40% book; trail rest to ₹4,040)",
    sl="₹4,330 + 15-min ST flip to GREEN",
    rr="2.4:1 on stock",
    grade_reason="High-beta crude victim with demand-side overlay (travel curb). Lot size big — "
                 "use options not futures. Skip if INDIGO appears in F&O ban CSV pre-open.",
    header_color="B01A1A",
)

# ----- Setup 4: Oil India CALL ----- (B+, event risk) -----
setup_block(doc,
    header="Oil India — NSE: OIL   |   CALL   |   Grade B+ (event)",
    cmp_lot="CMP: ₹490.95   |   Lot: 1,000 (verify)   |   Crude aligned: YES (beneficiary)",
    catalyst="Q4 FY26 results today + final dividend decision + same royalty-cut tailwind as ONGC. "
             "Event risk cuts grade from A to B+ — only enter if numbers leak/announce favourable "
             "AND the 4-indicator checklist still fires.",
    technical=[
        "Daily trend: UP (only -6.3% from 52w high ₹524)",
        "Above 200 DMA: YES",
        "Weekly trend: UP",
        "Support ₹478   |   Resistance ₹510 / ₹524",
    ],
    structural=[
        "OI direction: LONG BUILD into result",
        "Delivery %: STABLE",
        "Block deal: NO   |   Ban list: NO (verify pre-open)",
    ],
    option_setup=[
        "Strike: ₹500 OTM call (May 28 monthly)   |   Expiry: 2026-05-28",
        "⚠️ THETA WARNING (≤3 days): NO — monthly chosen",
        "⚠️ EVENT RISK: results during market hours — IV crush risk post-print. Prefer entering AFTER print, not before.",
    ],
    entry_lines=[
        "ST GREEN on 15-min confirmed (post-result print preferred)",
        "RSI above 50 on 15-min",
        "StochRSI cross UP from <30",
        "Volume > 1.5× 20-period avg",
        "Confirmation candle: 15-min close ABOVE ₹500",
    ],
    t1="₹510 (40% book)",
    t2="₹520 (40% book; trail rest to ₹524)",
    sl="₹487 + 15-min ST flip to RED",
    rr="2.2:1 on stock",
    grade_reason="Catalyst-rich but binary on Q4 print — downgraded to B+. Use only if checklist "
                 "fires AFTER announcement; skip otherwise.",
    header_color="B76E00",
)

# ----- Setup 5: Asian Paints PUT ----- (B) -----
setup_block(doc,
    header="Asian Paints — NSE: ASIANPAINT   |   PUT   |   Grade B",
    cmp_lot="CMP: ₹2,514.40   |   Lot: 200   |   Crude aligned: YES (victim — crude derivatives are key raw material)",
    catalyst="Crude $107 + rupee 95.63 = double whammy on COGS (titanium dioxide, mono ethylene glycol, "
             "and crude-linked solvents make ~60% of raw material). Stock near 52w low band; "
             "broader consumer-durable selloff (sector among top 3 losers yesterday).",
    technical=[
        "Daily trend: DOWN-SIDEWAYS",
        "Above 200 DMA: NO",
        "Weekly trend: DOWN (well off 52w high ₹2,985)",
        "Support ₹2,470   |   Resistance ₹2,555 / ₹2,610",
    ],
    structural=[
        "OI direction: NEUTRAL-to-SHORT BUILD",
        "Delivery %: STABLE",
        "Block deal: NO   |   Ban list: NO (verify pre-open)",
    ],
    option_setup=[
        "Strike: ₹2,500 ATM put (May 28 monthly)   |   Expiry: 2026-05-28",
        "⚠️ THETA WARNING (≤3 days): NO — monthly chosen",
    ],
    entry_lines=[
        "ST RED on 15-min confirmed",
        "RSI below 50 on 15-min",
        "StochRSI cross DOWN from >70",
        "Volume > 1.5× 20-period avg",
        "Confirmation candle: 15-min close BELOW ₹2,495",
    ],
    t1="₹2,470 (40% book)",
    t2="₹2,430 (40% book; trail rest to ₹2,395)",
    sl="₹2,545 + 15-min ST flip to GREEN",
    rr="2.0:1 on stock",
    grade_reason="Filler B-grade — lower conviction than BPCL/INDIGO because Asian Paints has "
                 "pricing-power buffer (phased hikes mooted by foreign broker). Take only if first "
                 "two A-grade setups don't fire and daily trade-count permits.",
    header_color="B76E00",
)

# ============== SECTION 5 — FINAL VERDICT ==============
divider(doc)
add_heading(doc, "SECTION 5 — FINAL VERDICT", level=1)

verdict_rows = [
    ("TODAY'S BIAS",      "BEAR"),
    ("CONFIDENCE",        "HIGH (war + rupee + FII selling + ADR weakness all aligned)"),
    ("CRUDE MODE",        "BEAR (₹9,659/bbl proxy, ₹9,000–10,000 band)"),
    ("VIX SIZING",        "HALF (India VIX ≈ 18)"),
    ("ENTRY PERMISSION",  "YELLOW (squeeze warning + headline risk)"),
]
add_kv_table(doc, verdict_rows, header=["Field", "Value"])

add_para(doc, "")
add_para(doc, "THE BULL CASE (always written, even when bias = BEAR):", bold=True, color=GREEN)
add_para(doc, "If a US–Iran de-escalation / Hormuz partial-reopen headline lands intraday, Brent "
              "drops $10+ in hours, OMCs/aviation squeeze 5–8% up, INR rallies 50 paise, and bears "
              "into weekly expiry get torched.")

add_para(doc, "THE BEAR CASE:", bold=True, color=RED)
add_para(doc, "Hormuz remains shut, Brent grinds toward $112+, INR breaks 96, FII selling continues, "
              "Nifty tests 23,200 then 23,000 floor, IT/discretionary/OMC/aviation all break supports.")

add_para(doc, "FLIP TRIGGER:", bold=True, color=AMBER)
add_para(doc, "If Brent prints below $100 OR any official Iran/US ceasefire wire crosses, "
              "bias flips to CAUTIOUS-BULL immediately — close all shorts, watch ONGC for trailing-stop hits.")

add_para(doc, "")
add_para(doc, "NIFTY KEY LEVELS:", bold=True, color=NAVY)
nifty_levels = [("S2", "23,100"), ("S1", "23,250"), ("CRITICAL", "23,380 (yest close)"),
                ("R1", "23,500"), ("R2", "23,700")]
add_kv_table(doc, nifty_levels, header=["Pivot", "Level"])

add_para(doc, "")
add_para(doc, "TOP 3 RANKED:", bold=True, color=NAVY)
add_bullets(doc, [
    "#1 ONGC — A+ — Crude beneficiary + royalty cut + breakout — entry ABOVE ₹284 (CALL)",
    "#2 BPCL — A — OMC crude victim, cleanest short — entry BELOW ₹284 (PUT)",
    "#3 INDIGO — A — Aviation fuel pain + travel-curb demand hit — entry BELOW ₹4,240 (PUT)",
])

add_para(doc, "")
add_para(doc, "ONE RISK THAT RUINS EVERYTHING TODAY:", bold=True, color=RED)
add_para(doc, "A single Iran–US ceasefire / Hormuz reopening wire during trading hours. It would "
              "simultaneously kill the ONGC long thesis (royalty win priced in but crude tanks) AND "
              "the OMC/aviation shorts (squeeze). Net: every position bleeds. Mitigation: half size, "
              "tight stops, exit all by 13:00.", color=RED)

# ============== ONE-LINE SUMMARY ==============
divider(doc)
add_heading(doc, "ONE-LINE SUMMARY (read at 9:10 AM)", level=2)
one_liner = (
    "Today is BEAR because Hormuz-shut crude + rupee 95.63 + FII heavy selling all align. "
    "Crude ₹9,659 = BEAR mode, falling 0.67% (mild — not enough to flip). Watch ONGC CALL above ₹284 "
    "and BPCL PUT below ₹284. Key risk: Iran ceasefire headline detonates both sides. Size HALF "
    "(VIX 18). Flips if Brent prints <$100 or any official US–Iran de-escalation wire crosses."
)
p = doc.add_paragraph()
r = p.add_run(one_liner)
r.italic = True
r.font.size = Pt(11)
r.font.color.rgb = NAVY

# Footer
divider(doc)
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run("Generated by Elite Daily F&O Intelligence Routine — for personal trading use only. "
                  "Not investment advice. Verify F&O ban list and live option chain at NSE pre-open.")
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = GREY

# Save to both paths (the requested Trading_Briefs/ flat dir + the repo's structured tree)
import os
os.makedirs(os.path.dirname(OUT_PATH_REPO), exist_ok=True)
os.makedirs(os.path.dirname(OUT_PATH_FLAT), exist_ok=True)
doc.save(OUT_PATH_REPO)
doc.save(OUT_PATH_FLAT)
print(f"Saved: {OUT_PATH_REPO}")
print(f"Saved: {OUT_PATH_FLAT}")
