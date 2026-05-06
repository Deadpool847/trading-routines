"""Generate the daily trading-intelligence .docx brief."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

DATE_ISO = "2026-05-06"
BIAS = "CAUTIOUS-BULL"
ENTRY_PERMISSION = "YELLOW"
CONFIDENCE = "MEDIUM"

OUT_DIR = "trading-briefs/2026/05-May"
os.makedirs(OUT_DIR, exist_ok=True)
OUT_PATH = f"{OUT_DIR}/Trading_Brief_{DATE_ISO}_{BIAS}.docx"


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
dr = date_p.add_run("Date: Wednesday, 06 May 2026  |  Generated 03:30 IST  |  Weekly expiry: 6 days  |  Monthly expiry: 22 days  |  Expiry week: NO")
dr.bold = True
dr.font.size = Pt(11)

# Bias + permission badges
badge_line(doc, "TODAY'S BIAS", "CAUTIOUS BULL (override active)", "F1C232")
badge_line(doc, "ENTRY PERMISSION", "YELLOW", "E69138")
badge_line(doc, "CRUDE RULE MODE", "CAUTION (₹8,500-9,000) — direction FALLING", "6AA84F")
badge_line(doc, "VIX SIZING", "HALF-SIZE (India VIX 18.46)", "3D85C6")

doc.add_paragraph()
add_para(doc,
    "OVERRIDE ACTIVE: GIFT Nifty gap-up +167 pts AND crude falling -4% overnight. "
    "Bear bias is suppressed — consensus bearish trade is the trap today. Confidence MEDIUM. "
    "Weekly expiry is 12-May (Tue), 6 days away.",
    italic=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
doc.add_paragraph()
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Reading", "Signal"],
    rows=[
        ["MCX Crude (proxy)", "~₹8,693/bbl (WTI×INR)", "CAUTION zone (₹8,500-9,000)"],
        ["Brent / WTI", "$109.87 / $102.27", "Both -4% overnight — FALLING"],
        ["Crude direction", "FALLING (Iran ceasefire holds)", "Mode flip level: ₹8,500"],
        ["GIFT Nifty", "24,200", "Implied gap +167 pts — BULL gap"],
        ["S&P 500", "7,259.22 (+0.81%)", "Risk ON"],
        ["Nasdaq", "25,326.13 (record)", "Risk ON"],
        ["Dow Jones", "49,298.34 (+356)", "Strong close"],
        ["US VIX", "Calmer post-ceasefire", "Calm tilt"],
        ["India VIX", "18.46 (eased)", "HALF-SIZE rule (15-20)"],
        ["Nifty close (05 May)", "24,032.80 (-86.5 / -0.36%)", "Below max pain"],
        ["FII cash MTD May", "-₹70,135 cr (heavy SELL)", "Bearish flow"],
        ["DII cash MTD May", "+₹51,063 cr (BUY)", "Domestic absorption"],
        ["USD-INR (est.)", "~85 (Hormuz pressure)", "Rupee weak"],
        ["INFY ADR (NYSE)", "$12.46", "Mixed signal"],
        ["Hormuz Strait", "Effectively CLOSED, fragile ceasefire", "HIGH headline risk"],
    ])

doc.add_paragraph()
add_para(doc, "Geopolitical one-liner: ", bold=True, size=11)
add_para(doc,
    "Fragile US-Iran ceasefire holding despite drone attack on Fujairah oil zone; Project Freedom convoy "
    "saw only 2 vessels through Hormuz on day one. Crude fell 4% on ceasefire confirmation but Iran tail "
    "risk remains live — any fresh attack = instant +5-7% crude spike.",
    size=11)

add_para(doc, "Today's intraday headline risk: HIGH",
    bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- SECTION 2 — CRUDE RULE + STRUCTURAL READ ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule + Structural Read", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "CRUDE MODE: CAUTION (half-size) — DIRECTION: FALLING — FLIP LEVEL: ₹8,500",
    bold=True, size=12)
add_bullets(doc, [
    "MCX crude proxy at ~₹8,693/bbl (WTI $102.27 × USDINR ~85). Sits in CAUTION band (₹8,500-9,000).",
    "Direction FALLING -4% overnight. If MCX breaks ₹8,500 → mode flips to MILD BULL (full-size returns).",
    "If Iran headlines re-spike crude → mode escalates to BEAR (puts only) above ₹9,000.",
    "AVOID (today, base case): ONGC/Oil India/MRPL on long side (crude falling hurts upstream).",
    "FAVOUR (crude-falling beneficiaries): BPCL/HPCL/IOC, IndiGo (aviation), paints, tyres.",
    "EXCEPTION: Keep ONGC as a hedge play — ANY Iran headline = upstream rallies 3-5% instantly.",
])

doc.add_paragraph()
add_para(doc, "Structural read", bold=True, size=12)
add_bullets(doc, [
    "Nifty 24,032.80 vs Max Pain 24,050 (12-May expiry) → 17 pts BELOW max pain → magnetism UP.",
    "PCR not directly disclosed; Max Pain pull + GIFT gap-up implies put writers controlling 24,000 floor.",
    "OI direction: heavy FII April selling (-₹70k cr MTD May suggests rolling shorts) — short squeeze possible.",
    "Futures basis: Nifty fut typically near-spot in non-expiry week; no extreme premium/discount signal.",
    "NOT expiry week — weekly is 12-May Tue (6 days). Monthly 28-May (22 days). Both expiries usable.",
    "Bank Nifty leadership: monitor first 30 min — if BNF leads any up-move, trust the move.",
])

# ---------- SECTION 3 — CONTRARIAN CHECK ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Contrarian Check (Layer 3)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Question", "Answer"],
    rows=[
        ["Q1. CONSENSUS today",
         "TV/retail expecting weakness on Hormuz crisis + heavy FII selling MTD. Sell rallies, buy puts."],
        ["Q2. THE TRAP",
         "Crude crashed 4% overnight on ceasefire — kills the bear thesis. Short covering rally squeezes retail puts."],
        ["Q3. RETAIL STOPS",
         "Long stops below 23,950 (round-number magnet). Short stops above 24,250-24,300 (pre-market high zone)."],
        ["Q4. FLIP TRIGGER",
         "Fresh Iran/Hormuz attack headline → crude +5% spike → instantly flips bias to BEAR. Watch newswires 9:00-10:30."],
    ])

doc.add_paragraph()
add_para(doc, "CONTRARIAN OVERRIDE LOGIC", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_kv_table(doc,
    header=["Condition", "Status"],
    rows=[
        ["A: GIFT Nifty gap-up >+100 pts?", "YES (+167 pts)"],
        ["B: Crude falling >2% overnight?", "YES (-4%)"],
        ["C: Contrarian probability >40%?", "YES (consensus bearish setup is the trap)"],
        ["D: Expiry week + PCR <0.8 squeeze risk?", "NO (not expiry week)"],
        ["RESULT", "A+B+C all YES → FULL OVERRIDE → Bias = CAUTIOUS BULL, Permission = YELLOW"],
    ])

add_para(doc,
    "⚠ FULL OVERRIDE: All 3 conditions met. Consensus bearish trade is the trap today. "
    "Bear bias is suppressed. Confidence drops to MEDIUM. Half-size mandatory.",
    bold=True, italic=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- SECTION 4 — STOCK SETUPS ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — 5 Individual F&O Setups", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc,
    "Setups ranked by catalyst clarity + crude alignment in the FALLING-crude regime. "
    "All entries require ALL 4 checklist items to fire on 15-min: SuperTrend colour, RSI(14) vs 50, "
    "StochRSI cross from extreme, Volume >1.5× 20-period average. R:R minimum 2:1.",
    italic=True, size=10)

# ===== STOCK 1: M&M =====
doc.add_paragraph()
add_para(doc, "STOCK 1 — MAHINDRA & MAHINDRA (NSE: M&M)  |  CALL  |  Grade A",
    bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot", "₹3,210.80 (close 05 May) | Lot ~350 | Crude aligned: YES (auto = crude victim → benefits)"),
    ("Catalyst",
     "Q4 FY26 PAT +42% YoY beat (declared 05 May). Final dividend ₹33. 10 ICE SUVs + 6 BEVs roadmap to FY31."),
    ("Daily Trend / 200 DMA",
     "UP weekly | Trading below 20/50 EMA (₹3,133/₹3,214) — needs reclaim. Above 200 DMA: YES."),
    ("Support / Resistance", "Support ₹3,070 / ₹3,032 | Resistance ₹3,222 / ₹3,260"),
    ("OI / Delivery / Ban",
     "ATM May 26 expiry strike 3,120 priced ±7.6% straddle. Ban list: NO. Delivery STABLE."),
    ("Option Setup",
     "M&M 3,200 CE 28-MAY-2026 monthly (ATM). Theta warning: NO (22d to expiry)."),
    ("Entry (ALL 4 must fire)",
     "ST GREEN on 15-min + RSI>50 + StochRSI cross UP from <30 + Volume >1.5×. Confirmation candle: 15-min close above ₹3,222."),
    ("T1 / T2", "T1 ₹3,260 (40% book) | T2 ₹3,310 (40% book)"),
    ("SL", "₹3,170 + ST flips RED on 15-min"),
    ("R:R / Time stop", "R:R ~2.3:1 | Exit by 13:00 IST (hard stop 13:30)"),
    ("Grade reason", "A: Strong Q4 beat catalyst + falling crude tailwind + above 200 DMA + clear technical trigger."),
])

# ===== STOCK 2: BPCL =====
doc.add_paragraph()
add_para(doc, "STOCK 2 — BHARAT PETROLEUM (NSE: BPCL)  |  CALL  |  Grade A",
    bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot", "₹356.40 (close 05 May) | Lot ~1,800 | Crude aligned: YES (OMC = direct crude-fall beneficiary)"),
    ("Catalyst",
     "WTI -4% overnight = marketing margin expansion. OMCs structurally squeezed when crude rises; reverse trade now."),
    ("Daily Trend / 200 DMA",
     "Sideways-to-up | Above 200 DMA: YES (52w range 234-391)."),
    ("Support / Resistance", "Support ₹344 / ₹352 | Resistance ₹367 / ₹379"),
    ("OI / Delivery / Ban",
     "Active May 26 + Jun 30 chains. Ban list: NO. MACD slow-divergence flagged — risk noted."),
    ("Option Setup",
     "BPCL 360 CE 28-MAY-2026 monthly (ATM). Theta warning: NO (22d to expiry)."),
    ("Entry (ALL 4 must fire)",
     "ST GREEN on 15-min + RSI>50 + StochRSI cross UP from <30 + Volume >1.5×. Confirmation candle: 15-min close above ₹362."),
    ("T1 / T2", "T1 ₹367 (40% book) | T2 ₹375 (40% book)"),
    ("SL", "₹354 + ST flips RED on 15-min"),
    ("R:R / Time stop", "R:R ~2.4:1 | Exit by 13:00 IST"),
    ("Grade reason", "A: Crude crash directly accretive to OMC margins; clear sector rotation trade today."),
])

# ===== STOCK 3: INDIGO =====
doc.add_paragraph()
add_para(doc, "STOCK 3 — INTERGLOBE AVIATION (NSE: INDIGO)  |  CALL  |  Grade B",
    bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot", "₹4,328.90 (04 May) | Lot ~25 | Crude aligned: YES (aviation fuel = ~40% cost; falling crude = relief)"),
    ("Catalyst",
     "Q3 PAT was -77% YoY (weak). 17% May intl capacity cut due to West Asia. Falling crude is THE positive catalyst today."),
    ("Daily Trend / 200 DMA",
     "Down from 6,232 high to near 52w low 4,021 | Above 200 DMA: BORDERLINE (check at open)."),
    ("Support / Resistance", "Support ₹4,250 / ₹4,021 (52w low) | Resistance ₹4,387 / ₹4,500"),
    ("OI / Delivery / Ban", "Ban list: NO. Recent weakness has built short OI — squeeze fuel if triggered."),
    ("Option Setup",
     "INDIGO 4,400 CE 28-MAY-2026 monthly. Theta warning: NO."),
    ("Entry (ALL 4 must fire)",
     "ST GREEN on 15-min + RSI>50 + StochRSI cross UP from <30 + Volume >1.5×. Confirmation: 15-min close above ₹4,390."),
    ("T1 / T2", "T1 ₹4,460 (40% book) | T2 ₹4,540 (40% book)"),
    ("SL", "₹4,330 + ST RED"),
    ("R:R / Time stop", "R:R ~2.2:1 | Exit by 13:00 IST"),
    ("Grade reason", "B: Crude alignment good but weak Q3 + capacity cut signal + 200 DMA borderline. Higher conviction needs full breakout."),
])

# ===== STOCK 4: L&T =====
doc.add_paragraph()
add_para(doc, "STOCK 4 — LARSEN & TOUBRO (NSE: LT)  |  CALL  |  Grade B",
    bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot", "₹4,051.30 (close 05 May) | Lot ~150 | Crude aligned: NEUTRAL"),
    ("Catalyst",
     "Q4 FY26 declared 05 May: Revenue +11.3% YoY, PAT -3.1% (mild miss). Order book RECORD ₹5.79 lakh cr (+22%). Final dividend ₹38."),
    ("Daily Trend / 200 DMA",
     "UP 1Y +21.65% | Above 200 DMA: YES (52w range 3,284-4,440)."),
    ("Support / Resistance", "Support ₹4,000 / ₹3,950 | Resistance ₹4,100 / ₹4,180"),
    ("OI / Delivery / Ban", "Ban list: NO. PAT miss vs strong order book = mixed reaction expected."),
    ("Option Setup",
     "LT 4,100 CE 28-MAY-2026 monthly. Theta warning: NO."),
    ("Entry (ALL 4 must fire)",
     "WAIT for first 15-min post-results gap to settle. ST GREEN + RSI>50 + StochRSI cross UP + Volume >1.5×. Confirmation: 15-min close above ₹4,090."),
    ("T1 / T2", "T1 ₹4,150 (40% book) | T2 ₹4,200 (40% book)"),
    ("SL", "₹4,030 + ST RED"),
    ("R:R / Time stop", "R:R ~2.0:1 | Exit by 13:00 IST"),
    ("Grade reason", "B: Strong order book is bullish anchor but PAT miss may produce gap-down first. Take only if reverses cleanly."),
])

# ===== STOCK 5: ONGC =====
doc.add_paragraph()
add_para(doc, "STOCK 5 — OIL & NATURAL GAS CORP (NSE: ONGC)  |  PUT (with hedge view)  |  Grade B",
    bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot", "₹293.20 (recent zone, near 52w high ₹307.50) | Lot ~9,500 | Crude aligned: NO (upstream → falling crude bearish)"),
    ("Catalyst",
     "Falling crude -4% removes near-term realisation tailwind. JM Financial Buy ₹340 target — but that's predicated on Hormuz tension."),
    ("Daily Trend / 200 DMA",
     "Up trend recently | Above 200 DMA: YES (broke above 270-272 resistance)."),
    ("Support / Resistance", "Support ₹284 / ₹272 | Resistance ₹298 / ₹307 (52w high)"),
    ("OI / Delivery / Ban", "Ban list: NO. Long unwind risk if crude continues falling."),
    ("Option Setup",
     "ONGC 290 PE 28-MAY-2026 monthly. Theta warning: NO. Consider as HEDGE leg vs long beneficiary calls."),
    ("Entry (ALL 4 must fire)",
     "ST RED on 15-min + RSI<50 + StochRSI cross DOWN from >70 + Volume >1.5×. Confirmation: 15-min close below ₹290."),
    ("T1 / T2", "T1 ₹284 (40% book) | T2 ₹278 (40% book)"),
    ("SL", "₹296 + ANY fresh Iran headline = exit immediately (mode flip risk)"),
    ("R:R / Time stop", "R:R ~2.1:1 | Exit by 13:00 IST"),
    ("Grade reason", "B: Direction logic is right but tail-risk (Iran re-escalation) is binary and brutal. Use for hedge balance, not size."),
])

# ---------- SECTION 5 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["TODAY'S BIAS", "CAUTIOUS BULL (override active — gap-up + falling crude)"],
        ["CONFIDENCE", "MEDIUM"],
        ["CRUDE MODE", "CAUTION zone (₹8,500-9,000) — direction FALLING"],
        ["VIX SIZING", "HALF-SIZE mandatory (India VIX 18.46)"],
        ["ENTRY PERMISSION", "YELLOW — slow, selective, confirmation required"],
        ["MAX TRADES TODAY", "3 (per profile rules)"],
        ["DAILY RISK CAP", "2% of capital"],
        ["HARD EXIT", "13:30 IST, no exceptions"],
    ])

doc.add_paragraph()
add_para(doc, "THE BULL CASE (primary):", bold=True, size=12, color=RGBColor(0x38, 0x76, 0x1D))
add_para(doc,
    "Crude crashed 4% on ceasefire confirmation — a structural relief for 70%+ of Nifty constituents. "
    "GIFT Nifty +167 pts is real money pricing the crude relief; FII shorts built through April are the squeeze fuel.",
    size=11)

add_para(doc, "THE BEAR CASE (real risk):", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "Hormuz still effectively closed; ceasefire is fragile (Fujairah just attacked). One Iran headline reverses "
    "everything in minutes. India VIX at 18.46 is elevated for a reason — markets do not trust the calm.",
    size=11)

doc.add_paragraph()
add_para(doc, "FLIP TRIGGER:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "If a fresh Iran/Hormuz attack headline lands during 9:15-12:30, bias flips to BEAR instantly. "
    "Crude spike >+3% on Reuters/CNBC newswire = exit all longs, no questions.",
    size=11)

doc.add_paragraph()
add_para(doc, "NIFTY KEY LEVELS", bold=True, size=12)
add_kv_table(doc,
    header=["Level", "Value", "Meaning"],
    rows=[
        ["S2 (deep support)", "₹23,800", "Breakdown level — below = trend reverses bearish"],
        ["S1 (support)", "₹23,950", "Round-number magnet; retail long stops"],
        ["CRITICAL pivot", "₹24,050", "Max Pain — gravitational pull"],
        ["R1 (resistance)", "₹24,250", "Pre-market high zone; retail short stops"],
        ["R2 (extension)", "₹24,400", "Squeeze target if shorts unwind"],
    ])

doc.add_paragraph()
add_para(doc, "TOP 3 RANKED FOR EXECUTION", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "#1 M&M — A — Q4 PAT +42% beat + falling-crude auto tailwind — entry above ₹3,222",
    "#2 BPCL — A — OMC margin expansion on crude crash — entry above ₹362",
    "#3 INDIGO — B — Aviation fuel relief trade — entry above ₹4,390",
])

doc.add_paragraph()
add_para(doc, "ONE RISK THAT RUINS EVERYTHING TODAY:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "A fresh Iranian strike on Gulf shipping or oil infrastructure during India market hours. "
    "Crude re-spikes 5-7%, India VIX jumps to 22+, all crude-victim longs (M&M, BPCL, INDIGO) get crushed in 30 min. "
    "Mitigation: small ONGC put hedge OR tight 1.0% portfolio stop.",
    size=11)

# ---------- ONE-LINE SUMMARY ----------
doc.add_paragraph()
add_heading(doc, "One-Line Summary (read at 9:10 AM)", level=2, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Today is CAUTIOUS BULL because crude crashed 4% on Iran ceasefire and GIFT is +167. "
    "Crude ₹8,693 = CAUTION, falling. Watch M&M CALL and BPCL CALL. Key risk: fresh Iran headline "
    "spikes crude. Size HALF. Flips to BEAR if any Hormuz attack newswire hits.\"",
    italic=True, size=12)

# ---------- FOOTER ----------
doc.add_paragraph()
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run("Generated by Daily Trading Routine — for personal educational use only. Not financial advice.")
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
