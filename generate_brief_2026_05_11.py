"""Generate the daily trading-intelligence .docx brief for 2026-05-11."""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-11"
BIAS = "BEAR"
CONFIDENCE = "MEDIUM"
ENTRY_PERMISSION = "YELLOW"
CRUDE_MODE = "MILD BULL (rising; nearing CAUTION flip)"
VIX_SIZING = "HALF SIZE"
GEN_TIME = "08:45 IST"

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
sub_run = sub.add_run("Indian F&O Day Trader - Nifty + Stocks")
sub_run.italic = True
sub_run.font.size = Pt(12)

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
dr = date_p.add_run(f"Date: Monday, 11 May 2026  |  Generated {GEN_TIME}  |  Weekly Expiry: 3 days  |  Monthly Expiry: 17 days  |  Expiry Week: YES")
dr.bold = True
dr.font.size = Pt(10)

# Bias + permission badges
badge_line(doc, "TODAY'S BIAS", f"{BIAS} ({CONFIDENCE} confidence)", "C00000")
badge_line(doc, "ENTRY PERMISSION", ENTRY_PERMISSION, "E69138")
badge_line(doc, "CRUDE MODE", CRUDE_MODE, "F1C232")
badge_line(doc, "VIX SIZING", VIX_SIZING, "1F3864")

doc.add_paragraph()
add_para(doc,
    "EXPIRY ALERT: Weekly Nifty expiry 3 days away (Thu 14 May). Options Rule: Days-to-expiry = 3 is BORDERLINE; "
    "all new option positions today must use MONTHLY expiry (28 May) to avoid theta cliff.",
    bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
doc.add_paragraph()
add_heading(doc, "Section 1 - Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Value", "Signal"],
    rows=[
        ["MCX Crude (proxy)", "~Rs 8,015/bbl (WTI 95.42 x USD-INR 84)", "MILD BULL, but rising toward CAUTION"],
        ["Brent / WTI", "$104.49 (+3.16%) / $95.42 (+3.08%)", "RISING sharply - Iran ceasefire fragile"],
        ["GIFT Nifty", "24,064 (-0.81% from 24,259.5)", "GAP DOWN ~195 pts - bearish open"],
        ["S&P 500 (Fri close)", "7,398.93 (+0.8%)", "Risk-ON in US, but dated 2-day-old"],
        ["Nasdaq (Fri close)", "26,247.08 (+1.7%)", "Tech tailwind for IT"],
        ["Dow Jones (Fri close)", "49,609.16 (+0.02%)", "Flat"],
        ["US VIX", "~mid-teens (calm)", "Risk-on US calm"],
        ["India VIX", "~16.5 (mid-teens)", "Elevated 15-20 zone -> HALF SIZE"],
        ["Nifty last close (Fri 8 May)", "Rs 24,176 (-0.62%)", "Below 24,200; weak close"],
        ["FII cash (latest)", "-Rs 340.89 Cr (SELL)", "5-day trend: SELLING; YTD outflow > Rs 2 lakh cr"],
        ["DII cash (latest)", "+Rs 441.07 Cr (BUY)", "Absorbing FII supply"],
        ["USD-INR", "~Rs 84 (estimate)", "Stable; rupee weakness mild"],
        ["INFY ADR", "$12.57 (~flat)", "INFY opens flat"],
        ["WIT ADR", "$2.00 (+0.50%)", "WIPRO mild positive"],
        ["IBN ADR", "$26.75 (+2.81%)", "ICICI opens GAP UP"],
        ["HDB ADR", "$25.63 (+3.22%)", "HDFC Bank opens GAP UP"],
    ])

doc.add_paragraph()
add_para(doc, "Geopolitical one-liner:", bold=True, size=11)
add_para(doc,
    "Iran-Hormuz: Strait largely blocked since 28 Feb 2026. Trump paused Project Freedom (US Navy escort op) on "
    "6 May citing 'great progress' on 14-point MOU. BUT crude up 3% overnight - market doubting durability of "
    "ceasefire. Status: PARTIAL. Headline risk: HIGH. Direction of risk: crude spike on any deal collapse.",
    size=11)

# ---------- SECTION 2 — CRUDE RULE + STRUCTURAL READ ----------
doc.add_paragraph()
add_heading(doc, "Section 2 - Crude Rule + Structural Read", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "CRUDE MODE: MILD BULL - DIRECTION: RISING (+3% overnight)", bold=True, size=12)
add_bullets(doc, [
    "MCX crude proxy ~Rs 8,015/bbl (using WTI $95.42 x USD-INR 84). Falls in Rs 7,500-8,500 MILD BULL band.",
    "CRUDE_FLIP_LEVEL: Rs 8,500 - a break above flips to CAUTION mode (half-size for everything).",
    "Crude is RISING - mode is STRENGTHENING toward bear. Flag: bear mode deepening.",
    "AVOID sectors: OMCs (BPCL, HPCL, IOC), aviation (IndiGo), paints, tyres, fertilisers.",
    "EXCEPTION_CALLS permitted: ONGC, Oil India, MRPL (upstream crude beneficiaries).",
    "Tail risk: any single Iran/Hormuz negative headline can re-spike crude 5-7%+ intraday.",
])

doc.add_paragraph()
add_para(doc, "Structural Read", bold=True, size=12)
add_bullets(doc, [
    "Nifty vs Max Pain: DATA_UNAVAILABLE for exact max pain; estimated zone 24,200-24,300 with downward pull.",
    "PCR: DATA_UNAVAILABLE for exact value; expiry-week dynamics suggest call writers active at 24,300/24,400.",
    "OI change direction (interpretation): With gap-down setup, expect call additions at 24,200-24,300 (bearish), put unwinding 24,000 (mildly bearish until floor tested).",
    "Futures basis: DATA_UNAVAILABLE; expiry week typically compresses to par/slight discount.",
    "Expiry week dynamics: 3 days to weekly expiry. With Nifty below 24,200 and gap-down, max pain magnetism likely toward 24,100-24,200 zone.",
    "Put writers floor: ~Rs 24,000 (psychological + last week's swing low cluster).",
    "Call writers ceiling: ~Rs 24,300-24,400 (gap zone).",
])

doc.add_paragraph()
add_para(doc, "Bank Nifty Leadership Check", bold=True, size=12)
add_bullets(doc, [
    "ADR signals show ICICI +2.81% and HDFC Bank +3.22% (Friday US close) - financial-sector strength baked in.",
    "If Bank Nifty opens UP while Nifty opens DOWN, this is a DIVERGENCE - Bank Nifty leading would suggest the gap-down in Nifty is fadable on dips.",
    "If Bank Nifty also opens DOWN (despite ADR strength), this confirms broad risk-off and the move is REAL.",
    "Watch first 30 minutes: Bank Nifty divergence is the single most important intraday tell.",
])

# ---------- SECTION 3 — CONTRARIAN CHECK ----------
doc.add_paragraph()
add_heading(doc, "Section 3 - Contrarian Check (Layer 3)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Q1 - CONSENSUS: What is every retail trader/TV expecting today?", bold=True, size=11)
add_para(doc,
    "GAP-DOWN at open, Nifty to test 24,000, BUY OMC puts, SELL aviation, hide in defensives, expect crude to keep spiking on Iran.",
    size=11)

add_para(doc, "Q2 - THE TRAP: How does the market punish that trade?", bold=True, size=11)
add_para(doc,
    "(a) Iran-US 14-point deal lands by mid-morning -> crude collapses 4-6% -> OMC puts get destroyed, aviation rips, ONGC dumps. "
    "(b) Gap-down at open gets bought aggressively by DIIs (Friday pattern: +Rs 441 cr buying); shorts get squeezed into expiry. "
    "(c) HDFC Bank/ICICI gap-up via ADR drags Bank Nifty up, Nifty reclaims 24,200, retail Nifty-put buyers get melted.",
    size=11)

add_para(doc, "Q3 - RETAIL STOPS: Where are they clustered?", bold=True, size=11)
add_para(doc,
    "Long stops below 24,000 (psychological). Short stops above 24,300 (Friday opening zone) and 24,400 (Thursday high cluster).",
    size=11)

add_para(doc, "Q4 - FLIP TRIGGER: The single event that reverses direction completely.", bold=True, size=11)
add_para(doc,
    "An Iran-US deal announcement (any time during India session) - confirmed by Brent dropping below $100 - flips bias from BEAR to "
    "CAUTIOUS BULL within an hour. Conversely, any reported Hormuz attack on a tanker flips bias to STRONG BEAR with crude limit-up risk.",
    size=11)

doc.add_paragraph()
add_para(doc, "Contrarian Override Check", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_kv_table(doc,
    header=["Condition", "Status", "Effect"],
    rows=[
        ["A. GIFT Nifty gap-up > +100 pts?", "NO (gap-DOWN -195)", "No override"],
        ["B. Crude falling > 2% overnight?", "NO (RISING +3%)", "No override"],
        ["C. Contrarian probability > 40%?", "NO (~25-30%)", "No override"],
        ["D. Expiry week + PCR < 0.8?", "PCR UNVERIFIED (expiry-week YES)", "Squeeze risk monitored, not triggered"],
    ])
add_para(doc,
    "Result: NO override triggered. Bear bias stands - but with mandatory yellow permission due to (i) half-size VIX rule, "
    "(ii) ADR-driven financial gap-up creating intraday divergence risk, (iii) Iran deal headline tail risk.",
    italic=True, size=11)

# ---------- SECTION 4 — 5 INDIVIDUAL F&O SETUPS ----------
doc.add_paragraph()
add_heading(doc, "Section 4 - Five Individual F&O Setups", level=1, color=RGBColor(0x1F, 0x38, 0x64))

# --- STOCK 1: ONGC ---
add_para(doc, "1. ONGC (NSE: ONGC) | CALL | Grade A", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot", "~Rs 293 (near 52w high Rs 307.50) | Lot ~3,850 | Crude aligned: YES"),
    ("Catalyst", "Crude up +3% overnight on Iran fragility. ONGC sensitivity: every $1 Brent up ~Rs 6,180 cr annual earnings impact."),
    ("Technical", "Daily UP. Above 200 DMA: YES. Weekly UP. Support Rs 280 (breakout retest) | Resistance Rs 307 (52w high)."),
    ("Structural", "OI: LONG BUILD ongoing. Delivery: RISING (institutional accumulation). Block deal: NO. F&O Ban: NO."),
    ("Option Setup", "ONGC 295 CE - MAY MONTHLY (28 May) expiry. ATM. Theta warning: NO (17 days)."),
    ("Entry (ALL 4)", "ST GREEN on 15-min CONFIRMED | RSI > 50 on 15-min | StochRSI cross UP from <30 OR fresh cross >70 | Volume > 1.5x 20-pd avg. Confirmation: 15-min close above Rs 295."),
    ("T1 / T2", "T1 Rs 300 (book 40%) | T2 Rs 306 (book 40%) | trail remainder."),
    ("SL", "Rs 287 (below breakout retest) OR ST flips RED. R:R approx 2.5:1."),
    ("Time Stop", "Exit by 13:00 IST."),
    ("Grade Reason", "Pure crude beneficiary in rising-crude regime; daily uptrend + above 200 DMA + clean structural setup."),
])

doc.add_paragraph()

# --- STOCK 2: HAL ---
add_para(doc, "2. HINDUSTAN AERONAUTICS (NSE: HAL) | CALL | Grade A", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot", "Rs 4,788 (Fri close) | Lot 150 | Crude aligned: N/A (defence theme)"),
    ("Catalyst", "Iran tensions + Hormuz blockade favours defence-export narrative. Q4 FY26 results on 14 May (3 days) - pre-result rally bias."),
    ("Technical", "Daily UP. Above 200 DMA: YES. Weekly UP. Support Rs 4,650 | Resistance Rs 4,900 (toward 52w high Rs 5,165)."),
    ("Structural", "OI: LONG BUILD (defence theme positioning). Delivery: STABLE-RISING. Block deal: NO. F&O Ban: NO."),
    ("Option Setup", "HAL 4800 CE - MAY MONTHLY (28 May). ATM/slight OTM. Theta warning: NO."),
    ("Entry (ALL 4)", "ST GREEN on 15-min | RSI > 50 | StochRSI cross UP from extreme | Volume > 1.5x 20-pd avg. Confirmation: 15-min close above Rs 4,810."),
    ("T1 / T2", "T1 Rs 4,880 (book 40%) | T2 Rs 4,950 (book 40%) | trail remainder."),
    ("SL", "Rs 4,720 (below Fri low) OR ST flips RED. R:R approx 2.2:1."),
    ("Time Stop", "Exit by 13:00 IST. Avoid holding into 14 May results."),
    ("Grade Reason", "Defence catalyst + Iran tail risk supportive + pre-result run-up + structurally above 200 DMA."),
])

doc.add_paragraph()

# --- STOCK 3: ICICI Bank ---
add_para(doc, "3. ICICI BANK (NSE: ICICIBANK) | CALL | Grade A", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot", "Rs 1,383.50 (up +2.01% recent) | Lot 700 | Crude aligned: NEUTRAL"),
    ("Catalyst", "IBN ADR +2.81% on Fri US close = strong gap-up signal. Q4 net profit +9.28% YoY (Rs 14,755 cr) already known."),
    ("Technical", "Daily UP. Above 200 DMA (Rs 1,371.76): YES (just breached). Above 50 DMA Rs 1,303. Support Rs 1,372 (200 DMA) | Resistance Rs 1,420 then Rs 1,500 (52w high)."),
    ("Structural", "OI: LONG BUILD (3 sessions up). Delivery: STABLE. Block deal: NO. F&O Ban: NO."),
    ("Option Setup", "ICICIBANK 1400 CE - MAY MONTHLY (28 May). Slight OTM. Theta warning: NO."),
    ("Entry (ALL 4)", "ST GREEN on 15-min | RSI > 50 | StochRSI cross UP | Volume > 1.5x avg. Confirmation: 15-min close above Rs 1,390 + 200 DMA holds as support on first dip."),
    ("T1 / T2", "T1 Rs 1,410 (book 40%) | T2 Rs 1,430 (book 40%) | trail remainder."),
    ("SL", "Rs 1,370 (below 200 DMA) OR ST flips RED. R:R approx 2.1:1."),
    ("Time Stop", "Exit by 13:00 IST."),
    ("Grade Reason", "ADR-confirmed gap up + just breached 200 DMA (institutional re-entry signal) + uptrend continuation + DII support."),
])

doc.add_paragraph()

# --- STOCK 4: BPCL ---
add_para(doc, "4. BHARAT PETROLEUM (NSE: BPCL) | PUT | Grade B+", bold=True, size=13, color=RGBColor(0xC0, 0x00, 0x00))
stock_card(doc, [
    ("CMP / Lot", "Rs 304-306 (Fri range) | Lot 1,800 | Crude aligned: NEGATIVE (OMC squeeze)"),
    ("Catalyst", "Crude +3% overnight squeezes OMC marketing margins. Brent at $104 implies under-recovery if retail prices held."),
    ("Technical", "Daily DOWN-SIDEWAYS. Below 50 DMA. 200 DMA: NO (below). Support Rs 302 then Rs 281 | Resistance Rs 344."),
    ("Structural", "OI: SHORT BUILD on crude spike days. Delivery: STABLE. Block deal: NO. F&O Ban: NO."),
    ("Option Setup", "BPCL 300 PE - MAY MONTHLY (28 May). ATM/slight OTM. Theta warning: NO."),
    ("Entry (ALL 4)", "ST RED on 15-min | RSI < 50 | StochRSI cross DOWN from >70 | Volume > 1.5x avg. Confirmation: 15-min close below Rs 302."),
    ("T1 / T2", "T1 Rs 297 (book 40%) | T2 Rs 290 (book 40%) | trail remainder."),
    ("SL", "Rs 310 (above Fri high) OR ST flips GREEN OR Brent breaks below $100. R:R approx 2.0:1."),
    ("Time Stop", "Exit by 13:00 IST. KILL trade if Iran-US deal headline lands."),
    ("Grade Reason", "Direct crude victim in rising regime; B+ (not A) because range-bound recent action and Iran-deal headline is binary kill-switch."),
])

doc.add_paragraph()

# --- STOCK 5: HDFC Bank ---
add_para(doc, "5. HDFC BANK (NSE: HDFCBANK) | CALL | Grade B", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
stock_card(doc, [
    ("CMP / Lot", "Rs 780.85 (Fri close) | Lot ~550 | Crude aligned: NEUTRAL"),
    ("Catalyst", "HDB ADR +3.22% on Fri = strongest gap-up signal of the day. Q4 FY26 PAT +8.05% YoY (Rs 20,350 cr) already known."),
    ("Technical", "Daily DOWNTREND (stock -23.48% from 52w high Rs 1,020.5). 200 DMA: NO (below). Support Rs 770 | Resistance Rs 800 then Rs 820."),
    ("Structural", "OI: SHORT COVERING expected on ADR-driven gap up. Delivery: STABLE. Block deal: NO. F&O Ban: NO."),
    ("Option Setup", "HDFCBANK 800 CE - MAY MONTHLY (28 May). Slight OTM. Theta warning: NO."),
    ("Entry (ALL 4)", "ST GREEN on 15-min CONFIRMED (not on opening minute) | RSI > 50 | StochRSI cross UP | Volume > 1.5x avg. Confirmation: 15-min close above Rs 790 after first 30 mins."),
    ("T1 / T2", "T1 Rs 802 (book 40%) | T2 Rs 815 (book 40%) | trail remainder."),
    ("SL", "Rs 775 (below Fri close) OR ST flips RED OR opening gap fails to hold. R:R approx 2.0:1."),
    ("Time Stop", "Exit by 13:00 IST."),
    ("Grade Reason", "B (not A) due to conflict: strong ADR signal vs daily downtrend. Pure catch-up trade; must see follow-through in first 30 mins or skip."),
])

# ---------- SECTION 5 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 5 - Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["TODAY'S BIAS", "BEAR"],
        ["CONFIDENCE", "MEDIUM"],
        ["CRUDE MODE", "MILD BULL (rising; nearing CAUTION flip at Rs 8,500)"],
        ["VIX SIZING", "HALF SIZE (India VIX ~16.5 in 15-20 band)"],
        ["ENTRY PERMISSION", "YELLOW"],
        ["EXPIRY ALERT", "Weekly expiry Thu 14 May (3 days) - use MONTHLY only for new options"],
    ])

doc.add_paragraph()
add_para(doc, "THE BULL CASE (2 lines):", bold=True, size=12, color=RGBColor(0x38, 0x76, 0x1D))
add_para(doc,
    "ADR-confirmed gap-up in IBN (+2.81%) and HDB (+3.22%) suggests financials will lead a Bank-Nifty-led reclaim of 24,200. "
    "Any Iran-US deal headline collapses crude and triggers a full BEAR-to-BULL flip with shorts squeezed into Thursday expiry.",
    size=11)

add_para(doc, "THE BEAR CASE (2 lines):", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "GIFT Nifty -195 pts + crude +3% + persistent FII selling (YTD outflow > Rs 2 lakh cr) + Iran ceasefire breakdown = "
    "Nifty test of 24,000 floor. Expiry-week pin near max-pain zone 24,100-24,200, then drift lower into Thursday.",
    size=11)

doc.add_paragraph()
add_para(doc, "FLIP TRIGGER:", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "If Brent drops below $100 intraday OR a confirmed Iran-US deal headline hits the tape, bias flips to CAUTIOUS BULL "
    "within the hour. Conversely, a Hormuz tanker attack or Brent > $108 cements STRONG BEAR with crude limit-up risk.",
    size=11)

doc.add_paragraph()
add_para(doc, "NIFTY KEY LEVELS", bold=True, size=12)
add_kv_table(doc,
    rows=[
        ["S2", "Rs 23,900 (last week swing low)"],
        ["S1", "Rs 24,000 (psychological floor)"],
        ["CRITICAL (pivot)", "Rs 24,176 (Fri close)"],
        ["R1", "Rs 24,260 (GIFT Nifty current)"],
        ["R2", "Rs 24,400 (Thu high zone)"],
    ])

doc.add_paragraph()
add_para(doc, "TOP 3 RANKED SETUPS", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "#1 ONGC - Grade A - Crude beneficiary; rising crude regime - entry above Rs 295 (CE MAY monthly)",
    "#2 ICICI Bank - Grade A - ADR +2.81% gap up + just breached 200 DMA - entry above Rs 1,390 (CE MAY monthly)",
    "#3 HAL - Grade A - Defence theme + Iran tensions + pre-result run-up - entry above Rs 4,810 (CE MAY monthly)",
])

doc.add_paragraph()
add_para(doc, "ONE RISK THAT RUINS EVERYTHING TODAY:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "A surprise Iran-US deal headline mid-session. It would simultaneously: (a) crash crude 5-8%, destroying ONGC/HAL longs, "
    "(b) rip aviation/OMCs higher, destroying BPCL puts, and (c) explode VIX lower, melting all option premiums. "
    "Single tape line, three trades killed. Set news alerts; be ready to exit ALL positions within 5 minutes if it hits.",
    size=11)

# ---------- ONE-LINE SUMMARY ----------
doc.add_paragraph()
add_para(doc, "ONE-LINE SUMMARY (read at 9:10 AM IST)", bold=True, size=13, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Today is BEAR because crude +3% on Iran fragility plus GIFT Nifty gap-down -195 pts plus persistent FII selling. "
    "Crude ~Rs 8,015 = MILD BULL but RISING toward CAUTION flip at Rs 8,500. Watch ONGC CE and ICICI Bank CE for entries. "
    "Key risk: Iran-US deal headline collapses crude, flips entire setup. Size HALF (VIX 16.5). "
    "Flips to BULL if Brent breaks below $100 intraday.\"",
    italic=True, size=11)

# ---------- FOOTER ----------
doc.add_paragraph()
doc.add_paragraph()
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run("Generated by Daily Trading Routine - personal educational use only. Not financial advice. "
                   "Trade discipline: max 3 trades, hard exit 13:30 IST, 2% daily risk cap. "
                   "ALL 4 entry conditions must fire simultaneously (ST + RSI + StochRSI + Volume).")
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
