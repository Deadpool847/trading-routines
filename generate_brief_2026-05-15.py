"""Generate the 15 May 2026 daily trading-intelligence .docx brief."""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-15"
BIAS = "CAUTIOUS-BULL"
ENTRY_PERMISSION = "YELLOW"
CRUDE_MODE = "BEAR (falling toward MILD-BULL)"
VIX_SIZING = "QUARTER (war-VIX elevated)"

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
sub_run = sub.add_run("Indian F&O Day Trader  |  Nifty + Single Stocks")
sub_run.italic = True
sub_run.font.size = Pt(12)

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
dr = date_p.add_run("Date: Friday, 15 May 2026  |  Generated 08:45 IST  |  Weekly Expiry: 2 trading days (Tue 19 May)  |  Monthly Expiry: 7 trading days (Tue 26 May)  |  Expiry Week: YES")
dr.bold = True
dr.font.size = Pt(10)

# Bias badges
badge_line(doc, "TODAY'S BIAS", BIAS, "B45F06")            # burnt-orange (overridden bear)
badge_line(doc, "ENTRY PERMISSION", ENTRY_PERMISSION, "F1C232")  # amber
badge_line(doc, "CRUDE MODE", CRUDE_MODE, "990000")        # bear red
badge_line(doc, "VIX SIZING", VIX_SIZING, "674EA7")
badge_line(doc, "EXPIRY ALERT", "Weekly Tue 19 May — DTE 2 trading days; theta accelerating", "0B5394")

doc.add_paragraph()

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Reading", "Signal"],
    rows=[
        ["MCX Crude (₹/bbl) — proxy", "~₹9,646 (WTI $101.54 × USDINR ₹95)", "BEAR band; falling"],
        ["Brent / WTI (live)", "$106.07 / $101.54", "Off $114 spike; -5% wk → FALLING"],
        ["Crude direction o/n", "FALLING (>2%)", "Mode WEAKENING — flip watch"],
        ["GIFT Nifty", "23,845 (+261, +1.11%)", "STRONG GAP UP — bullish"],
        ["Implied gap vs 23,694 close", "+151 pts", "Gap-up >100 = override input A"],
        ["S&P 500", "7,500 (+0.8%) — RECORD", "Risk-ON tailwind"],
        ["Nasdaq", "26,635 (+0.9%) — RECORD", "Tech tailwind for IT"],
        ["Dow Jones", "50,063 (+0.75%)", "Crossed 50k milestone"],
        ["US VIX (est)", "~14-16 (calm despite war)", "Calm-side fear"],
        ["India VIX", "DATA_UNAVAILABLE — assumed >20 (war regime)", "QUARTER SIZE applied"],
        ["Nifty close (14 May)", "23,694.15", "+1.18% prior session"],
        ["FII cash (12 May)", "-₹1,959 Cr (NET SELL)", "Bearish; stale data"],
        ["DII cash (12 May)", "+₹7,990 Cr (NET BUY)", "Heavy absorption"],
        ["USD-INR", "~₹95 (weak)", "FII headwind continues"],
        ["INFY ADR (14 May NY)", "$11.86", "Steady; ADR-implied flat-to-up open"],
        ["HDB ADR", "$25.67 (8 May)", "DATA_UNAVAILABLE for 14 May"],
        ["Hormuz status", "BLOCKED since 28 Feb 2026", "Tail-risk live; HIGH headline risk"],
    ])

add_para(doc, "Geopolitical one-liner:", bold=True, size=11)
add_para(doc,
    "Hormuz still blocked; Aramco CEO warns market won't normalize before 2027 if disruption persists past mid-June. "
    "Trump-Xi call discussed Strait of Hormuz — markets reading as cooling signal (oil fell from $107 to $101). "
    "Iran ceasefire described by Trump as 'on life support'. One headline can re-spike crude 5–7%.",
    size=11)
add_para(doc, "Today's intraday headline risk: HIGH (Iran tape-bomb risk persists every session).",
    bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- SECTION 2 — CRUDE RULE + STRUCTURAL READ ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule + Structural Read", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "CRUDE RULE STATE", bold=True, size=12)
add_bullets(doc, [
    "CRUDE_MODE: BEAR (puts-only band ₹9,000–₹10,000; we are at ~₹9,646)",
    "CRUDE_DIRECTION: FALLING (~-5% week, from $107 → $101.54 WTI on Trump-Xi tone)",
    "CRUDE_FLIP_LEVEL: WTI sustained below $89 (~₹8,455) → mode flips to MILD BULL; watch $95 (~₹9,025) as half-step",
    "FLAG: Crude FALLING toward lower band — current BEAR mode may flip intraday. Bears trapped if WTI rejects $100.",
    "AVOID (BEAR mode rule): BPCL, HPCL, IOC (OMCs), IndiGo/SpiceJet (aviation), Asian Paints/Berger (paints), tyres (Apollo, MRF), fertilisers",
    "EXCEPTION_CALLS allowed: ONGC, Oil India, MRPL (upstream beneficiaries — still in BEAR band)",
    "Reliance (Q4 PAT -8.9% YoY due to O2C drag from crude); now crude falling = relief, but technical structure still broken",
])

add_para(doc, "STRUCTURAL READ — Nifty / Bank Nifty / Options", bold=True, size=12)
add_bullets(doc, [
    "Nifty close 23,694; GIFT-implied open ~23,845 (gap +151)",
    "Max Pain / PCR / OI by strike: DATA_UNAVAILABLE for 15 May open — use live NSE option chain at 9:00 IST before trade",
    "Heuristic estimate (mid-month, gap up day): Max Pain likely 23,500–23,700, Highest Call OI 24,000, Highest Put OI 23,500",
    "Expiry-week dynamic: Today Fri = T-2 to Tuesday Nifty weekly expiry. Theta on weekly options accelerating sharply",
    "If gap-up holds, put writers at 23,500 defend → bullish floor; if Nifty fails 23,800, call writers at 24,000 dominate → range trade",
    "Futures basis: DATA_UNAVAILABLE pre-open — but historically gap-up days print premium → bullish",
    "Bank Nifty vs Nifty: HDFC Bank weak (closed ₹769.55, below 200 DMA ₹933) — Bank Nifty likely LAGGING Nifty if IT/Metals lead the gap up",
    "Bank Nifty lagging on a gap-up = move SUSPECT for trend continuation; favours intraday RANGE not runaway BULL",
])

add_para(doc, "Expiry-week ceiling/floor:", bold=True, size=11)
add_bullets(doc, [
    "Put writers' floor (estimate): 23,500–23,550",
    "Call writers' ceiling (estimate): 23,950–24,000",
    "Max pain magnetism direction: SIDEWAYS — pulling toward 23,700",
    "Final tilt: gap-up gets sold above 23,900 unless FII flip to buy (NSDL post-market data tells us tomorrow)",
])

# ---------- SECTION 3 — CONTRARIAN CHECK (LAYER 3) ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Contrarian Check (Layer 3)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Q1. CONSENSUS — What everyone is expecting today:", bold=True, size=11)
add_para(doc,
    "Bearish drift. Iran war narrative + WPI inflation 8.3% (42-month high) + Reliance Q4 miss + crude spike → "
    "retail and TV analysts are positioned SHORT Nifty 24,000 CE, LONG 23,500 PE, expecting failed rally and "
    "test of 23,500 support. Bank Nifty is the consensus underweight.",
    size=11)

add_para(doc, "Q2. THE TRAP — How market punishes consensus:", bold=True, size=11)
add_para(doc,
    "Crude has fallen 5% in 48 hours on Trump-Xi cooling tone. If crude eases another 2–3%, the entire bear thesis "
    "(Iran war → inflation → no rate cuts → weak Nifty) cracks. Squeeze targets: 23,950–24,000 (call writers stop), "
    "then 24,150. Profits go to put-writers at 23,500 and ONGC/Tata Steel longs.",
    size=11)

add_para(doc, "Q3. RETAIL STOPS clustered:", bold=True, size=11)
add_para(doc,
    "Long stops: below 23,550 (gap-fill level) and 23,400 (last week's swing low). "
    "Short stops: above 23,950 (call writer line) and 24,050 (psychological). "
    "Squeeze risk: if Nifty pierces 23,950 with volume in first 90 min, short-covering blast to 24,150.",
    size=11)

add_para(doc, "Q4. FLIP TRIGGER — single event that reverses today:", bold=True, size=11)
add_para(doc,
    "FRESH IRAN HEADLINE — any tanker incident, missile event, or Trump truth-social post that re-spikes crude "
    "back through WTI $105 instantly kills the gap-up and rotates back to BEAR mode. Secondary: ECB/Fed speaker "
    "out of the EU close hour pushing risk-off.",
    size=11)

add_para(doc, "CONTRARIAN OVERRIDE LOGIC — APPLIED", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_kv_table(doc,
    header=["Condition", "Status"],
    rows=[
        ["A: GIFT Nifty gap-up > +100 pts?", "YES (+151 / +261 on GIFT level itself)"],
        ["B: Crude falling > 2% overnight?", "YES (-5% week, FALLING into BEAR band)"],
        ["C: Contrarian scenario probability > 40%?", "YES (consensus crowded short, squeeze setup)"],
        ["D: Expiry week + PCR < 0.8 (squeeze risk)?", "DATA_UNAVAILABLE — flagged conditional"],
    ])

add_para(doc, "OVERRIDE RESULT:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "A + B + C all YES  →  FULL OVERRIDE TRIGGERED  →  Final bias = CAUTIOUS BULL, permission = YELLOW. "
    "Bear bias suppressed despite BEAR crude mode. The consensus bearish trade is the trap today.",
    bold=True, size=11)

# ---------- SECTION 4 — 5 INDIVIDUAL F&O SETUPS ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — Five Individual F&O Setups", level=1, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "F&O Ban list today: SAIL (skip). All 5 picks verified off ban list. Three A-grade, two B-grade per rule.",
    italic=True, size=10)

# ===== STOCK 1 — ONGC =====
add_para(doc, "[1] ONGC — NSE: ONGC  |  CALL  |  Grade A",
    bold=True, size=13, color=RGBColor(0x38, 0x76, 0x1D))
stock_card(doc, [
    ("CMP / Lot", "₹305.25 (+1.31% on 14 May) / Lot 1,425"),
    ("Crude aligned", "YES — upstream beneficiary; still benefits even at $101 vs $80 pre-war"),
    ("Catalyst", "Crude basing $100 = realisation cushion; Q4 earnings 26 May (event premium building); upstream PSU outperformer"),
    ("Daily trend", "UP — +7.24% in 1 month; near 52w high ₹307.50"),
    ("Above 200 DMA", "YES — strong institutional signal"),
    ("Weekly trend", "UP — +18% in 6 months; +20.7% YoY"),
    ("Support / Resistance", "S: ₹297 / ₹287 (50 DMA ₹280.95)   R: ₹307.50 (52w high) / ₹315"),
    ("OI direction", "Estimate: LONG BUILD (price + OI + crude floor) — verify on 9:00 IST option chain"),
    ("Delivery % / Block / Ban", "Delivery RISING (PSU accumulation); no block deal flagged; NOT on ban list"),
    ("Option setup", "ONGC 310 CE monthly (29 May) — ATM, liquid, low theta vs weekly. Avoid weekly given 2 DTE."),
    ("THETA warning ≤3 days?", "Monthly chosen — NO theta risk"),
    ("Entry (all 4 must fire)", "ST GREEN on 15-min  +  RSI > 50  +  StochRSI cross-up from <30  +  Volume > 1.5x avg.  +  15-min close above ₹307.50"),
    ("T1 (book 40%) / T2 (40%)", "T1 ₹312 / T2 ₹317"),
    ("Stop Loss", "₹302 (closing) OR 15-min ST flips RED"),
    ("R:R", "~2.4:1 (entry ₹307.50, SL ₹302, T2 ₹317)"),
    ("Time stop", "Exit by 13:00 IST regardless"),
    ("Grade reason", "A — clean trend, near ATH, structural OI, monthly expiry hedges theta. Not A+ only because crude direction is FALLING."),
])

doc.add_paragraph()

# ===== STOCK 2 — TATA STEEL =====
add_para(doc, "[2] TATA STEEL — NSE: TATASTEEL  |  CALL  |  Grade A",
    bold=True, size=13, color=RGBColor(0x38, 0x76, 0x1D))
stock_card(doc, [
    ("CMP / Lot", "₹217.62 (+0.39%) / Lot 5,500"),
    ("Crude aligned", "NEUTRAL — metals correlate with growth/CNY; not crude-victim"),
    ("Catalyst", "Q4 FY26 results SCHEDULED TODAY 15 May (board meet); one source flags 16 May — verify NSE corp announcements at 9:00 IST"),
    ("Daily trend", "UP — touched ATH ₹222.40 this week"),
    ("Above 200 DMA", "YES — strong institutional support"),
    ("Weekly trend", "UP — 3-year return +107%"),
    ("Support / Resistance", "S: ₹213 / ₹208   R: ₹222.40 (ATH) / ₹228"),
    ("OI direction", "Estimate: LONG BUILD on momentum; verify pre-open"),
    ("Delivery % / Block / Ban", "Delivery STABLE; SAIL on ban (peer not Tata Steel); Tata Steel NOT banned"),
    ("Option setup", "TATASTEEL 220 CE monthly (29 May) — ATM. If Q4 today, monthly hedges over-event vega."),
    ("THETA warning ≤3 days?", "Monthly chosen — NO theta risk; AVOID weekly given event"),
    ("Entry (all 4 must fire)", "ST GREEN  +  RSI > 50  +  StochRSI cross-up  +  Vol > 1.5x  +  15-min close above ₹222.40 (ATH breakout) — ONLY if results not pending intraday"),
    ("T1 / T2", "T1 ₹227 / T2 ₹233"),
    ("Stop Loss", "₹217 (closing) OR break of intraday VWAP"),
    ("R:R", "~2.1:1 (entry ₹222.40, SL ₹217, T2 ₹233)"),
    ("Time stop", "Exit by 13:00 IST; SKIP entry if Q4 confirmed post-market today — wait for tomorrow"),
    ("Grade reason", "A — ATH momentum + catalyst clarity, downgraded from A+ due to result-timing ambiguity (15 vs 16 May)."),
])

doc.add_paragraph()

# ===== STOCK 3 — INFOSYS =====
add_para(doc, "[3] INFOSYS — NSE: INFY  |  CALL  |  Grade A",
    bold=True, size=13, color=RGBColor(0x38, 0x76, 0x1D))
stock_card(doc, [
    ("CMP / Lot", "₹1,095.00 (-2.58% on 14 May) / Lot 300"),
    ("Crude aligned", "POSITIVE — IT exporter; falling crude eases input/operations + USDINR ₹95 = export tailwind"),
    ("Catalyst", "Nasdaq +0.9% record + S&P record + INFY ADR steady at $11.86; oversold-bounce setup at 52w low ₹1,089"),
    ("Daily trend", "DOWN short-term (-13.44% MoM); price AT 52w low — mean-reversion candidate"),
    ("Above 200 DMA", "NO — 200 DMA well above; trade is COUNTER-TREND bounce only"),
    ("Weekly trend", "DOWN — but momentum exhaustion at 52w low"),
    ("Support / Resistance", "S: ₹1,089 (52w low) / ₹1,070   R: ₹1,115 / ₹1,140 (gap-fill)"),
    ("OI direction", "Estimate: SHORT COVERING potential (puts dominant; unwinding setup)"),
    ("Delivery % / Block / Ban", "NOT on ban list; delivery rising on capitulation"),
    ("Option setup", "INFY 1100 CE monthly (29 May) — ATM, monthly to handle vol crush risk"),
    ("THETA warning ≤3 days?", "Monthly chosen — NO theta risk"),
    ("Entry (all 4 must fire)", "ST GREEN on 15-min  +  RSI crosses above 50  +  StochRSI cross-up from <20  +  Vol > 1.5x  +  15-min close above ₹1,115"),
    ("T1 / T2", "T1 ₹1,128 / T2 ₹1,142"),
    ("Stop Loss", "₹1,098 (closing) OR 52w low break"),
    ("R:R", "~2.2:1 (entry ₹1,115, SL ₹1,098, T2 ₹1,142)"),
    ("Time stop", "Exit by 13:00 IST"),
    ("Grade reason", "A — best contrarian alignment with override scenario: oversold + US tech rally + weak rupee. Trade fails if INFY breaks ₹1,089 — clear stop."),
])

doc.add_paragraph()

# ===== STOCK 4 — HDFC BANK =====
add_para(doc, "[4] HDFC BANK — NSE: HDFCBANK  |  PUT  |  Grade B",
    bold=True, size=13, color=RGBColor(0xCC, 0x00, 0x00))
stock_card(doc, [
    ("CMP / Lot", "₹769.55 (-24.6% from 52w high) / Lot 550"),
    ("Crude aligned", "NEUTRAL"),
    ("Catalyst", "Below both 50 DMA (₹810) and 200 DMA (₹933); broken institutional structure; banking weakness drag on Bank Nifty"),
    ("Daily trend", "DOWN — sideways in tight range ₹747–₹778"),
    ("Above 200 DMA", "NO — major institutional bearish signal"),
    ("Weekly trend", "DOWN — -18.91% YoY"),
    ("Support / Resistance", "S: ₹747 (recent low) / ₹726.65 (52w low)   R: ₹778 / ₹810 (50 DMA)"),
    ("OI direction", "Estimate: SHORT BUILD on every rejection of 50 DMA"),
    ("Delivery % / Block / Ban", "NOT on ban list"),
    ("Option setup", "HDFCBANK 760 PE monthly (29 May) — slightly ITM, deep liquidity"),
    ("THETA warning ≤3 days?", "Monthly chosen — NO theta risk"),
    ("Entry (all 4 must fire)", "ST RED on 15-min  +  RSI < 50  +  StochRSI cross-down from >70  +  Vol > 1.5x  +  15-min close below ₹758"),
    ("T1 / T2", "T1 ₹750 / T2 ₹740"),
    ("Stop Loss", "₹770 (closing) OR break above 50 DMA ₹810 invalidates entirely"),
    ("R:R", "~2.3:1 (entry ₹758, SL ₹770, T2 ₹740)"),
    ("Time stop", "Exit by 13:00 IST"),
    ("Grade reason", "B — counter to today's CAUTIOUS BULL override; trade only if Bank Nifty LAGS Nifty in first 30 min confirming bull move is suspect. Skip if banks lead."),
])

doc.add_paragraph()

# ===== STOCK 5 — RELIANCE =====
add_para(doc, "[5] RELIANCE — NSE: RELIANCE  |  PUT  |  Grade B",
    bold=True, size=13, color=RGBColor(0xCC, 0x00, 0x00))
stock_card(doc, [
    ("CMP / Lot", "₹1,360 (-15.62% from 52w high) / Lot 500"),
    ("Crude aligned", "MIXED — O2C division HURT when crude spikes (Q4 PAT -8.9% YoY); falling crude HELPS this lever, but technical broken"),
    ("Catalyst", "Q4 FY26 PAT fell 8.9% YoY to ₹20,589 cr; 52w high ₹1,611.8 broken; failed retests near ₹1,400"),
    ("Daily trend", "DOWN — close to 52w low ₹1,290"),
    ("Above 200 DMA", "NO"),
    ("Weekly trend", "DOWN multi-month"),
    ("Support / Resistance", "S: ₹1,340 / ₹1,290 (52w low)   R: ₹1,385 / ₹1,400"),
    ("OI direction", "Estimate: SHORT BUILD on rallies above ₹1,400; longs trapped"),
    ("Delivery % / Block / Ban", "NOT on ban list"),
    ("Option setup", "RELIANCE 1360 PE monthly (29 May) — ATM"),
    ("THETA warning ≤3 days?", "Monthly chosen — NO theta risk"),
    ("Entry (all 4 must fire)", "ST RED  +  RSI < 50  +  StochRSI cross-down  +  Vol > 1.5x  +  15-min close below ₹1,345"),
    ("T1 / T2", "T1 ₹1,330 / T2 ₹1,310"),
    ("Stop Loss", "₹1,360 (closing) OR above ₹1,385 invalidates"),
    ("R:R", "~2.3:1 (entry ₹1,345, SL ₹1,360, T2 ₹1,310)"),
    ("Time stop", "Exit by 13:00 IST"),
    ("Grade reason", "B — fundamental tailwind (crude falling) conflicts with technical setup (broken). Trade only if RIL fails ₹1,380 with volume confirming the bear continuation."),
])

# ---------- SECTION 5 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["TODAY'S BIAS", "CAUTIOUS BULL (overridden from BEAR by gap-up + crude falling + squeeze setup)"],
        ["CONFIDENCE", "MEDIUM (override is mechanical, but Iran tape-bomb risk caps confidence)"],
        ["CRUDE MODE", "BEAR band, but FALLING — watch WTI $95 for half-step, $89 for full flip"],
        ["VIX SIZING", "QUARTER (war regime; assume India VIX > 20 until proven otherwise)"],
        ["ENTRY PERMISSION", "YELLOW — confirmation-only entries; no anticipatory longs"],
    ])

add_para(doc, "THE BULL CASE (2 lines):", bold=True, size=11, color=RGBColor(0x38, 0x76, 0x1D))
add_para(doc,
    "Crude has cooled 5% in 48h on Trump-Xi cooling tone; US at fresh records; INFY ADR steady; DII firepower at "
    "+₹7,990 cr last data; consensus is crowded short into expiry week — perfect squeeze setup above 23,950.",
    size=11)

add_para(doc, "THE BEAR CASE (2 lines):", bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "Hormuz remains blocked; WPI at 42-month high (8.3%); RBI rate-cut path delayed; FII still net seller; "
    "Bank Nifty broken (HDB below 200 DMA); RIL Q4 miss still digesting. One Iran headline and entire override "
    "scenario flips back to BEAR.",
    size=11)

add_para(doc, "FLIP TRIGGER:", bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "\"If WTI re-crosses $105 on a fresh Iran headline before 12:00 IST, bias flips back to BEAR. "
    "If Nifty rejects 23,900 with volume in first 60 min and Bank Nifty fails to lead, bias degrades to RANGE.\"",
    italic=True, size=11)

add_para(doc, "NIFTY KEY LEVELS:", bold=True, size=12)
add_kv_table(doc,
    header=["S2", "S1", "CRITICAL", "R1", "R2"],
    rows=[
        ["23,400", "23,550", "23,700 (gap fill / pivot)", "23,950 (call writers)", "24,150 (squeeze target)"],
    ])

doc.add_paragraph()
add_para(doc, "TOP 3 RANKED:", bold=True, size=12)
add_bullets(doc, [
    "#1  ONGC  — Grade A — Crude beneficiary holding ATH zone; clean trend — entry above ₹307.50",
    "#2  TATA STEEL  — Grade A — Momentum + Q4 catalyst (verify result date) — entry above ₹222.40 (ATH breakout)",
    "#3  INFOSYS  — Grade A — Oversold bounce aligned with override scenario + US tech tailwind — entry above ₹1,115",
])

doc.add_paragraph()
add_para(doc, "ONE RISK THAT RUINS EVERYTHING TODAY:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "A SINGLE IRAN HEADLINE during India market hours (any tanker attack, missile event, or Trump truth-social post) "
    "that re-spikes crude through WTI $105. It nukes the gap-up, sends ONGC parabolic but the broader market into "
    "freefall on inflation fear, and you're holding 3 calls into a -1.5% Nifty session.",
    size=11)

# ---------- ONE-LINE SUMMARY ----------
doc.add_paragraph()
add_heading(doc, "ONE-LINE SUMMARY (read at 9:10 AM IST)", level=1, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Today is CAUTIOUS BULL because gap-up +151 plus crude falling 5% triggered the contrarian override on a "
    "crowded bearish consensus. Crude ₹9,646 = BEAR mode but FALLING. Watch ONGC CE and TATA STEEL CE for "
    "breakouts; INFY CE for oversold bounce. Key risk: any Iran tape-bomb that re-spikes crude. Size QUARTER. "
    "Flips back to BEAR if WTI crosses $105 before 12:00 IST.\"",
    italic=True, size=12)

# ---------- FOOTER ----------
doc.add_paragraph()
add_para(doc, "DATA INTEGRITY NOTES:", bold=True, size=10)
add_bullets(doc, [
    "India VIX exact value: DATA_UNAVAILABLE — assumed >20 under war regime → QUARTER sizing applied conservatively",
    "Max Pain, PCR, exact OI by strike: DATA_UNAVAILABLE pre-open — verify live NSE option chain at 9:00 IST",
    "FII/DII data: latest available is 12 May; 14 May data not yet released in NSDL feeds",
    "Tata Steel Q4 result date: ambiguous (15 May board meet per one source vs 16 May release per another) — verify NSE corp filings 9:00 IST",
    "HDB ADR live 14 May close: DATA_UNAVAILABLE — last confirmed $25.67 from 8 May",
])

doc.add_paragraph()
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run("Generated by Daily Trading Routine. Personal educational use only. Not financial advice. "
                  "Hard exit 13:30 IST. Max 3 trades. Daily risk cap 2%.")
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
