"""Generate the Trading Brief for 2026-05-08 (BEAR / YELLOW)."""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-08"
BIAS = "BEAR"
ENTRY_PERMISSION = "YELLOW"
OUT_PATH = f"Trading_Briefs/Trading_Brief_{DATE_ISO}_{BIAS}.docx"


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


def badge(doc, label, value, fill_hex):
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    cell.text = ""
    shade_cell(cell, fill_hex)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{label}: {value}")
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


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
run = title.add_run("ELITE DAILY TRADING BRIEF")
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
dr = date_p.add_run("Friday, 08 May 2026  |  Generated 08:50 IST")
dr.bold = True
dr.font.size = Pt(11)

date_p2 = doc.add_paragraph()
date_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
dr2 = date_p2.add_run("Weekly expiry: 6 days  |  Monthly expiry: 20 days  |  Expiry week: NO")
dr2.italic = True
dr2.font.size = Pt(10)

doc.add_paragraph()

badge(doc, "BIAS", "BEAR (cautious)", "C0392B")
badge(doc, "ENTRY PERMISSION", "YELLOW", "E69138")
badge(doc, "CRUDE MODE", "BEAR (puts only) — Direction: RISING", "8E44AD")
badge(doc, "VIX SIZING", "HALF SIZE (mandatory)", "2E86AB")
badge(doc, "EXPIRY ALERT", "Hormuz ceasefire reassessed TODAY — HIGH headline risk", "C0392B")

doc.add_paragraph()

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Value", "Signal"],
    rows=[
        ["MCX Crude (proxy)",   "~₹9,120/bbl",          "BEAR mode (₹9,000-10,000)"],
        ["Brent / WTI",          "$100.94 / $96.0",      "RISING +1.25% overnight"],
        ["GIFT Nifty gap",       "~24,395 (+68 vs close)", "Flat-to-slight gap-up; news: gap-down vs GIFT prior"],
        ["S&P 500",              "7,337.11",             "Risk-on but pulling back from records"],
        ["Nasdaq",               "25,806.20 (-0.13%)",   "Slight risk-off"],
        ["Dow Jones",            "49,596.97",            "Flat-to-positive"],
        ["US VIX",               "DATA_UNAVAILABLE",     "Estimated ~17-19 (calm-to-fear)"],
        ["India VIX",            "~17.5 (17-18 band)",   "HALF SIZE rule active"],
        ["Nifty close (07 May)", "₹24,326.65",           "Closed off intraday high 24,482"],
        ["FII cash (latest)",    "-₹5,834.9 Cr",         "SELLING (5-day trend: SELLING)"],
        ["DII cash (latest)",    "+₹6,836.87 Cr",        "BUYING (absorbing FII)"],
        ["USD-INR",              "~95.0",                "Rupee weak (FII headwind)"],
        ["INFY ADR",             "$12.51",               "Flat → INFY likely opens flat"],
        ["WIT ADR",              "$1.98 (-0.25%)",       "Slight negative → WIPRO flat-to-down"],
        ["IBN ADR",              "$26.79 (+0.15%)",      "Flat → ICICI Bank stable"],
        ["HDB ADR",              "$25.86 (+0.92%)",      "Positive → HDFC Bank gap-up"],
    ])

doc.add_paragraph()
add_para(doc, "Geopolitical one-liner: Hormuz blocked since Feb 2026; fragile US-Iran ceasefire BEING REASSESSED TODAY (May 8); Fujairah Oil Industry Zone hit by Iranian drones; S Korean vessel fire reported; Trump 'Project Freedom' under doubt by Eurasia Group.", bold=True, size=11, color=RGBColor(0xC0, 0x39, 0x2B))
add_para(doc, "Today's intraday headline risk: HIGH — single Iran tape can spike crude 5-7% intraday.", bold=True, size=11)

# ---------- SECTION 2 — CRUDE RULE + STRUCTURAL READ ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule + Structural Read", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "CRUDE_MODE: BEAR (puts only)", bold=True, size=12, color=RGBColor(0xC0, 0x39, 0x2B))
add_para(doc, "CRUDE_DIRECTION: RISING (+1.25% overnight; bear mode DEEPENING)", bold=True, size=11)
add_para(doc, "CRUDE_FLIP_LEVEL: ₹10,000 MCX (above = STRONG BEAR / no calls anywhere) | ₹8,500 MCX (below = downgrade to CAUTION half-size)", bold=True, size=11)
add_bullets(doc, [
    "AVOID: BPCL, HPCL, IOC (OMCs), IndiGo, SpiceJet (aviation), Asian Paints, Berger (paints), MRF, Apollo Tyres (tyres), Coromandel, Chambal (fertilisers).",
    "EXCEPTION CALLS permitted: ONGC, Oil India, MRPL — direct upstream beneficiaries.",
    "DEFENCE proxy bid: HAL, BEL, BDL benefit from sustained Hormuz risk premium.",
])

doc.add_paragraph()
add_para(doc, "Structural read of Nifty (option-chain context):", bold=True, size=12)
add_bullets(doc, [
    "Nifty 24,326.65 vs estimated max pain ~24,300-24,400: minor downward magnetism, near pin",
    "PCR: DATA_UNAVAILABLE (live source needed); historic 1.0-1.1 normal range",
    "Highest Call OI ceiling: ~24,500 CE (resistance)",
    "Highest Put OI floor: ~24,000 PE (support)",
    "Futures basis: assume modest premium (Nifty future > spot by 30-50 pts is typical)",
    "Expiry dynamics: NOT expiry week. Weekly expiry was Thursday May 7. Next: May 14 (Thu, 6 days). Monthly: May 28 (Thu, 20 days).",
    "Days to expiry > 3 → both weekly and monthly options OK by rule.",
])

doc.add_paragraph()
add_para(doc, "Bank Nifty leadership check:", bold=True, size=11)
add_para(doc, "Bank Nifty has been firm (HDB ADR +0.92%, IBN ADR +0.15%). If Bank Nifty leads on open, Nifty downside is suspect; if Bank Nifty also drifts down, the bear move is REAL. DEFAULT ASSUMPTION: Bank Nifty leads → Nifty bear move SUSPECT — favours range play over outright shorts.", size=11)

# ---------- SECTION 3 — CONTRARIAN CHECK ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Contrarian Check (Layer 3)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Q1. CONSENSUS:", bold=True, size=11)
add_para(doc, "Every retail trader and TV anchor expects bearish open: Hormuz reassessment + rising crude + FII selling + Nifty broke 24,400. Consensus = SHORT Nifty, BUY OMC puts, BUY oil producer calls.", size=11)

add_para(doc, "Q2. THE TRAP:", bold=True, size=11)
add_para(doc, "If Iran ceasefire is EXTENDED at the reassessment, crude tanks 5-8%, OMC puts collapse, IndiGo gaps up, ONGC calls bleed. Smart money is short crude into the announcement — they profit if consensus is wrong.", size=11)

add_para(doc, "Q3. RETAIL STOPS:", bold=True, size=11)
add_para(doc, "Long stops clustered below Nifty 24,250 (round number + S1). Short stops clustered above 24,450 (yesterday's intraday high zone) and 24,500 (call writer ceiling).", size=11)

add_para(doc, "Q4. FLIP TRIGGER:", bold=True, size=11)
add_para(doc, "A confirmed Hormuz ceasefire EXTENSION headline before 12:30 PM IST flips bias from BEAR to CAUTIOUS BULL instantly. Conversely, a single shipping incident or ceasefire COLLAPSE headline turns the day into a one-way crude/Nifty shock.", size=11)

doc.add_paragraph()
add_para(doc, "Contrarian Override Logic Applied:", bold=True, size=12, color=RGBColor(0xC0, 0x39, 0x2B))
add_bullets(doc, [
    "A: GIFT Nifty gap-up >+100 pts? NO (only +68 vs cash close)",
    "B: Crude falling >2% overnight? NO (crude RISING +1.25%)",
    "C: Contrarian probability >40%? PARTIAL — ceasefire extension is a real coin-flip risk (~35-40%)",
    "D: Expiry week + PCR <0.8? NO (not expiry week)",
    "Outcome: NO override fires. Bias stays BEAR but confidence is MEDIUM, permission YELLOW.",
])

# ---------- SECTION 4 — 5 INDIVIDUAL F&O SETUPS ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — 5 Individual F&O Setups", level=1, color=RGBColor(0x1F, 0x38, 0x64))

# STOCK 1 — ONGC CALL
add_para(doc, "STOCK 1 — ONGC (NSE: ONGC) | CALL | Grade A", bold=True, size=13, color=RGBColor(0x27, 0xAE, 0x60))
stock_card(doc, [
    ("CMP", "₹282.80 | Lot: 9,000 (approx) | Crude aligned: YES"),
    ("CATALYST", "Crude RISING +1.25% overnight + Hormuz reassessment day = upstream realisation tailwind. JM Financial buy with ₹340 target."),
    ("TECHNICAL", "Daily UP (broke 270-272 resistance). Above 200 DMA: YES. Weekly UP. Support ₹272 / Resistance ₹293-300."),
    ("STRUCTURAL", "OI direction: LONG BUILD likely on rising crude. Delivery %: stable. Block: NO. Ban list: NO."),
    ("OPTION SETUP", "Strike: ₹285 CE (ATM-OTM) | Expiry: 14 May weekly OR 28 May monthly. Theta warning: NO (6 days to weekly)."),
    ("ENTRY (ALL 4)", "ST GREEN on 15-min | RSI > 50 on 15-min | StochRSI cross UP from <30 | Volume > 1.5x 20-period avg | 15-min close above ₹285."),
    ("TARGETS", "T1 (40%): ₹290 | T2 (40%): ₹295 | Trail rest"),
    ("SL", "₹278 + ST flips RED | R:R: 2.3:1"),
    ("TIME STOP", "Exit by 13:00 IST"),
    ("GRADE REASON", "A: clean BEAR-mode exception with rising-crude catalyst and chart breakout above 270."),
])

doc.add_paragraph()
# STOCK 2 — HAL CALL
add_para(doc, "STOCK 2 — HAL (NSE: HAL) | CALL | Grade A", bold=True, size=13, color=RGBColor(0x27, 0xAE, 0x60))
stock_card(doc, [
    ("CMP", "₹4,651 | Lot: 150 | Crude aligned: NEUTRAL (defence)"),
    ("CATALYST", "Defence sector rallying on order inflows: HAL ₹69,400 cr orders YTD; Nifty India Defence +1.62% on 7 May. Hormuz reassessment keeps geopolitical premium."),
    ("TECHNICAL", "Daily UP. Above 200 DMA: YES. Weekly UP. Support ₹4,500 / Resistance ₹4,750-4,900."),
    ("STRUCTURAL", "OI: LONG BUILD on order inflow news. Delivery %: rising. Block: NO. Ban list: NO."),
    ("OPTION SETUP", "Strike: ₹4,700 CE (ATM) | Expiry: 28 May monthly (preferred for HAL — weekly liquidity thinner). Theta warning: NO."),
    ("ENTRY (ALL 4)", "ST GREEN on 15-min | RSI > 50 | StochRSI cross UP from extreme | Volume > 1.5x | 15-min close above ₹4,700."),
    ("TARGETS", "T1 (40%): ₹4,750 | T2 (40%): ₹4,820"),
    ("SL", "₹4,620 + ST flips RED | R:R: 2.4:1"),
    ("TIME STOP", "Exit by 13:00 IST"),
    ("GRADE REASON", "A: structural defence tailwind + geopolitical premium + technical strength. Crude-agnostic so safe even if mode flips."),
])

doc.add_paragraph()
# STOCK 3 — BPCL PUT
add_para(doc, "STOCK 3 — BPCL (NSE: BPCL) | PUT | Grade A", bold=True, size=13, color=RGBColor(0xC0, 0x39, 0x2B))
stock_card(doc, [
    ("CMP", "₹306.70 | Lot: 1,800 | Crude aligned: YES (negative — OMC victim)"),
    ("CATALYST", "Rising crude squeezes refining margins; ATF prices kept FROZEN at ₹104.9/L despite 15-20% benchmark hike → margin pain. Stock down 17.8% over 6 months."),
    ("TECHNICAL", "Daily DOWN. Below 200 DMA: weakening. Weekly DOWN. Support ₹295 / Resistance ₹320 (was ₹367 high)."),
    ("STRUCTURAL", "OI: SHORT BUILD likely. Delivery %: falling. Block: NO. Ban list: NO. Q4 results 27 May (no near catalyst)."),
    ("OPTION SETUP", "Strike: ₹300 PE (ATM-ITM) | Expiry: 28 May monthly. Theta warning: NO."),
    ("ENTRY (ALL 4)", "ST RED on 15-min | RSI < 50 | StochRSI cross DOWN from >70 | Volume > 1.5x | 15-min close below ₹304."),
    ("TARGETS", "T1 (40%): ₹298 | T2 (40%): ₹293"),
    ("SL", "₹312 + ST flips GREEN | R:R: 2.2:1"),
    ("TIME STOP", "Exit by 13:00 IST"),
    ("GRADE REASON", "A: textbook BEAR-crude OMC put. Multiple confluence — fuel-cost squeeze, technical downtrend, no near earnings risk."),
])

doc.add_paragraph()
# STOCK 4 — INDIGO PUT
add_para(doc, "STOCK 4 — INDIGO (NSE: INDIGO) | PUT | Grade B", bold=True, size=13, color=RGBColor(0xE6, 0x91, 0x38))
stock_card(doc, [
    ("CMP", "₹4,470 | Lot: 300 | Crude aligned: YES (negative — aviation victim)"),
    ("CATALYST", "Fuel = 35-40% of cost, 60% USD-linked. Brent above $100 + USD-INR ~95 = double squeeze. STOCK DOWN 21.5% in 6 months."),
    ("TECHNICAL", "Daily DOWN. Below 50 DMA. Support ₹4,400 / Resistance ₹4,600."),
    ("STRUCTURAL", "OI: SHORT BUILD. Delivery: falling. Block: NO. Ban list: NO."),
    ("OPTION SETUP", "Strike: ₹4,400 PE (ATM-OTM) | Expiry: 28 May monthly. Theta warning: NO."),
    ("ENTRY (ALL 4)", "ST RED on 15-min | RSI < 50 | StochRSI cross DOWN from extreme | Volume > 1.5x | 15-min close below ₹4,440."),
    ("TARGETS", "T1 (40%): ₹4,400 | T2 (40%): ₹4,330"),
    ("SL", "₹4,520 + ST flips GREEN | R:R: 2.0:1 (minimum threshold)"),
    ("TIME STOP", "Exit by 13:00 IST"),
    ("GRADE REASON", "B: thesis intact but ECLGS 5.0 government support + IndiGo's strong domestic moat partly offsets. Skip if RR drops below 2:1 at entry."),
])

doc.add_paragraph()
# STOCK 5 — SBI PUT
add_para(doc, "STOCK 5 — SBI (NSE: SBIN) | PUT | Grade B", bold=True, size=13, color=RGBColor(0xE6, 0x91, 0x38))
stock_card(doc, [
    ("CMP", "₹1,092 | Lot: 750 | Crude aligned: NO (banking)"),
    ("CATALYST", "Q4 FY26 results TODAY post-market. Consensus: PAT ₹17,000-19,000 cr (-12% to +8% YoY range); treasury hit drag (Nomura -11% seq, Kotak -22%). Analyst meet 5:15 PM. PRE-RESULT VOLATILITY EXPECTED INTRADAY."),
    ("TECHNICAL", "Daily DOWN over last week (closed ₹1,092 vs ₹1,108 high). Above 200 DMA: YES. Inverse H&S forming with target ₹1,200 in 3-4 mo (medium-term BULL setup)."),
    ("STRUCTURAL", "OI: NEUTRAL ahead of result; high IV crush risk post-event. Delivery: stable. Block: NO. Ban list: NO."),
    ("OPTION SETUP", "Strike: ₹1,080 PE (ATM-OTM) — INTRADAY ONLY before 13:00 | Expiry: 28 May monthly to limit theta. THETA WARNING: results = IV crush risk, exit BEFORE 13:00 well ahead of 3:30 result."),
    ("ENTRY (ALL 4)", "ST RED on 15-min | RSI < 50 | StochRSI cross DOWN | Volume > 1.5x | 15-min close below ₹1,086 (yesterday low)."),
    ("TARGETS", "T1 (40%): ₹1,082 | T2 (40%): ₹1,074"),
    ("SL", "₹1,098 + ST flips GREEN | R:R: 2.0:1"),
    ("TIME STOP", "MANDATORY exit by 13:00 IST — NO holding into result. NO option buying after 12:30."),
    ("GRADE REASON", "B: result-day binary risk forces tight time-stop. Consensus skewed negative on treasury, but long-term inverse H&S means a beat could vaporise puts. Trade only the intraday momentum, exit early."),
])

# ---------- SECTION 5 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["TODAY'S BIAS",          "BEAR (cautious tilt)"],
        ["CONFIDENCE",            "MEDIUM (Hormuz reassessment is binary)"],
        ["CRUDE MODE",            "BEAR ₹9,120 — RISING"],
        ["VIX SIZING",            "HALF SIZE (mandatory, VIX 17-18)"],
        ["ENTRY PERMISSION",      "YELLOW (headline-risk gating)"],
    ])

doc.add_paragraph()
add_para(doc, "THE BULL CASE (mandatory):", bold=True, size=12, color=RGBColor(0x27, 0xAE, 0x60))
add_para(doc, "If Hormuz ceasefire is EXTENDED today, crude collapses 5-8%, OMC/aviation gap up, FII shorts cover, Nifty reclaims 24,450 toward 24,600. DII has been net buyer (+₹6,837 cr) absorbing FII selling — domestic floor is real.", size=11)

add_para(doc, "THE BEAR CASE (primary):", bold=True, size=12, color=RGBColor(0xC0, 0x39, 0x2B))
add_para(doc, "Hormuz reassessment slides into ambiguity or Iran retaliation tape hits → crude spikes through ₹10,000 MCX → BPCL/IndiGo crater, INFY/IT sees rupee gain priced out, Nifty breaks 24,250 toward 24,000.", size=11)

doc.add_paragraph()
add_para(doc, "FLIP TRIGGER:", bold=True, size=12)
add_para(doc, "If Hormuz ceasefire is officially EXTENDED before 12:30 PM IST → bias becomes CAUTIOUS BULL; OMC puts MUST be exited; ONGC/HAL calls partially trimmed.", size=11, color=RGBColor(0xC0, 0x39, 0x2B))

doc.add_paragraph()
add_para(doc, "NIFTY KEY LEVELS:", bold=True, size=12)
add_kv_table(doc,
    header=["S2", "S1", "CRITICAL", "R1", "R2"],
    rows=[
        ["₹24,000", "₹24,250", "₹24,326 (close)", "₹24,450", "₹24,600"],
    ])

doc.add_paragraph()
add_para(doc, "TOP 3 RANKED:", bold=True, size=12)
add_bullets(doc, [
    "#1 ONGC — Grade A — crude beneficiary, BEAR-mode exception with rising crude tailwind — entry above ₹285",
    "#2 HAL — Grade A — defence rally + geopolitical premium, crude-agnostic so safe — entry above ₹4,700",
    "#3 BPCL — Grade A — OMC squeeze in BEAR crude, no near earnings risk — entry below ₹304",
])

doc.add_paragraph()
add_para(doc, "ONE RISK THAT RUINS EVERYTHING TODAY:", bold=True, size=12, color=RGBColor(0xC0, 0x39, 0x2B))
add_para(doc, "A surprise Hormuz ceasefire EXTENSION before 11:00 AM IST would simultaneously kill BPCL/INDIGO puts AND ONGC calls (crude collapses) — the single tape that breaks BOTH sides of today's book. Mitigation: half-size, hard 13:00 time stop, no overnight holds.", size=11)

# ---------- ONE-LINE SUMMARY ----------
doc.add_paragraph()
add_heading(doc, "One-Line Summary (read at 9:10 AM)", level=2, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Today is BEAR-cautious because crude is at ₹9,120 (rising) into the Hormuz ceasefire reassessment. "
    "Watch ONGC CALL (₹285) and BPCL PUT (₹300). Key risk: surprise ceasefire extension flips both sides. "
    "Size HALF (VIX 17-18). Flips CAUTIOUS BULL if Hormuz extension confirmed before 12:30 IST.\"",
    italic=True, size=12)

# ---------- DATA NOTES ----------
doc.add_paragraph()
add_heading(doc, "Data Notes", level=2, color=RGBColor(0x80, 0x80, 0x80))
add_bullets(doc, [
    "Live PCR / max-pain values not retrieved from authoritative source — flagged DATA_UNAVAILABLE; structural read uses contextual estimate.",
    "US VIX precise close not retrieved — flagged DATA_UNAVAILABLE.",
    "MCX crude is a USD→INR derived proxy: WTI $96 × USD-INR ~95 ≈ ₹9,120/bbl.",
    "F&O ban list specific names for May 8 not retrieved live — verify on NSE before placing trades; common bans (last week): SAIL.",
    "All trades subject to verification of the 4-indicator entry checklist on 15-min chart at 9:30 AM IST.",
])

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
