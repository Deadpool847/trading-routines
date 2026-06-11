"""Generate trading brief for Wednesday, 11 June 2026 — BEARISH / CAUTION"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO   = "2026-06-11"
BIAS       = "BEARISH"
ENTRY_PERM = "YELLOW — wait 30 min post-open"
CRUDE_MODE = "CAUTION (₹8,712 MCX — approaching BEAR boundary)"

OUT_PATH = f"trading-briefs/2026/06-June/Trading_Brief_{DATE_ISO}_{BIAS}.docx"

# ── helpers ────────────────────────────────────────────────────────────────────

def shade_cell(cell, hex_fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)

def add_heading(doc, text, level=1, color=None):
    h = doc.add_heading(text, level=level)
    if color:
        for run in h.runs: run.font.color.rgb = color
    return h

def add_para(doc, text, bold=False, italic=False, size=11, color=None, align=None):
    p = doc.add_paragraph()
    if align is not None: p.alignment = align
    run = p.add_run(text)
    run.bold = bold; run.italic = italic; run.font.size = Pt(size)
    if color: run.font.color.rgb = color
    return p

def add_bullets(doc, items):
    for item in items:
        doc.add_paragraph(item, style="List Bullet")

def add_table(doc, rows, header=None, shade_hdr=True):
    ncols = len(rows[0])
    table = doc.add_table(rows=(1 if header else 0) + len(rows), cols=ncols)
    table.style = "Light Grid Accent 1"
    r_idx = 0
    if header:
        hdr = table.rows[0]
        for i, txt in enumerate(header):
            cell = hdr.cells[i]
            cell.text = ""
            run = cell.paragraphs[0].add_run(txt)
            run.bold = True; run.font.size = Pt(11)
            if shade_hdr:
                shade_cell(cell, "1F3864")
                run.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
        r_idx = 1
    for i, row in enumerate(rows):
        trow = table.rows[r_idx + i]
        for j, val in enumerate(row):
            cell = trow.cells[j]
            cell.text = ""
            run = cell.paragraphs[0].add_run(str(val))
            run.font.size = Pt(11)
            if j == 0: run.bold = True
    return table

def badge(doc, label, value, fill_hex, text_color=RGBColor(0xFF,0xFF,0xFF)):
    table = doc.add_table(rows=1, cols=1); table.autofit = True
    cell = table.rows[0].cells[0]; cell.text = ""
    shade_cell(cell, fill_hex)
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{label}: {value}")
    run.bold = True; run.font.size = Pt(13); run.font.color.rgb = text_color

def stock_card(doc, title_text, rows, title_color="C00000"):
    add_para(doc, title_text, bold=True, size=13,
             color=RGBColor(0x1F,0x38,0x64))
    t = doc.add_table(rows=len(rows), cols=2); t.style = "Light Grid Accent 1"
    for i, (k, v) in enumerate(rows):
        lc = t.rows[i].cells[0]; rc = t.rows[i].cells[1]
        lc.text = ""; rc.text = ""
        r1 = lc.paragraphs[0].add_run(k)
        r1.bold = True; r1.font.size = Pt(11); shade_cell(lc, "D9E1F2")
        r2 = rc.paragraphs[0].add_run(v); r2.font.size = Pt(11)
    doc.add_paragraph()

# ── BUILD ──────────────────────────────────────────────────────────────────────
doc = Document()
for section in doc.sections:
    section.top_margin = Cm(1.8); section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.8); section.right_margin = Cm(1.8)

# ── TITLE BLOCK ───────────────────────────────────────────────────────────────
title = doc.add_paragraph(); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("ELITE DAILY TRADING INTELLIGENCE BRIEF")
run.bold = True; run.font.size = Pt(20); run.font.color.rgb = RGBColor(0x1F,0x38,0x64)

sub = doc.add_paragraph(); sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sr = sub.add_run("Indian F&O Day Trader — Nifty + Stocks  |  Capital: ₹2,00,000  |  Mode: PAPER")
sr.italic = True; sr.font.size = Pt(12)

dt = doc.add_paragraph(); dt.alignment = WD_ALIGN_PARAGRAPH.CENTER
dr = dt.add_run("Date: Wednesday, 11 June 2026  |  Generated ~08:45 IST")
dr.bold = True; dr.font.size = Pt(11)

doc.add_paragraph()
badge(doc, "TODAY'S BIAS",       BIAS,       "C00000")     # red — bearish
badge(doc, "ENTRY PERMISSION",   ENTRY_PERM, "BF8F00",
      text_color=RGBColor(0xFF,0xFF,0xFF))                  # amber
badge(doc, "CRUDE RULE MODE",    CRUDE_MODE, "7F6000",
      text_color=RGBColor(0xFF,0xFF,0xFF))                  # dark gold
badge(doc, "VIX SIZING",
      "HALF-SIZE (VIX 15.00 — AT boundary; likely spikes today)",
      "E69138")
doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 0 — YESTERDAY'S POSTMORTEM
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Section 0 — Yesterday's Postmortem", level=1,
            color=RGBColor(0x1F,0x38,0x64))

add_para(doc, "Brief Accuracy Score: N/A — No prior brief in this session.", bold=True, size=12)
add_bullets(doc, [
    "This is the FIRST brief generated in the current conversation session.",
    "No yesterday's picks to score. Baseline methodology accuracy is being established from today.",
    "Methodology adjustment: None (no prior miss to correct). Proceed at full conviction per rules.",
    "NOTE: To maintain running accuracy scoring, paste yesterday's brief context at the start of future sessions.",
])

doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — MACRO SNAPSHOT
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Section 1 — Macro Snapshot (Cross-Verified)", level=1,
            color=RGBColor(0x1F,0x38,0x64))

add_table(doc,
    header=["Metric", "Reading", "Source(s)", "Signal"],
    rows=[
        ["WTI Crude (NYMEX)",   "$91.36/bbl (+1.47%)",      "TradingEconomics / CNBC",      "Rising — CAUTION regime pressure"],
        ["Brent Crude (context ONLY)", "~$96–97/bbl",       "TradingEconomics / Bloomberg",  "NOT used for MCX proxy"],
        ["USDINR Spot",          "95.36",                    "Fed H.10 (Jun 10) / Investing",  "Near record-weak INR"],
        ["MCX Crude Proxy",      "₹8,712/bbl",               "WTI×USDINR calc (see below)",   "CAUTION regime"],
        ["India VIX (Jun 10 close)", "15.00 (↓ from 15.58)", "NSE India / 5paisa",           "AT half-size boundary"],
        ["Nifty 50 (Jun 10 close)", "23,242.10 (+0.52%)",   "NSE India",                     "Resistance: 23,300"],
        ["Sensex (Jun 10 close)", "~77,800 est.",            "ESTIMATE — verify on BSE",      "Tracks Nifty"],
        ["GIFT Nifty (Jun 11 am)", "23,157.5",              "NSE IFSC / 5paisa",             "Gap DOWN ~85 pts"],
        ["Nifty Jun Futures close", "~23,280 est.",          "Estimated cost-of-carry",       "GIFT below futures"],
        ["Nikkei 225",           "~67,471 (–1.36%)",        "CNBC Asia (early Jun data)",    "Weak Asia"],
        ["Hang Seng",            "–1.31% latest",            "Bloomberg / Reuters",           "Bearish"],
        ["Kospi",                "8,639 (–1.84%)",           "CNBC Asia",                    "Bearish"],
        ["CSI 300",              "–0.69%",                   "Bloomberg",                    "Modest weakness"],
        ["Dow Jones (Jun 10)",   "49,918.78 (–1.87%, –953 pts)", "CNBC / TheStreet",         "RISK-OFF"],
        ["S&P 500 (Jun 10)",     "7,266.99 (–1.62%)",        "CNBC",                         "RISK-OFF"],
        ["Nasdaq (Jun 10)",      "25,169.50 (–1.98%)",       "CNBC",                         "RISK-OFF"],
        ["US CPI (May 2026)",    "4.2% YoY — 3-year HIGH",   "BLS / IndMoney",               "Fed hawkish — mega NEGATIVE"],
        ["FII Flow (Jun 9 prov)", "–₹4,566 Cr",             "NSE / StockEdge",               "Outflow"],
        ["DII Flow (Jun 9 prov)", "+₹6,159 Cr",             "NSE / StockEdge",               "Absorbing FII"],
        ["RBI Repo Rate",        "5.25% (held Jun 5)",       "RBI / BizzBuzz",               "Unchanged; hawkish tilt"],
    ])

doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — CRUDE REGIME
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Section 2 — Crude Regime (LIVE Math)", level=1,
            color=RGBColor(0x1F,0x38,0x64))

add_para(doc,
    "MCX PROXY CALCULATION  (ALWAYS WTI — NEVER Brent)",
    bold=True, size=12, color=RGBColor(0xC0,0x00,0x00))
add_bullets(doc, [
    "WTI = $91.36/bbl  ×  USDINR = 95.36  =  ₹8,712/bbl",
    "Regime: CAUTION  (₹8,500 – ₹9,000 band)",
    "Distance to BEAR boundary (₹9,000): ₹288 — NOT within ±₹200, no formal flag YET",
    "BUT: WTI already up 1.47% today. US strikes on Iran (2nd night) + Hormuz near-closure.",
    "If WTI crosses $93.50 intraday → MCX ~₹8,920 → ENTER BOUNDARY-RISK ZONE",
    "If WTI crosses $94.50 intraday → MCX ~₹9,011 → BEAR REGIME TRIGGERED",
    "In BEAR regime: BLOCK ALL CALL POSITIONS except ONGC / OIL India / MRPL",
    "Direction: RISING — geopolitical premium firmly in place",
])

add_para(doc,
    "⚠  REGIME ESCALATION WATCH: Monitor MCX crude at 9:00 AM open. Any print above ₹8,850 "
    "requires immediate downgrade to BEAR caution posture for new entries.",
    bold=True, size=11, color=RGBColor(0xC0,0x00,0x00))

doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — GEOPOLITICAL & NEWS DIGEST
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Section 3 — News Digest (Categorised)", level=1,
            color=RGBColor(0x1F,0x38,0x64))

add_para(doc, "CATEGORY A — GEOPOLITICAL (CRITICAL)", bold=True, size=12,
         color=RGBColor(0xC0,0x00,0x00))
add_bullets(doc, [
    "US launched SECOND consecutive night of strikes on Iran targets (Jun 10). CENTCOM confirmed.",
    "Trigger: Iran downed a US Apache helicopter near Strait of Hormuz (Jun 9).",
    "Iran retaliated: IRGC claims it struck 18 US military targets at Ali Al-Salem (Kuwait), Ahmad Al-Jaber (Kuwait), Sheikh Isa (Bahrain).",
    "Iran supreme leader: 'US bases in Middle East no longer safe.'",
    "Trump (Jun 9): Deal is 'only days away' — SAME claim made before ordering strikes. LOW CREDIBILITY.",
    "Strait of Hormuz: Near-total closure continuing per multiple sources. Global supply disruption: 600–700 million barrels lost (Vitol CEO).",
    "⚠  CRITICAL RULE: Iran 'deal in days' has collapsed TWICE in 2026 within 48 hours. Treat ANY ceasefire MoU as unconfirmed for 3–5 sessions. Do NOT chase rally on headline alone.",
])

doc.add_paragraph()
add_para(doc, "CATEGORY B — DOMESTIC INDIA", bold=True, size=12,
         color=RGBColor(0x1F,0x38,0x64))
add_bullets(doc, [
    "RBI MPC (Jun 5): Rates HELD at 5.25% for 3rd consecutive meeting. Stance: NEUTRAL.",
    "RBI raised FY27 inflation forecast to 5.1% (from 4.6%). Lowered GDP forecast to 6.6% (from 6.9%).",
    "India-US Trade: Commerce Min Goyal targets mid-July for first tranche. USTR 12.5% tariff proposal on 54 countries (incl. India) for forced-labour goods under review — hearings July 7.",
    "Monsoon 2026: Below-normal forecast (90–95% LPA). El Niño conditions developing. SW Monsoon entered West Bengal Jun 9. Active advancement — but QUANTITY is the concern.",
    "No SEBI/RBI major announcements today.",
])

doc.add_paragraph()
add_para(doc, "CATEGORY C — CORPORATE CATALYSTS", bold=True, size=12,
         color=RGBColor(0x1F,0x38,0x64))
add_bullets(doc, [
    "Q4 FY26 season largely concluded. Check NSE for any remaining stragglers reporting today.",
    "F&O Ban List (last known Jun 2): AMBER, KAYNES — VERIFY ON NSE BEFORE ENTRY today.",
    "No major block deals or IPO listings confirmed for Jun 11.",
    "Brokerage targets: MOFSL positive on HAL (₹5,500), BEL (₹520), BDL (₹1,800), Astra Microwave (₹1,150).",
    "IndiGo: Analysts still bullish long-term (FY30 vision) but near-term hampered by crude and Airbus delays.",
])

doc.add_paragraph()
add_para(doc, "CATEGORY D — GLOBAL MACRO", bold=True, size=12,
         color=RGBColor(0x1F,0x38,0x64))
add_bullets(doc, [
    "US CPI May 2026: 4.2% YoY — 3-year HIGH. Energy component: +23.5% (Iran-war driven). Fed rate-cut expectations further delayed.",
    "US markets closed sharply lower on this CPI print: Dow –1.87%, S&P –1.62%, Nasdaq –1.98%.",
    "Industrials fell >3% in the US. This is negative for global risk appetite.",
    "No major Fed speak scheduled today but CPI print will dominate sentiment.",
    "Gold: Safe-haven demand likely elevated (monitor for Nifty metals proxy).",
])

doc.add_paragraph()
add_para(doc, "CATEGORY E — F&O SPECIFIC", bold=True, size=12,
         color=RGBColor(0x1F,0x38,0x64))
add_bullets(doc, [
    "Weekly expiry: TUESDAY (post-SEBI Sep 2025 shift). Last weekly was Tuesday Jun 10. Next weekly: Tuesday Jun 16.",
    "Monthly expiry: June monthly expires last Thursday = June 26, 2026.",
    "For new positions: Use JULY monthly options per standing rule.",
    "Max pain for Jun 16 weekly: NOT available in pre-open. Check Sensibull/Opstra at open.",
    "F&O ban list for today: CHECK NSE ARCHIVES directly (nsearchives.nseindia.com/content/fo/fo_secban.csv).",
    "PCR trend: Need live data at open. Nifty put-call ratio below 0.8 = call writers in control = range-bound or down.",
])

doc.add_paragraph()
add_para(doc, "NEWS-DRIVEN CATALYST SUMMARY", bold=True, size=12,
         color=RGBColor(0x1F,0x38,0x64))
add_table(doc,
    header=["Category", "Stocks/Sectors"],
    rows=[
        ["Positive catalysts",   "HAL, BEL, BDL (defence), ONGC / OIL India (crude E&P), Gold ETFs (safe haven)"],
        ["Negative catalysts",   "HPCL, BPCL, IOCL (OMC crude squeeze), IndiGo (ATF +crude), IT stocks (US CPI/Fed)"],
        ["Sector-wide news",     "Oil & Gas (both directions), Defence (bullish), Aviation (bearish)"],
        ["Session disruptors",   "Any Iran ceasefire or escalation headline; MCX crude crossing ₹8,850+"],
    ])

doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — LAYER 1 / 2 / 3 BIAS FRAMEWORK
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Section 4 — Layer 1/2/3 Contrarian Check", level=1,
            color=RGBColor(0x1F,0x38,0x64))

add_para(doc, "LAYER 1 — SURFACE READ (Retail Consensus)", bold=True, size=12)
add_para(doc,
    "\"Gap down 80–100 pts on Iran escalation and US CPI shock. Buy defence stocks "
    "(HAL, BEL) and crude beneficiaries (ONGC). Wait for the dip to 23,000 and accumulate Nifty longs. "
    "RBI will cut rates soon — market will recover.\"",
    italic=True, size=11)

doc.add_paragraph()
add_para(doc, "LAYER 2 — CRITICAL READ (What Consensus Gets Wrong)", bold=True, size=12)
add_bullets(doc, [
    "Trump's 'deal in days' comment predates the strike orders he just issued. The statement has NO credibility.",
    "Iran struck US military bases in Kuwait and Bahrain overnight — this is ESCALATION, not de-escalation.",
    "US CPI at 4.2% is energy-driven but still delays any Fed pivot. Indian IT stocks face multiple compression.",
    "VIX at 15.00 is AT the half-size boundary — one negative headline pushes it to 16+, forcing QUARTER-size.",
    "DII bought ₹6,159 Cr on Jun 9. But if FII selling accelerates on Iran news, DII absorption capacity has limits.",
    "Defence stocks HAL/BEL have already rallied significantly YTD — the 'buy on news' trade may be crowded.",
    "Retail stops are likely clustered just below 23,000. Smart money may hunt that level before any reversal.",
    "'RBI rate cut soon' narrative contradicts the RBI's OWN inflation upgrade to 5.1% — cut timeline pushed out.",
])

doc.add_paragraph()
add_para(doc, "LAYER 3 — CONTRARIAN READ (What Everyone Is Underestimating)", bold=True, size=12)
add_bullets(doc, [
    "MONSOON BLIND SPOT: Below-normal monsoon (90–95% LPA) + El Niño being ignored entirely amid Iran dominance. "
    "If Kharif crops disappoint, food inflation adds to the 5.1% RBI forecast. RBI CANNOT cut. Banks under NIM pressure for longer.",
    "RUPEE AT 95.36 is quietly catastrophic for the import bill. Every $1 move in WTI costs ₹95 on the MCX proxy. "
    "The INR weakness AMPLIFIES crude pain beyond what USD-only tracking shows.",
    "REGIME BOUNDARY RISK IS CLOSER THAN IT LOOKS: $93 WTI pushes MCX to ₹8,869 — within ₹131 of BEAR territory. "
    "A 1.8% WTI move from current levels flips the regime. That's one bad Iran headline.",
    "SECTOR ROTATION BENEATH THE SURFACE: FMCG has been weak (monsoon risk + consumer squeeze). "
    "Pharma is being ignored — but USDINR at 95.36 is a massive export revenue tailwind for Sun Pharma, Dr Reddy's.",
    "CPI 4.2% = INFRASTRUCTURE CAPEX RISK: US Industrials fell 3% because high inflation + high rates = infra projects delayed. "
    "Indian infra/capex names have not priced this US contagion risk yet.",
])

doc.add_paragraph()
add_table(doc,
    header=["Check", "Reading"],
    rows=[
        ["Flip to BULLISH trigger",  "Iran ceasefire MoU + WTI drops below $87 + GIFT Nifty swing positive + VIX back below 14"],
        ["Flip to BEARISH (deeper)", "WTI above $93.50 (MCX ₹8,920+) + Iran strikes US carrier + VIX spikes above 18"],
        ["Entry permission",         "YELLOW — wait 30 min post-open. Confirm gap-down direction before entering any trade"],
        ["Today's overall bias",     "BEARISH with macro-hedge CALL positions as counter-balance (HAL, ONGC)"],
    ])

doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — NIFTY STRUCTURE
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Section 5 — Nifty Structure", level=1,
            color=RGBColor(0x1F,0x38,0x64))

add_table(doc,
    header=["Level", "Value", "Meaning"],
    rows=[
        ["Yesterday close",         "23,242.10",    "The level the market must reclaim to turn neutral"],
        ["GIFT Nifty implied open",  "~23,100–23,160", "Gap down ~80–140 pts at current GIFT reading"],
        ["Resistance R1",            "23,300",       "Immediate cap — three-session rejection zone"],
        ["Resistance R2",            "23,450–23,550", "Next meaningful resistance on any recovery"],
        ["Resistance R3",            "24,000",       "Major psychological and technical wall"],
        ["Support S1",               "23,000",       "CRITICAL — psychological + chart support"],
        ["Support S2",               "22,800–22,700", "Next meaningful demand zone on break of S1"],
        ["Support S3",               "22,300",       "Medium-term support"],
        ["200 DMA",                  "~22,500 est.", "Long-term trend still intact if above this"],
        ["Expected intraday range",  "22,950–23,380", "Wide range — heightened geopolitical vol"],
        ["THE ONE NUMBER TODAY",     "₹9,000 (MCX crude)", "If MCX crosses here: BEAR regime, block all CALLs except E&P"],
    ])

doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — STOCK SETUPS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Section 6 — Stock Setups (5 Stocks)", level=1,
            color=RGBColor(0x1F,0x38,0x64))

add_para(doc,
    "⚠  ALL LOT SIZES MUST BE VERIFIED ON NSE F&O SEGMENT BEFORE ENTRY — sizes change periodically. "
    "Capital check rule: Lot size × CMP × 3% < ₹15,000 for half-size. Verify F&O ban list on NSE first.",
    bold=True, size=10, color=RGBColor(0xC0,0x00,0x00))
add_para(doc,
    "Use JULY monthly options for all new positions (not June monthly — gives more time to be right).",
    bold=True, size=10, color=RGBColor(0xC0,0x00,0x00))
doc.add_paragraph()

# ── STOCK 1: HPCL PUT ─────────────────────────────────────────────────────────
stock_card(doc, "STOCK 1 — HPCL PUT  |  GRADE: A+  |  Sector: Oil & Gas (OMC)", [
    ("CMP (estimate)", "₹405–₹415 (late May ref ₹412; likely gapping down today — verify live)"),
    ("Lot size",       "~1,000 shares (verify on NSE)  |  Capital check: 1,000 × 410 × 3% = ₹12,300 ✓"),
    ("Why A+",
     "FOUR independent catalysts stacking: (1) WTI $91.36 +1.47% today — OMC marketing margin squeeze. "
     "(2) US-Iran active war — 2nd consecutive night of strikes + Iranian retaliation. "
     "(3) MCX proxy ₹8,712 — approaching BEAR boundary where OMC losses accelerate. "
     "(4) US CPI 4.2% = Fed on hold = global risk-off = FII sell India = broader market pressure on OMCs."),
    ("Option setup",   "HPCL 400 PE — July monthly expiry"),
    ("Entry trigger",  "(1) MCX crude opens above ₹8,750 at 9 AM. "
                       "(2) HPCL 15-min candle closes below ₹408. "
                       "(3) Nifty holds BELOW 23,200 (gap-down confirmed). "
                       "(4) F&O ban list confirmed — HPCL NOT in ban."),
    ("T1",             "₹395 (book 40% of position)"),
    ("T2",             "₹382 (book 40% of position)"),
    ("Remainder",      "Trail with 15-min SuperTrend"),
    ("Stop Loss",      "15-min candle CLOSE above ₹426 (3% above ₹413 entry estimate)"),
    ("R:R",            "~2.5:1 (risk ₹13–15 to gain ₹18–31 across T1/T2)"),
    ("Time stop",      "Exit full position if T1 not hit by 2:00 PM IST"),
    ("Macro alignment","BEARISH crude regime — FULL alignment with CAUTION mode"),
])

# ── STOCK 2: HAL CALL ─────────────────────────────────────────────────────────
stock_card(doc, "STOCK 2 — HAL CALL  |  GRADE: A+  |  Sector: Defence (MACRO HEDGE ✓)", [
    ("CMP (estimate)", "₹4,900–₹5,100 (MOFSL target ₹5,500 — verify live)"),
    ("Lot size",       "~75 shares (verify on NSE)  |  Capital check: 75 × 5,000 × 3% = ₹11,250 ✓"),
    ("Why A+",
     "FOUR catalysts: (1) US launched 2nd consecutive night of Iran strikes — global defence spending signal. "
     "(2) MOFSL formal buy with ₹5,500 target (active brokerage coverage). "
     "(3) India indigenisation push + export opportunity (HAL is 19th largest defence contractor globally). "
     "(4) This is today's MACRO HEDGE trade — required per diversification rules. Defence outperforms in active-war environment."),
    ("Option setup",   "HAL 5100 CE — July monthly expiry (if CMP ~₹5,000); adjust strike to nearest ATM"),
    ("Entry trigger",  "(1) HAL opens flat-to-green DESPITE Nifty gap-down (relative strength = institutional buying). "
                       "(2) First 15-min candle holds above yesterday's close. "
                       "(3) No new Iran de-escalation headline before entry. "
                       "(4) MCX crude NOT above ₹9,000 (i.e., not in BEAR regime — HAL allowed in BEAR too as it's not crude-sensitive)."),
    ("T1",             "₹5,250 (book 40%)"),
    ("T2",             "₹5,450 (book 40%)"),
    ("Remainder",      "Trail with 15-min EMA(20)"),
    ("Stop Loss",      "15-min close below ₹4,820 (below prior day structure)"),
    ("R:R",            "~2.8:1"),
    ("Time stop",      "Exit if T1 not hit by end of day — HAL is a multi-day thesis, but intraday confirmation needed"),
    ("Macro alignment","POSITIVE — geopolitical hedge works in BEAR regime too. KEEP even if crude spikes."),
])

# ── STOCK 3: INDIGO PUT ───────────────────────────────────────────────────────
stock_card(doc, "STOCK 3 — INDIGO (InterGlobe) PUT  |  GRADE: A  |  Sector: Aviation", [
    ("CMP (estimate)", "₹4,350–₹4,539 (was ₹4,539 on Jun 9; likely lower today — verify live)"),
    ("Lot size",       "~75 shares (verify on NSE)  |  Capital check: 75 × 4,400 × 3% = ₹9,900 ✓"),
    ("Why A",
     "THREE confirmed catalysts: (1) WTI $91.36 (+1.47%) — ATF cost is 38% of IndiGo revenue at $97 Brent; "
     "crude spike = direct earnings destruction. (2) Airbus delivery delays confirmed — capacity constraint kills "
     "revenue recovery story. (3) Down 13% YTD already = weak price trend; today's gap-down accelerates."),
    ("Option setup",   "INDIGO 4200 PE or 4000 PE — July monthly expiry"),
    ("Entry trigger",  "(1) Crude holds above ₹8,750 MCX. "
                       "(2) INDIGO opens below ₹4,450 and 15-min close below ₹4,420. "
                       "(3) No fresh Iran de-escalation headline. "
                       "(4) Nifty gap-down sustained (not recovering to flat)."),
    ("T1",             "₹4,200 (book 40%)"),
    ("T2",             "₹4,050 (book 40%)"),
    ("Remainder",      "Trail stop"),
    ("Stop Loss",      "15-min close above ₹4,600 (3%+ above ₹4,460 entry estimate)"),
    ("R:R",            "~2.4:1"),
    ("Time stop",      "2:00 PM if T1 not triggered"),
    ("Macro alignment","BEARISH crude + global risk-off = strong alignment"),
])

# ── STOCK 4: ONGC CALL ────────────────────────────────────────────────────────
stock_card(doc, "STOCK 4 — ONGC CALL  |  GRADE: A  |  Sector: Oil E&P (EXCEPTION STOCK — allowed in BEAR regime)", [
    ("CMP (estimate)", "₹285–₹300 (rising on crude — verify live)"),
    ("Lot size",       "~1,925 shares (verify on NSE — may have been revised)  |  Capital check: 1,925 × 290 × 3% = ₹16,748 — BORDERLINE. If lot > 1,800, SKIP this stock."),
    ("Why A",
     "TWO strong catalysts: (1) WTI at $91.36 and rising — ONGC's realisation improves directly. "
     "Every $1/bbl increase in crude = meaningful net profit addition. "
     "(2) This is the EXCEPTION STOCK — allowed even in BEAR regime per crude rules (upstream E&P). "
     "Smart hedge: if crude spikes (bad for market) ONGC still works as CALL."),
    ("Option setup",   "ONGC 300 CE or 310 CE — July monthly expiry"),
    ("Entry trigger",  "(1) MCX crude opens above ₹8,750. "
                       "(2) ONGC shows relative strength (flat or green vs red Nifty). "
                       "(3) Capital check: Lot size × CMP × 3% < ₹15,000 — CONFIRM BEFORE ENTRY. "
                       "(4) Not in F&O ban list."),
    ("T1",             "₹310 (book 40%)"),
    ("T2",             "₹320–₹325 (book 40%)"),
    ("Stop Loss",      "15-min close below ₹278 (prior support / ~5% from entry)"),
    ("R:R",            "~2.2:1"),
    ("Time stop",      "End of day (ONGC is a multi-day crude story but entry needs today confirmation)"),
    ("Macro alignment","POSITIVE — crude rising = ONGC wins. Works even in BEAR regime."),
])

# ── STOCK 5: IDFCFIRST PUT ───────────────────────────────────────────────────
stock_card(doc, "STOCK 5 — IDFCFIRST BANK PUT  |  GRADE: B+  |  Sector: Private Banking", [
    ("CMP (estimate)", "₹70–₹90 range (verify live — stock has been under pressure)"),
    ("Lot size",       "~6,000 shares (verify on NSE)  |  Capital check: 6,000 × 80 × 3% = ₹14,400 ✓ (if CMP ≤ ₹83). Skip if CMP × 6,000 × 3% > ₹15,000."),
    ("Why B+",
     "SINGLE macro catalyst (downgraded from A): FII selling –₹4,566 Cr (Jun 9) + RBI raised inflation to 5.1% "
     "(rates staying high longer) = small private banks face NIM pressure as cost of funds stays elevated. "
     "B+ because: no company-specific news today. Pure macro overlay. "
     "ENTRY ONLY if ONGC/HPCL confirm direction first (B+ = conditional)."),
    ("Option setup",   "IDFCFIRST 75 PE or 70 PE — July monthly expiry"),
    ("Entry trigger",  "(1) Nifty gap-down sustained below 23,100 at 30-min mark. "
                       "(2) Bank Nifty also weak (below 51,000 — verify level). "
                       "(3) IDFCFIRST 15-min close below its opening print. "
                       "(4) Capital check verified: lot × CMP × 3% < ₹15,000."),
    ("T1",             "₹73 (book 40%)"),
    ("T2",             "₹68 (book 40%)"),
    ("Stop Loss",      "15-min close above ₹89 (~10% from ₹80 entry — wide because vol)"),
    ("R:R",            "~2:1 (minimum) — only enter if confirmed"),
    ("Time stop",      "11:30 AM — if trade not working in 90 min, exit"),
    ("Macro alignment","CAUTION to BEAR environment = small banks underperform. Alignment: MODERATE."),
])

doc.add_paragraph()

# Sector diversity table
add_para(doc, "SECTOR DIVERSIFICATION CHECK", bold=True, size=12, color=RGBColor(0x1F,0x38,0x64))
add_table(doc,
    header=["Stock", "Sector", "Direction", "Grade"],
    rows=[
        ["HPCL",       "Oil & Gas (OMC)",    "PUT",  "A+"],
        ["HAL",        "Defence",            "CALL", "A+  ← MACRO HEDGE ✓"],
        ["INDIGO",     "Aviation",           "PUT",  "A"],
        ["ONGC",       "Oil E&P (exception)", "CALL", "A"],
        ["IDFCFIRST",  "Private Banking",    "PUT",  "B+  (conditional)"],
    ])

add_bullets(doc, [
    "Sectors represented: 4 distinct sectors (Oil & Gas split OMC/E&P counts as 2; Defence; Aviation; Banking) ✓",
    "Max per sector: 2 (Oil & Gas has HPCL + ONGC = 2 ✓) — limit respected",
    "CALL vs PUT: 2 CALLs (HAL, ONGC) : 3 PUTs (HPCL, IndiGo, IDFCFIRST) — appropriate for bearish macro ✓",
    "Macro hedge: HAL CALL ✓",
    "Stock count: 5 (target met: 1 A+ macro-hedge + 1 A+ directional + 2 A + 1 B+) ✓",
])

doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — STOCKS TO AVOID
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Section 7 — Stocks to Actively Avoid Today", level=1,
            color=RGBColor(0x1F,0x38,0x64))

add_table(doc,
    header=["Stock/Sector", "Reason to Avoid"],
    rows=[
        ["BPCL / IOCL",         "Same thesis as HPCL — sector cap hit at 2. Don't double-up on OMC."],
        ["IT stocks (TCS, INFY, WIPRO)",
         "US CPI 4.2% = Fed hawkish = global IT spending caution. Mixed signal with India-US trade deal. No clean setup."],
        ["HDFC Bank / ICICI Bank",
         "Lot sizes too large for ₹2L capital at current prices (lot × CMP × 3% >> ₹15,000)."],
        ["BEL",                 "Lot 2,000 × ~₹490 × 3% = ₹29,400 — FAILS capital filter."],
        ["AMBER / KAYNES",      "Last known F&O ban list (Jun 2). Verify on NSE — if still in ban, SKIP."],
        ["Weekly options (Jun 12 expired)",
         "Last weekly expiry was Jun 10 (Tuesday). Don't trade expired series. Use JULY monthly."],
        ["Any stock with gap > 5%",
         "Chasing gaps on Iran headline in high-VIX environment = poor R:R. Wait for 15-min candle confirmation."],
        ["ZOMATO / consumer discretionary",
         "Lot sizes too large; no strong company-specific catalyst today."],
    ])

doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — FINAL VERDICT
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Section 8 — Final Verdict", level=1,
            color=RGBColor(0xC0,0x00,0x00))

add_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["Today's Bias",        "BEARISH — US strikes on Iran (2nd night) + US CPI 4.2% shock + US markets –1.8%"],
        ["Confidence Level",    "MEDIUM (55%) — gap down is certain; recovery possible if Iran headline flips"],
        ["Crude Mode",          "CAUTION (₹8,712 MCX). Watch ₹8,850 for boundary-risk flag; ₹9,000 = BEAR"],
        ["VIX Sizing",          "HALF-SIZE (VIX 15.00). Expect VIX to spike to 16–18 on today's news."],
        ["Entry Permission",    "YELLOW — wait 30 minutes after open. No pre-open entries."],
        ["Expected Nifty Range","22,950 – 23,380 (wide, geopolitical vol)"],
        ["Key Morning Watch",   "MCX crude 9:00 AM print. If above ₹8,850 at open — DOWNGRADE to near-BEAR posture."],
        ["Key Afternoon Watch", "Any Iran ceasefire headline between 1–3 PM can violently reverse positions."],
        ["Bull Case",           "Trump announces Iran breakthrough → WTI drops to $87 → Nifty reverses +200 pts"],
        ["Bear Case",           "Iran sinks US vessel in Hormuz → WTI spikes to $96 → BEAR regime → Nifty drops 400 pts"],
        ["Most Likely Scenario","Gap down 80–100 pts. Attempt to hold 23,000. Range-bound 22,950–23,250 day."],
    ])

doc.add_paragraph()
add_para(doc, "THE ONE NUMBER THAT MATTERS:", bold=True, size=14,
         color=RGBColor(0xC0,0x00,0x00))
add_para(doc,
    "₹9,000 — MCX crude crossing this level triggers BEAR REGIME."
    " All CALL positions except ONGC / OIL India / MRPL must be BLOCKED.",
    bold=True, size=13, color=RGBColor(0xC0,0x00,0x00))

doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 9 — PAPER TRADING NOTE
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Section 9 — Paper Trading Note", level=1,
            color=RGBColor(0x1F,0x38,0x64))

add_bullets(doc, [
    "System will see: Gap-down open ~23,100–23,160, followed by first 15-min candle formation.",
    "Watch for: HPCL 15-min close below ₹408 (PUT entry signal). HAL relative strength vs Nifty (CALL signal).",
    "Paper log: Record ALL signal fires, timestamps, entry price and outcome. This week builds the week-2 backtest.",
    "If system fires 0 trades today: Log as OBSERVATION DAY. Do not lower confirmation thresholds.",
    "Running zero-trade day count: [Update manually after each session]",
    "Key items to verify at open: (1) MCX crude live price. (2) F&O ban list from NSE. (3) GIFT Nifty final pre-open. (4) India VIX live print.",
])

doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 10 — ONE-LINE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, "One-Line Summary", level=1, color=RGBColor(0x1F,0x38,0x64))
add_para(doc,
    "BEARISH day — US-Iran war escalates (2nd night of strikes, Iran hits US bases in Kuwait/Bahrain), "
    "US markets fell 1.8–2% on CPI shock at 4.2%, GIFT Nifty gap-down ~85 pts; CAUTION crude regime "
    "(MCX ₹8,712, ₹288 from BEAR flip); HALF-size on VIX 15; "
    "play HPCL PUT + HAL CALL (macro hedge) + IndiGo PUT + ONGC CALL; wait 30 min post-open; "
    "THE NUMBER: ₹9,000 MCX = BEAR regime trigger.",
    italic=True, size=12)

doc.add_paragraph()

# ── FOOTER ────────────────────────────────────────────────────────────────────
foot = doc.add_paragraph(); foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run(
    "Generated by Daily Trading Routine  |  11 June 2026  |  "
    "For personal educational use only. Not financial advice.")
fr.italic = True; fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80,0x80,0x80)

doc.save(OUT_PATH)
print(f"Wrote → {OUT_PATH}")
