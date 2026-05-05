"""Generate the daily trading-intelligence .docx brief for 2026-05-05."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-05"
BIAS = "BEAR"
ENTRY_PERMISSION = "GREEN (puts only + crude beneficiaries)"

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
sub_run = sub.add_run("Indian F&O Day Trader — Nifty + Stocks  |  Phase 2 Brief")
sub_run.italic = True
sub_run.font.size = Pt(12)

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
dr = date_p.add_run("Date: Tuesday, 5 May 2026  |  Generated 08:45 IST  |  Phase 1 context MISSING — rebuilt inline")
dr.bold = True
dr.font.size = Pt(11)
dr.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

# Bias + permission badges
badge_line(doc, "TODAY'S BIAS", "BEAR (HIGH confidence)", "C00000")
badge_line(doc, "ENTRY PERMISSION", "GREEN — puts + ONGC call only", "548235")
badge_line(doc, "CRUDE RULE MODE", "STRONG BEAR (>₹10,000) — RISING", "7F1D1D")
badge_line(doc, "VIX SIZING", "HALF SIZE (India VIX 18.30)", "BF9000")
badge_line(doc, "EXPIRY WARNING", "Tuesday = Weekly Nifty expiry day", "E69138")

doc.add_paragraph()

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
add_heading(doc, "Section 1 — Macro Snapshot ⚠️[REBUILT live]", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Reading", "Interpretation"],
    rows=[
        ["MCX Crude proxy (₹/bbl)", "~₹10,121 (WTI×USDINR)", "STRONG BEAR (>₹10,000) — no calls except crude beneficiaries"],
        ["Brent / WTI (live)", "$114.44 / $106.42", "+6% / +4% overnight; 4-yr high; peaks reported $125+"],
        ["Crude direction", "RISING aggressively", "Iran-UAE missile attack overnight; Hormuz partial closure"],
        ["GIFT Nifty", "24,010.5 (-199.5, -0.82%)", "Discount ~159 pts to Nifty close → BEAR gap"],
        ["Implied Nifty open", "~23,960", "Gap down -159 pts — BEAR signal (border of STRONG BEAR)"],
        ["Nifty close (May 4)", "24,119.30 (+121.75, +0.51%)", "Election-day rally; profit-booking risk today"],
        ["Sensex close (May 4)", "77,269.40 (+355.90, +0.46%)", "Above 77,000 support; vulnerable on gap"],
        ["India VIX", "18.30 (-0.87%)", "Elevated 15-20 band → MANDATORY HALF SIZE"],
        ["US S&P 500", "7,200.75 (-0.41%)", "Risk-off; pulled back from records on oil shock"],
        ["US Dow", "48,942 (-1.13%)", "Broader weakness — bearish global cue"],
        ["US Nasdaq", "25,067.80 (-0.19%)", "Mild weakness vs Dow — IT pocket relatively resilient"],
        ["US VIX (CBOE)", "18.24 (+7.36%)", "Fear rising sharply — confirms global risk-off"],
        ["Gold", "$4,579.60 (-1.40%)", "Down despite risk-off — unusual; USD strength dominates"],
        ["USD-INR", "~95.10 (record low)", "INR under stress: oil import bill + FII outflows"],
        ["FII cash (May 4)", "-₹4,539.49 Cr", "Heavy SELLERS — bearish flow signal"],
        ["DII cash (May 4)", "+₹4,493.73 Cr", "BUYERS absorbing; rare divergence — caution"],
    ])

# ADR check
doc.add_paragraph()
add_para(doc, "ADR Gap Indicators (May 4 US close)", bold=True, size=11)
add_bullets(doc, [
    "INFY ADR: DATA_UNAVAILABLE in current pull — Nasdaq -0.19% suggests mild softness",
    "WIT ADR: DATA_UNAVAILABLE — sector likely flat on rupee weakness offset",
    "IBN ADR: DATA_UNAVAILABLE — banking pocket likely under FII pressure",
    "HDB ADR: DATA_UNAVAILABLE — same risk signal as IBN",
    "Note: full ADR pull blocked by Phase 1 context absence; treat as risk-off bias",
])

# ---------- SECTION 2 — CRUDE RULE APPLIED ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule Applied (LIVE verified)", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_para(doc, "MODE: STRONG BEAR — RISING — NO CALLS except ONGC / Oil India / MRPL.", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_bullets(doc, [
    "LIVE Brent: $114.44 (+~6% overnight). LIVE WTI: $106.42 (+~4%). USD-INR: ~95.10.",
    "MCX proxy = $106.42 × 95.10 ≈ ₹10,121 — clearly inside STRONG BEAR (>₹10,000) band.",
    "Crude DIRECTION = RISING. Iran missile attack on UAE (Fujairah Oil Industry Zone fire); Hormuz partial closure.",
    "US Operation 'Project Freedom' escorting tankers; six Iranian boats sunk Monday — kinetic phase active.",
    "Oil reportedly touched $125 peak — four-year high; jet fuel shortage warnings in 1-2 months.",
    "FLIP_LEVEL: Brent below $100 + Hormuz reopening headline = downgrade to BEAR (₹9,000-10,000) → still puts only.",
    "FLIP_LEVEL deeper: Brent below $85 + ceasefire = upgrade to MILD BULL (₹7,500-8,500). Not on table today.",
    "AVOID SECTORS: Aviation, OMCs (BPCL/HPCL/IOC), Paints, Autos, Tyres, Lubricants, FMCG-with-packaging.",
    "EXCEPTION CALLS allowed: ONGC, Oil India, MRPL only. RIL refining margin tailwind but neutral net (index-weighted drag).",
])

# ---------- SECTION 3 — MARKET STRUCTURE & EXPIRY ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Market Structure & Expiry Mechanics", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Parameter", "Reading", "Notes"],
    rows=[
        ["Day", "Tuesday", "Under new NSE rule, weekly Nifty expiry on Tuesday → TODAY"],
        ["Days to weekly expiry", "0 (TODAY)", "EXTREME caution — theta crush, max-pain magnet"],
        ["Days to monthly expiry", "21 (May 26 Tue)", "All puts must use MAY MONTHLY series, not weekly"],
        ["Max Pain", "DATA_UNAVAILABLE", "Pull from NSE option chain at 9:00 AM"],
        ["Highest Call OI strike", "DATA_UNAVAILABLE", "Likely 24,200-24,300 zone basis prior session"],
        ["Highest Put OI strike", "DATA_UNAVAILABLE", "Likely 24,000 zone (psychological floor)"],
        ["PCR", "DATA_UNAVAILABLE", "Verify before entering — squeeze risk if PCR<0.8 today"],
        ["Nifty Futures basis", "DATA_UNAVAILABLE", "GIFT discount suggests futures discount to spot"],
        ["F&O Ban list", "Verify on NSE", "If any pick is on ban list → replace immediately"],
    ])

doc.add_paragraph()
add_para(doc, "Why expiry day matters TODAY:", bold=True, size=11)
add_bullets(doc, [
    "Theta on today's weekly options goes near-zero by 13:00 — NEVER buy weekly options for directional plays.",
    "Expiry-day pin to max pain is real — index can drift sideways even if direction is clear.",
    "If PCR is < 0.8 with bears positioned → SHORT SQUEEZE risk (consensus puts get unwound).",
    "Use MONTHLY expiry (May 26) for all directional bets today — 21 days of vega buffers a sudden ceasefire reversal.",
])

# ---------- SECTION 4 — GEOPOLITICS ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — Geopolitical Overlay (THE driver)", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_bullets(doc, [
    "Iran-Hormuz status: PARTIAL closure; US escorting ships (Operation Project Freedom). 2 American-flag tankers transited Mon.",
    "Overnight escalation: Iran missile/drone attack on UAE; UAE intercepted 19; Fujairah oil zone hit. 3 Indian nationals injured.",
    "US response: Sank 6 Iranian boats targeting civilian ships. Trump declined to confirm ceasefire status.",
    "OPEC+ symbolic output rise during closure — insufficient to offset.",
    "Crude tail risk today: ESCALATION-loaded. Any further attack on Saudi/UAE infra OR US strike = +5-10% crude spike.",
    "De-escalation tail: a sudden ceasefire headline = -10% crude crash → flips today's BEAR thesis violently. Must watch.",
])
add_para(doc, "If a ceasefire/escalation headline drops during 9:15-12:30, ALL positions reviewed immediately regardless of technicals.",
    bold=True, italic=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- SECTION 5 — STOCK SETUPS ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Five F&O Stock Setups", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc,
    "All puts use MAY MONTHLY expiry (26 May 2026) due to weekly expiry today. ONGC call uses monthly. "
    "Verify F&O ban list before entry. Apply HALF-SIZE on all picks (India VIX 18.30).",
    italic=True, size=10)

# STOCK 1 — ONGC CALL
add_para(doc, "STOCK 1 — ONGC (NSE: ONGC)  |  CALL  |  Grade A+", bold=True, size=13, color=RGBColor(0x00, 0x5A, 0x32))
stock_card(doc, [
    ("CMP (approx)", "₹275-290 zone (verify at 9:15)"),
    ("Lot size", "1,150 (verify at NSE)"),
    ("Crude Aligned", "YES — exception rule (upstream beneficiary, $1 Brent ≈ ₹6,180 cr ONGC EBITDA)"),
    ("Catalyst", "Brent at $114 (4-yr high) + Hormuz crisis = realisation surge; only segment of market with positive earnings tailwind"),
    ("Daily trend / 200 DMA", "UP / above (institutional accumulation in PSU upstream)"),
    ("Weekly trend", "UP — multi-month uptrend on geopolitical premium"),
    ("Key support", "₹265 (recent breakout zone)"),
    ("Key resistance", "₹305 (52w high zone)"),
    ("OI build-up", "Long build-up expected (verify); stock PCR likely >1.2"),
    ("Option Setup", "ATM CALL ₹280 strike — 26 May 2026 monthly expiry"),
    ("Theta warning", "NO (monthly = 21 days)"),
    ("Entry trigger (ALL must align)", "1) ST GREEN 15-min  2) RSI>50 15-min  3) StochRSI cross UP from <30  4) Volume>1.5×20MA  5) Brent holds >$110 + MCX>₹9,500"),
    ("Confirmation candle", "15-min close above ₹285"),
    ("Target T1 (40%)", "₹295 (≈ +3.5% on stock; call ≈ +35-50%)"),
    ("Target T2 (40%)", "₹305 (≈ +7% on stock; call ≈ +70-90%)"),
    ("Trail (20%)", "Trail with 15-min ST flip"),
    ("Stop loss", "Stock close <₹272 OR Brent breaks <$105 OR ST flips RED 15-min"),
    ("R:R", "≈ 2.5:1"),
    ("Time stop", "Exit by 13:00 IST"),
])

doc.add_paragraph()

# STOCK 2 — INDIGO PUT
add_para(doc, "STOCK 2 — InterGlobe Aviation / IndiGo (NSE: INDIGO)  |  PUT  |  Grade A+", bold=True, size=13, color=RGBColor(0xC0, 0x00, 0x00))
stock_card(doc, [
    ("CMP (approx)", "~₹4,260 (close May 4 zone — was 4,675 on April ceasefire spike, now retracing)"),
    ("Lot size", "15 (verify)"),
    ("Crude Aligned", "YES (puts) — every $10 Brent = ₹3,500-4,000 cr extra annual fuel bill"),
    ("Catalyst", "ATF crossed ₹2 lakh/KL in April (record); Brent $114 + Hormuz crisis = direct margin destruction; jet fuel shortage warnings 1-2 months"),
    ("Daily trend / 200 DMA", "DOWN / pressure (recent rally was ceasefire-driven, now reversing)"),
    ("Weekly trend", "DOWN — repeated 5-6% drops on each crude spike"),
    ("Key support", "₹4,100 (Apr swing low)"),
    ("Key resistance", "₹4,400 (50 DMA)"),
    ("OI build-up", "Short build-up expected as fuel cost shock plays through"),
    ("Option Setup", "Slightly OTM PUT ₹4,200 strike — 26 May 2026 monthly expiry"),
    ("Theta warning", "NO (monthly)"),
    ("Entry trigger (ALL must align)", "1) ST RED 15-min  2) RSI<50 15-min  3) StochRSI cross DOWN from >70  4) Volume>1.5×20MA  5) Brent stays >$110"),
    ("Confirmation candle", "15-min close below ₹4,230"),
    ("Target T1 (40%)", "₹4,150 (≈ -2.5% stock; put ≈ +40-60%)"),
    ("Target T2 (40%)", "₹4,050 (≈ -5% stock; put ≈ +90-130%)"),
    ("Trail (20%)", "Trail with 15-min ST flip"),
    ("Stop loss", "Stock close >₹4,330 OR Brent breaks <$100 OR ST flips GREEN 15-min"),
    ("R:R", "≈ 2.5:1"),
    ("Time stop", "Exit by 13:00 IST"),
])

doc.add_paragraph()

# STOCK 3 — ASIAN PAINTS PUT
add_para(doc, "STOCK 3 — Asian Paints (NSE: ASIANPAINT)  |  PUT  |  Grade A", bold=True, size=13, color=RGBColor(0xC0, 0x00, 0x00))
stock_card(doc, [
    ("CMP (approx)", "Verify at open"),
    ("Lot size", "200 (verify)"),
    ("Crude Aligned", "YES (puts) — 55-60% of input costs are crude derivatives (resins, solvents, binders)"),
    ("Catalyst", "Industry hiking prices 3-5% from 5 May (today!) — confirms cost pressure; price hikes never fully offset margin compression in rural-skewed demand"),
    ("Daily trend / 200 DMA", "DOWN / under pressure"),
    ("Weekly trend", "DOWN — paints sector down up to 5% on prior crude spikes"),
    ("Key support", "Verify last swing low at open"),
    ("Key resistance", "Prior day high"),
    ("OI build-up", "Short build-up expected (institutional crude-hedge trade)"),
    ("Option Setup", "ATM PUT — 26 May 2026 monthly expiry"),
    ("Theta warning", "NO (monthly)"),
    ("Entry trigger (ALL must align)", "1) ST RED 15-min  2) RSI<50 15-min  3) StochRSI cross DOWN from >70  4) Volume>1.5×20MA  5) Brent holds >$110"),
    ("Confirmation candle", "15-min close below prior-day low"),
    ("Target T1 (40%)", "≈ -2% on stock"),
    ("Target T2 (40%)", "≈ -4% on stock"),
    ("Trail (20%)", "Trail with 15-min ST flip"),
    ("Stop loss", "Stock close > prior-day high OR Brent breaks <$100"),
    ("R:R", "≈ 2:1"),
    ("Time stop", "Exit by 13:00 IST"),
])

doc.add_paragraph()

# STOCK 4 — BPCL PUT
add_para(doc, "STOCK 4 — BPCL (NSE: BPCL)  |  PUT  |  Grade A", bold=True, size=13, color=RGBColor(0xC0, 0x00, 0x00))
stock_card(doc, [
    ("CMP (approx)", "Verify at open"),
    ("Lot size", "1,800 (verify)"),
    ("Crude Aligned", "YES (puts) — OMC margin squeeze; cannot pass through cost at retail under govt pricing pressure"),
    ("Catalyst", "Brent $114 + INR 95 = dual hit on import bill; OMCs historically lose ₹3-4 per litre of marketing margin per $5 Brent; downstream-only exposure (no upstream offset like RIL/IOC integrated)"),
    ("Daily trend / 200 DMA", "DOWN / vulnerable"),
    ("Weekly trend", "DOWN — paint and OMC downside cluster"),
    ("Key support", "Verify last swing low"),
    ("Key resistance", "Prior day high"),
    ("OI build-up", "Short build-up expected"),
    ("Option Setup", "ATM PUT — 26 May 2026 monthly expiry"),
    ("Theta warning", "NO (monthly)"),
    ("Entry trigger (ALL must align)", "Standard 5-condition entry checklist + Brent > $110 confirmation"),
    ("Confirmation candle", "15-min close below prior-day low"),
    ("Target T1 (40%)", "≈ -2.5% on stock"),
    ("Target T2 (40%)", "≈ -5% on stock"),
    ("Trail (20%)", "Trail with 15-min ST flip"),
    ("Stop loss", "Stock close > prior-day high OR Brent breaks <$100"),
    ("R:R", "≈ 2:1"),
    ("Time stop", "Exit by 13:00 IST"),
])

doc.add_paragraph()

# STOCK 5 — MARUTI PUT
add_para(doc, "STOCK 5 — Maruti Suzuki (NSE: MARUTI)  |  PUT  |  Grade B", bold=True, size=13, color=RGBColor(0xE6, 0x91, 0x38))
stock_card(doc, [
    ("CMP (approx)", "₹13,314 (close May 4)"),
    ("Lot size", "50 (verify)"),
    ("Crude Aligned", "YES (puts) — fuel-cost shock crimps demand + commodity input pressure on steel/plastic"),
    ("Catalyst", "Stock rallied 2.8% May 4 on April sales — primed for mean reversion on crude shock; rupee at record-low 95 hits import-content cost"),
    ("Daily trend / 200 DMA", "UP / above (this is a counter-trend reversal play, not trend continuation — that's why B grade)"),
    ("Weekly trend", "UP (still strong) — wait for confirmation candle, do NOT pre-position"),
    ("Key support", "₹13,100 (recent base)"),
    ("Key resistance", "₹13,500 (yesterday high)"),
    ("OI build-up", "Long unwinding likely on crude spike"),
    ("Option Setup", "Slightly OTM PUT ₹13,200 strike — 26 May 2026 monthly expiry"),
    ("Theta warning", "NO (monthly)"),
    ("Entry trigger (ALL must align)", "1) ST RED 15-min  2) RSI<50  3) StochRSI cross DOWN from >70  4) Volume>1.5×20MA  5) Brent holds >$110 + Nifty trading below 24,000"),
    ("Confirmation candle", "15-min close below ₹13,180 — STRICT (counter-trend needs strong confirmation)"),
    ("Target T1 (40%)", "₹13,000 (≈ -1.5% stock)"),
    ("Target T2 (40%)", "₹12,800 (≈ -3% stock)"),
    ("Trail (20%)", "Trail with 15-min ST flip"),
    ("Stop loss", "Stock close >₹13,350 (TIGHT — counter-trend) OR Brent breaks <$100"),
    ("R:R", "≈ 2:1"),
    ("Time stop", "Exit by 13:00 IST"),
])

# ---------- SECTION 6 — CONTRARIAN CHECK ----------
doc.add_paragraph()
add_heading(doc, "Section 6 — Contrarian Check (Layer 3)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc, "Q1. Consensus Narrative", bold=True, size=11)
add_para(doc,
    "Every retail trader, TV anchor, and Twitter analyst expects a gap-down BEAR open today driven by Iran-UAE escalation, "
    "$114 Brent, and FII selling. The obvious trade is short Nifty / IndiGo puts / Asian Paints puts. The whole street is on one side.",
    size=10)

add_para(doc, "Q2. The Trap", bold=True, size=11)
add_para(doc,
    "If a sudden Hormuz reopening or US-Iran ceasefire headline drops during the session, crude can crash 8-12% intraday and "
    "all bear trades blow up simultaneously. Retail puts get crushed by vega collapse + spot reversal — a textbook short squeeze "
    "on PUT sellers (note: this means PUT BUYERS lose, CALL writers lose, IndiGo can hit upper circuit again as it did April 8).",
    size=10)

add_para(doc, "Q3. Smart Money vs Retail Divergence", bold=True, size=11)
add_para(doc,
    "FII derivative positioning: DATA_UNAVAILABLE — must check NSE/NSDL. FII cash was net SELL ₹4,539 cr; if FII derivatives "
    "are net LONG index futures, that's a classic divergence (selling cash, hedged long futures = trap loading for shorts). "
    "DII heavy buying (₹4,494 cr) vs FII selling = institutional disagreement. Verify before sizing up puts.",
    size=10)

add_para(doc, "Q4. The Flip Trigger", bold=True, size=11)
add_para(doc,
    "ANY of these single events between 9:15-12:30 flip BEAR to BULL: (a) Hormuz reopening confirmed by US Navy statement, "
    "(b) Iran ceasefire announcement / Trump tweet, (c) Brent break below $100, (d) Nifty reclaims 24,200 with broad strength. "
    "Specific level to watch: NIFTY 24,000. A clean break below = bear thesis confirmed; a defended bounce from 23,900 = trap.",
    size=10)

# ---------- SECTION 7 — CONTRARIAN OVERRIDE ----------
doc.add_paragraph()
add_heading(doc, "Section 7 — Contrarian Override Check ⚠️ MANDATORY", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Condition", "Status", "Effect"],
    rows=[
        ["A. GIFT Nifty gap up > +100 pts?", "NO (gap DOWN -159)", "No bullish override"],
        ["B. Crude falling > 1% overnight?", "NO (rising +5-6%)", "No bullish override"],
        ["C. Contrarian flip identified > 40% prob?", "PARTIAL — sudden ceasefire risk real but not >40%", "Watchful, not overriding"],
        ["D. Expiry week + PCR < 0.8?", "Today IS expiry; PCR DATA_UNAVAILABLE", "Squeeze risk POSSIBLE — verify PCR before entry"],
    ])

add_para(doc, "VERDICT:", bold=True, size=12)
add_bullets(doc, [
    "A+B both NO → no bullish override on BEAR bias",
    "C alone (without D fully confirmed) → keep BEAR but flag ceasefire-headline tail risk",
    "If D confirms (PCR <0.8 verified at 9:00 AM) → SIZE DROPS TO QUARTER on all puts; ONGC call unaffected",
    "BEAR bias confirmed. HIGH confidence with squeeze-risk caveat on puts.",
])

# ---------- SECTION 8 — TIME-STAMPED CALENDAR ----------
doc.add_paragraph()
add_heading(doc, "Section 8 — Time-Stamped Event Calendar", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Time (IST)", "Event", "Action"],
    rows=[
        ["08:45", "Brief delivered, trader review window", "Read brief + verify F&O ban list, PCR, max pain on NSE"],
        ["09:00", "Pre-open auction begins", "Confirm GIFT Nifty implied open; check ADR moves"],
        ["09:15", "Market open", "DO NOT enter at open — wait for first 15-min candle close"],
        ["09:30", "First 15-min candle close", "Earliest entry permitted with full 5-condition checklist"],
        ["12:30", "LAST ENTRY PERMITTED (HARD RULE)", "No new entries after this time"],
        ["13:00", "Time stop on all positions", "Exit weakening positions; tighten stops on winners"],
        ["13:30", "HARD EXIT (no exceptions)", "All F&O positions flat regardless of P&L"],
        ["15:30", "Weekly expiry settlement", "Today's weekly Nifty options settle; do not hold any"],
        ["Anytime", "Iran/Hormuz/ceasefire headline", "Immediate review of all positions"],
    ])

# ---------- SECTION 9 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 9 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["TODAY'S BIAS", "BEAR"],
        ["CONFIDENCE", "HIGH"],
        ["CRUDE MODE", "STRONG BEAR (>₹10,000 proxy) — RISING"],
        ["VIX SIZING", "HALF SIZE (India VIX 18.30)"],
        ["ENTRY PERMISSION", "GREEN — puts + ONGC call only; RED for index/banking calls"],
        ["Nifty Support S1", "23,900"],
        ["Nifty Support S2", "23,700"],
        ["Nifty Resistance R1", "24,200"],
        ["Nifty Resistance R2", "24,300 (key sell-on-rise zone)"],
        ["Critical Level", "24,000 — clean break = bear confirmation; defended = trap"],
    ])

doc.add_paragraph()
add_para(doc, "THE BULL CASE (write even though bias = BEAR):", bold=True, size=11, color=RGBColor(0x00, 0x5A, 0x32))
add_para(doc,
    "If a sudden Hormuz reopening or ceasefire headline drops, crude crashes 10%+ and the entire risk-off trade unwinds. "
    "DII buying force (₹4,494 cr May 4) signals domestic floor; election results day rally May 4 shows underlying domestic optimism intact.",
    size=10)

add_para(doc, "THE BEAR CASE (today's primary thesis):", bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "Brent at 4-yr high $114, MCX proxy ₹10,121, INR record low 95.10, FII heavy seller ₹4,539 cr, GIFT discount 159 pts, "
    "active kinetic Iran-US conflict, jet fuel shortage warnings — five aligned bearish signals dominate the tape.",
    size=10)

add_para(doc, "FLIP TRIGGER:", bold=True, size=11)
add_para(doc,
    "If Brent breaks $100 OR Nifty reclaims 24,200 with volume OR US/Iran ceasefire headline → bias flips to CAUTIOUS BULL "
    "(close all puts, switch to ONGC PUT for crude collapse trade, watch IndiGo for upper-circuit replay).",
    italic=True, size=10)

doc.add_paragraph()
add_para(doc, "TOP 3 STOCK PICKS RANKED", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "#1 ONGC CALL — A+ — only crude beneficiary call permitted; entry above ₹285 monthly call",
    "#2 INDIGO PUT — A+ — direct fuel-cost destruction, ATF at record; entry below ₹4,230 monthly put",
    "#3 ASIAN PAINTS PUT — A — 55-60% input costs crude-linked, sector hiking prices today; entry below prior-day low",
])

doc.add_paragraph()
add_para(doc, "THE ONE RISK THAT RUINS EVERYTHING TODAY:", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "A surprise Hormuz reopening or US-Iran ceasefire headline during 9:15-12:30. Crude crashes 8-12% intraday, "
    "all puts blow up simultaneously, IndiGo upper-circuits like April 8, and ONGC call also collapses. "
    "Mitigation: HALF-SIZE (mandatory anyway), MONTHLY expiries (vega buffer), strict 13:00 time stop.",
    size=10)

# ---------- ONE-LINE SUMMARY ----------
doc.add_paragraph()
add_heading(doc, "ONE-LINE SUMMARY (read at 9:10 AM)", level=1, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Today is BEAR because Brent at 4-yr high $114 with active Iran-UAE kinetic crisis. Crude at ₹10,121 = STRONG BEAR, "
    "RISING. Watch ONGC call and IndiGo put. Key risk: surprise ceasefire headline. Size HALF. "
    "Flips if Brent breaks $100 or Nifty reclaims 24,200.\"",
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
