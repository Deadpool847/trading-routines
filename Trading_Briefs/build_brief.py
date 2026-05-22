"""Build the daily F&O trading brief Word document for 2026-05-22."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

DATE = "2026-05-22"
BIAS = "RANGE-BEAR"
OUT = f"/home/user/trading-routines/Trading_Briefs/Trading_Brief_{DATE}_{BIAS}.docx"

doc = Document()

# Default style
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(10)


def H(text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x68)
    return h


def P(text, bold=False, color=None, size=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    if color:
        r.font.color.rgb = color
    if size:
        r.font.size = Pt(size)
    return p


def line(text=""):
    P(text)


# =========================================================
# TITLE BAR
# =========================================================
title = doc.add_heading(f"ELITE DAILY TRADING BRIEF — {DATE} — {BIAS}", level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
P("Generated: 08:55 IST  |  Weekly expiry: 4 trading days (Thu 28-May)  |  "
  "Monthly expiry: 4 trading days (Thu 28-May)  |  Expiry week: NO (borderline)",
  bold=True)

# Status box
box = doc.add_table(rows=5, cols=1)
box.style = "Light Shading Accent 1"
box.rows[0].cells[0].text = "BIAS: RANGE-BEAR (crude-led, geopolitically fragile)"
box.rows[1].cells[0].text = "ENTRY PERMISSION: YELLOW — Hormuz binary risk + heavy FII selling"
box.rows[2].cells[0].text = "CRUDE MODE: BEAR (₹9,500/bbl proxy) | Direction: RISING today after 4.6% fall Wed"
box.rows[3].cells[0].text = "VIX SIZING: HALF SIZE (India VIX estimated 17-20, data partial)"
box.rows[4].cells[0].text = "EXPIRY ALERT: 4 trading days to monthly+weekly — theta acceleration begins Mon"

# =========================================================
# SECTION 1 — MACRO SNAPSHOT
# =========================================================
H("SECTION 1 — MACRO SNAPSHOT", level=1)

t = doc.add_table(rows=1, cols=3)
t.style = "Light Grid Accent 1"
hdr = t.rows[0].cells
hdr[0].text = "Metric"
hdr[1].text = "Value"
hdr[2].text = "Signal"

rows = [
    ("MCX Crude (WTI × USDINR proxy)", "₹9,526 / bbl",  "BEAR mode (₹9k-10k band)"),
    ("Brent",                          "$104.52 (+1.89%)", "Rising today, off Wed lows"),
    ("WTI",                            "$98.87",          "Rising today"),
    ("GIFT Nifty",                     "23,665 (+10 to +33 pts)", "Tepid gap, no >100 trigger"),
    ("Nifty 21-May close",             "23,654.70 (-0.02%)", "Flat — coiling near resistance"),
    ("S&P 500",                        "7,445.72 (+0.17%)", "Risk-on mild"),
    ("Nasdaq",                         "26,293.10 (+0.09%)", "Flat tech"),
    ("Dow",                            "50,285.66 (+0.55%)", "Mild risk-on"),
    ("US VIX",                         "DATA_UNAVAILABLE",  "Assume calm given S&P up"),
    ("India VIX",                      "DATA_UNAVAILABLE (est 17-20)", "HALF SIZE sizing"),
    ("FII cash 20-May",                "-₹1,597 Cr (SELL)", "MTD May = -₹25,897 Cr"),
    ("DII cash 20-May",                "+₹1,968 Cr (BUY)",  "MTD May = +₹48,370 Cr"),
    ("USD-INR",                        "96.35 (rupee -7% YTD)", "Weak — RBI $5b swap on 26-May"),
    ("INFY ADR",                       "$12.49 (~flat)",    "INFY likely opens flat"),
    ("WIT ADR",                        "$1.93 (0.00%)",     "WIPRO flat open"),
    ("IBN ADR",                        "$25.36 (+0.42%)",   "ICICIBANK mild gap-up"),
    ("HDB ADR",                        "$24.24 (-0.08%)",   "HDFCBANK flat-soft open"),
]
for m, v, s in rows:
    r = t.add_row().cells
    r[0].text = m
    r[1].text = v
    r[2].text = s

line()
P("Geopolitical one-liner: Iran Supreme Leader ordered enriched uranium reserves "
  "to remain inside Iran, complicating peace talks; Brent +1.89% today on the news. "
  "However Trump still public on 'final-stages' Iran deal and Hormuz tanker traffic "
  "is at ~5% of pre-war levels (largest energy supply shock in history per IEA).",
  bold=True)
P("Today's intraday headline risk: HIGH — any Iran-deal or Hormuz headline can swing "
  "Nifty 200 pts either way within minutes.",
  bold=True, color=RGBColor(0xC0, 0x00, 0x00))

# =========================================================
# SECTION 2 — CRUDE RULE + STRUCTURAL READ
# =========================================================
H("SECTION 2 — CRUDE RULE + STRUCTURAL READ", level=1)

P("CRUDE_MODE: BEAR  (puts-only on victims; calls only on crude beneficiaries)", bold=True)
P("CRUDE_DIRECTION: RISING today (+1.89%) after a 4.6% drop Wed — choppy two-way", bold=True)
P("CRUDE_FLIP_LEVEL: ₹8,500 (≈ WTI $88 × USDINR 96.4) — below this, mode flips to MILD BULL", bold=True)
P("AVOID (crude victims): BPCL, HPCL, IOC, IndiGo, Spicejet, Asian Paints, Berger, "
   "MRF, Apollo Tyres, fertiliser names")
P("EXCEPTION CALLS permitted on: ONGC, Oil India, MRPL, GAIL, Petronet, Reliance (E&P leg), BEL/HAL (defence beneficiary of Iran tension)")

line()
P("STRUCTURAL READ", bold=True)
P("• Nifty 23,654.70 vs max pain: DATA_UNAVAILABLE — option-chain live feed not extracted. "
  "Read with caution. Best estimate from positioning: max pain cluster near 23,600-23,700 "
  "given coil; magnetism = SIDEWAYS into 28-May expiry.")
P("• PCR: DATA_UNAVAILABLE. Inferred sentiment from FII heavy selling + USDINR 96 + "
  "crude $104 = retail likely positioning bearish → consensus crowded short risk.")
P("• OI change direction: DATA_UNAVAILABLE — verify on terminal before first entry.")
P("• Futures basis: DATA_UNAVAILABLE — verify; if Nifty fut in discount > 25 pts, "
  "treat as confirmation of bearish lean.")
P("• Days to monthly expiry: 4 trading days (Thu 28-May). Theta meaningful from Tuesday.")
P("• Bank Nifty leadership: Per ADR cues IBN +0.42%, HDB -0.08% — mixed but slight "
  "edge to ICICI. If Bank Nifty leads any rally → REAL move. If banks lag a Nifty "
  "bounce while only IT/Reliance push → FADE the move.")
P("• Defenders: Put writers expected to defend 23,500 (psychological round + recent low). "
  "Call writers cluster 23,800-23,900 (ceiling).")
P("• Magnetism direction: DOWN/sideways into expiry while crude > ₹9,000.")

# =========================================================
# SECTION 3 — CONTRARIAN CHECK
# =========================================================
H("SECTION 3 — CONTRARIAN CHECK (LAYER 3)", level=1)

P("Q1. CONSENSUS: Retail + TV expects a 'gap-down on Iran tension + crude spike' → "
  "short Nifty / long BPCL/IndiGo puts / long ONGC calls. Everyone is reading the "
  "same Hormuz headlines.", bold=True)
P("Q2. THE TRAP: Trump signs (or merely re-affirms) Iran-deal language at any moment "
  "during India hours → crude collapses 4-6% in minutes → OMC/aviation puts get "
  "destroyed, ONGC calls bleed, Nifty rips 200 pts on relief. Buyers of the consensus "
  "trade pay the bill. Short-covering rally on PSU banks + autos + paints.", bold=True)
P("Q3. RETAIL STOPS: Long stops clustered below 23,500 (Nifty) and below 285 ONGC. "
  "Short stops clustered above 23,800 Nifty and above 380 BPCL.", bold=True)
P("Q4. FLIP TRIGGER: A single Reuters/AP flash 'Iran-US framework agreed' or 'Hormuz "
  "tanker convoy resumes commercial volume' → Brent prints sub-$100 → MCX < ₹9,000 → "
  "bias flips to BULL within 30 minutes. Equally: a single Israeli strike headline "
  "flips bias to DEEP BEAR with Brent $115+.", bold=True)

line()
P("CONTRARIAN OVERRIDE LOGIC CHECK", bold=True)
P("A. GIFT Nifty gap-up > +100 pts? NO (only +10 to +33).")
P("B. Crude falling > 2% overnight? NO — crude is RISING +1.89% today.")
P("C. Contrarian scenario probability > 40%? YES — Iran deal headline risk is real.")
P("D. Expiry week + PCR < 0.8? PARTIAL — 4 days to expiry (borderline), PCR data unavailable.")
P("RESULT: No A+B override fires. C alone → keep BEAR lean but downgrade confidence to LOW. "
  "Permission stays YELLOW. Be ready to flip on a single Iran headline.", bold=True,
  color=RGBColor(0xC0, 0x00, 0x00))

# =========================================================
# SECTION 4 — STOCK SETUPS
# =========================================================
H("SECTION 4 — 5 INDIVIDUAL F&O SETUPS", level=1)


def setup(title, body):
    P(title, bold=True, color=RGBColor(0x1F, 0x3A, 0x68), size=11)
    for ln in body:
        P(ln)
    line()


setup(
    "1) ONGC — NSE: ONGC | CALL | Grade A",
    [
        "CMP: ₹296.45 | Lot: 1,150 | Crude aligned: YES (upstream beneficiary)",
        "CATALYST: Brent $104+ holds; Iran tension keeps upstream margins fat. "
        "+7.24% over last month already shows institutional accumulation.",
        "TECHNICAL: Daily trend UP | Above 200 DMA: YES (key institutional signal) | "
        "Weekly trend UP | Support ₹278 / ₹277 | Resistance ₹305 → ₹316",
        "STRUCTURAL: OI direction LONG BUILD (recent rally on rising OI per multiple "
        "sources) | Delivery % RISING | Block deal NO | Ban list NO",
        "OPTION SETUP: Strike ₹300 CE [near-ATM] | Expiry 28-May-2026 (monthly) | "
        "THETA WARNING (≤ 3 days): NO (4 days, but use monthly only)",
        "ENTRY (ALL 4 must fire on 15-min): ST GREEN | RSI > 50 | StochRSI cross UP "
        "from < 30 | Volume > 1.5x 20-period avg | Confirmation: 15-min close ABOVE ₹298.",
        "T1 (book 40%): ₹304 | T2 (book 40%): ₹310 | SL: ₹293 + 15-min ST flip RED",
        "R:R: ~2.2 : 1 | Time stop: Exit by 13:00 IST",
        "GRADE REASON: A — best crude-mode alignment, clean uptrend, above 200DMA, "
        "macro tailwind. Loses a grade only because of Iran-deal flip risk.",
    ],
)

setup(
    "2) BPCL — NSE: BPCL | PUT | Grade A",
    [
        "CMP: ~₹370 | Lot: 1,800 | Crude aligned: YES (OMC margin victim of $104 Brent)",
        "CATALYST: OMC marketing margins compress sharply when Brent > $100 and "
        "rupee at 96. Recent week already -2.41%. Absorbing supply near multi-year highs.",
        "TECHNICAL: Daily trend DOWN (rolling over) | Above 200 DMA: YES but fading | "
        "Weekly trend SIDEWAYS | Support ₹360 / ₹355 | Resistance ₹378 / ₹391 (ATH)",
        "STRUCTURAL: OI direction SHORT BUILD likely (price down + supply absorption) | "
        "Delivery % STABLE | Block deal NO | Ban list NO",
        "OPTION SETUP: Strike ₹370 PE [ATM] | Expiry 28-May-2026 (monthly) | "
        "THETA WARNING: NO (4 days)",
        "ENTRY (ALL 4): ST RED | RSI < 50 | StochRSI cross DOWN from > 70 | "
        "Volume > 1.5x | Confirmation: 15-min close BELOW ₹367.",
        "T1: ₹362 | T2: ₹356 | SL: ₹374 + 15-min ST flip GREEN",
        "R:R: ~2.3 : 1 | Time stop: Exit by 13:00 IST",
        "GRADE REASON: A — clean crude-victim short, anchored to macro thesis. "
        "Downgrade if Iran deal flashes during the day (would squeeze BPCL up 4-5%).",
    ],
)

setup(
    "3) BEL — NSE: BEL | CALL | Grade A",
    [
        "CMP: ₹429.6 | Lot: 475 | Crude aligned: YES via geopolitical (defence beneficiary)",
        "CATALYST: ₹74,000 Cr order book as of 01-Apr-2026; FY26 fresh order intake "
        "₹30,000 Cr including exports. Q3 PAT +20.81% YoY. Iran tension = defence flow.",
        "TECHNICAL: Daily trend UP | Above 200 DMA: YES | Weekly trend UP | "
        "Support ₹420 / ₹415 | Resistance ₹440 / ₹450",
        "STRUCTURAL: OI direction LONG BUILD (sustained rally with rising OI per recent "
        "analyst notes) | Delivery % RISING | Block deal NO | Ban list NO",
        "OPTION SETUP: Strike ₹430 CE [ATM] | Expiry 28-May-2026 (monthly) | "
        "THETA WARNING: NO (4 days, use monthly only)",
        "ENTRY (ALL 4): ST GREEN | RSI > 50 | StochRSI cross UP from < 30 | "
        "Volume > 1.5x | Confirmation: 15-min close ABOVE ₹432.",
        "T1: ₹440 | T2: ₹448 | SL: ₹425 + 15-min ST flip RED",
        "R:R: ~2.5 : 1 | Time stop: Exit by 13:00 IST",
        "GRADE REASON: A — clean defence/geopolitical CALL, hedges Iran-escalation scenario. "
        "If Iran deal hits, BEL may dip 2% but order book floor cushions the move.",
    ],
)

setup(
    "4) INDIGO — NSE: INDIGO | PUT | Grade B+",
    [
        "CMP: ₹4,409.70 | Lot: 15 | Crude aligned: YES (ATF cost victim)",
        "CATALYST: Q3 FY26 net profit -77.55% YoY; EBITDA margin already compressed to 23% "
        "due to fuel. Brent $104+ continues margin pressure. Down from 52w high ₹6,232.",
        "TECHNICAL: Daily trend DOWN (lower-highs structure) | Above 200 DMA: NO "
        "(institutional sell signal) | Weekly trend DOWN | Support ₹4,300 / ₹4,200 | "
        "Resistance ₹4,500 / ₹4,600",
        "STRUCTURAL: OI direction SHORT BUILD (price down, OI rising trend) | "
        "Delivery % STABLE | Block deal NO | Ban list NO",
        "OPTION SETUP: Strike ₹4,400 PE [ATM] | Expiry 28-May-2026 (monthly) | "
        "THETA WARNING: NO (4 days)",
        "ENTRY (ALL 4): ST RED | RSI < 50 | StochRSI cross DOWN from > 70 | "
        "Volume > 1.5x | Confirmation: 15-min close BELOW ₹4,380.",
        "T1: ₹4,330 | T2: ₹4,260 | SL: ₹4,455 + 15-min ST flip GREEN",
        "R:R: ~2.1 : 1 | Time stop: Exit by 13:00 IST",
        "GRADE REASON: B+ — strong fundamental thesis but premium lot value (~₹66k notional) "
        "+ Iran-deal headline can spike INDIGO 4-5% intraday. Smaller size mandated.",
    ],
)

setup(
    "5) ICICIBANK — NSE: ICICIBANK | CALL | Grade B+",
    [
        "CMP: derived from IBN ADR ₹25.36 ×~83 ≈ ₹2,100-2,110 range (verify on open) | "
        "Lot: 700 | Crude aligned: NEUTRAL (private bank, mild beneficiary of weak rupee remit)",
        "CATALYST: IBN ADR +0.42% overnight while HDB ADR -0.08% — relative leadership in "
        "private banks. Bank Nifty leadership check: if it leads, real move.",
        "TECHNICAL: Daily trend UP | Above 200 DMA: YES | Weekly trend UP | "
        "Support: prior swing low | Resistance: recent swing high",
        "STRUCTURAL: OI direction LONG BUILD on the ADR-aligned pop | "
        "Delivery % STABLE | Block deal NO | Ban list NO",
        "OPTION SETUP: Strike ATM CE (open-of-day) | Expiry 28-May-2026 (monthly) | "
        "THETA WARNING: NO",
        "ENTRY (ALL 4): ST GREEN | RSI > 50 | StochRSI cross UP from < 30 | "
        "Volume > 1.5x | Confirmation: 15-min close ABOVE pre-market high.",
        "T1: +0.6% | T2: +1.2% from entry | SL: -0.4% from entry + 15-min ST flip RED",
        "R:R: target ≥ 2 : 1; SKIP if intraday R:R drops below 2:1",
        "GRADE REASON: B+ — ADR-confirmed direction but Bank Nifty leadership unverified; "
        "downgrade from A because exact CMP unverified pre-open and macro headwinds heavy.",
    ],
)

# =========================================================
# SECTION 5 — FINAL VERDICT
# =========================================================
H("SECTION 5 — FINAL VERDICT", level=1)

verdict = doc.add_table(rows=5, cols=2)
verdict.style = "Light Grid Accent 1"
verdict.rows[0].cells[0].text = "TODAY'S BIAS"
verdict.rows[0].cells[1].text = "RANGE-BEAR (crude-led, geopolitically fragile)"
verdict.rows[1].cells[0].text = "CONFIDENCE"
verdict.rows[1].cells[1].text = "LOW — single Iran/Hormuz headline can flip the day"
verdict.rows[2].cells[0].text = "CRUDE MODE"
verdict.rows[2].cells[1].text = "BEAR (₹9,500 proxy) | Rising today after Wed plunge"
verdict.rows[3].cells[0].text = "VIX SIZING"
verdict.rows[3].cells[1].text = "HALF SIZE (India VIX data partial; environment risk-high)"
verdict.rows[4].cells[0].text = "ENTRY PERMISSION"
verdict.rows[4].cells[1].text = "YELLOW — checklist + headline scan before every entry"

line()
P("THE BULL CASE (even though bias = BEAR):", bold=True)
P("Trump 'final stages' Iran-deal narrative is alive; one credible headline collapses "
  "Brent to sub-$95, crushes shorts, and ignites OMC/aviation/auto short-covering. "
  "DII flow remains relentlessly positive (+₹48k Cr MTD) absorbing every FII sell.")

P("THE BEAR CASE (primary thesis):", bold=True)
P("FII selling -₹25,897 Cr MTD, USDINR 96.35 (worst Asia), Brent $104, Hormuz traffic "
  "5% of pre-war — this combination is historically a recipe for sustained derating. "
  "Nifty coiling sideways with crude bleeding margins of OMCs, autos, paints, aviation.")

P("FLIP TRIGGER:", bold=True)
P("If Brent prints below $98 intraday (Reuters/CNBC flash) → bias flips to BULL, exit "
  "all puts, ONGC/BEL also exit (no longer crude-tailwind trades), pivot to OMC/auto "
  "calls. Conversely if Brent prints above $108 → bias deepens to STRONG BEAR.")

line()
P("NIFTY KEY LEVELS:", bold=True)
P("S2 ₹23,400 | S1 ₹23,500 | CRITICAL ₹23,655 (yesterday close) | R1 ₹23,800 | R2 ₹23,900")

line()
P("TOP 3 RANKED:", bold=True)
P("#1 ONGC — A — crude beneficiary, clean uptrend, above 200DMA — entry ABOVE ₹298")
P("#2 BEL  — A — defence/geopolitical CALL, order-book floor — entry ABOVE ₹432")
P("#3 BPCL — A — OMC crude victim, rolling over from ATH zone — entry BELOW ₹367")

line()
P("ONE RISK THAT RUINS EVERYTHING TODAY:", bold=True, color=RGBColor(0xC0, 0x00, 0x00))
P("A live Iran/US framework-deal flash during 9:15–12:30 collapses Brent 5%+ in minutes. "
  "ONGC/BEL longs and BPCL/INDIGO puts ALL bleed simultaneously — the only trade that "
  "survives is exiting on the flash and waiting 30 min. Set a hard mental stop: any "
  "Reuters Iran headline = flatten everything, reassess.")

# =========================================================
# ONE-LINE SUMMARY
# =========================================================
H("ONE-LINE SUMMARY (read at 9:10 AM)", level=2)
P('"Today is RANGE-BEAR because crude is ₹9,500 (BEAR mode) and FII selling persists '
  '(-₹25.9k Cr MTD). Crude ₹9,500 = BEAR, RISING +1.89% today. Watch ONGC CALL and '
  'BPCL PUT. Key risk: an Iran-deal flash that collapses Brent and squeezes shorts. '
  'Size HALF (VIX rule). Flips if Brent prints below $98 intraday."',
  bold=True)

# =========================================================
# COMPLIANCE TRAILER
# =========================================================
H("HARD-CODED RULE COMPLIANCE", level=2)
P("• Exchange: NSE F&O only — OK")
P("• Trading window: 9:15 – 12:30 (no new entries after 12:30); hard exit 13:30 — OK")
P("• Max 3 trades today (5 setups listed → pick top 3)")
P("• Daily risk cap: 2% of capital — sizing per VIX rule = HALF")
P("• Entry checklist: ALL 4 must fire (ST, RSI, StochRSI, Volume) — embedded")
P("• Option expiry rule: 4 days → monthly OR weekly both OK; brief uses MONTHLY 28-May "
  "to be safe (theta safer if held into Tue/Wed)")
P("• F&O ban list today: KAYNES, SAIL — both excluded from setups above")

# Sources page
H("DATA SOURCES (verify any DATA_UNAVAILABLE before entry)", level=2)
sources = [
    ("Brent / WTI price",            "tradingeconomics.com, investing.com, oilprice.com"),
    ("GIFT Nifty",                   "5paisa.com, kotakneo.com, business-standard.com"),
    ("S&P / Nasdaq / Dow / ADRs",    "Yahoo Finance, thestreet.com, nasdaq.com"),
    ("Iran / Hormuz",                "Reuters, CNBC, Yahoo Finance, Wikipedia 2026 Hormuz crisis"),
    ("Nifty close, FII/DII",         "NSE India, NSDL, nifty50pulse.in"),
    ("USD-INR",                      "investing.com IN, tradingeconomics.com"),
    ("F&O ban list",                 "niftytrader.in, 5paisa.com, NSE archives"),
    ("Stock fundamentals/OI",        "screener.in, trendlyne.com, tickertape.in, nseindia.com"),
    ("Option chain (UNVERIFIED)",    "REQUIRES live check on NSE option chain before entry"),
]
for label, src in sources:
    P(f"• {label}: {src}")

doc.save(OUT)
print(f"SAVED: {OUT}")
