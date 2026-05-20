#!/usr/bin/env python3
"""Generate the daily trading brief Word document from in-memory content."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pathlib import Path

DATE = "2026-05-20"
BIAS = "RANGE-BEAR"
OUT_DIR = Path("/home/user/trading-routines/Trading_Briefs")
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PATH = OUT_DIR / f"Trading_Brief_{DATE}_{BIAS}.docx"

doc = Document()

style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10)

def H(text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.color.rgb = RGBColor(0, 0, 0)
    return h

def P(text, bold=False, size=10):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    return p

def BOX(lines):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.1)
    for i, line in enumerate(lines):
        r = p.add_run(line + ("\n" if i < len(lines)-1 else ""))
        r.bold = True
        r.font.size = Pt(10)

# ============ TITLE ============
title = doc.add_heading(f"ELITE DAILY TRADING BRIEF — {DATE} — {BIAS}", level=0)
for r in title.runs:
    r.font.color.rgb = RGBColor(0, 0, 0)

P("Generated: 03:30 IST  |  Weekly expiry: 1 day  |  Monthly expiry: 6 days  |  Expiry week: YES",
  bold=True)

# ============ BIAS BOX ============
BOX([
    "BIAS: RANGE with BEAR lean",
    "ENTRY PERMISSION: YELLOW",
    "CRUDE MODE: CAUTION (half size) | Direction: FALLING",
    "VIX SIZING: HALF (India VIX = 19.0)",
    "EXPIRY ALERT: Weekly expiry tomorrow (May 21) — theta accelerates today",
])

# ============ SECTION 1 ============
H("SECTION 1 — MACRO SNAPSHOT", level=1)

table = doc.add_table(rows=1, cols=3)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
hdr[0].text = "Metric"; hdr[1].text = "Value"; hdr[2].text = "Signal"

rows = [
    ("MCX Crude (proxy)", "~₹8,755/bbl", "CAUTION zone (half size)"),
    ("Brent / WTI", "$110.12 / $103.00", "FALLING (Trump backed off Iran strike)"),
    ("GIFT Nifty gap", "23,520 (–98 pts vs close)", "Mild gap-down → bearish open"),
    ("S&P 500", "–0.55% (3rd down day)", "Risk-OFF"),
    ("Nasdaq / Dow", "–0.55% / –0.85%", "US weakness persists"),
    ("US VIX", "Elevated (data unavailable)", "Caution"),
    ("India VIX", "19.00 (+4.47%)", "HALF SIZE rule active"),
    ("Nifty close", "₹23,618", "Just above 23,600 max pain"),
    ("FII cash (18-May)", "+₹2,813 Cr", "BUY (latest); MTD net –₹21,842 Cr (SELL)"),
    ("DII cash (18-May)", "+₹2,682 Cr", "BUY (steady support)"),
    ("USD-INR", "~₹85.5 (RBI intervening)", "Weak, rupee pressured by oil"),
    ("INFY ADR", "$12.92 (≈ ₹1,098)", "Slight discount vs NSE ₹1,142–1,656; FLAT-to-DOWN open"),
    ("WIT/IBN/HDB ADR", "DATA_UNAVAILABLE", "—"),
]
for m, v, s in rows:
    row = table.add_row().cells
    row[0].text = m; row[1].text = v; row[2].text = s

P("")
P("Geopolitical one-liner: Iran-Hormuz partially blocked since 28-Feb; Trump halted strike at Saudi/Qatar/UAE request, talks ongoing — crude easing but headline risk LIVE.", bold=True)
P("Today's intraday headline risk: HIGH", bold=True)

# ============ SECTION 2 ============
H("SECTION 2 — CRUDE RULE + STRUCTURAL READ", level=1)
P("CRUDE_MODE: CAUTION (₹8,500–9,000 band) — HALF SIZE", bold=True)
P("CRUDE_DIRECTION: FALLING (Brent down from ~$115 peaks toward $110)", bold=True)
P("CRUDE_FLIP_LEVEL: ₹8,500 — drop below = MILD BULL mode (re-arm crude beneficiaries on the call side)", bold=True)
P("FLAG: Crude falling — watch ₹8,500 for mode flip. Direction tailwind for OMCs / aviation / paints / tyres; headwind for ONGC / Oil India / Reliance E&P segment.", bold=True)
P("AVOID (lean against): pure crude producers (ONGC favoured on PUT side)", bold=False)
P("EXCEPTION CALLS: None required — not in BEAR mode")

P("")
P("STRUCTURAL READ:", bold=True)
P("• Nifty vs Max Pain: 23,618 is +18 pts above max pain (23,600) → magnetism pulling DOWN gently into expiry.")
P("• PCR = 1.35, expiry week → Put writers DEFENDING 23,500 floor; not bullish breakout, just floor support.")
P("• OI change: 23,500 PE has highest OI (floor); 23,700 CE highest OI (ceiling) → range expectation 23,500–23,700.")
P("• Futures basis: monthly expires 26-May; no clear discount/premium reported → DATA_UNAVAILABLE (assume neutral).")
P("• Expiry week dynamics (1 day to weekly): Put writers have more skin (PCR 1.35) → defend ₹23,500. Magnetism direction: DOWN toward 23,600.")
P("• Bank Nifty leadership: DATA_UNAVAILABLE — caution before trusting Nifty direction; verify in first 15-min.")

# ============ SECTION 3 ============
H("SECTION 3 — CONTRARIAN CHECK (Layer 3)", level=1)
P("Q1. CONSENSUS: Every retail tape watcher expects a gap-down sell-day because (a) US down 3rd day, (b) India VIX up 4.47%, (c) GIFT Nifty –98, (d) Hormuz uncertainty.", bold=False)
P("Q2. THE TRAP: A sharp gap-down open that immediately reverses as Trump's de-escalation lands in Indian hours and crude prints lower. Shorts entered at open get squeezed up toward 23,650–23,700. Profit goes to put writers and 23,500 PE shorts.", bold=False)
P("Q3. RETAIL STOPS: Long stops cluster below 23,500 (PE wall). Short stops cluster above 23,700 (CE wall). Whipsaw zone in between.", bold=False)
P("Q4. FLIP TRIGGER: A confirmed Iran-deal headline (Hormuz reopening) during Indian hours → crude crashes –5% → Nifty rips toward 23,800; INDIGO/Asian Paints fly, ONGC tanks. Conversely, any tanker incident in Hormuz → crude spikes +5% → opposite move.", bold=False)

P("")
P("CONTRARIAN OVERRIDE LOGIC:", bold=True)
P("A: GIFT Nifty gap-up >+100 pts?  NO (–98 pts gap-down)")
P("B: Crude falling >2% overnight?  YES")
P("C: Contrarian scenario probability >40%?  YES (Iran talks progressing)")
P("D: Expiry week + PCR <0.8 (squeeze)?  NO (PCR is 1.35 — put writers strong)")
P("→ A is NO → no full bear-override fires. But B+C are YES → consensus bear trade carries reversal risk. Final bias stays RANGE with BEAR lean, NOT outright BEAR. Permission YELLOW.", bold=True)

# ============ SECTION 4 ============
H("SECTION 4 — 5 INDIVIDUAL F&O SETUPS", level=1)

def stock_block(title, lines):
    P("=" * 60)
    P(title, bold=True)
    P("=" * 60)
    for l in lines:
        P(l)

stock_block(
    "[RELIANCE — NSE: RELIANCE] | PUT | Grade A",
    [
        "CMP: ₹1,322  |  Lot: 250  |  Crude aligned: YES (oil-petchem weakness)",
        "CATALYST: Heavyweight that hit late selling on 19-May; bearish technical score 32/100; sits 18% below 52w high. With falling crude squeezing refining + petchem margins and weak USD-INR adding import cost, the index anchor is the cleanest expression of expiry-week downside drift.",
        "TECHNICAL: Daily DOWN | Below 200 DMA: YES (stock is 8% YTD lower) | Weekly DOWN | Support ₹1,290 | Resistance ₹1,360",
        "STRUCTURAL: OI direction: SHORT BUILD (price down + OI up trend) | Delivery %: STABLE | Block deal: NO | Ban list: NO",
        "OPTION SETUP: Strike 1,320 PE (ATM) — use MONTHLY 26-May expiry (only 6 days — still acceptable, weekly is 1 day = theta trap)",
        "⚠️ THETA WARNING (expiry ≤3 days)? NO for monthly; YES if weekly used → AVOID weekly.",
        "ENTRY (ALL 4 must fire):",
        "  • SuperTrend RED on 15-min confirmed",
        "  • RSI(14) < 50 on 15-min",
        "  • StochRSI crossover DOWN from >70",
        "  • Volume > 1.5× 20-period average",
        "  • Confirmation candle: 15-min close BELOW ₹1,315",
        "T1 (book 40%): ₹1,300  |  T2 (book 40%): ₹1,288",
        "SL: ₹1,342 (above intraday VWAP + ST flip)",
        "R:R: ~2.3:1   |   Time stop: Exit by 13:00 IST",
        "GRADE REASON: A — Multi-driver alignment (crude + technical + macro), but R:R only 2.3 caps it below A+.",
    ],
)

stock_block(
    "[ONGC — NSE: ONGC] | PUT | Grade A",
    [
        "CMP: ₹296.40  |  Lot: 4,500  |  Crude aligned: YES (direct beneficiary INVERSE — falling crude hurts)",
        "CATALYST: Crude in CAUTION + FALLING direction. ONGC is the cleanest single-name short proxy on lower oil. Sideways print on 19-May with weakness at the highs suggests distribution near 52w high (₹307).",
        "TECHNICAL: Daily SIDEWAYS-DOWN | Above 200 DMA: YES (₹254.96) — but losing 50 DMA (₹282) would accelerate | Weekly UP-fading | Support ₹290 | Resistance ₹300",
        "STRUCTURAL: OI direction: LONG UNWIND on price weakness | Delivery %: STABLE | Block deal: NO | Ban list: NO",
        "OPTION SETUP: Strike 295 PE (ATM-slightly OTM) — MONTHLY 26-May expiry mandatory (weekly too thin on theta)",
        "⚠️ THETA WARNING (expiry ≤3 days)? NO with monthly used.",
        "ENTRY (ALL 4 must fire):",
        "  • SuperTrend RED on 15-min confirmed",
        "  • RSI(14) < 50 on 15-min",
        "  • StochRSI crossover DOWN from >70",
        "  • Volume > 1.5× 20-period average",
        "  • Confirmation candle: 15-min close BELOW ₹294.50",
        "T1 (book 40%): ₹292  |  T2 (book 40%): ₹289",
        "SL: ₹300.50 (above the 19-May intraday high)",
        "R:R: ~2.1:1   |   Time stop: Exit by 13:00 IST",
        "GRADE REASON: A — Clean crude-inverse thesis; R:R right at the 2:1 floor, so size strictly HALF per VIX rule.",
    ],
)

stock_block(
    "[INTERGLOBE AVIATION — NSE: INDIGO] | CALL | Grade A",
    [
        "CMP: ₹4,230  |  Lot: 75  |  Crude aligned: YES (aviation = direct beneficiary of falling ATF)",
        "CATALYST: Twin tailwinds — (1) Maharashtra cut ATF VAT to 7% from 18% effective 15-May (six-month window); Delhi also cut. (2) Crude FALLING toward $103 WTI. Both compress ATF cost. Mutual funds have been raising stake consistently (Upstox data).",
        "TECHNICAL: Daily DOWN intraday but holding ₹4,224 low | Above 200 DMA: NO (still under recovery; 52w low ₹3,895, high ₹6,232) | Weekly basing | Support ₹4,220 | Resistance ₹4,280 / ₹4,320",
        "STRUCTURAL: MF stake rising = LONG BUILD trend | Delivery %: RISING | Block deal: NO | Ban list: NO",
        "OPTION SETUP: Strike 4,300 CE (slight OTM) — MONTHLY 26-May expiry",
        "⚠️ THETA WARNING (expiry ≤3 days)? NO with monthly used.",
        "ENTRY (ALL 4 must fire):",
        "  • SuperTrend GREEN on 15-min confirmed",
        "  • RSI(14) > 50 on 15-min",
        "  • StochRSI crossover UP from <30",
        "  • Volume > 1.5× 20-period average",
        "  • Confirmation candle: 15-min close ABOVE ₹4,260",
        "T1 (book 40%): ₹4,300  |  T2 (book 40%): ₹4,335",
        "SL: ₹4,220 (below 19-May low + ST flip)",
        "R:R: ~2.6:1   |   Time stop: Exit by 13:00 IST",
        "GRADE REASON: A — Best fundamental tailwind today (VAT cut + crude); only held back by Nifty risk-off backdrop.",
    ],
)

stock_block(
    "[ASIAN PAINTS — NSE: ASIANPAINT] | CALL | Grade B",
    [
        "CMP: ₹2,614  |  Lot: 100  |  Crude aligned: YES (TiO2 / monomer cost relief on falling oil)",
        "CATALYST: Falling crude eases raw material cost (~30% input link). Demonstrated resilience even during crude RISE — should see beta on the way down. No earnings catalyst today.",
        "TECHNICAL: Daily SIDEWAYS-UP | Above 200 DMA: borderline (DATA_UNAVAILABLE precise) | Weekly basing | Support ₹2,580 | Resistance ₹2,650",
        "STRUCTURAL: OI direction: NEUTRAL | Delivery %: STABLE | Block deal: NO | Ban list: NO",
        "OPTION SETUP: Strike 2,640 CE (OTM) — MONTHLY 26-May expiry",
        "⚠️ THETA WARNING (expiry ≤3 days)? NO with monthly used.",
        "ENTRY (ALL 4 must fire):",
        "  • SuperTrend GREEN on 15-min confirmed",
        "  • RSI(14) > 50 on 15-min",
        "  • StochRSI crossover UP from <30",
        "  • Volume > 1.5× 20-period average",
        "  • Confirmation candle: 15-min close ABOVE ₹2,625",
        "T1 (book 40%): ₹2,650  |  T2 (book 40%): ₹2,675",
        "SL: ₹2,598 (below pivot)",
        "R:R: ~2.2:1   |   Time stop: Exit by 13:00 IST",
        "GRADE REASON: B — Right theme but momentum lukewarm; consumer-defensive doesn't move enough for half-size sizing.",
    ],
)

stock_block(
    "[INFOSYS — NSE: INFY] | PUT | Grade B",
    [
        "CMP: ₹1,656 (NSE) / ADR $12.92 (≈ ₹1,098 implied) — flag DATA CONFLICT across sources; use NSE close",
        "Lot: 400  |  Crude aligned: N/A (USD-INR weakness is a partial offset for IT, but US tape down 3 days dominates).",
        "CATALYST: S&P 500, Nasdaq down 3 straight; INFY itself –25% over 6 months; ADR signal flat-to-soft. Index-heavy IT name to express US risk-off without committing full size.",
        "TECHNICAL: Daily DOWN | Above 200 DMA: NO (52w low ₹1,089 vs CMP ₹1,656 — but trend is down) | Weekly DOWN | Support ₹1,640 | Resistance ₹1,680",
        "STRUCTURAL: OI direction: SHORT BUILD | Delivery %: STABLE | Block deal: NO | Ban list: NO",
        "OPTION SETUP: Strike 1,640 PE (ATM-OTM) — MONTHLY 26-May expiry",
        "⚠️ THETA WARNING (expiry ≤3 days)? NO with monthly used.",
        "ENTRY (ALL 4 must fire):",
        "  • SuperTrend RED on 15-min confirmed",
        "  • RSI(14) < 50 on 15-min",
        "  • StochRSI crossover DOWN from >70",
        "  • Volume > 1.5× 20-period average",
        "  • Confirmation candle: 15-min close BELOW ₹1,648",
        "T1 (book 40%): ₹1,635  |  T2 (book 40%): ₹1,620",
        "SL: ₹1,672 (above 19-May intraday high)",
        "R:R: ~2.0:1   |   Time stop: Exit by 13:00 IST",
        "GRADE REASON: B — Source-price conflict (₹1,142 vs ₹1,656 across data) is a data-integrity risk; verify on terminal before entry. Setup valid only with NSE-confirmed price.",
    ],
)

# ============ SECTION 5 ============
H("SECTION 5 — FINAL VERDICT", level=1)
BOX([
    "TODAY'S BIAS:    RANGE with BEAR lean",
    "CONFIDENCE:      MEDIUM",
    "CRUDE MODE:      CAUTION (half size) — FALLING",
    "VIX SIZING:      HALF (VIX = 19.00)",
    "ENTRY PERMISSION: YELLOW",
])

P("")
P("THE BULL CASE (even if bias = BEAR):", bold=True)
P("Iran/Hormuz de-escalation headline during Indian hours crashes crude → INDIGO/ASIANPAINT rip; max-pain 23,600 + PCR 1.35 + DII steady buying defends 23,500 floor; FIIs flipped buyers on 18-May.")

P("THE BEAR CASE (even if bias = BULL):", bold=True)
P("US down 3rd straight day, INFY/RELIANCE technically broken, GIFT Nifty –98 pts gap-down, India VIX up 4.47% pre-expiry — magnetism toward 23,600 then 23,500.")

P("")
P("FLIP TRIGGER:", bold=True)
P("If Nifty closes 15-min above 23,700 with Bank Nifty leading → bias flips to BULL (intraday squeeze toward 23,800). If WTI prints below $100 on Iran-deal headline → flip to CAUTIOUS BULL, dump PUT setups on RELIANCE/ONGC, ride INDIGO/ASIANPAINT CALLs hard.")

P("")
P("NIFTY KEY LEVELS:", bold=True)
P("S2 ₹23,400  |  S1 ₹23,500 (put wall)  |  CRITICAL ₹23,600 (max pain)  |  R1 ₹23,700 (call wall)  |  R2 ₹23,800")

P("")
P("TOP 3 RANKED:", bold=True)
P("#1 RELIANCE PUT — A — heavyweight tech-break + crude/petchem drag — entry below ₹1,315")
P("#2 ONGC PUT — A — clean crude-inverse short proxy — entry below ₹294.50")
P("#3 INDIGO CALL — A — ATF VAT cut + falling crude double-tailwind — entry above ₹4,260")

P("")
P("ONE RISK THAT RUINS EVERYTHING TODAY:", bold=True)
P("A surprise Iran retaliation / tanker incident in the Hormuz \"vast operational area\" during Indian hours. Crude spikes +5%, ALL bear setups (RELIANCE, INFY, ONGC PUTs) implode while ONGC inverts and the aviation/paints CALL legs evaporate. The book is double-exposed in opposite directions to crude — risk-down position size accordingly.")

# ============ ONE-LINER ============
H("ONE-LINE SUMMARY (read at 9:10 AM)", level=1)
P("\"Today is RANGE-BEAR because US is down 3 days and GIFT Nifty gaps –98 into expiry. Crude ₹8,755 = CAUTION, falling. Watch RELIANCE PUT and INDIGO CALL. Key risk: Hormuz tanker incident. Size HALF. Flips if Nifty closes 15-min above 23,700 with Bank Nifty leading.\"",
  bold=True)

# ============ DATA NOTES ============
H("DATA NOTES & FLAGS", level=1)
P("• INFY price conflict across sources (₹1,142 vs ₹1,656) — verify on terminal before trading.")
P("• WIT / IBN / HDB ADRs: DATA_UNAVAILABLE — use Yahoo Finance manual check pre-open.")
P("• USD-INR: precise value DATA_UNAVAILABLE; RBI noted intervening to defend rupee against oil-driven weakness.")
P("• Bank Nifty close & leadership vs Nifty: DATA_UNAVAILABLE — verify Sectoral Index in first 15-min.")
P("• Sector top gainers/losers for 19-May: DATA_UNAVAILABLE from search aggregation.")
P("• FII/DII figures shown are for 18-May (latest in search); 19-May provisional figures not yet released at brief generation time.")
P("• Crude price-to-MCX proxy uses USD-INR ≈ ₹85 — adjust if open print differs.")

doc.save(OUT_PATH)
print(f"Saved: {OUT_PATH}")
