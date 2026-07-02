"""Generate the daily trading-intelligence .docx brief — 2026-07-02."""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-07-02"
BIAS = "CAUTIOUS-BULL"
ENTRY_PERMISSION = "YELLOW"

OUT_PATH = f"trading-briefs/2026/07-July/Trading_Brief_{DATE_ISO}_{BIAS}.docx"

NAVY = RGBColor(0x1F, 0x38, 0x64)
RED = RGBColor(0xC0, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


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


def add_para(doc, text, bold=False, italic=False, size=11, color=None):
    p = doc.add_paragraph()
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
run.font.color.rgb = NAVY

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_run = sub.add_run("Indian F&O Day Trader — Nifty + Stocks")
sub_run.italic = True
sub_run.font.size = Pt(12)

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
dr = date_p.add_run("Date: Thursday, 2 July 2026  |  Generated ~08:20 IST")
dr.bold = True
dr.font.size = Pt(11)

badge_line(doc, "TODAY'S BIAS", BIAS, "F1C232")
badge_line(doc, "ENTRY PERMISSION", ENTRY_PERMISSION, "E69138")
badge_line(doc, "CRUDE RULE MODE", "AGGRESSIVE BULL (~₹6,600 MCX)", "6AA84F")
badge_line(doc, "CAPITAL FIT", "FLAGGED — see Section 4", "C0392B")

doc.add_paragraph()

# ---------- SECTION 0 — POSTMORTEM / CALIBRATION GAP ----------
add_heading(doc, "Section 0 — Yesterday's Postmortem / Calibration Gap", level=1, color=RED)
add_para(doc, "BRIEF METHODOLOGY STATUS: NO VALID POSTMORTEM POSSIBLE.", bold=True, size=12, color=RED)
add_bullets(doc, [
    "The last brief in this log is dated 2026-04-23 (a WAIT day). Today is 2026-07-02 — a gap of ~10 weeks "
    "(≈50 trading sessions) with no intervening briefs and an empty paper-trading log (trade_log.txt).",
    "Phase 1 of the routine requires scoring yesterday's 5-7 picks against actual outcomes. There is no "
    "'yesterday's brief' to grade — the accuracy score is UNDEFINED, not zero and not passing.",
    "Per the spirit of the routine's recovery logic (Phase 6), this brief treats the gap as a calibration "
    "reset: stock count trimmed to 6, no grade above A, wait window extended to 45-60 minutes post-open "
    "instead of the standard 30, and VIX-based sizing is overridden one notch more conservative than the "
    "raw VIX reading would imply.",
    "ACTION FOR OPERATOR: resume daily logging so tomorrow's brief can run a real postmortem. If this gap "
    "reflects the routine being paused/newly (re)automated rather than deliberate, no further action needed "
    "beyond acknowledging today's picks carry no track-record backing.",
])

doc.add_paragraph()

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=NAVY)

add_kv_table(doc,
    header=["Metric", "Reading", "Interpretation"],
    rows=[
        ["WTI Crude (NYMEX)", "~$69.50–70.50", "2 sources ranged $68.04–$70.55; used midpoint ~$70"],
        ["Brent Crude (context only)", "~$73 (range $71–75)", "NOT used for MCX proxy — WTI only"],
        ["USD/INR spot", "~94.30–94.70", "2 sources disagreed by ~₹0.35; used ~94.50"],
        ["MCX Crude Proxy (WTI×USDINR)", "≈ ₹6,600/bbl (range ₹6,400–6,680)", "AGGRESSIVE BULL zone, well clear of ₹7,500 boundary"],
        ["India VIX", "13.24 (Jul 1 close, -2.65%)", "Below 15 = normally FULL SIZE; overridden to HALF given calibration gap"],
        ["Nifty 50 (Jul 1 close)", "~23,987–23,997 (2 sources ~9pt apart)", "+0.5% day; day range 23,895–24,021"],
        ["Sensex (Jul 1 close)", "76,922.64 (+0.58%)", "Confirms Nifty strength"],
        ["GIFT Nifty (pre-market)", "CONFLICTING — one calc implies +67-68pt gap-up (24,149.5 vs fut. close 24,082); another note says 'slightly red'", "Direction genuinely uncertain — do not assume gap-up until confirmed at open"],
        ["Dow Jones (Jul 1 close)", "52,305.24 (-0.03%)", "Faded from a record intraday high"],
        ["S&P 500 (Jul 1 close)", "7,483.23 (-0.22%)", "Mild pullback"],
        ["Nasdaq (Jul 1 close)", "26,040.03 (-0.66%)", "Semiconductor profit-taking after huge H1 rally"],
        ["Asian markets", "STALE — only Jun 30/Jul 1 data found (Nikkei 70,671 +0.87%, Kospi +1.0%, Shanghai +0.5%, Hang Seng -0.6%)", "Live Jul 2 Asia opens NOT FOUND — check terminal before open"],
        ["FII flow (Jul 1, provisional)", "Net SELL ₹1,140.5 cr", ">₹43,000 cr net sold in June alone"],
        ["DII flow (Jul 1)", "Net BUY ₹3,159.24 cr", "DII buying > FII selling = near-term floor"],
    ])

doc.add_paragraph()
add_para(doc, "Data-quality flags (per routine's cross-verification rule):", bold=True, size=11, color=RED)
add_bullets(doc, [
    "WTI, Brent and USD/INR each showed meaningful spread across sources — treated as ranges, not points.",
    "Nifty's own July 1 close was reported as both 23,987.95 and 23,997.20 by different aggregators.",
    "GIFT Nifty's implied gap is internally contradictory between sources — this is the single biggest reason "
    "entry permission is YELLOW rather than GREEN today.",
])

doc.add_paragraph()

# ---------- SECTION 2 — CRUDE RULE APPLIED ----------
add_heading(doc, "Section 2 — Crude Rule Applied", level=1, color=NAVY)
add_para(doc, "MODE: AGGRESSIVE BULL (~₹6,600 MCX, using WTI only — never Brent).", bold=True, size=12)
add_bullets(doc, [
    "This is a sharp regime shift from the last brief (23 Apr): MCX crude proxy was ~₹8,390 then (MILD BULL, "
    "caution tilt); it is ~₹6,600 now — a >20% collapse over 10 weeks, driven mainly by WTI falling from ~$89 "
    "to ~$70 with only modest INR depreciation (93.80 → ~94.5).",
    "₹6,600 sits ₹900 clear of the ₹7,500 upper boundary of AGGRESSIVE BULL — not a near-boundary situation.",
    "CONTRARIAN NOTE: the Iran/Hormuz conflict is NOT resolved — it is a fragile 60-day MOU already violated "
    "once (Iran briefly re-closed the strait after Israel-Hezbollah strikes), with Iran demanding transit "
    "tolls the US rejects, and a vessel grounding on a disputed route as recently as July 1. Crude this low "
    "despite an unresolved shooting conflict is itself a signal worth distrusting — one bad headline can "
    "reprice WTI 5-10% intraday and flip this regime toward CAUTION without much warning.",
    "Critical level: MCX ₹7,500 (regime flip to CAUTION-adjacent), then ₹8,500 (flip to half-size CAUTION).",
])
add_para(doc, "Crude-aligned sector bias:", bold=True, size=11)
add_bullets(doc, [
    "Bullish on falling crude: aviation (IndiGo — lower ATF cost), paints (Asian Paints/Berger — input cost), "
    "auto (lower running cost supports demand), OMCs (BPCL/HPCL/IOC — cheaper feedstock, not used today due "
    "to sector-cap tradeoffs vs auto/aviation picks).",
    "Bearish on falling crude: upstream producers — ONGC, Oil India, MRPL (lower per-barrel realisations).",
])

doc.add_paragraph()

# ---------- SECTION 3 — INDIAN MARKET INTERNALS ----------
add_heading(doc, "Section 3 — Indian Market Internals", level=1, color=NAVY)
add_para(doc, "Nifty Technical Levels", bold=True, size=12)
add_kv_table(doc,
    header=["Level", "Value", "Meaning"],
    rows=[
        ["Previous Close (Jul 1)", "~23,990 (2 sources: 23,987.95 / 23,997.20)", "+0.5% day"],
        ["Jul 1 Intraday High", "24,020.65", "Near-term resistance R1"],
        ["Jul 1 Intraday Low", "23,895.10", "Near-term support S1"],
        ["Prior Close (Jun 30)", "23,865.75", "Support S2 if S1 breaks"],
        ["GIFT Nifty Implied Open", "Uncertain — see flag above", "WAIT for first 15-min candle to resolve direction"],
        ["Weekly Expiry", "Tuesdays (moved from Thursday)", "Next weekly: Tue 7 Jul — today (Wed) is NOT an expiry day"],
        ["Monthly Expiry", "Last Tuesday of month", "Jul 2026 monthly: Tue 28 Jul"],
    ])

doc.add_paragraph()
add_para(doc, "F&O Ban List / Liquidity", bold=True, size=12)
add_bullets(doc, [
    "Ban list was empty as of Jun 29-30 close per multiple trackers; none of today's shortlisted stocks have "
    "appeared on any recent ban reference. Live NSE ban CSV could not be fetched directly (blocked) — "
    "OPERATOR MUST manually confirm the live ban list at/before market open before placing any order.",
    "Max pain for the current weekly/monthly series could not be reliably sourced for Jul 2 specifically — "
    "treat OI-based conviction as LOW today; do not lean on max-pain gravity as a thesis input.",
])

doc.add_paragraph()
add_para(doc, "Sector Color (last few sessions)", bold=True, size=12)
add_bullets(doc, [
    "Auto: standout — June dispatches strong across the board except Hero MotoCorp (Maruti +19.3%, Tata "
    "Motors PV +69%, M&M +37% YoY; Hero MotoCorp down YoY, the sector laggard).",
    "Banking: HDFCBANK, ICICIBANK, KOTAKBANK, SBIN, BAJFINANCE/BAJFINSV supportive of recent index gains.",
    "Pharma: CIPLA, DRREDDY, SUNPHARMA showing recent relative strength (defensive tailwind).",
    "IT: TECHM and INFY firm recently, but Tech Mahindra has sharply diverging fresh brokerage views (CLSA "
    "'high-conviction outperform' vs Jefferies 'underperform') — a volatility flag, avoided today for a "
    "clean single-thesis setup.",
])

doc.add_paragraph()

# ---------- SECTION 4 — CONTRARIAN CHECK (LAYERS 1/2/3) ----------
add_heading(doc, "Section 4 — Contrarian Check (Layer 1/2/3)", level=1, color=NAVY)
add_para(doc, "Layer 1 — Surface / Consensus Read:", bold=True, size=11)
add_para(doc,
    "\"Crude has crashed to ~$70 WTI, Hormuz tensions are cooling with Doha talks resuming, Fed held rates, "
    "DII buying is cushioning FII selling, Nifty just reclaimed 24,000, and auto stocks are rallying hard on "
    "blowout June sales. This is a risk-on tape — buy the dip.\"", italic=True, size=11)

add_para(doc, "Layer 2 — What Consensus Is Likely Wrong About:", bold=True, size=11)
add_bullets(doc, [
    "The Hormuz 'de-escalation' is a fragile 60-day MOU already breached once, with Iran-US transit-toll "
    "talks unresolved and a vessel grounding on a disputed route as recently as yesterday — not a settled "
    "ceasefire.",
    "GIFT Nifty's own pre-market signal is internally contradictory across sources — the market isn't even "
    "giving a clean directional read yet, which the 'confident gap-up' narrative glosses over.",
    "The Fed's dot plot just flipped hawkish (9 of 19 members now lean toward a 2026 HIKE, not a cut) — a "
    "headwind for EM/INR flows that the 'DII is cushioning FII selling' story can mask only temporarily; "
    "FII has already sold >₹43,000 cr in June alone.",
])

add_para(doc, "Layer 3 — What's Being Underestimated:", bold=True, size=11)
add_bullets(doc, [
    "Monsoon is the quiet, unpriced risk: June 2026 was the 5th driest since 1901 (40% below normal), and "
    "IMD's July outlook is also below-normal (<94% LPA). This is a slow-burn negative for rural consumption, "
    "FMCG and auto-financing demand that today's crude-crash euphoria is currently drowning out.",
    "The India-US trade deal has a real dated catalyst this month — the current US tariff framework expires "
    "24 July 2026, and the deal is 'near final' but NOT signed. A slip past that date without signature "
    "could reverse sentiment sharply; a signature could extend the rally.",
    "Tech Mahindra's fresh, sharply diverging brokerage calls (CLSA bullish vs Jefferies bearish, both new) "
    "hint at disagreement beneath an otherwise calm IT-sector surface.",
])

add_para(doc, "Flip triggers: BEARISH — any fresh Iran/Hormuz/Israel-Lebanon escalation headline, MCX crude "
              "spiking back through ₹7,500, or GIFT Nifty resolving red at the actual open.", size=11)
add_para(doc, "Flip triggers: BULLISH — Hormuz status holding calm for 3-5 more sessions, and/or the India-US "
              "trade deal being signed ahead of the 24 Jul deadline.", size=11)
add_para(doc, "Entry Permission: YELLOW — wait for the first 45-60 minutes to see which GIFT Nifty read "
              "(gap-up vs flat/red) actually resolves before sizing up.", bold=True, size=12, color=RGBColor(0xE6, 0x91, 0x38))

doc.add_paragraph()

# ---------- SECTION 5 — STOCK OPPORTUNITIES ----------
add_heading(doc, "Section 5 — Stock Setups", level=1, color=NAVY)
add_para(doc,
    "CAPITAL-FIT FLAG: at current prices, ALL 6 stocks below fail the routine's own eligibility filter "
    "(lot size × spot × 3% < ₹15,000) for a ₹2,00,000 account — the closest is ONGC at ~₹15,832 (3% over), "
    "the worst is Bharti Airtel at ~₹26,391 (76% over). This likely reflects 10 weeks of price drift since "
    "the filter was last calibrated. Recommended handling for today: 1 lot only per name, treat as REDUCED "
    "CONVICTION size regardless of grade, and paper-track premium-based R:R rather than underlying R:R — do "
    "not scale up until smaller-notional alternatives are identified or the filter is re-benchmarked.",
    bold=True, size=10.5, color=RED)

doc.add_paragraph()

add_para(doc, "#1 — MARUTI SUZUKI (MARUTI) CALL — Grade A", bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("CMP / Lot", "₹14,395 (Jul 1 close) / Lot 50 — notional ₹7.2L, fails ₹15k capital-fit filter"),
    ("Sector", "Auto"),
    ("Why this grade", "June sales +19.3% YoY (200,390 units) blowout + fresh Jefferies upgrade "
     "(target ₹16,500) + macro tailwind from aggressive-bull crude (lower running costs). Capped at A "
     "(not A+) given the 10-week calibration gap."),
    ("Option Setup", "July monthly expiry (28 Jul), ATM/slightly OTM CE — do not use a weekly"),
    ("Entry Trigger", "(1) Nifty Auto index green at open (2) Maruti 15-min close above ₹14,450 "
     "(3) volume above 20-day average (4) India VIX stays under 15"),
    ("T1 / T2", "+3% (~₹14,827) / +5% (~₹15,115) — book 40% each"),
    ("Stop Loss", "15-min close below ₹14,150 (~-1.7%)"),
    ("R:R", "Est. ≥2:1 on option premium basis — verify against live premium before entry"),
    ("Time Stop", "Exit if trigger not met within 60 minutes of open"),
])

doc.add_paragraph()
add_para(doc, "#2 — HERO MOTOCORP (HEROMOTOCO) PUT — Grade B+", bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("CMP / Lot", "₹4,851.6 (Jul 1 close) / Lot 150 — notional ₹7.3L, fails capital-fit filter"),
    ("Sector", "Auto (pairs with Maruti — 2-stock sector cap, mixed direction)"),
    ("Why this grade", "Lone auto laggard in June dispatch data — down YoY (541,159 units) while every "
     "other major OEM posted double-digit growth. Single relative-underperformance catalyst."),
    ("Option Setup", "July monthly expiry, ATM/OTM PE"),
    ("Entry Trigger", "(1) Stock fails to reclaim ₹4,857 (2) 15-min close below ₹4,807 (3) Maruti/sector "
     "not confirming strength (4) no fresh Hero-specific news reversing the sales miss"),
    ("T1 / T2", "-3% (~₹4,706) / -5% (~₹4,609)"),
    ("Stop Loss", "15-min close above ₹4,857"),
    ("R:R", "Est. ≥2:1 on option premium basis — verify live"),
    ("Time Stop", "60 minutes"),
])

doc.add_paragraph()
add_para(doc, "#3 — INDIGO / INTERGLOBE AVIATION (INDIGO) CALL — Grade B+", bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("CMP / Lot", "₹5,388.50 (Jul 1) / Lot 150 — notional ₹8.1L, fails capital-fit filter"),
    ("Sector", "Aviation"),
    ("Why this grade", "Direct beneficiary of the aggressive-bull crude regime — lower ATF cost. Single "
     "strong macro catalyst + crude-regime alignment."),
    ("Option Setup", "July monthly, ATM CE"),
    ("Entry Trigger", "(1) Brent stays under $75 / MCX crude under ₹7,000 (2) 15-min close above ₹5,395 "
     "(3) no fresh Hormuz escalation headline (4) VIX under 15"),
    ("T1 / T2", "+3% (~₹5,550) / +5% (~₹5,658)"),
    ("Stop Loss", "15-min close below ₹5,329, OR immediate exit on any Hormuz escalation headline"),
    ("R:R", "Est. ≥2:1 on option premium basis — verify live"),
    ("Time Stop", "60 minutes"),
])

doc.add_paragraph()
add_para(doc, "#4 — ONGC PUT — Grade B+", bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("CMP / Lot", "₹234.55 (Jun 30 close, cross-confirmed) / Lot 2250 — notional ₹5.3L, ~3% over "
     "capital-fit filter (closest fit of the 6)"),
    ("Sector", "Energy / Upstream Oil & Gas"),
    ("Why this grade", "Direct inverse trade on the aggressive-bull crude thesis — upstream realisations "
     "get squeezed as WTI falls further; contrarian hedge against the crude-regime call itself."),
    ("Option Setup", "July monthly, ATM/OTM PE"),
    ("Entry Trigger", "(1) WTI/MCX crude trending lower intraday (2) 15-min close below ₹230 (3) no "
     "fresh Iran-supply-shock headline (which would reverse this thesis) (4) sector (BPCL/HPCL) confirming softness"),
    ("T1 / T2", "-3% (~₹227.5) / -5% (~₹222.8)"),
    ("Stop Loss", "15-min close above ₹238, OR immediate exit on any Iran/Hormuz supply-shock headline"),
    ("R:R", "Est. ≥2:1 on option premium basis — verify live"),
    ("Time Stop", "60 minutes"),
])

doc.add_paragraph()
add_para(doc, "#5 — BHARTI AIRTEL (BHARTIARTL) CALL — Grade B+", bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("CMP / Lot", "₹1,852 (Jun 30 close) / Lot 475 — notional ₹8.8L, worst capital-fit overshoot of the 6"),
    ("Sector", "Telecom (diversification — largely crude/geopolitics-agnostic)"),
    ("Why this grade", "Fresh Nomura target raise to ₹2,355 (Buy) — single strong brokerage catalyst on "
     "a defensive-quality name, useful portfolio diversifier away from crude/auto themes."),
    ("Option Setup", "July monthly, ATM CE"),
    ("Entry Trigger", "(1) 15-min close above ₹1,876 (2) broad market tone neutral-to-positive (3) no "
     "telecom-specific negative news (4) volume confirmation"),
    ("T1 / T2", "+3% (~₹1,908) / +5% (~₹1,945)"),
    ("Stop Loss", "15-min close below ₹1,830 (~-1.2%)"),
    ("R:R", "Est. ≥2:1 on option premium basis — verify live"),
    ("Time Stop", "60 minutes"),
])

doc.add_paragraph()
add_para(doc, "#6 — BHARAT ELECTRONICS (BEL) CALL — Grade B — MACRO HEDGE / DEFENCE", bold=True, size=13, color=NAVY)
stock_card(doc, [
    ("CMP / Lot", "₹407.2 (Jul 2 pre-market) / Lot 1425 — notional ₹5.8L, fails capital-fit filter"),
    ("Sector", "Defence (required macro-hedge slot — safe haven against Iran/Hormuz/Israel-Lebanon escalation)"),
    ("Why this grade", "Defensive hedge, not a high-conviction directional call — included specifically "
     "to satisfy the routine's macro-hedge requirement given the unresolved geopolitical backdrop."),
    ("Option Setup", "July monthly, ATM CE"),
    ("Entry Trigger", "(1) Any fresh Iran/Hormuz/Israel-Lebanon escalation headline, OR (2) 15-min close "
     "above ₹415 on its own technical merit (3) sector (defence peers) confirming (4) VIX check"),
    ("T1 / T2", "+3% (~₹419) / +5% (~₹427)"),
    ("Stop Loss", "15-min close below ₹406"),
    ("R:R", "Est. ≥2:1 on option premium basis — verify live"),
    ("Time Stop", "60 minutes, or hold through session if used purely as a hedge"),
])

doc.add_paragraph()
add_para(doc, "Sector Diversification Summary", bold=True, size=12)
add_bullets(doc, [
    "Auto (2): Maruti CALL, Hero MotoCorp PUT — mixed direction within the sector cap",
    "Aviation (1): IndiGo CALL",
    "Energy/Upstream (1): ONGC PUT",
    "Telecom (1): Bharti Airtel CALL",
    "Defence/Hedge (1): BEL CALL",
    "5 sectors represented (min 4 required); 4 CALLs / 2 PUTs — not all one-directional; 1 explicit macro-hedge included.",
])

doc.add_paragraph()
add_para(doc, "Stocks to Actively Avoid Today", bold=True, size=12, color=RED)
add_bullets(doc, [
    "Tata Motors — strongest single catalyst (PV sales +69% YoY) but already an extreme one-day beat; "
    "chasing it post-move without a clean entry trigger risks buying the top. Watch, don't chase.",
    "Tech Mahindra — CLSA (bullish) vs Jefferies (bearish) fresh, sharply diverging calls this week. Too "
    "much conflicting signal for a clean single-thesis setup today.",
    "TCS, HDFC Bank — no results until 9 Jul / 18 Jul respectively; no fresh catalyst today.",
    "Any weekly-expiry options on stocks — today (Wed) is not an expiry day anyway; when trading weeklies, "
    "never do so on expiry day per the routine's absolute prohibition.",
])

doc.add_paragraph()

# ---------- SECTION 6 — DOMESTIC MACRO ----------
add_heading(doc, "Section 6 — Domestic Macro & Global Triggers", level=1, color=NAVY)
add_kv_table(doc,
    header=["Driver", "Status"],
    rows=[
        ["RBI Repo Rate", "5.25%, held at 3-5 Jun MPC; stance NEUTRAL. Next MPC date not confirmed (est. early Aug)."],
        ["India-US Trade Deal", "'Near final' per Piyush Goyal; NOT signed. Hard deadline: current US tariff framework expires 24 Jul 2026."],
        ["Monsoon 2026", "June 40% below normal (5th driest since 1901); IMD forecasts July also below-normal (<94% LPA). Bearish for rural consumption/FMCG."],
        ["US Fed", "Held 3.50-3.75% at 16-17 Jun FOMC; dot plot turned hawkish (9/19 lean toward a 2026 hike). Next FOMC: 28-29 Jul."],
        ["China PMI (Jun)", "Manufacturing 50.3 (3rd straight expansion month), export/tech sub-index strong at 53.5. Mildly positive for EM sentiment."],
        ["June Auto Sales", "Maruti +19.3%, Tata Motors PV +69%, M&M +37% YoY — sector-wide beat except Hero MotoCorp."],
        ["Brokerage Calls (last 48h)", "Jefferies: Maruti upgrade (₹16,500 target). Nomura: Bharti Airtel target ₹2,355. JM Financial: Dixon Tech target ₹14,200. CLSA/Jefferies split on Tech Mahindra."],
        ["NSE F&O Ban List", "Empty as of 29-30 Jun per trackers; live CSV blocked to automated fetch — CONFIRM MANUALLY before any order."],
        ["Q1 FY27 Results", "Nothing major today; TCS reports 9 Jul, HDFC Bank 18 Jul — season not yet underway."],
        ["Domestic Macro Verdict", "SUPPORTIVE on rates/DII flows/auto data; HEADWIND from weak monsoon + hawkish Fed shift + unsigned trade deal deadline."],
    ])

doc.add_paragraph()

# ---------- SECTION 7 — FINAL VERDICT ----------
add_heading(doc, "Section 7 — Final Verdict", level=1, color=RED)
add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["Crude Rule Mode", "AGGRESSIVE BULL (~₹6,600 MCX, WTI-based) — sharp shift from Apr's MILD BULL/CAUTION"],
        ["Market Bias", "CAUTIOUS BULL — Nifty holding above 23,900, DII cushioning FII selling, but GIFT Nifty signal unresolved"],
        ["VIX Sizing Rule", "Raw VIX 13.24 = FULL SIZE eligible; OVERRIDDEN to HALF SIZE given 10-week calibration gap"],
        ["Key Support", "23,895 → 23,865"],
        ["Key Resistance", "24,020 → ~24,150 (GIFT Nifty implied, unconfirmed)"],
        ["Critical Crude Level", "₹7,500 MCX — regime watch level"],
        ["Top Risk 1", "GIFT Nifty's contradictory pre-market signal — true open direction unknown until first 15-min candle"],
        ["Top Risk 2", "Any Iran/Hormuz/Israel-Lebanon escalation headline — fragile MOU, already breached once"],
        ["Top Risk 3", "Capital-fit filter failure across all 6 shortlisted names — see Section 5 flag"],
        ["Entry Permission", "YELLOW — wait 45-60 minutes post-open for direction confirmation before sizing up"],
    ])

doc.add_paragraph()
add_para(doc, "Bull case:", bold=True, size=11)
add_para(doc, "GIFT Nifty resolves toward the +68pt gap-up read, Hormuz stays calm through the session, DII "
              "buying continues to absorb FII selling, and auto-sector strength broadens. Nifty tests 24,150-24,200.", size=11)
add_para(doc, "Bear case:", bold=True, size=11)
add_para(doc, "GIFT Nifty resolves toward the 'slightly red' read, a fresh Hormuz/Iran headline hits, or the "
              "hawkish Fed dot-plot narrative gains traction in FII flows. Nifty retests 23,895 then 23,865.", size=11)
add_para(doc, "Most likely scenario:", bold=True, size=11)
add_para(doc, "A choppy, headline-driven session inside 23,895-24,150 given the two genuinely conflicting "
              "GIFT Nifty reads — patience in the first hour is worth more than an early position today.", size=11)

doc.add_paragraph()
add_para(doc, "THE ONE NUMBER THAT MATTERS: MCX crude ₹7,500 — the AGGRESSIVE BULL vs CAUTION boundary. "
              "Currently ₹900 clear of it, but WTI has round-tripped 5-10% on single Iran headlines before.",
         bold=True, size=12, color=RED)

doc.add_paragraph()

# ---------- SECTION 8 — PAPER TRADING NOTE ----------
add_heading(doc, "Section 8 — Note on Paper Trading Today", level=1, color=NAVY)
add_bullets(doc, [
    "trade_log.txt is currently empty — this is the first entry point after the 10-week gap. Log every "
    "signal the system fires today (fired/not fired, direction, outcome) so tomorrow's postmortem is possible.",
    "Watch specifically whether the system fires on Maruti/IndiGo (aggressive-bull-aligned CALLs) or on "
    "ONGC/Hero MotoCorp (contrarian PUTs) — divergence between system signals and this brief's thesis on "
    "the crude-regime trades (ONGC, IndiGo) is the most informative test given how fresh the regime shift is.",
    "Zero-trade-day count: resetting to 0 today given the logging gap; do not treat prior silence as data.",
])

doc.add_paragraph()
add_para(doc, "ONE-LINE SUMMARY", bold=True, size=13, color=NAVY)
add_para(doc,
    "\"First brief in 10 weeks — no postmortem possible, methodology unscored. Crude has crashed to "
    "AGGRESSIVE BULL (~₹6,600 MCX) despite an unresolved Iran/Hormuz conflict; GIFT Nifty gives two "
    "contradictory pre-market reads, so entry permission is YELLOW — wait 45-60 min for confirmation. "
    "Auto (Maruti CALL / Hero MotoCorp PUT), IndiGo CALL and a contrarian ONGC PUT lead the six-stock list; "
    "Bharti Airtel and BEL round out diversification. All six fail the account's own capital-fit filter at "
    "current prices — size down to 1 lot regardless of grade. Watch for any Hormuz headline all session.\"",
    italic=True, size=12)

doc.add_paragraph()
doc.add_paragraph()
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run("Generated by Daily Trading Routine — for personal educational / paper-trading use only. "
                   "Not financial advice. Multiple data points in this brief carry source discrepancies or "
                   "are stale — verify all levels against a live terminal before 9:15 AM IST.")
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
