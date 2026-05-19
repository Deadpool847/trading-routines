"""Generate the daily trading-intelligence .docx brief for 2026-05-19."""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-19"
DATE_LONG = "Tuesday, 19 May 2026"
GEN_TIME = "08:50 IST"
BIAS = "RANGE"
ENTRY_PERMISSION = "YELLOW"
CRUDE_MODE = "BEAR (puts-only, FALLING)"
VIX_SIZING = "QUARTER"
CONFIDENCE = "LOW"

OUT_DIR = "trading-briefs/2026/05-May"
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
os.makedirs(OUT_DIR, exist_ok=True)
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
sub_run = sub.add_run(f"{DATE_LONG}  —  {BIAS} bias  —  Generated {GEN_TIME}")
sub_run.italic = True
sub_run.font.size = Pt(12)

meta = doc.add_paragraph()
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
mr = meta.add_run("Weekly expiry: 0 days (TODAY)  |  Monthly expiry: 7 days (26 May)  |  Expiry week: YES")
mr.bold = True
mr.font.size = Pt(11)

# Badges
badge_line(doc, "TODAY'S BIAS", BIAS, "F1C232")
badge_line(doc, "ENTRY PERMISSION", ENTRY_PERMISSION, "E69138")
badge_line(doc, "CRUDE MODE", CRUDE_MODE, "C00000")
badge_line(doc, "VIX SIZING", VIX_SIZING + " (India VIX surged +4%)", "990000")
badge_line(doc, "EXPIRY ALERT", "WEEKLY EXPIRY TODAY — intraday scalps only, exit ≤ 13:00 IST", "660000")

doc.add_paragraph()

# ---------- SECTION 1 — MACRO SNAPSHOT ----------
add_heading(doc, "Section 1 — Macro Snapshot", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Metric", "Value", "Signal"],
    rows=[
        ["MCX Crude proxy (₹/bbl)", "~₹9,910 (WTI 103.24 × 96.13)", "BEAR band — falling"],
        ["Brent / WTI", "$109.66 / $103.24", "WTI -5% from $108.66 settle (Iran ceasefire)"],
        ["Brent settle (18 May)", "$112.10 (+2%)", "Eased back below $110 after Trump call-off"],
        ["GIFT Nifty", "23,653 (-0.24%)", "Implied gap-up ~+64 pts vs Nifty close"],
        ["S&P 500 (18 May)", "7,403.05 (-0.07%)", "Flat — risk-neutral"],
        ["Nasdaq (18 May)", "26,090.73 (-0.51%)", "Risk-off tilt"],
        ["Dow Jones (18 May)", "49,686.12 (+0.32%)", "Defensive bid"],
        ["US 10Y yield", "Highest in a year", "RISK-OFF for emerging mkts"],
        ["US VIX", "Elevated (n/a exact)", "Iran headline risk"],
        ["India VIX (18 May)", "+4% surge", "QUARTER size mandate"],
        ["Nifty close (18 May)", "₹23,589.25 (-0.23%)", "Range 23,317-23,610"],
        ["Bank Nifty close (18 May)", "53,537 (recovered from -428)", "LATE-SESSION buy, watch lead"],
        ["FII cash (18 May)", "DATA_UNAVAILABLE", "Verify NSE before 9:00 AM"],
        ["DII cash (18 May)", "DATA_UNAVAILABLE", "Verify NSE before 9:00 AM"],
        ["USD-INR", "₹96.13 (+0.15%)", "Rupee weak — FII headwind"],
        ["INFY ADR", "$12.07 flat", "INFY likely flat open"],
        ["WIT ADR", "$1.93 (+2.1%)", "WIPRO likely positive open"],
        ["HDB ADR", "$24.64", "HDFC Bank flat-to-soft open"],
        ["IBN ADR", "$25-26 zone (est.)", "ICICI Bank stable open"],
    ])

doc.add_paragraph()
add_para(doc, "Geopolitical one-liner:", bold=True, size=11)
add_para(doc,
    "Strait of Hormuz still largely blocked since 28 Feb 2026; Trump CALLED OFF Tuesday strike on Iran "
    "at request of Gulf Arab allies; WTI fell ~5% in extended hours; possible US sanctions relief on "
    "Iranian oil exports rumoured to be driving GIFT Nifty rally earlier overnight.", size=11)

add_para(doc, "Intraday headline risk: HIGH — single Iran/Hormuz tweet can whipsaw crude ±5% within an hour.",
    bold=True, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- SECTION 2 — CRUDE RULE + STRUCTURAL READ ----------
doc.add_paragraph()
add_heading(doc, "Section 2 — Crude Rule + Structural Read", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Crude Parameter", "Reading"],
    rows=[
        ["CRUDE_MODE", "BEAR (₹9,000-10,000 band) — puts only, except producers"],
        ["CRUDE_DIRECTION", "FALLING (-5% overnight)"],
        ["CRUDE_FLIP_LEVEL", "MCX ₹9,000 (WTI ~$93.75) → CAUTION/half-size"],
        ["MODE TRANSITION FLAG", "Crude falling toward lower band — WATCH ₹9,000 for mode flip"],
        ["AVOID (today)", "BPCL, HPCL, IOC, IndiGo, Asian Paints, MRF, fertilisers"],
        ["EXCEPTION_CALLS permitted", "ONGC, Oil India, MRPL (crude beneficiaries)"],
    ])

doc.add_paragraph()
add_para(doc, "Structural Read", bold=True, size=12)
add_bullets(doc, [
    "Nifty close vs Max Pain: Max Pain 23,820 (11 May data; refresh at 9:00 AM); Nifty at 23,589 = ~231 pts BELOW — pull UP if max pain holds.",
    "Highest Call OI strike (ceiling): ~₹24,000 zone (last known) — refresh on NSE option chain before entry.",
    "Highest Put OI strike (floor): ~₹23,500 zone — supports near current price.",
    "PCR: DATA_UNAVAILABLE for fresh value — verify on NSE/Sensibull before 9:00 AM.",
    "Days to weekly expiry: 0 (TODAY). Days to monthly expiry: 7 (26 May).",
    "Expiry day dynamics: Max-pain magnetism dominant. Theta decay extreme. NO option BUYING beyond intraday scalp.",
    "Futures basis: DATA_UNAVAILABLE — check NSE near 9:10 AM.",
    "HDFC Bank 770 PE saw heavy 10,000-contract trade on 18 May — bearish hedge/bet positioning.",
])

doc.add_paragraph()
add_para(doc, "Expiry Week — Skin in the Game", bold=True, size=12)
add_bullets(doc, [
    "Put writers' floor likely ₹23,500 (Put OI cluster + recent low at 23,317)",
    "Call writers' ceiling likely ₹23,800-24,000 (Max Pain pull + Call OI)",
    "Magnetism direction: UP toward 23,800 max pain IF gap-up holds first 30 min",
    "Range to expect: 23,400 — 23,800 unless Iran headline breaks it",
])

doc.add_paragraph()
add_para(doc, "Bank Nifty Leadership Check", bold=True, size=12)
add_bullets(doc, [
    "Bank Nifty opened 428 pts WEAK then recovered to 53,537 — V-shaped reversal = potential leadership",
    "If Bank Nifty leads Nifty up today → move is REAL (financials buying = institutional)",
    "If Bank Nifty lags despite gap-up → move is SUSPECT — fade calls, prefer puts",
    "HDFC Bank near 52w low ₹726; recovery off lows yesterday but ADR weak overnight — MIXED signal",
])

# ---------- SECTION 3 — CONTRARIAN CHECK ----------
doc.add_paragraph()
add_heading(doc, "Section 3 — Contrarian Check (Layer 3)", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_kv_table(doc,
    header=["Q", "Read"],
    rows=[
        ["Q1. CONSENSUS",
         "Retail expects: crude high → bearish for OMC/aviation; Iran headlines = sell rally; "
         "rupee at 96 = FII outflows; near 52w lows in HDFCBANK/RELIANCE/INFY = more downside."],
        ["Q2. THE TRAP",
         "Trump CALLED OFF Tuesday Iran strike → crude crashing -5% overnight. OMCs/aviation could "
         "gap UP hard; oil producers (ONGC) face profit-taking. The 'short OMC, long ONGC' "
         "consensus trade is the trap today."],
        ["Q3. RETAIL STOPS",
         "Long stops below 23,400 (recent low cushion). Short stops above 23,800 (max pain magnet). "
         "BPCL short stops above ₹290; ONGC long stops below ₹290."],
        ["Q4. FLIP TRIGGER",
         "WTI breaking $95 → MCX crude prints ₹9,000 → mode shifts to CAUTION → OMC short-covering "
         "wave; ONGC sell-off. OR formal Iran sanctions-relief headline → both happen simultaneously."],
    ])

doc.add_paragraph()
add_para(doc, "CONTRARIAN OVERRIDE LOGIC", bold=True, size=12, color=RGBColor(0xC0, 0x00, 0x00))
add_kv_table(doc,
    header=["Condition", "Status"],
    rows=[
        ["A: GIFT Nifty gap-up > +100 pts?", "NO (+64 pts implied)"],
        ["B: Crude falling > 2% overnight?", "YES (-5%)"],
        ["C: Contrarian scenario probability > 40%?", "YES (Iran ceasefire + sanctions relief rumours)"],
        ["D: Expiry day + PCR < 0.8?", "Expiry YES; PCR DATA_UNAVAILABLE — assume neutral"],
    ])

add_para(doc,
    "⚠ A + B not both YES (gap +64 below threshold). Override DOES NOT force BEAR-bias suppression. "
    "However B + C combine to lock confidence at LOW. Bias kept at RANGE with downside-fade tactical lean "
    "ONLY when ₹23,800 max-pain ceiling holds.",
    bold=True, italic=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))

# ---------- SECTION 4 — STOCK SETUPS ----------
doc.add_paragraph()
add_heading(doc, "Section 4 — 5 Individual F&O Setups", level=1, color=RGBColor(0x1F, 0x38, 0x64))

add_para(doc,
    "F&O Ban list (15 May data): KAYNES, SAIL. Refresh NSE ban list before 9:00 AM — replace any setup that gets banned.",
    italic=True, size=10, color=RGBColor(0xC0, 0x00, 0x00))

# Stock 1 — HDFC BANK PUT
add_para(doc, "1. HDFC BANK (NSE: HDFCBANK)  —  PUT  —  Grade A", bold=True, size=13,
         color=RGBColor(0xC0, 0x00, 0x00))
stock_card(doc, [
    ("CMP", "₹767.50  |  Lot: 550  |  Crude aligned: NEUTRAL"),
    ("Catalyst", "HDB ADR weak overnight at $24.64; ₹770 PE saw 10,000-contract block trade on 18 May "
                 "(bearish hedge/positioning); FII Q4 sell-off concentrated in HDFC Bank (₹41,449 cr)."),
    ("Technical", "Daily DOWN. BELOW 200 DMA (key institutional bearish signal). Weekly DOWN. "
                  "52w high ₹1,020 / 52w low ₹726 (only ~₹41 above floor). Support ₹751 / Resistance ₹780."),
    ("Structural", "OI direction: SHORT BUILD (770 PE block). Delivery %: STABLE. Block deal: YES (770 PE). "
                   "Ban list: NO."),
    ("Option Setup", "770 PE MONTHLY (26 May expiry) — Strike ₹770 ATM. ⚠ THETA WARNING: NO weekly options "
                     "today (expiry). Monthly only."),
    ("Entry Trigger",
     "ALL 4 must fire: SuperTrend RED on 15-min confirmed; RSI(14) below 50 on 15-min; "
     "StochRSI cross DOWN from >70; Volume > 1.5× 20-period avg; Confirmation: 15-min close BELOW ₹765."),
    ("T1 / T2", "T1 ₹755 (book 40%)  |  T2 ₹745 (book 40%)"),
    ("Stop Loss", "₹775 + 15-min close above 50 EMA"),
    ("R:R", "≥ 2.3:1 (₹10 risk / ₹22.5 reward at T2)"),
    ("Time Stop", "Exit ALL by 13:00 IST (expiry-day hard rule)"),
    ("Grade Reason", "A — strongest structural signal (block PE trade + near 52w low + FII selling pressure)"),
])

# Stock 2 — RELIANCE PUT
doc.add_paragraph()
add_para(doc, "2. RELIANCE INDUSTRIES (NSE: RELIANCE)  —  PUT  —  Grade A", bold=True, size=13,
         color=RGBColor(0xC0, 0x00, 0x00))
stock_card(doc, [
    ("CMP", "₹1,338  |  Lot: 500  |  Crude aligned: PARTIAL (refining loses on falling crude)"),
    ("Catalyst", "Bearish technical bias confirmed for week 19-23 May; momentum score 31.6/100; "
                 "stock -6.90% week, -12% over 6m. Only ₹48 above 52w low ₹1,290 — high break-down risk."),
    ("Technical", "Daily DOWN. BELOW 200 DMA. Weekly DOWN. Support ₹1,290 / Resistance ₹1,365. "
                  "Range yesterday: ₹1,329-1,365."),
    ("Structural", "OI direction: SHORT BUILD likely. Refining margins compress as Brent falls "
                   "(petchem/refining 60%+ of revenue). Ban list: NO."),
    ("Option Setup", "1340 PE MONTHLY (26 May). ATM. Avoid weekly (expiry today)."),
    ("Entry Trigger",
     "SuperTrend RED on 15-min; RSI < 50 on 15-min; StochRSI cross DOWN from >70; "
     "Volume > 1.5× avg; Confirmation: 15-min close BELOW ₹1,330."),
    ("T1 / T2", "T1 ₹1,315  |  T2 ₹1,295"),
    ("Stop Loss", "₹1,350 + 15-min close above VWAP"),
    ("R:R", "≥ 2.1:1"),
    ("Time Stop", "Exit by 13:00 IST"),
    ("Grade Reason", "A — multiple confirming negatives (technical, momentum, refining headwind)"),
])

# Stock 3 — INFOSYS CALL
doc.add_paragraph()
add_para(doc, "3. INFOSYS (NSE: INFY)  —  CALL  —  Grade B+", bold=True, size=13,
         color=RGBColor(0x38, 0x76, 0x1D))
stock_card(doc, [
    ("CMP", "₹1,142.50  |  Lot: 400  |  Crude aligned: NEUTRAL (IT services, USD revenue)"),
    ("Catalyst", "Nifty IT closed higher 2nd straight session — all 10 constituents green on 18 May. "
                 "INFY at 52w low zone (₹1,089) — oversold bounce setup. INFY ADR flat at $12.07; "
                 "WIT ADR up +2.1% — IT sector tailwind. Weak rupee (96.13) = export tailwind."),
    ("Technical", "Daily intraday UPTREND noted. BELOW 50 DMA (₹1,247) and BELOW 200 DMA (₹1,454) — "
                  "structurally weak but oversold; bounce trade only. Support ₹1,128 / Resistance ₹1,165."),
    ("Structural", "OI direction: SHORT COVERING risk in IT. Delivery %: rising on bounce. Ban list: NO."),
    ("Option Setup", "1140 CE or 1150 CE MONTHLY (26 May). Avoid weekly (expiry today)."),
    ("Entry Trigger",
     "SuperTrend GREEN on 15-min; RSI > 50 on 15-min; StochRSI cross UP from <30; "
     "Volume > 1.5× avg; Confirmation: 15-min close ABOVE ₹1,148."),
    ("T1 / T2", "T1 ₹1,158  |  T2 ₹1,168"),
    ("Stop Loss", "₹1,135 + 15-min close below VWAP"),
    ("R:R", "≥ 2.0:1"),
    ("Time Stop", "Exit by 13:00 IST"),
    ("Grade Reason", "B+ — counter-trend bounce trade; IT sector strong but stock structurally weak"),
])

# Stock 4 — ONGC CALL
doc.add_paragraph()
add_para(doc, "4. ONGC (NSE: ONGC)  —  CALL  —  Grade B", bold=True, size=13,
         color=RGBColor(0x38, 0x76, 0x1D))
stock_card(doc, [
    ("CMP", "₹297.45  |  Lot: 2,250  |  Crude aligned: YES (mandatory crude-beneficiary call in BEAR mode)"),
    ("Catalyst", "Crude FALLING overnight is HEADWIND — but stock has +7.24% 1-month momentum and is "
                 "₹10 below 52w high ₹307.50. Required crude-beneficiary slot per BEAR-mode rule."),
    ("Technical", "Daily UP. ABOVE 200 DMA. Weekly UP. Support ₹290 / Resistance ₹305. "
                  "Tight base 290-300."),
    ("Structural", "OI direction: LONG BUILD into Hormuz crisis. Royalty cuts boost realisations. Ban list: NO."),
    ("Option Setup", "300 CE MONTHLY (26 May). Slightly OTM."),
    ("Entry Trigger",
     "SuperTrend GREEN on 15-min; RSI > 50 on 15-min; StochRSI cross UP from <30; "
     "Volume > 1.5× avg; Confirmation: 15-min close ABOVE ₹300 AND WTI HOLDS > $100. "
     "If WTI breaks $98 → SKIP TRADE (crude-fall risk dominates)."),
    ("T1 / T2", "T1 ₹304  |  T2 ₹308"),
    ("Stop Loss", "₹295 + WTI breaking $98"),
    ("R:R", "≥ 2.0:1"),
    ("Time Stop", "Exit by 13:00 IST"),
    ("Grade Reason", "B — required slot but trade quality degraded by overnight crude crash; "
                     "conditional entry only"),
])

# Stock 5 — TCS PUT
doc.add_paragraph()
add_para(doc, "5. TATA STEEL (NSE: TATASTEEL)  —  CALL  —  Grade B  [LIMITED LIVE DATA]", bold=True, size=13,
         color=RGBColor(0x38, 0x76, 0x1D))
stock_card(doc, [
    ("CMP", "DATA_UNAVAILABLE — verify on NSE before 9:00 AM"),
    ("Catalyst", "Metals (YTD leader sector). Defensive choice in absence of cleaner mid-cap setups; "
                 "China stimulus tailwind ongoing. Filler position — not core conviction."),
    ("Technical", "DATA_UNAVAILABLE on intraday levels — verify before entry."),
    ("Structural", "Ban list: NO (verify in morning refresh)."),
    ("Option Setup", "ATM CE MONTHLY (26 May)."),
    ("Entry Trigger",
     "All 4 standard signals must fire AND 15-min close above prior day high. "
     "Skip if any single check fails — this is the lowest-conviction setup of the five."),
    ("T1 / T2", "Set after price discovery — 1% / 2% targets from breakout"),
    ("Stop Loss", "Below 15-min entry candle low"),
    ("R:R", "Compute live — skip if not ≥ 2:1"),
    ("Time Stop", "Exit by 13:00 IST"),
    ("Grade Reason", "B — placeholder/diversifier; flagged DATA_UNAVAILABLE; "
                     "trade ONLY if morning refresh confirms ALL filters"),
])

# ---------- SECTION 5 — FINAL VERDICT ----------
doc.add_paragraph()
add_heading(doc, "Section 5 — Final Verdict", level=1, color=RGBColor(0xC0, 0x00, 0x00))

add_kv_table(doc,
    header=["Parameter", "Reading"],
    rows=[
        ["TODAY'S BIAS", "RANGE (fade extremes; 23,400-23,800 expected)"],
        ["CONFIDENCE", "LOW (expiry day + crude crashing + Iran headline risk)"],
        ["CRUDE MODE", "BEAR (₹9,910 proxy) — direction FALLING"],
        ["VIX SIZING", "QUARTER (India VIX +4% on 18 May; expiry-day vol spike likely)"],
        ["ENTRY PERMISSION", "YELLOW"],
        ["Max trades today", "3 (cap absolute)"],
        ["Daily risk cap", "2% of capital"],
        ["Hard exit", "13:00 IST (expiry-day rule overrides 13:30 default)"],
        ["No new entries after", "12:00 IST (expiry-day conservative)"],
    ])

doc.add_paragraph()
add_para(doc, "THE BULL CASE (2 lines):", bold=True, size=11)
add_para(doc,
    "If Iran sanctions relief becomes formal news intraday, OMCs/aviation rip higher, "
    "USD-INR firms below 96, and Nifty grinds back to max-pain 23,820. Banking leadership confirms "
    "the move (Bank Nifty reclaims 54,000), short-covering wave in HDFCBANK above ₹780.", size=11)

add_para(doc, "THE BEAR CASE (2 lines):", bold=True, size=11)
add_para(doc,
    "If Iran rejects ceasefire / new Hormuz incident, crude spikes 5%+ back to $108, USD-INR breaches 96.50, "
    "Nifty rejects 23,650 and slices through 23,400 toward put-writer floor 23,250. "
    "RELIANCE breaks ₹1,329 day-low → ₹1,295 cascade.", size=11)

add_para(doc, "FLIP TRIGGER:", bold=True, size=11, color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "If WTI breaks $95 (= MCX ₹9,000) → crude mode flips BEAR→CAUTION; bias upgrades from RANGE to "
    "CAUTIOUS BULL; ONGC long → CLOSE; BPCL short-covering wave (off the avoid list as mode shifts).",
    bold=True, size=11)

doc.add_paragraph()
add_para(doc, "NIFTY KEY LEVELS", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_kv_table(doc,
    rows=[
        ["S2", "₹23,250 (put-writer defended floor)"],
        ["S1", "₹23,400 (recent intraday low cushion)"],
        ["CRITICAL", "₹23,589 (yesterday close — bias line)"],
        ["R1", "₹23,800 (max-pain magnet zone)"],
        ["R2", "₹24,000 (call-writer ceiling)"],
    ])

doc.add_paragraph()
add_para(doc, "TOP 3 RANKED SETUPS", bold=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))
add_bullets(doc, [
    "#1 HDFCBANK PE — A — Block PE trade + ADR weak + 52w low proximity — entry below ₹765",
    "#2 RELIANCE PE — A — Momentum bearish + crude-falling pressures refining — entry below ₹1,330",
    "#3 INFY CE — B+ — Oversold bounce + IT sector strength + weak rupee tailwind — entry above ₹1,148",
])

doc.add_paragraph()
add_para(doc, "ONE RISK THAT RUINS EVERYTHING TODAY", bold=True, size=12,
         color=RGBColor(0xC0, 0x00, 0x00))
add_para(doc,
    "A fresh Iran/Hormuz headline (renewed strike, tanker attack, sanctions deal collapse) — crude "
    "whiplashes ±7% in minutes, every setup invalidated, USD-INR gaps to 97, both bull and bear "
    "stops triggered. On expiry day with QUARTER sizing, sit out for 15 min and re-read.", size=11)

# ---------- ONE-LINE SUMMARY ----------
doc.add_paragraph()
add_heading(doc, "ONE-LINE SUMMARY (read at 9:10 AM)", level=2, color=RGBColor(0x1F, 0x38, 0x64))
add_para(doc,
    "\"Today is RANGE because expiry day + crude crashing -5% overnight + Iran headline overhang lock "
    "confidence at LOW. Crude ₹9,910 = BEAR mode but FALLING. Watch HDFCBANK PE and RELIANCE PE on "
    "the downside, INFY CE as oversold bounce. Key risk: any Iran/Hormuz headline. "
    "Size QUARTER. Flips CAUTIOUS BULL if WTI breaks $95.\"",
    italic=True, size=12, color=RGBColor(0x1F, 0x38, 0x64))

# ---------- FOOTER ----------
doc.add_paragraph()
doc.add_paragraph()
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = foot.add_run("Generated by Daily Trading Routine — for personal educational use only. Not financial advice. "
                  "Verify all data against NSE official before market open.")
fr.italic = True
fr.font.size = Pt(9)
fr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
