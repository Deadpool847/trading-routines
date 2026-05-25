"""Generate the daily F&O trading-intelligence .docx brief for 2026-05-25."""
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-25"
BIAS = "RANGE"
OUT_PATH = f"trading-briefs/2026/05-May/Trading_Brief_{DATE_ISO}_{BIAS}.docx"

NAVY = RGBColor(0x1F, 0x38, 0x64)
RED = RGBColor(0xC0, 0x00, 0x00)
GREEN = RGBColor(0x1E, 0x7A, 0x33)
AMBER = RGBColor(0xB8, 0x6A, 0x00)
GREY = RGBColor(0x55, 0x55, 0x55)


def shade(cell, hex_fill):
    pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    pr.append(shd)


def heading(doc, text, level=1, color=NAVY):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.color.rgb = color
    return h


def para(doc, text, bold=False, italic=False, size=10.5, color=None, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return p


def bullets(doc, items):
    for it in items:
        doc.add_paragraph(it, style="List Bullet")


def table(doc, header, rows):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Light Grid Accent 1"
    for i, txt in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        run = c.paragraphs[0].add_run(txt)
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(c, "1F3864")
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run(str(val))
            run.font.size = Pt(9.5)
    return t


doc = Document()
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(10.5)

# ===== TITLE =====
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
tr = title.add_run("ELITE DAILY TRADING BRIEF")
tr.bold = True
tr.font.size = Pt(20)
tr.font.color.rgb = NAVY
sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sr = sub.add_run(f"{DATE_ISO} (Monday)  |  BIAS: {BIAS}  |  NSE F&O")
sr.bold = True
sr.font.size = Pt(13)
sr.font.color.rgb = AMBER
meta = doc.add_paragraph()
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
mr = meta.add_run("Generated 08:50 IST  |  Weekly expiry: 1 day (Tue 26 May)  |  "
                  "Monthly expiry: 1 day (May) / 36 days (Jun 30)  |  Expiry week: YES")
mr.italic = True
mr.font.size = Pt(9)
mr.font.color.rgb = GREY

# ===== DASHBOARD BOX =====
dash = table(doc, ["FIELD", "READING"], [
    ["BIAS", "RANGE  (gap-up fade risk into expiry; explosive headline tail)"],
    ["ENTRY PERMISSION", "YELLOW  (expiry-1, half size, Iran headline risk HIGH)"],
    ["CRUDE MODE", "BEAR  (MCX ~Rs 9,160/bbl)  |  Direction: FLAT-to-FIRM (elevated)"],
    ["VIX SIZING", "HALF SIZE  (India VIX ~18.7, band 15-20)"],
    ["EXPIRY ALERT", "Tomorrow Tue-26-May is weekly+monthly expiry. Max-pain magnetism active; "
                     "buy options on JUNE (30-Jun) series only to dodge 1-day theta cliff."],
    ["CONFIDENCE", "MEDIUM"],
])

# ===== SECTION 1 =====
heading(doc, "SECTION 1 - MACRO SNAPSHOT", 1)
table(doc, ["Metric", "Value", "Signal"], [
    ["MCX Crude (proxy)", "~Rs 9,160/bbl", "BEAR band (9,000-10,000)"],
    ["Brent / WTI", "$100.21 / $96.60", "Elevated; +0.94% / +0.26%"],
    ["GIFT Nifty gap", "+249 pts (23,970 vs 23,719)", "Bullish open, +1.05%"],
    ["S&P 500", "DATA_UNAVAILABLE (close)", "US-Iran peace hope = risk-on tone"],
    ["US VIX", "DATA_UNAVAILABLE", "—"],
    ["India VIX", "~18.7", "HALF SIZE (mandatory)"],
    ["Nifty close (22-May)", "Rs 23,719.30 (+0.27%)", "Below max pain 23,850"],
    ["FII cash", "-Rs 1,891 Cr (21-May)", "SELL (MTD -27,788 Cr)"],
    ["DII cash", "+Rs 2,492 Cr (21-May)", "BUY (MTD +50,862 Cr)"],
    ["USD-INR", "~94.84", "WEAK rupee (-11% YTD)"],
    ["INFY ADR", "$12.64 (last)", "% move DATA_UNAVAILABLE"],
    ["HDB ADR", "$24.62 (last)", "% move DATA_UNAVAILABLE"],
    ["IBN ADR", "$25.89 (last)", "% move DATA_UNAVAILABLE"],
])
para(doc, "Geopolitical one-liner: Strait of Hormuz crisis ongoing (largely blocked since 28-Feb-2026); "
          "GIFT Nifty gapping up on US-Iran peace-talk optimism, but Iran reportedly keeping enriched "
          "uranium in-country = talks 'complicated'. Fragile.", bold=False)
para(doc, "Today's intraday headline risk: HIGH  (direction of risk = crude SPIKE / equity gap-fill if talks wobble).",
     bold=True, color=RED)

# ===== SECTION 2 =====
heading(doc, "SECTION 2 - CRUDE RULE + STRUCTURAL READ", 1)
para(doc, "CRUDE MODE: BEAR (puts only).  MCX proxy = WTI $96.60 x USD-INR 94.84 ~ Rs 9,160/bbl "
          "(weak rupee pushes rupee-crude into bear band even with WTI sub-$100).", bold=True, color=RED)
bullets(doc, [
    "CRUDE_DIRECTION: FLAT-to-FIRM. Brent eased from Fri ~$105 spike to ~$100 but ticking up today; NOT falling >2%.",
    "CRUDE_FLIP_LEVEL: Rs 9,000 (below -> CAUTION) ; Rs 8,500 (mild-bull) ; Rs 10,000 (STRONG BEAR, no calls).",
    "AVOID: OMCs (BPCL/HPCL/IOC), aviation, paints, tyres, fertilisers.",
    "EXCEPTION_CALLS permitted: ONGC, Oil India, MRPL (upstream beneficiaries) + defence (BEL).",
])
para(doc, "Structural read:", bold=True)
bullets(doc, [
    "Nifty vs max pain: close 23,719 is ~131 pts BELOW max pain 23,850 -> magnetism pulls UP toward 23,850; "
    "but GIFT 23,970 is ABOVE max pain -> gap-up likely fades back toward 23,850 into expiry.",
    "PCR 1.35 (>1.3). Expiry-week interpretation: PUT WRITERS DEFENDING FLOOR (bullish cushion, buy-on-dips). NOT a squeeze setup.",
    "OI change: puts being WRITTEN at 23,500 (bullish floor); call resistance building near 24,000 (ceiling).",
    "Futures basis: DATA_UNAVAILABLE (treat neutral).",
    "Expiry-week dynamics: Put writers' floor Rs 23,500 ; Call writers' ceiling ~Rs 24,000 ; max-pain magnetism -> ~23,850.",
    "Put side has more skin to defend (PCR 1.35) -> downside cushioned unless an Iran headline breaks the floor.",
])
para(doc, "Bank Nifty leadership check: ICICI Bank trades BELOW both 50-DMA (1,278) and 200-DMA (1,361) -> "
          "banks are LAGGING. A bank-light gap-up is SUSPECT -> fade-prone. (Live Bank Nifty print DATA_UNAVAILABLE.)",
     bold=True, color=AMBER)

# ===== SECTION 3 =====
heading(doc, "SECTION 3 - CONTRARIAN CHECK (Layer 3)", 1)
bullets(doc, [
    "Q1 CONSENSUS: 'Peace deal = gap-and-go. Buy the open, crude crashes, Nifty 24,000+.' TV is risk-on.",
    "Q2 THE TRAP: Talks are NOT signed (Iran kept uranium). Gap-up into expiry with banks lagging fades to "
    "max pain 23,850; option buyers bleed theta on a pin. Sellers/writers profit.",
    "Q3 RETAIL STOPS: Long stops clustered below 23,800 (gap-fill) and 23,700 (Fri close). Short stops above 24,020.",
    "Q4 FLIP TRIGGER: A single Hormuz/Iran headline. Peace confirmed -> explosive BULL >24,000. "
    "Talks collapse / tanker incident -> crude re-spikes, RANGE breaks BEAR through 23,700.",
])
heading(doc, "Contrarian Override Logic", 2)
table(doc, ["Condition", "Status"], [
    ["A: GIFT gap-up > +100 pts?", "YES (+249)"],
    ["B: Crude falling > 2% overnight?", "NO (flat-to-firm)"],
    ["C: Contrarian (fade) probability > 40%?", "YES (~45%, expiry pin + banks lagging)"],
    ["D: Expiry week + PCR < 0.8 (squeeze)?", "NO (PCR 1.35)"],
])
bullets(doc, [
    "A+B both YES? NO -> bear-suppression override NOT triggered.",
    "C+D both YES? NO -> no forced squeeze warning.",
    "A+B+C all YES? NO -> no full override.",
    "NET: No override fires. Bias stands on merit = RANGE with explosive headline tail. "
    "Confidence held at MEDIUM (C alone keeps conviction capped).",
])

# ===== SECTION 4 =====
heading(doc, "SECTION 4 - 5 INDIVIDUAL F&O SETUPS", 1)
para(doc, "Lot sizes are approximate NSE values - VERIFY on NSE before sizing. All option buys use the "
          "JUNE (30-Jun-2026) series: with weekly+monthly expiry tomorrow, the May series is a 1-day theta trap. "
          "Per rule (days-to-expiry <= 3 -> monthly only), and May monthly dies tomorrow, June is the valid monthly.",
     italic=True, color=GREY)


def setup(doc, header_line, cmp_lot, catalyst, technical, structural, option, entry, targets, grade_reason,
          grade_color=NAVY):
    p = doc.add_paragraph()
    r = p.add_run(header_line)
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = grade_color
    para(doc, cmp_lot, bold=True, size=10)
    para(doc, "CATALYST: " + catalyst, size=10)
    para(doc, "TECHNICAL:", bold=True, size=10)
    bullets(doc, technical)
    para(doc, "STRUCTURAL:", bold=True, size=10)
    bullets(doc, structural)
    para(doc, "OPTION SETUP:", bold=True, size=10)
    bullets(doc, option)
    para(doc, "ENTRY (ALL 4 must fire):", bold=True, size=10)
    bullets(doc, entry)
    para(doc, targets, bold=True, size=10, color=GREEN)
    para(doc, "GRADE REASON: " + grade_reason, italic=True, size=10)
    doc.add_paragraph("———————————————————————————————")


# --- 1. INDIGO PUT ---
setup(doc,
      "1) INTERGLOBE / INDIGO  -  NSE: INDIGO  |  PUT  |  Grade A",
      "CMP: Rs 4,411.60  |  Lot: ~150 (verify NSE)  |  Crude aligned: YES (crude victim)",
      "ATF up ~130% MoM; jet fuel >Rs 1 lakh/KL; fuel = 35-40% of opex; Mideast crisis denting margins. "
      "Pure BEAR-crude-mode play.",
      ["Daily trend: DOWN (-24.5% over 6M)",
       "Above 200 DMA: NO (well below)",
       "Weekly trend: DOWN",
       "Support Rs 4,200 / 3,895 (52w low) | Resistance Rs 4,500"],
      ["OI direction: SHORT BUILD bias (fuel-cost overhang)",
       "Delivery %: DATA_UNAVAILABLE",
       "Block deal: NO | Ban list: NO"],
      ["Strike: Rs 4,400 PE [ATM] | Expiry: 30-Jun-2026",
       "THETA WARNING (expiry <= 3d on May series): YES if May used -> use JUNE"],
      ["ST RED on 15-min confirmed",
       "RSI below 50 on 15-min",
       "StochRSI cross DOWN from >70",
       "Volume > 1.5x 20-period avg",
       "Confirmation candle: 15-min close BELOW Rs 4,380"],
      "T1 (book 40%): Rs 4,280  |  T2 (book 40%): Rs 4,180  |  SL: Rs 4,480 (15-min close above) "
      "|  R:R ~2.3:1  |  Time stop: exit by 13:00 IST",
      "Cleanest crude-victim short in BEAR mode; structural downtrend + macro fuel shock aligned.",
      grade_color=RED)

# --- 2. BEL CALL ---
setup(doc,
      "2) BHARAT ELECTRONICS  -  NSE: BEL  |  CALL  |  Grade A",
      "CMP: Rs 416.55  |  Lot: ~5,700 (verify NSE)  |  Crude aligned: YES (geopolitical/defence beneficiary)",
      "Hormuz/Iran tension = defence tailwind. Strong Q4 (21-May, +2%); order book Rs 73,882 Cr; "
      "fresh orders (GBMES Rs 1,251 Cr).",
      ["Daily trend: UP (consolidating off 429)",
       "Above 200 DMA: YES (52w 322-473, CMP 416)",
       "Weekly trend: UP",
       "Support Rs 400 / 390 | Resistance Rs 430 / 473 (52w high)"],
      ["OI direction: LONG BUILD on order-flow news",
       "Delivery %: DATA_UNAVAILABLE",
       "Block deal: NO | Ban list: NO"],
      ["Strike: Rs 420 CE [ATM] | Expiry: 30-Jun-2026",
       "THETA WARNING: YES if May used -> use JUNE"],
      ["ST GREEN on 15-min confirmed",
       "RSI above 50 on 15-min",
       "StochRSI cross UP from <30",
       "Volume > 1.5x 20-period avg",
       "Confirmation candle: 15-min close ABOVE Rs 420"],
      "T1 (book 40%): Rs 428  |  T2 (book 40%): Rs 438  |  SL: Rs 408 (15-min close below) "
      "|  R:R ~2.5:1  |  Time stop: exit by 13:00 IST",
      "Defence beneficiary of live geopolitics; uptrend + order momentum. Caveat: rich P/E ~50.",
      grade_color=GREEN)

# --- 3. ONGC CALL ---
setup(doc,
      "3) ONGC  -  NSE: ONGC  |  CALL  |  Grade A",
      "CMP: Rs 296.45  |  Lot: ~3,850 (verify NSE)  |  Crude aligned: YES (upstream beneficiary)",
      "Upstream producer gains on higher crude realisation. Q4 earnings TOMORROW (26-May) = catalyst AND event "
      "risk -> intraday only, exit before close, NO overnight.",
      ["Daily trend: UP (+7.24% 1M, near 52w high 307.5)",
       "Above 200 DMA: YES (near 52w high)",
       "Weekly trend: UP",
       "Support Rs 282 / 278 | Resistance Rs 307 (52w high)"],
      ["OI direction: LONG BUILD (crude beneficiary bid)",
       "Delivery %: DATA_UNAVAILABLE",
       "Block deal: NO | Ban list: NO"],
      ["Strike: Rs 300 CE [ATM] | Expiry: 30-Jun-2026",
       "THETA WARNING: YES if May used -> use JUNE; avoid holding into 26-May earnings"],
      ["ST GREEN on 15-min confirmed",
       "RSI above 50 on 15-min",
       "StochRSI cross UP from <30",
       "Volume > 1.5x 20-period avg",
       "Confirmation candle: 15-min close ABOVE Rs 298"],
      "T1 (book 40%): Rs 304  |  T2 (book 40%): Rs 310  |  SL: Rs 291 (15-min close below) "
      "|  R:R ~2.0:1  |  Time stop: exit by 13:00 IST (pre-earnings)",
      "Crude-beneficiary call legal in BEAR mode; trend strong, but earnings-eve adds event risk -> A not A+.",
      grade_color=GREEN)

# --- 4. BPCL PUT ---
setup(doc,
      "4) BHARAT PETROLEUM  -  NSE: BPCL  |  PUT  |  Grade B",
      "CMP: Rs 295.20  |  Lot: ~1,800 (verify NSE)  |  Crude aligned: YES (OMC crude victim)",
      "OMC margin crush: LPG under-recovery ~Rs 670/cyl (May), petrol/diesel marketing losses Rs 10-20/litre. "
      "S&P flags margin pressure from oil spike.",
      ["Daily trend: DOWN (-19% 6M) but +5% post-Q4 bounce (21-May)",
       "Above 200 DMA: NO",
       "Weekly trend: DOWN",
       "Support Rs 285 / 266 (52w low) | Resistance Rs 300"],
      ["OI direction: SHORT BUILD into rallies",
       "Delivery %: DATA_UNAVAILABLE",
       "Block deal: NO | Ban list: NO"],
      ["Strike: Rs 290 PE [ATM] | Expiry: 30-Jun-2026",
       "THETA WARNING: YES if May used -> use JUNE"],
      ["ST RED on 15-min confirmed",
       "RSI below 50 on 15-min",
       "StochRSI cross DOWN from >70",
       "Volume > 1.5x 20-period avg",
       "Confirmation candle: 15-min close BELOW Rs 292"],
      "T1 (book 40%): Rs 286  |  T2 (book 40%): Rs 278  |  SL: Rs 299 (15-min close above) "
      "|  R:R ~2.0:1  |  Time stop: exit by 13:00 IST",
      "Macro crude-victim thesis strong, but recent post-Q4 bounce = chop risk -> B until ST flips RED.",
      grade_color=RED)

# --- 5. ICICI BANK CALL ---
setup(doc,
      "5) ICICI BANK  -  NSE: ICICIBANK  |  CALL  |  Grade B",
      "CMP: Rs 1,264.30  |  Lot: ~700 (verify NSE)  |  Crude aligned: N/A (financial)",
      "Gap-up / risk-on proxy + Bank Nifty leadership tell. IBN ADR $25.89. Tactical bounce play only.",
      ["Daily trend: SIDEWAYS-up (intraday uptrend)",
       "Above 200 DMA: NO (200=1,361; 50=1,278; CMP below both)",
       "Weekly trend: SIDEWAYS",
       "Support Rs 1,242 / 1,188 (52w low) | Resistance Rs 1,278 (50-DMA) / 1,300"],
      ["OI direction: NEUTRAL",
       "Delivery %: DATA_UNAVAILABLE",
       "Block deal: NO | Ban list: NO"],
      ["Strike: Rs 1,280 CE [ATM/OTM] | Expiry: 30-Jun-2026",
       "THETA WARNING: YES if May used -> use JUNE"],
      ["ST GREEN on 15-min confirmed",
       "RSI above 50 on 15-min",
       "StochRSI cross UP from <30",
       "Volume > 1.5x 20-period avg",
       "Confirmation candle: 15-min close ABOVE Rs 1,270 (reclaim 50-DMA)"],
      "T1 (book 40%): Rs 1,284  |  T2 (book 40%): Rs 1,298  |  SL: Rs 1,256 (15-min close below) "
      "|  R:R ~2.4:1  |  Time stop: exit by 13:00 IST",
      "Below 200-DMA = institutional headwind + gap-fade risk; only a YELLOW-permission tactical long -> B.",
      grade_color=NAVY)

# ===== SECTION 5 =====
heading(doc, "SECTION 5 - FINAL VERDICT", 1)
table(doc, ["Field", "Call"], [
    ["TODAY'S BIAS", "RANGE (gap-up fade into expiry; explosive headline tail)"],
    ["CONFIDENCE", "MEDIUM"],
    ["CRUDE MODE", "BEAR (~Rs 9,160/bbl, elevated)"],
    ["VIX SIZING", "HALF SIZE"],
    ["ENTRY PERMISSION", "YELLOW"],
])
para(doc, "THE BULL CASE: PCR 1.35 with put writers defending 23,500; DII buying hard (+50,862 MTD); "
          "peace-deal headline could trigger explosive break above 24,000.", color=GREEN)
para(doc, "THE BEAR CASE: Banks lagging (ICICI below 200-DMA), FII selling (-27,788 MTD), elevated crude, "
          "expiry pin to 23,850 below GIFT -> gap-up fades; one Iran headline breaks 23,700.", color=RED)
para(doc, 'FLIP TRIGGER: "If Nifty 15-min closes below 23,700, bias becomes BEAR (gap-fill, target 23,500 floor). '
          'If a confirmed US-Iran peace headline hits AND Nifty holds above 24,020, bias becomes EXPLOSIVE BULL."',
     bold=True)
para(doc, "NIFTY KEY LEVELS:  S2 23,500  |  S1 23,700  |  CRITICAL 23,850 (max pain)  |  R1 24,000  |  R2 24,200",
     bold=True)
para(doc, "TOP 3 RANKED:", bold=True)
bullets(doc, [
    "#1 INDIGO - A - cleanest crude-victim short, structural downtrend - entry PUT below Rs 4,380",
    "#2 BEL - A - defence beneficiary of live geopolitics, uptrend - entry CALL above Rs 420",
    "#3 ONGC - A - upstream crude beneficiary (intraday only, earnings-eve) - entry CALL above Rs 298",
])
para(doc, "ONE RISK THAT RUINS EVERYTHING TODAY: A Strait-of-Hormuz escalation or collapse of US-Iran talks "
          "mid-session -> crude re-spikes through $105, equity gap fills violently, and even the 'safe' put-floor "
          "23,500 gives way. Headline-driven, un-hedgeable intraday.", bold=True, color=RED)

# ===== ONE-LINE SUMMARY =====
heading(doc, "ONE-LINE SUMMARY (read at 9:10 AM)", 1)
para(doc,
     '"Today is RANGE because a +249 gap-up on fragile US-Iran peace hope runs into expiry-day-minus-1 '
     'magnetism toward max pain 23,850 with banks lagging. Crude ~Rs 9,160 = BEAR, elevated/flat. '
     'Watch INDIGO PUT and BEL CALL. Key risk: a Hormuz headline re-spiking crude. Size HALF. '
     'Flips BEAR if Nifty 15-min closes below 23,700."',
     bold=True, size=11, color=NAVY)

# ===== DISCLAIMER =====
doc.add_paragraph()
para(doc, "Auto-generated pre-market intelligence brief. Educational/analytical use only; not investment advice. "
          "Figures sourced from public web search at ~08:50 IST 2026-05-25 and may be intraday-stale; verify live "
          "before execution. Fields marked DATA_UNAVAILABLE were not reliably resolvable at generation time.",
     italic=True, size=8, color=GREY)

doc.save(OUT_PATH)
print("Saved:", OUT_PATH)
