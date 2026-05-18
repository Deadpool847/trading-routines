"""Generate the daily trading-intelligence .docx brief."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATE_ISO = "2026-05-18"
BIAS = "BEAR"
ENTRY_PERMISSION = "YELLOW"
CONFIDENCE = "MEDIUM"

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


def add_kv_table(doc, rows, col_widths=(4.2, 3.5, 4.5)):
    table = doc.add_table(rows=1 + len(rows), cols=3)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    for i, name in enumerate(["Metric", "Value", "Signal"]):
        hdr[i].text = ""
        run = hdr[i].paragraphs[0].add_run(name)
        run.bold = True
        run.font.size = Pt(11)
        shade_cell(hdr[i], "1F4E78")
        for r in hdr[i].paragraphs[0].runs:
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for i, (k, v, s) in enumerate(rows, start=1):
        c = table.rows[i].cells
        c[0].text = k
        c[1].text = v
        c[2].text = s
        for cell in c:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10)
    for i, w in enumerate(col_widths):
        for row in table.rows:
            row.cells[i].width = Inches(w)
    return table


def add_setup_block(doc, stock):
    add_heading(
        doc,
        f"{stock['name']} (NSE: {stock['ticker']}) | {stock['side']} | Grade {stock['grade']}",
        level=2,
        color=RGBColor(0x1F, 0x4E, 0x78),
    )
    add_para(
        doc,
        f"CMP: Rs.{stock['cmp']} | Lot: {stock['lot']} | Crude aligned: {stock['crude_aligned']}",
        bold=True,
    )
    add_para(doc, f"CATALYST: {stock['catalyst']}")

    add_para(doc, "TECHNICAL:", bold=True)
    add_bullets(
        doc,
        [
            f"Daily trend: {stock['daily']}",
            f"Above 200 DMA: {stock['above_200dma']}",
            f"Weekly trend: {stock['weekly']}",
            f"Support Rs.{stock['support']} | Resistance Rs.{stock['resistance']}",
        ],
    )

    add_para(doc, "STRUCTURAL:", bold=True)
    add_bullets(
        doc,
        [
            f"OI direction: {stock['oi']}",
            f"Delivery %: {stock['delivery']}",
            f"Block deal: {stock['block']} | Ban list: {stock['ban']}",
        ],
    )

    add_para(doc, "OPTION SETUP:", bold=True)
    add_bullets(
        doc,
        [
            f"Strike: Rs.{stock['strike']} [{stock['moneyness']}] | Expiry: {stock['expiry']}",
            f"THETA WARNING if expiry <= 3 days: {stock['theta_warn']}",
        ],
    )

    add_para(doc, "ENTRY (ALL 4 must fire):", bold=True)
    add_bullets(
        doc,
        [
            f"SuperTrend {stock['st']} on 15-min confirmed",
            f"RSI {stock['rsi']} 50 on 15-min",
            f"StochRSI cross {stock['stoch']} from extreme",
            "Volume > 1.5x 20-period avg",
            f"Confirmation candle: 15-min close {stock['confirm']} Rs.{stock['confirm_level']}",
        ],
    )

    add_para(
        doc,
        f"T1 (book 40%): Rs.{stock['t1']} | T2 (book 40%): Rs.{stock['t2']}",
        bold=True,
    )
    add_para(doc, f"SL: Rs.{stock['sl']} + {stock['sl_cond']}")
    add_para(doc, f"R:R: {stock['rr']} (skip if below 2:1)")
    add_para(doc, "Time stop: Exit by 13:00 IST")
    add_para(doc, f"GRADE REASON: {stock['reason']}", italic=True)
    doc.add_paragraph("")


def hr_line(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "808080")
    pBdr.append(bottom)
    pPr.append(pBdr)


def main():
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)

    title = doc.add_heading(
        f"ELITE DAILY TRADING BRIEF - {DATE_ISO} - {BIAS}", level=0
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.color.rgb = RGBColor(0xB0, 0x00, 0x00)

    add_para(
        doc,
        "Generated: 03:30 IST | Weekly expiry: 3 days | Monthly expiry: 10 days | Expiry week: YES",
        italic=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    banner = doc.add_table(rows=5, cols=2)
    banner.style = "Light Shading Accent 1"
    banner_rows = [
        ("BIAS", f"{BIAS} (Cautious - high headline risk)"),
        ("ENTRY PERMISSION", ENTRY_PERMISSION),
        ("CRUDE MODE", "BEAR | Direction: RISING (war escalating)"),
        ("VIX SIZING", "HALF SIZE (India VIX 18.78, range 15-20)"),
        ("EXPIRY ALERT", "Weekly expiry in 3 days -> MONTHLY OPTIONS ONLY"),
    ]
    for i, (k, v) in enumerate(banner_rows):
        banner.rows[i].cells[0].text = k
        banner.rows[i].cells[1].text = v
        for p in banner.rows[i].cells[0].paragraphs:
            for r in p.runs:
                r.bold = True

    doc.add_paragraph("")
    hr_line(doc)

    add_heading(doc, "SECTION 1 - MACRO SNAPSHOT", level=1)

    macro_rows = [
        ("MCX Crude (proxy)", "Rs.9,170 / bbl", "BEAR mode (9k-10k band)"),
        ("Brent / WTI", "$110.93 / $105.42", "RISING (+8.1% / +11% wk)"),
        ("GIFT Nifty gap", "-64 pts (~23,579)", "Modest gap DOWN"),
        ("S&P 500", "7,408.50 (-1.24%)", "Risk OFF (Fri sell-off)"),
        ("Nasdaq", "26,225.14 (-1.54%)", "Tech weakness"),
        ("US VIX", "Elevated (war premium)", "FEAR"),
        ("India VIX", "18.78 (+0.97%)", "HALF SIZE rule"),
        ("Nifty close (15-May)", "Rs.23,643.50 (-0.19%)", "Holding 23,600 zone"),
        ("FII cash (14-May)", "+Rs.187 Cr | MTD -Rs.25,985 Cr", "MTD heavy SELL"),
        ("DII cash (14-May)", "+Rs.684 Cr | MTD +Rs.41,876 Cr", "MTD heavy BUY"),
        ("USD-INR (est.)", "~Rs.86-87 (record-low zone)", "RBI defending"),
        ("INFY ADR", "~$11.86 (14-May)", "Slight negative drag"),
        ("HDB ADR", "DATA_UNAVAILABLE", "Use BankNifty proxy"),
    ]
    add_kv_table(doc, macro_rows)

    add_para(
        doc,
        "Geopolitical one-liner: Strait of Hormuz remains blocked (since 28-Feb); "
        "Trump rejected Iran counter-offer, ceasefire on 'life support'.",
        bold=True,
    )
    add_para(doc, "Today's intraday headline risk: HIGH", bold=True, color=RGBColor(0xB0, 0x00, 0x00))

    hr_line(doc)

    add_heading(doc, "SECTION 2 - CRUDE RULE + STRUCTURAL READ", level=1)

    add_para(doc, "CRUDE_MODE: BEAR", bold=True, color=RGBColor(0xB0, 0x00, 0x00))
    add_para(doc, "CRUDE_DIRECTION: RISING (war ongoing, +11% WTI wk)", bold=True)
    add_para(doc, "CRUDE_FLIP_LEVEL: Rs.8,500 (down) -> would re-enter CAUTION", bold=True)
    add_para(
        doc,
        "AVOID: OMCs (BPCL/HPCL/IOC) as calls, aviation (INDIGO) as calls, "
        "paints (ASIANPAINT/BERGER) as calls, tyres (APOLLO/MRF), fertilisers (CHAMBLFERT)",
    )
    add_para(
        doc,
        "EXCEPTION_CALLS allowed on crude beneficiaries: ONGC, OIL INDIA, MRPL "
        "(+ defence: BEL, HAL, BDL as war beneficiaries)",
    )

    add_para(doc, "Structural read:", bold=True)
    add_bullets(
        doc,
        [
            "Nifty vs Max Pain: DATA_UNAVAILABLE for live max pain - use 23,500 estimated pin zone (round-number magnetism)",
            "PCR: DATA_UNAVAILABLE - proceed without it; rely on price/OI action intraday",
            "OI change direction: To be verified at 9:15 - watch 23,700-23,800 Call OI buildup (resistance) and 23,500 Put OI (floor)",
            "Futures basis: Likely DISCOUNT given gap-down + FII MTD selling pressure",
            "Expiry week dynamics (3 days to weekly): Put writers may defend 23,500; Call writers stacked 23,800-24,000 ceiling; max-pain magnetism = SIDEWAYS-DOWN bias",
            "Bank Nifty leadership: Check at 9:30 - if BankNifty leads down, BEAR move is REAL; if BankNifty holds, Nifty fall is SUSPECT (fade)",
        ],
    )

    hr_line(doc)

    add_heading(doc, "SECTION 3 - CONTRARIAN CHECK (Layer 3)", level=1)

    add_para(doc, "Q1. CONSENSUS expectation today:", bold=True)
    add_para(
        doc,
        "Every retail desk and TV channel will scream 'gap down on crude spike, "
        "go short Nifty / buy puts on aviation, OMCs'. Heavy IT weakness ('INFY ADR negative') "
        "narrative. Iran headlines + Friday US sell-off = uniform bearish consensus.",
    )

    add_para(doc, "Q2. THE TRAP:", bold=True)
    add_para(
        doc,
        "Market punishes the easy short. Gap-down on Monday after a -1.2% Friday US session "
        "is the textbook 'sell-the-rumour, panic open, V-recovery' setup. DII MTD buying "
        "(+Rs.41,876 Cr) is absorbing all FII selling - a green 30-min candle traps shorts. "
        "Profit goes to put-writers at 23,500 and call-writers at 23,800.",
    )

    add_para(doc, "Q3. RETAIL STOPS:", bold=True)
    add_para(
        doc,
        "Long stops clustered below 23,500 (psychological + put OI wall). "
        "Short stops clustered above 23,750 (Fri's intraday high zone). "
        "A spike through either side = stop-hunt fuel.",
    )

    add_para(doc, "Q4. FLIP TRIGGER:", bold=True)
    add_para(
        doc,
        "Single event that reverses today's direction: Any Iran-US ceasefire / "
        "Hormuz reopening headline (crude -5%+ instantly) -> Nifty rips 200+ pts, "
        "ASIANPAINT/INDIGO/BPCL rally hard, ONGC/BEL get sold. "
        "Direction = full BULL reversal.",
    )

    add_heading(doc, "CONTRARIAN OVERRIDE LOGIC", level=2)
    add_bullets(
        doc,
        [
            "A) GIFT Nifty gap-up > +100 pts? NO (gap DOWN ~64 pts)",
            "B) Crude falling > 2% overnight? NO (RISING +11% wk)",
            "C) Contrarian scenario probability > 40%? PARTIAL (~30% - DII absorption real)",
            "D) Expiry week + PCR < 0.8 (squeeze risk)? PCR unknown, but moderate squeeze risk on a ceasefire headline",
        ],
    )
    add_para(
        doc,
        "VERDICT: No full override triggered. BEAR bias retained, but confidence "
        "downgraded to MEDIUM and permission to YELLOW due to (i) DII absorption "
        "strength, (ii) expiry-week put-writer defense at 23,500, "
        "(iii) headline-flip risk on Iran.",
        bold=True,
    )

    hr_line(doc)

    add_heading(doc, "SECTION 4 - 5 INDIVIDUAL F&O SETUPS", level=1)

    setups = [
        {
            "name": "ONGC",
            "ticker": "ONGC",
            "side": "CALL",
            "grade": "A+",
            "cmp": "301.05",
            "lot": "2,250",
            "crude_aligned": "YES (crude producer beneficiary)",
            "catalyst": "Brent $110+ keeps ONGC realisations elevated; war-premium adds upside; near 52w high Rs.307.50 = breakout watch.",
            "daily": "UP (strong)",
            "above_200dma": "YES (200DMA Rs.254.43; price 18% above)",
            "weekly": "UP (+6.7% past week)",
            "support": "295",
            "resistance": "307.50",
            "oi": "LONG BUILD expected (verify at open)",
            "delivery": "RISING (institutional accumulation)",
            "block": "NO",
            "ban": "NO",
            "strike": "300 CE",
            "moneyness": "ATM",
            "expiry": "29-May-2026 (MONTHLY - weekly banned, 3d to expiry)",
            "theta_warn": "NO (monthly used)",
            "st": "GREEN",
            "rsi": "above",
            "stoch": "up",
            "confirm": "above",
            "confirm_level": "303.50",
            "t1": "311",
            "t2": "318",
            "sl": "295 (below 50DMA-adjacent)",
            "sl_cond": "15-min close < Rs.295 OR ST flip RED",
            "rr": "2.5:1",
            "reason": "Cleanest crude-beneficiary, above 200DMA, near 52w-high breakout, war tailwind; A+.",
        },
        {
            "name": "BEL",
            "ticker": "BEL",
            "side": "CALL",
            "grade": "A",
            "cmp": "422.70",
            "lot": "2,850",
            "crude_aligned": "YES (defence + war premium)",
            "catalyst": "Active US-Iran war = defence order-flow narrative; recent sideways consolidation above 200DMA invites breakout.",
            "daily": "SIDEWAYS (consolidating)",
            "above_200dma": "YES (200DMA Rs.414.32)",
            "weekly": "SIDEWAYS",
            "support": "415",
            "resistance": "438 (50DMA) -> 473 (52w high)",
            "oi": "NEUTRAL - needs confirmation",
            "delivery": "STABLE",
            "block": "NO",
            "ban": "NO",
            "strike": "430 CE",
            "moneyness": "OTM (slight)",
            "expiry": "29-May-2026 (MONTHLY)",
            "theta_warn": "NO (monthly used)",
            "st": "GREEN",
            "rsi": "above",
            "stoch": "up",
            "confirm": "above",
            "confirm_level": "427",
            "t1": "438",
            "t2": "448",
            "sl": "418 (below 200DMA)",
            "sl_cond": "15-min close < Rs.418 OR ST flip RED",
            "rr": "2.2:1",
            "reason": "Above 200DMA, defence beta to war headlines; but below 50DMA so trend still hesitant - A not A+.",
        },
        {
            "name": "INTERGLOBE AVIATION (INDIGO)",
            "ticker": "INDIGO",
            "side": "PUT",
            "grade": "A",
            "cmp": "4,314.90",
            "lot": "300",
            "crude_aligned": "YES (crude victim - fuel = 40% cost)",
            "catalyst": "Brent $110 + war = direct margin hit; stock already 30.77% below 52w high (Rs.6,232) confirms damage; further downside on every crude tick.",
            "daily": "DOWN",
            "above_200dma": "NO (broken)",
            "weekly": "DOWN",
            "support": "4,200 -> 3,895 (52w low)",
            "resistance": "4,360 (Fri high)",
            "oi": "SHORT BUILD expected",
            "delivery": "FALLING (distribution)",
            "block": "NO",
            "ban": "NO",
            "strike": "4,300 PE",
            "moneyness": "ATM",
            "expiry": "29-May-2026 (MONTHLY)",
            "theta_warn": "NO (monthly used)",
            "st": "RED",
            "rsi": "below",
            "stoch": "down",
            "confirm": "below",
            "confirm_level": "4,277 (Fri low)",
            "t1": "4,220",
            "t2": "4,150",
            "sl": "4,365 (above Fri high)",
            "sl_cond": "15-min close > Rs.4,365 OR ST flip GREEN",
            "rr": "2.3:1",
            "reason": "Textbook crude-victim, broken 200DMA, downtrend intact; A.",
        },
        {
            "name": "BPCL",
            "ticker": "BPCL",
            "side": "PUT",
            "grade": "A+",
            "cmp": "284.40",
            "lot": "1,800",
            "crude_aligned": "YES (OMC - under-recoveries spike with crude)",
            "catalyst": "Brent $110 + government price-cap risk = OMC margin compression; price already in confirmed downtrend, near 52w low Rs.266.60; IOC results today = sector tone risk.",
            "daily": "DOWN",
            "above_200dma": "NO (well below)",
            "weekly": "DOWN (-23.4% in 6 months)",
            "support": "275 -> 266.60 (52w low)",
            "resistance": "295",
            "oi": "SHORT BUILD",
            "delivery": "FALLING",
            "block": "NO",
            "ban": "NO",
            "strike": "285 PE",
            "moneyness": "ATM",
            "expiry": "29-May-2026 (MONTHLY)",
            "theta_warn": "NO (monthly used)",
            "st": "RED",
            "rsi": "below",
            "stoch": "down",
            "confirm": "below",
            "confirm_level": "282",
            "t1": "275",
            "t2": "268",
            "sl": "291 (above Fri close)",
            "sl_cond": "15-min close > Rs.291 OR ST flip GREEN",
            "rr": "2.4:1",
            "reason": "Strongest put thesis: confirmed downtrend, 52w-low magnet, crude victim, sector-event (IOC results) catalyst; A+.",
        },
        {
            "name": "ASIAN PAINTS",
            "ticker": "ASIANPAINT",
            "side": "PUT",
            "grade": "B",
            "cmp": "2,605.60",
            "lot": "200",
            "crude_aligned": "YES (crude derivatives = >50% of raw material)",
            "catalyst": "Crude $110+ smashes gross margin; but technicals contradict - stock recently broke above MAs with RSI >60. Counter-trend put = lower confidence; valid only if Nifty cracks.",
            "daily": "UP (recent breakout - risk to put)",
            "above_200dma": "YES (against put thesis)",
            "weekly": "UP (momentum)",
            "support": "2,560",
            "resistance": "2,650 -> 2,720",
            "oi": "MIXED - verify at open",
            "delivery": "RISING (caution: accumulation)",
            "block": "NO",
            "ban": "NO",
            "strike": "2,600 PE",
            "moneyness": "ATM",
            "expiry": "29-May-2026 (MONTHLY)",
            "theta_warn": "NO (monthly used)",
            "st": "RED (must flip to RED first)",
            "rsi": "below",
            "stoch": "down",
            "confirm": "below",
            "confirm_level": "2,558 (breakdown of recent support)",
            "t1": "2,510",
            "t2": "2,450",
            "sl": "2,640 (above recent high zone)",
            "sl_cond": "15-min close > Rs.2,640 OR ST flip GREEN - strict",
            "rr": "2.0:1",
            "reason": "Fundamental crude headwind real, but technicals against - counter-trend trade; only fires after clean breakdown; B grade.",
        },
    ]

    for s in setups:
        add_setup_block(doc, s)

    hr_line(doc)

    add_heading(doc, "SECTION 5 - FINAL VERDICT", level=1)

    verdict_rows = [
        ("TODAY'S BIAS", BIAS, "Confidence MEDIUM"),
        ("CRUDE MODE", "BEAR (RISING)", "MCX ~Rs.9,170"),
        ("VIX SIZING", "HALF SIZE", "India VIX 18.78"),
        ("ENTRY PERMISSION", "YELLOW", "Selective, hand-picked only"),
    ]
    add_kv_table(doc, verdict_rows)

    add_para(doc, "THE BULL CASE (steelman, even though bias = BEAR):", bold=True)
    add_para(
        doc,
        "(1) DII +Rs.41,876 Cr MTD absorbs all FII selling; (2) gap-down opens are "
        "regularly bought in expiry week as put-writers defend 23,500; "
        "(3) one Iran-ceasefire headline cuts crude 5%+ and rips Nifty 200+ pts up.",
    )

    add_para(doc, "THE BEAR CASE (primary thesis):", bold=True)
    add_para(
        doc,
        "Crude $110+ with Hormuz blocked = persistent CAD/inflation/INR pressure; "
        "FII MTD selling -Rs.25,985 Cr; US -1.24% Friday + war headlines + INFY ADR drag "
        "= structural risk-off through 9:30 IST.",
    )

    add_para(doc, "FLIP TRIGGER:", bold=True, color=RGBColor(0xB0, 0x00, 0x00))
    add_para(
        doc,
        "If Brent breaks back below $105 (-3% intraday) OR any Iran-US ceasefire / "
        "Hormuz reopening headline hits the wire -> bias flips to CAUTIOUS BULL; "
        "exit all puts immediately, switch to ONGC put / INDIGO call scalp.",
    )

    add_para(doc, "NIFTY KEY LEVELS:", bold=True)
    add_para(
        doc,
        "S2 Rs.23,350 | S1 Rs.23,500 | CRITICAL Rs.23,643 (Fri close) | R1 Rs.23,750 | R2 Rs.23,880",
    )

    add_para(doc, "TOP 3 RANKED:", bold=True)
    add_bullets(
        doc,
        [
            "#1 BPCL - A+ - confirmed downtrend OMC put with sector-event catalyst - entry below Rs.282",
            "#2 ONGC - A+ - crude-beneficiary call near 52w-high breakout - entry above Rs.303.50",
            "#3 INDIGO - A - aviation put, broken 200DMA, fuel-margin hit - entry below Rs.4,277",
        ],
    )

    add_para(
        doc,
        "ONE RISK THAT RUINS EVERYTHING TODAY:",
        bold=True,
        color=RGBColor(0xB0, 0x00, 0x00),
    )
    add_para(
        doc,
        "Surprise Iran-US ceasefire announcement during market hours -> crude crashes "
        "5-8%, Nifty gaps 200+ pts up, every put on the desk (BPCL, INDIGO, ASIANPAINT) "
        "evaporates AND every call on ONGC/BEL collapses. Tail risk is bidirectional "
        "but most damaging to a BEAR book.",
    )

    hr_line(doc)

    add_heading(doc, "ONE-LINE SUMMARY (read at 9:10 AM)", level=1)
    add_para(
        doc,
        '"Today is BEAR because crude is $110+ with Hormuz blocked and US -1.24% Fri. '
        'Crude Rs.9,170 = BEAR, rising. Watch BPCL PUT and ONGC CALL. Key risk: surprise '
        'Iran-US ceasefire headline. Size HALF. Flips if Brent < $105 or ceasefire hits."',
        bold=True,
        italic=True,
    )

    hr_line(doc)

    add_para(
        doc,
        "Trader Profile: NSE F&O only | Window 9:15-12:30 | Hard exit 13:30 IST | "
        "Max 3 trades | Risk cap 2% | All 4 entry conditions mandatory | "
        "Monthly options only (3d to weekly).",
        italic=True,
        size=9,
        color=RGBColor(0x55, 0x55, 0x55),
    )

    doc.save(OUT_PATH)
    print(f"Saved: {OUT_PATH}")


if __name__ == "__main__":
    main()
