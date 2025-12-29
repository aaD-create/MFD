from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import ChartData
from datetime import datetime

# Demo scheme database (replace with live factsheet metrics)
SCHEMES = [
    {
        "name": "Parag Parikh Flexi Cap Fund – Direct Growth",
        "group": "Equity", "subCat": "Flexi Cap",
        "metrics": {
            "cagr3": 23.06, "cagrCat3": 17.76, "sharpe": 1.55, "beta": 0.57, "stddev": 8.41, "pe": 20.62,
            "exitLoad": "For units >10%: 2% ≤365d; 1% 366–730d", "aumCr": 129783, "brokerageBps": None,
            "top3Sectors": 48.28, "top10Stocks": 47.89,
            "expense": 0.63, "turnoverPct": 18.81, "alpha": 7.42, "maxDDPct": -6.05,
            "riskometer": "Very High"
        },
        "notes": "Core flexi‑cap; lower beta and strong risk‑adjusted profile vs category"
    },
    {
        "name": "HDFC Mid‑Cap Fund – Direct Growth",
        "group": "Equity", "subCat": "Mid Cap",
        "metrics": {
            "cagr3": 26.95, "cagrCat3": 24.74, "sharpe": 1.316, "beta": 0.858, "stddev": 13.847, "pe": None,
            "exitLoad": "1% ≤365 days; Nil thereafter", "aumCr": 92168.85, "brokerageBps": None,
            "top3Sectors": 56.23, "top10Stocks": 33.44,
            "expense": 1.36, "turnoverPct": 18.22, "alpha": 4.77, "maxDDPct": -16.42,
            "riskometer": "Very High"
        },
        "notes": "Quality‑biased mid‑cap engine; strong 3Y risk metrics"
    },
    {
        "name": "ICICI Prudential Balanced Advantage Fund — Regular Growth",
        "group": "Hybrid", "subCat": "Balanced Advantage",
        "metrics": {
            "cagr3": 13.61, "cagrCat3": 11.96, "sharpe": 1.19, "beta": 0.55, "stddev": 5.46, "pe": None,
            "exitLoad": ">30% units: 1% within 1 year", "aumCr": 69868, "brokerageBps": None,
            "top3Sectors": 41.80, "top10Stocks": 55.00,
            "expense": 1.43, "turnoverPct": 33.97, "alpha": 3.06, "maxDDPct": -6.09,
            "riskometer": "Very High"
        },
        "notes": "Dynamic allocation moderates drawdowns; stabiliser for 3–5 yrs"
    },
    {
        "name": "Axis Liquid Fund – Direct Growth",
        "group": "Debt", "subCat": "Liquid",
        "metrics": {
            "cagr3": 7.08, "cagrCat3": 7.02, "sharpe": 3.65, "beta": 0.38, "stddev": 0.80, "pe": None,
            "exitLoad": "Day‑1 0.0070% … Day‑7+ 0%", "aumCr": 37357.87, "brokerageBps": None,
            "top3Sectors": None, "top10Stocks": None,
            "expense": 0.24, "turnoverPct": None, "alpha": 1.30, "maxDDPct": None,
            "riskometer": "Low to Moderate"
        },
        "notes": "High‑liquidity sleeve to ring‑fence near‑term obligations"
    }
]

POLICY = {
    "LT1": {"Debt":100,"Hybrid":0,"Equity":0},
    "1to3": {"Debt":60,"Hybrid":40,"Equity":0},
    "3to5": {"Debt":20,"Hybrid":30,"Equity":50},
    "5to10": {"Debt":15,"Hybrid":20,"Equity":65},
    "GT10": {"Debt":10,"Hybrid":10,"Equity":80}
}

TENURE_LABELS = {"LT1":"< 1 year","1to3":"1–3 years","3to5":"3–5 years","5to10":"5–10 years","GT10":"10+ years"}


def build_proposal_ppt(buf, inputs):
    # derive asset mix
    tenure = inputs.get('tenure','3to5')
    mix = POLICY.get(tenure, POLICY['3to5'])
    # sleeve selection: top per group
    picks = []
    # choose Axis Liquid for Debt, ICICI BAF for Hybrid, Parag + HDFC for Equity
    for s in SCHEMES:
        g = s['group']
        if g=='Debt' and mix['Debt']>0:
            picks.append((s, 20))
        if g=='Hybrid' and mix['Hybrid']>0:
            picks.append((s, 30))
    # equity picks
    equities = [s for s in SCHEMES if s['group']=='Equity']
    # split equity 50 as 25/25
    for s in equities[:2]:
        picks.append((s, 25))

    # sector distribution (static example; could be computed from holdings)
    sector_dist = {
        'Debt': 29.5, 'Others': 28.6, 'Financials': 18.9, 'IT': 6.4,
        'Automobile': 4.3, 'Energy': 3.6, 'Healthcare': 3.0, 'Capital Goods': 2.3
    }

    prs = Presentation()

    # Slide 1
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = 'Investment Proposal — Live Assessment'
    s1.placeholders[1].text = f"Investor: {inputs.get('investorName','Investor')}\nPrepared by: {inputs.get('distributorName','Advisor')}\nDate: {datetime.now().strftime('%Y-%m-%d')}"
    s1.shapes.title.text_frame.paragraphs[0].font.size = Pt(36)

    # Slide 2
    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = 'Asset Mix & Suggested Allocation'
    tf = s2.placeholders[1].text_frame
    tf.clear()
    p = tf.add_paragraph(); p.text = f"Asset mix: Equity {mix['Equity']}% • Hybrid {mix['Hybrid']}% • Debt {mix['Debt']}%"; p.level=0; p.font.size=Pt(16)
    for sch, wt in picks:
        q = tf.add_paragraph(); q.level=0; q.text = f"• {sch['name']} — {wt}%"; q.font.size=Pt(14)

    # Slide 3 — Table
    s3 = prs.slides.add_slide(prs.slide_layouts[1])
    s3.shapes.title.text = 'Selected Funds — Live Metrics & Rationale'
    left, top, width, height = Inches(0.5), Inches(1.6), Inches(9.0), Inches(3.6)
    rows = len(picks)+1
    cols = 17
    table = s3.shapes.add_table(rows, cols, left, top, width, height).table
    headers = ["Scheme","Group/Category","3Y CAGR (vs cat)","Sharpe","Beta","Std.Dev","P/E","Exit‑load","AUM (₹ Cr)","Brokerage (bps)","Top‑3 sectors %","Top‑10 stocks %","Expense","Turnover","Alpha","Max DD","Riskometer"]
    for i,h in enumerate(headers):
        table.cell(0,i).text = h
    for r,(sch,wt) in enumerate(picks, start=1):
        m = sch['metrics']
        rowdata = [
            sch['name'], f"{sch['group']}/{sch['subCat']}", f"{m['cagr3']}% (vs {m['cagrCat3']}%)", str(m['sharpe']), str(m['beta']), f"{m['stddev']}%",
            str(m['pe']) if m['pe'] is not None else '—', m['exitLoad'], f"{m['aumCr']:,}", (str(m['brokerageBps']) if m['brokerageBps'] else 'N/A'),
            (f"{m['top3Sectors']}%" if m['top3Sectors'] is not None else '—'), (f"{m['top10Stocks']}%" if m['top10Stocks'] is not None else '—'),
            f"{m['expense']}%" if m['expense'] is not None else '—', (f"{m['turnoverPct']}%" if m['turnoverPct'] is not None else '—'),
            (f"{m['alpha']}%" if m['alpha'] is not None else '—'), (f"{m['maxDDPct']}%" if m['maxDDPct'] is not None else '—'), m['riskometer']
        ]
        for c,val in enumerate(rowdata):
            table.cell(r,c).text = val
    # Rationale bullets
    tf3 = s3.placeholders[1].text_frame
    b = tf3.add_paragraph(); b.text = "\nRationale:"; b.level=0; b.font.bold=True
    for sch,wt in picks:
        rp = tf3.add_paragraph(); rp.level=0; rp.text = f"• {sch['name']}: {sch['notes']}"

    # Slide 4 — Sector chart
    s4 = prs.slides.add_slide(prs.slide_layouts[1])
    s4.shapes.title.text = 'Estimated ₹30,000 SIP distribution by sector'
    cd = ChartData(); cd.categories = list(sector_dist.keys()); cd.add_series('Allocation %', list(sector_dist.values()))
    s4.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.8), Inches(9.2), Inches(4.5), cd)

    # Slide 5 — Key metrics & Compliance
    s5 = prs.slides.add_slide(prs.slide_layouts[1])
    s5.shapes.title.text = 'Key Metrics & Compliance Notes'
    notes = s5.placeholders[1].text_frame
    notes.clear()
    bullets = [
        'Sharpe = (Return − Risk‑free) / Standard deviation — higher ⇒ better risk‑adjusted performance.',
        'Beta: sensitivity to market (1.0 = market); lower beta ⇒ lower volatility.',
        'Standard deviation: volatility of returns; lower ⇒ more stability.',
        'Sortino: Sharpe variant penalising downside volatility only.',
        'Rolling returns: consistency across overlapping periods; preferred to single point‑to‑point.',
        'Riskometer: SEBI/AMFI label (Low → Very High); must align with investor profile.',
        'Exit‑load: SEBI maximum cap is 3%; scheme grid applies.',
        'Disclaimer: MF investments are subject to market risks. Past performance may or may not be sustained. Suitability first; brokerage does not override investor interest.'
    ]
    for it in bullets:
        p = notes.add_paragraph(); p.text = '• ' + it; p.level=0

    prs.save(buf)
