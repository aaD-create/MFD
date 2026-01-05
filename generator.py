
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

