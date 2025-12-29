# Branded Mutual Fund Proposal Generator (Web + PPT)

Generate a professional investment proposal deck in **PowerPoint (.pptx)** from a simple web form (the 7 questions).

## Features
- Branded web UI (HTML/CSS) — add your logo and colours
- Flask backend — builds PPT on the fly using `python-pptx`
- Suitability logic (tenure × risk) + sample scheme metrics
- Immediate **download link** in the browser

> ⚠️ Replace demo scheme metrics in `generator.py` with **live factsheet values** before client use.

## Quick start
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```
Open `http://localhost:5000` in your browser.

## Branding
- Place your logo at `static/logo.png` (and reference in `templates/index.html` if needed)
- Adjust theme colours in the `<style>` section

## Customise scheme metrics
Edit `generator.py` → `SCHEMES` list with **live data** (3Y CAGR, Sharpe, Beta, Std. Dev., P/E, Exit‑load, AUM, Expense, Turnover, Alpha, Drawdown, Riskometer, etc.).

## Compliance
- **SEBI exit‑load cap**: 3% (Sep‑2025). Individual scheme grids apply.
- **Riskometer** must match investor risk profile.

## Notes
- For production, deploy behind HTTPS.
- Add authentication and logging if used across sub‑distributors.
```
