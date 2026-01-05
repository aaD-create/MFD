
from flask import Flask, render_template, request, send_file
from datetime import datetime
import io
from generator import build_proposal_ppt

app = Flask(__name__)

# ---- Helper: choose Bajaj MAF vs Bajaj Flexi Cap when not explicitly provided ----
def _select_bajaj_choice(payload: dict) -> str:
    """
    Returns one of: 'maf', 'flexi'
    Rule of thumb:
      - Very high / High risk, long tenure, or no short-term liquidity -> 'flexi'
      - Low / Conservative risk, or declared liquidity needs -> 'maf'
    If the user explicitly sent bajajFundChoice ('maf'/'flexi'), honor it.
    """
    explicit = (payload.get("bajajFundChoice") or "").strip().lower()
    if explicit in {"maf", "flexi"}:
        return explicit

    risk_appetite = (payload.get("riskAppetite") or "").strip().lower()
    risk_tolerance = (payload.get("riskTolerance") or "").strip().lower()
    tenure = (payload.get("tenure") or "").strip().lower()      # e.g., '3to5', '5to10', '10+ years'
    liquidity = (payload.get("liquidityNeeds") or "").strip().lower()

    # If investor can stomach more equity risk, prefer Bajaj Flexi Cap
    if "very" in risk_appetite or "high" in risk_appetite or "aggressive" in risk_tolerance:
        return "flexi"

    # If short/near-term liquidity is declared, prefer Bajaj MAF (hybrid multi-asset)
    if liquidity.startswith("yes") or "yes" in liquidity:
        return "maf"

    # Tenure-based nudge: longer tenures can carry more equity exposure
    if tenure in {"5to10", "gt10", "10+ years"}:
        return "flexi"

    # Default fallback
    return "maf"


@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json(force=True) or {}

    # ---- Enforce Bajaj AMC inclusion on the server ----
    # 1) Always set the enforcement flag the generator looks for
    data["mustIncludeBajaj"] = True

    # 2) If the frontend sent a choice, we honor it; otherwise select based on inputs
    if not data.get("bajajFundChoice"):
        data["bajajFundChoice"] = _select_bajaj_choice(data)

    # Optional: guard against unexpected values
    if data["bajajFundChoice"] not in {"maf", "flexi", "auto"}:
        data["bajajFundChoice"] = _select_bajaj_choice(data)

    # 3) (Optional) If exit-load tolerance is very strict (e.g., 0%), nudge 'maf' to keep equity sleeve lighter
    try:

