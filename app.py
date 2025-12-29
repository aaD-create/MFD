from flask import Flask, render_template, request, send_file
from datetime import datetime
import io
from generator import build_proposal_ppt

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json(force=True)
    filename = f"Investment_Proposal_{data.get('investorName','Investor').replace(' ','_')}_{datetime.now().strftime('%Y-%m-%d')}.pptx"
    # Build PPT in memory
    buf = io.BytesIO()
    build_proposal_ppt(buf, data)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
