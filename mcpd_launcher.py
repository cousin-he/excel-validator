# -*- coding: utf-8 -*-
"""
mcpd_launcher.py
Web interface to validate files according to:
- MCPD v4.3 (genetic resources)
- Climate (environmental data)

Usage:
    python mcpd_launcher.py

Then open http://localhost:5050
"""

import os
import sys
import tempfile
from flask import Flask, request, render_template_string

sys.path.insert(0, r"C:\Users\rahma\developement")
from validator import run_validation as run_mcpd
from validator_climate import run_validation as run_climate

app = Flask(__name__)

HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Data Validator - MCPD & Climate</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 1000px; margin: 40px auto; padding: 20px; }
        h1 { color: #2c3e50; }
        .upload-box {
            border: 2px dashed #3498db;
            border-radius: 10px;
            padding: 30px;
            text-align: center;
            margin: 20px 0;
            background: #f8f9fa;
        }
        input[type=file] { margin: 10px 0; }
        button {
            background: #3498db;
            color: white;
            border: none;
            padding: 10px 25px;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
        }
        button:hover { background: #2980b9; }
        .report {
            background: #1e1e1e;
            color: #d4d4d4;
            padding: 20px;
            border-radius: 8px;
            font-family: monospace;
            font-size: 13px;
            white-space: pre-wrap;
            margin-top: 20px;
            max-height: 600px;
            overflow-y: auto;
        }
        .valid   { background: #27ae60; color: white; padding: 10px 20px; border-radius: 5px; font-weight: bold; }
        .invalid { background: #e74c3c; color: white; padding: 10px 20px; border-radius: 5px; font-weight: bold; }
        .badge { display: inline-block; margin-bottom: 15px; font-size: 18px; }
        .nav { background: #2c3e50; color: white; padding: 10px 20px; margin: -20px -20px 30px -20px; }
        .nav a { color: #3498db; text-decoration: none; margin-right: 20px; }
        .selector {
            margin: 15px 0;
            text-align: left;
            display: inline-block;
        }
        .selector label {
            margin-right: 20px;
            font-weight: normal;
        }
        input[type="radio"] {
            margin-right: 5px;
        }
    </style>
</head>
<body>
    <div class="nav">
        <strong>Excel Validator</strong>
        <a href="http://localhost:8888">&larr; Back to MIAPPE</a>
    </div>

    <h1>Data Validator</h1>
    <p>Choose validation type, then select an Excel file (.xlsx).</p>

    <div class="upload-box">
        <form method="POST" enctype="multipart/form-data">
            <div class="selector">
                <label>
                    <input type="radio" name="validation_type" value="mcpd" {% if validation_type == 'mcpd' %}checked{% endif %}>
                    🌾 MCPD (genetic resources)
                </label>
                <label>
                    <input type="radio" name="validation_type" value="climate" {% if validation_type == 'climate' %}checked{% endif %}>
                    🌡️ Climate (environmental data)
                </label>
            </div>
            <br>
            <p>📎 Excel file (.xlsx):</p>
            <input type="file" name="fichier" accept=".xlsx" required><br><br>
            <button type="submit">Run validation</button>
        </form>
    </div>

    {% if report %}
        <div class="badge">
            {% if valid %}
                <span class="valid">VALID</span>
            {% else %}
                <span class="invalid">INVALID</span>
            {% endif %}
        </div>
        <div class="report">{{ report }}</div>
    {% endif %}
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    report = None
    valid = False
    validation_type = "mcpd"

    if request.method == "POST":
        file = request.files.get("fichier")
        validation_type = request.form.get("validation_type", "mcpd")

        if file and file.filename.endswith(".xlsx"):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                file.save(tmp.name)
                tmp_path = tmp.name
            try:
                if validation_type == "mcpd":
                    report = run_mcpd(tmp_path)
                else:
                    report = run_climate(tmp_path)
                valid = "Blocking errors      : 0" in report
            except Exception as e:
                report = f"[ERROR] {e}"
                valid = False
            finally:
                try:
                    os.unlink(tmp_path)
                except:
                    pass

    return render_template_string(HTML, report=report, valid=valid, validation_type=validation_type)


if __name__ == "__main__":
    print("=" * 50)
    print("  Unified Validator - MCPD & Climate")
    print("  http://localhost:5050")
    print("=" * 50)
    app.run(port=5050, debug=True)