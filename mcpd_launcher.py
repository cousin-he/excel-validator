# -*- coding: utf-8 -*-
"""
mcpd_launcher.py
Lance une interface web simple pour valider des fichiers MCPD
en utilisant ton validator.py existant.

Usage:
    python mcpd_launcher.py

Puis ouvre http://localhost:5050
"""

import os
import sys
import tempfile
from flask import Flask, request, render_template_string

# Charge ton validator.py
sys.path.insert(0, r"C:\Users\rahma\developement")
from validator import run_validation

app = Flask(__name__)

HTML = """
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>MCPD Validator v4.3</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 900px; margin: 40px auto; padding: 20px; }
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
        .rapport {
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
    </style>
</head>
<body>
    <div class="nav">
        <strong>Excel Validator</strong>
        <a href="http://localhost:8888">&larr; Retour MIAPPE</a>
    </div>

    <h1>MCPD v2.1 &mdash; Germplasm Passport Validator</h1>
    <p>Validate your Excel files according to the MCPD v2.1 standard (Multi-Crop Passport Descriptors).</p>

    <div class="upload-box">
        <form method="POST" enctype="multipart/form-data">
            <p>Please select your Excel file (.xlsx)</p>
            <input type="file" name="fichier" accept=".xlsx" required><br><br>
            <button type="submit">Valider</button>
        </form>
    </div>

    {% if rapport %}
        <div class="badge">
            {% if valide %}
                <span class="valid">VALIDE</span>
            {% else %}
                <span class="invalid">INVALIDE</span>
            {% endif %}
        </div>
        <div class="rapport">{{ rapport }}</div>
    {% endif %}
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    rapport = None
    valide = False

    if request.method == "POST":
        fichier = request.files.get("fichier")
        if fichier and fichier.filename.endswith(".xlsx"):
            # Sauvegarde temporaire
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                fichier.save(tmp.name)
                tmp_path = tmp.name
            try:
                rapport = run_validation(tmp_path)
                valide = "Blocking errors   : 0" in rapport
            except Exception as e:
                rapport = f"[ERREUR] {e}"
                valide = False
            finally:
                try:
                    os.unlink(tmp_path)
                except:
                    pass

    return render_template_string(HTML, rapport=rapport, valide=valide)


if __name__ == "__main__":
    print("=" * 50)
    print("  MCPD Validator v4.3")
    print("  http://localhost:5050")
    print("=" * 50)
    app.run(port=5050, debug=True)
    