# -*- coding: utf-8 -*-
"""
mcpd_launcher.py
Interface web pour valider des fichiers selon :
- MCPD v4.3 (ressources génétiques)
- Climatique (données environnementales)

Usage:
    python mcpd_launcher.py

Puis ouvrir http://localhost:5050
"""

import os
import sys
import tempfile
from flask import Flask, request, render_template_string

# Chargement des deux validateurs
sys.path.insert(0, r"C:\Users\rahma\developement")
from validator import run_validation as run_mcpd
from validator_climate import run_validation as run_climate

app = Flask(__name__)

HTML = """
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Validateur de données - MCPD & Climat</title>
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
        <a href="http://localhost:8888">&larr; Retour MIAPPE</a>
    </div>

    <h1>Validateur de données normalisées</h1>
    <p>Choisissez le type de validation, puis sélectionnez un fichier Excel (.xlsx).</p>

    <div class="upload-box">
        <form method="POST" enctype="multipart/form-data">
            <div class="selector">
                <label>
                    <input type="radio" name="validation_type" value="mcpd" {% if validation_type == 'mcpd' %}checked{% endif %}>
                    🌾 MCPD (ressources génétiques)
                </label>
                <label>
                    <input type="radio" name="validation_type" value="climate" {% if validation_type == 'climate' %}checked{% endif %}>
                    🌡️ Climatique (environnemental)
                </label>
            </div>
            <br>
            <p>📎 Fichier Excel (.xlsx) :</p>
            <input type="file" name="fichier" accept=".xlsx" required><br><br>
            <button type="submit">Lancer la validation</button>
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
    validation_type = "mcpd"  # valeur par défaut

    if request.method == "POST":
        fichier = request.files.get("fichier")
        validation_type = request.form.get("validation_type", "mcpd")

        if fichier and fichier.filename.endswith(".xlsx"):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                fichier.save(tmp.name)
                tmp_path = tmp.name
            try:
                if validation_type == "mcpd":
                    rapport = run_mcpd(tmp_path)
                else:  # climate
                    rapport = run_climate(tmp_path)
                # Détection de validité : on cherche "Blocking errors      : 0" dans le rapport
                valide = "Blocking errors      : 0" in rapport
            except Exception as e:
                rapport = f"[ERREUR] {e}"
                valide = False
            finally:
                try:
                    os.unlink(tmp_path)
                except:
                    pass

    return render_template_string(HTML, rapport=rapport, valide=valide, validation_type=validation_type)


if __name__ == "__main__":
    print("=" * 50)
    print("  Validateur unifié - MCPD & Climat")
    print("  http://localhost:5050")
    print("=" * 50)
    app.run(port=5050, debug=True)