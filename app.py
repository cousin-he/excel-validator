# app.py — API REST asynchrone (remplace la partie web de mcpd_launcher.py)
import os
import json
import tempfile
from flask import Flask, request, jsonify, render_template_string
from redis import Redis
from rq import Queue
from rq.job import Job

from tasks import valider_job

app = Flask(__name__)
redis_conn = Redis()          # se connecte à localhost:6379 par défaut
queue = Queue(connection=redis_conn)

RESULT_TTL = 86400            # le résultat rq reste 24h dans Redis

HTML = """
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Excel Validator (async)</title>
<style>
body { font-family: Arial, sans-serif; max-width: 900px; margin: 40px auto; padding: 20px; }
h1 { color: #2c3e50; }
.upload-box { border: 2px dashed #3498db; border-radius: 10px; padding: 30px; text-align: center; margin: 20px 0; background: #f8f9fa; }
button { background: #3498db; color: white; border: none; padding: 10px 25px; border-radius: 5px; cursor: pointer; font-size: 16px; }
button:disabled { background: #95a5a6; cursor: not-allowed; }
.selector { margin: 15px 0; }
.selector label { margin-right: 20px; }
.report { background: #1e1e1e; color: #d4d4d4; padding: 20px; border-radius: 8px; font-family: monospace; font-size: 13px; white-space: pre-wrap; margin-top: 20px; max-height: 600px; overflow-y: auto; display: none; }
.badge { display: none; margin-bottom: 15px; font-size: 18px; padding: 10px 20px; border-radius: 5px; font-weight: bold; color: white; }
.valid { background: #27ae60; }
.invalid { background: #e74c3c; }
.status { margin-top: 15px; font-style: italic; color: #7f8c8d; }
</style>
</head>
<body>
<h1>Excel Validator (async)</h1>
<p>Choisissez un type de validation et un fichier Excel (.xlsx).</p>

<div class="upload-box">
  <form id="form">
    <div class="selector">
      <label><input type="radio" name="validation_type" value="mcpd" checked> MCPD</label>
      <label><input type="radio" name="validation_type" value="climate"> Climate</label>
    </div>
    <input type="file" id="fichier" accept=".xlsx" required><br><br>
    <button type="submit" id="btn">Lancer la validation</button>
  </form>
  <div class="status" id="status"></div>
</div>

<div class="badge" id="badge"></div>
<div class="report" id="report"></div>

<script>
const form = document.getElementById('form');
const btn = document.getElementById('btn');
const statusEl = document.getElementById('status');
const badge = document.getElementById('badge');
const report = document.getElementById('report');

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  badge.style.display = 'none';
  report.style.display = 'none';
  btn.disabled = true;

  const fd = new FormData();
  fd.append('fichier', document.getElementById('fichier').files[0]);
  fd.append('validation_type', form.validation_type.value);

  statusEl.textContent = 'Envoi du fichier...';
  const res = await fetch('/validate', { method: 'POST', body: fd });
  const data = await res.json();

  if (!res.ok) {
    statusEl.textContent = 'Erreur : ' + (data.error || 'inconnue');
    btn.disabled = false;
    return;
  }

  statusEl.textContent = 'Job en cours (id: ' + data.job_id + ')...';
  poll(data.job_id);
});

async function poll(jobId) {
  const res = await fetch('/result/' + jobId);
  const data = await res.json();

  if (data.status === 'done') {
    statusEl.textContent = 'Terminé. (job id: ' + jobId + ')';
    badge.textContent = data.valid ? 'VALIDE' : 'INVALIDE';
    badge.className = 'badge ' + (data.valid ? 'valid' : 'invalid');
    badge.style.display = 'inline-block';
    report.textContent = data.report;
    report.style.display = 'block';
    btn.disabled = false;
  } else if (data.status === 'failed') {
    statusEl.textContent = 'Échec du job.';
    btn.disabled = false;
  } else {
    statusEl.textContent = 'Statut : ' + data.status + '...';
    setTimeout(() => poll(jobId), 1500);
  }
}
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/validate", methods=["POST"])
def validate():
    file = request.files.get("fichier")
    validation_type = request.form.get("validation_type", "mcpd")

    if not file or not file.filename.endswith(".xlsx"):
        return jsonify({"error": "fichier .xlsx requis"}), 400

    # fichier temporaire (supprimé par le worker après validation)
    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    file.save(tmp_path)

    job = queue.enqueue(
        valider_job, tmp_path, validation_type,
        job_timeout=600,
        result_ttl=RESULT_TTL,
    )
    return jsonify({"job_id": job.id, "status": "pending"}), 202


@app.route("/result/<job_id>")
def result(job_id):
    """Résultat via rq (utilisé par la page web pour le polling)."""
    try:
        job = Job.fetch(job_id, connection=redis_conn)
    except Exception:
        return jsonify({"error": "job introuvable"}), 404

    if job.is_finished:
        return jsonify({"status": "done", **job.result})
    if job.is_failed:
        return jsonify({"status": "failed", "error": str(job.exc_info)}), 500
    return jsonify({"status": job.get_status()})


@app.route("/report/<job_id>")
def report(job_id):
    """Rapport lu directement dans la clé Redis 'report:<job_id>'."""
    data = redis_conn.get(f"report:{job_id}")
    if not data:
        return jsonify({"error": "rapport introuvable"}), 404
    return jsonify(json.loads(data))


if __name__ == "__main__":
    app.run(port=5050, debug=True)
