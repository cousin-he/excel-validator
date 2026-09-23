# autre_api.py — "l'autre API" qui reçoit / lit le rapport du validateur
# Elle montre les 3 façons de récupérer le rapport :
#   (A) POST HTTP envoyé par le worker  -> /reports
#   (B) Lecture directe dans Redis      -> /lire/<job_id>
#   (C) Écoute pub/sub Redis            -> thread en arrière-plan
import json
import threading
from flask import Flask, request, jsonify
from redis import Redis

REDIS_HOST = "localhost"
REDIS_PORT = 6379
PUBSUB_CHANNEL = "validation_reports"

app = Flask(__name__)
r = Redis(host=REDIS_HOST, port=REDIS_PORT)


# (A) Reçoit le POST envoyé par le worker
@app.route("/reports", methods=["POST"])
def recevoir():
    data = request.get_json()
    print("[POST reçu]  job_id =", data["job_id"], "| valid =", data["valid"])
    return jsonify({"ok": True})


# (B) Lit le rapport dans Redis quand on veut
@app.route("/lire/<job_id>")
def lire(job_id):
    raw = r.get(f"report:{job_id}")
    if not raw:
        return jsonify({"error": "introuvable"}), 404
    return jsonify(json.loads(raw))


# (C) Écoute les notifications pub/sub
def ecouter():
    p = Redis(host=REDIS_HOST, port=REDIS_PORT).pubsub()
    p.subscribe(PUBSUB_CHANNEL)
    for msg in p.listen():
        if msg["type"] == "message":
            d = json.loads(msg["data"])
            print("[PUBSUB reçu] job_id =", d["job_id"], "| valid =", d["valid"])


if __name__ == "__main__":
    threading.Thread(target=ecouter, daemon=True).start()
    # use_reloader=False pour éviter de lancer le thread pub/sub deux fois
    app.run(port=6000, debug=True, use_reloader=False)
