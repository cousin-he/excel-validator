# tasks.py — enveloppe rq autour des validateurs existants (inchangés)
# Le worker :
#   1) exécute le validateur
#   2) stocke le rapport dans Redis (clé "report:<job_id>")
#   3) publie une notification Redis (pub/sub)
#   4) envoie le rapport en HTTP POST vers l'autre API
import os
import sys
import json
import requests
from redis import Redis
from rq import get_current_job

sys.path.insert(0, r"C:\Users\rahma\developement")  # même chemin que mcpd_launcher.py

from validator import run_validation as run_mcpd
from validator_climate import run_validation as run_climate

# ---------------------------------------------------------------------------
# Configuration  (à adapter)
# ---------------------------------------------------------------------------
REDIS_HOST = "localhost"
REDIS_PORT = 6379
AUTRE_API_URL = "http://localhost:6000/reports"   # <-- adresse de l'autre API
REPORT_TTL = 86400                                # durée de vie du rapport : 24h
PUBSUB_CHANNEL = "validation_reports"

redis_conn = Redis(host=REDIS_HOST, port=REDIS_PORT)


def envoyer_rapport(job_id: str, validation_type: str, result: dict) -> None:
    """Distribue le rapport : Redis (stockage) + pub/sub + HTTP POST."""
    payload = {"job_id": job_id, "validation_type": validation_type, **result}
    data = json.dumps(payload)

    # 1) Stockage Redis : c'est la garantie, l'autre API peut lire quand elle veut
    redis_conn.set(f"report:{job_id}", data, ex=REPORT_TTL)

    # 2) Pub/Sub : notification instantanée (perdue si personne n'écoute)
    redis_conn.publish(PUBSUB_CHANNEL, data)

    # 3) HTTP POST : push direct vers l'autre API
    try:
        r = requests.post(AUTRE_API_URL, json=payload, timeout=10)
        r.raise_for_status()
    except requests.RequestException as e:
        # Un échec d'envoi ne doit pas faire échouer la validation :
        # le rapport reste disponible dans Redis.
        print(f"[WARN] envoi HTTP impossible : {e}")


def valider_job(chemin_fichier: str, validation_type: str) -> dict:
    """Exécuté par le worker rq."""
    job = get_current_job()
    try:
        if validation_type == "climate":
            report = run_climate(chemin_fichier)
        else:
            report = run_mcpd(chemin_fichier)
        valid = "Blocking errors      : 0" in report
        result = {"valid": valid, "report": report}
    except Exception as e:
        result = {"valid": False, "report": f"[ERROR] {e}"}
    finally:
        # nettoyage du fichier temporaire
        try:
            os.remove(chemin_fichier)
        except OSError:
            pass

    envoyer_rapport(job.id, validation_type, result)
    return result
