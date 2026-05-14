"""
mcpdPlugin.py — Plugin MCPD pour excel-validator (Mathis)
Appelle le validator.py custom (logique MCPD v4.3) et injecte
le rapport dans l'interface web de Mathis.
"""

import os
import sys
import importlib.util


def _load_validator():
    """Charge validator.py depuis le dossier parent (config/mcpd/)."""
    plugin_dir   = os.path.dirname(os.path.abspath(__file__))   # .../mcpd/plugins/
    mcpd_dir     = os.path.dirname(plugin_dir)                   # .../mcpd/
    validator_path = os.path.join(mcpd_dir, "validator.py")

    if not os.path.isfile(validator_path):
        raise FileNotFoundError(
            f"validator.py introuvable dans : {mcpd_dir}\n"
            "Assure-toi de l'avoir copié dans le dossier mcpd/"
        )

    spec   = importlib.util.spec_from_file_location("mcpd_validator", validator_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class McpdPlugin:
    """
    Plugin appelé par excel-validator après la validation standard.
    Il exécute run_validation() de ton validator.py et ajoute
    le rapport MCPD complet dans les erreurs/warnings de Mathis.
    """

    def __init__(self, **kwargs):
        self.validator_module = None
        try:
            self.validator_module = _load_validator()
        except Exception as e:
            print(f"[mcpdPlugin] WARN: impossible de charger validator.py : {e}",
                  file=sys.stderr)

    def check(self, excel_filename, **kwargs):
        """
        Appelé par excel-validator avec le chemin du fichier uploadé.
        Retourne une liste de dicts {"type": "error"|"warning", "message": str}
        """
        if self.validator_module is None:
            return [{
                "type": "warning",
                "message": "validator.py MCPD non chargé — vérification custom ignorée."
            }]

        try:
            rapport = self.validator_module.run_validation(excel_filename)
        except Exception as e:
            return [{
                "type": "error",
                "message": f"Erreur lors de l'exécution du validator MCPD : {e}"
            }]

        # Analyse le rapport texte pour extraire erreurs et warnings
        results = []
        current_section = None

        for line in rapport.splitlines():
            line_stripped = line.strip()

            if "ERREURS BLOQUANTES" in line_stripped:
                current_section = "error"
            elif "AVERTISSEMENTS QUALITÉ" in line_stripped:
                current_section = "warning"
            elif "RÉSUMÉ" in line_stripped:
                current_section = None

            # Lignes de détail (commencent par Ligne ou !!)
            if line_stripped.startswith("Ligne ") or line_stripped.startswith("!!"):
                if current_section:
                    results.append({
                        "type": current_section,
                        "message": line_stripped
                    })

        # Si aucune erreur/warning parsée, ajoute le rapport complet en info
        if not results:
            results.append({
                "type": "warning",
                "message": "Rapport MCPD complet :\n" + rapport
            })

        return results
