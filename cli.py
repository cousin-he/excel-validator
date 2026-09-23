"""
cli.py — Interface ligne de commande pour le validateur MCPD v4.3

Usage
-----
    python cli.py --file dataset.xlsx --output rapport.txt
    python cli.py --file dataset.xlsx        # affiche dans le terminal

Prérequis
---------
    pip install frictionless "frictionless[excel]" pandas openpyxl
"""

import argparse
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from validator import run_validation   # ← nom corrigé v4.3


def main() -> None:
    parser = argparse.ArgumentParser(
        description="MCPD Validator v4.3 — frictionless edition",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemples :
  python cli.py --file dataset.xlsx --output rapport.txt
  python cli.py --file dataset.xlsx
        """,
    )
    parser.add_argument("--file",   required=True,  help="Fichier Excel à valider (.xlsx)")
    parser.add_argument("--output", required=False, help="Fichier de sortie (.txt). Si absent : terminal.")

    args = parser.parse_args()

    if not os.path.isfile(args.file):
        sys.exit(f"[ERREUR] Fichier introuvable : {args.file}")

    result = run_validation(args.file)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(result)
        print(f"[OK] Rapport sauvegardé → {args.output}")
    else:
        print(result)


if __name__ == "__main__":
    main()