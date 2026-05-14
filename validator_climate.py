"""
validator_climate.py – Validation de données climatiques
Structure du rapport identique au validateur MCPD v4.3
Colonnes obligatoires : Name, Data Type, Location/Climate
"""

from __future__ import annotations

import os
import sys
import tempfile
import pandas as pd
from collections import defaultdict

try:
    from frictionless import Schema, Field, Resource, validate
    from frictionless import fields as fl_fields
    import frictionless
except ImportError:
    sys.exit("[FATAL] frictionless non installé.\n  pip install frictionless 'frictionless[excel]' pandas openpyxl")

# ==============================================================================
# 1. CONSTANTES CLIMATIQUES (plages de validation)
# ==============================================================================
TEMP_MIN_ABS = -90.0
TEMP_MAX_ABS = 60.0
PRECIP_MAX_ABS = 2000.0
WIND_MAX_ABS = 200.0
SUN_HOURS_MAX = 24.0

VALID_ISO3_CODES = {
    "AFG","ALB","DZA","AND","AGO","ARG","ARM","AUS","AUT","BEL","BGD","BGR","BIH",
    "BLR","BOL","BRA","BRN","CAN","CHE","CHL","CHN","CMR","COL","CRI","CUB","CYP",
    "CZE","DEU","DNK","DOM","DZA","ECU","EGY","ESP","EST","ETH","FIN","FRA","GAB",
    "GBR","GEO","GHA","GRC","GTM","HKG","HND","HRV","HUN","IDN","IND","IRL","IRN",
    "IRQ","ISL","ISR","ITA","JAM","JOR","JPN","KAZ","KEN","KGZ","KHM","KOR","KWT",
    "LAO","LBN","LBY","LKA","LTU","LUX","LVA","MAR","MDA","MDG","MEX","MLI","MLT",
    "MMR","MNE","MNG","MOZ","MRT","MWI","MYS","NAM","NCL","NGA","NIC","NLD","NOR",
    "NPL","NZL","OMN","PAK","PAN","PER","PHL","POL","PRT","PRY","QAT","ROU","RUS",
    "RWA","SAU","SDN","SEN","SGP","SLV","SOM","SRB","SUR","SVK","SVN","SWE","SYR",
    "TCD","TGO","THA","TJK","TKM","TUN","TUR","TWN","TZA","UGA","UKR","URY","USA",
    "UZB","VEN","VNM","YEM","ZAF","ZMB","ZWE",
}

# ==============================================================================
# 2. COLONNES OBLIGATOIRES (selon votre tableau)
# ==============================================================================
MANDATORY_FIELDS = ("Name", "Data Type", "Location/Climate")

# ==============================================================================
# 3. DÉTECTION DES COLONNES CLIMATIQUES (flexible)
# ==============================================================================
CLIMATE_KEYWORDS = {
    "station_id":   ["name", "station", "location", "site", "code"],
    "date":         ["date", "year", "annee", "time", "timestamp"],
    "tmin":         ["tmin", "tn", "temp_min", "min_temp"],
    "tmax":         ["tmax", "tx", "temp_max", "max_temp"],
    "precip":       ["precip", "precipitation", "rain", "pluie", "rr"],
    "latitude":     ["latitude", "lat", "y"],
    "longitude":    ["longitude", "lon", "long", "x"],
    "altitude":     ["altitude", "elevation", "alt", "z"],
    "country":      ["country", "pays", "iso3", "cntry"],
    "wind_speed":   ["wind_speed", "wind", "vent", "wspd"],
    "sun_hours":    ["sun_hours", "ensoleillement", "sunshine"],
}

def _map_columns(df: pd.DataFrame) -> dict:
    """Retourne {nom_normalisé: nom_original} pour les colonnes climatiques."""
    mapping = {}
    df_cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for standard, patterns in CLIMATE_KEYWORDS.items():
        for p in patterns:
            p_low = p.lower()
            for col_lower, orig in df_cols_lower.items():
                if p_low in col_lower or col_lower in p_low:
                    mapping[standard] = orig
                    break
            if standard in mapping:
                break
    return mapping

def _standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = _map_columns(df)
    rename_dict = {v: k for k, v in mapping.items() if v in df.columns}
    df_std = df.rename(columns=rename_dict)
    if "date" in df_std.columns:
        df_std["date"] = pd.to_datetime(df_std["date"], errors="coerce")
    numeric_fields = ["tmin", "tmax", "precip", "latitude", "longitude", "altitude", "wind_speed", "sun_hours"]
    for f in numeric_fields:
        if f in df_std.columns:
            df_std[f] = pd.to_numeric(df_std[f], errors="coerce")
    return df_std

# ==============================================================================
# 4. GESTION MULTI-FEUILLES (fusion possible)
# ==============================================================================
def _detect_join_key(sheets: dict[str, pd.DataFrame]) -> str | None:
    if len(sheets) == 1:
        return None
    all_cols = [set(df.columns) for df in sheets.values()]
    common = set.intersection(*all_cols)
    if not common:
        return None
    # Priorité aux colonnes obligatoires ou climatiques
    for priority in ["Name", "station_id", "date"]:
        if priority in common:
            return priority
    for col in sorted(common):
        if any(k in col.lower() for k in ("id", "code", "key", "name")):
            return col
    return sorted(common)[0]

def _load_multi_sheet(excel_path: str) -> tuple[pd.DataFrame, str]:
    xl = pd.ExcelFile(excel_path)
    sheets_raw = {}
    for sheet_name in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet_name)
        df.dropna(how="all", inplace=True)
        df.dropna(axis=1, how="all", inplace=True)
        if not df.empty:
            sheets_raw[sheet_name] = df
            print(f"[INFO] Feuille '{sheet_name}' : {len(df)} lignes, {len(df.columns)} colonnes", file=sys.stderr)
    if not sheets_raw:
        raise ValueError("Aucune feuille non vide trouvée.")
    # Standardisation des colonnes climatiques (sans renommer les obligatoires)
    sheets_std = {}
    for name, df in sheets_raw.items():
        df_std = _standardize_columns(df)
        # Conserver les colonnes obligatoires telles quelles (pas renommées)
        sheets_std[name] = df_std
    join_key = _detect_join_key(sheets_std)
    if join_key is None:
        print("[WARN] Aucune colonne commune trouvée pour fusionner. Utilisation de la première feuille uniquement.", file=sys.stderr)
        df_merged = next(iter(sheets_std.values()))
    else:
        print(f"[INFO] Fusion sur la clé : '{join_key}'", file=sys.stderr)
        base_name, df_merged = next(iter(sheets_std.items()))
        df_merged = df_merged.drop_duplicates(subset=[join_key])
        for name, df_other in list(sheets_std.items())[1:]:
            df_other = df_other.drop_duplicates(subset=[join_key])
            overlap = (set(df_merged.columns) & set(df_other.columns)) - {join_key}
            if overlap:
                df_merged = df_merged.rename(columns={c: f"{c}_{base_name[:3]}" for c in overlap})
                df_other = df_other.rename(columns={c: f"{c}_{name[:3]}" for c in overlap})
            df_merged = df_merged.merge(df_other, on=join_key, how="outer")
        print(f"[INFO] Fusion terminée : {len(df_merged)} lignes, {len(df_merged.columns)} colonnes", file=sys.stderr)
    fd, csv_path = tempfile.mkstemp(suffix=".csv", dir=".", prefix="climate_")
    os.close(fd)
    df_merged.to_csv(csv_path, index=False, encoding="utf-8")
    return df_merged, csv_path

# ==============================================================================
# 5. SCHÉMA FRICTIONLESS (uniquement pour colonnes climatiques reconnues)
# ==============================================================================
def _build_climate_schema(df_columns: list) -> Schema:
    field_defs = {
        "station_id": fl_fields.StringField(name="station_id"),
        "date": fl_fields.DateField(name="date"),
        "tmin": fl_fields.NumberField(name="tmin", constraints={"minimum": TEMP_MIN_ABS, "maximum": TEMP_MAX_ABS}),
        "tmax": fl_fields.NumberField(name="tmax", constraints={"minimum": TEMP_MIN_ABS, "maximum": TEMP_MAX_ABS}),
        "precip": fl_fields.NumberField(name="precip", constraints={"minimum": 0, "maximum": PRECIP_MAX_ABS}),
        "latitude": fl_fields.NumberField(name="latitude", constraints={"minimum": -90, "maximum": 90}),
        "longitude": fl_fields.NumberField(name="longitude", constraints={"minimum": -180, "maximum": 180}),
        "altitude": fl_fields.NumberField(name="altitude", constraints={"minimum": -500, "maximum": 9000}),
        "country": fl_fields.StringField(name="country", constraints={"enum": sorted(VALID_ISO3_CODES)}),
        "wind_speed": fl_fields.NumberField(name="wind_speed", constraints={"minimum": 0, "maximum": WIND_MAX_ABS}),
        "sun_hours": fl_fields.NumberField(name="sun_hours", constraints={"minimum": 0, "maximum": SUN_HOURS_MAX}),
    }
    fields = []
    for fname, fobj in field_defs.items():
        if fname in df_columns:
            fields.append(fobj)
    return Schema(fields=fields)

# ==============================================================================
# 6. VÉRIFICATIONS MÉTIER
# ==============================================================================
def _check_mandatory_fields(df: pd.DataFrame) -> list[dict]:
    """Vérifie la présence des colonnes obligatoires (Name, Data Type, Location/Climate)."""
    issues = []
    for field in MANDATORY_FIELDS:
        if field not in df.columns:
            issues.append({
                "row": 0,
                "accenumb": "[STRUCTURE]",
                "category": "MISSING_COLUMN",
                "field": field,
                "level": "ERROR",
                "message": f"Colonne obligatoire '{field}' absente. Données climatiques exigent Name, Data Type et Location/Climate."
            })
        else:
            # Vérifier les valeurs manquantes dans la colonne (optionnel mais utile)
            missing = df[field].isna().sum()
            if missing > 0:
                issues.append({
                    "row": 0,
                    "accenumb": "[STRUCTURE]",
                    "category": "MISSING_VALUES",
                    "field": field,
                    "level": "WARNING",
                    "message": f"Colonne '{field}' présente mais {missing} ligne(s) avec valeur manquante."
                })
    return issues

def _check_tmin_tmax_consistency(df: pd.DataFrame) -> list[dict]:
    issues = []
    if "tmin" in df.columns and "tmax" in df.columns:
        for idx, row in df.iterrows():
            tmin = row.get("tmin")
            tmax = row.get("tmax")
            if pd.notna(tmin) and pd.notna(tmax) and tmin > tmax:
                station = row.get("station_id", f"ligne{idx+2}")
                issues.append({
                    "row": idx+2, "accenumb": str(station),
                    "category": "TMIN_TMAX_INCONSISTENCY", "field": "tmin/tmax", "level": "ERROR",
                    "message": f"tmin ({tmin}) > tmax ({tmax}) : valeur impossible."
                })
    return issues

def _check_negative_precip(df: pd.DataFrame) -> list[dict]:
    issues = []
    if "precip" in df.columns:
        for idx, row in df.iterrows():
            p = row.get("precip")
            if pd.notna(p) and p < 0:
                station = row.get("station_id", f"ligne{idx+2}")
                issues.append({
                    "row": idx+2, "accenumb": str(station),
                    "category": "NEGATIVE_PRECIP", "field": "precip", "level": "ERROR",
                    "message": f"Précipitation négative : {p}."
                })
    return issues

def _check_date_order(df: pd.DataFrame) -> list[dict]:
    issues = []
    if "date" in df.columns and "station_id" in df.columns:
        grouped = df.sort_values(["station_id", "date"]).groupby("station_id")
        for station, group in grouped:
            dates = group["date"].dropna()
            if len(dates) > 1 and not dates.is_monotonic_increasing:
                issues.append({
                    "row": 0, "accenumb": str(station),
                    "category": "DATE_ORDER", "field": "date", "level": "WARNING",
                    "message": f"Dates non croissantes pour la station {station}."
                })
    return issues

def _check_coordinates_range(df: pd.DataFrame) -> list[dict]:
    issues = []
    for coord in ["latitude", "longitude"]:
        if coord in df.columns:
            for idx, row in df.iterrows():
                val = row.get(coord)
                if pd.notna(val):
                    station = row.get("station_id", f"ligne{idx+2}")
                    if coord == "latitude" and (val < -90 or val > 90):
                        issues.append({
                            "row": idx+2, "accenumb": str(station),
                            "category": "LATITUDE_RANGE", "field": coord, "level": "ERROR",
                            "message": f"Latitude {val} hors de [-90,90]."
                        })
                    elif coord == "longitude" and (val < -180 or val > 180):
                        issues.append({
                            "row": idx+2, "accenumb": str(station),
                            "category": "LONGITUDE_RANGE", "field": coord, "level": "ERROR",
                            "message": f"Longitude {val} hors de [-180,180]."
                        })
    return issues

def _check_country_code(df: pd.DataFrame) -> list[dict]:
    issues = []
    if "country" in df.columns:
        for idx, row in df.iterrows():
            c = row.get("country")
            if pd.notna(c) and str(c).strip().upper() not in VALID_ISO3_CODES:
                station = row.get("station_id", f"ligne{idx+2}")
                issues.append({
                    "row": idx+2, "accenumb": str(station),
                    "category": "INVALID_COUNTRY", "field": "country", "level": "WARNING",
                    "message": f"Code pays '{c}' non reconnu (ISO 3166-1 alpha-3 attendu)."
                })
    return issues

# ==============================================================================
# 7. RAPPORT STYLE MCPD (identique en présentation)
# ==============================================================================
def _format_report(excel_path: str, df: pd.DataFrame, frictionless_report, extra_issues: list[dict]) -> str:
    W = 72
    def rule(c="="): return c * W
    def box(text):
        pad = W - 4
        return [
            "+" + "-" * (W - 2) + "+",
            "|  " + text.ljust(pad) + "|",
            "+" + "-" * (W - 2) + "+",
        ]

    # Erreurs frictionless (types, plages, enum)
    fl_issues = []
    if frictionless_report:
        for task in (frictionless_report.tasks or []):
            for err in (task.errors or []):
                row_num = getattr(err, "row_number", None)
                if row_num is None:
                    continue
                code = getattr(err, "code", "")
                field = str(getattr(err, "field_name", "") or "")
                note = getattr(err, "note", "")
                if row_num < len(df):
                    acc = df.iloc[row_num].get("station_id", f"row{row_num+1}")
                else:
                    acc = f"row{row_num+1}"
                fl_issues.append({
                    "row": row_num + 1,
                    "accenumb": str(acc),
                    "category": code.upper(),
                    "field": field,
                    "level": "ERROR",
                    "message": note or f"Erreur {code} sur {field}"
                })

    all_issues = fl_issues + extra_issues
    errors = [i for i in all_issues if i["level"] == "ERROR"]
    warnings = [i for i in all_issues if i["level"] == "WARNING"]

    errors_by_cat = defaultdict(list)
    for e in errors: errors_by_cat[e["category"]].append(e)
    warnings_by_cat = defaultdict(list)
    for w in warnings: warnings_by_cat[w["category"]].append(w)

    n_err = len(errors)
    n_warn = len(warnings)
    status = "INVALIDE" if errors else "VALIDE"

    lines = []
    lines += [
        rule("="),
        "  CLIMATE DATA VALIDATION REPORT — MCPD‑like structure",
        f"(frictionless v{frictionless.__version__})",
        rule("="),
        f"  File                 : {excel_path}",
        f"  Total rows           : {len(df)}",
        f"  Status               : {status}",
        f"  Blocking errors      : {n_err}",
        f"  Warnings             : {n_warn}",
        rule("="),
    ]

    # Règles de validation (rappel)
    lines += ["", rule("-"), "  VALIDATION RULES (only if columns exist)", rule("-")]
    lines.append("  * Mandatory columns : Name, Data Type, Location/Climate")
    lines.append("  * tmin & tmax : numeric, tmin ≤ tmax, ranges [-90,60]")
    lines.append("  * precip      : ≥ 0, max 2000 mm/day")
    lines.append("  * latitude    : [-90,90], longitude [-180,180]")
    lines.append("  * date        : parsable, order warning per station")
    lines.append("  * country     : ISO 3166-1 alpha-3 if provided")
    lines.append("  * Other columns are ignored")
    lines.append(rule("="))

    # Erreurs
    if errors:
        lines += [""] + box("BLOCKING ERRORS — Data cannot be ingested as is") + [""]
        for cat, cat_issues in sorted(errors_by_cat.items()):
            n_aff = len({i["row"] for i in cat_issues if i["row"] != 0})
            n_struct = len([i for i in cat_issues if i["row"] == 0])
            suffix = "  [STRUCTURAL]" if n_struct > 0 else ""
            lines.append(f"  [{cat}]  {n_aff} affected row(s){suffix}")
            lines.append(rule("-"))
            for iss in cat_issues[:15]:  # un peu plus pour les détails
                lines.append(
                    f"  Row {str(iss['row']).rjust(4)}  |  "
                    f"Identifier: {iss['accenumb'][:25].ljust(25)}  |  "
                    f"Field: {iss['field']}"
                )
                for ml in iss["message"].split("\n"):
                    lines.append("              " + ml)
            if len(cat_issues) > 15:
                lines.append(f"              ... and {len(cat_issues)-15} more")
            lines.append("")
    else:
        lines += ["", "  ✅ No critical errors detected.", ""]

    # Avertissements
    if warnings:
        lines += [rule("=")] + box("WARNINGS — Quality recommendations") + [""]
        for cat, cat_issues in sorted(warnings_by_cat.items()):
            n_aff = len({i["row"] for i in cat_issues})
            lines.append(f"  [{cat}]  {n_aff} affected row(s)")
            lines.append(rule("-"))
            for iss in cat_issues[:15]:
                lines.append(
                    f"  Row {str(iss['row']).rjust(4)}  |  "
                    f"Identifier: {iss['accenumb'][:25].ljust(25)}  |  "
                    f"Field: {iss['field']}"
                )
                for ml in iss["message"].split("\n"):
                    lines.append("              " + ml)
            if len(cat_issues) > 15:
                lines.append(f"              ... and {len(cat_issues)-15} more")
            lines.append("")
    else:
        lines += ["", "  ✅ No warnings detected.", ""]

    # Résumé
    lines += [
        rule("="),
        "  SUMMARY",
        rule("-"),
        f"  Rows analyzed        : {len(df)}",
        f"  Blocking errors      : {n_err}",
        f"  Warnings             : {n_warn}",
    ]
    if errors_by_cat:
        lines.append("    . Error categories:")
        for cat in sorted(errors_by_cat.keys()):
            lines.append(f"        - {cat}: {len(errors_by_cat[cat])}")
    if warnings_by_cat:
        lines.append("    . Warning categories:")
        for cat in sorted(warnings_by_cat.keys()):
            lines.append(f"        - {cat}: {len(warnings_by_cat[cat])}")
    lines.append(rule("="))

    return "\n".join(lines)

# ==============================================================================
# 8. FONCTION PRINCIPALE EXPOSÉE
# ==============================================================================
def run_validation(excel_path: str) -> str:
    try:
        df, csv_path = _load_multi_sheet(excel_path)
    except Exception as e:
        return f"[FATAL] Failed to load file: {e}"

    # Appliquer le schéma frictionless uniquement pour les colonnes climatiques reconnues
    schema = _build_climate_schema(df.columns.tolist())
    frictionless_report = None
    try:
        resource = Resource(path=csv_path, schema=schema)
        frictionless_report = validate(resource)
    except Exception as e:
        print(f"[WARN] frictionless validate() : {e}", file=sys.stderr)
    finally:
        try:
            os.unlink(csv_path)
        except OSError:
            pass

    extra = []
    extra.extend(_check_mandatory_fields(df))
    extra.extend(_check_tmin_tmax_consistency(df))
    extra.extend(_check_negative_precip(df))
    extra.extend(_check_date_order(df))
    extra.extend(_check_coordinates_range(df))
    extra.extend(_check_country_code(df))

    return _format_report(excel_path, df, frictionless_report, extra)

# ==============================================================================
# 9. EXÉCUTION DIRECTE (test)
# ==============================================================================
if __name__ == "__main__":
    if len(sys.argv) > 1:
        print(run_validation(sys.argv[1]))
    else:
        print("Usage: python validator_climate.py <fichier_climat.xlsx>")