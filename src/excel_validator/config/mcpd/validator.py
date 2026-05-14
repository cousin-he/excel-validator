"""
MCPD Validator Pipeline v4.3 — corrections ACCENUMB + unicité composite
========================================================================

CORRECTIONS v4.3 par rapport à v4.2 :
  1. ACCENUMB est du texte LIBRE (free text). La regex ^[A-Z]{1,4}\\s?\\d{1,9}$
     est SUPPRIMÉE. Des valeurs comme 'KU 11483a', 'KU 11635b', 'Ae. Biuncialis - MVGB - 378'
     sont parfaitement valides. Il n'y a RIEN à valider sur le format de l'ACCENUMB seul.

  2. UNICITÉ : la contrainte d'unicité porte sur la COMBINAISON
     (INSTCODE + GENUS + ACCENUMB), pas sur ACCENUMB seul.
     Deux accessions peuvent avoir le même numéro dans des institutions ou genres
     différents — c'est normal dans les banques de gènes.

  3. cli.py importe 'validator_frictionless' — le fichier est donc renommé
     validator_frictionless.py pour correspondre.

  4. Toutes les autres logiques (frictionless, CATEGORICAL, coordonnées,
     ORIGCTY, triade qualité, DMS, GENUS/SPECIES/SUBTAXA) restent identiques à v4.2.
"""

from __future__ import annotations

import re
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
    sys.exit(
        "[FATAL] frictionless non installé.\n"
        "  pip install frictionless 'frictionless[excel]'"
    )


# ==============================================================================
# SECTION 1 — VALEURS CATÉGORIQUES (Germinate CATEGORICAL)
# ==============================================================================

MLSSTAT_VALUES: dict[int, str] = {
    10: "In the Multilateral System (MLS)",
    20: "Not in the Multilateral System (MLS)",
    30: "Unspecified",
}

SAMPSTAT_VALUES: dict[int, str] = {
    100: "Wild",
    110: "Natural",
    120: "Semi-natural/wild",
    130: "Semi-natural/sown",
    200: "Weedy",
    300: "Traditional cultivar/landrace",
    400: "Breeding/research material",
    410: "Breeders' line",
    411: "Synthetic population",
    412: "Hybrid",
    413: "Founder stock/base population",
    414: "Inbred line (parent of hybrid)",
    415: "Segregating population",
    416: "Clonal selection",
    420: "Genetic stock",
    421: "Mutant (e.g. induced/insertion mutant, tilling population)",
    422: "Cytogenetic stocks (e.g. chromosome addition/substitution lines)",
    423: "Other genetic stocks (e.g. mapping populations)",
    500: "Advanced or improved cultivar (modern cultivar)",
    600: "GMO (Genetically Modified Organism)",
    999: "Other (elaborate in REMARKS field)",
}

STORAGE_VALUES: dict[int, str] = {
    10: "Seed collection",
    11: "Short-term seed storage",
    12: "Medium-term seed storage",
    13: "Long-term seed storage",
    20: "Field collection",
    30: "In vitro collection",
    40: "Cryopreservation",
    50: "DNA collection",
    99: "Other (elaborate in REMARKS field)",
}

COLLSRC_VALUES: dict[int, str] = {
    10: "Wild habitat",
    11: "Forest or woodland",
    12: "Shrubland",
    13: "Grassland",
    14: "Desert or tundra",
    15: "Aquatic habitat",
    20: "Farm or cultivated habitat",
    21: "Field",
    22: "Orchard",
    23: "Backyard/kitchen garden",
    24: "Fallow land",
    25: "Pasture",
    26: "Cultivated land (unspecified)",
    27: "Meadow",
    28: "Other farm habitat",
    30: "Market or shop",
    40: "Institute, Experimental station, Research organisation",
    50: "Seed company",
    60: "Weedy/disturbed habitat",
    61: "Roadside",
    62: "Field margin",
    99: "Other (elaborate in REMARKS field)",
}

CATEGORICAL_FIELDS: dict[str, dict[int, str]] = {
    "MLSSTAT":  MLSSTAT_VALUES,
    "SAMPSTAT": SAMPSTAT_VALUES,
    "STORAGE":  STORAGE_VALUES,
    "COLLSRC":  COLLSRC_VALUES,
}


# ==============================================================================
# SECTION 2 — CODES PAYS ISO 3166-1 ALPHA-3 (champ ORIGCTY)
# ==============================================================================

VALID_ISO3_CODES: set[str] = {
    "AFG","ALB","DZA","ASM","AND","AGO","AIA","ATA","ATG","ARG","ARM","ABW","AUS","AUT",
    "AZE","BHS","BHR","BGD","BRB","BLR","BEL","BLZ","BEN","BMU","BTN","BOL","BES","BIH",
    "BWA","BVT","BRA","IOT","BRN","BGR","BFA","BDI","CPV","KHM","CMR","CAN","CYM","CAF",
    "TCD","CHL","CHN","CXR","CCK","COL","COM","COD","COG","COK","CRI","HRV","CUB","CUW",
    "CYP","CZE","CIV","DNK","DJI","DMA","DOM","ECU","EGY","SLV","GNQ","ERI","EST","SWZ",
    "ETH","FLK","FRO","FJI","FIN","FRA","GUF","PYF","ATF","GAB","GMB","GEO","DEU","GHA",
    "GIB","GRC","GRL","GRD","GLP","GUM","GTM","GGY","GIN","GNB","GUY","HTI","HMD","VAT",
    "HND","HKG","HUN","ISL","IND","IDN","IRN","IRQ","IRL","IMN","ISR","ITA","JAM","JPN",
    "JEY","JOR","KAZ","KEN","KIR","PRK","KOR","KWT","KGZ","LAO","LVA","LBN","LSO","LBR",
    "LBY","LIE","LTU","LUX","MAC","MDG","MWI","MYS","MDV","MLI","MLT","MHL","MTQ","MRT",
    "MUS","MYT","MEX","FSM","MDA","MCO","MNG","MNE","MSR","MAR","MOZ","MMR","NAM","NRU",
    "NPL","NLD","NCL","NZL","NIC","NER","NGA","NIU","NFK","MKD","MNP","NOR","OMN","PAK",
    "PLW","PSE","PAN","PNG","PRY","PER","PHL","PCN","POL","PRT","PRI","QAT","ROU","RUS",
    "RWA","REU","BLM","SHN","KNA","LCA","MAF","SPM","VCT","WSM","SMR","STP","SAU","SEN",
    "SRB","SYC","SLE","SGP","SXM","SVK","SVN","SLB","SOM","ZAF","SGS","SSD","ESP","LKA",
    "SDN","SUR","SJM","SWE","CHE","SYR","TWN","TJK","TZA","THA","TLS","TGO","TKL","TON",
    "TTO","TUN","TUR","TKM","TCA","TUV","UGA","UKR","ARE","GBR","UMI","USA","URY","UZB",
    "VUT","VEN","VNM","VGB","VIR","WLF","ESH","YEM","ZMB","ZWE",
    # Codes historiques (fréquents dans les banques de gènes — décision Alexandre)
    "XKX",  # Kosovo
    "SCG",  # Serbie-et-Monténégro
    "YUG",  # Yougoslavie
    "CSK",  # Tchécoslovaquie
    "DDR",  # Allemagne de l'Est
    "SUN",  # URSS
}


# ==============================================================================
# SECTION 3 — PATTERNS ET CONSTANTES DE VALIDATION
# ==============================================================================

# ── NOTE v4.3 ──────────────────────────────────────────────────────────────────
# ACCENUMB_RE EST SUPPRIMÉ.
# L'ACCENUMB est un texte LIBRE assigné par l'institution qui tient la collection.
# Des exemples valides : 'KU 11483a', 'KU 11635b', 'PI 113869',
#                        'Ae. Biuncialis - MVGB - 378', 'MG 00123'.
# La SEULE vérification possible sur l'ACCENUMB seul est sa présence (non vide).
# L'unicité est vérifiée sur la COMBINAISON (INSTCODE + GENUS + ACCENUMB).
# ──────────────────────────────────────────────────────────────────────────────

# LATITUDE DMS : DDMMSS + N ou S (7 caractères)
LATITUDE_DMS_RE  = re.compile(r"^\d{2}\d{2}\d{2}[NS]$")

# LONGITUDE DMS : DDDMMSS + E ou W (8 caractères)
LONGITUDE_DMS_RE = re.compile(r"^\d{3}\d{2}\d{2}[EW]$")

# SPECIES : tout en minuscules ou "sp."
SPECIES_RE = re.compile(r"^([a-z]+(\s[a-z]+)*|sp\.)$")

# SUBTAXA : préfixes autorisés par le standard MCPD
VALID_SUBTAXA_PREFIXES = ("subsp.", "convar.", "var.", "f.", "Group")

# Champs dont la présence est OBLIGATOIRE (décision Stephan + Alexandre)
# INSTCODE et GENUS sont obligatoires en pratique (mid-term triad).
# Note : pour ce dataset INSTCODE et GENUS sont absents -> quality warnings.
MANDATORY_FIELDS = ("INSTCODE", "ACCENUMB", "GENUS")

# Mots-clés MCPD pour détection automatique de l'en-tête
MCPD_KNOWN_COLUMNS: set[str] = {
    "INSTCODE", "ACCENUMB", "ACCENAME", "GENUS", "SPECIES", "SUBTAXA",
    "LATITUDE", "LONGITUDE", "DECLATITUDE", "DECLONGITUDE",
    "COORDUNCERT", "GEOREFMETH",
    "COLLDATE", "COLLSITE", "COLLNUMB", "COLLCODE", "COLLNAME",
    "ORIGCTY", "ORIGCTYNAME",
    "SAMPSTAT", "MLSSTAT", "STORAGE", "COLLSRC",
    "ACQDATE", "DONORCODE", "DONORNUMB", "DONORNAME",
    "BREDCODE", "BREDNAME",
    "REMARKS", "PUID", "ACCEURL",
    # Colonnes métier de ce dataset
    "GENOTYPE", "GENOTYPE NAME", "UCP CODE", "UCP_CODE",
    "ACCESSION", "ACCESSION NUMBER", "ACCESSION NAME",
}

MIN_HEADER_HITS = 2


# ==============================================================================
# SECTION 4 — DÉTECTION AUTOMATIQUE DE L'EN-TÊTE
# ==============================================================================

def _detect_header_row(xl: pd.ExcelFile, sheet_name: str, max_scan: int = 20) -> int:
    try:
        df_raw = xl.parse(sheet_name, header=None, nrows=max_scan, dtype=str)
    except Exception as e:
        print(f"[WARN] Scan header '{sheet_name}' : {e}", file=sys.stderr)
        return 0

    for row_idx, row in df_raw.iterrows():
        cells = {str(v).strip().upper() for v in row.dropna()}
        hits  = cells & MCPD_KNOWN_COLUMNS
        if len(hits) >= MIN_HEADER_HITS:
            print(
                f"[INFO] Feuille '{sheet_name}' : en-tête détecté ligne {row_idx} "
                f"(hits : {sorted(hits)})",
                file=sys.stderr,
            )
            return int(row_idx)

    print(
        f"[WARN] Feuille '{sheet_name}' : aucun mot-clé MCPD trouvé dans les "
        f"{max_scan} premières lignes. Ligne 0 utilisée par défaut.",
        file=sys.stderr,
    )
    return 0


# ==============================================================================
# SECTION 5 — TRANSFORMATION Genotype -> ACCENUMB
# ==============================================================================

def genotype_to_accenumb(genotype: str) -> str:
    """
    Mappe la valeur brute de la colonne Genotype vers ACCENUMB.

    Note Stephan v4.3 : ACCENUMB est du texte LIBRE.
    On conserve la valeur brute telle quelle — aucune transformation n'est appliquée.
    Le format 'Ae. Biuncialis - MVGB - 378' est une valeur parfaitement valide.
    """
    return str(genotype).strip()


# ==============================================================================
# SECTION 6 — CONSTRUCTION DU SCHÉMA FRICTIONLESS
# ==============================================================================

def _build_schema(df: pd.DataFrame) -> Schema:
    """
    Construit le Schema frictionless pour les champs MCPD.
    Les champs obligatoires n'ont PAS constraints={"required": True} :
    la double-garde _check_mandatory_fields() est plus fiable.
    """
    fields: list[Field] = []

    # -- A) Champs OBLIGATOIRES (texte libre, sans required frictionless) ------
    fields.append(fl_fields.StringField(
        name="ACCENUMB",
        title="Accession Number",
        description=(
            "Identifiant libre assigné par l'institution. TEXTE LIBRE. "
            "Exemples valides : 'KU 11483a', 'PI 113869', 'Ae. Biuncialis - MVGB - 378'."
        ),
    ))
    fields.append(fl_fields.StringField(
        name="INSTCODE",
        title="Institute Code (FAO WIEWS)",
        description="Code FAO WIEWS obligatoire (ex. 'NOR039'). OBLIGATOIRE.",
    ))
    fields.append(fl_fields.StringField(
        name="GENUS",
        title="Genus",
        description="Genre botanique, initiale majuscule (ex. 'Triticum'). OBLIGATOIRE.",
    ))

    # -- B) Champs CATEGORICAL -------------------------------------------------
    for cat_field, cat_map in CATEGORICAL_FIELDS.items():
        fields.append(fl_fields.IntegerField(
            name=cat_field,
            title=f"{cat_field} (CATEGORICAL)",
            description=f"Type CATEGORICAL Germinate. Codes valides : {sorted(cat_map.keys())}.",
            constraints={"enum": sorted(cat_map.keys())},
        ))

    # -- C) ORIGCTY -----------------------------------------------------------
    fields.append(fl_fields.StringField(
        name="ORIGCTY",
        title="Country of Origin",
        description="Code ISO 3166-1 alpha-3 (optionnel). Si présent, doit être valide.",
        constraints={"enum": sorted(VALID_ISO3_CODES)},
    ))

    # -- D) Coordonnées décimales ---------------------------------------------
    fields.append(fl_fields.NumberField(
        name="DECLATITUDE",
        constraints={"minimum": -90, "maximum": 90},
    ))
    fields.append(fl_fields.NumberField(
        name="DECLONGITUDE",
        constraints={"minimum": -180, "maximum": 180},
    ))

    # -- E) Coordonnées DMS ---------------------------------------------------
    fields.append(fl_fields.StringField(name="LATITUDE"))
    fields.append(fl_fields.StringField(name="LONGITUDE"))

    # -- F) Autres champs MCPD optionnels -------------------------------------
    fields.append(fl_fields.StringField(name="SPECIES"))
    fields.append(fl_fields.StringField(name="SUBTAXA"))
    for col in ("REMARKS", "ACCENAME", "COLLSITE", "COLLNUMB", "ACQDATE",
                "COLLDATE", "BREDCODE", "DONORCODE", "DONORNUMB", "COLLCODE", "PUID"):
        fields.append(fl_fields.StringField(name=col))

    # -- G) Toutes les colonnes restantes (passthrough) -----------------------
    known_cols = {f.name for f in fields}
    for col in df.columns:
        if col not in known_cols:
            fields.append(fl_fields.StringField(name=col))

    return Schema(fields=fields)


# ==============================================================================
# SECTION 7 — UTILITAIRES
# ==============================================================================

def _is_set(val) -> bool:
    return val is not None and str(val).strip() not in ("", "nan", "None", "NaT")


# ==============================================================================
# SECTION 8 — CHARGEMENT EXCEL ET PRÉPARATION DU CSV
# ==============================================================================

def _load_and_prepare(excel_path: str) -> tuple[pd.DataFrame, str]:
    xl          = pd.ExcelFile(excel_path)
    sheet_names = xl.sheet_names

    if not sheet_names:
        raise ValueError("Le fichier Excel ne contient aucune feuille.")

    sheets: dict[str, pd.DataFrame] = {}
    for name in sheet_names:
        header_row = _detect_header_row(xl, name)
        try:
            df_tmp = xl.parse(name, header=header_row)
            df_tmp.dropna(how="all", inplace=True)
            df_tmp.dropna(axis=1, how="all", inplace=True)
            df_tmp.columns = df_tmp.columns.str.strip()
            df_tmp = df_tmp[~(
                df_tmp.astype(str)
                      .apply(lambda r: r.str.strip().eq("").all(), axis=1)
            )]
            if not df_tmp.empty:
                sheets[name] = df_tmp
                print(
                    f"[INFO] Feuille '{name}' : {len(df_tmp)} lignes, "
                    f"{len(df_tmp.columns)} colonnes. "
                    f"En-têtes : {list(df_tmp.columns[:8])}",
                    file=sys.stderr,
                )
            else:
                print(f"[WARN] Feuille '{name}' vide après nettoyage.", file=sys.stderr)
        except Exception as e:
            print(f"[WARN] Feuille '{name}' ignorée : {e}", file=sys.stderr)

    if not sheets:
        raise ValueError("Aucune feuille lisible dans le fichier.")

    # Détection clé de jointure
    def _find_join_key(sheets: dict[str, pd.DataFrame]) -> str | None:
        if len(sheets) == 1:
            return None
        all_cols = [set(df.columns) for df in sheets.values()]
        common   = set.intersection(*all_cols)
        if not common:
            return None
        for c in common:
            if c.strip().lower() == "ucp code":
                return c
        for c in sorted(common):
            if any(k in c.lower() for k in ("id", "code", "num", "key", "ref", "accen")):
                return c
        return sorted(common)[0]

    sheet_list = list(sheets.items())
    base_name, df = sheet_list[0]
    df = df.copy()

    if len(sheet_list) > 1:
        join_key = _find_join_key(sheets)
        if join_key is None:
            raise ValueError("Aucune colonne commune trouvée pour la jointure entre les feuilles.")
        df = df.drop_duplicates(subset=[join_key])
        for sname, df_other in sheet_list[1:]:
            df_other = df_other.drop_duplicates(subset=[join_key])
            overlap  = (set(df.columns) & set(df_other.columns)) - {join_key}
            if overlap:
                df       = df.rename(columns={c: c + f"_{base_name[:3]}"  for c in overlap})
                df_other = df_other.rename(columns={c: c + f"_{sname[:3]}" for c in overlap})
            df = df.merge(df_other, on=join_key, how="left")

    print(f"[INFO] Colonnes fusionnées : {list(df.columns)}", file=sys.stderr)

    # Recherche de la colonne Genotype (exacte, suffixée, ou par inclusion)
    genotype_col = None
    for col in df.columns:
        if col.strip().lower() == "genotype":
            genotype_col = col
            break
    if genotype_col is None:
        for col in df.columns:
            if col.strip().lower().split("_")[0] == "genotype":
                genotype_col = col
                break
    if genotype_col is None:
        for col in df.columns:
            if "genotype" in col.strip().lower():
                genotype_col = col
                break

    print(f"[INFO] Colonne Genotype retenue : {repr(genotype_col)}", file=sys.stderr)

    if genotype_col and "ACCENUMB" not in df.columns:
        df.insert(
            df.columns.get_loc(genotype_col) + 1,
            "ACCENUMB",
            df[genotype_col].apply(
                lambda g: genotype_to_accenumb(g) if _is_set(g) else ""
            ),
        )
        print(f"[INFO] ACCENUMB créée depuis '{genotype_col}'.", file=sys.stderr)
    elif "ACCENUMB" in df.columns:
        print("[INFO] Colonne ACCENUMB déjà présente.", file=sys.stderr)
    else:
        print(
            f"[ERREUR] Colonne 'Genotype' introuvable. "
            f"Colonnes disponibles : {list(df.columns)}",
            file=sys.stderr,
        )

    print(f"[INFO] Lignes après fusion : {len(df)}", file=sys.stderr)

    fd, csv_path = tempfile.mkstemp(suffix=".csv", dir=".", prefix="mcpd_tmp_")
    os.close(fd)
    df.to_csv(csv_path, index=False, encoding="utf-8")

    return df, csv_path


# ==============================================================================
# SECTION 9 — VÉRIFICATIONS POST-FRICTIONLESS
# ==============================================================================

def _check_mandatory_fields(df: pd.DataFrame) -> list[dict]:
    """
    Double-garde sur INSTCODE, ACCENUMB et GENUS.
    Niveau 1 : MISSING_COLUMN  — colonne absente du DataFrame (1 issue globale).
    Niveau 2 : MANDATORY_TRIAD — cellule vide dans une colonne présente (1 issue/ligne).
    """
    issues: list[dict] = []

    # Niveau 1 : colonnes structurellement absentes
    for field in MANDATORY_FIELDS:
        if field in df.columns:
            continue
        if field == "ACCENUMB":
            msg = (
                "Colonne 'ACCENUMB' introuvable. Elle est générée automatiquement "
                "depuis la colonne 'Genotype'. Vérifiez que 'Genotype' existe dans le fichier."
            )
        else:
            msg = (
                f"Colonne '{field}' ABSENTE du fichier Excel. "
                "OBLIGATOIRE pour l'ingestion Germinate."
            )
        issues.append({
            "row": 0, "accenumb": "[STRUCTURE]",
            "category": "MISSING_COLUMN", "field": field,
            "level": "ERROR", "message": msg,
        })

    # Niveau 2 : cellules vides dans colonnes présentes
    present_mandatory = [f for f in MANDATORY_FIELDS if f in df.columns]
    for idx, row in df.iterrows():
        row_num   = int(idx) + 2
        acc_val   = str(row.get("ACCENUMB", "")).strip()
        acc_label = acc_val if acc_val else f"[ligne {row_num}]"
        for field in present_mandatory:
            if not _is_set(row.get(field)):
                display_acc = f"[ligne {row_num}]" if field == "ACCENUMB" else acc_label
                issues.append({
                    "row": row_num, "accenumb": display_acc,
                    "category": "MANDATORY_TRIAD", "field": field,
                    "level": "ERROR",
                    "message": _friendly_message(field, "required-error", ""),
                })
    return issues


def _check_accenumb_uniqueness(df: pd.DataFrame) -> list[dict]:
    """
    ── NOTE v4.3 ──────────────────────────────────────────────────────────────
    ACCENUMB est du texte LIBRE. Il n'y a AUCUNE contrainte de format.
    La seule vérification valide est l'UNICITÉ DE LA COMBINAISON :
        INSTCODE + GENUS + ACCENUMB

    Deux accessions peuvent avoir le même ACCENUMB si elles appartiennent à
    des institutions ou des genres différents — c'est la réalité des banques de gènes.
    ──────────────────────────────────────────────────────────────────────────
    """
    issues: list[dict] = []

    # On ne peut vérifier l'unicité composite que si les trois colonnes existent
    cols_present = [c for c in ("INSTCODE", "GENUS", "ACCENUMB") if c in df.columns]
    if "ACCENUMB" not in df.columns:
        return issues  # déjà signalé comme MISSING_COLUMN

    if len(cols_present) < 3:
        # Colonnes partielles : unicité sur ACCENUMB seul comme fallback
        missing_triad = [c for c in ("INSTCODE", "GENUS") if c not in df.columns]
        counts = df["ACCENUMB"].value_counts()
        dupes  = counts[counts > 1].index.tolist()
        for dup in dupes:
            rows = df[df["ACCENUMB"] == dup].index.tolist()
            row_nums = [r + 2 for r in rows]
            issues.append({
                "row": row_nums[1], "accenumb": str(dup),
                "category": "ACCENUMB_UNIQUENESS",
                "field":    "ACCENUMB",
                "level":    "WARNING",
                "message": (
                    f"ACCENUMB '{dup}' apparaît {len(row_nums)} fois (lignes {row_nums}). "
                    f"Note : l'unicité devrait être vérifiée sur la combinaison "
                    f"INSTCODE + GENUS + ACCENUMB, mais {', '.join(missing_triad)} "
                    f"{'est absente' if len(missing_triad)==1 else 'sont absentes'} "
                    "de ce dataset. Vérification sur ACCENUMB seul en fallback."
                ),
            })
        return issues

    # Vérification sur la combinaison composite INSTCODE + GENUS + ACCENUMB
    df["_composite_key"] = (
        df["INSTCODE"].astype(str).str.strip() + " | " +
        df["GENUS"].astype(str).str.strip()    + " | " +
        df["ACCENUMB"].astype(str).str.strip()
    )
    counts = df["_composite_key"].value_counts()
    dupes  = counts[counts > 1].index.tolist()
    df.drop(columns=["_composite_key"], inplace=True)

    for dup_key in dupes:
        mask     = (
            df["INSTCODE"].astype(str).str.strip() + " | " +
            df["GENUS"].astype(str).str.strip()    + " | " +
            df["ACCENUMB"].astype(str).str.strip()
        ) == dup_key
        rows     = df[mask].index.tolist()
        row_nums = [r + 2 for r in rows]
        inst, genus, acc = dup_key.split(" | ")
        issues.append({
            "row": row_nums[1], "accenumb": acc,
            "category": "ACCENUMB_UNIQUENESS",
            "field":    "INSTCODE + GENUS + ACCENUMB",
            "level":    "ERROR",
            "message": (
                f"Combinaison dupliquée détectée aux lignes {row_nums} :\n"
                f"  INSTCODE = '{inst}'\n"
                f"  GENUS    = '{genus}'\n"
                f"  ACCENUMB = '{acc}'\n"
                "La combinaison (INSTCODE + GENUS + ACCENUMB) doit être unique "
                "au sein d'un même fichier. Deux accessions différentes ne peuvent "
                "pas avoir les trois valeurs identiques simultanément."
            ),
        })
    return issues


def _check_coordinate_exclusivity(df: pd.DataFrame) -> list[dict]:
    """LATITUDE XOR DECLATITUDE ; LONGITUDE XOR DECLONGITUDE (décision Alexandre)."""
    issues = []
    for idx, row in df.iterrows():
        acc = str(row.get("ACCENUMB", "")).strip() or f"ligne {int(idx) + 2}"
        rn  = int(idx) + 2
        if _is_set(row.get("LATITUDE")) and _is_set(row.get("DECLATITUDE")):
            issues.append({
                "row": rn, "accenumb": acc,
                "category": "COORD_EXCLUSIVITY",
                "field":    "LATITUDE / DECLATITUDE",
                "level":    "ERROR",
                "message": (
                    "Conflit latitude : LATITUDE (DMS) et DECLATITUDE (décimal) "
                    "sont tous deux renseignés. Germinate rejette les deux simultanément. "
                    "Conserver DECLATITUDE (décimal) de préférence."
                ),
            })
        if _is_set(row.get("LONGITUDE")) and _is_set(row.get("DECLONGITUDE")):
            issues.append({
                "row": rn, "accenumb": acc,
                "category": "COORD_EXCLUSIVITY",
                "field":    "LONGITUDE / DECLONGITUDE",
                "level":    "ERROR",
                "message": (
                    "Conflit longitude : LONGITUDE (DMS) et DECLONGITUDE (décimal) "
                    "sont tous deux renseignés. Germinate rejette les deux simultanément. "
                    "Conserver DECLONGITUDE (décimal) de préférence."
                ),
            })
    return issues


def _check_dms_format(df: pd.DataFrame) -> list[dict]:
    """Format et plages DMS pour LATITUDE et LONGITUDE."""
    issues = []
    for idx, row in df.iterrows():
        acc = str(row.get("ACCENUMB", "")).strip() or f"ligne {int(idx) + 2}"
        rn  = int(idx) + 2

        lat_dms = str(row.get("LATITUDE", "")).strip()
        if _is_set(lat_dms):
            if not LATITUDE_DMS_RE.match(lat_dms):
                issues.append({
                    "row": rn, "accenumb": acc,
                    "category": "LATITUDE_FORMAT", "field": "LATITUDE", "level": "ERROR",
                    "message": (
                        f"Format DMS invalide : '{lat_dms}'. "
                        "Attendu : DDMMSS + N ou S (7 chars). Ex : '103020S'."
                    ),
                })
            else:
                deg, mnt, sec = int(lat_dms[0:2]), int(lat_dms[2:4]), int(lat_dms[4:6])
                if deg > 90 or mnt > 59 or sec > 59:
                    issues.append({
                        "row": rn, "accenumb": acc,
                        "category": "LATITUDE_RANGE", "field": "LATITUDE", "level": "ERROR",
                        "message": (
                            f"DMS hors plage : degrés={deg}(max 90), "
                            f"minutes={mnt}(max 59), secondes={sec}(max 59)."
                        ),
                    })

        lon_dms = str(row.get("LONGITUDE", "")).strip()
        if _is_set(lon_dms):
            if not LONGITUDE_DMS_RE.match(lon_dms):
                issues.append({
                    "row": rn, "accenumb": acc,
                    "category": "LONGITUDE_FORMAT", "field": "LONGITUDE", "level": "ERROR",
                    "message": (
                        f"Format DMS invalide : '{lon_dms}'. "
                        "Attendu : DDDMMSS + E ou W (8 chars). Ex : '0762510W'."
                    ),
                })
            else:
                deg, mnt, sec = int(lon_dms[0:3]), int(lon_dms[3:5]), int(lon_dms[5:7])
                if deg > 180 or mnt > 59 or sec > 59:
                    issues.append({
                        "row": rn, "accenumb": acc,
                        "category": "LONGITUDE_RANGE", "field": "LONGITUDE", "level": "ERROR",
                        "message": (
                            f"DMS hors plage : degrés={deg}(max 180), "
                            f"minutes={mnt}(max 59), secondes={sec}(max 59)."
                        ),
                    })
    return issues


def _check_genus_format(df: pd.DataFrame) -> list[dict]:
    """GENUS doit commencer par une lettre majuscule (standard MCPD)."""
    issues = []
    for idx, row in df.iterrows():
        genus = row.get("GENUS")
        if _is_set(genus):
            g = str(genus).strip()
            if g and not g[0].isupper():
                acc = str(row.get("ACCENUMB", "")).strip() or f"ligne {int(idx) + 2}"
                issues.append({
                    "row": int(idx) + 2, "accenumb": acc,
                    "category": "GENUS_FORMAT", "field": "GENUS", "level": "ERROR",
                    "message": (
                        f"GENUS '{g}' : initiale minuscule. "
                        f"Correction : '{g[0].upper() + g[1:]}'."
                    ),
                })
    return issues


def _check_species_format(df: pd.DataFrame) -> list[dict]:
    """SPECIES doit être en minuscules ou valoir 'sp.'."""
    issues = []
    for idx, row in df.iterrows():
        species = row.get("SPECIES")
        if _is_set(species):
            s = str(species).strip()
            if s and not SPECIES_RE.match(s):
                acc = str(row.get("ACCENUMB", "")).strip() or f"ligne {int(idx) + 2}"
                issues.append({
                    "row": int(idx) + 2, "accenumb": acc,
                    "category": "SPECIES_FORMAT", "field": "SPECIES", "level": "WARNING",
                    "message": (
                        f"SPECIES '{s}' doit être en minuscules. "
                        f"Correction : '{s.lower()}'."
                    ),
                })
    return issues


def _check_subtaxa_format(df: pd.DataFrame) -> list[dict]:
    """SUBTAXA doit commencer par un préfixe MCPD autorisé."""
    issues = []
    for idx, row in df.iterrows():
        subtaxa = row.get("SUBTAXA")
        if _is_set(subtaxa):
            s = str(subtaxa).strip()
            if s and not any(s.startswith(p) for p in VALID_SUBTAXA_PREFIXES):
                acc = str(row.get("ACCENUMB", "")).strip() or f"ligne {int(idx) + 2}"
                issues.append({
                    "row": int(idx) + 2, "accenumb": acc,
                    "category": "SUBTAXA_FORMAT", "field": "SUBTAXA", "level": "WARNING",
                    "message": (
                        f"SUBTAXA '{s}' : préfixe non reconnu. "
                        f"Autorisés : {', '.join(VALID_SUBTAXA_PREFIXES)}."
                    ),
                })
    return issues


# ==============================================================================
# SECTION 10 — CONSTRUCTION DU RAPPORT
# ==============================================================================

def _accenumb_for_row(df: pd.DataFrame, row_number: int) -> str:
    idx = row_number - 2
    if 0 <= idx < len(df):
        v = str(df.iloc[idx].get("ACCENUMB", "")).strip()
        return v if v else f"[ligne {row_number}]"
    return f"[ligne {row_number}]"


def _friendly_message(field_name: str, code: str, note: str) -> str:
    if "required" in code:
        hints = {
            "ACCENUMB": (
                "ACCENUMB (colonne Genotype) ABSENT ou VIDE. "
                "Ce champ est OBLIGATOIRE — l'accession ne peut pas être créée dans Germinate. "
                "ACCENUMB est du TEXTE LIBRE : toute valeur non-vide est acceptée."
            ),
            "INSTCODE": (
                "INSTCODE ABSENT ou VIDE. "
                "Code FAO WIEWS obligatoire (ex. 'NOR039', 'TUN001')."
            ),
            "GENUS": (
                "GENUS ABSENT ou VIDE. "
                "Genre botanique obligatoire, initiale MAJUSCULE (ex. 'Triticum', 'Aegilops')."
            ),
        }
        return hints.get(field_name, f"Champ '{field_name}' obligatoire absent. {note}")

    if "enum" in code or "constraint" in code:
        if field_name in CATEGORICAL_FIELDS:
            lines = [f"Valeur non reconnue pour '{field_name}' (CATEGORICAL Germinate)."]
            lines.append("Valeurs autorisées (code -> description) :")
            for k, v in sorted(CATEGORICAL_FIELDS[field_name].items()):
                lines.append(f"  {str(k).rjust(4)}  ->  {v}")
            return "\n              ".join(lines)
        if field_name == "ORIGCTY":
            return (
                "Code pays inconnu. ISO 3166-1 alpha-3 requis (3 lettres MAJ). "
                "Codes historiques acceptés : XKX, SCG, YUG, CSK, DDR, SUN."
            )
        if field_name == "DECLATITUDE":
            return f"DECLATITUDE hors plage [-90, +90]. Valeur : {note}"
        if field_name == "DECLONGITUDE":
            return f"DECLONGITUDE hors plage [-180, +180]. Valeur : {note}"
        return f"Valeur hors contrainte pour '{field_name}'. Détail : {note}"

    if "type" in code:
        if field_name in CATEGORICAL_FIELDS:
            return (
                f"Type invalide pour '{field_name}' : entier attendu, reçu : {note}. "
                f"Codes valides : {sorted(CATEGORICAL_FIELDS[field_name].keys())}"
            )
        return f"Type invalide pour '{field_name}' : {note}"

    return note or f"Erreur '{code}' sur '{field_name}'."


def _format_report(
    excel_path: str,
    df: pd.DataFrame,
    frictionless_report,
    extra_issues: list[dict],
) -> str:
    W = 72

    def rule(c="="): return c * W
    def box(text):
        pad = W - 4
        return [
            "+" + "-" * (W - 2) + "+",
            "|  " + text.ljust(pad) + "|",
            "+" + "-" * (W - 2) + "+",
        ]

    fl_issues: list[dict] = []

    if frictionless_report:
        for task in (frictionless_report.tasks or []):
            for err in (task.errors or []):
                raw_row = getattr(err, "row_number", None)
                code    = getattr(err, "code", "") or ""
                fn      = str(getattr(err, "field_name", "") or getattr(err, "field_number", "") or "")
                note    = getattr(err, "note", "") or str(err)

                if raw_row is None:
                    continue  # erreur structurelle frictionless — couverte par double-garde Python
                if "required" in code:
                    continue  # doublon avec _check_mandatory_fields

                row_num = raw_row + 1
                acc     = _accenumb_for_row(df, row_num)

                if "type" in code:
                    category = "TYPE_" + fn.upper() if fn else "TYPE_ERROR"
                    level    = "ERROR"
                elif "constraint" in code or "enum" in code:
                    if fn == "ORIGCTY":
                        category = "ORIGCTY_INVALID"
                    elif fn in CATEGORICAL_FIELDS:
                        category = "CATEGORICAL_VALUE"
                    elif fn in ("DECLATITUDE", "DECLONGITUDE"):
                        category = "COORDINATE_RANGE"
                    else:
                        category = "CONSTRAINT_" + fn.upper() if fn else "CONSTRAINT_ERROR"
                    level = "ERROR"
                else:
                    category = code.upper().replace("-", "_") or "UNKNOWN"
                    level    = "ERROR"

                fl_issues.append({
                    "row": row_num, "accenumb": acc,
                    "category": category, "field": fn, "level": level,
                    "message": _friendly_message(fn, code, note),
                })

    all_issues = fl_issues + extra_issues
    errors     = [i for i in all_issues if i["level"] == "ERROR"]
    warnings   = [i for i in all_issues if i["level"] == "WARNING"]

    errors_by_cat:   dict[str, list[dict]] = defaultdict(list)
    warnings_by_cat: dict[str, list[dict]] = defaultdict(list)
    for i in errors:   errors_by_cat[i["category"]].append(i)
    for i in warnings: warnings_by_cat[i["category"]].append(i)

    n_err  = len(errors)
    n_warn = len(warnings)
    status = "INVALIDE" if errors else "VALIDE"

    lines: list[str] = []

    lines += [
        rule("="),
        "  RAPPORT DE VALIDATION MCPD — PIPELINE GERMINATE  "
        f"(frictionless v{frictionless.__version__})",
        rule("="),
        f"  Fichier              : {excel_path}",
        f"  Lignes totales       : {len(df)}",
        f"  Statut               : {status}",
        f"  Erreurs bloquantes   : {n_err}",
        f"  Avertissements       : {n_warn}",
        rule("="),
    ]

    # Rappel règle ACCENUMB — mise en évidence dans le rapport
    lines += [
        "",
        rule("-"),
        "  RÈGLE ACCENUMB — TEXTE LIBRE (v4.3)",
        rule("-"),
        "  L'ACCENUMB est un identifiant LIBRE assigné par l'institution.",
        "  Il n'y a AUCUNE contrainte de format sur sa valeur.",
        "  Exemples valides : 'KU 11483a', 'KU 11635b', 'PI 113869',",
        "                     'Ae. Biuncialis - MVGB - 378'.",
        "  La vérification porte UNIQUEMENT sur :",
        "    1. Sa PRÉSENCE (non vide)  — obligatoire",
        "    2. L'UNICITÉ de la combinaison INSTCODE + GENUS + ACCENUMB",
        rule("="),
    ]

    # Référentiel CATEGORICAL
    lines += ["", rule("-"), "  RÉFÉRENTIEL — CHAMPS CATEGORICAL GERMINATE", rule("-")]
    for fname, fmap in CATEGORICAL_FIELDS.items():
        lines.append(f"  {fname} :")
        for k, v in sorted(fmap.items()):
            lines.append(f"      {str(k).rjust(4)}  ->  {v}")
        lines.append("")
    lines.append(rule("="))

    # Règles principales
    lines += [
        "  CHAMPS OBLIGATOIRES (Stephan + Alexandre)",
        rule("-"),
        "  * INSTCODE  : Code FAO WIEWS (ex. 'NOR039') — OBLIGATOIRE",
        "  * ACCENUMB  : Texte libre depuis colonne 'Genotype' — OBLIGATOIRE",
        "  * GENUS     : Genre botanique, initiale MAJUSCULE — OBLIGATOIRE",
        "",
        "  TRIADE QUALITÉ GERMINATE (mid-term requirement — Alexandre)",
        rule("-"),
        "  INSTCODE + ACCENUMB + GENUS = minimum recommandé pour des passeports exploitables.",
        "",
        "  COORDONNÉES — Exclusivité par axe (Alexandre)",
        rule("-"),
        "  LATITUDE (DMS) XOR DECLATITUDE (décimal)  — jamais les deux",
        "  LONGITUDE (DMS) XOR DECLONGITUDE (décimal) — jamais les deux",
        rule("="),
    ]

    # ERREURS BLOQUANTES
    if errors:
        lines += [""] + box("ERREURS BLOQUANTES — Ingestion Germinate IMPOSSIBLE") + [""]
        for cat, issues in sorted(errors_by_cat.items()):
            n_aff    = len({i["row"] for i in issues} - {0})
            n_struct = len([i for i in issues if i["row"] == 0])
            suffix   = ""
            if n_struct > 0 and n_aff == 0:
                suffix = "  [ERREUR STRUCTURELLE — colonne absente]"
            elif n_struct > 0:
                suffix = f"  + {n_struct} erreur(s) structurelle(s)"
            lines.append(f"  [{cat}]  {n_aff} accession(s) affectée(s){suffix}")
            lines.append(rule("-"))

            if cat == "MISSING_COLUMN":
                for iss in issues:
                    lines.append(f"  !! COLONNE MANQUANTE : '{iss['field']}'")
                    for ml in iss["message"].split("\n"):
                        lines.append("              " + ml)
                lines.append("")

            elif cat == "MANDATORY_TRIAD":
                seen:     dict[int, set] = defaultdict(set)
                row_repr: dict[int, str] = {}
                for iss in issues:
                    seen[iss["row"]].add(iss["field"])
                    row_repr[iss["row"]] = iss["accenumb"]
                by_combo: dict[str, list] = defaultdict(list)
                for rn, fset in sorted(seen.items()):
                    combo = ", ".join(sorted(fset))
                    by_combo[combo].append(row_repr[rn])
                for combo, accs in sorted(by_combo.items(), key=lambda x: -len(x[1])):
                    lines.append(
                        f"  Champ(s) MANQUANT(S) : {combo}  ->  {len(accs)} accession(s)"
                    )
                    ex_msg = next(
                        (i["message"] for i in issues if i["field"] in combo), ""
                    )
                    for ml in ex_msg.split("\n"):
                        lines.append("    !  " + ml)
                    for i in range(0, len(accs), 6):
                        lines.append("      " + "  |  ".join(accs[i:i + 6]))
                lines.append("")

            elif cat == "ACCENUMB_UNIQUENESS":
                for iss in issues:
                    lines.append(
                        f"  Ligne {str(iss['row']).rjust(4)}  |  "
                        f"ACCENUMB: {iss['accenumb'].ljust(30)}  |  "
                        f"Champ: {iss['field']}"
                    )
                    for ml in iss["message"].split("\n"):
                        lines.append("              " + ml)
                lines.append("")

            else:
                for iss in issues:
                    lines.append(
                        f"  Ligne {str(iss['row']).rjust(4)}  |  "
                        f"ACCENUMB: {iss['accenumb'].ljust(25)}  |  "
                        f"Champ: {iss['field']}"
                    )
                    for ml in iss["message"].split("\n"):
                        lines.append("              " + ml)
                lines.append("")
    else:
        lines += ["", "  OK  Aucune erreur bloquante détectée.", ""]

    # AVERTISSEMENTS
    if warnings:
        lines += [rule("=")] + box("AVERTISSEMENTS QUALITÉ — Enrichissement conseillé") + [""]
        lines += [
            "  Ces points ne bloquent pas l'ingestion mais dégradent",
            "  la qualité des passeports. Correction conseillée avant import.",
            "",
        ]
        for cat, issues in sorted(warnings_by_cat.items()):
            n_aff = len({i["row"] for i in issues})
            lines.append(f"  [{cat}]  {n_aff} accession(s) affectée(s)")
            lines.append(rule("-"))
            for iss in issues:
                lines.append(
                    f"  Ligne {str(iss['row']).rjust(4)}  |  "
                    f"ACCENUMB: {iss['accenumb'].ljust(25)}  |  "
                    f"Champ: {iss['field']}"
                )
                for ml in iss["message"].split("\n"):
                    lines.append("              " + ml)
            lines.append("")
    else:
        lines += ["", "  OK  Aucun avertissement qualité.", ""]

    # RÉSUMÉ
    missing_col_issues = errors_by_cat.get("MISSING_COLUMN", [])
    mandatory_issues   = errors_by_cat.get("MANDATORY_TRIAD", [])
    n_missing_cols     = len(missing_col_issues)
    n_mandatory_rows   = len({i["row"] for i in mandatory_issues})

    lines += [
        rule("="),
        "  RÉSUMÉ",
        rule("-"),
        f"  Lignes analysées     : {len(df)}",
        f"  Erreurs bloquantes   : {n_err}",
        f"  Avertissements       : {n_warn}",
    ]

    if n_missing_cols > 0 or n_mandatory_rows > 0:
        lines.append("")
        lines.append("  Détail champs obligatoires :")
        if n_missing_cols > 0:
            cols_absent = [i["field"] for i in missing_col_issues]
            lines.append(
                f"    . COLONNES ABSENTES    : {n_missing_cols} "
                f"({', '.join(cols_absent)})"
            )
        if n_mandatory_rows > 0:
            by_field: dict[str, int] = {}
            for i in mandatory_issues:
                by_field[i["field"]] = by_field.get(i["field"], 0) + 1
            for fname, count in sorted(by_field.items()):
                lines.append(f"    . CELLULES VIDES [{fname}] : {count} ligne(s)")

    other_error_cats = {
        cat: cat_issues
        for cat, cat_issues in sorted(errors_by_cat.items())
        if cat not in ("MISSING_COLUMN", "MANDATORY_TRIAD")
    }
    if other_error_cats or warnings:
        lines.append("")
        if other_error_cats:
            lines.append("  Autres catégories d'erreurs :")
            for cat, cat_issues in other_error_cats.items():
                lines.append(f"    . {cat}: {len(cat_issues)} occurrence(s)")
        if warnings:
            lines.append("  Avertissements :")
            for cat, cat_issues in sorted(warnings_by_cat.items()):
                lines.append(f"    ! {cat}: {len(cat_issues)} occurrence(s)")
    lines.append(rule("="))

    return "\n".join(lines)


# ==============================================================================
# SECTION 11 — POINT D'ENTRÉE PUBLIC
# ==============================================================================

def run_validation(excel_path: str) -> str:
    try:
        df, csv_path = _load_and_prepare(excel_path)
    except Exception as e:
        return f"[FATAL] Impossible de charger le fichier : {e}"

    schema = _build_schema(df)
    report = None
    try:
        resource = Resource(path=csv_path, schema=schema)
        report   = validate(resource)
    except Exception as e:
        print(f"[WARN] frictionless validate() : {e}", file=sys.stderr)
    finally:
        try:
            os.unlink(csv_path)
        except OSError:
            pass

    extra: list[dict] = []
    extra.extend(_check_mandatory_fields(df))
    extra.extend(_check_accenumb_uniqueness(df))      # COMPOSITE uniqueness
    extra.extend(_check_coordinate_exclusivity(df))
    extra.extend(_check_dms_format(df))
    extra.extend(_check_genus_format(df))
    extra.extend(_check_species_format(df))
    extra.extend(_check_subtaxa_format(df))

    return _format_report(excel_path, df, report, extra)


# ==============================================================================
# SECTION 12 — EXÉCUTION DIRECTE
# ==============================================================================

if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "Wheat_Minerals sheet_UCP_COUSIN.xlsx"
    print(run_validation(path))