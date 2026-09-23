"""
MCPD Validator Pipeline v4.3.1 — final fix for INSTCODE/GENUS duplication
==========================================================================
This module validates an Excel file (multiple sheets) against MCPD and Germinate.
It no longer renames INSTCODE or GENUS during merge.
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
        "[FATAL] frictionless not installed.\n"
        "  pip install frictionless 'frictionless[excel]'"
    )

print("*** USING CORRECTED VALIDATOR v4.3.1 ***", file=sys.stderr)

# ==============================================================================
# SECTION 1 — CATEGORICAL VALUES (Germinate CATEGORICAL)
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
# SECTION 2 — ISO 3166-1 ALPHA-3 COUNTRY CODES
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
    "XKX","SCG","YUG","CSK","DDR","SUN",
}

# ==============================================================================
# SECTION 3 — PATTERNS AND CONSTANTS
# ==============================================================================

LATITUDE_DMS_RE  = re.compile(r"^\d{2}\d{2}\d{2}[NS]$")
LONGITUDE_DMS_RE = re.compile(r"^\d{3}\d{2}\d{2}[EW]$")
SPECIES_RE = re.compile(r"^([a-z]+(\s[a-z]+)*|sp\.)$")
VALID_SUBTAXA_PREFIXES = ("subsp.", "convar.", "var.", "f.", "Group")
MANDATORY_FIELDS = ("INSTCODE", "ACCENUMB", "GENUS")

MCPD_KNOWN_COLUMNS: set[str] = {
    "INSTCODE", "ACCENUMB", "ACCENAME", "GENUS", "SPECIES", "SUBTAXA",
    "LATITUDE", "LONGITUDE", "DECLATITUDE", "DECLONGITUDE",
    "COORDUNCERT", "GEOREFMETH", "COLLDATE", "COLLSITE", "COLLNUMB",
    "COLLCODE", "COLLNAME", "ORIGCTY", "ORIGCTYNAME", "SAMPSTAT",
    "MLSSTAT", "STORAGE", "COLLSRC", "ACQDATE", "DONORCODE", "DONORNUMB",
    "DONORNAME", "BREDCODE", "BREDNAME", "REMARKS", "PUID", "ACCEURL",
    "GENOTYPE", "GENOTYPE NAME", "UCP CODE", "UCP_CODE",
    "ACCESSION", "ACCESSION NUMBER", "ACCESSION NAME",
}
MIN_HEADER_HITS = 2

# ==============================================================================
# SECTION 4 — HEADER DETECTION
# ==============================================================================

def _detect_header_row(xl: pd.ExcelFile, sheet_name: str, max_scan: int = 20) -> int:
    try:
        df_raw = xl.parse(sheet_name, header=None, nrows=max_scan, dtype=str)
    except Exception as e:
        print(f"[WARN] Header scan '{sheet_name}' : {e}", file=sys.stderr)
        return 0
    for row_idx, row in df_raw.iterrows():
        cells = {str(v).strip().upper() for v in row.dropna()}
        hits  = cells & MCPD_KNOWN_COLUMNS
        if len(hits) >= MIN_HEADER_HITS:
            print(f"[INFO] Sheet '{sheet_name}' : header at row {row_idx} hits {sorted(hits)}", file=sys.stderr)
            return int(row_idx)
    print(f"[WARN] Sheet '{sheet_name}' : no MCPD keywords, using row 0", file=sys.stderr)
    return 0

# ==============================================================================
# SECTION 5 — Genotype -> ACCENUMB (free text)
# ==============================================================================

def genotype_to_accenumb(genotype: str) -> str:
    return str(genotype).strip()

# ==============================================================================
# SECTION 6 — FRICTIONLESS SCHEMA
# ==============================================================================

def _build_schema(df: pd.DataFrame) -> Schema:
    fields: list[Field] = []
    fields.append(fl_fields.StringField(name="ACCENUMB", title="Accession Number"))
    fields.append(fl_fields.StringField(name="INSTCODE", title="Institute Code (FAO WIEWS)"))
    fields.append(fl_fields.StringField(name="GENUS", title="Genus"))
    for cat_field, cat_map in CATEGORICAL_FIELDS.items():
        fields.append(fl_fields.IntegerField(
            name=cat_field,
            constraints={"enum": sorted(cat_map.keys())},
        ))
    fields.append(fl_fields.StringField(name="ORIGCTY", constraints={"enum": sorted(VALID_ISO3_CODES)}))
    fields.append(fl_fields.NumberField(name="DECLATITUDE", constraints={"minimum": -90, "maximum": 90}))
    fields.append(fl_fields.NumberField(name="DECLONGITUDE", constraints={"minimum": -180, "maximum": 180}))
    fields.append(fl_fields.StringField(name="LATITUDE"))
    fields.append(fl_fields.StringField(name="LONGITUDE"))
    fields.append(fl_fields.StringField(name="SPECIES"))
    fields.append(fl_fields.StringField(name="SUBTAXA"))
    for col in ("REMARKS", "ACCENAME", "COLLSITE", "COLLNUMB", "ACQDATE",
                "COLLDATE", "BREDCODE", "DONORCODE", "DONORNUMB", "COLLCODE", "PUID"):
        fields.append(fl_fields.StringField(name=col))
    known_cols = {f.name for f in fields}
    for col in df.columns:
        if col not in known_cols:
            fields.append(fl_fields.StringField(name=col))
    return Schema(fields=fields)

# ==============================================================================
# SECTION 7 — UTILITIES
# ==============================================================================

def _is_set(val) -> bool:
    return val is not None and str(val).strip() not in ("", "nan", "None", "NaT")

# ==============================================================================
# SECTION 8 — EXCEL LOADING AND MERGING (FIXED)
# ==============================================================================

def _load_and_prepare(excel_path: str) -> tuple[pd.DataFrame, str]:
    xl = pd.ExcelFile(excel_path)
    sheet_names = xl.sheet_names
    if not sheet_names:
        raise ValueError("No sheets found.")

    sheets: dict[str, pd.DataFrame] = {}
    for name in sheet_names:
        header_row = _detect_header_row(xl, name)
        try:
            df_tmp = xl.parse(name, header=header_row)
            df_tmp.dropna(how="all", inplace=True)
            df_tmp.dropna(axis=1, how="all", inplace=True)
            df_tmp.columns = df_tmp.columns.str.strip()
            df_tmp = df_tmp[~(df_tmp.astype(str).apply(lambda r: r.str.strip().eq("").all(), axis=1))]
            if not df_tmp.empty:
                sheets[name] = df_tmp
                print(f"[INFO] Sheet '{name}' : {len(df_tmp)} rows, {len(df_tmp.columns)} cols. Headers: {list(df_tmp.columns[:8])}", file=sys.stderr)
            else:
                print(f"[WARN] Sheet '{name}' empty after cleaning.", file=sys.stderr)
        except Exception as e:
            print(f"[WARN] Sheet '{name}' skipped: {e}", file=sys.stderr)

    if not sheets:
        raise ValueError("No readable sheets.")

    def _find_join_key(sheets: dict[str, pd.DataFrame]) -> str | None:
        if len(sheets) == 1:
            return None
        all_cols = [set(df.columns) for df in sheets.values()]
        common = set.intersection(*all_cols)
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
            raise ValueError("No common join column.")
        df = df.drop_duplicates(subset=[join_key])
        for sname, df_other in sheet_list[1:]:
            df_other = df_other.drop_duplicates(subset=[join_key])
            overlap = (set(df.columns) & set(df_other.columns)) - {join_key}
            # Columns that must NOT be duplicated: we drop them from the other sheet
            keep_cols = {"INSTCODE", "GENUS", "ACCENUMB"}
            # Columns to rename (all overlapping except the ones we keep)
            rename_cols = overlap - keep_cols
            if rename_cols:
                df_other = df_other.rename(columns={c: f"{c}_{sname[:3]}" for c in rename_cols})
            # Drop the kept columns from the other sheet to avoid _x/_y
            drop_cols = overlap & keep_cols
            if drop_cols:
                df_other = df_other.drop(columns=drop_cols)
            df = df.merge(df_other, on=join_key, how="left")

    print(f"[INFO] Merged columns : {list(df.columns)}", file=sys.stderr)

    # Locate Genotype column
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

    print(f"[INFO] Genotype column selected : {repr(genotype_col)}", file=sys.stderr)

    if genotype_col and "ACCENUMB" not in df.columns:
        df.insert(df.columns.get_loc(genotype_col) + 1, "ACCENUMB",
                  df[genotype_col].apply(lambda g: genotype_to_accenumb(g) if _is_set(g) else ""))
        print(f"[INFO] ACCENUMB created from '{genotype_col}'.", file=sys.stderr)
    elif "ACCENUMB" in df.columns:
        print("[INFO] ACCENUMB already present.", file=sys.stderr)
    else:
        print(f"[ERROR] 'Genotype' column not found. Columns: {list(df.columns)}", file=sys.stderr)

    print(f"[INFO] Rows after merge : {len(df)}", file=sys.stderr)

    fd, csv_path = tempfile.mkstemp(suffix=".csv", dir=".", prefix="mcpd_tmp_")
    os.close(fd)
    df.to_csv(csv_path, index=False, encoding="utf-8")
    return df, csv_path

# ==============================================================================
# SECTION 9 — POST‑FRICTIONLESS CHECKS
# ==============================================================================

def _check_mandatory_fields(df: pd.DataFrame) -> list[dict]:
    """
    Double-check for mandatory columns (INSTCODE, ACCENUMB, GENUS).
    Level 1: MISSING_COLUMN – column completely absent.
    Level 2: MANDATORY_TRIAD – empty cells in present columns.
    """
    issues: list[dict] = []

    # Level 1: structurally absent columns
    for field in MANDATORY_FIELDS:
        if field in df.columns:
            continue
        if field == "ACCENUMB":
            msg = (
                "Column 'ACCENUMB' not found. It is automatically created "
                "from the 'Genotype' column. Please ensure 'Genotype' exists."
            )
        else:
            msg = (
                f"Column '{field}' MISSING from the Excel file. "
                "REQUIRED for Germinate ingestion."
            )
        issues.append({
            "row": 0, "accenumb": "[STRUCTURE]",
            "category": "MISSING_COLUMN", "field": field,
            "level": "ERROR", "message": msg,
        })

    # Level 2: empty cells in present mandatory columns
    present_mandatory = [f for f in MANDATORY_FIELDS if f in df.columns]
    for idx, row in df.iterrows():
        row_num   = int(idx) + 2
        acc_val   = str(row.get("ACCENUMB", "")).strip()
        acc_label = acc_val if acc_val else f"[row {row_num}]"
        for field in present_mandatory:
            if not _is_set(row.get(field)):
                display_acc = f"[row {row_num}]" if field == "ACCENUMB" else acc_label
                issues.append({
                    "row": row_num, "accenumb": display_acc,
                    "category": "MANDATORY_TRIAD", "field": field,
                    "level": "ERROR",
                    "message": _friendly_message(field, "required-error", ""),
                })
    return issues


def _check_accenumb_uniqueness(df: pd.DataFrame) -> list[dict]:
    """
    NOTE v4.3: ACCENUMB is FREE TEXT. No format constraints.
    The only valid uniqueness check is on the combination:
        INSTCODE + GENUS + ACCENUMB

    Two accessions may share the same ACCENUMB if they belong to different
    institutes or genera – this is realistic in gene banks.
    """
    issues: list[dict] = []

    cols_present = [c for c in ("INSTCODE", "GENUS", "ACCENUMB") if c in df.columns]
    if "ACCENUMB" not in df.columns:
        return issues  # already reported as MISSING_COLUMN

    if len(cols_present) < 3:
        # Fallback: uniqueness on ACCENUMB alone
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
                    f"ACCENUMB '{dup}' appears {len(row_nums)} times (rows {row_nums}). "
                    f"Uniqueness should be checked on INSTCODE+GENUS+ACCENUMB, but "
                    f"{', '.join(missing_triad)} {'is' if len(missing_triad)==1 else 'are'} "
                    f"missing from this dataset. Falling back to ACCENUMB alone."
                ),
            })
        return issues

    # Full composite key check
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
                f"Duplicate combination detected at rows {row_nums}:\n"
                f"  INSTCODE = '{inst}'\n"
                f"  GENUS    = '{genus}'\n"
                f"  ACCENUMB = '{acc}'\n"
                "The combination (INSTCODE + GENUS + ACCENUMB) must be unique "
                "within the same file."
            ),
        })
    return issues


def _check_coordinate_exclusivity(df: pd.DataFrame) -> list[dict]:
    """LATITUDE XOR DECLATITUDE ; LONGITUDE XOR DECLONGITUDE (Germinate rule)."""
    issues = []
    for idx, row in df.iterrows():
        acc = str(row.get("ACCENUMB", "")).strip() or f"row {int(idx) + 2}"
        rn  = int(idx) + 2
        if _is_set(row.get("LATITUDE")) and _is_set(row.get("DECLATITUDE")):
            issues.append({
                "row": rn, "accenumb": acc,
                "category": "COORD_EXCLUSIVITY",
                "field":    "LATITUDE / DECLATITUDE",
                "level":    "ERROR",
                "message": (
                    "Latitude conflict: LATITUDE (DMS) and DECLATITUDE (decimal) "
                    "are both provided. Germinate rejects both simultaneously. "
                    "Prefer DECLATITUDE (decimal)."
                ),
            })
        if _is_set(row.get("LONGITUDE")) and _is_set(row.get("DECLONGITUDE")):
            issues.append({
                "row": rn, "accenumb": acc,
                "category": "COORD_EXCLUSIVITY",
                "field":    "LONGITUDE / DECLONGITUDE",
                "level":    "ERROR",
                "message": (
                    "Longitude conflict: LONGITUDE (DMS) and DECLONGITUDE (decimal) "
                    "are both provided. Germinate rejects both simultaneously. "
                    "Prefer DECLONGITUDE (decimal)."
                ),
            })
    return issues


def _check_dms_format(df: pd.DataFrame) -> list[dict]:
    """Validate DMS format and ranges for LATITUDE and LONGITUDE."""
    issues = []
    for idx, row in df.iterrows():
        acc = str(row.get("ACCENUMB", "")).strip() or f"row {int(idx) + 2}"
        rn  = int(idx) + 2

        lat_dms = str(row.get("LATITUDE", "")).strip()
        if _is_set(lat_dms):
            if not LATITUDE_DMS_RE.match(lat_dms):
                issues.append({
                    "row": rn, "accenumb": acc,
                    "category": "LATITUDE_FORMAT", "field": "LATITUDE", "level": "ERROR",
                    "message": (
                        f"Invalid DMS format: '{lat_dms}'. "
                        "Expected: DDMMSS + N or S (7 chars). Example: '103020S'."
                    ),
                })
            else:
                deg, mnt, sec = int(lat_dms[0:2]), int(lat_dms[2:4]), int(lat_dms[4:6])
                if deg > 90 or mnt > 59 or sec > 59:
                    issues.append({
                        "row": rn, "accenumb": acc,
                        "category": "LATITUDE_RANGE", "field": "LATITUDE", "level": "ERROR",
                        "message": (
                            f"DMS out of range: degrees={deg} (max 90), "
                            f"minutes={mnt} (max 59), seconds={sec} (max 59)."
                        ),
                    })

        lon_dms = str(row.get("LONGITUDE", "")).strip()
        if _is_set(lon_dms):
            if not LONGITUDE_DMS_RE.match(lon_dms):
                issues.append({
                    "row": rn, "accenumb": acc,
                    "category": "LONGITUDE_FORMAT", "field": "LONGITUDE", "level": "ERROR",
                    "message": (
                        f"Invalid DMS format: '{lon_dms}'. "
                        "Expected: DDDMMSS + E or W (8 chars). Example: '0762510W'."
                    ),
                })
            else:
                deg, mnt, sec = int(lon_dms[0:3]), int(lon_dms[3:5]), int(lon_dms[5:7])
                if deg > 180 or mnt > 59 or sec > 59:
                    issues.append({
                        "row": rn, "accenumb": acc,
                        "category": "LONGITUDE_RANGE", "field": "LONGITUDE", "level": "ERROR",
                        "message": (
                            f"DMS out of range: degrees={deg} (max 180), "
                            f"minutes={mnt} (max 59), seconds={sec} (max 59)."
                        ),
                    })
    return issues


def _check_genus_format(df: pd.DataFrame) -> list[dict]:
    """GENUS must start with an uppercase letter (MCPD standard)."""
    issues = []
    for idx, row in df.iterrows():
        genus = row.get("GENUS")
        if _is_set(genus):
            g = str(genus).strip()
            if g and not g[0].isupper():
                acc = str(row.get("ACCENUMB", "")).strip() or f"row {int(idx) + 2}"
                issues.append({
                    "row": int(idx) + 2, "accenumb": acc,
                    "category": "GENUS_FORMAT", "field": "GENUS", "level": "ERROR",
                    "message": (
                        f"GENUS '{g}' : lowercase initial letter. "
                        f"Suggestion: '{g[0].upper() + g[1:]}'."
                    ),
                })
    return issues


def _check_species_format(df: pd.DataFrame) -> list[dict]:
    """SPECIES must be all lowercase or equal to 'sp.'."""
    issues = []
    for idx, row in df.iterrows():
        species = row.get("SPECIES")
        if _is_set(species):
            s = str(species).strip()
            if s and not SPECIES_RE.match(s):
                acc = str(row.get("ACCENUMB", "")).strip() or f"row {int(idx) + 2}"
                issues.append({
                    "row": int(idx) + 2, "accenumb": acc,
                    "category": "SPECIES_FORMAT", "field": "SPECIES", "level": "WARNING",
                    "message": (
                        f"SPECIES '{s}' must be in lowercase. "
                        f"Suggestion: '{s.lower()}'."
                    ),
                })
    return issues


def _check_subtaxa_format(df: pd.DataFrame) -> list[dict]:
    """SUBTAXA must start with an allowed MCPD prefix."""
    issues = []
    for idx, row in df.iterrows():
        subtaxa = row.get("SUBTAXA")
        if _is_set(subtaxa):
            s = str(subtaxa).strip()
            if s and not any(s.startswith(p) for p in VALID_SUBTAXA_PREFIXES):
                acc = str(row.get("ACCENUMB", "")).strip() or f"row {int(idx) + 2}"
                issues.append({
                    "row": int(idx) + 2, "accenumb": acc,
                    "category": "SUBTAXA_FORMAT", "field": "SUBTAXA", "level": "WARNING",
                    "message": (
                        f"SUBTAXA '{s}' : prefix not recognized. "
                        f"Allowed: {', '.join(VALID_SUBTAXA_PREFIXES)}."
                    ),
                })
    return issues


# ==============================================================================
# SECTION 10 — REPORT GENERATION
# ==============================================================================

def _accenumb_for_row(df: pd.DataFrame, row_number: int) -> str:
    idx = row_number - 2
    if 0 <= idx < len(df):
        v = str(df.iloc[idx].get("ACCENUMB", "")).strip()
        return v if v else f"[row {row_number}]"
    return f"[row {row_number}]"


def _friendly_message(field_name: str, code: str, note: str) -> str:
    if "required" in code:
        hints = {
            "ACCENUMB": (
                "ACCENUMB (Genotype column) MISSING or EMPTY. "
                "This field is REQUIRED – the accession cannot be created in Germinate. "
                "ACCENUMB is FREE TEXT: any non‑empty value is accepted."
            ),
            "INSTCODE": (
                "INSTCODE MISSING or EMPTY. "
                "Mandatory FAO WIEWS code (e.g. 'NOR039', 'TUN001')."
            ),
            "GENUS": (
                "GENUS MISSING or EMPTY. "
                "Mandatory botanical genus, uppercase initial letter (e.g. 'Triticum', 'Aegilops')."
            ),
        }
        return hints.get(field_name, f"Field '{field_name}' is required. {note}")

    if "enum" in code or "constraint" in code:
        if field_name in CATEGORICAL_FIELDS:
            lines = [f"Invalid value for '{field_name}' (CATEGORICAL Germinate)."]
            lines.append("Permitted values (code -> description):")
            for k, v in sorted(CATEGORICAL_FIELDS[field_name].items()):
                lines.append(f"  {str(k).rjust(4)}  ->  {v}")
            return "\n              ".join(lines)
        if field_name == "ORIGCTY":
            return (
                "Unknown country code. ISO 3166-1 alpha-3 required (3 uppercase letters). "
                "Historical codes accepted: XKX, SCG, YUG, CSK, DDR, SUN."
            )
        if field_name == "DECLATITUDE":
            return f"DECLATITUDE out of range [-90, +90]. Value: {note}"
        if field_name == "DECLONGITUDE":
            return f"DECLONGITUDE out of range [-180, +180]. Value: {note}"
        return f"Invalid value for '{field_name}'. Detail: {note}"

    if "type" in code:
        if field_name in CATEGORICAL_FIELDS:
            return (
                f"Invalid type for '{field_name}' : integer expected, received: {note}. "
                f"Valid codes: {sorted(CATEGORICAL_FIELDS[field_name].keys())}"
            )
        return f"Invalid type for '{field_name}' : {note}"

    return note or f"Error '{code}' on '{field_name}'."


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
                    continue  # structural frictionless error – already covered by Python double-check
                if "required" in code:
                    continue  # duplicate of _check_mandatory_fields

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
    status = "INVALID" if errors else "VALID"

    lines: list[str] = []

    lines += [
        rule("="),
        "  MCPD VALIDATION REPORT — GERMINATE PIPELINE  "
        f"(frictionless v{frictionless.__version__})",
        rule("="),
        f"  File                 : {excel_path}",
        f"  Total rows           : {len(df)}",
        f"  Status               : {status}",
        f"  Blocking errors      : {n_err}",
        f"  Warnings             : {n_warn}",
        rule("="),
    ]

    # ACCENUMB rule reminder
    lines += [
        "",
        rule("-"),
        "  ACCENUMB VALIDATION RULE — FREE TEXT FIELD ",
        rule("-"),
        "  ACCENUMB is a free identifier assigned by the institution.",
        "  There is no formatting constraint on its value.",
        "  Valid examples: 'KU 11483a', 'KU 11635b', 'PI 113869',",
        "                  'Ae. Biuncialis - MVGB - 378'.",
        "  Validation is based ONLY on:",
        "    1. Presence (must not be empty)",
        "    2. Uniqueness of the combination INSTCODE + GENUS + ACCENUMB",
        rule("="),
    ]

    # Categorical reference
    lines += ["", rule("-"), "  REFERENCE — CATEGORICAL FIELDS (GERMINATE)", rule("-")]
    for fname, fmap in CATEGORICAL_FIELDS.items():
        lines.append(f"  {fname} :")
        for k, v in sorted(fmap.items()):
            lines.append(f"      {str(k).rjust(4)}  ->  {v}")
        lines.append("")
    lines.append(rule("="))

    # Main rules
    lines += [
        "  MANDATORY FIELDS ",
        rule("-"),
        "  * INSTCODE  : FAO WIEWS code (e.g. 'NOR039') — REQUIRED",
        "  * ACCENUMB  : Free text from 'Genotype' column — REQUIRED",
        "  * GENUS     : Botanical genus, uppercase initial — REQUIRED",
        "",
        "  QUALITY TRIAD FOR GERMINATE ",
        rule("-"),
        "  INSTCODE + ACCENUMB + GENUS = minimum recommended for usable passports.",
        "",
        "  COORDINATES ",
        rule("-"),
        "  LATITUDE (DMS) XOR DECLATITUDE (decimal)  — never both",
        "  LONGITUDE (DMS) XOR DECLONGITUDE (decimal) — never both",
        rule("="),
    ]

    # BLOCKING ERRORS
    if errors:
        lines += [""] + box("BLOCKING ERRORS — Germinate ingestion IMPOSSIBLE") + [""]
        for cat, issues in sorted(errors_by_cat.items()):
            n_aff    = len({i["row"] for i in issues} - {0})
            n_struct = len([i for i in issues if i["row"] == 0])
            suffix   = ""
            if n_struct > 0 and n_aff == 0:
                suffix = "  [SCHEMA ERROR — column missing]"
            elif n_struct > 0:
                suffix = f"  + {n_struct} STRUCTURAL ERROR(S)"
            lines.append(f"  [{cat}]  {n_aff} Modified accession(s){suffix}")
            lines.append(rule("-"))

            if cat == "MISSING_COLUMN":
                for iss in issues:
                    lines.append(f"  !! REQUIRED COLUMN NOT FOUND : '{iss['field']}'")
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
                        f"  Required field(s) missing :{combo}  ->  {len(accs)} affected accession(s)"
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
                        f"  Row {str(iss['row']).rjust(4)}  |  "
                        f"ACCENUMB: {iss['accenumb'].ljust(30)}  |  "
                        f"Field: {iss['field']}"
                    )
                    for ml in iss["message"].split("\n"):
                        lines.append("              " + ml)
                lines.append("")

            else:
                for iss in issues:
                    lines.append(
                        f"  Row {str(iss['row']).rjust(4)}  |  "
                        f"ACCENUMB: {iss['accenumb'].ljust(25)}  |  "
                        f"Field: {iss['field']}"
                    )
                    for ml in iss["message"].split("\n"):
                        lines.append("              " + ml)
                lines.append("")
    else:
        lines += ["", "  OK   No critical errors detected.", ""]

    # WARNINGS
    if warnings:
        lines += [rule("=")] + box("WARNINGS — Quality enrichment recommended") + [""]
        lines += [
            "  These points do not block ingestion but degrade",
            "  the quality of the passports. Recommended correction before import.",
            "",
        ]
        for cat, issues in sorted(warnings_by_cat.items()):
            n_aff = len({i["row"] for i in issues})
            lines.append(f"  [{cat}]  {n_aff} accession(s) affected")
            lines.append(rule("-"))
            for iss in issues:
                lines.append(
                    f"  Row {str(iss['row']).rjust(4)}  |  "
                    f"ACCENUMB: {iss['accenumb'].ljust(25)}  |  "
                    f"Field: {iss['field']}"
                )
                for ml in iss["message"].split("\n"):
                    lines.append("              " + ml)
            lines.append("")
    else:
        lines += ["", "  OK   No warnings detected.", ""]

    # SUMMARY
    missing_col_issues = errors_by_cat.get("MISSING_COLUMN", [])
    mandatory_issues   = errors_by_cat.get("MANDATORY_TRIAD", [])
    n_missing_cols     = len(missing_col_issues)
    n_mandatory_rows   = len({i["row"] for i in mandatory_issues})

    lines += [
        rule("="),
        "  SUMMARY",
        rule("-"),
        f"  Rows analyzed        : {len(df)}",
        f"  Blocking errors      : {n_err}",
        f"  Warnings             : {n_warn}",
    ]

    if n_missing_cols > 0 or n_mandatory_rows > 0:
        lines.append("")
        lines.append("  Detail of required fields :")
        if n_missing_cols > 0:
            cols_absent = [i["field"] for i in missing_col_issues]
            lines.append(
                f"    . REQUIRED COLUMNS MISSING    : {n_missing_cols} "
                f"({', '.join(cols_absent)})"
            )
        if n_mandatory_rows > 0:
            by_field: dict[str, int] = {}
            for i in mandatory_issues:
                by_field[i["field"]] = by_field.get(i["field"], 0) + 1
            for fname, count in sorted(by_field.items()):
                lines.append(f"    . MISSING VALUES [{fname}] : {count} row(s)")

    other_error_cats = {
        cat: cat_issues
        for cat, cat_issues in sorted(errors_by_cat.items())
        if cat not in ("MISSING_COLUMN", "MANDATORY_TRIAD")
    }
    if other_error_cats or warnings:
        lines.append("")
        if other_error_cats:
            lines.append("  Other error categories :")
            for cat, cat_issues in other_error_cats.items():
                lines.append(f"    . {cat}: {len(cat_issues)} occurrence(s) found")
        if warnings:
            lines.append("  Warnings :")
            for cat, cat_issues in sorted(warnings_by_cat.items()):
                lines.append(f"    ! {cat}: {len(cat_issues)} occurrence(s) found")
    lines.append(rule("="))

    return "\n".join(lines)


# ==============================================================================
# SECTION 11 — PUBLIC ENTRY POINT
# ==============================================================================
def run_validation(excel_path: str) -> str:
    try:
        df, csv_path = _load_and_prepare(excel_path)
    except Exception as e:
        return f"[FATAL] Failed to load file: {e}"
    schema = _build_schema(df)
    report = None
    try:
        resource = Resource(path=csv_path, schema=schema)
        report = validate(resource)
    except Exception as e:
        print(f"[WARN] frictionless validate() : {e}", file=sys.stderr)
    finally:
        try:
            os.unlink(csv_path)
        except OSError:
            pass
    extra = []
    extra.extend(_check_mandatory_fields(df))
    extra.extend(_check_accenumb_uniqueness(df))
    extra.extend(_check_coordinate_exclusivity(df))
    extra.extend(_check_dms_format(df))
    extra.extend(_check_genus_format(df))
    extra.extend(_check_species_format(df))
    extra.extend(_check_subtaxa_format(df))
    return _format_report(excel_path, df, report, extra)

if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "Wheat_Minerals sheet_UCP_COUSIN.xlsx"
    print(run_validation(path))