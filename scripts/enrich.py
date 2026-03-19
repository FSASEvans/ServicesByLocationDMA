"""
enrich.py
---------
Takes the raw Snowflake CSV export and adds columns:
  BRAND         — extracted from STORE_NAME_FULL via pattern matching
  DMA_CODE      — Nielsen DMA numeric code
  DMA_NAME      — Nielsen DMA market name
  STORE_MODEL   — Current store model tier (QL / QL+ / QL++ / FS / EM)
  PROPOSED_MODEL — 2026 proposed model tier (from operations assessment)

Usage:
  python scripts/enrich.py \
    --input data/raw_export.csv \
    --output data/store_services_enriched.csv \
    --dma-names scripts/dma_names.csv \
    --dma-boundaries scripts/dma_boundaries.csv \
    --store-models scripts/store_models.csv
"""

import csv
import sys
import argparse
from collections import defaultdict

# ── BRAND EXTRACTION ─────────────────────────────────────────────────────────
BRAND_MAP = [
    # Most specific first — order matters
    ("Grease Monkey (ADI)", lambda s: "GREASE MONKEY" in s and "ADI" in s),
    ("Grease Monkey",       lambda s: "GREASE MONKEY" in s),
    ("American Lubefast",   lambda s: "LUBEFAST" in s or "AMERICAN LUBE" in s),
    ("Uncle Ed's",          lambda s: "UNCLE ED" in s),
    ("Kwik Kar",            lambda s: "KWIK KAR" in s),
    ("SpeeDee",             lambda s: "SPEEDEE" in s),
    ("Herbert Automotive",  lambda s: "HERBERT" in s),
    ("Economy Oil Change",  lambda s: "ECONOMY OIL" in s),
]

def extract_brand(store_name_full):
    s = (store_name_full or "").upper()
    for brand, fn in BRAND_MAP:
        if fn(s):
            return brand
    return "Other"


# ── DMA POLYGON MATCHING ──────────────────────────────────────────────────────
def load_dma_names(path):
    names = {}
    with open(path) as f:
        for row in csv.DictReader(f):
            names[int(row["dma_code"])] = row["geo_dma"].strip('"')
    return names

def load_dma_polygons(path, dma_names):
    """Build shapely polygons from the boundary coordinate file."""
    try:
        from shapely.geometry import Polygon, MultiPolygon
    except ImportError:
        print("ERROR: shapely not installed. Run: pip install shapely")
        sys.exit(1)

    pieces = defaultdict(lambda: defaultdict(list))
    with open(path) as f:
        for row in csv.DictReader(f):
            code  = int(float(row["geo_dma"]))
            piece = row["piece"]
            hole  = row["hole"].strip('"') == "TRUE"
            if not hole:
                pieces[code][piece].append(
                    (float(row["long"]), float(row["lat"]))
                )

    polygons = {}
    for code, ps in pieces.items():
        polys = [Polygon(coords) for coords in ps.values() if len(coords) >= 3]
        if not polys:
            continue
        try:
            polygons[code] = MultiPolygon(polys) if len(polys) > 1 else polys[0]
        except Exception:
            polygons[code] = polys[0]

    print(f"  Loaded {len(polygons)} DMA polygons")
    return polygons

def assign_dma(lat, lon, polygons, dma_names):
    """Point-in-polygon lookup. Returns (dma_code, dma_name)."""
    try:
        from shapely.geometry import Point
        if not lat or not lon:
            return "", "Unknown"
        pt = Point(float(lon), float(lat))
        for code, poly in polygons.items():
            if poly.contains(pt):
                return str(code), dma_names.get(code, "Unknown")
    except Exception:
        pass
    return "", "Unknown"

# ── MANUAL OVERRIDES (coastal stores that fall outside polygon boundaries) ────
MANUAL_DMA = {
    # store_number: (dma_code, dma_name)
    "8057": ("600", "Corpus Christi"),            # Rockport, TX
    "8306": ("686", "Mobile-Pensacola (Ft Walt)"), # Gulf Breeze, FL
}


# ── STORE MODEL LOOKUP ────────────────────────────────────────────────────────
def load_store_models(path):
    """Load store_models.csv → {store_number: {STORE_MODEL, PROPOSED_MODEL}}"""
    models = {}
    if not path:
        return models
    try:
        with open(path) as f:
            for row in csv.DictReader(f):
                models[str(row["STORE_NUMBER"]).strip()] = {
                    "STORE_MODEL":    row.get("STORE_MODEL", "").strip(),
                    "PROPOSED_MODEL": row.get("PROPOSED_MODEL", "").strip(),
                }
        print(f"  Loaded {len(models)} store model tiers")
    except FileNotFoundError:
        print(f"  WARNING: store_models.csv not found at {path} — STORE_MODEL will be empty")
    return models


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input",          required=True)
    parser.add_argument("--output",         required=True)
    parser.add_argument("--dma-names",      required=True)
    parser.add_argument("--dma-boundaries", required=True)
    parser.add_argument("--store-models",   default="scripts/store_models.csv",
                        help="Path to store_models.csv (optional, adds STORE_MODEL column)")
    args = parser.parse_args()

    print(f"Loading DMA reference data...")
    dma_names    = load_dma_names(args.dma_names)
    dma_polygons = load_dma_polygons(args.dma_boundaries, dma_names)

    print(f"Loading store model tiers...")
    store_models = load_store_models(args.store_models)

    print(f"Reading input: {args.input}")
    with open(args.input, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    print(f"  {len(rows):,} rows")

    # Validate required columns
    required = {"STORE_NUMBER", "STORE_NAME", "LATITUDE", "LONGITUDE",
                "L1_CATEGORY", "L2_SUBCATEGORY", "SERVICE_NAME", "TRANSACTION_COUNT"}
    missing = required - set(rows[0].keys())
    if missing:
        print(f"ERROR: Missing columns in CSV: {missing}")
        print(f"Available columns: {list(rows[0].keys())}")
        sys.exit(1)

    # Cache enrichment per store so we don't re-run polygon lookup for every row
    store_cache = {}
    matched = 0
    total_stores = 0
    model_matched = 0

    enriched_rows = []
    for row in rows:
        sn = row["STORE_NUMBER"]

        if sn not in store_cache:
            total_stores += 1
            brand = extract_brand(
                row.get("STORE_NAME_FULL", row.get("STORE_NAME", ""))
            )

            # DMA assignment
            if sn in MANUAL_DMA:
                dma_code, dma_name = MANUAL_DMA[sn]
            else:
                dma_code, dma_name = assign_dma(
                    row.get("LATITUDE"), row.get("LONGITUDE"),
                    dma_polygons, dma_names
                )

            if dma_name != "Unknown":
                matched += 1

            # Store model tier
            model_info = store_models.get(sn, {})
            if model_info:
                model_matched += 1

            store_cache[sn] = {
                "BRAND":          brand,
                "DMA_CODE":       dma_code,
                "DMA_NAME":       dma_name,
                "STORE_MODEL":    model_info.get("STORE_MODEL", ""),
                "PROPOSED_MODEL": model_info.get("PROPOSED_MODEL", ""),
            }

        dims = store_cache[sn]
        out_row = {k: v for k, v in row.items() if k != "STORE_NAME_FULL"}
        out_row.update(dims)
        enriched_rows.append(out_row)

    print(f"  DMA matched: {matched}/{total_stores} stores")
    print(f"  Model tier matched: {model_matched}/{total_stores} stores")
    unmatched_dma = [sn for sn, d in store_cache.items() if d["DMA_NAME"] == "Unknown"]
    if unmatched_dma:
        print(f"  Unmatched DMA stores: {unmatched_dma}")

    # Write output
    fieldnames = list(enriched_rows[0].keys())
    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(enriched_rows)

    print(f"  Written: {args.output} ({len(enriched_rows):,} rows)")


if __name__ == "__main__":
    main()
