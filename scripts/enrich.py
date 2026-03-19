"""
enrich.py
---------
Takes the raw Snowflake CSV export and adds:
  BRAND         — extracted from STORE_NAME_FULL via pattern matching
  DMA_CODE      — Nielsen DMA numeric code
  DMA_NAME      — Nielsen DMA market name
  STORE_MODEL   — Current store model tier (QL / QL+ / QL++ / FS / EM)
  PROPOSED_MODEL — 2026 proposed model tier
  REGION_NUM    — FSA ops region number (1–4)
  REGION_VP     — Regional VP name
  DISTRICT_NUM  — FSA ops district/area number
  DISTRICT_MGR  — District/Area Manager name

Usage:
  python scripts/enrich.py \
    --input data/raw_export.csv \
    --output data/store_services_enriched.csv \
    --dma-names scripts/dma_names.csv \
    --dma-boundaries scripts/dma_boundaries.csv \
    --store-models scripts/store_models.csv \
    --region-district scripts/region_district.csv
"""

import csv, sys, argparse
from collections import defaultdict

BRAND_MAP = [
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

def load_dma_names(path):
    names = {}
    with open(path) as f:
        for row in csv.DictReader(f):
            names[int(row["dma_code"])] = row["geo_dma"].strip('"')
    return names

def load_dma_polygons(path, dma_names):
    try:
        from shapely.geometry import Polygon, MultiPolygon
    except ImportError:
        print("ERROR: shapely not installed."); sys.exit(1)
    pieces = defaultdict(lambda: defaultdict(list))
    with open(path) as f:
        for row in csv.DictReader(f):
            code = int(float(row["geo_dma"])); piece = row["piece"]
            hole = row["hole"].strip('"') == "TRUE"
            if not hole:
                pieces[code][piece].append((float(row["long"]), float(row["lat"])))
    polygons = {}
    for code, ps in pieces.items():
        polys = [Polygon(coords) for coords in ps.values() if len(coords) >= 3]
        if not polys: continue
        try: polygons[code] = MultiPolygon(polys) if len(polys) > 1 else polys[0]
        except: polygons[code] = polys[0]
    print(f"  Loaded {len(polygons)} DMA polygons")
    return polygons

def assign_dma(lat, lon, polygons, dma_names):
    try:
        from shapely.geometry import Point
        if not lat or not lon: return "", "Unknown"
        pt = Point(float(lon), float(lat))
        for code, poly in polygons.items():
            if poly.contains(pt):
                return str(code), dma_names.get(code, "Unknown")
    except: pass
    return "", "Unknown"

MANUAL_DMA = {
    "8057": ("600", "Corpus Christi"),
    "8306": ("686", "Mobile-Pensacola (Ft Walt)"),
}

def load_csv_lookup(path, key_col, value_cols, label=""):
    lookup = {}
    if not path: return lookup
    try:
        with open(path) as f:
            for row in csv.DictReader(f):
                k = str(row[key_col]).strip()
                lookup[k] = {c: row.get(c, '').strip() for c in value_cols}
        print(f"  Loaded {len(lookup)} {label} records")
    except FileNotFoundError:
        print(f"  WARNING: {path} not found — {label} columns will be empty")
    return lookup

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input",           required=True)
    parser.add_argument("--output",          required=True)
    parser.add_argument("--dma-names",       required=True)
    parser.add_argument("--dma-boundaries",  required=True)
    parser.add_argument("--store-models",    default="scripts/store_models.csv")
    parser.add_argument("--region-district", default="scripts/region_district.csv")
    args = parser.parse_args()

    print("Loading DMA reference data...")
    dma_names    = load_dma_names(args.dma_names)
    dma_polygons = load_dma_polygons(args.dma_boundaries, dma_names)

    print("Loading store model tiers...")
    store_models = load_csv_lookup(args.store_models, "STORE_NUMBER",
                                   ["STORE_MODEL","PROPOSED_MODEL"], "store models")

    print("Loading region/district assignments...")
    region_dist = load_csv_lookup(args.region_district, "STORE_NUMBER",
                                  ["REGION_NUM","REGION_VP","DISTRICT_NUM","DISTRICT_MGR"],
                                  "region/district")

    print(f"Reading input: {args.input}")
    with open(args.input, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    print(f"  {len(rows):,} rows")

    required = {"STORE_NUMBER","STORE_NAME","LATITUDE","LONGITUDE",
                "L1_CATEGORY","L2_SUBCATEGORY","SERVICE_NAME","TRANSACTION_COUNT"}
    missing = required - set(rows[0].keys())
    if missing:
        print(f"ERROR: Missing columns: {missing}"); sys.exit(1)

    store_cache = {}
    matched_dma = matched_model = matched_region = total_stores = 0

    enriched_rows = []
    for row in rows:
        sn = row["STORE_NUMBER"]
        if sn not in store_cache:
            total_stores += 1
            brand = extract_brand(row.get("STORE_NAME_FULL", row.get("STORE_NAME", "")))
            if sn in MANUAL_DMA:
                dma_code, dma_name = MANUAL_DMA[sn]
            else:
                dma_code, dma_name = assign_dma(row.get("LATITUDE"), row.get("LONGITUDE"),
                                                 dma_polygons, dma_names)
            if dma_name != "Unknown": matched_dma += 1
            model_info = store_models.get(sn, {})
            if model_info: matched_model += 1
            rd_info = region_dist.get(sn, {})
            if rd_info: matched_region += 1
            store_cache[sn] = {
                "BRAND":          brand,
                "DMA_CODE":       dma_code,
                "DMA_NAME":       dma_name,
                "STORE_MODEL":    model_info.get("STORE_MODEL", ""),
                "PROPOSED_MODEL": model_info.get("PROPOSED_MODEL", ""),
                "REGION_NUM":     rd_info.get("REGION_NUM", ""),
                "REGION_VP":      rd_info.get("REGION_VP", ""),
                "DISTRICT_NUM":   rd_info.get("DISTRICT_NUM", ""),
                "DISTRICT_MGR":   rd_info.get("DISTRICT_MGR", ""),
            }
        dims = store_cache[sn]
        out_row = {k: v for k, v in row.items() if k != "STORE_NAME_FULL"}
        out_row.update(dims)
        enriched_rows.append(out_row)

    print(f"  DMA matched:           {matched_dma}/{total_stores}")
    print(f"  Model tier matched:    {matched_model}/{total_stores}")
    print(f"  Region/dist matched:   {matched_region}/{total_stores}")
    unmatched = [sn for sn, d in store_cache.items() if d["DMA_NAME"] == "Unknown"]
    if unmatched: print(f"  Unmatched DMA stores:  {unmatched}")

    fieldnames = list(enriched_rows[0].keys())
    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(enriched_rows)
    print(f"  Written: {args.output} ({len(enriched_rows):,} rows)")

if __name__ == "__main__":
    main()
