"""
build_dashboard.py
------------------
Reads the enriched CSV (with BRAND + DMA_NAME columns),
builds the JSON data payload, and injects it into the
dashboard HTML template to produce index.html.

Usage:
  python scripts/build_dashboard.py \
    --input  data/store_services_enriched.csv \
    --output index.html
"""

import csv
import json
import argparse
import os
from collections import defaultdict
from datetime import date, timedelta


# ── CATEGORY + BRAND COLOR MAPS (keep in sync with template JS) ──────────────
CAT_COLORS = {
    "Air Conditioning":       "#0EA5E9",
    "Air Filters":            "#16A34A",
    "Battery":                "#D97706",
    "Brakes":                 "#DC2626",
    "Differentials":          "#7C3AED",
    "Emissions & Inspections":"#4F46E5",
    "Engine":                 "#2563EB",
    "Fluids & Cooling":       "#0891B2",
    "Fuel System":            "#B45309",
    "Lighting":               "#CA8A04",
    "Shop & Misc":            "#6B7280",
    "Suspension & Steering":  "#059669",
    "Tires":                  "#65A30D",
    "Transmission":           "#9333EA",
    "Wiper Blades":           "#475569",
}


def build_data_payload(csv_path):
    """Read enriched CSV and return the dashboard JSON data dict."""
    print(f"  Reading: {csv_path}")
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    print(f"  {len(rows):,} rows")

    # Validate
    required = {"STORE_NUMBER", "STORE_NAME", "CITY", "STATE",
                "LATITUDE", "LONGITUDE", "ADDRESS1", "ZIP_CODE",
                "L1_CATEGORY", "TRANSACTION_COUNT", "BRAND", "DMA_NAME"}
    missing = required - set(rows[0].keys())
    if missing:
        raise ValueError(f"Missing columns in enriched CSV: {missing}")

    stores = {}
    for r in rows:
        sn = r["STORE_NUMBER"]
        if sn not in stores:
            try:
                lat = float(r["LATITUDE"]) if r["LATITUDE"] else None
                lon = float(r["LONGITUDE"]) if r["LONGITUDE"] else None
            except ValueError:
                lat = lon = None

            stores[sn] = {
                "store_number": sn,
                "store_name":   r["STORE_NAME"],
                "brand":        r.get("BRAND", "Other"),
                "address":      r.get("ADDRESS1", ""),
                "city":         r["CITY"],
                "state":        r["STATE"],
                "zip":          r.get("ZIP_CODE", ""),
                "lat":          lat,
                "lon":          lon,
                "dma_code":     r.get("DMA_CODE", ""),
                "dma_name":     r.get("DMA_NAME", "Unknown"),
                "region_num":   r.get("REGION_NUM", ""),
                "region_vp":    r.get("REGION_VP", ""),
                "district_num": r.get("DISTRICT_NUM", ""),
                "district_mgr": r.get("DISTRICT_MGR", ""),
                "categories":     {},   # non-oil transactions only
                "categories_all": {},   # oil-inclusive
            }

        l1 = r["L1_CATEGORY"]
        if l1 and l1 != "Other":
            try:
                tx = int(r["TRANSACTION_COUNT"])
            except (ValueError, KeyError):
                tx = 0
            is_oil = str(r.get("IS_OIL_TRANSACTION", "0")).strip() == "1"
            stores[sn]["categories_all"][l1] = stores[sn]["categories_all"].get(l1, 0) + tx
            if not is_oil:
                stores[sn]["categories"][l1] = stores[sn]["categories"].get(l1, 0) + tx

    stores_list = sorted(stores.values(), key=lambda s: (
        int(s["store_number"]) if s["store_number"].isdigit() else 0
    ))

    # Derive filter dimension lists
    all_l1     = sorted(set(
        l1 for s in stores_list
        for l1 in s["categories"]
    ))
    all_brands   = sorted(set(s["brand"] for s in stores_list))
    all_states   = sorted(set(s["state"] for s in stores_list))
    all_dmas     = sorted(set(
        s["dma_name"] for s in stores_list
        if s["dma_name"] and s["dma_name"] != "Unknown"
    ))
    all_regions  = sorted(set(
        s["region_num"] for s in stores_list if s["region_num"]
    ), key=lambda x: int(x) if x.isdigit() else 999)
    all_districts = sorted(set(
        s["district_num"] for s in stores_list if s["district_num"]
    ), key=lambda x: int(x) if x.isdigit() else 999)

    print(f"  {len(stores_list)} stores · {len(all_l1)} categories · "
          f"{len(all_brands)} brands · {len(all_dmas)} DMAs · "
          f"{len(all_regions)} regions · {len(all_districts)} districts")

    return {
        "stores":         stores_list,
        "l1_categories":  all_l1,
        "brands":         all_brands,
        "states":         all_states,
        "dmas":           all_dmas,
        "regions":        all_regions,
        "districts":      all_districts,
    }


def compute_date_window():
    """Return human-readable date range string for 'last 3 months'."""
    today = date.today()
    # Approximate 3 months as 91 days for display purposes
    start = today - timedelta(days=91)
    months = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]
    return (f"{months[start.month-1]} {start.day}, {start.year} "
            f"&ndash; "
            f"{months[today.month-1]} {today.day}, {today.year}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input",    required=True,  help="Path to enriched CSV")
    parser.add_argument("--output",   required=True,  help="Output HTML path (index.html)")
    parser.add_argument("--template", default=None,   help="Dashboard template HTML (auto-detected)")
    args = parser.parse_args()

    # Auto-detect template path relative to this script
    if args.template is None:
        script_dir   = os.path.dirname(os.path.abspath(__file__))
        args.template = os.path.join(script_dir, "dashboard_template.html")

    if not os.path.exists(args.template):
        raise FileNotFoundError(f"Template not found: {args.template}")

    print("Building dashboard...")
    payload = build_data_payload(args.input)
    data_json = json.dumps(payload, separators=(",", ":"))

    date_window = compute_date_window()

    print(f"  Reading template: {args.template}")
    with open(args.template) as f:
        template = f.read()

    # Inject data
    html = template.replace("__DASHBOARD_DATA__", data_json)

    # Inject date window (template has a placeholder)
    html = html.replace("__DATE_WINDOW__", date_window)

    # Write output
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = len(html) // 1024
    print(f"  Written: {args.output} ({size_kb} KB)")
    print("Done.")


if __name__ == "__main__":
    main()
