# FullSpeed Automotive — Store Service Intelligence Dashboard

> **Marketing Performance & Insights** · Confidential Internal Use

---

## Overview

This project delivers a recurring, data-driven intelligence system that maps **which services each FullSpeed Automotive location actively offers**, based on real transaction behavior from the Cinch CRM platform. The output is an executive-facing web dashboard that enables filtering by brand, geography, and Nielsen DMA, with a live store map and service penetration KPIs across all 285 active locations.

The system was designed to answer a core operational question:

> *"Across our network of 285 stores and 8 brands, which service categories are truly being performed — and where?"*

---

## Project Goals

| # | Goal | Status |
|---|------|--------|
| 1 | Recurring Snowflake SQL query to pull service availability per location | ✅ Complete |
| 2 | Automated workflow to transform raw data into a structured, classified matrix | ✅ Complete |
| 3 | Executive HTML dashboard hosted via GitHub Pages with full filter capability | ✅ Complete |
| 4 | Automated GitHub Actions pipeline: CSV commit → rebuild → redeploy | ✅ Complete |

---

## Architecture

```
Cinch Snowflake Data Share
        │
        ▼
store_services_matrix_v_FILTERS.sql
(TRANSACTIONS + TRANSACTIONS_DETAILS + LOCATIONS)
        │  27,678 rows · store × service grain · last 3 months
        ▼
/data/ folder in GitHub repo
        │  (manual step: export CSV from Snowflake, upload to /data)
        ▼
GitHub Actions: rebuild-dashboard.yml
        ├── scripts/enrich.py
        │   ├── Brand extraction (pattern match on STORE_NAME_FULL)
        │   └── DMA assignment (point-in-polygon · Shapely · Nielsen polygons)
        └── scripts/build_dashboard.py
            └── Injects enriched JSON into dashboard_template.html
                        │
                        ▼
                    index.html
            (self-contained · data embedded · GitHub Pages)
                        │
                        ▼
        https://fsasevans.github.io/ServicesByLocationDMA/
```

---

## Data Source

**Platform:** Snowflake (read-only data share)
**Schema:** `CINCH_DATA.CINCH_TO_FULLSPEED_AUTOMOTIVE_DATA_SHARE`

| Table | Rows | Role |
|---|---|---|
| `TRANSACTIONS` | 13.6M | Header — date filter via `TRANSACTION_AT` |
| `TRANSACTIONS_DETAILS` | 88.9M | Line items — `ITEM_DESCR`, `ITEM_TYPE` |
| `LOCATIONS` | 327 | Store reference — coordinates, address, name |

**Key column notes:**
- Date filter uses `TRANSACTIONS.TRANSACTION_AT` — the actual service date. **Not** `CREATED_AT`, which is the Cinch system ingestion timestamp and does not reflect when the service occurred.
- Store coordinates come from `LOCATIONS.LATITUDE` / `LOCATIONS.LONGITUDE` — used for DMA point-in-polygon matching.
- `TRANSACTIONS_DETAILS.ITEM_TYPE` is a **free-text POS field**, not a controlled enum. Classification relies entirely on `ITEM_DESCR` pattern matching.

---

## Methodology

### Oil Change Exclusion

Entire transactions are excluded if **any line item** matches known oil change patterns. This is a transaction-level exclusion — if a transaction has an oil change line plus a wiper blade line, the entire transaction is removed.

**Excluded `ITEM_TYPE` values:**
`OIL`, `OL1`, `OL6`, `OF`, `SYNTHETIC`, `OIL FILTERS`, `ENGINE OIL FS04`, `ENGINE OIL FS01`

**Excluded `ITEM_DESCR` patterns:**
`%OIL CHANGE%`, `%OIL & FILTER%`, `%OIL AND FILTER%`, `%OIL/FILTER%`

**Rationale:** 83.6% of all transactions are oil-related. Excluding them isolates true "beyond-oil" service capability per location.

### Coverage Validation (last 3-month period)

| Metric | Value |
|---|---|
| Total transactions (last 3 months) | 526,841 |
| Oil-excluded transactions | 440,337 (83.6%) |
| Non-oil transactions with detail rows | 83,542 |
| Transactions represented in service matrix | 140,481 |
| Transactions classified into hierarchy | 132,021 (94.0%) |
| Unclassified ("Other") transactions | 8,460 (6.0%) |

The unclassified 6% is intentional — long-tail single-store free-text POS entries (e.g. `O-RING`, `HOSE`, technician notes) that do not represent meaningful service categories.

### Brand Extraction

Derived from `LOCATIONS.FRIENDLY_NAME` via Python pattern matching in `scripts/enrich.py`. Match order matters — most specific patterns checked first.

| Brand | Stores | Match Pattern |
|---|---|---|
| Grease Monkey (ADI) | 19 | Contains "GREASE MONKEY" AND "ADI" |
| Grease Monkey | 156 | Contains "GREASE MONKEY" |
| American Lubefast | 38 | Contains "LUBEFAST" |
| Uncle Ed's | 29 | Contains "UNCLE ED" |
| Kwik Kar | 25 | Contains "KWIK KAR" |
| SpeeDee | 9 | Contains "SPEEDEE" |
| Herbert Automotive | 6 | Contains "HERBERT" |
| Economy Oil Change | 3 | Contains "ECONOMY OIL" |

### DMA Assignment

**Method:** Point-in-polygon using Python `shapely` against official Nielsen DMA polygon boundaries.
**Source:** 206 of 210 Nielsen DMA polygons (public domain boundary data).
**Input:** `LOCATIONS.LATITUDE` / `LOCATIONS.LONGITUDE`.
**Coverage:** 283/285 stores matched automatically. 2 coastal stores assigned manually (polygon edge artifacts):
- Store 8057 (Rockport, TX) → Corpus Christi DMA (600)
- Store 8306 (Gulf Breeze, FL) → Mobile-Pensacola DMA (686)

**Result:** 50 unique DMAs across 285 stores in 24 states.

---

## Product Service Hierarchy

Service names come from `TRANSACTIONS_DETAILS.ITEM_DESCR` — a free-text POS field containing a mix of controlled SKU codes (`A5642 AIR FILTER`), standardized names (`TIRE ROTATION`), technician notes, and part numbers (`H11-55 HALOGEN BULB`).

Classification is implemented as SQL `CASE/WHEN` in `store_services_matrix_v_FILTERS.sql` using `UPPER(ITEM_DESCR) LIKE`, `RLIKE` (regex), and exact string matching. **First matching rule wins.**

**Classification rate:** 94.0% of transactions (132,021 / 140,481)
**Structure:** 15 L1 Categories → 80 L2 Subcategories

---

### L1 Category Summary

*Note: L1 categories updated March 2026 following cross-validation against operations data (see Validation section). "Tires" split into "Tire Sales" and "Tire Services"; "Alignment" reclassified from Tires into Suspension & Steering; Transmission, Differentials, and Fuel System patterns strengthened.*

| L1 Category | Stores (v1) | Notes |
|---|---:|---|
| Wiper Blades | 284 | Stable — direct match |
| Shop & Misc | 283 | Stable — catch-all |
| Lighting | 283 | Stable — direct match |
| Air Filters | 277 | Stable — includes cabin + engine air |
| Battery | 263 | Stable — direct match |
| Fluids & Cooling | 264 | Stable — includes PS fluid, coolant, radiator |
| Emissions & Inspections | 260 | Stable — state programs dominate volume |
| Fuel System | 243 | **Patterns strengthened** — added injector clean, throttle body, BG 44K |
| Engine | 240 | Stable |
| Transmission | 229 | **Patterns strengthened** — added MERCON, DEXRON, ATF, TRANS FLUID variants |
| Brakes | 224 | Stable — direct match |
| **Tire Services** *(new)* | ~193 | Rotation, TPMS, repair, disposal — split from former "Tires" L1 |
| **Tire Sales** *(new)* | ~73 | Physical tire replacement/installation — split from former "Tires" L1 |
| Differentials | 182 | **Patterns strengthened** — added GEARBOX, GEAR LUBE, AXLE OIL, viscosity variants |
| Suspension & Steering | ~95 | **Now includes Alignment** — was 88 stores; alignment transactions absorbed from former Tires L1 |
| Air Conditioning | 54 | Stable — not tracked in operations data |

*Store counts for updated categories reflect v1 baseline; will update after next Snowflake refresh.*

---

### Full L1 → L2 Hierarchy with Matching Logic

Classification is applied in `store_services_matrix_v_FILTERS.sql` via `CASE/WHEN` on `UPPER(ITEM_DESCR)`. The Python enrichment script (`enrich.py`) also applies a supplemental patch layer for items the SQL marks as "Other" — catching case variants and free-text synonyms missed by the SQL regex. Both layers are documented below.

---

#### BRAKES · 224 stores · 7,156 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Brake Pads — Front | 171 | 1,454 | `FRONT BRAKE PAD`, `FRONT BRAKE REPLACE`, `FRONT PADS`, `FRONT BRAKES` · lower: `front pads`, `front brake pads` |
| Brake Pads — Rear | 159 | 1,203 | `REAR BRAKE PAD`, `REAR BRAKE REPLACE`, `REAR PADS`, `REAR BRAKES` · lower: `rear pads`, `pads` |
| Brake Fluid | 183 | 1,659 | RLIKE `BRAKE FLUID\|DOT 3\|DOT 4` |
| Rotors — Front | 113 | 380 | `FRONT ROTOR`, `FRONT ROTORS`, `FRONT BRAKE ROTOR` · lower: `front rotor` |
| Rotors — Rear | 98 | 324 | `REAR ROTOR`, `REAR ROTORS`, `REAR BRAKE ROTOR` · lower: `rear rotor` |
| Rotors (General) | 109 | 845 | Contains `ROTOR` (not matched by Front/Rear rules) |
| Brake Labor | 133 | 862 | `BRAKE LABOR`, `BRAKE INSPECTION`, `CALIPER`, `CHECKLIST` |
| Brake Pads (General) | 58 | 119 | Contains `BRAKE PAD` or lower equals `brake pads` |

---

#### TIRES · 202 stores · 4,204 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Tire Rotation | 193 | 1,015 | `TIRE ROTATION` |
| TPMS | 34 | 570 | `TPMS` |
| Tire Mount & Balance | 30 | 631 | `TIRE MOUNT`, `TIRE BALANCE`, `TIRE BALANCING` |
| Wheel Alignment | 25 | 346 | `ALIGNMENT` |
| Tire Repair | 23 | 232 | `FLAT TIRE`, `TIRE REPAIR` |
| Tire Disposal | 22 | 331 | `TIRE DISPOSAL` (excludes `NO TIRE TAX`) |
| Tire Replacement | 10 | 144 | RLIKE `REPLACED.*TIRE` (e.g. `!REPLACED 1 TIRE` through `!REPLACED 6 TIRES`) |

---

#### BATTERY · 263 stores · 2,521 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Battery Replacement | 251 | 1,402 | RLIKE `DHGOL\|DHGO\|DHSVR\|DHGMP\|AGM PLATM\|GOLD DH\|SILVER DH\|H4 DIEHARD\|BATTERY REPLACEMENT\|DIEHARD GOLD` · RLIKE `^T[0-9]+ GOLD` (DieHard Gold part# format) |
| Battery Fees | 76 | 530 | `BATTERY SALES FEE`, `STATE BATTERY`, `TEXAS BATTERY` |
| Battery Service | 105 | 235 | Contains `BATTERY` (catch-all when not matched above) |
| Starting & Charging | 88 | 354 | `ALTERNATOR`, `STARTING & CHARGING`, `STARTER` · lower: `alternator`, `starter` |

---

#### ENGINE · 240 stores · 3,880 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Spark Plugs | 163 | 628 | RLIKE `SPARK PLUG\|COIL PAK` · lower: `sparkplugs`, `plugs`, `tune up`, `coil packs` |
| Engine Treatments | 136 | 357 | `OIL SYSTEM CLEANER`, `ENGINE CLEANER`, `ENGINE MAX`, `ENGINE TREATMENT`, `STOP LEAK` |
| Valve Cover Gasket | 98 | 275 | `VALVE COVER` |
| Ignition / Coils | 112 | 409 | `IGNITION COIL`, `IGNITION LABOR`, `COIL PACK` |
| Engine Labor | 86 | 908 | `ENGINE LABOR`, `ENGINE DIAGNOSTIC`, `COOLING SYSTEM LABOR`, `OVERLAP LABOR`, `EXHAUST LABOR` |
| Thermostat / Water Pump | 86 | 313 | RLIKE `THERMOSTAT\|WATER PUMP` |
| Belts | 79 | 274 | `SERPENTINE`, `BELT TENSION`, `DRIVE BELT` · lower: `belt`, `serp belt`, `idler pulley` |
| O2 / Sensors | 62 | 159 | RLIKE `O2 SENSOR\|02 SENSOR\|OXYGEN SENSOR\|DOWN STREAM\|UP STREAM\|BANK.*SENSOR\|PCV VALVE` |

---

#### TRANSMISSION · 229 stores · 2,498 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Transmission Fluid Exchange | 224 | 2,089 | `TRANSMISSION FLUID`, `CVT`, `DMX GLOBAL SYN ATF`, `HONDA DUAL PUMP FLUID`, `ATF` (as word) |
| Transfer Case | 90 | 139 | `TRANSFER CASE` |
| Drivetrain Labor | 24 | 76 | `DRIVETRAIN LABOR`, `TRANSMISSION LABOR` |

---

#### FLUIDS & COOLING · 264 stores · 5,281 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Coolant / Antifreeze | 256 | 3,717 | RLIKE `ANTIFREEZE\|COOLANT\|RADIATOR SERVICE\|TOP OFF ANTIFREEZE` |
| Power Steering Fluid | 119 | 390 | `POWER STEERING` (excludes labor) |
| Window Wash | 88 | 791 | RLIKE `WINDOW WASH\|WINDSHIELD WASH` |
| Radiator Replacement | 86 | 198 | `RADIATOR` (excludes `SERVICE`, `LABOR`) |
| Cooling System Labor | 49 | 181 | Contains both `COOLING` and `LABOR` |

---

#### DIFFERENTIALS · 182 stores · 1,007 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Gear Oil | 180 | 494 | RLIKE `GL5\|GEAR OIL\|75W-90\|75W-140\|80W-90\|LIMITED SLIP` |
| Rear Differential | 162 | 367 | `REAR DIFFERENTIAL` |
| Front Differential | 89 | 142 | `FRONT DIFFERENTIAL` |

---

#### FUEL SYSTEM · 243 stores · 1,562 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Fuel System Cleaning | 216 | 1,167 | RLIKE `FUEL SYSTEM CLEAN\|GDI FUEL\|FUEL JUELS` (note: "FUEL JUELS" is a recurring POS typo in the data) |
| Fuel Filter | 130 | 360 | `FUEL FILTER` · RLIKE `^F[0-9]{5}` (Fram/Wix filter part numbers e.g. `F80409`) |

---

#### AIR FILTERS · 277 stores · 3,052 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Engine Air Filter | 272 | 1,856 | `AIR FILTER` (excl. CABIN) · RLIKE `^A[0-9]{4}` (Wix/Fram part# format e.g. `A5642`) |
| Cabin Air Filter | 254 | 1,196 | `CABIN AIR FILTER` · lower: `cabin filter` · RLIKE `^C[0-9]{5}` (Wix cabin part# format e.g. `C38175`) |

---

#### WIPER BLADES · 284 stores · 8,873 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Front Wiper Blades | 284 | 8,441 | `WIPER BLADE`, `BB PREMIUM WIPER`, `CB WIPER` (blade size codes e.g. `26BB PREMIUM WIPER BLADE`) |
| Rear Wipers | 159 | 432 | `REAR WIPER` · RLIKE `\d+-[12] REAR WIPER` (size-coded e.g. `12-1 REAR WIPER`, `11-2 REAR WIPER`) |

---

#### LIGHTING · 283 stores · 5,841 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Interior / Signal Bulbs | 272 | 3,258 | RLIKE ANSI bulb number codes: `168`, `194`, `906`, `912`, `921`, `1156`, `1157`, `2057`, `2357`, `3057`, `3157`, `3457`, `3757`, `4157`, `6418`, `7440`, `7443`, `7444`, or contains `LIGHT BULB`, `BULB` |
| Headlight Bulbs | 269 | 2,452 | RLIKE H-series and 9xxx codes: `H1`, `H3`, `H4`, `H7`, `H8`, `H9`, `H10`, `H11`, `H13`, `9003`–`9012`, `HALOGEN BULB`, `HEADLIGHT BULB` |
| Headlight Restoration | 48 | 131 | `HEADLIGHT RESTORATION` |

---

#### EMISSIONS & INSPECTIONS · 260 stores · 66,741 transactions

*Note: High transaction count is driven by state vehicle inspection programs in Georgia, Texas, North Carolina, and Utah. This inflates this category significantly versus organic service demand.*

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Vehicle Check / Inspection | 237 | 9,662 | RLIKE `VEHICLE CHECK\|DIAGNOSTIC\|SERVICE CHECKLIST\|INSPECTION\|UBER VEHICLE` |
| Emissions Test | 95 | 56,886 | RLIKE `EMISSIONS TEST\|GEORGIA EMISSIONS\|TEXAS EMISSIONS\|UTAH EMISSIONS\|EMISSION CERTIFICATE\|EMISSION STICKER\|SAFETY STICKER\|NC STATE EMIS\|NC STATE SAFE\|FAILED NC STATE\|ON THE SPOT RENEWAL` |

---

#### AIR CONDITIONING · 54 stores · 342 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| A/C Service | 54 | 342 | RLIKE `AIR CONDITION\|REFRIGERANT\|R-134\|R134\|A/C COMPRESSOR\|AC COMPRESSOR` |

---

#### SUSPENSION & STEERING · 88 stores · 499 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Suspension / Steering Labor | 88 | 499 | RLIKE `SUSPENSION\|STEERING LABOR\|LOWER CONTROL ARM\|UPPER CONTROL ARM\|SWAY BAR\|STRUT\|WHEEL HUB\|TIE ROD\|WHEEL STUD` |

---

#### SHOP & MISC · 283 stores · 18,530 transactions

| L2 Subcategory | Stores | Tx | Matching Logic |
|---|---:|---:|---|
| Shop Supplies | 282 | 12,505 | `SHOP SUPPLIES` |
| Shop Labor | 227 | 3,219 | RLIKE `SHOP LABOR\|MISC.*LABOR\|LABOR.*NO CHARGE` · lower: `install`, `labor`, `top off` |
| Lug Nut Service | 212 | 2,360 | `LUG NUT` |
| Car Wash | 14 | 324 | `CAR WASH` |

---

### Classification Design Decisions

**Why SQL CASE/WHEN rather than Python post-processing?**
The hierarchy lives in the Snowflake query itself. Every refresh produces pre-classified output. The enrichment script handles only brand and DMA assignment — not reclassification. This means the hierarchy is self-documenting at the query level and does not require running additional code.

**Why exclude oil change transactions entirely (not just the oil lines)?**
An oil change visit that also includes a wiper blade swap should not count as evidence of wiper blade service capability. The intent is to identify stores where customers proactively seek out non-oil services. Transaction-level exclusion enforces this distinction.

**Why 94% classification and not 100%?**
The remaining 6% is genuine noise — technician-typed notes, one-off part descriptions, and single-store POS idiosyncrasies. Classifying them would require either overfitting patterns to individual stores or creating meaninglessly broad catch-all rules. All meaningful service categories are captured at scale in the 94%.

**Known classification edge cases:**

| Issue | Detail |
|---|---|
| `ITEM_TYPE` is free-text | Not a controlled enum. Values include full descriptions (`WIPER BLADES`), codes (`ATF`), and garbage. All classification uses `ITEM_DESCR` only. |
| Fuel System / Radiator near-zero rates | Flagged in prior Recommended Services journey analysis — may be POS code mapping gaps, not true zero service provision. Treat with caution. |
| Emissions & Inspections volume | 56,886 emissions test transactions at 95 stores is state-program-driven (GA, TX, NC, UT). This inflates the L1 % significantly beyond organic service demand. |
| Brake-Inspection overlap | Some brake-related technician notes containing inspection language fall into Vehicle Check. Known overlap, low volume impact. |
| Tires L1 split | "Tires" (v1) was split into "Tire Sales" (replacement/installation) and "Tire Services" (rotation, TPMS, repair, disposal) after cross-validation showed our original Tires L1 inflated store count by ~138 stores vs operations data — we were capturing rotation-only stores as "tire" stores. |
| Alignment reclassified | Wheel alignment transactions moved from Tires L1 to Suspension & Steering L1, consistent with operations framework which treats alignment as a capability indicator separate from tire service. |
| Gearbox = Differentials | Operations data labels gear oil service as "Gearbox"; our hierarchy uses "Differentials." Same service, different names. SQL patterns now include GEARBOX as a match term for Differentials. |

---

## Dashboard Features

**Live URL:** `https://fsasevans.github.io/ServicesByLocationDMA/`

### Global Filter Bar (sticky)
Four multi-select filters apply globally to all modules. Filters are hierarchically linked — selecting Brand/State/DMA narrows available City options. Active filter badge shows count of applied filters.

| Filter | Source | Options |
|---|---|---|
| Brand | Extracted from `STORE_NAME_FULL` | 8 brands |
| State | `LOCATIONS.STATE` | 24 states |
| DMA | Point-in-polygon assignment | 50 DMAs |
| City | `LOCATIONS.CITY` | 208 cities |

### Summary Strip
Four live KPI cards: Stores in View, DMAs Covered, States Active, Brands. All update on every filter change.

### Service Category Penetration (KPI Grid)
15 tiles, one per L1 category. Shows % of filtered stores with at least one qualifying transaction during the data window. Click any tile to open the store drilldown sidebar.

### Store Drilldown Sidebar
Slides in from the right on KPI tile click. Shows all stores in the current filter set offering that service, grouped by brand, sorted by transaction count. Includes live search across store name, city, state, and DMA.

### Store Location Map
CartoDB light-style basemap with Leaflet.js. Three color modes: Brand, DMA, Service (highlight/dim by selected category). Clickable pins with store details popup.

---

## Repository Structure

```
ServicesByLocationDMA/
├── index.html                             ← Live dashboard (auto-rebuilt by Actions)
├── store_services_matrix_v_FILTERS.sql    ← Primary Snowflake query
├── .gitignore
├── .github/
│   └── workflows/
│       └── rebuild-dashboard.yml          ← GitHub Actions automation
├── data/
│   └── README.md                          ← Data refresh instructions
│   └── [Snowflake_Export_YYYY-MM-DD.csv]  ← Drop new exports here
└── scripts/
    ├── enrich.py                           ← Brand extraction + DMA assignment + store model tier
    ├── build_dashboard.py                  ← Builds index.html from enriched data
    ├── dashboard_template.html             ← HTML shell with data placeholder
    ├── dma_names.csv                       ← Nielsen DMA code → name (210 DMAs)
    ├── dma_boundaries.csv                  ← Nielsen DMA polygon boundaries
    └── store_models.csv                    ← Store model tiers (QL/QL+/QL++/FS) from operations
```

---

## Refresh Cycle

**Recommended cadence:** Monthly (or quarterly)

```
Step 1 — Run Snowflake query (~2 min, manual)
   Open store_services_matrix_v_FILTERS.sql
   Run → Results → Download CSV

Step 2 — Upload CSV to GitHub (~1 min, manual)
   github.com/FSASEvans/ServicesByLocationDMA/data
   Add file → Upload files → Commit to main

Step 3 — Automated rebuild (~2 min, automatic)
   GitHub Actions: enrich.py → build_dashboard.py → commit index.html

Step 4 — Dashboard live for all users
   https://fsasevans.github.io/ServicesByLocationDMA/
```

The date window label in the dashboard header updates automatically to reflect the current run date on every rebuild.

---


---

## Cross-Validation Against Operations Data

In March 2026, our transaction-based hierarchy was cross-validated against the operations service capability matrix maintained by FullSpeed Automotive's Regional Operations team (Brian Brooks). This validation compared 285 stores appearing in both datasets and produced several hierarchy improvements.

### Methodology Comparison

| Dimension | Operations Matrix (Brian's Work) | Our Transaction Model |
|---|---|---|
| Data source | Self-reported + DM-verified annual activity | Snowflake POS transactions |
| Signal | Annual transaction volume per service type | 90-day presence (any transaction) |
| Primary question | What is this store *capable* of offering? | What has this store *actually transacted*? |
| Store count | 286 stores | 285 stores |
| Service taxonomy | 26 types across 4 groups (Mag 7, Preventive, Additives, Maintenance) | 16 L1 categories, 80 L2 subcategories |
| Store model tiers | QL / QL+ / QL++ / FS | Now included via `store_models.csv` join |

Their framework organizes services into four tiers by store model capability:
- **QL** — Oil changes + Magnificent 7 (air filters, batteries, cabin filters, light bulbs, wiper blades, oil system cleaner) + fluid preventive services
- **QL+** — All QL services + brakes, belt service, tune-ups, and limited maintenance. Requires lift + minimum B/C tech. Prohibited from internal engine work, transmission rebuilds, PCM reprogramming.
- **QL++** — Being eliminated in 2026; 26 stores being reclassified to QL+ or FS
- **FS (Full Service)** — Any mechanical service. Requires alignment rack + minimum A tech. Currently 8 stores, proposed 16.

### Key Validation Findings

**Strong alignment (no action needed):**
- Air Filters, Batteries, Wiper Blades, Lighting — near-perfect match across both datasets
- Brakes — 207 stores agree; 38 discrepancies explained by 90-day vs annual window
- Fluids & Cooling — 262/285 stores agree on radiator/cooling service

**Hierarchy improvements made (v2 SQL):**

| Issue | Root Cause | Fix Applied |
|---|---|---|
| Transmission under-count (48 stores) | SQL patterns too narrow — missing MERCON, DEXRON, ATF, TRANS FLUID variants | Added to L1 + L2 RLIKE patterns |
| Differentials under-count (98 stores) | "Gearbox" in their taxonomy = gear oil = our "Differentials"; GEARBOX not in our patterns | Added GEARBOX, GEAR LUBE, AXLE OIL, viscosity variants |
| Tires over-count (138 stores inflated) | Our "Tires" L1 included rotation, TPMS, and alignment — they count tire *replacement* only | Split into "Tire Sales" + "Tire Services" L1 categories |
| Alignment misclassified | Alignment matched under Tires L1; operations treats it as a capability indicator under Suspension | Moved to Suspension & Steering L1 |
| Fuel System under-count (40 stores) | Missing injector cleaning, throttle body, BG 44K product codes | Added to L1 + L2 RLIKE patterns |

**Structural differences (intentional — no change):**

| Our Category | Their Coverage | Reason for Difference |
|---|---|---|
| Emissions & Inspections | Not tracked | State inspection programs are regulatory, not capability-based |
| Air Conditioning | Not tracked | Equipment penetration inconsistently tracked by operations |
| Shop & Misc | No equivalent | Our catch-all for supplies, labor, lug nuts |
| Cabin Filters | Separate "Cabin Filters" column | We combine cabin + engine air under one "Air Filters" L1; split at L2 |

### Audit Queries for Future Snowflake Validation

To further validate Transmission and Differentials classification after next refresh, run this audit against specific stores where gaps were identified:

```sql
-- Identify ITEM_DESCR values at transmission-gap stores not landing in Transmission L1
-- Replace store list with current gap stores after next run
SELECT
    td.ITEM_DESCR,
    COUNT(DISTINCT t.TRANSACTION_ID) AS tx_count
FROM CINCH_DATA.CINCH_TO_FULLSPEED_AUTOMOTIVE_DATA_SHARE.TRANSACTIONS t
JOIN CINCH_DATA.CINCH_TO_FULLSPEED_AUTOMOTIVE_DATA_SHARE.TRANSACTIONS_DETAILS td
    ON t.TRANSACTION_ID = td.TRANSACTION_ID
WHERE t.LOCATION_ID IN (
    -- Add LOCATION_IDs for transmission-gap stores here
    SELECT LOCATION_ID FROM CINCH_DATA.CINCH_TO_FULLSPEED_AUTOMOTIVE_DATA_SHARE.LOCATIONS
    WHERE NAME IN ('1027 GMI Woodstock GA','1040 GMI Morrow GA','147 GMI Morrisville NC')
)
AND UPPER(td.ITEM_DESCR) NOT RLIKE '.*(TRANSMISSION|CVT|ATF|MERCON|DEXRON|TRANSFER CASE).*'
AND UPPER(td.ITEM_DESCR) RLIKE '.*(FLUID|TRANS|GEAR).*'
GROUP BY td.ITEM_DESCR
ORDER BY tx_count DESC;
```

## Known Data Quality Notes

| Issue | Detail | Status |
|---|---|---|
| Fuel System near-zero rates | May be POS code mapping gaps — SQL patterns now strengthened with injector clean, throttle body, BG 44K | Updated v2 |
| Radiator near-zero rates | Cross-validation confirms Fluids & Cooling captures radiator service correctly via COOLANT/ANTIFREEZE/RADIATOR patterns | Monitored |
| Transmission under-count | Cross-validation found 48 stores with operations-confirmed trans fluid but zero in our Transmission L1 — SQL patterns strengthened with MERCON, DEXRON, ATF, TRANS FLUID variants | Updated v2 |
| Differentials under-count | Cross-validation found 98 stores with operations-confirmed gearbox service but zero in our Differentials L1 — SQL patterns strengthened with GEARBOX, GEAR LUBE, AXLE OIL | Updated v2 |
| `ITEM_TYPE` free-text field | Not a controlled enum — all classification uses `ITEM_DESCR` pattern matching only | Ongoing |
| 6% unclassified transactions | Long-tail single-store POS free-text — intentionally excluded | Ongoing |
| Emissions & Inspections volume | State inspection programs in GA, TX, NC, UT inflate this category (56,886 of 66,741 tx) | Ongoing |
| 2 coastal stores manual DMA | Rockport TX → Corpus Christi (600); Gulf Breeze FL → Mobile-Pensacola (686) | Fixed |
| Store #8130 | Exists in operations data but not in our Snowflake pull — investigate | Open |

---

## Team

**Produced by:** Marketing Performance & Insights
**Analyst:** Sam Evans
**Strategic input:** Zach (journey objective frameworks and service review); Brian Brooks (operations service capability matrix and store model tier framework)
**Data platform:** Cinch / Snowflake (read-only data share)
**Dashboard hosting:** GitHub Pages — `FSASEvans/ServicesByLocationDMA`
**Internal use only — not for external distribution**
