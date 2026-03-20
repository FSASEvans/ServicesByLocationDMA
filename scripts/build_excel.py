"""
build_excel.py
--------------
Builds the FSA Store Service Intelligence Excel workbook from the enriched CSV.
Produces a date-stamped file in the /data folder so every run is preserved.

Usage:
  python scripts/build_excel.py \
    --input  data/store_services_enriched.csv \
    --output data/FSA_Store_Service_Matrix_YYYY-MM-DD.xlsx

Output:
  8-sheet workbook:
    Store Directory, L1 Category Matrix, L2 Subcategory Matrix,
    Brand Summary, Model Summary, DMA Summary,
    Hierarchy Reference, Raw Data
"""

import csv, sys, argparse
from collections import defaultdict
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip install openpyxl")
    sys.exit(1)

# ── CONSTANTS ──────────────────────────────────────────────────────────────────
CAT_COLORS = {
    'Air Conditioning':       '0EA5E9',
    'Air Filters':            '16A34A',
    'Battery':                'D97706',
    'Brakes':                 'DC2626',
    'Differentials':          '7C3AED',
    'Emissions & Inspections':'4F46E5',
    'Engine':                 '2563EB',
    'Fluids & Cooling':       '0891B2',
    'Fuel System':            'B45309',
    'Lighting':               'CA8A04',
    'Shop & Misc':            '6B7280',
    'Suspension & Steering':  '059669',
    'Tires':                  '65A30D',
    'Transmission':           '9333EA',
    'Wiper Blades':           '475569',
}

AUTHORITATIVE_L2 = {
    'Air Conditioning':        ['A/C Service'],
    'Air Filters':             ['Cabin Air Filter', 'Engine Air Filter'],
    'Battery':                 ['Battery Fees', 'Battery Replacement', 'Battery Service', 'Starting & Charging'],
    'Brakes':                  ['Brake Labor', 'Brake Pads — Front', 'Brake Pads — Rear',
                                'Brake Pads (General)', 'Rotors — Front', 'Rotors — Rear', 'Rotors (General)'],
    'Differentials':           ['Front Differential', 'Gear Oil', 'Rear Differential'],
    'Emissions & Inspections': ['Emissions Test', 'Vehicle Check / Inspection'],
    'Engine':                  ['Belts', 'Engine Labor', 'Engine Treatments', 'Ignition / Coils',
                                'O2 / Sensors', 'Spark Plugs', 'Thermostat / Water Pump', 'Valve Cover Gasket'],
    'Fluids & Cooling':        ['Brake Fluid', 'Coolant / Antifreeze', 'Cooling System Labor', 'Power Steering Fluid',
                                'Radiator Replacement', 'Window Wash'],
    'Fuel System':             ['Fuel Filter', 'Fuel System Cleaning'],
    'Lighting':                ['Headlight Bulbs', 'Headlight Restoration', 'Interior / Signal Bulbs'],
    'Shop & Misc':             ['Car Wash', 'Lug Nut Service', 'Shop Labor', 'Shop Supplies'],
    'Suspension & Steering':   ['Suspension / Steering Labor', 'Wheel Alignment'],
    'Oil Change':              ['Oil Change Service'],
    'Tires':                   ['Tire Disposal', 'Tire Mount & Balance', 'Tire Repair',
                                'Tire Replacement', 'Tire Rotation', 'TPMS'],
    'Transmission':            ['Drivetrain Labor', 'Transfer Case', 'Transmission Fluid Exchange'],
    'Wiper Blades':            ['Front Wiper Blades', 'Rear Wipers'],
}

L1_CATS   = list(AUTHORITATIVE_L2.keys())
BRANDS    = sorted(['American Lubefast','Economy Oil Change','Grease Monkey','Grease Monkey (ADI)',
                    'Herbert Automotive','Kwik Kar','SpeeDee',"Uncle Ed's"])
ALL_MODELS = ['QL', 'QL+', 'QL++', 'FS']
MODEL_COLORS = {'QL':'1D4ED8','QL+':'7C3AED','QL++':'D97706','FS':'DC2626'}

FSA_RED = 'C8102E'; DARK = '1F2937'; DARK2 = '374151'
WHITE   = 'FFFFFF'; LGREY = 'F9FAFB'; MGREY = 'F3F4F6'
DGREY   = '6B7280'; BLACK = '111827'

# ── STYLE HELPERS ──────────────────────────────────────────────────────────────
def fill(h):
    return PatternFill('solid', start_color=h, end_color=h)

def font(bold=False, color='000000', size=11, italic=False):
    return Font(name='Arial', bold=bold, color=color, size=size, italic=italic)

def bdr():
    s = Side(style='thin', color='D1D5DB')
    return Border(top=s, bottom=s, left=s, right=s)

def align(h='left', v='center', wrap=False, rot=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, textRotation=rot)

def hm_fill(pct):
    if not pct: return fill(WHITE)
    if pct >= 0.60: return fill('D1FAE5')
    if pct >= 0.20: return fill('FEF3C7')
    return fill('FEE2E2')

def hm_font(pct):
    if not pct: return font(color=DGREY, size=10)
    if pct >= 0.60: return font(bold=True, color='065F46', size=10)
    if pct >= 0.20: return font(color='92400E', size=10)
    return font(color='991B1B', size=10)

def hm_val(pct):
    return f'{pct:.0%}' if pct else '—'

def set_title(ws, text, n_cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(row=1, column=1, value=text)
    c.font = font(bold=True, color=WHITE, size=13)
    c.fill = fill(DARK); c.alignment = align('left', 'center')
    ws.row_dimensions[1].height = 26

def set_left_hdr(ws, row, col, text, width=None):
    c = ws.cell(row=row, column=col, value=text)
    c.font = font(bold=True, color=WHITE, size=11)
    c.fill = fill(DARK2); c.alignment = align('left', 'center'); c.border = bdr()
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width

def set_cat_hdr(ws, row, col, cat, rotation=60, width=11):
    color = CAT_COLORS.get(cat, '6B7280')
    c = ws.cell(row=row, column=col, value=cat)
    c.font = font(bold=True, color=WHITE, size=11)
    c.fill = fill(color); c.alignment = align('center', 'bottom', rot=rotation); c.border = bdr()
    ws.column_dimensions[get_column_letter(col)].width = width

def set_count_cell(ws, row, col, count, color_hex):
    c = ws.cell(row=row, column=col, value=count)
    c.font = font(bold=True, color=color_hex, size=10)
    c.fill = fill(MGREY); c.alignment = align('center', 'center'); c.border = bdr()

def set_count_label(ws, row, col_start, col_end, text='Stores offering service →'):
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    c = ws.cell(row=row, column=col_start, value=text)
    c.font = font(bold=True, color=DGREY, size=9, italic=True)
    c.fill = fill(MGREY); c.alignment = align('right', 'center'); c.border = bdr()
    ws.row_dimensions[row].height = 16


# ── MAIN BUILD FUNCTION ────────────────────────────────────────────────────────
def build(input_path, output_path, region_district_path=None):
    print(f"Reading: {input_path}")
    with open(input_path, newline='', encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))
    print(f"  {len(rows):,} rows")

    # Derive date window from today for title bar
    today = datetime.now()
    from_date_str = "Dec 17, 2025"  # fallback; ideally passed as arg
    to_date_str   = today.strftime('%b %d, %Y')
    date_window   = f"{from_date_str} – {to_date_str}"

    # ── Build lookups ──────────────────────────────────────────────────────────
    # Load region/district lookup
    region_dist = {}
    if region_district_path:
        try:
            with open(region_district_path) as f:
                for row in csv.DictReader(f):
                    region_dist[str(row['STORE_NUMBER']).strip()] = {
                        'REGION_NUM':  row.get('REGION_NUM','').strip(),
                        'REGION_VP':   row.get('REGION_VP','').strip(),
                        'DISTRICT_NUM':row.get('DISTRICT_NUM','').strip(),
                        'DISTRICT_MGR':row.get('DISTRICT_MGR','').strip(),
                    }
            print(f"  Loaded {len(region_dist)} region/district records")
        except FileNotFoundError:
            print(f"  WARNING: region_district.csv not found — region/district columns will be empty")

    store_map = {}
    for r in rows:
        sn = r['STORE_NUMBER']
        if sn not in store_map:
            rd = region_dist.get(sn, {'REGION_NUM':'','REGION_VP':'','DISTRICT_NUM':'','DISTRICT_MGR':''})
            if not rd['REGION_NUM'] and r.get('REGION_NUM'):
                rd = {'REGION_NUM':r.get('REGION_NUM',''),'REGION_VP':r.get('REGION_VP',''),
                      'DISTRICT_NUM':r.get('DISTRICT_NUM',''),'DISTRICT_MGR':r.get('DISTRICT_MGR','')}
            store_map[sn] = dict(
                STORE_NUMBER=sn, STORE_NAME=r['STORE_NAME'],
                BRAND=r['BRAND'], CITY=r['CITY'], STATE=r['STATE'],
                ADDRESS1=r['ADDRESS1'], ZIP_CODE=r['ZIP_CODE'],
                DMA_NAME=r['DMA_NAME'],
                STORE_MODEL=r.get('STORE_MODEL',''),
                PROPOSED_MODEL=r.get('PROPOSED_MODEL',''),
                REGION_NUM=rd['REGION_NUM'], REGION_VP=rd['REGION_VP'],
                DISTRICT_NUM=rd['DISTRICT_NUM'], DISTRICT_MGR=rd['DISTRICT_MGR'],
                l1_tx=defaultdict(int), l2_tx=defaultdict(int),      # non-oil only
                l1_tx_all=defaultdict(int), l2_tx_all=defaultdict(int)  # oil-inclusive
            )
        if r['OFFERS_SERVICE'] == '1':
            is_oil = str(r.get('IS_OIL_TRANSACTION','0')).strip() == '1'
            store_map[sn]['l1_tx_all'][r['L1_CATEGORY']] += int(r['TRANSACTION_COUNT'])
            store_map[sn]['l2_tx_all'][(r['L1_CATEGORY'], r['L2_SUBCATEGORY'])] += int(r['TRANSACTION_COUNT'])
            if not is_oil:
                store_map[sn]['l1_tx'][r['L1_CATEGORY']] += int(r['TRANSACTION_COUNT'])
                store_map[sn]['l2_tx'][(r['L1_CATEGORY'], r['L2_SUBCATEGORY'])] += int(r['TRANSACTION_COUNT'])

    stores = sorted(store_map.values(), key=lambda x: (x['REGION_NUM'] or '9', x['DISTRICT_NUM'] or '999', x['BRAND'], x['STORE_NAME']))
    n_stores = len(stores)
    print(f"  {n_stores} unique stores")

    l1_stats = defaultdict(lambda: {'stores': set(), 'tx': 0})
    l2_stats = defaultdict(lambda: {'stores': set(), 'tx': 0})
    brand_stores = defaultdict(set);  brand_l1 = defaultdict(lambda: defaultdict(set))
    model_stores = defaultdict(set);  model_l1 = defaultdict(lambda: defaultdict(set))
    dma_stores   = defaultdict(set);  dma_l1   = defaultdict(lambda: defaultdict(set))

    for r in rows:
        if r['OFFERS_SERVICE'] == '1':
            l1_stats[r['L1_CATEGORY']]['stores'].add(r['STORE_NUMBER'])
            l1_stats[r['L1_CATEGORY']]['tx'] += int(r['TRANSACTION_COUNT'])
            if r['L2_SUBCATEGORY']:
                l2_stats[(r['L1_CATEGORY'], r['L2_SUBCATEGORY'])]['stores'].add(r['STORE_NUMBER'])
                l2_stats[(r['L1_CATEGORY'], r['L2_SUBCATEGORY'])]['tx'] += int(r['TRANSACTION_COUNT'])

    for s in stores:
        brand_stores[s['BRAND']].add(s['STORE_NUMBER'])
        model_stores[s['STORE_MODEL']].add(s['STORE_NUMBER'])
        dma_stores[s['DMA_NAME']].add(s['STORE_NUMBER'])
        for cat in s['l1_tx']:
            brand_l1[s['BRAND']][cat].add(s['STORE_NUMBER'])
            model_l1[s['STORE_MODEL']][cat].add(s['STORE_NUMBER'])
            dma_l1[s['DMA_NAME']][cat].add(s['STORE_NUMBER'])

    DMAS = sorted(dma_stores.keys())
    n_cats = len(L1_CATS)
    all_l2 = [(l1, l2) for l1 in L1_CATS for l2 in AUTHORITATIVE_L2[l1]]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── SHEET 1: STORE DIRECTORY ───────────────────────────────────────────────
    ws = wb.create_sheet('Store Directory')
    ws.freeze_panes = 'A4'
    set_title(ws, f'FullSpeed Automotive — Store Directory  |  {n_stores} Active Locations  |  {date_window}', 10)
    ws.merge_cells('A2:J2')
    c = ws.cell(row=2, column=1, value='All active locations with brand, geography, DMA assignment, and store model tier')
    c.font = font(color=DGREY, size=10, italic=True); c.fill = fill(MGREY); c.alignment = align('left','center')
    ws.row_dimensions[2].height = 16

    for col, (hdr, w) in enumerate([('Store #',9),('Brand',22),('Store Name',32),('Address',26),
                                     ('City',16),('State',7),('Zip',8),('DMA',26),('Model',9),('Proposed 2026',13)], 1):
        set_left_hdr(ws, 3, col, hdr, width=w)
    ws.row_dimensions[3].height = 18

    for i, s in enumerate(stores, 4):
        bg = LGREY if i % 2 == 0 else WHITE
        for col, val in enumerate([s['STORE_NUMBER'],s['BRAND'],s['STORE_NAME'],s['ADDRESS1'],
                                    s['CITY'],s['STATE'],s['ZIP_CODE'],s['DMA_NAME'],
                                    s['STORE_MODEL'],s['PROPOSED_MODEL']], 1):
            c = ws.cell(row=i, column=col, value=val)
            c.font = font(size=10, color=BLACK); c.fill = fill(bg)
            c.alignment = align('left','center'); c.border = bdr()
        ws.row_dimensions[i].height = 15

    ws.auto_filter.ref = f'A3:J{3+n_stores}'

    # ── SHEET 2: L1 CATEGORY MATRIX ───────────────────────────────────────────
    ws = wb.create_sheet('L1 Category Matrix')
    ws.freeze_panes = 'E4'
    set_title(ws, f'FullSpeed Automotive — Service Category Offering Matrix  |  Transaction count per store per category', 4+n_cats)

    for col, (hdr, w) in enumerate([('Store #',9),('Brand',22),('City, State',20),('DMA',26)], 1):
        set_left_hdr(ws, 2, col, hdr, width=w)
    for col, cat in enumerate(L1_CATS, 5):
        set_cat_hdr(ws, 2, col, cat, rotation=60)
    ws.row_dimensions[2].height = 80

    set_count_label(ws, 3, 1, 4)
    for col, cat in enumerate(L1_CATS, 5):
        set_count_cell(ws, 3, col, len(l1_stats[cat]['stores']), CAT_COLORS.get(cat,'6B7280'))

    for i, s in enumerate(stores, 4):
        bg = LGREY if i % 2 == 0 else WHITE
        for col, val in enumerate([s['STORE_NUMBER'],s['BRAND'],f"{s['CITY']}, {s['STATE']}",s['DMA_NAME']], 1):
            c = ws.cell(row=i, column=col, value=val)
            c.font = font(size=10, bold=(col==2), color=BLACK)
            c.fill = fill(bg); c.alignment = align('left','center'); c.border = bdr()
        for col, cat in enumerate(L1_CATS, 5):
            tx = s['l1_tx'].get(cat, 0)
            c = ws.cell(row=i, column=col, value=tx if tx else None)
            c.fill = fill('ECFDF5') if tx else fill(bg)
            c.font = font(bold=bool(tx), color='065F46' if tx else DGREY, size=10)
            c.alignment = align('center','center'); c.border = bdr()
        ws.row_dimensions[i].height = 15

    ws.auto_filter.ref = f'A2:{get_column_letter(4+n_cats)}{3+n_stores}'

    # ── SHEET 2b: L1 MATRIX — OIL INCLUSIVE ──────────────────────────────────
    ws = wb.create_sheet('L1 Matrix (Oil Incl.)')
    ws.freeze_panes = 'E4'
    set_title(ws, f'FullSpeed Automotive — Service Matrix (Oil-Inclusive)  |  Includes services bundled with oil change visits', 4+n_cats)
    for col, (hdr, w) in enumerate([('Store #',9),('Brand',22),('City, State',20),('DMA',26)], 1):
        set_left_hdr(ws, 2, col, hdr, width=w)
    for col, cat in enumerate(L1_CATS, 5):
        set_cat_hdr(ws, 2, col, cat, rotation=60)
    ws.row_dimensions[2].height = 80
    set_count_label(ws, 3, 1, 4)
    for col, cat in enumerate(L1_CATS, 5):
        oil_stores = {r['STORE_NUMBER'] for r in rows if r['OFFERS_SERVICE']=='1' and r['L1_CATEGORY']==cat}
        set_count_cell(ws, 3, col, len(oil_stores), CAT_COLORS.get(cat,'6B7280'))
    for i, s in enumerate(stores, 4):
        bg = LGREY if i % 2 == 0 else WHITE
        for col, val in enumerate([s['STORE_NUMBER'],s['BRAND'],f"{s['CITY']}, {s['STATE']}",s['DMA_NAME']], 1):
            c = ws.cell(row=i, column=col, value=val)
            c.font = font(size=10, bold=(col==2), color=BLACK)
            c.fill = fill(bg); c.alignment = align('left','center'); c.border = bdr()
        for col, cat in enumerate(L1_CATS, 5):
            tx = s['l1_tx_all'].get(cat, 0)
            c = ws.cell(row=i, column=col, value=tx if tx else None)
            c.fill = fill('EFF6FF') if tx else fill(bg)
            c.font = font(bold=bool(tx), color='1E40AF' if tx else DGREY, size=10)
            c.alignment = align('center','center'); c.border = bdr()
        ws.row_dimensions[i].height = 15
    ws.auto_filter.ref = f'A2:{get_column_letter(4+n_cats)}{3+n_stores}'

    # ── SHEET 3: L2 SUBCATEGORY MATRIX ────────────────────────────────────────
    ws = wb.create_sheet('L2 Subcategory Matrix')
    ws.freeze_panes = 'E5'
    set_title(ws, f'FullSpeed Automotive — Service Subcategory Matrix  |  {len(all_l2)} L2 subcategories across {n_stores} stores', 4+len(all_l2))

    for col in range(1, 5):
        c = ws.cell(row=2, column=col, value='')
        c.fill = fill(DARK2); c.border = bdr()
    col = 5
    for l1 in L1_CATS:
        l2s = AUTHORITATIVE_L2[l1]; color = CAT_COLORS.get(l1,'6B7280')
        if len(l2s) > 1:
            ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col+len(l2s)-1)
        c = ws.cell(row=2, column=col, value=l1)
        c.font = font(bold=True, color=WHITE, size=10)
        c.fill = fill(color); c.alignment = align('center','center'); c.border = bdr()
        col += len(l2s)
    ws.row_dimensions[2].height = 18

    for col, (hdr, w) in enumerate([('Store #',9),('Brand',20),('City, State',18),('DMA',24)], 1):
        set_left_hdr(ws, 3, col, hdr, width=w)
    for col, (l1, l2) in enumerate(all_l2, 5):
        color = CAT_COLORS.get(l1,'6B7280')
        c = ws.cell(row=3, column=col, value=l2)
        c.font = font(bold=True, color=WHITE, size=10)
        c.fill = fill(color); c.alignment = align('center','bottom', rot=60); c.border = bdr()
        ws.column_dimensions[get_column_letter(col)].width = 9
    ws.row_dimensions[3].height = 80

    set_count_label(ws, 4, 1, 4)
    for col, (l1, l2) in enumerate(all_l2, 5):
        set_count_cell(ws, 4, col, len(l2_stats[(l1,l2)]['stores']), CAT_COLORS.get(l1,'6B7280'))

    for i, s in enumerate(stores, 5):
        bg = LGREY if i % 2 == 0 else WHITE
        for col, val in enumerate([s['STORE_NUMBER'],s['BRAND'],f"{s['CITY']}, {s['STATE']}",s['DMA_NAME']], 1):
            c = ws.cell(row=i, column=col, value=val)
            c.font = font(size=10, bold=(col==2), color=BLACK)
            c.fill = fill(bg); c.alignment = align('left','center'); c.border = bdr()
        for col, (l1, l2) in enumerate(all_l2, 5):
            tx = s['l2_tx'].get((l1,l2), 0)
            c = ws.cell(row=i, column=col, value=tx if tx else None)
            c.fill = fill('ECFDF5') if tx else fill(bg)
            c.font = font(bold=bool(tx), color='065F46' if tx else DGREY, size=10)
            c.alignment = align('center','center'); c.border = bdr()
        ws.row_dimensions[i].height = 15

    ws.auto_filter.ref = f'A3:{get_column_letter(4+len(all_l2))}{4+n_stores}'

    # ── SHEET 4: BRAND SUMMARY ─────────────────────────────────────────────────
    ws = wb.create_sheet('Brand Summary')
    ws.freeze_panes = 'B4'
    n_brands = len(BRANDS)
    set_title(ws, 'FullSpeed Automotive — Service Penetration by Brand  |  % of brand stores offering each category', 2+n_brands)

    c = ws.cell(row=2, column=1, value='Service Category')
    c.font = font(bold=True, color=WHITE, size=11); c.fill = fill(DARK2)
    c.alignment = align('left','center'); c.border = bdr()
    ws.column_dimensions['A'].width = 28

    for col, brand in enumerate(BRANDS, 2):
        c = ws.cell(row=2, column=col, value=brand)
        c.font = font(bold=True, color=WHITE, size=11); c.fill = fill(FSA_RED)
        c.alignment = align('center','bottom', rot=45); c.border = bdr()
        ws.column_dimensions[get_column_letter(col)].width = 14

    c = ws.cell(row=2, column=2+n_brands, value='All Brands')
    c.font = font(bold=True, color=WHITE, size=11); c.fill = fill(DARK)
    c.alignment = align('center','bottom', rot=45); c.border = bdr()
    ws.column_dimensions[get_column_letter(2+n_brands)].width = 14
    ws.row_dimensions[2].height = 80

    c = ws.cell(row=3, column=1, value='Store count →')
    c.font = font(bold=True, color=DGREY, size=9, italic=True)
    c.fill = fill(MGREY); c.alignment = align('right','center'); c.border = bdr()
    for col, brand in enumerate(BRANDS, 2):
        set_count_cell(ws, 3, col, len(brand_stores[brand]), FSA_RED)
    set_count_cell(ws, 3, 2+n_brands, n_stores, DARK)
    ws.row_dimensions[3].height = 16

    for i, l1 in enumerate(L1_CATS, 4):
        color = CAT_COLORS.get(l1,'6B7280')
        c = ws.cell(row=i, column=1, value=l1)
        c.font = font(bold=True, color=WHITE, size=11); c.fill = fill(color)
        c.alignment = align('left','center'); c.border = bdr()
        for j, brand in enumerate(BRANDS, 2):
            bs = len(brand_stores[brand]); os_ = len(brand_l1[brand][l1])
            pct = os_ / bs if bs else 0
            c = ws.cell(row=i, column=j, value=hm_val(pct))
            c.fill = hm_fill(pct); c.font = hm_font(pct)
            c.alignment = align('center','center'); c.border = bdr()
        pct = len(l1_stats[l1]['stores']) / n_stores
        c = ws.cell(row=i, column=2+n_brands, value=hm_val(pct))
        c.fill = hm_fill(pct); c.font = hm_font(pct)
        c.alignment = align('center','center'); c.border = bdr()
        ws.row_dimensions[i].height = 18

    # ── SHEET 5: MODEL SUMMARY ─────────────────────────────────────────────────
    ws = wb.create_sheet('Model Summary')
    ws.freeze_panes = 'B4'
    n_models = len(ALL_MODELS)
    set_title(ws, 'FullSpeed Automotive — Service Penetration by Store Model  |  % of model-tier stores offering each category', 2+n_models)

    c = ws.cell(row=2, column=1, value='Service Category')
    c.font = font(bold=True, color=WHITE, size=11); c.fill = fill(DARK2)
    c.alignment = align('left','center'); c.border = bdr()
    ws.column_dimensions['A'].width = 28

    for col, m in enumerate(ALL_MODELS, 2):
        c = ws.cell(row=2, column=col, value=m)
        c.font = font(bold=True, color=WHITE, size=11); c.fill = fill(MODEL_COLORS[m])
        c.alignment = align('center','bottom', rot=45); c.border = bdr()
        ws.column_dimensions[get_column_letter(col)].width = 14

    c = ws.cell(row=2, column=2+n_models, value='All Brands')
    c.font = font(bold=True, color=WHITE, size=11); c.fill = fill(DARK)
    c.alignment = align('center','bottom', rot=45); c.border = bdr()
    ws.column_dimensions[get_column_letter(2+n_models)].width = 14
    ws.row_dimensions[2].height = 80

    c = ws.cell(row=3, column=1, value='Store count →')
    c.font = font(bold=True, color=DGREY, size=9, italic=True)
    c.fill = fill(MGREY); c.alignment = align('right','center'); c.border = bdr()
    for col, m in enumerate(ALL_MODELS, 2):
        set_count_cell(ws, 3, col, len(model_stores.get(m,set())), MODEL_COLORS[m])
    set_count_cell(ws, 3, 2+n_models, n_stores, DARK)
    ws.row_dimensions[3].height = 16

    for i, l1 in enumerate(L1_CATS, 4):
        color = CAT_COLORS.get(l1,'6B7280')
        c = ws.cell(row=i, column=1, value=l1)
        c.font = font(bold=True, color=WHITE, size=11); c.fill = fill(color)
        c.alignment = align('left','center'); c.border = bdr()
        for j, m in enumerate(ALL_MODELS, 2):
            ms = len(model_stores.get(m,set())); os_ = len(model_l1[m][l1])
            pct = os_ / ms if ms else 0
            c = ws.cell(row=i, column=j, value=hm_val(pct))
            c.fill = hm_fill(pct); c.font = hm_font(pct)
            c.alignment = align('center','center'); c.border = bdr()
        pct = len(l1_stats[l1]['stores']) / n_stores
        c = ws.cell(row=i, column=2+n_models, value=hm_val(pct))
        c.fill = hm_fill(pct); c.font = hm_font(pct)
        c.alignment = align('center','center'); c.border = bdr()
        ws.row_dimensions[i].height = 18

    # ── SHEET 6: DMA SUMMARY ───────────────────────────────────────────────────
    ws = wb.create_sheet('DMA Summary')
    ws.freeze_panes = 'C4'
    set_title(ws, 'FullSpeed Automotive — Service Penetration by DMA  |  Store count & % offering each category per market', 2+n_cats)

    set_left_hdr(ws, 2, 1, 'DMA Market', width=28)
    set_left_hdr(ws, 2, 2, 'Stores', width=9)
    for col, cat in enumerate(L1_CATS, 3):
        set_cat_hdr(ws, 2, col, cat, rotation=60)
    ws.row_dimensions[2].height = 80

    ws.merge_cells('A3:B3')
    c = ws.cell(row=3, column=1, value='Stores offering service →')
    c.font = font(bold=True, color=DGREY, size=9, italic=True)
    c.fill = fill(MGREY); c.alignment = align('right','center'); c.border = bdr()
    for col, cat in enumerate(L1_CATS, 3):
        set_count_cell(ws, 3, col, len(l1_stats[cat]['stores']), CAT_COLORS.get(cat,'6B7280'))
    ws.row_dimensions[3].height = 16

    for i, dma in enumerate(DMAS, 4):
        bg = LGREY if i % 2 == 0 else WHITE
        n = len(dma_stores[dma])
        c = ws.cell(row=i, column=1, value=dma)
        c.font = font(bold=True, size=10, color=BLACK); c.fill = fill(bg)
        c.alignment = align('left','center'); c.border = bdr()
        c = ws.cell(row=i, column=2, value=n)
        c.font = font(size=10, color=BLACK); c.fill = fill(bg)
        c.alignment = align('center','center'); c.border = bdr()
        for col, cat in enumerate(L1_CATS, 3):
            os_ = len(dma_l1[dma][cat]); pct = os_ / n if n else 0
            c = ws.cell(row=i, column=col, value=hm_val(pct))
            c.fill = hm_fill(pct); c.font = hm_font(pct)
            c.alignment = align('center','center'); c.border = bdr()
        ws.row_dimensions[i].height = 15

    ws.auto_filter.ref = f'A2:{get_column_letter(2+n_cats)}{3+len(DMAS)}'

    # ── SHEET 7: HIERARCHY REFERENCE ──────────────────────────────────────────
    ws = wb.create_sheet('Hierarchy Reference')
    ws.freeze_panes = 'A3'
    set_title(ws, 'Product Hierarchy — L1 Categories & L2 Subcategories  |  Coverage statistics  |  v2', 6)

    for col, (hdr, w) in enumerate([('L1 Category',28),('L2 Subcategory',30),('Total Transactions',16),
                                     ('Stores Offering',14),('% of All Stores',14),('Change Notes',42)], 1):
        set_left_hdr(ws, 2, col, hdr, width=w)
    ws.row_dimensions[2].height = 18

    rn = 3
    for l1 in L1_CATS:
        color = CAT_COLORS.get(l1,'6B7280')
        l2s = AUTHORITATIVE_L2[l1]
        l1_s = len(l1_stats[l1]['stores']); l1_tx = l1_stats[l1]['tx']
        for j, l2 in enumerate(l2s):
            bg = LGREY if rn % 2 == 0 else WHITE
            stats = l2_stats.get((l1,l2), {'stores':set(),'tx':0})
            c = ws.cell(row=rn, column=1, value=l1 if j==0 else '')
            c.font = font(bold=True, color=WHITE, size=11) if j==0 else font(size=10, color=BLACK)
            c.fill = fill(color) if j==0 else fill(bg)
            c.alignment = align('left','center'); c.border = bdr()
            c = ws.cell(row=rn, column=2, value=l2)
            c.font = font(size=10, color=BLACK); c.fill = fill(bg)
            c.alignment = align('left','center'); c.border = bdr()
            c = ws.cell(row=rn, column=3, value=stats['tx'])
            c.font = font(size=10); c.fill = fill(bg); c.number_format = '#,##0'
            c.alignment = align('center','center'); c.border = bdr()
            cnt = len(stats['stores'])
            c = ws.cell(row=rn, column=4, value=cnt)
            c.font = font(size=10); c.fill = fill(bg)
            c.alignment = align('center','center'); c.border = bdr()
            c = ws.cell(row=rn, column=5, value=cnt/n_stores if n_stores else 0)
            c.font = font(size=10); c.fill = fill(bg); c.number_format = '0%'
            c.alignment = align('center','center'); c.border = bdr()
            ws.cell(row=rn, column=6, value='').fill = fill(bg)
            ws.cell(row=rn, column=6).border = bdr()
            rn += 1
        # L1 total row
        c = ws.cell(row=rn, column=1, value=f'TOTAL — {l1}')
        c.font = font(bold=True, color=WHITE, size=11); c.fill = fill(DARK)
        c.alignment = align('left','center'); c.border = bdr()
        for col, val, fmt in [(2,'',None),(3,l1_tx,'#,##0'),(4,l1_s,None),(5,l1_s/n_stores if n_stores else 0,'0%'),(6,'',None)]:
            c = ws.cell(row=rn, column=col, value=val)
            c.font = font(bold=True, color=WHITE, size=11); c.fill = fill(DARK)
            c.alignment = align('center','center'); c.border = bdr()
            if fmt: c.number_format = fmt
        rn += 1

    # ── SHEET 8: RAW DATA ──────────────────────────────────────────────────────
    raw_rows = []
    with open(input_path, newline='', encoding='utf-8-sig') as f:
        raw_rows = list(csv.DictReader(f))

    ws = wb.create_sheet('Raw Data')
    ws.freeze_panes = 'A3'
    set_title(ws, f'Raw Data — All Store × Service Combinations  |  {len(raw_rows):,} rows  |  {date_window}', 12)

    RCOLS = [('STORE_NUMBER',9),('STORE_NAME',30),('BRAND',22),('CITY',14),('STATE',6),('DMA_NAME',26),
             ('STORE_MODEL',10),('PROPOSED_MODEL',13),('L1_CATEGORY',22),('L2_SUBCATEGORY',24),
             ('SERVICE_NAME',34),('TRANSACTION_COUNT',14),('IS_OIL_TRANSACTION',14)]
    for col, (hdr, w) in enumerate(RCOLS, 1):
        set_left_hdr(ws, 2, col, hdr, width=w)
    ws.row_dimensions[2].height = 18

    for i, r in enumerate(raw_rows, 3):
        bg = LGREY if i % 2 == 0 else WHITE
        for col, (key, _) in enumerate(RCOLS, 1):
            val = r.get(key, '')
            if key == 'TRANSACTION_COUNT' and val:
                try: val = int(val)
                except: pass
            c = ws.cell(row=i, column=col, value=val)
            c.font = font(size=9, color=BLACK); c.fill = fill(bg)
            c.alignment = align('left','center'); c.border = bdr()
        ws.row_dimensions[i].height = 14

    ws.auto_filter.ref = f'A2:{get_column_letter(len(RCOLS))}{2+len(raw_rows)}'

    # ── SAVE ───────────────────────────────────────────────────────────────────
    wb.save(output_path)
    print(f"  Written: {output_path}")


# ── CLI ────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input',           required=True)
    parser.add_argument('--output',          required=True)
    parser.add_argument('--region-district', default='scripts/region_district.csv',
                        help='Path to region_district.csv (optional, adds REGION/DISTRICT columns)')
    args = parser.parse_args()
    build(args.input, args.output, args.region_district)

if __name__ == '__main__':
    main()
