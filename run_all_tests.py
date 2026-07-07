"""
Comprehensive test suite for the Godrej Market Share Model.
Run: python3 run_all_tests.py
Every test must pass before the model is considered release-ready.
"""
import sys
import pandas as pd
import numpy as np
import openpyxl

from cleaner import (
    process_all_files, process_workbook, get_sheet_overview,
    parse_region_sheet_name, classify_flag, classify_company,
)
from calculator import (
    add_calculations, build_factor_lookup, DEFAULT_INDIVIDUAL_FACTOR_UR,
    DEFAULT_ZONE_MAPPING, STATE_TO_ZONE, build_company_summary,
    compute_variance, smart_guess_companies, find_potential_duplicates,
    get_period_list,
)
from exporter import export_to_excel
from default_brand_mapping import DEFAULT_BRAND_MAPPING
from default_company_mapping import DEFAULT_COMPANY_MAPPING

RAW = '/mnt/user-data/uploads/Rajasthan_SHC.xlsx'
TRUTH = '/mnt/user-data/uploads/HC_KWP_Working_CompanyWise_V3_18062026__2_.xlsx'

PASS, FAIL = 0, 0
def check(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  PASS  {name}")
    else:
        FAIL += 1
        print(f"  FAIL  {name}  {detail}")


def load_truth(state='Rajasthan', fmt='SHC'):
    wb = openpyxl.load_workbook(TRUTH, data_only=True, read_only=True)
    ws = wb['Sheet1']
    headers = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0] == state and r[3] == fmt]
    df = pd.DataFrame(rows, columns=headers)
    df['UR'] = df['Urban/Rural'].map({'Urban': 'U', 'Rural': 'R', 'U+R': 'U+R'})
    return df

QMAP = {
    '2023 Apr - 2023 Jun': "01. AMJ'23", '2023 Jul - 2023 Sep': "02. JAS'23",
    '2023 Oct - 2023 Dec': "03. OND'23", '2024 Jan - 2024 Mar': "04. JFM'24",
    '2024 Apr - 2024 Jun': "05. AMJ'24", '2024 Jul - 2024 Sep': "06. JAS'24",
    '2024 Oct - 2024 Dec': "07. OND'24", '2025 Jan - 2025 Mar': "08. JFM'25",
    '2025 Apr - 2025 Jun': "09. AMJ'25", '2025 Jul - 2025 Sep': "10. JAS'25",
    '2025 Oct - 2025 Dec': "11. OND'25", '2026 Jan - 2026 Mar': "12. JFM'26",
    '2023 Apr - 2024 Mar': 'MAT Mar24', '2024 Apr - 2025 Mar': 'MAT Mar25',
    '2025 Apr - 2026 Mar': 'MAT Mar26',
}

print("=" * 70)
print("TEST GROUP 1: Sheet name parsing")
print("=" * 70)
cases = [
    ('Rajasthan (U+R)', ('Rajasthan', 'U+R', False)),
    ('Rajasthan (U)', ('Rajasthan', 'U', False)),
    ('Rajasthan (R)', ('Rajasthan', 'R', False)),
    ('North (U+R)', ('North', 'U+R', True)),
    ('South (R)', ('South', 'R', True)),
    ('East (U)', ('East', 'U', True)),
    ('West (U+R)', ('West', 'U+R', True)),
    ('GCPL East (U+R)', ('GCPL East', 'U+R', True)),
    ('GCPL West (U)', ('GCPL West', 'U', True)),
    ('North Zone(U)', ('North', 'U', True)),
    ('All India Urban', ('All India', 'U', False)),
    ('All India Rural', ('All India', 'R', False)),
    ('All India U+R', ('All India', 'U+R', False)),
    ('Andhra Pradesh excl Tela (U+R)', ('Andhra Pradesh excl Tela', 'U+R', False)),
    ('Madhya Pradesh excl Chha (R)', ('Madhya Pradesh excl Chha', 'R', False)),
]
for name, expected in cases:
    got = parse_region_sheet_name(name)
    check(f"parse '{name}'", got == expected, f"got {got}, expected {expected}")

print("=" * 70)
print("TEST GROUP 2: Flag / Company classification")
print("=" * 70)
check("Category row detected", classify_flag('[HCEXL] ANY HAIR COL.SHAMPOO') == 'Category')
check("Brand row detected", classify_flag('[HCEXL] GODREJ SELFIE HAIR COLOUR SHAMPOO', DEFAULT_BRAND_MAPPING) == 'Brand')
check("Others row detected", classify_flag("[HCEXL] OTH.BRDED CLR'ANT/HENNA", DEFAULT_BRAND_MAPPING) == 'Others')
check("SKU row detected", classify_flag('Godrej Selfie Shampoo Sachet 15ml') == 'SKU')
check("Non-bracket HBP brand '5 Star' via mapping", classify_flag('5 Star', DEFAULT_BRAND_MAPPING) == 'Brand')
check("Sub-brand via mapping", classify_flag('[HCEXL] GARNIER BLACK NATURALS', DEFAULT_BRAND_MAPPING) == 'Sub-brand')
check("Category company = Category Total", classify_company('[HCEXL] ANY X', 'Category') == 'Category Total')
check("GCPL company lookup", classify_company('[HCEXL] GODREJ SELFIE HAIR COLOUR SHAMPOO', 'Brand', DEFAULT_COMPANY_MAPPING) == 'GODREJ CONSUMER PRODS')
check("Unmapped brand -> Others / Unmapped", classify_company('[HCEXL] SOMETHING BRAND NEW', 'Brand', DEFAULT_COMPANY_MAPPING) == 'Others / Unmapped')

print("=" * 70)
print("TEST GROUP 3: Sheet overview (Region Selection defaults)")
print("=" * 70)
f = open(RAW, 'rb')
ov = get_sheet_overview(f)
check("3 sheets found in Rajasthan file", len(ov) == 3)
check("All Rajasthan sheets default-included", all(o['Include'] for o in ov))
check("No sheet misdetected as zone", all(not o['Is_Zone'] for o in ov))

print("=" * 70)
print("TEST GROUP 4: Full pipeline vs GROUND TRUTH (the core correctness test)")
print("=" * 70)
f2 = open(RAW, 'rb')
master_df, sku_df = process_all_files(
    {'SHC': f2},
    flag_overrides=dict(DEFAULT_BRAND_MAPPING),
    company_overrides=dict(DEFAULT_COMPANY_MAPPING),
    included_sheets_by_format={'SHC': {'Rajasthan (U+R)', 'Rajasthan (U)', 'Rajasthan (R)'}},
)
check("Master rows extracted", len(master_df) == 270, f"got {len(master_df)}")
check("SKU rows extracted for variance", len(sku_df) == 1245, f"got {len(sku_df)}")

factor_lookup = build_factor_lookup(dict(DEFAULT_INDIVIDUAL_FACTOR_UR))
zone_mapping = {'North': ['Rajasthan']}  # what the app would auto-build here
result, missing, unmapped = add_calculations(master_df, factor_lookup=factor_lookup, zone_mapping=zone_mapping)
check("No missing factor regions", len(missing) == 0, str(missing))
check("No unmapped zones", len(unmapped) == 0, str(unmapped))

truth = load_truth()
res = result[result['TG_Segment'] == 'TOTAL'].set_index(['Urban_Rural', 'Brand_SKU_Item']).sort_index()
truth_idx = truth.set_index(['UR', 'Brand_SKU Item'])

# 4a. Sales Derived exactness: every brand, every UR, ALL 15 periods
checked, max_diff, flag_mm = 0, 0.0, 0
for (ur, item), trow in truth_idx.iterrows():
    if (ur, item) not in res.index:
        continue
    rrow = res.loc[(ur, item)]
    if isinstance(rrow, pd.DataFrame):
        rrow = rrow.iloc[0]
    if rrow['Flag'] != trow['Flag']:
        flag_mm += 1
    for our_p, truth_p in QMAP.items():
        t = trow.get(f'Sales Value_{truth_p}')
        o = rrow.get(f'Sales Derived__{our_p}')
        if o is None or pd.isna(o) or t in (None, 0) or pd.isna(t):
            continue
        d = abs(o - t) / t * 100
        checked += 1
        if d > max_diff:
            max_diff = d
check(f"Sales Derived exact across ALL periods ({checked} datapoints)", max_diff < 1e-9, f"max diff {max_diff}%")
check("Zero flag mismatches", flag_mm == 0, f"{flag_mm} mismatches")

# 4b. Units Estd exactness
checked_u, max_diff_u = 0, 0.0
for (ur, item), trow in truth_idx.iterrows():
    if (ur, item) not in res.index:
        continue
    rrow = res.loc[(ur, item)]
    if isinstance(rrow, pd.DataFrame):
        rrow = rrow.iloc[0]
    for our_p, truth_p in QMAP.items():
        t = trow.get(f'Units_{truth_p}')
        o = rrow.get(f'Units Estd__{our_p}')
        if o is None or pd.isna(o) or t in (None, 0) or pd.isna(t):
            continue
        d = abs(o - t) / t * 100
        checked_u += 1
        if d > max_diff_u:
            max_diff_u = d
check(f"Units Estd exact across ALL periods ({checked_u} datapoints)", max_diff_u < 1e-9, f"max diff {max_diff_u}%")

# 4c. Value MS% exactness
checked_m, max_diff_m = 0, 0.0
for (ur, item), trow in truth_idx.iterrows():
    if (ur, item) not in res.index:
        continue
    rrow = res.loc[(ur, item)]
    if isinstance(rrow, pd.DataFrame):
        rrow = rrow.iloc[0]
    for our_p, truth_p in QMAP.items():
        t = trow.get(f'Value MS%_{truth_p}')
        o = rrow.get(f'Value MS%__{our_p}')
        if o is None or pd.isna(o) or t is None or pd.isna(t):
            continue
        d = abs(o - t)
        checked_m += 1
        if d > max_diff_m:
            max_diff_m = d
check(f"Value MS% exact across ALL periods ({checked_m} datapoints)", max_diff_m < 1e-9, f"max abs diff {max_diff_m}")

# 4d. Units MS% exactness
checked_um, max_diff_um = 0, 0.0
for (ur, item), trow in truth_idx.iterrows():
    if (ur, item) not in res.index:
        continue
    rrow = res.loc[(ur, item)]
    if isinstance(rrow, pd.DataFrame):
        rrow = rrow.iloc[0]
    for our_p, truth_p in QMAP.items():
        t = trow.get(f'Units MS%_{truth_p}')
        o = rrow.get(f'Units MS%__{our_p}')
        if o is None or pd.isna(o) or t is None or pd.isna(t):
            continue
        d = abs(o - t)
        checked_um += 1
        if d > max_diff_um:
            max_diff_um = d
check(f"Units MS% exact across ALL periods ({checked_um} datapoints)", max_diff_um < 1e-9, f"max abs diff {max_diff_um}")

print("=" * 70)
print("TEST GROUP 5: Company summary vs ground truth (GCPL and every competitor)")
print("=" * 70)
cs = build_company_summary(result)
truth_ur = truth[truth['UR'] == 'U+R']
company_col = 'Company Flag '  # note: trailing space exists in the source header
truth_company_sums = {}
for _, r in truth_ur.iterrows():
    comp = r[company_col]
    if comp in ('Category Total',):
        continue
    for our_p, truth_p in QMAP.items():
        v = r.get(f'Sales Value_{truth_p}')
        if v is None or pd.isna(v):
            continue
        truth_company_sums.setdefault((comp, our_p), 0.0)
        truth_company_sums[(comp, our_p)] += v

cs_ur = cs[(cs['State_Zone'] == 'Rajasthan') & (cs['Urban_Rural'] == 'U+R') & (cs['TG_Segment'] == 'TOTAL')]
checked_c, max_diff_c = 0, 0.0
for (comp, our_p), t in truth_company_sums.items():
    row = cs_ur[cs_ur['Company'] == comp]
    if len(row) == 0 or t == 0:
        continue
    o = row[f'Sales Derived__{our_p}'].values[0]
    if pd.isna(o):
        continue
    d = abs(o - t) / t * 100
    checked_c += 1
    if d > max_diff_c:
        max_diff_c = d
check(f"Company-level Sales exact for every company/period ({checked_c} datapoints)", max_diff_c < 1e-9, f"max diff {max_diff_c}%")

print("=" * 70)
print("TEST GROUP 6: U+R = U + R identity (structural invariant)")
print("=" * 70)
periods = get_period_list(result)
tot = result[(result['TG_Segment'] == 'TOTAL') & (result['State_Zone'] == 'Rajasthan')]
u = tot[tot['Urban_Rural'] == 'U'].set_index('Brand_SKU_Item')
r_ = tot[tot['Urban_Rural'] == 'R'].set_index('Brand_SKU_Item')
ur = tot[tot['Urban_Rural'] == 'U+R'].set_index('Brand_SKU_Item')
bad = 0
for item in ur.index:
    if item not in u.index or item not in r_.index:
        continue
    for p in ['2025 Apr - 2026 Mar']:
        s_ur = ur.loc[item, f'Sales Derived__{p}']
        s_sum = u.loc[item, f'Sales Derived__{p}'] + r_.loc[item, f'Sales Derived__{p}']
        if pd.notna(s_ur) and abs(s_ur - s_sum) > 1e-6:
            bad += 1
check("U+R equals U + R for every brand", bad == 0, f"{bad} violations")

print("=" * 70)
print("TEST GROUP 7: Zone synthesis correctness")
print("=" * 70)
north = result[(result['State_Zone'] == 'North') & (result['Urban_Rural'] == 'U+R') & (result['TG_Segment'] == 'TOTAL') & (result['Flag'] == 'Category')]
raj = result[(result['State_Zone'] == 'Rajasthan') & (result['Urban_Rural'] == 'U+R') & (result['TG_Segment'] == 'TOTAL') & (result['Flag'] == 'Category')]
check("Zone North exists in output", len(north) == 1)
check("Zone North == sum of member states (Rajasthan only here)",
      abs(north['Sales Derived__2025 Apr - 2026 Mar'].values[0] - raj['Sales Derived__2025 Apr - 2026 Mar'].values[0]) < 1e-6)

# Zone with garbage sheet data present in input must be ignored
fake_zone = master_df[master_df['Urban_Rural'] == 'U'].copy()
fake_zone['State_Zone'] = 'North'
fake_zone['Is_Zone'] = True
for c in fake_zone.columns:
    if c.startswith('HH__') or c.startswith('Val__') or c.startswith('Avg NOP__'):
        fake_zone[c] = 999999
df_with_zone_sheet = pd.concat([master_df, fake_zone], ignore_index=True)
result2, _, _ = add_calculations(df_with_zone_sheet, factor_lookup=factor_lookup, zone_mapping={'North': ['Rajasthan']})
north2 = result2[(result2['State_Zone'] == 'North') & (result2['Urban_Rural'] == 'U') & (result2['TG_Segment'] == 'TOTAL') & (result2['Flag'] == 'Category')]
raj_u = result2[(result2['State_Zone'] == 'Rajasthan') & (result2['Urban_Rural'] == 'U') & (result2['TG_Segment'] == 'TOTAL') & (result2['Flag'] == 'Category')]
check("Zone's own sheet data ignored (synthesized from states only)",
      len(north2) == 1 and abs(north2['Sales Derived__2025 Apr - 2026 Mar'].values[0] - raj_u['Sales Derived__2025 Apr - 2026 Mar'].values[0]) < 1e-6)

# Zone with no member states present -> reported, not guessed
result3, _, unmapped3 = add_calculations(master_df, factor_lookup=factor_lookup, zone_mapping={'South': ['Kerala']})
check("Zone with absent member states reported as unmapped", 'South' in unmapped3)
check("Unmapped zone produces no rows", len(result3[result3['State_Zone'] == 'South']) == 0)

print("=" * 70)
print("TEST GROUP 8: Variance calculation")
print("=" * 70)
var_df = compute_variance(master_df, sku_df)
check("Variance rows produced", len(var_df) > 0)
sel = var_df[(var_df['State_Zone'] == 'Rajasthan') & (var_df['Urban_Rural'] == 'U+R') & (var_df['TG_Segment'] == 'TOTAL') & (var_df['Brand_SKU_Item'].str.contains('SELFIE'))]
check("Selfie variance row exists", len(sel) == 1)
if len(sel) == 1:
    bu = sel['Brand_Units__2025 Apr - 2026 Mar'].values[0]
    su = sel['SKU_Sum_Units__2025 Apr - 2026 Mar'].values[0]
    # manual: brand HH*NOP = 288.146*1.811
    manual_bu = 288.146 * 1.811
    check("Brand units match manual HH*NOP", abs(bu - manual_bu) < 1e-6, f"{bu} vs {manual_bu}")
    check("Variance = Brand - SKU sum", abs(sel['Variance__2025 Apr - 2026 Mar'].values[0] - (bu - su)) < 1e-9)

print("=" * 70)
print("TEST GROUP 9: Smart guess and duplicates")
print("=" * 70)
g = smart_guess_companies(['[HCEXL] GODREJ BRAND NEW THING'], DEFAULT_COMPANY_MAPPING)
check("Godrej keyword guess", g['[HCEXL] GODREJ BRAND NEW THING'][0] == 'GODREJ CONSUMER PRODS')
g2 = smart_guess_companies(['[HCEXL] TOTALLY UNKNOWN THING QQQ'], DEFAULT_COMPANY_MAPPING)
check("Unknown brand falls to Others / Unmapped", g2['[HCEXL] TOTALLY UNKNOWN THING QQQ'][0] == 'Others / Unmapped')
g3 = smart_guess_companies(['[HCEXL] GODREJ SELFIE HAIR COLOR SHAMPO'], DEFAULT_COMPANY_MAPPING)
check("Near-duplicate typo matched to GCPL", g3['[HCEXL] GODREJ SELFIE HAIR COLOR SHAMPO'][0] == 'GODREJ CONSUMER PRODS')
dupes = find_potential_duplicates(list(DEFAULT_BRAND_MAPPING.keys()))
check("Duplicate scan runs and finds known near-dupes", len(dupes) >= 1)

print("=" * 70)
print("TEST GROUP 10: Exporter (all sheets, openable workbook)")
print("=" * 70)
out = export_to_excel(result, dict(DEFAULT_INDIVIDUAL_FACTOR_UR), missing,
                      company_summary_df=cs, zone_mapping=zone_mapping,
                      unmapped_zones=unmapped, variance_df=var_df)
data = out.getvalue()
check("Export non-empty", len(data) > 100000)
import io
wb_out = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
expected_sheets = {'Master_Clean', 'Company_Summary', 'Variance', 'Brand_Mapping', 'Company_Mapping', 'Individual_Factor', 'Zone_Mapping'}
check("All expected sheets present", expected_sheets.issubset(set(wb_out.sheetnames)), str(wb_out.sheetnames))
ws_mc = wb_out['Master_Clean']
check("Master_Clean has data rows", ws_mc.max_row > 100)

print("=" * 70)
print("TEST GROUP 11: Mapping data integrity")
print("=" * 70)
check("Brand mapping 180 items", len(DEFAULT_BRAND_MAPPING) == 180)
check("Company mapping 180 items", len(DEFAULT_COMPANY_MAPPING) == 180)
check("Same keys in both mappings", set(DEFAULT_BRAND_MAPPING) == set(DEFAULT_COMPANY_MAPPING))
# Re-extract from source and diff
wb_t = openpyxl.load_workbook(TRUTH, data_only=True, read_only=True)
ws_t = wb_t['Sheet1']
fresh_flag, fresh_comp = {}, {}
for row in ws_t.iter_rows(min_row=2, values_only=True):
    fresh_flag[row[5]] = row[4]
    fresh_comp[row[5]] = row[96]
check("Brand mapping matches source byte-for-byte", fresh_flag == DEFAULT_BRAND_MAPPING)
check("Company mapping matches source byte-for-byte", fresh_comp == DEFAULT_COMPANY_MAPPING)

print("=" * 70)
print("TEST GROUP 12: App page logic simulation (session-state flows)")
print("=" * 70)
# Simulate the exact dict-building the app does before running the pipeline
mapping_df = pd.DataFrame([{'Brand_SKU_Item': k, 'Flag': v, 'Status': 'Verified'} for k, v in DEFAULT_BRAND_MAPPING.items()])
company_mapping_df = pd.DataFrame([{'Brand_SKU_Item': k, 'Company': v, 'Status': 'Verified'} for k, v in DEFAULT_COMPANY_MAPPING.items()])
fo = dict(zip(mapping_df['Brand_SKU_Item'], mapping_df['Flag']))
co = dict(zip(company_mapping_df['Brand_SKU_Item'], company_mapping_df['Company']))
check("flag_overrides built from df with Status col", fo == DEFAULT_BRAND_MAPPING)
check("company_overrides built from df with Status col", co == DEFAULT_COMPANY_MAPPING)

# Zone auto-build logic from app
sheet_sel = pd.DataFrame([{**o, 'Format': 'SHC'} for o in get_sheet_overview(open(RAW, 'rb'))])
included_states = set(sheet_sel[(sheet_sel['Include']) & (~sheet_sel['Is_Zone'])]['State_Zone'].unique())
zrows = []
for state in sorted(included_states):
    z = STATE_TO_ZONE.get(state)
    if z:
        zrows.append({'Zone': z, 'Member_State': state})
check("App zone auto-build produces North->Rajasthan", zrows == [{'Zone': 'North', 'Member_State': 'Rajasthan'}])

# Factor dict building as the app does
factor_df = pd.DataFrame([{'State_Zone': s, 'Individual_Factor_U': v.get('U'), 'Individual_Factor_R': v.get('R')} for s, v in DEFAULT_INDIVIDUAL_FACTOR_UR.items()])
fud = {}
for _, row in factor_df.iterrows():
    fud[row['State_Zone']] = {'U': row.get('Individual_Factor_U'), 'R': row.get('Individual_Factor_R')}
fl2 = build_factor_lookup(fud)
check("Factor lookup from app's df == direct lookup", fl2 == build_factor_lookup(dict(DEFAULT_INDIVIDUAL_FACTOR_UR)))
check("Rajasthan U factor exact", fl2[('Rajasthan', 'U')] == 1.32)
check("Rajasthan R factor exact", fl2[('Rajasthan', 'R')] == 1.24)

print("=" * 70)
print("TEST GROUP 13: Robustness edge cases")
print("=" * 70)
# Empty sku_df to variance
empty_var = compute_variance(master_df, pd.DataFrame())
check("Variance with empty SKU df doesn't crash", len(empty_var) > 0)
# Included sheets filter actually filters
f3 = open(RAW, 'rb')
m_only_ur, _ = process_workbook(f3, 'SHC', included_sheet_names={'Rajasthan (U+R)'})
check("Sheet whitelist honored", set(m_only_ur['Urban_Rural'].unique()) == {'U+R'})
# get_sheet_overview twice on same handle (seek safety)
f4 = open(RAW, 'rb')
ov1 = get_sheet_overview(f4)
ov2 = get_sheet_overview(f4)
check("Sheet overview re-readable on same handle", len(ov1) == len(ov2) == 3)
# process after overview on same handle (the app does exactly this)
m_after, _ = process_workbook(f4, 'SHC')
check("Workbook processable after overview on same handle", len(m_after) == 270, f"got {len(m_after)}")

print("=" * 70)
print("TEST GROUP 14: All India auto-rollup (never its own factor/sheet)")
print("=" * 70)
f5 = open(RAW, 'rb')
m5, s5 = process_all_files({'SHC': f5}, flag_overrides=dict(DEFAULT_BRAND_MAPPING), company_overrides=dict(DEFAULT_COMPANY_MAPPING))
result5, _, _ = add_calculations(m5, factor_lookup=factor_lookup, zone_mapping={})
ai5 = result5[(result5['State_Zone'] == 'All India') & (result5['Urban_Rural'] == 'U+R') & (result5['TG_Segment'] == 'TOTAL') & (result5['Flag'] == 'Category')]
raj5 = result5[(result5['State_Zone'] == 'Rajasthan') & (result5['Urban_Rural'] == 'U+R') & (result5['TG_Segment'] == 'TOTAL') & (result5['Flag'] == 'Category')]
check("All India auto-synthesized (no manual mapping needed)", len(ai5) == 1)
check("All India == sum of real states present (Rajasthan only here)",
      abs(ai5['Sales Derived__2025 Apr - 2026 Mar'].values[0] - raj5['Sales Derived__2025 Apr - 2026 Mar'].values[0]) < 1e-6)

# Plant a real All India sheet (with the OLD unverified 1.50/1.80 factor) alongside
# Rajasthan and confirm its own data/factor are still fully ignored.
fake_ai = m5[m5['Urban_Rural'] == 'U'].copy()
fake_ai['State_Zone'] = 'All India'
fake_ai['Is_Zone'] = False
for c in fake_ai.columns:
    if c.startswith('HH__') or c.startswith('Val__') or c.startswith('Avg NOP__'):
        fake_ai[c] = 999999
m5_with_ai_sheet = pd.concat([m5, fake_ai], ignore_index=True)
result6, _, _ = add_calculations(m5_with_ai_sheet, factor_lookup=factor_lookup, zone_mapping={})
ai6 = result6[(result6['State_Zone'] == 'All India') & (result6['Urban_Rural'] == 'U+R') & (result6['TG_Segment'] == 'TOTAL') & (result6['Flag'] == 'Category')]
check("All India's own planted sheet+factor fully ignored, still == Rajasthan",
      len(ai6) == 1 and abs(ai6['Sales Derived__2025 Apr - 2026 Mar'].values[0] - raj5['Sales Derived__2025 Apr - 2026 Mar'].values[0]) < 1e-6)

print("=" * 70)
print("TEST GROUP 15: Grammage/SU blank-inconsistency bug (regression)")
print("=" * 70)
# Two states with the SAME real values but DIFFERENT blank conventions for
# Grammage/SU (None vs ''), which is common across different real workbooks.
# Before the fix, this silently split zone/All India rollups into separate
# under-counted groups instead of combining them.
data = {
    'Format': ['SHC'] * 4,
    'State_Zone': ['StateA', 'StateA', 'StateB', 'StateB'],
    'Is_Zone': [False] * 4,
    'Urban_Rural': ['U', 'R', 'U', 'R'],
    'TG_Segment': ['TOTAL'] * 4,
    'Flag': ['Brand'] * 4,
    'Company': ['GCPL'] * 4,
    'Grammage': [None, None, '', ''],
    'SU': [None, None, None, None],
    'Brand_SKU_Item': ['Test Brand'] * 4,
    'HH__P1': [100.0, 100.0, 100.0, 100.0],
    'Val__P1': [10.0, 10.0, 10.0, 10.0],
    'Avg NOP__P1': [2.0, 2.0, 2.0, 2.0],
}
df_blank_test = pd.DataFrame(data)
fl_test = build_factor_lookup({'StateA': {'U': 1.0, 'R': 1.0}, 'StateB': {'U': 1.0, 'R': 1.0}})
result_blank, _, _ = add_calculations(df_blank_test, factor_lookup=fl_test, zone_mapping={})
ai_blank = result_blank[result_blank['State_Zone'] == 'All India']
ai_u_total = ai_blank[ai_blank['Urban_Rural'] == 'U']['Sales Derived__P1'].sum()
ai_r_total = ai_blank[ai_blank['Urban_Rural'] == 'R']['Sales Derived__P1'].sum()
check("All India combines states with inconsistent blank Grammage (U)", ai_u_total == 20000000.0, f"got {ai_u_total}")
check("All India combines states with inconsistent blank Grammage (R)", ai_r_total == 20000000.0, f"got {ai_r_total}")
check("No row-count inflation from blank-value grouping mismatch", len(ai_blank[ai_blank['Urban_Rural'] == 'U']) == 1)

print("=" * 70)
print(f"RESULT: {PASS} passed, {FAIL} failed")
print("=" * 70)
sys.exit(1 if FAIL else 0)
