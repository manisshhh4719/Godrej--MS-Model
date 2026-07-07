"""
Godrej Market Share Model - Data Cleaning Module
Reads raw Kantar 68-sheet Excel files (one per Format/Category) and produces
a single clean long-form dataframe: one row per Region x TG_Segment x Brand x Period.

Key facts validated against ground truth before writing this code:
- Column A (TG) in the raw sheet is filled only on real data rows, with the
  segment name (TOTAL, SEC A, SEC B, SEC C, SEC D/E). Blank / header / Universe
  rows always have column A empty. This lets us detect segments and skip
  junk rows without guessing row numbers.
- The metric header row (HH, Vol, Val, Avg Cons, Avg FOP, Avg POC, Avg NOP) and
  the period header row are identical everywhere in the sheet, so we only need
  to read them once (row 4 and row 5) to build column names.
- The "[X] ANY ..." row appearing first in every segment is the Category total.
  Any other bracketed row ("[HCEXL] ..." or "[COLOR] ...") is a Brand row.
  Un-bracketed rows are individual SKUs and are NOT needed for this model
  (Sagar's brand-level share model only needs Category and Brand rows; Brand
  totals from Kantar already include unlisted small SKUs, so summing children
  would double count / under count).
"""

import openpyxl
import pandas as pd
import re
from default_brand_mapping import DEFAULT_BRAND_MAPPING
from default_company_mapping import DEFAULT_COMPANY_MAPPING

TG_SEGMENTS = {"TOTAL", "SEC A", "SEC B", "SEC C", "SEC D/E"}

# Sheets to always exclude (per Sagar's rule: keep "Andhra Pradesh excl Tela",
# drop any combined "AP+Tel" sheet to avoid double counting)
EXCLUDE_SHEET_KEYWORDS = ["AP+Tel", "AP + Tel", "AP&Tel"]

# Known zone identifiers. A sheet base-name is treated as a Zone (not a
# State) if it matches one of these, or contains "zone"/"gcpl". This matters
# because real zone sheets are often just named "North", "GCPL East" etc,
# with no literal word "Zone" in them - a name-pattern-only check would
# misclassify them as ordinary states.
KNOWN_ZONE_NAMES = {"north", "south", "east", "west"}


def parse_region_sheet_name(sheet_name):
    """
    Split a sheet name into (State/Zone name, Urban_Rural, is_zone).
    Handles patterns like:
      'Rajasthan (U+R)', 'Rajasthan (U)', 'Rajasthan (R)'
      'All India Urban', 'All India Rural', 'All India U+R'
      'North (U)', 'North Zone(U)', 'GCPL East (U+R)'
      'Andhra Pradesh excl Tela (U+R)'
    """
    name = sheet_name.strip()

    # All India special case (no parentheses, word Urban/Rural/U+R at the end)
    m = re.match(r"^All India\s*(Urban|Rural|U\+R)$", name, re.IGNORECASE)
    if m:
        ur_raw = m.group(1)
        ur = {"urban": "U", "rural": "R", "u+r": "U+R"}[ur_raw.lower()]
        return "All India", ur, False

    # Generic pattern: 'BaseName (U+R)' / 'BaseName(U)' etc - covers both
    # states ('Rajasthan') and zones ('North', 'GCPL East', 'North Zone').
    m = re.match(r"^(.*?)\s*\((U\+R|U|R)\)$", name)
    if m:
        base = m.group(1).strip()
        ur = m.group(2).upper()
        base_clean = re.sub(r"\s*Zone\s*$", "", base, flags=re.IGNORECASE).strip()
        is_zone = (
            base_clean.lower() in KNOWN_ZONE_NAMES
            or "zone" in base.lower()
            or "gcpl" in base.lower()
        )
        return base_clean, ur, is_zone

    # Fallback: couldn't parse, return as-is
    return name, "UNKNOWN", False


def is_excluded_sheet(sheet_name, exclude_keywords=None):
    keywords = exclude_keywords if exclude_keywords is not None else EXCLUDE_SHEET_KEYWORDS
    return any(kw.lower() in sheet_name.lower() for kw in keywords if kw)


def get_sheet_overview(file_path_or_bytes, exclude_keywords=None):
    """
    Lists every sheet in a raw workbook with its classification, without
    reading any row data (cheap - just sheet names). Used by the Region
    Selection page so the person can see and choose exactly which sheets
    get processed, instead of a hidden keyword rule.

    Default suggested inclusion:
      - Plain State sheets -> included
      - Zone sheets (North, GCPL East, etc) -> excluded (zones are
        calculated by summing states via Zone Mapping, not from their own
        sheet - see calculator.add_calculations)
      - AP+Tel-style combined sheets -> excluded (avoids double counting
        with Andhra Pradesh excl Tela + Telangana reported separately)
    """
    if hasattr(file_path_or_bytes, "seek"):
        file_path_or_bytes.seek(0)
    wb = openpyxl.load_workbook(file_path_or_bytes, read_only=True)
    overview = []
    for sheet_name in wb.sheetnames:
        state_or_zone, urban_rural, is_zone = parse_region_sheet_name(sheet_name)
        is_ap_tel = is_excluded_sheet(sheet_name, exclude_keywords=exclude_keywords)
        default_include = (not is_zone) and (not is_ap_tel)
        overview.append({
            "Sheet_Name": sheet_name,
            "State_Zone": state_or_zone,
            "Urban_Rural": urban_rural,
            "Is_Zone": is_zone,
            "Is_AP_Tel": is_ap_tel,
            "Include": default_include,
        })
    wb.close()
    return overview


def build_column_map(rows, metric_header_row_idx=3, period_header_row_idx=4, id_cols=5):
    """
    Build a list of column names for columns starting at index `id_cols`.
    rows: list of row tuples (0-indexed, as returned by ws.iter_rows(values_only=True))
    Returns list of strings like 'HH__2023 Apr - 2023 Jun'.
    """
    metric_row = rows[metric_header_row_idx]
    period_row = rows[period_header_row_idx]

    col_names = []
    current_metric = None
    for i in range(id_cols, len(metric_row)):
        if metric_row[i] is not None:
            current_metric = str(metric_row[i]).strip()
        period = period_row[i]
        if period is None:
            col_names.append(None)
        else:
            col_names.append(f"{current_metric}__{str(period).strip()}")
    return col_names


def classify_flag(product_name, flag_overrides=None):
    """
    Returns one of: 'Category', 'Brand', 'Sub-brand', 'Others', 'SKU'

    Priority order:
      1. flag_overrides (the Brand Mapping table) - checked FIRST, regardless
         of bracket notation. This matters because not every format uses the
         same naming convention (e.g. HBP/Powder has at least one real Brand
         row, "5 Star", with no brackets at all - the mapping catches this,
         a pure bracket-based guess would not).
      2. '[X] ANY ...' pattern -> Category (one per segment, the format total)
      3. Any other bracketed row -> Others (if it contains "OTH.") else Brand
      4. Un-bracketed and not in the mapping -> SKU (excluded from output,
         since Sagar's brand-level share model doesn't need individual SKUs)

    flag_overrides: dict {Brand_SKU_Item (exact text): 'Category'|'Brand'|'Sub-brand'|'Others'}
    """
    name = str(product_name).strip()

    if flag_overrides and name in flag_overrides:
        return flag_overrides[name]

    if not name.startswith("["):
        return "SKU"

    after_bracket = re.sub(r"^\[[^\]]+\]\s*", "", name)
    if after_bracket.upper().startswith("ANY "):
        return "Category"

    if "OTH." in name.upper() or "OTHERS" in name.upper():
        return "Others"

    return "Brand"


def classify_company(product_name, flag, company_overrides=None):
    """
    Returns the parent Company for a Brand_SKU_Item, e.g. 'GODREJ CONSUMER PRODS'.
    Category rows get 'Category Total'. Others/unmapped items not found in the
    mapping get 'Others / Unmapped' rather than being silently blank, so they're
    still visible to review in the output.

    company_overrides: dict {Brand_SKU_Item (exact text): Company name}
    """
    if flag == "Category":
        return "Category Total"

    name = str(product_name).strip()
    if company_overrides and name in company_overrides:
        return company_overrides[name]

    return "Others / Unmapped"


def process_workbook(file_path_or_bytes, format_tag, flag_overrides=None, company_overrides=None, exclude_keywords=None, included_sheet_names=None, verbose=False):
    """
    Reads one raw Kantar workbook (all sheets = all regions) for a single
    Format/Category, and returns (df, sku_df):
      df     - Category/Brand/Sub-brand/Others rows (the main Master_Clean data)
      sku_df - individual SKU rows, each tagged with its Parent_Brand, kept
               separately (not part of Master_Clean) purely so a Variance
               check can compare a Brand's own reported total against the
               sum of its listed SKUs.

    flag_overrides (if given) are layered on top of DEFAULT_BRAND_MAPPING,
    i.e. user edits/uploads always win over the built-in default.

    included_sheet_names: if given (a set/list of exact sheet names), ONLY
    those sheets are processed - this is what the Region Selection page
    drives. If None, falls back to the old keyword-based exclude_keywords
    rule (kept for programmatic/test use).
    """
    combined_mapping = dict(DEFAULT_BRAND_MAPPING)
    if flag_overrides:
        combined_mapping.update(flag_overrides)

    combined_company_mapping = dict(DEFAULT_COMPANY_MAPPING)
    if company_overrides:
        combined_company_mapping.update(company_overrides)

    if hasattr(file_path_or_bytes, "seek"):
        file_path_or_bytes.seek(0)
    wb = openpyxl.load_workbook(file_path_or_bytes, data_only=True, read_only=True)

    all_records = []
    sku_records = []
    skipped_sheets = []

    for sheet_name in wb.sheetnames:
        if included_sheet_names is not None:
            if sheet_name not in included_sheet_names:
                skipped_sheets.append(sheet_name)
                continue
        elif is_excluded_sheet(sheet_name, exclude_keywords=exclude_keywords):
            skipped_sheets.append(sheet_name)
            continue

        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 6:
            continue

        state_or_zone, urban_rural, is_zone = parse_region_sheet_name(sheet_name)
        col_names = build_column_map(rows)

        current_parent_brand = None


        for r in rows[5:]:  # data starts after the two header rows (row 6 onward, 0-indexed 5)
            tg_segment = r[0]
            product = r[4]
            if tg_segment is None or product is None:
                continue  # junk / blank / universe / TG base row
            if tg_segment not in TG_SEGMENTS:
                continue

            flag = classify_flag(product, flag_overrides=combined_mapping)

            if flag == "SKU":
                if current_parent_brand is not None:
                    sku_record = {
                        "Format": format_tag,
                        "State_Zone": state_or_zone,
                        "Is_Zone": is_zone,
                        "Urban_Rural": urban_rural,
                        "TG_Segment": tg_segment,
                        "Parent_Brand": current_parent_brand,
                        "Brand_SKU_Item": str(product).strip(),
                    }
                    for i, col_name in enumerate(col_names):
                        if col_name is None:
                            continue
                        sku_record[col_name] = r[5 + i]
                    sku_records.append(sku_record)
                continue  # SKU rows are never part of Master_Clean

            # This is a Category/Brand/Sub-brand/Others row - it becomes the
            # parent for any SKU rows that follow it, until the next one.
            current_parent_brand = str(product).strip()

            company = classify_company(product, flag, company_overrides=combined_company_mapping)

            record = {
                "Format": format_tag,
                "Region_Raw": sheet_name,
                "State_Zone": state_or_zone,
                "Is_Zone": is_zone,
                "Urban_Rural": urban_rural,
                "TG_Segment": tg_segment,
                "Flag": flag,
                "Company": company,
                "Grammage": r[2],
                "SU": r[3],
                "Brand_SKU_Item": str(product).strip(),
            }
            for i, col_name in enumerate(col_names):
                if col_name is None:
                    continue
                val = r[5 + i]
                record[col_name] = val

            all_records.append(record)

    df = pd.DataFrame(all_records)
    sku_df = pd.DataFrame(sku_records)
    if verbose:
        print(f"[{format_tag}] Sheets read: {len(wb.sheetnames) - len(skipped_sheets)}, "
              f"skipped (AP+Tel): {skipped_sheets}, rows extracted: {len(df)}, SKU rows: {len(sku_df)}")
    return df, sku_df


def process_all_files(file_dict, flag_overrides=None, company_overrides=None, exclude_keywords=None, included_sheets_by_format=None, verbose=False):
    """
    file_dict: {format_tag: file_path_or_bytes}
    included_sheets_by_format: optional {format_tag: set of exact sheet names to include}
    Returns (master_df, sku_df) - the combined main data and the combined
    SKU-level data (used only for the Variance check), across all formats.
    """
    frames = []
    sku_frames = []
    for format_tag, f in file_dict.items():
        included = included_sheets_by_format.get(format_tag) if included_sheets_by_format else None
        df, sku_df = process_workbook(
            f, format_tag, flag_overrides=flag_overrides,
            company_overrides=company_overrides, exclude_keywords=exclude_keywords,
            included_sheet_names=included, verbose=verbose,
        )
        frames.append(df)
        sku_frames.append(sku_df)
    master_df = pd.concat(frames, ignore_index=True)
    sku_master_df = pd.concat(sku_frames, ignore_index=True) if sku_frames else pd.DataFrame()
    return master_df, sku_master_df
