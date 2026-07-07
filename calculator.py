"""
Godrej Market Share Model - Calculation Module

Formulas (validated against ground truth on Rajasthan/SHC before use):
  Units Estd    = HH * Avg NOP * Individual Factor * 1000
  Sales Derived = Units Estd * [(Val * 1000) / (HH * Avg NOP)]
                = Individual Factor * 1,000,000 * Val
  (Val is in million Rs, HH is in 000s households, per the sheet's own legend)

Market Share % (within a Region + Urban_Rural + TG_Segment + Format + Period):
  Units MS%  = Brand Units Estd  / Category Units Estd  (for the same group)
  Value MS%  = Brand Sales Derived / Category Sales Derived (for the same group)

Individual Factor is looked up per (State_Zone, Urban_Rural) region. U+R value
defaults to the average of U and R if not separately given.
"""

import pandas as pd
import numpy as np
import difflib

FUZZY_MATCH_THRESHOLD = 0.87
DUPLICATE_THRESHOLD = 0.92
NON_COMPANY_VALUES = {"Category Total", "Others / Unmapped", "Review Required"}


def _normalize(name):
    return str(name).strip().upper()


GENERIC_WORDS = {
    "COMPANY", "COMPANIES", "INDIA", "LTD", "LIMITED", "PVT", "PRIVATE",
    "CORP", "CORPORATION", "INTERNATIONAL", "INDS", "INDUSTRIES", "INDUSTRY",
    "CARE", "EXPORT", "EXPORTS", "GROUP", "ENTERPRISE", "ENTERPRISES",
    "CONSUMER", "PRODUCTS", "PRODS", "AND", "THE", "CONCEPTS",
}


def build_company_keywords(company_mapping):
    """
    Derives a {keyword: Company} lookup from a distinctive word in each real
    company name already in the mapping (skips placeholder values like
    'Others / Unmapped', and skips generic corporate words like 'Company' or
    'Industries' that would cause false-positive matches across unrelated
    companies). Used to smart-guess a new brand's company by keyword, e.g. a
    new item containing 'GODREJ' likely belongs to whichever company name
    also contains 'GODREJ'.
    """
    keywords = {}
    companies = sorted(set(v for v in company_mapping.values() if v not in NON_COMPANY_VALUES))
    for comp in companies:
        for word in comp.replace("'", " ").split():
            word_upper = word.upper()
            if len(word_upper) > 3 and word_upper not in GENERIC_WORDS and word_upper not in keywords:
                keywords[word_upper] = comp
                break
    return keywords


def fuzzy_match_company(item_name, existing_mapping, threshold=FUZZY_MATCH_THRESHOLD):
    """
    Looks for a near-duplicate of item_name already in existing_mapping
    (e.g. a typo'd or slightly reformatted repeat of a brand already
    classified) and, if similar enough, returns that brand's company.
    Returns (company_or_None, best_score, matched_item_or_None).
    """
    target = _normalize(item_name)
    best_item, best_score = None, 0.0
    for existing_item in existing_mapping:
        score = difflib.SequenceMatcher(None, target, _normalize(existing_item)).ratio()
        if score > best_score:
            best_score, best_item = score, existing_item
    if best_score >= threshold and best_item is not None:
        return existing_mapping[best_item], best_score, best_item
    return None, best_score, best_item


def smart_guess_companies(new_items, existing_mapping):
    """
    For each item in new_items (not already in existing_mapping), tries to
    guess its Company:
      1. Fuzzy match against existing items (catches near-duplicate names)
      2. Keyword match against existing company names
      3. Falls back to 'Others / Unmapped'
    Returns {item: (guessed_company, status_label)}.
    """
    keywords = build_company_keywords(existing_mapping)
    results = {}
    for item in new_items:
        comp, score, matched_item = fuzzy_match_company(item, existing_mapping)
        if comp:
            results[item] = (comp, f"Smart-guessed: {int(score*100)}% match to '{matched_item}'")
            continue
        comp = None
        name_upper = _normalize(item)
        for kw, kw_comp in keywords.items():
            if kw in name_upper:
                comp = kw_comp
                break
        if comp:
            results[item] = (comp, "Smart-guessed: name keyword match")
        else:
            results[item] = ("Others / Unmapped", "New - needs review")
    return results


def find_potential_duplicates(item_list, threshold=DUPLICATE_THRESHOLD):
    """
    Scans a list of Brand_SKU_Item names for near-duplicate pairs (possible
    typos / inconsistent naming of the same brand), excluding exact matches.
    Returns a list of (item1, item2, similarity_score) sorted by score desc.
    """
    items = list(dict.fromkeys(item_list))  # de-dupe while preserving order
    pairs = []
    for i in range(len(items)):
        for j in range(i + 1, len(items)):
            a, b = items[i], items[j]
            if a == b:
                continue
            score = difflib.SequenceMatcher(None, _normalize(a), _normalize(b)).ratio()
            if score >= threshold:
                pairs.append((a, b, score))
    return sorted(pairs, key=lambda x: -x[2])

# Individual Factor list (U and R values as given; U+R computed as average).
# These were spot-checked against the ground truth file for Rajasthan and
# landed within ~1% of the true value, so this table is a solid starting point.
# Editable in the dashboard before running.
DEFAULT_INDIVIDUAL_FACTOR_UR = {
    # NOTE: "All India" is deliberately NOT listed here. It is always
    # auto-synthesized as the sum of every real state present (same as a
    # Zone), never given its own Individual Factor - see add_calculations().
    # A single constant factor for All India does not reproduce ground truth
    # exactly, confirming it must be built as a rollup, not a standalone factor.
    "Delhi": {"U": 0.00, "R": 1.00},
    "Punjab_Haryana": {"U": 1.00, "R": 1.90},
    "Rajasthan": {"U": 1.32, "R": 1.24},
    "Uttar Pradesh": {"U": 2.10, "R": 1.87},
    "West Bengal": {"U": 1.80, "R": 1.62},
    "Bihar excl Jharkhand": {"U": 1.00, "R": 1.00},
    "Jharkhand": {"U": 1.00, "R": 1.00},
    "Guwahati": {"U": 1.00, "R": 1.00},
    "Orissa": {"U": 1.00, "R": 1.70},
    "Maharashtra": {"U": 1.85, "R": 1.85},
    "Gujarat": {"U": 1.90, "R": 1.85},
    "Madhya Pradesh excl Chha": {"U": 1.40, "R": 1.30},
    "Chhattisgarh": {"U": 1.00, "R": 1.00},
    "Tamil Nadu": {"U": 2.84, "R": 3.90},
    "Karnataka": {"U": 1.60, "R": 1.73},
    "Kerala": {"U": 1.10, "R": 7.88},
    "Andhra Pradesh excl Tela": {"U": 1.50, "R": 1.55},
    "Telangana": {"U": 1.00, "R": 1.63},
}

# Zone -> member States, confirmed directly from Sagar's ground truth reference
# file (the 'Zone' column against each Market Cut). Only the standard
# North/South/East/West zones are confirmed this way; any other zone (e.g.
# a Godrej-internal "GCPL East"/"GCPL West" split) is NOT in that file and
# must be supplied separately, there is no safe way to guess it.
DEFAULT_ZONE_MAPPING = {
    "North": ["Delhi", "Punjab_Haryana", "Rajasthan", "Uttar Pradesh"],
    "South": ["Andhra Pradesh excl Tela", "Karnataka", "Kerala", "Tamil Nadu", "Telangana"],
    "East": ["Bihar excl Jharkhand", "Chhattisgarh", "Guwahati", "Jharkhand", "Orissa", "West Bengal"],
    "West": ["Gujarat", "Madhya Pradesh excl Chha", "Maharashtra"],
    # Not confirmed from Sagar's file (that file has no "GCPL East/West" sheets
    # at all). Treated as = the standard East / West zones per your instruction,
    # still fully editable in the dashboard if that turns out to be wrong.
    "GCPL East": ["Bihar excl Jharkhand", "Chhattisgarh", "Guwahati", "Jharkhand", "Orissa", "West Bengal"],
    "GCPL West": ["Gujarat", "Madhya Pradesh excl Chha", "Maharashtra"],
}

# Reverse lookup (State -> Zone), built only from the 4 confirmed standard
# zones (not the GCPL East/West duplicates, which would make this ambiguous).
# Used to auto-populate Zone Mapping from whichever states are actually
# present in the uploaded data, instead of always showing a static full list.
STATE_TO_ZONE = {}
for _zone in ("North", "South", "East", "West"):
    for _state in DEFAULT_ZONE_MAPPING.get(_zone, []):
        STATE_TO_ZONE[_state] = _zone

# Which regions have actually been checked against a ground truth file and
# confirmed exact (see calibration done for Rajasthan/SHC). Everything else
# above is an unverified placeholder from an early session, kept only as a
# starting point, not a confirmed number.
VERIFIED_REGIONS = {"Rajasthan"}


def factor_status(state):
    return "Verified (exact match)" if state in VERIFIED_REGIONS else "Unverified - placeholder, please confirm"


def build_factor_lookup(factor_ur_dict=None):
    """
    Turns the {state: {U:.., R:..}} dict into a flat {(state, urban_rural): factor}
    lookup, including a computed U+R = average(U, R).
    """
    if factor_ur_dict is None:
        factor_ur_dict = DEFAULT_INDIVIDUAL_FACTOR_UR

    flat = {}
    for state, urvals in factor_ur_dict.items():
        u = urvals.get("U")
        r = urvals.get("R")
        if u is not None:
            flat[(state, "U")] = u
        if r is not None:
            flat[(state, "R")] = r
        if u is not None and r is not None:
            flat[(state, "U+R")] = round((u + r) / 2, 4)
        elif u is not None:
            flat[(state, "U+R")] = u
        elif r is not None:
            flat[(state, "U+R")] = r
    return flat


def get_period_list(df):
    """All distinct periods found in HH__ columns, in original column order."""
    periods = []
    for col in df.columns:
        if col.startswith("HH__"):
            periods.append(col.replace("HH__", ""))
    return periods


ID_COLS = ["Format", "State_Zone", "Is_Zone", "TG_Segment", "Flag", "Brand_SKU_Item"]


def _compute_units_sales(df, factor_lookup, default_factor, missing_factor_regions):
    """Compute Units Estd / Sales Derived columns in place for a (U or R only) dataframe."""
    periods = get_period_list(df)

    def lookup_factor(row):
        key = (row["State_Zone"], row["Urban_Rural"])
        if key in factor_lookup:
            return factor_lookup[key]
        missing_factor_regions.add(key)
        return default_factor

    df["Individual_Factor"] = df.apply(lookup_factor, axis=1)

    for period in periods:
        hh_col, val_col, nop_col = f"HH__{period}", f"Val__{period}", f"Avg NOP__{period}"
        if hh_col not in df.columns or val_col not in df.columns or nop_col not in df.columns:
            continue
        hh = pd.to_numeric(df[hh_col], errors="coerce")
        val = pd.to_numeric(df[val_col], errors="coerce")
        nop = pd.to_numeric(df[nop_col], errors="coerce")
        factor = df["Individual_Factor"]

        df[f"Units Estd__{period}"] = hh * nop * factor * 1000
        df[f"Sales Derived__{period}"] = factor * 1_000_000 * val

    return periods


def add_calculations(df, factor_lookup=None, default_factor=1.0, zone_mapping=None):
    """
    Adds Units Estd / Sales Derived / Units MS% / Value MS% for every period.

    Three tiers, each handled differently (an Individual Factor only ever
    applies at the STATE level, never above it):

    1. State (U) / State (R): looked up directly from factor_lookup and
       calculated from that state's own HH/Val/Avg NOP.

    2. State (U+R): NOT calculated from the U+R sheet's own HH/Val/Avg NOP.
       Validated exactly against ground truth: U+R = U result + R result for
       the same Format/State/TG_Segment/Flag/Brand_SKU_Item/Period.

    3. Zone (U) / Zone (R) / Zone (U+R) (e.g. "North", "GCPL East") AND
       "All India": all of these are pure rollups of member states - their
       rows are SYNTHESIZED fresh from the already-calculated state results,
       and NEVER read from their own sheet's HH/Val/Avg NOP, even if one is
       present in the raw file (it's ignored for calculation). "All India"
       is treated exactly like a zone here: it is automatically the sum of
       every real state present in the data, with no factor of its own and
       no manual zone_mapping entry needed. This matters because a single
       constant "All India Individual Factor" does not reproduce the ground
       truth exactly (unlike states, which do), meaning All India is not a
       simple single-factor panel - it must be built as a rollup, the same
       as any other zone. Any zone in zone_mapping whose member states are
       entirely absent from the current data is reported back in
       `unmapped_zones` rather than silently skipped.

    factor_lookup: dict {(State_Zone, Urban_Rural): factor} where Urban_Rural
    is 'U' or 'R' only (U+R is derived, never looked up). Only applies to
    real states (rows with Is_Zone == False and State_Zone != "All India").

    zone_mapping: dict {Zone name: [State name, ...]}.
    """
    if factor_lookup is None:
        factor_lookup = build_factor_lookup()
    if zone_mapping is None:
        zone_mapping = {}

    df = df.copy()
    missing_factor_regions = set()
    unmapped_zones = set()

    # Any Is_Zone=True rows, and any "All India" rows, are ignored entirely
    # for calculation purposes - both are always synthesized as a rollup of
    # real states, never calculated from their own sheet.
    state_df = df[(df["Is_Zone"] != True) & (df["State_Zone"] != "All India")].copy()  # noqa: E712

    # ---- Tier 1 & 2: states (U, R direct; U+R derived) ----
    ur_mask = state_df["Urban_Rural"] == "U+R"
    non_ur_df = state_df[~ur_mask].copy()
    ur_raw_df = state_df[ur_mask].copy()

    periods = _compute_units_sales(non_ur_df, factor_lookup, default_factor, missing_factor_regions)

    u_df = non_ur_df[non_ur_df["Urban_Rural"] == "U"]
    r_df = non_ur_df[non_ur_df["Urban_Rural"] == "R"]

    metric_cols = [c for c in non_ur_df.columns if c.startswith("Units Estd__") or c.startswith("Sales Derived__")]

    u_indexed = u_df.set_index(ID_COLS)
    r_indexed = r_df.set_index(ID_COLS)

    common_keys = u_indexed.index.intersection(r_indexed.index)
    derived_ur = u_indexed.loc[common_keys, metric_cols].add(
        r_indexed.loc[common_keys, metric_cols].values
    )
    derived_ur = derived_ur.reset_index()
    derived_ur["Urban_Rural"] = "U+R"

    ur_raw_df = ur_raw_df.merge(derived_ur, on=ID_COLS + ["Urban_Rural"], how="left", suffixes=("", ""))

    missing_ur_mask = ur_raw_df["Units Estd__" + periods[0]].isna() if periods else pd.Series(False, index=ur_raw_df.index)
    if missing_ur_mask.any():
        fallback = ur_raw_df[missing_ur_mask].drop(columns=metric_cols)
        _compute_units_sales(fallback, factor_lookup, default_factor, missing_factor_regions)
        ur_raw_df.loc[missing_ur_mask, metric_cols] = fallback[metric_cols].values

    state_result = pd.concat([non_ur_df, ur_raw_df], ignore_index=True)

    # ---- Tier 3: zones, synthesized purely from zone_mapping + state_result ----
    zone_group_cols = ["Format", "TG_Segment", "Flag", "Company", "Brand_SKU_Item", "Urban_Rural"]
    zone_group_cols = [c for c in zone_group_cols if c in state_result.columns]

    zone_records = []
    for zone_name, member_states in zone_mapping.items():
        sub = state_result[state_result["State_Zone"].isin(member_states)]
        if len(sub) == 0:
            unmapped_zones.add(zone_name)
            continue
        grouped = sub.groupby(zone_group_cols, dropna=False)[metric_cols].sum(min_count=1).reset_index()
        grouped["State_Zone"] = zone_name
        grouped["Is_Zone"] = True
        zone_records.append(grouped)

    # "All India" is always auto-synthesized as the sum of every real state
    # present in the data (not dependent on zone_mapping listing every state
    # manually). If no states are present at all, it's simply not produced.
    if len(state_result) > 0:
        grouped_national = state_result.groupby(zone_group_cols, dropna=False)[metric_cols].sum(min_count=1).reset_index()
        grouped_national["State_Zone"] = "All India"
        grouped_national["Is_Zone"] = True
        zone_records.append(grouped_national)

    zone_result = pd.concat(zone_records, ignore_index=True) if zone_records else pd.DataFrame(columns=state_result.columns)

    result = pd.concat([state_result, zone_result], ignore_index=True)

    # Market share %, computed within each Region+UR+TG_Segment+Format+Period group,
    # relative to that group's Category row.
    group_cols = ["Format", "State_Zone", "Urban_Rural", "TG_Segment"]

    for period in periods:
        units_col = f"Units Estd__{period}"
        sales_col = f"Sales Derived__{period}"
        if units_col not in result.columns:
            continue

        cat_units = result[result["Flag"] == "Category"].groupby(group_cols)[units_col].first()
        cat_sales = result[result["Flag"] == "Category"].groupby(group_cols)[sales_col].first()

        keys = list(zip(result["Format"], result["State_Zone"], result["Urban_Rural"], result["TG_Segment"]))
        result[f"Units MS%__{period}"] = [
            (result[units_col].iloc[i] / cat_units.get(k)) if cat_units.get(k) not in (None, 0) and pd.notna(cat_units.get(k)) else np.nan
            for i, k in enumerate(keys)
        ]
        result[f"Value MS%__{period}"] = [
            (result[sales_col].iloc[i] / cat_sales.get(k)) if cat_sales.get(k) not in (None, 0) and pd.notna(cat_sales.get(k)) else np.nan
            for i, k in enumerate(keys)
        ]

    return result, missing_factor_regions, unmapped_zones


def build_company_summary(result_df):
    """
    Rolls the brand-level result up to Company level (GCPL, L'Oreal India,
    Cavinkare, etc.) within each Format + State_Zone + Urban_Rural + TG_Segment,
    giving Company Units MS% and Company Value MS% against the Category total,
    same concept as the "gcpl" row in Sagar's own pivot table.

    Category rows (Company == 'Category Total') are excluded from the summed
    companies but their totals are used as the denominator for MS%.
    """
    periods = get_period_list(result_df)
    group_cols = ["Format", "State_Zone", "Urban_Rural", "TG_Segment"]

    cat_mask = result_df["Flag"] == "Category"
    company_mask = ~cat_mask

    sum_cols = [f"Units Estd__{p}" for p in periods] + [f"Sales Derived__{p}" for p in periods]
    sum_cols = [c for c in sum_cols if c in result_df.columns]

    company_totals = (
        result_df[company_mask]
        .groupby(group_cols + ["Company"])[sum_cols]
        .sum()
        .reset_index()
    )

    cat_totals = (
        result_df[cat_mask]
        .groupby(group_cols)[sum_cols]
        .first()
    )

    keys = list(zip(
        company_totals["Format"], company_totals["State_Zone"],
        company_totals["Urban_Rural"], company_totals["TG_Segment"],
    ))

    for period in periods:
        units_col = f"Units Estd__{period}"
        sales_col = f"Sales Derived__{period}"
        if units_col not in company_totals.columns:
            continue
        cat_units = cat_totals[units_col] if units_col in cat_totals.columns else pd.Series(dtype=float)
        cat_sales = cat_totals[sales_col] if sales_col in cat_totals.columns else pd.Series(dtype=float)

        company_totals[f"Units MS%__{period}"] = [
            (company_totals[units_col].iloc[i] / cat_units.get(k)) if cat_units.get(k) not in (None, 0) else np.nan
            for i, k in enumerate(keys)
        ]
        company_totals[f"Value MS%__{period}"] = [
            (company_totals[sales_col].iloc[i] / cat_sales.get(k)) if cat_sales.get(k) not in (None, 0) else np.nan
            for i, k in enumerate(keys)
        ]

    return company_totals


def compute_variance(master_df, sku_df):
    """
    Variance check: for every Brand/Sub-brand/Others row, compares its own
    reported panel total against the sum of its listed SKUs, using the raw
    panel formula (no Individual Factor, no x1000):
        Units (brand's own row)      = HH * Avg NOP
        Units (each SKU under it)    = HH * Avg NOP
        SKU_Sum                      = sum of that brand's SKU Units
        Variance                     = Brand's own Units - SKU_Sum
        Variance %                  = Variance / Brand's own Units

    A non-zero Variance is expected and normal (Kantar brand totals include
    small/unlisted SKUs not individually broken out); a large Variance % is
    what's worth flagging for review.

    Returns a dataframe: Format, State_Zone, Urban_Rural, TG_Segment,
    Brand_SKU_Item, and per period: Brand_Units, SKU_Sum_Units, Variance,
    Variance_Pct.
    """
    periods = get_period_list(master_df)
    group_cols = ["Format", "State_Zone", "Urban_Rural", "TG_Segment", "Brand_SKU_Item"]

    brand_rows = master_df[master_df["Flag"] != "Category"].copy()
    for period in periods:
        hh_col, nop_col = f"HH__{period}", f"Avg NOP__{period}"
        if hh_col in brand_rows.columns and nop_col in brand_rows.columns:
            brand_rows[f"__brand_units__{period}"] = (
                pd.to_numeric(brand_rows[hh_col], errors="coerce") *
                pd.to_numeric(brand_rows[nop_col], errors="coerce")
            )

    if sku_df is None or len(sku_df) == 0:
        result = brand_rows[group_cols].drop_duplicates().reset_index(drop=True)
        for period in periods:
            result[f"Brand_Units__{period}"] = np.nan
            result[f"SKU_Sum_Units__{period}"] = np.nan
            result[f"Variance__{period}"] = np.nan
            result[f"Variance_Pct__{period}"] = np.nan
        return result

    sku_df = sku_df.copy()
    for period in periods:
        hh_col, nop_col = f"HH__{period}", f"Avg NOP__{period}"
        if hh_col in sku_df.columns and nop_col in sku_df.columns:
            sku_df[f"__sku_units__{period}"] = (
                pd.to_numeric(sku_df[hh_col], errors="coerce") *
                pd.to_numeric(sku_df[nop_col], errors="coerce")
            )

    sku_group_cols = ["Format", "State_Zone", "Urban_Rural", "TG_Segment", "Parent_Brand"]
    sku_sum_cols = [c for c in sku_df.columns if c.startswith("__sku_units__")]
    sku_sums = sku_df.groupby(sku_group_cols)[sku_sum_cols].sum().reset_index()
    sku_sums = sku_sums.rename(columns={"Parent_Brand": "Brand_SKU_Item"})

    result = brand_rows[group_cols + [c for c in brand_rows.columns if c.startswith("__brand_units__")]].copy()
    result = result.merge(sku_sums, on=group_cols, how="left")

    for period in periods:
        brand_col = f"__brand_units__{period}"
        sku_col = f"__sku_units__{period}"
        if brand_col not in result.columns:
            continue
        brand_units = result.get(brand_col)
        sku_units = result[sku_col] if sku_col in result.columns else pd.Series(np.nan, index=result.index)
        result[f"Brand_Units__{period}"] = brand_units
        result[f"SKU_Sum_Units__{period}"] = sku_units
        result[f"Variance__{period}"] = brand_units - sku_units
        result[f"Variance_Pct__{period}"] = np.where(
            (brand_units.notna()) & (brand_units != 0),
            (brand_units - sku_units) / brand_units,
            np.nan,
        )

    drop_cols = [c for c in result.columns if c.startswith("__brand_units__") or c.startswith("__sku_units__")]
    result = result.drop(columns=drop_cols)

    return result
