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
# Zero factors that are CONFIRMED intentional (Sagar's actual given values,
# not placeholders). The run page shows these as a quiet informational note
# instead of the loud red zero-factor error. Any zero factor NOT in this set
# still triggers the full red error, because an accidental 0 silently zeroes
# out all Sales Derived / Units Estd for that state.
KNOWN_INTENTIONAL_ZERO_FACTORS = {("Delhi", "U")}

DEFAULT_INDIVIDUAL_FACTOR_UR = {
    # NOTE: "All India" is deliberately NOT listed here. It is always
    # auto-synthesized as the sum of every real state present (same as a
    # Zone), never given its own Individual Factor - see add_calculations().
    # A single constant factor for All India does not reproduce ground truth
    # exactly, confirming it must be built as a rollup, not a standalone factor.
    "Delhi": {"U": 0.00, "R": 1.00},  # Confirmed intentional: this is Sagar's
    # actual given value, not a placeholder. The zero-factor warning below will
    # still flag it every run, but only as a visible confirmation, not an error
    # to fix - Delhi Urban is meant to be 0 per Sagar.
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


def _compute_units_sales(df, factor_lookup, default_factor, missing_factor_regions, zero_factor_regions=None):
    """Compute Units Estd / Sales Derived columns in place for a (U or R only) dataframe."""
    periods = get_period_list(df)

    def lookup_factor(row):
        key = (row["State_Zone"], row["Urban_Rural"])
        if key in factor_lookup:
            val = factor_lookup[key]
            if val == 0 and zero_factor_regions is not None:
                zero_factor_regions.add(key)
            return val
        missing_factor_regions.add(key)
        return default_factor

    df["Individual_Factor"] = df.apply(lookup_factor, axis=1)

    for period in periods:
        hh_col, val_col, nop_col = f"HH__{period}", f"Val__{period}", f"Avg NOP__{period}"
        if hh_col not in df.columns or val_col not in df.columns:
            continue
        hh = pd.to_numeric(df[hh_col], errors="coerce")
        val = pd.to_numeric(df[val_col], errors="coerce")
        factor = df["Individual_Factor"]

        # Sales Derived only needs Val and the factor - compute it whenever
        # possible, independent of whether Avg NOP exists for this period.
        df[f"Sales Derived__{period}"] = factor * 1_000_000 * val

        # Units Estd genuinely needs Avg NOP - only compute it when that
        # column is actually present (e.g. Format B's self-computed annual
        # period deliberately has no confirmed Avg NOP yet, so Units Estd is
        # correctly left uncomputed for that period rather than guessed).
        if nop_col in df.columns:
            nop = pd.to_numeric(df[nop_col], errors="coerce")
            df[f"Units Estd__{period}"] = hh * nop * factor * 1000

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

    Returns (result, missing_factor_regions, unmapped_zones, zero_factor_regions).
    zero_factor_regions flags any (State, Urban_Rural) whose Individual Factor
    is exactly 0 - a real factor should never be truly zero, and a literal 0
    silently zeroes out every Sales Derived/Units Estd value for that entire
    state, with no other visible symptom (Value MS%/Units MS% just go blank
    from the divide-by-zero guard). This is reported so it's never missed.
    """
    if factor_lookup is None:
        factor_lookup = build_factor_lookup()
    if zone_mapping is None:
        zone_mapping = {}

    df = df.copy()
    missing_factor_regions = set()
    zero_factor_regions = set()
    unmapped_zones = set()

    # ---- Tier 0: collapse duplicate-identity rows ----
    # Some production files (e.g. the HC "Monthly & MAT Apr & May" workbook)
    # split one logical row across SEVERAL sheets: the same brand appears as
    # 2-3 rows, each carrying a different period's columns (one sheet holds
    # Monthly, another MAT Apr, another MAT May). Left as-is, every later
    # stage counts that brand 2-3 times: the U+R sum got stamped onto each
    # duplicate and rollups multiplied it (the exact 2x/3x tally failures).
    # Collapsing here, before ANY calculation, makes duplicates structurally
    # impossible downstream. sum(min_count=1) coalesces disjoint period
    # columns; for the vast majority of files (Formats A/B/C with one row per
    # identity) every group has exactly 1 row and this is a pure no-op,
    # which is why the Format A ground-truth match is unaffected.
    _raw_metric_cols = [c for c in df.columns if "__" in c]
    if _raw_metric_cols:
        _collapse_key = [c for c in ["Format", "State_Zone", "Is_Zone", "Urban_Rural",
                                     "TG_Segment", "Flag", "Company", "Brand_SKU_Item"]
                         if c in df.columns]
        for _ck in ["Grammage", "SU"]:
            if _ck in df.columns:
                df[f"__collapse_{_ck}"] = df[_ck].fillna("").astype(str).str.strip()
                _collapse_key.append(f"__collapse_{_ck}")
        _other_cols = [c for c in df.columns
                       if c not in _raw_metric_cols and c not in _collapse_key
                       and not c.startswith("__collapse_")]
        _agg = {c: "sum" for c in _raw_metric_cols} | {c: "first" for c in _other_cols}
        _before_n = len(df)
        df = (df.groupby(_collapse_key, dropna=False, sort=False)
                .agg({k: (lambda s: s.sum(min_count=1)) if v == "sum" else "first"
                      for k, v in _agg.items()})
                .reset_index())
        df = df.drop(columns=[c for c in df.columns if c.startswith("__collapse_")])

    # Any Is_Zone=True rows, and any "All India" rows, are ignored entirely
    # for calculation purposes - both are always synthesized as a rollup of
    # real states, never calculated from their own sheet.
    state_df = df[(df["Is_Zone"] != True) & (df["State_Zone"] != "All India")].copy()  # noqa: E712

    # ---- Tier 1 & 2: states (U, R direct; U+R derived) ----
    ur_mask = state_df["Urban_Rural"] == "U+R"
    non_ur_df = state_df[~ur_mask].copy()
    ur_raw_df = state_df[ur_mask].copy()

    periods = _compute_units_sales(non_ur_df, factor_lookup, default_factor, missing_factor_regions, zero_factor_regions)

    u_df = non_ur_df[non_ur_df["Urban_Rural"] == "U"]
    r_df = non_ur_df[non_ur_df["Urban_Rural"] == "R"]

    metric_cols = [c for c in non_ur_df.columns if c.startswith("Units Estd__") or c.startswith("Sales Derived__")]

    u_indexed = u_df.set_index(ID_COLS)
    r_indexed = r_df.set_index(ID_COLS)

    # U+R is ALWAYS the sum of whatever U and R rows exist for the same
    # identity (both if both exist, just U for urban-only rows, just R for
    # rural-only rows). This is a hard guarantee: a U+R value that doesn't
    # tally with its own U and R rows must be impossible by construction.
    #
    # History of why: the previous implementation summed only the U-and-R
    # intersection, and any raw U+R row it failed to pair fell back to being
    # computed from the raw U+R sheet's own HH/AvgNOP with an AVERAGED
    # factor ((U+R)/2). That fallback silently produced numbers that did not
    # equal U + R (e.g. Gujarat Units Estd Dec: 874,520 from the fallback vs
    # the correct 878,819 = U + R), and whether it triggered depended on
    # fragile key alignment that differed between sessions. The raw-sheet
    # fallback is now used ONLY when a raw U+R row has no U row and no R row
    # at all to sum from.
    # Pair U/R rows with U+R rows on the FULL row identity, including
    # Grammage and SU. Production files (e.g. HC) contain several rows that
    # share the same Brand_SKU_Item and differ only by Grammage; pairing
    # without Grammage collapsed all of a brand's grammage rows into one sum
    # and stamped that TOTAL onto every one of its U+R grammage rows -
    # inflating each of them (the 314-datapoint tally failure). Grammage and
    # SU are normalized (NaN -> '', stripped) purely for the pairing so blank
    # vs missing can never break a match; stored values are untouched.
    extra_pair_cols = [c for c in ["Grammage", "SU"] if c in non_ur_df.columns]
    pair_cols = ID_COLS + [f"__pair_{c}" for c in extra_pair_cols]
    for frame in (non_ur_df, ur_raw_df):
        for c in extra_pair_cols:
            frame[f"__pair_{c}"] = frame[c].fillna("").astype(str).str.strip()

    sum_ur = non_ur_df.groupby(pair_cols, dropna=False)[metric_cols].sum(min_count=1).reset_index()
    sum_ur["Urban_Rural"] = "U+R"

    ur_raw_df = ur_raw_df.merge(sum_ur, on=pair_cols + ["Urban_Rural"], how="left", suffixes=("", ""))

    # A raw U+R row genuinely missed the sum-merge only if ALL its metric
    # columns are blank. Never test just one period column: in a multi-format
    # run the period columns span every format, and a monthly row is
    # legitimately blank in another format's quarterly column. The old check
    # used periods[0] alone, so in any multi-file run every row of a format
    # that didn't own periods[0] was falsely treated as missed and its
    # correct summed U+R was overwritten by the raw-sheet fallback - the
    # exact source of the mass tally failures.
    missing_ur_mask = ur_raw_df[metric_cols].isna().all(axis=1) if metric_cols else pd.Series(False, index=ur_raw_df.index)
    if missing_ur_mask.any():
        fallback = ur_raw_df[missing_ur_mask].drop(columns=metric_cols)
        _compute_units_sales(fallback, factor_lookup, default_factor, missing_factor_regions, zero_factor_regions)
        ur_raw_df.loc[missing_ur_mask, metric_cols] = fallback[metric_cols].values

    # Urban-only states: some states report only an Urban sheet, with no Rural
    # and no U+R sheet at all (e.g. Delhi, Guwahati in the Hair Colour files).
    # For these the market's total IS its urban figure, so U+R = U. Without
    # this they are dropped from every U+R rollup (All India U+R, zone U+R),
    # understating the national total - exactly the "U+R not adding in Delhi
    # and Guwahati" issue Sagar raised. We synthesize a full U+R row (copy of
    # the U row, relabelled) ONLY for states that have NO Rural row anywhere,
    # so a state with both U and R is never touched and this can't double count.
    states_with_u = set(u_df["State_Zone"].unique())
    states_with_r = set(r_df["State_Zone"].unique())
    states_with_raw_ur = set(ur_raw_df["State_Zone"].unique())
    u_only = u_df[
        (~u_df["State_Zone"].isin(states_with_r))
        & (~u_df["State_Zone"].isin(states_with_raw_ur))
    ]
    synthesized_u_only_states = sorted(u_only["State_Zone"].unique())
    if len(u_only) > 0:
        u_only_ur = u_only.copy()
        u_only_ur["Urban_Rural"] = "U+R"
        ur_raw_df = pd.concat([ur_raw_df, u_only_ur], ignore_index=True)

    # Mirror case: a state with ONLY a Rural sheet (no Urban, no raw U+R).
    # Then rural IS the market total, so U+R = R, symmetric with the
    # urban-only rule above. No current file has this shape, but the rule
    # must be symmetric rather than silently dropping such a state from
    # every U+R rollup.
    r_only = r_df[
        (~r_df["State_Zone"].isin(states_with_u))
        & (~r_df["State_Zone"].isin(states_with_raw_ur))
    ]
    if len(r_only) > 0:
        r_only_ur = r_only.copy()
        r_only_ur["Urban_Rural"] = "U+R"
        ur_raw_df = pd.concat([ur_raw_df, r_only_ur], ignore_index=True)

    # Drop the temporary pairing columns before assembling the result.
    for frame in (non_ur_df, ur_raw_df):
        frame.drop(columns=[c for c in frame.columns if c.startswith("__pair_")], inplace=True, errors="ignore")

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

    return result, missing_factor_regions, unmapped_zones, zero_factor_regions


def build_tally_check(result_df, tolerance=1.0):
    """
    Automatic QC: for every state that has U and/or R rows, verify that the
    U+R row's additive metrics (Units Estd, Sales Derived) equal the sum of
    the U and R rows, per identity, per period. Returns a DataFrame of every
    mismatch beyond `tolerance` (absolute). Empty DataFrame = fully tallies.

    Runs on every pipeline run and is exported into the Excel, so a
    non-tallying number can never ship silently. Zones and All India are
    covered implicitly: they are sums of state rows, so if states tally,
    rollups tally.
    """
    cols = ["Format", "State_Zone", "TG_Segment", "Flag", "Brand_SKU_Item",
            "Metric", "Period", "U_plus_R", "UR_row", "Difference"]
    if result_df is None or len(result_df) == 0:
        return pd.DataFrame(columns=cols)

    is_zone = result_df["Is_Zone"] == True if "Is_Zone" in result_df.columns else pd.Series(False, index=result_df.index)  # noqa: E712
    states = result_df[(~is_zone) & (result_df["State_Zone"] != "All India")]

    metric_cols = [c for c in states.columns
                   if c.startswith("Units Estd__") or c.startswith("Sales Derived__")]
    if not metric_cols:
        return pd.DataFrame(columns=cols)

    key = ["Format", "State_Zone", "TG_Segment", "Flag", "Brand_SKU_Item"]
    # Include Grammage/SU in the identity (normalized) so rows that share a
    # brand name but differ by grammage are checked individually. Comparing
    # GROUP SUMS on both sides (rather than row-by-row label alignment) also
    # makes duplicated keys impossible to hide behind: if each duplicate U+R
    # row were wrongly stamped with the full total, the U+R group sum would
    # exceed the U/R group sum and be flagged.
    states = states.copy()
    for c in ["Grammage", "SU"]:
        if c in states.columns:
            states[f"__k_{c}"] = states[c].fillna("").astype(str).str.strip()
            key.append(f"__k_{c}")

    ur_rows = states[states["Urban_Rural"] == "U+R"].groupby(key, dropna=False)[metric_cols].sum(min_count=1)
    sum_rows = (states[states["Urban_Rural"].isin(["U", "R"])]
                .groupby(key, dropna=False)[metric_cols].sum(min_count=1))

    common = ur_rows.index.intersection(sum_rows.index)
    display_key = ["Format", "State_Zone", "TG_Segment", "Flag", "Brand_SKU_Item"]
    records = []
    for c in metric_cols:
        expected = sum_rows.loc[common, c]
        actual = ur_rows.loc[common, c]
        diff = (actual - expected).abs()
        bad = diff[diff.fillna(0) > tolerance]
        metric, period = c.split("__", 1)
        for idx, d in bad.items():
            rec = dict(zip(key, idx if isinstance(idx, tuple) else (idx,)))
            rec = {k: v for k, v in rec.items() if not k.startswith("__k_")}
            rec.update({
                "Metric": metric, "Period": period,
                "U_plus_R": expected.loc[idx], "UR_row": actual.loc[idx],
                "Difference": d,
            })
            records.append(rec)
    return pd.DataFrame(records, columns=cols)


def build_rollup_coverage(result_df, zone_mapping=None):
    """
    Reports, for every rollup (All India and each mapped Zone) at every U/R
    cut in every Format, how many member states actually contributed vs how
    many exist in the data at all. Purely reads the finished result - it can
    never change a number.

    Why this exists: rollups are sums of whatever states are present. If a
    state's sheet was de-selected or absent (e.g. only Telangana had a U+R
    sheet), the rollup is still produced and labelled "All India" while
    actually being one state. That must be reported loudly, not discovered
    by the reader.

    "Expected" for All India = every state present in that Format at ANY
    U/R cut. "Expected" for a Zone = its mapped member states present in
    that Format at any cut (members entirely absent from the data are
    already reported separately via unmapped_zones and are not counted
    here). A row is emitted whenever expected > 0, including when included
    is 0 (the rollup cut then simply doesn't exist in the output).

    States are identified by the Is_Zone flag, never by name matching -
    "West Bengal" is a state and must never be confused with the "West"
    zone.
    """
    cols = ["Format", "Rollup", "Urban_Rural", "States_Included",
            "States_Expected", "Missing_States", "Missing_Detail"]
    if result_df is None or len(result_df) == 0:
        return pd.DataFrame(columns=cols)
    zone_mapping = zone_mapping or {}

    df = result_df
    is_zone = df["Is_Zone"] == True if "Is_Zone" in df.columns else pd.Series(False, index=df.index)  # noqa: E712
    states_df = df[(~is_zone) & (df["State_Zone"] != "All India")]

    def _describe_missing(fsub, missing_states, ur_cut):
        """Plain-language, per-state explanation of WHY a state is absent
        from a rollup cut: what the raw file actually contains for it, and
        what the model therefore did. Written for a reader who should never
        have to guess whether the model dropped data or the data never
        existed."""
        parts = []
        for st in missing_states:
            cuts = set(fsub[fsub["State_Zone"] == st]["Urban_Rural"].unique())
            if cuts == {"U", "U+R"} or cuts == {"U"}:
                parts.append(f"{st} (raw file has Urban only - no Rural sheet exists; "
                             f"the model set its U+R = Urban, so it IS counted in U and U+R rollups; "
                             f"it is absent only from Rural because no Rural data was ever supplied)")
            elif cuts == {"R", "U+R"} or cuts == {"R"}:
                parts.append(f"{st} (raw file has Rural only - no Urban sheet exists; "
                             f"the model set its U+R = Rural, so it IS counted in R and U+R rollups; "
                             f"it is absent only from Urban because no Urban data was ever supplied)")
            elif ur_cut not in cuts and cuts:
                parts.append(f"{st} (raw file contains {', '.join(sorted(cuts))} but no {ur_cut} sheet)")
            else:
                parts.append(f"{st} (present in the data but excluded from this cut - REVIEW: "
                             f"this is the dangerous case, e.g. a de-selected sheet)")
        return "; ".join(parts)

    rows = []
    for fmt in sorted(states_df["Format"].dropna().unique()):
        fsub = states_df[states_df["Format"] == fmt]
        all_states = set(fsub["State_Zone"].unique())
        if not all_states:
            continue
        for ur in ["U", "R", "U+R"]:
            present = set(fsub[fsub["Urban_Rural"] == ur]["State_Zone"].unique())
            missing = sorted(all_states - present)
            rows.append({
                "Format": fmt, "Rollup": "All India", "Urban_Rural": ur,
                "States_Included": len(present), "States_Expected": len(all_states),
                "Missing_States": ", ".join(missing),
                "Missing_Detail": _describe_missing(fsub, missing, ur),
            })
            for zone, members in zone_mapping.items():
                expected = set(members) & all_states
                if not expected:
                    continue
                contributing = expected & present
                z_missing = sorted(expected - contributing)
                rows.append({
                    "Format": fmt, "Rollup": zone, "Urban_Rural": ur,
                    "States_Included": len(contributing), "States_Expected": len(expected),
                    "Missing_States": ", ".join(z_missing),
                    "Missing_Detail": _describe_missing(fsub, z_missing, ur),
                })
    return pd.DataFrame(rows, columns=cols)


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
    # 'Subtotal' rows are redundant company/brand subtotals baked into some
    # source files (e.g. 'ANY GCPL SHC (GEE+SELFIE)'). They are NEITHER the
    # category denominator NOR a summable company row - counting them double-
    # counts the brands they aggregate. Exclude them from both sides.
    subtotal_mask = result_df["Flag"] == "Subtotal"
    company_mask = ~cat_mask & ~subtotal_mask

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
