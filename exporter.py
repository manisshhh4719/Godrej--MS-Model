"""
Godrej Market Share Model - Export Module
Writes the final Master_Clean output plus two editable reference sheets:
  - Brand_Mapping: every unique Brand_SKU_Item with its current Flag
    (Category / Brand / Sub-brand / Others). Edit this and re-upload next
    time to change how any item is classified.
  - Individual_Factor: the Urban / Rural factor used per region. Edit this
    and re-upload next time to correct or add regions.
"""

import pandas as pd
from io import BytesIO


def build_brand_mapping_sheet(df):
    mapping = (
        df[["Brand_SKU_Item", "Flag"]]
        .drop_duplicates()
        .sort_values(["Flag", "Brand_SKU_Item"])
        .reset_index(drop=True)
    )
    return mapping


def build_individual_factor_sheet(factor_ur_dict):
    rows = []
    for state, urvals in factor_ur_dict.items():
        rows.append({
            "State_Zone": state,
            "Individual_Factor_U": urvals.get("U"),
            "Individual_Factor_R": urvals.get("R"),
        })
    return pd.DataFrame(rows).sort_values("State_Zone").reset_index(drop=True)


def build_company_mapping_sheet(df):
    mapping = (
        df[df["Flag"] != "Category"][["Brand_SKU_Item", "Company"]]
        .drop_duplicates()
        .sort_values(["Company", "Brand_SKU_Item"])
        .reset_index(drop=True)
    )
    return mapping


def export_to_excel(master_df, factor_ur_dict, missing_factor_regions=None, company_summary_df=None, zone_mapping=None, unmapped_zones=None, variance_df=None):
    """
    Returns a BytesIO Excel file with sheets:
      Master_Clean, Company_Summary (if provided), Variance (if provided),
      Brand_Mapping, Company_Mapping, Individual_Factor, Zone_Mapping (if provided)
    """
    output = BytesIO()

    # Drop the noisy per-row Region_Raw / Is_Zone helper columns from the
    # user-facing Master_Clean sheet, keep everything else.
    display_df = master_df.drop(columns=["Is_Zone"], errors="ignore")

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        display_df.to_excel(writer, sheet_name="Master_Clean", index=False)
        if company_summary_df is not None:
            company_summary_df.to_excel(writer, sheet_name="Company_Summary", index=False)
        if variance_df is not None:
            variance_df.to_excel(writer, sheet_name="Variance", index=False)
        build_brand_mapping_sheet(master_df).to_excel(writer, sheet_name="Brand_Mapping", index=False)
        build_company_mapping_sheet(master_df).to_excel(writer, sheet_name="Company_Mapping", index=False)
        build_individual_factor_sheet(factor_ur_dict).to_excel(writer, sheet_name="Individual_Factor", index=False)

        if zone_mapping:
            rows = [{"Zone": z, "Member_State": s} for z, states in zone_mapping.items() for s in states]
            pd.DataFrame(rows).to_excel(writer, sheet_name="Zone_Mapping", index=False)

        if missing_factor_regions:
            pd.DataFrame(
                [{"State_Zone": s, "Urban_Rural": u} for (s, u) in sorted(missing_factor_regions)]
            ).to_excel(writer, sheet_name="Missing_Factor_Regions", index=False)

        if unmapped_zones:
            pd.DataFrame(
                [{"Zone": z, "Note": "No Zone Mapping entry - not calculated"} for z in sorted(unmapped_zones)]
            ).to_excel(writer, sheet_name="Unmapped_Zones", index=False)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)

    output.seek(0)
    return output
