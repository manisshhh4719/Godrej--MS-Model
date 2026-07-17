import streamlit as st
import pandas as pd

from cleaner import process_all_files, get_sheet_overview
from calculator import (
    add_calculations, build_factor_lookup, DEFAULT_INDIVIDUAL_FACTOR_UR,
    KNOWN_INTENTIONAL_ZERO_FACTORS, build_rollup_coverage,
    build_company_summary, DEFAULT_ZONE_MAPPING, STATE_TO_ZONE, compute_variance,
    smart_guess_companies, find_potential_duplicates,
)
from exporter import export_to_excel
from default_brand_mapping import DEFAULT_BRAND_MAPPING
from default_company_mapping import DEFAULT_COMPANY_MAPPING

st.set_page_config(page_title="Godrej Market Share Model", page_icon=None, layout="wide", initial_sidebar_state="expanded")

KNOWN_FORMATS = ["SHC", "Creme", "Henna", "HBP", "OBL", "VAP", "BP", "KM"]

PAGES = [
    ("upload", "Upload Files"),
    ("region", "Region Selection"),
    ("factor", "Individual Factor"),
    ("zone", "Zone Mapping"),
    ("brand", "Brand Mapping"),
    ("manufacturer", "Manufacturer Mapping"),
    ("run", "Run and Download"),
    ("explorer", "KPI Explorer"),
]
PAGE_ACCENTS = {
    "upload": ("#DCEEFB", "#3E7CB1"),
    "region": ("#FBE4E4", "#C1666B"),
    "factor": ("#FDEBD3", "#C97F2E"),
    "zone": ("#E3F5E1", "#4F9A63"),
    "brand": ("#EAE4F7", "#7B65B5"),
    "manufacturer": ("#FFF3D6", "#B8912B"),
    "run": ("#D8F0EA", "#33897A"),
    "explorer": ("#E8EAF6", "#4A5AA8"),
}

# ============================== GLOBAL STYLE ==============================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }

.stApp { background: #FAF9FC; }

section[data-testid="stSidebar"] {
    background: #FFFFFF;
    border-right: 1px solid #ECEAF3;
}
.sidebar-brand {
    font-family: 'Fraunces', serif;
    font-size: 1.35rem;
    font-weight: 700;
    color: #2E2A45;
    padding: 0.5rem 0 0.25rem 0;
    line-height: 1.25;
}
.sidebar-sub {
    color: #8B889E;
    font-size: 0.82rem;
    padding-bottom: 1.2rem;
    border-bottom: 1px solid #ECEAF3;
    margin-bottom: 0.8rem;
}

.page-header {
    display: flex;
    align-items: center;
    gap: 0.9rem;
    margin-bottom: 1.6rem;
}
.page-header .icon-chip {
    width: 48px; height: 48px;
    border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    font-family: 'Fraunces', serif;
    font-weight: 700;
    font-size: 1.25rem;
    flex-shrink: 0;
}
.page-header h1 {
    font-family: 'Fraunces', serif;
    font-weight: 700;
    font-size: 1.9rem;
    color: #2E2A45;
    margin: 0;
    line-height: 1.15;
}
.page-header p {
    color: #6B6884;
    font-size: 0.95rem;
    margin: 0.15rem 0 0 0;
}

.stat-card {
    border-radius: 16px;
    padding: 1.1rem 1.3rem;
    height: 100%;
}
.stat-card .stat-value {
    font-family: 'Fraunces', serif;
    font-weight: 700;
    font-size: 1.9rem;
    color: #2E2A45;
    line-height: 1.1;
}
.stat-card .stat-label {
    color: #6B6884;
    font-size: 0.82rem;
    margin-top: 0.15rem;
    font-weight: 500;
}

.soft-card {
    background: #FFFFFF;
    border: 1px solid #ECEAF3;
    border-radius: 18px;
    padding: 1.4rem 1.5rem;
    margin-bottom: 1.2rem;
}
.soft-card h3 {
    font-family: 'Fraunces', serif;
    font-size: 1.1rem;
    color: #2E2A45;
    margin: 0 0 0.5rem 0;
}
.soft-note {
    color: #6B6884;
    font-size: 0.88rem;
    line-height: 1.5;
}
.highlight-badge {
    display: inline-block;
    background: #FFE8CC;
    color: #94590E;
    font-size: 0.75rem;
    font-weight: 600;
    padding: 0.15rem 0.55rem;
    border-radius: 999px;
    margin-left: 0.4rem;
}

.stButton > button {
    background: #6C63A6;
    color: white;
    font-weight: 600;
    border: none;
    padding: 0.6rem 1.6rem;
    border-radius: 10px;
    transition: background 0.15s ease;
}
.stButton > button:hover { background: #574E8F; }
.stDownloadButton > button {
    background: #33897A;
    color: white;
    font-weight: 600;
    border: none;
    padding: 0.6rem 1.6rem;
    border-radius: 10px;
}
.stDownloadButton > button:hover { background: #276E62; }

[data-testid="stFileUploaderDropzone"] {
    background: #FAF9FC;
    border-radius: 14px;
}
[data-testid="stDataFrame"], [data-testid="stDataEditor"] {
    border-radius: 12px;
    overflow: hidden;
}
hr { margin: 1.2rem 0; }
</style>
""", unsafe_allow_html=True)


def page_header(page_key, title, subtitle):
    bg, fg = PAGE_ACCENTS[page_key]
    letter = title[0]
    st.markdown(f"""
    <div class="page-header">
        <div class="icon-chip" style="background:{bg}; color:{fg};">{letter}</div>
        <div>
            <h1>{title}</h1>
            <p>{subtitle}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)


def stat_card(label, value, bg):
    st.markdown(f"""
    <div class="stat-card" style="background:{bg};">
        <div class="stat-value">{value}</div>
        <div class="stat-label">{label}</div>
    </div>
    """, unsafe_allow_html=True)


def soft_card_open(title=None):
    html = '<div class="soft-card">'
    if title:
        html += f'<h3>{title}</h3>'
    st.markdown(html, unsafe_allow_html=True)


def soft_card_close():
    st.markdown('</div>', unsafe_allow_html=True)


def is_dirty(df_key):
    orig_key = df_key + "__orig"
    cur = st.session_state.get(df_key)
    orig = st.session_state.get(orig_key)
    if cur is None:
        return False
    if orig is None:
        st.session_state[orig_key] = cur.copy()
        return False
    try:
        return not cur.reset_index(drop=True).equals(orig.reset_index(drop=True))
    except Exception:
        return True


def nav_footer(current_key, df_keys=None):
    changed = any(is_dirty(k) for k in (df_keys or []))
    label = "Save and Next" if changed else "Next"
    idx = [p[0] for p in PAGES].index(current_key)
    if idx >= len(PAGES) - 1:
        return
    st.write("")
    if st.button(label, key=f"nav_next_{current_key}", type="primary"):
        for k in (df_keys or []):
            st.session_state[k + "__orig"] = st.session_state[k].copy()
        st.session_state.current_page = PAGES[idx + 1][0]
        st.rerun()


def get_included_sheets_by_format():
    included = {}
    if "sheet_selection_df" in st.session_state and len(st.session_state.sheet_selection_df):
        for fmt, grp in st.session_state.sheet_selection_df.groupby("Format"):
            included[fmt] = set(grp[grp["Include"]]["Sheet_Name"])
    return included


def get_cached_master_df():
    """Cheap re-parse of the currently uploaded + selected raw files, cached
    until the file set or sheet selection actually changes. Used to detect
    brand names that aren't in Brand Mapping / Manufacturer Mapping yet."""
    file_format_map = st.session_state.file_format_map
    if not file_format_map:
        return None
    included_sheets_by_format = get_included_sheets_by_format()
    sig = (
        tuple(sorted((fmt, f.name, f.size) for fmt, f in file_format_map.items())),
        tuple(sorted((k, tuple(sorted(v))) for k, v in included_sheets_by_format.items())),
    )
    if st.session_state.get("master_df_sig") != sig:
        try:
            master_df, _ = process_all_files(file_format_map, included_sheets_by_format=included_sheets_by_format)
        except Exception as e:
            # A sheet whose layout matches none of the known formats is refused
            # rather than guessed at (correct behaviour - guessing would read the
            # wrong cells silently). But that must never dump a raw traceback on
            # an unrelated page: show what to do about it instead.
            st.error(
                f"Could not read one of the selected sheets, so this page can't be built yet.\n\n"
                f"Details: {e}\n\n"
                f"Go to **Region Selection** and untick that sheet, then come back. "
                f"All India and Zone sheets are never needed - those numbers are always "
                f"rebuilt by summing the states you include."
            )
            st.stop()
        st.session_state.cached_master_df = master_df
        st.session_state.master_df_sig = sig
    return st.session_state.cached_master_df


# ============================== SESSION STATE ==============================
if "factor_df" not in st.session_state:
    st.session_state.factor_df = pd.DataFrame([
        {"State_Zone": s, "Individual_Factor_U": v.get("U"), "Individual_Factor_R": v.get("R")}
        for s, v in DEFAULT_INDIVIDUAL_FACTOR_UR.items()
    ])
    st.session_state["factor_df__orig"] = st.session_state.factor_df.copy()

if "mapping_df" not in st.session_state:
    st.session_state.mapping_df = pd.DataFrame(
        [{"Brand_SKU_Item": k, "Flag": v, "Status": "Verified"} for k, v in DEFAULT_BRAND_MAPPING.items()]
    ).sort_values(["Flag", "Brand_SKU_Item"]).reset_index(drop=True)
    st.session_state["mapping_df__orig"] = st.session_state.mapping_df.copy()

if "company_mapping_df" not in st.session_state:
    st.session_state.company_mapping_df = pd.DataFrame(
        [{"Brand_SKU_Item": k, "Company": v, "Status": "Verified"} for k, v in DEFAULT_COMPANY_MAPPING.items()]
    ).sort_values(["Company", "Brand_SKU_Item"]).reset_index(drop=True)
    st.session_state["company_mapping_df__orig"] = st.session_state.company_mapping_df.copy()

if "zone_mapping_df" not in st.session_state:
    st.session_state.zone_mapping_df = pd.DataFrame(columns=["Zone", "Member_State"])
    st.session_state["zone_mapping_df__orig"] = st.session_state.zone_mapping_df.copy()

if "file_format_map" not in st.session_state:
    st.session_state.file_format_map = {}

if "current_page" not in st.session_state:
    st.session_state.current_page = "upload"


def sync_file_format_map():
    """Rebuild file_format_map from the staged files + whatever is currently
    typed in each Format box.

    This MUST run before the sidebar and the page body render. Streamlit draws
    the sidebar first, and the upload page used to only rebuild this map near
    the bottom of its own body - so the sidebar caption and the 'Files added'
    card were always showing the PREVIOUS run's value. Typing a format and
    hitting enter appeared to do nothing until you edited the box a second
    time and triggered another rerun. Reading the text_input's own session
    state key here (which already holds the latest typed value) fixes that,
    so one enter is enough.
    """
    staged = st.session_state.get("staged_files", {})
    fmap = {}
    for fname, entry in staged.items():
        typed = st.session_state.get(f"fmt_{fname}", entry[1])
        typed = (typed or "").strip()
        entry[1] = typed  # keep the staged copy in sync so it survives reruns
        if typed:
            fmap[typed] = entry[0]
    st.session_state.file_format_map = fmap


sync_file_format_map()

# ============================== SIDEBAR NAV ==============================
with st.sidebar:
    st.markdown('<div class="sidebar-brand">Godrej Market<br/>Share Model</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-sub">Raw Kantar files to verified Master_Clean output</div>', unsafe_allow_html=True)

    for key, label in PAGES:
        is_active = st.session_state.current_page == key
        if st.button(label, key=f"nav_{key}", use_container_width=True,
                     type="primary" if is_active else "secondary"):
            st.session_state.current_page = key
            st.rerun()

    st.markdown("<hr/>", unsafe_allow_html=True)
    n_files = len(st.session_state.file_format_map)
    st.caption(f"{n_files} file(s) added" + (f" - {', '.join(st.session_state.file_format_map.keys())}" if n_files else ""))
    if "output_bytes" in st.session_state:
        st.caption("Last run complete, ready to download")

page = st.session_state.current_page

# ============================== PAGE: UPLOAD ==============================
if page == "upload":
    page_header("upload", "Upload Files", "Upload one raw 68-sheet Excel file per Format/Category.")

    stat_card("Files added", len(st.session_state.file_format_map), PAGE_ACCENTS["upload"][0])

    st.write("")
    soft_card_open()
    st.caption(f"Common formats: {', '.join(KNOWN_FORMATS)} - or type your own if it's not one of these.")
    uploaded_files = st.file_uploader(
        "Upload Excel files (.xlsx)", type=["xlsx"], accept_multiple_files=True,
        label_visibility="collapsed", key="main_file_uploader",
    )

    # Merge newly uploaded files into what's already staged, rather than
    # replacing it. This avoids wiping out previously staged files if this
    # widget briefly returns nothing when the page is revisited.
    if "staged_files" not in st.session_state:
        st.session_state.staged_files = {}  # file.name -> (file_obj, format_text)

    if uploaded_files:
        for f in uploaded_files:
            if f.name not in st.session_state.staged_files:
                st.session_state.staged_files[f.name] = [f, ""]
            else:
                st.session_state.staged_files[f.name][0] = f  # refresh file object reference

    to_remove = None
    for fname, (fobj, fmt_val) in list(st.session_state.staged_files.items()):
        col1, col2, col3 = st.columns([3, 2, 1])
        with col1:
            st.write(f"**{fname}**")
        with col2:
            fmt = st.text_input("Format", value=fmt_val, key=f"fmt_{fname}", label_visibility="collapsed",
                                 placeholder=f"Type a format, e.g. {KNOWN_FORMATS[0]}")
            st.session_state.staged_files[fname][1] = (fmt or "").strip()
        with col3:
            if st.button("Remove", key=f"remove_{fname}"):
                to_remove = fname
    if to_remove:
        del st.session_state.staged_files[to_remove]
        st.rerun()

    # Same sync as the one that already ran before the sidebar - re-run it here
    # so any format typed during THIS run is reflected immediately.
    sync_file_format_map()
    soft_card_close()

    nav_footer("upload")

# ============================== PAGE: REGION SELECTION ==============================
elif page == "region":
    page_header("region", "Region Selection", "Choose which sheets from your raw files should be processed.")

    file_format_map = st.session_state.file_format_map
    sig = tuple(sorted((fmt, f.name, f.size) for fmt, f in file_format_map.items()))
    have_existing = "sheet_selection_df" in st.session_state and len(st.session_state.sheet_selection_df) > 0
    if st.session_state.get("region_sig") != sig and (sig or not have_existing):
        rows = []
        for fmt, f in file_format_map.items():
            for o in get_sheet_overview(f):
                o["Format"] = fmt
                rows.append(o)
        cols = ["Format", "Sheet_Name", "State_Zone", "Urban_Rural", "Is_Zone", "Is_All_India", "Is_AP_Tel", "Include"]
        new_df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
        st.session_state.sheet_selection_df = new_df
        st.session_state["sheet_selection_df__orig"] = new_df.copy()
        st.session_state.region_sig = sig

    sel = st.session_state.get("sheet_selection_df", pd.DataFrame())
    n_total = len(sel)
    n_included = int(sel["Include"].sum()) if n_total else 0

    c1, c2 = st.columns(2)
    with c1: stat_card("Sheets found", n_total, PAGE_ACCENTS["region"][0])
    with c2: stat_card("Sheets selected", n_included, PAGE_ACCENTS["region"][0])

    st.write("")
    if not file_format_map:
        soft_card_open()
        st.markdown('<p class="soft-note">No files added yet. Go to <b>Upload Files</b> first.</p>', unsafe_allow_html=True)
        soft_card_close()
    else:
        soft_card_open("Sheets in your uploaded files")
        st.caption(
            "State sheets are selected by default. AP+Tel, Zone sheets (North, GCPL East, etc) and "
            "All India sheets start unselected. Zones and All India are always recalculated by summing "
            "the states you include, so their own sheets are never used for the numbers. "
            "Tick or untick any row to change what gets processed."
        )
        edited = st.data_editor(
            st.session_state.sheet_selection_df, use_container_width=True, key="sheet_selection_editor",
            disabled=["Format", "Sheet_Name", "State_Zone", "Urban_Rural", "Is_Zone", "Is_All_India", "Is_AP_Tel"],
            height=420,
        )
        st.session_state.sheet_selection_df = edited
        soft_card_close()

    nav_footer("region", df_keys=["sheet_selection_df"])

# ============================== PAGE: INDIVIDUAL FACTOR ==============================
elif page == "factor":
    page_header("factor", "Individual Factor", "Urban and Rural factors per state.")

    stat_card("States loaded", len(st.session_state.factor_df), PAGE_ACCENTS["factor"][0])

    zero_rows = st.session_state.factor_df[
        (st.session_state.factor_df["Individual_Factor_U"] == 0) |
        (st.session_state.factor_df["Individual_Factor_R"] == 0)
    ]
    if len(zero_rows) > 0:
        st.error(
            f"{len(zero_rows)} state(s) have a factor of exactly 0, which will zero out ALL Sales Derived "
            f"and Units Estd for that state/Urban_Rural, with no other warning sign: "
            f"{zero_rows['State_Zone'].tolist()}. A real factor is never truly zero, fix this before running."
        )

    st.write("")
    soft_card_open("Your Individual Factor list")
    st.session_state.factor_df = st.data_editor(
        st.session_state.factor_df, num_rows="dynamic", use_container_width=True, key="factor_editor",
    )
    soft_card_close()

    soft_card_open("Or upload a factor file instead")
    st.caption("Columns: State_Zone, Individual_Factor_U, Individual_Factor_R. Loads into the table above.")
    factor_upload = st.file_uploader("Upload Individual Factor file", type=["csv", "xlsx"], key="factor_upload", label_visibility="collapsed")
    if factor_upload is not None:
        upload_sig = (factor_upload.name, factor_upload.size)
        if st.session_state.get("factor_upload_sig") != upload_sig:
            try:
                up_df = pd.read_csv(factor_upload) if factor_upload.name.endswith(".csv") else pd.read_excel(factor_upload)
                cols = list(up_df.columns)
                rename = {cols[0]: "State_Zone"}
                for c in cols[1:]:
                    cl = str(c).lower()
                    if cl.endswith("_u") or cl == "u" or "urban" in cl:
                        rename[c] = "Individual_Factor_U"
                    elif cl.endswith("_r") or cl == "r" or "rural" in cl:
                        rename[c] = "Individual_Factor_R"
                up_df = up_df.rename(columns=rename)
                needed = ["State_Zone", "Individual_Factor_U", "Individual_Factor_R"]
                if all(c in up_df.columns for c in needed):
                    st.session_state.factor_df = up_df[needed].reset_index(drop=True)
                    st.session_state["factor_df__orig"] = st.session_state.factor_df.copy()
                    st.session_state.factor_upload_sig = upload_sig
                    st.success(f"Loaded {len(up_df)} rows into the table above.")
                    st.rerun()
                else:
                    st.error("Could not detect State/U/R columns. Expected columns like: State_Zone, Individual_Factor_U, Individual_Factor_R.")
            except Exception as e:
                st.error(f"Could not read file: {e}")
    soft_card_close()

    nav_footer("factor", df_keys=["factor_df"])

# ============================== PAGE: ZONE MAPPING ==============================
elif page == "zone":
    page_header("zone", "Zone Mapping", "Zones are calculated by summing their member states.")

    included_states = set()
    if "sheet_selection_df" in st.session_state and len(st.session_state.sheet_selection_df):
        sel = st.session_state.sheet_selection_df
        included_states = set(sel[(sel["Include"]) & (~sel["Is_Zone"])]["State_Zone"].unique())

    zone_sig = tuple(sorted(included_states))
    have_existing_zone = "zone_mapping_df" in st.session_state and len(st.session_state.zone_mapping_df) > 0
    if st.session_state.get("zone_sig") != zone_sig and (zone_sig or not have_existing_zone):
        rows = []
        for state in sorted(included_states):
            zone = STATE_TO_ZONE.get(state)
            if zone:
                rows.append({"Zone": zone, "Member_State": state})
        for gcpl_zone, base_zone in [("GCPL East", "East"), ("GCPL West", "West")]:
            for state in included_states:
                if STATE_TO_ZONE.get(state) == base_zone:
                    rows.append({"Zone": gcpl_zone, "Member_State": state})
        new_df = pd.DataFrame(rows, columns=["Zone", "Member_State"]).drop_duplicates().reset_index(drop=True) if rows else pd.DataFrame(columns=["Zone", "Member_State"])
        st.session_state.zone_mapping_df = new_df
        st.session_state["zone_mapping_df__orig"] = new_df.copy()
        st.session_state.zone_sig = zone_sig

    n_zone_rows = len(st.session_state.zone_mapping_df)
    n_zones_mapped = st.session_state.zone_mapping_df["Zone"].nunique() if n_zone_rows else 0

    c1, c2 = st.columns(2)
    with c1: stat_card("Zones mapped", n_zones_mapped, PAGE_ACCENTS["zone"][0])
    with c2: stat_card("State/Zone rows", n_zone_rows, PAGE_ACCENTS["zone"][0])

    st.write("")
    soft_card_open("Zone to member states")
    st.caption(
        "Built from the states selected on Region Selection. GCPL East/West are suggested equal to "
        "East/West, edit if needed."
    )
    known_states = sorted(st.session_state.factor_df["State_Zone"].dropna().unique().tolist())
    st.session_state.zone_mapping_df = st.data_editor(
        st.session_state.zone_mapping_df, num_rows="dynamic", use_container_width=True, key="zone_mapping_editor",
        column_config={"Member_State": st.column_config.SelectboxColumn("Member_State", options=known_states)},
    )
    soft_card_close()

    soft_card_open("Or upload a Zone Mapping file")
    st.caption("Columns: Zone, Member_State")
    zone_upload = st.file_uploader("Upload Zone Mapping file", type=["csv", "xlsx"], key="zone_mapping_upload", label_visibility="collapsed")
    if zone_upload is not None:
        try:
            zmap_df = pd.read_csv(zone_upload) if zone_upload.name.endswith(".csv") else pd.read_excel(zone_upload)
            st.session_state.zone_mapping_df = zmap_df[["Zone", "Member_State"]]
            st.success(f"Loaded {len(zmap_df)} zone/state rows.")
        except Exception as e:
            st.error(f"Could not read file: {e}")
    soft_card_close()

    nav_footer("zone", df_keys=["zone_mapping_df"])

# ============================== PAGE: BRAND MAPPING ==============================
elif page == "brand":
    page_header("brand", "Brand Mapping", "How each brand is classified: Category, Brand, Sub-brand, or Others.")

    master_df = get_cached_master_df()
    if master_df is not None and len(master_df):
        current_items = set(master_df["Brand_SKU_Item"].unique())
        mapped_items = set(st.session_state.mapping_df["Brand_SKU_Item"])
        new_items = sorted(current_items - mapped_items)
        if new_items:
            item_flags = master_df.drop_duplicates("Brand_SKU_Item").set_index("Brand_SKU_Item")["Flag"]
            new_rows = [{"Brand_SKU_Item": item, "Flag": item_flags.get(item, "Brand"), "Status": "Unverified"} for item in new_items]
            st.session_state.mapping_df = pd.concat([st.session_state.mapping_df, pd.DataFrame(new_rows)], ignore_index=True)
            st.info(f"{len(new_items)} new brand(s) found and added below.")

    counts = st.session_state.mapping_df["Flag"].value_counts()
    n_unverified = int((st.session_state.mapping_df["Status"] == "Unverified").sum())
    c1, c2, c3, c4 = st.columns(4)
    with c1: stat_card("Category", int(counts.get("Category", 0)), PAGE_ACCENTS["brand"][0])
    with c2: stat_card("Brand", int(counts.get("Brand", 0)), PAGE_ACCENTS["brand"][0])
    with c3: stat_card("Sub-brand", int(counts.get("Sub-brand", 0)), PAGE_ACCENTS["brand"][0])
    with c4: stat_card("Unverified", n_unverified, PAGE_ACCENTS["brand"][0])

    st.write("")
    soft_card_open("Classification list")
    filter_options = ["All", "Category", "Brand", "Sub-brand", "Others", "Unverified"]
    filter_choice = st.selectbox("Show", filter_options, index=0, key="brand_filter")
    full_df = st.session_state.mapping_df
    if filter_choice == "All":
        view_df = full_df
    elif filter_choice == "Unverified":
        view_df = full_df[full_df["Status"] == "Unverified"]
    else:
        view_df = full_df[full_df["Flag"] == filter_choice]
    st.caption(f"Showing {len(view_df)} of {len(full_df)} items.")
    edited_view_df = st.data_editor(
        view_df, num_rows="dynamic", use_container_width=True, key=f"mapping_editor_{filter_choice}",
        column_config={
            "Flag": st.column_config.SelectboxColumn("Flag", options=["Category", "Brand", "Sub-brand", "Others"]),
            "Status": st.column_config.SelectboxColumn("Status", options=["Verified", "Unverified"]),
        },
        height=350,
    )
    if filter_choice == "All":
        st.session_state.mapping_df = edited_view_df
    else:
        mask = (full_df["Status"] == "Unverified") if filter_choice == "Unverified" else (full_df["Flag"] == filter_choice)
        untouched = full_df[~mask]
        st.session_state.mapping_df = pd.concat([untouched, edited_view_df], ignore_index=True)
    soft_card_close()

    soft_card_open("Or upload a whole new Brand Mapping file")
    mapping_upload = st.file_uploader("Upload Brand_Mapping file", type=["xlsx", "csv"], key="mapping_upload", label_visibility="collapsed")
    if mapping_upload is not None:
        try:
            map_df = pd.read_csv(mapping_upload) if mapping_upload.name.endswith(".csv") else pd.read_excel(mapping_upload, sheet_name="Brand_Mapping")
            if "Status" not in map_df.columns:
                map_df["Status"] = "Verified"
            st.session_state.mapping_df = map_df[["Brand_SKU_Item", "Flag", "Status"]]
            st.success(f"Loaded {len(map_df)} item mappings.")
        except Exception as e:
            st.error(f"Could not read mapping file: {e}")
    soft_card_close()

    nav_footer("brand", df_keys=["mapping_df"])

# ============================== PAGE: MANUFACTURER MAPPING ==============================
elif page == "manufacturer":
    page_header("manufacturer", "Manufacturer Mapping", "Which parent company each brand belongs to.")

    master_df = get_cached_master_df()
    if master_df is not None and len(master_df):
        current_items = set(master_df[master_df["Flag"] != "Category"]["Brand_SKU_Item"].unique())
        mapped_items = set(st.session_state.company_mapping_df["Brand_SKU_Item"])
        new_items = sorted(current_items - mapped_items)
        if new_items:
            existing_mapping = dict(zip(st.session_state.company_mapping_df["Brand_SKU_Item"], st.session_state.company_mapping_df["Company"]))
            guesses = smart_guess_companies(new_items, existing_mapping)
            new_rows = [{"Brand_SKU_Item": item, "Company": comp, "Status": "Unverified"} for item, (comp, _status) in guesses.items()]
            st.session_state.company_mapping_df = pd.concat([st.session_state.company_mapping_df, pd.DataFrame(new_rows)], ignore_index=True)
            st.info(f"{len(new_items)} new brand(s) found and added below.")

    n_companies = st.session_state.company_mapping_df["Company"].nunique()
    n_company_items = len(st.session_state.company_mapping_df)
    n_unverified = int((st.session_state.company_mapping_df["Status"] == "Unverified").sum())

    c1, c2, c3 = st.columns(3)
    with c1: stat_card("Items mapped", n_company_items, PAGE_ACCENTS["manufacturer"][0])
    with c2: stat_card("Companies", n_companies, PAGE_ACCENTS["manufacturer"][0])
    with c3: stat_card("Unverified", n_unverified, PAGE_ACCENTS["manufacturer"][0])

    st.write("")

    dupes = find_potential_duplicates(st.session_state.company_mapping_df["Brand_SKU_Item"].tolist())
    if dupes:
        soft_card_open("Possible duplicate brand names")
        dup_df = pd.DataFrame(dupes, columns=["Item A", "Item B", "Similarity"])
        dup_df["Similarity"] = (dup_df["Similarity"] * 100).round(0).astype(int).astype(str) + "%"
        st.dataframe(dup_df, use_container_width=True, height=min(250, 45 + 35 * len(dup_df)))
        soft_card_close()

    soft_card_open("Brand to Company")
    filter_options = ["All"] + sorted(st.session_state.company_mapping_df["Company"].unique().tolist()) + ["Unverified"]
    company_filter = st.selectbox("Show", filter_options, index=0, key="company_filter")
    full_company_df = st.session_state.company_mapping_df
    if company_filter == "All":
        view_company_df = full_company_df
    elif company_filter == "Unverified":
        view_company_df = full_company_df[full_company_df["Status"] == "Unverified"]
    else:
        view_company_df = full_company_df[full_company_df["Company"] == company_filter]
    st.caption(f"Showing {len(view_company_df)} of {len(full_company_df)} items.")
    edited_company_view_df = st.data_editor(
        view_company_df, num_rows="dynamic", use_container_width=True, key=f"company_editor_{company_filter}", height=350,
        column_config={"Status": st.column_config.SelectboxColumn("Status", options=["Verified", "Unverified"])},
    )
    if company_filter == "All":
        st.session_state.company_mapping_df = edited_company_view_df
    else:
        mask = (full_company_df["Status"] == "Unverified") if company_filter == "Unverified" else (full_company_df["Company"] == company_filter)
        untouched_company = full_company_df[~mask]
        st.session_state.company_mapping_df = pd.concat([untouched_company, edited_company_view_df], ignore_index=True)
    soft_card_close()

    soft_card_open("Or upload a whole new Manufacturer Mapping file")
    company_mapping_upload = st.file_uploader("Upload Company_Mapping file", type=["xlsx", "csv"], key="company_mapping_upload", label_visibility="collapsed")
    if company_mapping_upload is not None:
        try:
            cmap_df = pd.read_csv(company_mapping_upload) if company_mapping_upload.name.endswith(".csv") else pd.read_excel(company_mapping_upload, sheet_name="Company_Mapping")
            if "Status" not in cmap_df.columns:
                cmap_df["Status"] = "Verified"
            st.session_state.company_mapping_df = cmap_df[["Brand_SKU_Item", "Company", "Status"]]
            st.success(f"Loaded {len(cmap_df)} item mappings.")
        except Exception as e:
            st.error(f"Could not read mapping file: {e}")
    soft_card_close()

    nav_footer("manufacturer", df_keys=["company_mapping_df"])

# ============================== PAGE: RUN & DOWNLOAD ==============================
elif page == "run":
    page_header("run", "Run and Download", "Clean, calculate, and export the verified Master_Clean workbook.")

    file_format_map = st.session_state.file_format_map
    c1, c2, c3 = st.columns(3)
    with c1: stat_card("Files ready", len(file_format_map), PAGE_ACCENTS["run"][0])
    with c2: stat_card("Formats", ", ".join(file_format_map.keys()) if file_format_map else "-", PAGE_ACCENTS["run"][0])
    with c3: stat_card("Status", "Ready" if file_format_map else "Upload files first", PAGE_ACCENTS["run"][0])

    st.write("")

    if not file_format_map:
        soft_card_open()
        st.markdown('<p class="soft-note">No files added yet. Go to <b>Upload Files</b> first.</p>', unsafe_allow_html=True)
        soft_card_close()
    else:
        soft_card_open()
        run_clicked = st.button("Run Pipeline", type="primary")
        soft_card_close()

        if run_clicked:
            flag_overrides = dict(zip(st.session_state.mapping_df["Brand_SKU_Item"], st.session_state.mapping_df["Flag"]))
            company_overrides = dict(zip(st.session_state.company_mapping_df["Brand_SKU_Item"], st.session_state.company_mapping_df["Company"]))
            included_sheets_by_format = get_included_sheets_by_format()

            with st.spinner("Cleaning and combining files..."):
                try:
                    master_df, sku_df = process_all_files(
                        file_format_map, flag_overrides=flag_overrides,
                        company_overrides=company_overrides, included_sheets_by_format=included_sheets_by_format,
                        verbose=False,
                    )
                    st.success(f"Cleaned {len(master_df):,} rows across {len(file_format_map)} format(s): {', '.join(file_format_map.keys())}.")
                except Exception as e:
                    st.error(f"Cleaning failed: {e}")
                    st.stop()

            with st.spinner("Calculating Units Estd, Sales Derived, and Market Share %..."):
                try:
                    factor_ur_dict = {}
                    source_df = st.session_state.factor_df
                    state_col = source_df.columns[0]
                    for _, row in source_df.iterrows():
                        factor_ur_dict[row[state_col]] = {"U": row.get("Individual_Factor_U"), "R": row.get("Individual_Factor_R")}
                    factor_lookup = build_factor_lookup(factor_ur_dict)

                    zone_mapping = {}
                    for _, row in st.session_state.zone_mapping_df.iterrows():
                        zone, state = row.get("Zone"), row.get("Member_State")
                        if pd.isna(zone) or pd.isna(state):
                            continue
                        zone_mapping.setdefault(zone, []).append(state)

                    result_df, missing_regions, unmapped_zones, zero_factor_regions = add_calculations(master_df, factor_lookup=factor_lookup, zone_mapping=zone_mapping)
                    company_summary_df = build_company_summary(result_df)
                    if zero_factor_regions:
                        unexpected_zeros = set(zero_factor_regions) - KNOWN_INTENTIONAL_ZERO_FACTORS
                        intentional_zeros = set(zero_factor_regions) & KNOWN_INTENTIONAL_ZERO_FACTORS
                        if unexpected_zeros:
                            st.error(f"{len(unexpected_zeros)} region(s) have an Individual Factor of exactly 0, which zeroes out ALL Sales Derived/Units Estd for that state: {sorted(unexpected_zeros)}. Go to Individual Factor and fix this before trusting the output.")
                        if intentional_zeros:
                            st.caption(f"Note: {', '.join(f'{s} ({u})' for s, u in sorted(intentional_zeros))} has an Individual Factor of 0 as per the given factor table, so its Sales Derived and Units Estd are 0 by design. This is expected, not an error.")
                    if missing_regions:
                        st.warning(f"{len(missing_regions)} state region(s) had no Individual Factor and used the default (1.0): {sorted(missing_regions)}.")
                    if unmapped_zones:
                        st.warning(f"{len(unmapped_zones)} Zone(s) have no member states present in the data and were NOT calculated: {sorted(unmapped_zones)}.")

                    # Rollup coverage: warn loudly when a rollup is built from
                    # fewer states than exist in the data (e.g. only one state
                    # had a U+R sheet, so "All India (U+R)" is really just that
                    # state). Read-only report - cannot change any number.
                    coverage_df = build_rollup_coverage(result_df, zone_mapping=zone_mapping)
                    gaps = coverage_df[coverage_df["States_Included"] < coverage_df["States_Expected"]]
                    if len(gaps):
                        st.warning(f"{len(gaps)} rollup cut(s) are built from FEWER states than exist in your data. Their numbers are partial, not complete rollups. Details below.")
                        for _, g in gaps.head(8).iterrows():
                            st.warning(f"{g['Rollup']} ({g['Urban_Rural']}) in {g['Format']}: built from {g['States_Included']} of {g['States_Expected']} states. Missing: {g['Missing_States']}.")
                        if len(gaps) > 8:
                            st.warning(f"...and {len(gaps) - 8} more partial rollup cut(s), see the full table below.")
                    with st.expander("Rollup coverage (which states fed All India and each Zone)"):
                        st.dataframe(coverage_df, use_container_width=True, hide_index=True)
                    st.success("Calculations complete.")
                except Exception as e:
                    st.error(f"Calculation failed: {e}")
                    st.stop()

            with st.spinner("Computing Variance (brand total vs sum of its SKUs)..."):
                try:
                    variance_df = compute_variance(master_df, sku_df)
                except Exception as e:
                    st.error(f"Variance calculation failed: {e}")
                    variance_df = None

            with st.spinner("Building Excel output..."):
                output = export_to_excel(
                    result_df, factor_ur_dict, missing_regions, company_summary_df=company_summary_df,
                    zone_mapping=zone_mapping, unmapped_zones=unmapped_zones, variance_df=variance_df,
                )

            st.session_state.output_bytes = output.getvalue()
            st.session_state.result_df = result_df
            st.session_state.company_summary_df = company_summary_df
            st.session_state.variance_df = variance_df

    if "output_bytes" in st.session_state:
        st.write("")
        soft_card_open("Download")
        st.download_button(
            "Download Master_Clean Excel", data=st.session_state.output_bytes,
            file_name="Godrej_Master_Clean.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        soft_card_close()

        soft_card_open("Preview - Master_Clean (first 100 rows)")
        preview_cols = ["Format", "State_Zone", "Urban_Rural", "TG_Segment", "Flag", "Company", "Brand_SKU_Item"]
        available = [c for c in preview_cols if c in st.session_state.result_df.columns]
        st.dataframe(st.session_state.result_df[available].head(100), use_container_width=True)
        soft_card_close()

        if st.session_state.get("company_summary_df") is not None:
            soft_card_open("Preview - Company Summary (GCPL vs competitors)")
            st.dataframe(st.session_state.company_summary_df.head(50), use_container_width=True)
            soft_card_close()

        if st.session_state.get("variance_df") is not None:
            soft_card_open("Preview - Variance (brand total vs sum of its SKUs)")
            st.dataframe(st.session_state.variance_df.head(50), use_container_width=True)
            soft_card_close()

elif page == "explorer":
    page_header("explorer", "KPI Explorer", "Look up any KPI for any company or brand, with filters. Reads directly from the verified pipeline output.")

    if "result_df" not in st.session_state or st.session_state.result_df is None:
        soft_card_open()
        st.markdown('<p class="soft-note">No pipeline output yet. Go to <b>Run and Download</b> and run the pipeline first. The explorer reads the exact same numbers as the Excel output.</p>', unsafe_allow_html=True)
        soft_card_close()
    else:
        result_df = st.session_state.result_df
        company_df = st.session_state.get("company_summary_df")

        # KPIs that can be validly summed across categories. MS% and HH cannot.
        ADDITIVE_KPIS = {"Sales Derived", "Units Estd", "Val", "Vol"}
        PCT_KPIS = {"Value MS%", "Units MS%"}
        COMPANY_KPIS = ["Value MS%", "Units MS%", "Sales Derived", "Units Estd"]
        BRAND_KPIS = ["Value MS%", "Units MS%", "Sales Derived", "Units Estd", "HH", "Vol", "Val"]

        def periods_in(df):
            seen, out = set(), []
            for c in df.columns:
                if "__" in c:
                    p = c.split("__", 1)[1]
                    if p not in seen:
                        seen.add(p)
                        out.append(p)
            return out

        def fmt_val(kpi, v):
            if pd.isna(v):
                return "-"
            if kpi in PCT_KPIS:
                return f"{v * 100:,.2f}%"
            return f"{v:,.2f}"

        soft_card_open("Selection")
        level = st.radio("Level", ["Company", "Brand"], horizontal=True, key="kx_level")

        if level == "Company":
            base = company_df if company_df is not None else pd.DataFrame()
            kpis = COMPANY_KPIS
            entities = sorted([c for c in base["Company"].dropna().unique() if c != "Category Total"]) if len(base) else []
            ent_col = "Company"
        else:
            base = result_df[result_df["Flag"].isin(["Brand", "Others"])].copy()
            kpis = BRAND_KPIS
            entities = sorted(base["Brand_SKU_Item"].dropna().unique()) if len(base) else []
            ent_col = "Brand_SKU_Item"

        if not len(base) or not entities:
            st.warning("No rows available at this level in the current output.")
            soft_card_close()
        else:
            default_ent = 0
            for i, e in enumerate(entities):
                if "GODREJ" in str(e).upper():
                    default_ent = i
                    break

            c1, c2 = st.columns(2)
            with c1:
                entity = st.selectbox(level, entities, index=default_ent, key="kx_entity")
            with c2:
                kpi = st.selectbox("KPI", kpis, key="kx_kpi")

            all_formats = sorted(base["Format"].dropna().unique())
            all_markets = list(base["State_Zone"].dropna().unique())
            market_order = sorted(all_markets, key=lambda m: (m != "All India", m))
            ur_opts = [u for u in ["U+R", "U", "R"] if u in set(base["Urban_Rural"].dropna().unique())]
            tg_opts = sorted(base["TG_Segment"].dropna().unique(), key=lambda t: (t != "TOTAL", t))
            all_periods = periods_in(base)

            f1, f2, f3 = st.columns(3)
            with f1:
                sel_formats = st.multiselect("Category (Format)", all_formats, default=all_formats, key="kx_fmt")
            with f2:
                market = st.selectbox("Market", market_order, key="kx_mkt")
            with f3:
                ur = st.selectbox("Urban / Rural", ur_opts, key="kx_ur")

            f4, f5 = st.columns(2)
            with f4:
                tg = st.selectbox("TG Segment", tg_opts, key="kx_tg") if len(tg_opts) > 1 else tg_opts[0]
            with f5:
                period_choice = st.selectbox("Period", ["All periods (trend)"] + all_periods,
                                             index=len(all_periods), key="kx_period")
            soft_card_close()

            if not sel_formats:
                st.info("Select at least one Category.")
            else:
                sub = base[
                    (base[ent_col] == entity)
                    & (base["Format"].isin(sel_formats))
                    & (base["State_Zone"] == market)
                    & (base["Urban_Rural"] == ur)
                    & (base["TG_Segment"] == tg)
                ]

                if sub.empty:
                    st.warning(f"No data for {entity} with these filters. Try a different Market, U/R cut, or Category.")
                else:
                    multi_cat = sub["Format"].nunique() > 1
                    show_periods = all_periods if period_choice == "All periods (trend)" else [period_choice]
                    show_periods = [p for p in show_periods if f"{kpi}__{p}" in sub.columns]

                    if not show_periods:
                        st.warning(f"{kpi} is not available for the selected period in this data.")
                    else:
                        st.write("")
                        if multi_cat and kpi not in ADDITIVE_KPIS:
                            st.info(f"{kpi} cannot be added across categories (the total would be meaningless), so it is shown per category below.")

                        if len(show_periods) == 1:
                            p = show_periods[0]
                            col = f"{kpi}__{p}"
                            if multi_cat:
                                rows = sub.groupby("Format", as_index=False)[col].sum(min_count=1)
                                if kpi in ADDITIVE_KPIS:
                                    stat_card(f"{entity} | {kpi} | {p} | Total across {len(rows)} categories",
                                              fmt_val(kpi, rows[col].sum()), PAGE_ACCENTS["explorer"][0])
                                    st.write("")
                                disp = rows.rename(columns={col: kpi, "Format": "Category"}).copy()
                                disp[kpi] = disp[kpi].map(lambda v: fmt_val(kpi, v))
                                st.dataframe(disp, use_container_width=True, hide_index=True)
                            else:
                                val = sub[col].sum(min_count=1) if kpi in ADDITIVE_KPIS else sub[col].iloc[0]
                                stat_card(f"{entity} | {kpi} | {p} | {market} ({ur}, {tg})",
                                          fmt_val(kpi, val), PAGE_ACCENTS["explorer"][0])
                        else:
                            recs = []
                            group_cats = multi_cat
                            for p in show_periods:
                                col = f"{kpi}__{p}"
                                if group_cats:
                                    for fmt_name, grp in sub.groupby("Format"):
                                        v = grp[col].sum(min_count=1) if kpi in ADDITIVE_KPIS else grp[col].iloc[0]
                                        recs.append({"Period": p, "Category": fmt_name, kpi: v})
                                else:
                                    v = sub[col].sum(min_count=1) if kpi in ADDITIVE_KPIS else sub[col].iloc[0]
                                    recs.append({"Period": p, kpi: v})
                            trend = pd.DataFrame(recs)

                            plot_df = trend.copy()
                            if kpi in PCT_KPIS:
                                plot_df[kpi] = plot_df[kpi] * 100
                            if group_cats:
                                chart_df = plot_df.pivot(index="Period", columns="Category", values=kpi).reindex(show_periods)
                            else:
                                chart_df = plot_df.set_index("Period")[[kpi]].reindex(show_periods)
                            st.line_chart(chart_df)

                            disp = trend.copy()
                            disp[kpi] = disp[kpi].map(lambda v: fmt_val(kpi, v))
                            st.dataframe(disp, use_container_width=True, hide_index=True)

                        st.caption("These numbers come from the same verified pipeline output as the downloaded Excel, no separate calculation.")
