import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from pathlib import Path

try:
    from streamlit_pdf_viewer import pdf_viewer
except ImportError:  # The app still starts if the optional viewer is unavailable.
    pdf_viewer = None

# ==========================================
# 1. CORE PALETTE & DICTIONARY LOOKUPS
# ==========================================
PRIMARY_COLORS = {"Electric Blue": "#001391", "Serene Blue": "#85C8FF", "Midnight": "#060E46"}
# BBVA digital palette (used for controls and chart accents).
BBVA_NAVY, BBVA_BLUE, BBVA_AZURE = "#072146", "#004481", "#2DCCCD"
ACCENT_COLORS = ["#88E783", "#FFB56B", "#FFE761", "#8BE1E9", "#9694FF"]
GREYS = {"Sand": "#F7F8F8", "Grey-1": "#E2E6EA", "Grey-2": "#CAD1D8", "Grey-3": "#ADB8C2", "Grey-4": "#46536D"}
ASSESSMENT_COLORS = {"Liberalising": "#1E5631", "Distortive": "#B22222"}

CPC_SECTIONS = {
    "0": "Agriculture, forestry & fishery products",
    "1": "Ores and minerals; Electricity, gas & water",
    "2": "Food, beverages & tobacco; Textiles, apparel & leather",
    "3": "Other transportable goods",
    "4": "Metal products; Machinery & equipment",
    "5": "Construction",
    "6": "Trade; Hospitality; Transport; Distribution services",
    "7": "Financial; Real estate; Leasing services",
    "8": "Business & production services",
    "9": "Community, social & personal services"
}

CPC_PRODUCTS_2D = {
    "01": "Agriculture & Horticulture", "02": "Live Animals", "03": "Forestry", "04": "Fish Products",
    "11": "Coal & Lignite", "12": "Crude Petroleum", "13": "Metal Ores", "14": "Stone & Sand",
    "21": "Meat, Fish & Dairy", "22": "Grain Mill Products", "23": "Beverages & Tobacco", "24": "Textiles & Apparel",
    "31": "Wood & Paper", "32": "Petroleum Products", "33": "Chemicals", "34": "Basic Metals",
    "41": "Fabricated Metals", "42": "Machinery & Equipment", "43": "Office & Computing", "44": "Electrical Machinery"
}

HS_PRODUCTS_2D = {
    "01": "Live Animals", "02": "Meat", "03": "Fish", "04": "Dairy", "05": "Animal Derived",
    "06": "Live Trees", "07": "Vegetables", "08": "Fruits & Nuts", "09": "Coffee & Tea", "10": "Cereals",
    "25": "Salt & Sulfur", "26": "Ores", "27": "Mineral Fuels", "28": "Inorganic Chemicals", "29": "Organic Chemicals",
    "30": "Pharmaceuticals", "39": "Plastics", "40": "Rubber", "72": "Iron & Steel", "84": "Nuclear Reactors & Machinery",
    "85": "Electrical Machinery", "87": "Vehicles", "88": "Aircraft", "90": "Optical & Medical Instruments"
}

HS_PRODUCTS_1D = {
    "0": "Live animals & animal products", "1": "Vegetable products",
    "2": "Foodstuffs", "3": "Mineral products", "4": "Chemicals",
    "5": "Plastics, rubber & leather", "6": "Textiles & apparel",
    "7": "Stone, glass & precious metals", "8": "Base metals",
    "9": "Machinery, transport & other goods",
}

HS_SECTIONS = {
    ("01", "02", "03", "04", "05"): "Live animals; Animal products",
    tuple(str(i).zfill(2) for i in range(6, 15)): "Vegetable products",
    tuple(str(i).zfill(2) for i in range(15, 25)): "Food, beverages & tobacco",
    ("25", "26", "27"): "Mineral products",
    tuple(str(i).zfill(2) for i in range(28, 39)): "Chemicals",
    ("39", "40"): "Plastics & rubber",
    ("44", "45", "46", "47", "48", "49"): "Wood",
    tuple([str(i).zfill(2) for i in range(41, 44)] + [str(i).zfill(2) for i in range(50, 68)]): "Textiles & apparel; Leather",
    ("68", "69", "70", "71"): "Stone, glass & precious metals",
    tuple(str(i).zfill(2) for i in range(72, 84)): "Base metals",
    ("84", "85"): "Machinery & equipment",
    ("86", "87", "88", "89"): "Vehicles & transport equipment",
    ("90", "91", "92", "94", "95", "96"): "Miscellaneous (Furniture, Optical, Musical, etc.)",
    ("93",): "Arms & ammunition", ("97",): "Art",
}

HS_SECTION_BY_CODE = {code: label for codes, label in HS_SECTIONS.items() for code in codes}
HS_SECTION_CODES = {label: set(codes) for codes, label in HS_SECTIONS.items()}

COUNTRY_GROUPS = {
    "G-7": ["United States of America", "United Kingdom", "Germany", "France", "Italy", "Japan", "Canada"],
    "G-20": ["United States of America", "United Kingdom", "Germany", "France", "Italy", "Japan", "Canada", "Argentina", "Australia", "Brazil", "China", "India", "Indonesia", "Mexico", "Saudi Arabia", "South Africa", "South Korea", "Turkey", "Russia"],
    "EU-27": ["Austria", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Czechia", "Denmark", "Estonia", "Finland", "France", "Germany", "Greece", "Hungary", "Ireland", "Italy", "Latvia", "Lithuania", "Luxembourg", "Malta", "Netherlands", "Poland", "Portugal", "Romania", "Slovakia", "Slovenia", "Spain", "Sweden"],
    "Europe": ["Austria", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Czechia", "Denmark", "Estonia", "Finland", "France", "Germany", "Greece", "Hungary", "Ireland", "Italy", "Latvia", "Lithuania", "Luxembourg", "Malta", "Netherlands", "Poland", "Portugal", "Romania", "Slovakia", "Slovenia", "Spain", "Sweden", "United Kingdom", "Switzerland", "Norway"],
    "North America": ["United States of America", "Canada", "Mexico"],
    "South America": ["Argentina", "Brazil", "Chile", "Colombia", "Peru", "Uruguay", "Venezuela"],
    "LatAm": ["Argentina", "Brazil", "Chile", "Colombia", "Mexico", "Peru", "Uruguay", "Venezuela"],
    "Asia": ["China", "Japan", "India", "South Korea", "Indonesia", "Saudi Arabia", "Turkey", "Singapore"],
    "Middle East": ["Saudi Arabia", "Turkey", "United Arab Emirates", "Israel", "Qatar"],
    "Africa": ["South Africa", "Nigeria", "Egypt", "Kenya", "Morocco"],
    "Oceania": ["Australia", "New Zealand"]
}

SECTOR_COLS = ["Sector: Low Carbon Technology", "Sector: Dual-Use Products", "Sector: Critical Minerals", "Sector: Advanced Technology Products", "Sector: Medical Products", "Sector: Chemicals", "Sector: Includes IT or Digital Services"]
MOTIVE_COLS = ["Motive: National Security or Geopolitical Concern", "Motive: Resilience/Security of Supply (Non-Food)", "Motive: Strategic Competitiveness", "Motive: Climate Change Mitigation", "Motive: Digital Transformation"]
POLICY_COLS = ["Is Export Policy", "Is Import Policy", "Is Trade Defence", "Is Subsidy", "Is Export Incentive", "Is FDI Policy", "Is Procurement Policy", "Is Localisation Policy", "Is Other Policy"]

# ==========================================
# 2. SOURCE DATA CLEANING & RE-APPORTIONMENT
# ==========================================
@st.cache_data
def load_source_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df["Announcement Date"] = pd.to_datetime(df["Announcement Date"], errors="coerce")
    for date_col in ["Implementation Date", "Removal Date"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["Announcement Date"])
    df = df[df["Levels of Policy Intervention"] != "Firm-specific"]
    
    df["Trade Covered (USD Million)"] = pd.to_numeric(df["Trade Covered (USD Million)"], errors='coerce').fillna(0)
    df["Size of Subsidy (USD Million)"] = pd.to_numeric(df["Size of Subsidy (USD Million)"], errors='coerce').fillna(0)
    df["Total_USD_Value"] = df["Trade Covered (USD Million)"] + df["Size of Subsidy (USD Million)"]
    
    df["Level of Government Implementation"] = df["Level of Government Implementation"].replace({
        "IFI": "Independent Fiscal Institutions (IFI)", "NFI": "National Framework Implementations (NFI)"
    })
    
    for col in SECTOR_COLS + MOTIVE_COLS + POLICY_COLS:
        if col in df.columns:
            df[col] = df[col].astype(str).str.upper() == "TRUE"
            
    df["Initial Assessment"] = df["Initial Assessment (Change Relative to 1 Jan 2009)"].astype(str).str.capitalize()
    df["Affected List"] = df["Affected Jurisdiction"].astype(str).apply(lambda x: [i.strip() for i in x.split(",") if i.strip()])
    return df

def get_dynamic_palette(categories, category_type):
    if category_type == "Assessment Type":
        return [ASSESSMENT_COLORS.get(c, GREYS["Grey-3"]) for c in categories]
    base_corporate = [PRIMARY_COLORS["Electric Blue"], PRIMARY_COLORS["Serene Blue"]] + ACCENT_COLORS
    # "Other transportable goods (3)" is a proper CPC section, not an unclassified residual category, so it must retain a distinct colour.
    is_product_view = category_type in {"Product (CPC v2.1 Sectors)", "Product (1-digit HS 2022)"}
    primary_cats = [
        c for c in categories
        if not (str(c) == "Others" if is_product_view else str(c).startswith("Other"))
    ]
    
    # Keep the app self-contained: this used to call sns.color_palette without
    # importing seaborn, which caused the dashboard to fail as soon as a chart
    # had more categories than the corporate palette.
    extra_count = max(0, len(primary_cats) - len(base_corporate))
    extra_colours = [
        f"hsl({round(360 * i / max(extra_count, 1))}, 58%, 48%)"
        for i in range(extra_count)
    ]
    color_pool = base_corporate + extra_colours
    palette_map = {}
    idx = 0
    for cat in categories:
        if str(cat).startswith("Other"):
            palette_map[cat] = GREYS["Grey-3"]
        else:
            palette_map[cat] = color_pool[idx]
            idx += 1
    return [palette_map[c] for c in categories]

# ==========================================
# 3. BOOLEAN QUERY & FILTER PIPELINE
# ==========================================
def evaluate_boolean_query(title_text, query_str):
    if not query_str or query_str.strip() == "":
        return True
    target_lower = str(title_text).lower()
    pattern = re.compile(r'(\bAND\b|\bOR\b|\(|\))', re.IGNORECASE)
    parts = pattern.split(query_str)
    
    new_parts = []
    for part in parts:
        if part is None: continue
        up_part = part.strip().upper()
        if up_part in ["AND", "OR"]:
            new_parts.append(up_part.lower())
        elif up_part in ["(", ")"]:
            new_parts.append(up_part)
        else:
            term = part.strip().lower()
            if term == "": continue
            # Match complete words/phrases.  In particular, searching for
            # ``AI`` must not match the ``ai`` inside words such as "certain".
            is_match = "True" if re.search(
                rf"(?<!\w){re.escape(term)}(?!\w)", target_lower,
                flags=re.IGNORECASE,
            ) else "False"
            new_parts.append(is_match)
            
    try:
        return eval(" ".join(new_parts), {"__builtins__": None}, {})
    except:
        return False

def execute_filter_pipeline(df, config):
    df_out = df.copy()
    if len(config.get("dates", [])) == 2:
        df_out = df_out[(df_out["Announcement Date"] >= pd.to_datetime(config["dates"][0])) & 
                        (df_out["Announcement Date"] <= pd.to_datetime(config["dates"][1]))]
    for date_col, config_key in [("Implementation Date", "implementation_dates"), ("Removal Date", "removal_dates")]:
        if date_col in df_out.columns and len(config.get(config_key, [])) == 2:
            df_out = df_out[
                df_out[date_col].isna() |
                df_out[date_col].between(
                    pd.to_datetime(config[config_key][0]),
                    pd.to_datetime(config[config_key][1]),
                    inclusive="both",
                )
            ]
    if config.get("gov_level"):
        df_out = df_out[df_out["Level of Government Implementation"].isin(config["gov_level"])]
    if config.get("trade_flow"):
        df_out = df_out[df_out["Affected Trade Flow"].isin(config["trade_flow"])]
    if config.get("assessments"):
        df_out = df_out[df_out["Initial Assessment"].isin(config["assessments"])]
        
    if config.get("imp_jurisdiction"):
        resolved_imp = set()
        for item in config["imp_jurisdiction"]:
            clean = item.replace("Group: ", "")
            resolved_imp.update(COUNTRY_GROUPS.get(clean, [clean]))
        df_out = df_out[df_out["Implementing Jurisdiction"].isin(list(resolved_imp))]
        
    if config.get("aff_jurisdiction"):
        resolved_aff = set()
        for item in config["aff_jurisdiction"]:
            clean = item.replace("Group: ", "")
            resolved_aff.update(COUNTRY_GROUPS.get(clean, [clean]))
        df_out = df_out[df_out["Affected List"].apply(lambda x: any(i in resolved_aff for i in x))]
        
    if config.get("hs_2d"):
        codes = set().union(*(HS_SECTION_CODES.get(item, set()) for item in config["hs_2d"]))
        df_out = df_out[df_out["Product: HS 6-digit (2022)"].apply(lambda x: any(c in [t.strip()[:2].zfill(2) for t in str(x).split(",") if t.strip()] for c in codes))]
    if config.get("cpc_2d"):
        codes = [item.split("(")[1].replace(")", "").strip() for item in config["cpc_2d"]]
        df_out = df_out[df_out["Sector: CPC 3-digit (v2.1)"].apply(lambda x: any(c == t.strip()[:1] for t in str(x).split(",") if t.strip() for c in codes))]
    if config.get("policies"):
        df_out = df_out[df_out[config["policies"]].any(axis=1)]
    if config.get("sectors"):
        if "Others" in config["sectors"]:
            df_out = df_out[~df_out[SECTOR_COLS].any(axis=1)]
        else:
            df_out = df_out[df_out[config["sectors"]].any(axis=1)]
    if config.get("motives"):
        if "Others" in config["motives"]:
            df_out = df_out[~df_out[MOTIVE_COLS].any(axis=1)]
        else:
            df_out = df_out[df_out[config["motives"]].any(axis=1)]
            
    if config.get("keyword_search"):
        df_out = df_out[df_out["Title"].apply(lambda x: evaluate_boolean_query(x, config["keyword_search"]))]
        
    return df_out

def apply_fractional_allocation(df, col_type):
    df_temp = df.copy()
    if col_type == "Assessment Type":
        df_temp["Active_Categories"] = df_temp["Initial Assessment"].apply(lambda x: [x] if x in ["Liberalising", "Distortive"] else ["Other Assessments"])
        df_temp["Denominator"] = 1.0
    elif col_type in ["Product (CPC v2.1 Sectors)", "Product (CPC v2.1 Sections)", "Product (1-digit HS 2022)"]:
        target_col = "Sector: CPC 3-digit (v2.1)" if col_type in ["Product (CPC v2.1 Sectors)", "Product (CPC v2.1 Sections)"] else "Product: HS 6-digit (2022)"
        
        # DEFINED INNER FUNCTION FOR SYSTEM ALLOCATIONS
        def split_codes(val):
            val = str(val).strip()
            if val.upper() in ["NAN", "NONE", ""]: return [f"Other {col_type}"]
            tokens = list(set([t.strip() for t in val.split(",") if t.strip()]))
            if col_type in ["Product (CPC v2.1 Sectors)", "Product (CPC v2.1 Sections)"]:
                return list(set([
                    f"{CPC_SECTIONS.get(t[:1], 'Other Sections')} ({t[:1]})" for t in tokens
                ]))
            return list(set([
                HS_SECTION_BY_CODE.get(t[:2].zfill(2), "Other HS products") for t in tokens
            ]))
            
        df_temp["Active_Categories"] = df_temp[target_col].apply(split_codes)
        df_temp["Denominator"] = df_temp["Active_Categories"].apply(len)
    else:
        cols = SECTOR_COLS if col_type == "Sector" else MOTIVE_COLS if col_type == "Motive" else POLICY_COLS
        lbl = "Sectors" if col_type == "Sector" else "Motives" if col_type == "Motive" else "Policies"
        df_temp["True_Count"] = df_temp[cols].sum(axis=1)
        df_temp["Active_Categories"] = df_temp.apply(lambda r: [c.split(": ")[-1].replace("Is ", "") for c in cols if r[c]] or [f"Other {lbl}"], axis=1)
        df_temp["Denominator"] = df_temp["True_Count"].replace(0, 1)

    df_temp["Allocated_Combined_USD"] = df_temp["Total_USD_Value"] / df_temp["Denominator"]
    df_temp["Allocated_Subsidy_USD"] = df_temp["Size of Subsidy (USD Million)"] / df_temp["Denominator"]
    df_temp["Allocated_Trade_USD"] = df_temp["Trade Covered (USD Million)"] / df_temp["Denominator"]
    df_temp["Allocated_Count"] = 1.0 / df_temp["Denominator"]
    return df_temp.explode("Active_Categories")

# ==========================================
# 5. INLINE FILTER SELECTION MAPPING BUILDER
# ==========================================
def render_inline_filters(df_source, key_prefix, master_ref=None, compact=False, include_title=True):
    groups_list = [f"Group: {k}" for k in COUNTRY_GROUPS.keys()]
    all_imp = groups_list + sorted(df_source["Implementing Jurisdiction"].dropna().unique().tolist())
    all_aff = groups_list + sorted(list(set(x for l in df_source["Affected List"].dropna() for x in l)))
    all_gov = ["Independent Fiscal Institutions (IFI)", "National Framework Implementations (NFI)"] + sorted([x for x in df_source["Level of Government Implementation"].dropna().unique().tolist() if x not in ["Independent Fiscal Institutions (IFI)", "National Framework Implementations (NFI)"]])
    all_flow = sorted(df_source["Affected Trade Flow"].dropna().unique().tolist())
    
    hs_opts = list(dict.fromkeys(HS_SECTIONS.values()))
    cpc_opts = [f"{v} ({k})" for k, v in CPC_SECTIONS.items()]

    def get_fallback(field, default):
        return master_ref[field] if master_ref and field in master_ref else default

    def date_bounds(column):
        if column not in df_source.columns:
            return [df_source["Announcement Date"].min().date(), df_source["Announcement Date"].max().date()]
        values = pd.to_datetime(df_source[column], errors="coerce").dropna()
        return [values.min().date(), values.max().date()] if not values.empty else date_bounds("Announcement Date")

    chart_title = st.text_input("Chart title", get_fallback("title", ""), key=f"{key_prefix}_title") if include_title else ""
    dt = st.date_input("Announcement Date", get_fallback("dates", [df_source["Announcement Date"].min(), df_source["Announcement Date"].max()]), key=f"{key_prefix}_dt")
    imp = st.multiselect("Implementing Jurisdictions", all_imp, default=get_fallback("imp_jurisdiction", []), key=f"{key_prefix}_imp")
    aff = st.multiselect("Affected Jurisdictions", all_aff, default=get_fallback("aff_jurisdiction", []), key=f"{key_prefix}_aff")
    kw = st.text_input(
        "Keyword Search", get_fallback("keyword_search", ""), key=f"{key_prefix}_kw",
        help="Search for interventions with a title matching your query. Use parentheses to group terms and AND/OR to combine them. Example: (AI OR artificial intelligence) AND (chip OR semiconductor). Search is case-insensitive and matches complete words."
    )

    advanced = st.expander("More filters", expanded=not compact)
    with advanced:
        implementation_dates = st.date_input("Implementation Date", get_fallback("implementation_dates", date_bounds("Implementation Date")), key=f"{key_prefix}_implementation_dates", help="Select the implementation-date range to include.")
        removal_dates = st.date_input("Removal Date", get_fallback("removal_dates", date_bounds("Removal Date")), key=f"{key_prefix}_removal_dates", help="Select the removal-date range to include.")
        gov = st.multiselect("Government Level", all_gov, default=get_fallback("gov_level", []), key=f"{key_prefix}_gov")
        flow = st.multiselect("Trade Flow", all_flow, default=get_fallback("trade_flow", []), key=f"{key_prefix}_flow")
        assess = st.multiselect("Assessment", ["Liberalising", "Distortive"], default=get_fallback("assessments", []), key=f"{key_prefix}_assess")
        hs2d = st.multiselect("Product (1-digit HS 2022)", hs_opts, default=get_fallback("hs_2d", []), key=f"{key_prefix}_hs2d")
        cpc2d = st.multiselect("Product (CPC v2.1 Sectors)", cpc_opts, default=get_fallback("cpc_2d", []), key=f"{key_prefix}_cpc2d")
        pols = st.multiselect("Policy Instrument", POLICY_COLS, default=get_fallback("policies", []), key=f"{key_prefix}_pols")
        secs = st.multiselect("Sector", SECTOR_COLS + ["Others"], default=get_fallback("sectors", []), key=f"{key_prefix}_secs")
        mots = st.multiselect("Motive", MOTIVE_COLS + ["Others"], default=get_fallback("motives", []), key=f"{key_prefix}_mots")

    return {
        "keyword_search": kw, "title": chart_title, "dates": dt,
        "implementation_dates": implementation_dates, "removal_dates": removal_dates,
        "imp_jurisdiction": imp, "aff_jurisdiction": aff,
        "gov_level": gov, "trade_flow": flow, "assessments": assess, "hs_2d": hs2d, "cpc_2d": cpc2d,
        "policies": pols, "sectors": secs, "motives": mots
    }

def fill_missing_with_master(child_cfg, master_cfg):
    effective = {}
    for k in master_cfg.keys():
        if k == "dates":
            effective[k] = child_cfg[k] if len(child_cfg[k]) == 2 else master_cfg[k]
        elif k == "keyword_search":
            effective[k] = child_cfg[k] if child_cfg[k].strip() != "" else master_cfg[k]
        else:
            effective[k] = child_cfg[k] if child_cfg[k] else master_cfg[k]
    return effective

def saved_override_config(key_prefix):
    """Return a previously edited chart override without rendering its form."""
    fields = {
        "keyword_search": "kw", "title": "title", "dates": "dt", "implementation_dates": "implementation_dates", "removal_dates": "removal_dates", "imp_jurisdiction": "imp",
        "aff_jurisdiction": "aff", "gov_level": "gov", "trade_flow": "flow",
        "assessments": "assess", "hs_2d": "hs2d", "cpc_2d": "cpc2d",
        "policies": "pols", "sectors": "secs", "motives": "mots",
    }
    return {
        field: st.session_state.get(f"{key_prefix}_{suffix}", "" if field == "keyword_search" else [])
        for field, suffix in fields.items()
    }

# ==========================================
# 6. STREAMLIT APP FRAMEWORK WORKSPACE
# ==========================================
st.set_page_config(page_title="NIPO Industrial Policy Explorer", layout="wide")
st.markdown("""
<style>
    .stApp { background: #F7F8F8; }
    .stMainBlockContainer, .block-container { max-width: none; padding: 2.5rem 4rem 3rem; }
    .stAppHeader { display: none; }
    div[data-testid="stHorizontalBlock"] { align-items: flex-start; gap: 1.5rem; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        background: white; border: 1px solid #E2E6EA; border-radius: 10px;
        padding: 1.5rem; box-shadow: 0 4px 8px rgba(0,0,0,.05);
    }
    button[kind="primary"] { background: #072146 !important; border-color: #072146 !important; color: white !important; }
    button[kind="primary"]:hover { background: #004481 !important; border-color: #004481 !important; }
    div[data-baseweb="select"] > div:focus-within { border-color: #004481 !important; box-shadow: 0 0 0 1px #004481 !important; }
    div[data-testid="stExpander"] details { border: 0; }
    div[data-testid="stExpander"] summary p { font-weight: 600; }
    [data-testid="stTabs"] button { font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.title("BBVA Research | Industrial Policy Explorer")
st.markdown("Explore and analyze industrial policy interventions worldwide. Built by [BBVA Research](https://www.bbvaresearch.com/) based on the [Global Trade Alert's New Industrial Policy Observatory (GTA-NIPO)](https://globaltradealert.org/reports/new-industrial-policy-observatory-nipo) database.")

default_source = Path(__file__).with_name("GTA NIPO - January 2026.xlsx")
uploaded_file = st.file_uploader("Please upload the NIPO database (XLSX file).", type="xlsx")
source_file = uploaded_file if uploaded_file is not None else default_source

if uploaded_file is not None or default_source.exists():
    raw_df = load_source_data(source_file)
    tab_inspect, tab_viz, tab_methodology = st.tabs(["🔎 Data inspection", "📊 Visualization", "❓ Methodology"])

    # ------------------------------------------
    # DATA INSPECTOR WORKSPACE TAB
    # ------------------------------------------
    with tab_inspect:
        filter_col, plot_col = st.columns([1, 3])
        with filter_col:
            st.markdown("### ⚙️ Configure the output table.")
            st.caption("Choose the interventions to include in the output table.")
            inspector_config = render_inline_filters(raw_df, "inspector", compact=True, include_title=False)
            trigger_inspect = st.button("Generate Table", type="primary", use_container_width=True)
            
        with plot_col:
            st.markdown("### ⭐ Results")
            if trigger_inspect:
                ins_df = execute_filter_pipeline(raw_df, inspector_config)
                st.metric("Matching interventions", f"{len(ins_df):,}")
                drop_fields = ["NEW", "Entry ID", "Was First Reported Before This Inventory Month?", "Initial Assessment (Change Relative to 1 Jan 2009)", "Affected List"]
                display_df = ins_df.drop(columns=[c for c in drop_fields if c in ins_df.columns], errors="ignore")
                st.dataframe(display_df, use_container_width=True, hide_index=True)
                st.caption("Click on the button in the top right corner to download the selected data.")
            else:
                st.info("Adjust the menu choices on the left column pane and select 'Generate Table'.")

    # ------------------------------------------
    # MATRIX DASHBOARD GRID VISUALIZATION TAB
    # ------------------------------------------
    with tab_viz:
        filter_col, plot_col = st.columns([1, 3])
        with filter_col:
            st.markdown("### 1️⃣ Configure the general figure settings.")
            st.caption("Set the general figure settings first.")
            disaggregation = st.selectbox("Split series by", [
                "Sector", "Motive", "Policy Instrument", "Assessment Type",
                "Product (CPC v2.1 Sectors)", "Product (1-digit HS 2022)",
            ])
            freq_choice = st.selectbox("Time frequency", ["Daily", "Monthly", "Quarterly", "Yearly"], index=3)
            metric_choice = st.selectbox("Measure", ["Policy Count", "Subsidy USD Amount", "Trade Covered USD Amount", "Combined USD Amount"])
            smoothing = st.slider("Smoothing (periods)", min_value=1, max_value=100, value=1, help="A value of 1 leaves the series unchanged.")

            st.markdown("### 2️⃣ Customize the subplots.")
            st.caption("Now configure the individual subplots to be displayed.")
            chart_to_customize = st.selectbox("Chart to customize", ["Chart 1", "Chart 2", "Chart 3", "Chart 4"], help="Configure one chart at a time. New charts inherit Chart 1's settings by default.")
            chart_number = int(chart_to_customize.split()[-1])
            saved_configs = st.session_state.get("saved_subplot_configs", {})
            override_prefix = {"Chart 1": "v_p1", "Chart 2": "v_p2", "Chart 3": "v_p3", "Chart 4": "v_p4"}[chart_to_customize]
            # Saved child settings take precedence over Chart 1 inheritance.
            # A child inherits Chart 1 only until it has been saved once.
            child_master = saved_configs.get(chart_number) or saved_configs.get(1) or saved_override_config("v_p1")
            selected_override = render_inline_filters(raw_df, override_prefix, master_ref=None if chart_number == 1 else child_master, compact=True)
            st.caption("Save each chart when it is ready. Charts 2–4 begin with Chart 1's settings, which you can then change.")
            save_chart = st.button("Save chart", type="primary", use_container_width=True)

        with plot_col:
            st.markdown("### ⭐ Results")
            freq_code = {"Daily": "D", "Monthly": "M", "Quarterly": "Q", "Yearly": "Y"}[freq_choice]
            metric_col = {"Policy Count": "Allocated_Count", "Subsidy USD Amount": "Allocated_Subsidy_USD", "Trade Covered USD Amount": "Allocated_Trade_USD", "Combined USD Amount": "Allocated_Combined_USD"}[metric_choice]
            
        if "saved_subplot_configs" not in st.session_state:
            st.session_state.saved_subplot_configs = {}

        p1_config = selected_override if chart_number == 1 else st.session_state.get("saved_subplot_configs", {}).get(1, saved_override_config("v_p1"))
        p1_config = fill_missing_with_master(p1_config, p1_config)
        selected_effective = (
            p1_config if chart_number == 1
            else fill_missing_with_master(selected_override, p1_config)
        )

        if save_chart:
            # Store the fully resolved settings. This both makes inheritance
            # reliable for every child chart and prevents later form changes
            # from altering a chart that has already been saved.
            st.session_state.saved_subplot_configs[chart_number] = selected_effective

        saved_configs = st.session_state.saved_subplot_configs
        chart_numbers = []
        for number in range(1, 5):
            if number in saved_configs:
                chart_numbers.append(number)
            else:
                break

        if chart_numbers:
            configs = [saved_configs[number] for number in chart_numbers]
            titles = [cfg.get("title") or f"Chart {number}" for number, cfg in zip(chart_numbers, configs)]
            chart_count = len(configs)
            rows, cols = (1, 1) if chart_count == 1 else ((1, 2) if chart_count == 2 else (2, 2))
            fig = make_subplots(rows=rows, cols=cols, subplot_titles=titles, vertical_spacing=0.14, horizontal_spacing=0.16)
            all_periods = pd.period_range(start="2010-01-01", end="2025-12-31", freq=freq_code)
            
            global_categories = set()
            data_matrices = []
            
            for cfg in configs:
                sub_filtered = execute_filter_pipeline(raw_df, cfg)
                if not sub_filtered.empty:
                    allocated_df = apply_fractional_allocation(sub_filtered, disaggregation)
                    global_categories.update(allocated_df["Active_Categories"].dropna().unique())
                    data_matrices.append(allocated_df)
                else:
                    data_matrices.append(pd.DataFrame())
                    
            sorted_categories = sorted(list(global_categories), key=lambda x: (str(x).startswith("Other"), x))
            plot_colors = get_dynamic_palette(sorted_categories, disaggregation)
            color_map = dict(zip(sorted_categories, plot_colors))
            
            for idx in range(chart_count):
                row, col = (idx // cols) + 1, (idx % cols) + 1
                df_allocated = data_matrices[idx]
                
                if not df_allocated.empty:
                    df_allocated['Period'] = df_allocated['Announcement Date'].dt.to_period(freq_code)
                    grouped = df_allocated.groupby(["Period", "Active_Categories"])[metric_col].sum().unstack(fill_value=0)
                    plot_data = grouped.reindex(index=all_periods, columns=sorted_categories, fill_value=0)
                    if smoothing > 1:
                        plot_data = plot_data.rolling(window=smoothing, min_periods=1).mean()
                else:
                    plot_data = pd.DataFrame(0.0, index=all_periods, columns=sorted_categories)
                    
                x_axis_labels = plot_data.index.astype(str).tolist()
                
                for cat in sorted_categories:
                    y_vals = plot_data[cat].tolist()
                    fig.add_trace(
                        go.Bar(
                            x=x_axis_labels, y=y_vals, name=cat, marker_color=color_map[cat],
                            hovertemplate="%{fullData.name}: %{y}<extra></extra>",
                            showlegend=(idx == 0), legendgroup=cat
                        ),
                        row=row, col=col
                    )
            
            metric_axis_label = {
                "Policy Count": "Policy Count",
                "Subsidy USD Amount": "Subsidy (USD Million)",
                "Trade Covered USD Amount": "Trade covered (USD Million)",
                "Combined USD Amount": "Combined (USD Million)",
            }[metric_choice]
            fig.update_layout(
                barmode='stack', hovermode='x unified', height=500 if chart_count == 1 else 750,
                paper_bgcolor="white", plot_bgcolor="white",
                margin=dict(l=50, r=30, t=60, b=100),
                legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="center", x=0.5)
            )
            fig.update_xaxes(showline=True, linewidth=1, linecolor=GREYS["Grey-2"], tickangle=45, automargin=True)
            # Let Plotly use its original compact SI formatting (e.g. 250k,
            # 250.3k) for tick values; the axis title provides the USD-million
            # unit context.
            fig.update_yaxes(title_text=metric_axis_label, showline=True, linewidth=1, linecolor=GREYS["Grey-2"], gridcolor=GREYS["Grey-1"], gridwidth=0.5, automargin=True)
            
            with plot_col:
                st.plotly_chart(fig, use_container_width=True)
                st.caption("Tip: click a legend item to hide or show it; double-click an item to isolate it in the chart.")
        else:
            with plot_col:
                st.info("Configure Chart 1 and select **Save chart** to start the figure.")

    # ------------------------------------------
    # METHODOLOGY TAB
    # ------------------------------------------
    with tab_methodology:
        summary_col, pdf_col = st.columns([1, 1])
        with summary_col:
            st.markdown("### NIPO methodology")
            st.markdown(
                """
                The [Global Trade Alert New Industrial Policy Observatory (NIPO)](https://globaltradealert.org/reports/new-industrial-policy-observatory-nipo) tracks government interventions that shape international trade and competition. The database records when a measure is announced, implemented, and removed, as well as the jurisdictions, sectors, products, policy instruments, motives, and trade flows affected.

                The methodology explains how interventions are identified and coded, how policy measures are classified, and how their trade and subsidy values are recorded. It also describes the assessment framework used to distinguish liberalising and distortive interventions and the treatment of measures affecting multiple products or jurisdictions.

                Use the document viewer to the right to browse the complete methodology.
                """
            )
        with pdf_col:
            methodology_path = Path(__file__).with_name("1774854873454_NIPO - Methodology.pdf")

            @st.cache_data
            def load_pdf_bytes(path):
                return path.read_bytes()

            if methodology_path.exists():
                methodology_bytes = load_pdf_bytes(methodology_path)
                if pdf_viewer is not None:
                    pdf_viewer(
                        methodology_bytes, height=700, viewer_align="center",
                        show_page_separator=True, render_text=True, resolution_boost=2,
                    )
                elif hasattr(st, "pdf"):
                    st.pdf(methodology_bytes, height=700)
                else:
                    st.warning("Install streamlit-pdf-viewer to enable the embedded PDF viewer.")
                st.download_button(
                    "Download methodology PDF", methodology_bytes,
                    file_name=methodology_path.name, mime="application/pdf",
                )
            else:
                st.error("The methodology PDF could not be found in the industrial-policy folder.")
else:
    st.warning("Please upload the NIPO database (XLSX file) to start exploring.")
