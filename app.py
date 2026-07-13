import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl

st.set_page_config(page_title="NIPO Industrial Policy Dashboard", layout="wide", page_icon="📊")

# ==========================================
# 1. CORE DICTIONARIES & REGIONAL GROUPS
# ==========================================
PRIMARY_COLORS = {"Electric Blue": "#001391", "Serene Blue": "#85C8FF", "Midnight": "#060E46"}
ACCENT_COLORS = ["#88E783", "#FFB56B", "#FFE761", "#8BE1E9", "#9694FF"]
GREYS = {"Sand": "#F7F8F8", "Grey-1": "#E2E6EA", "Grey-2": "#CAD1D8", "Grey-3": "#ADB8C2", "Grey-4": "#46536D"}
ASSESSMENT_COLORS = {"Liberalising": "#1E5631", "Distortive": "#B22222"}

CPC_SECTIONS = {
    "0": "Agriculture, forestry, fishery", "1": "Ores, minerals, electricity, gas, water",
    "2": "Food, beverages, apparel, leather", "3": "Other transportable goods",
    "4": "Metal products, machinery", "5": "Constructions and services",
    "6": "Distributive trade, transport, hospitality", "7": "Financial, real estate, leasing",
    "8": "Business and production", "9": "Community, social, personal services"
}

# 2-digit macro lookup wrappers
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

COUNTRY_GROUPS = {
    "G-7": ["United States of America", "United Kingdom", "Germany", "France", "Italy", "Japan", "Canada"],
    "G-20": ["United States of America", "United Kingdom", "Germany", "France", "Italy", "Japan", "Canada", "Argentina", "Australia", "Brazil", "China", "India", "Indonesia", "Mexico", "Saudi Arabia", "South Africa", "South Korea", "Turkey", "Russia"],
    "EU-27": ["Austria", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Czechia", "Denmark", "Estonia", "Finland", "France", "Germany", "Greece", "Hungary", "Ireland", "Italy", "Latvia", "Lithuania", "Luxembourg", "Malta", "Netherlands", "Poland", "Portugal", "Romania", "Slovakia", "Slovenia", "Spain", "Sweden"],
    "Europe": ["Albania", "Andorra", "Austria", "Belarus", "Belgium", "Bosnia and Herzegovina", "Bulgaria", "Croatia", "Cyprus", "Czechia", "Denmark", "Estonia", "Finland", "France", "Germany", "Greece", "Hungary", "Iceland", "Ireland", "Italy", "Latvia", "Liechtenstein", "Lithuania", "Luxembourg", "Malta", "Moldova", "Monaco", "Montenegro", "Netherlands", "North Macedonia", "Norway", "Poland", "Portugal", "Romania", "Russia", "San Marino", "Serbia", "Slovakia", "Slovenia", "Spain", "Sweden", "Switzerland", "Ukraine", "United Kingdom", "Vatican City"],
    "North America": ["United States of America", "Canada", "Mexico", "Cuba", "Guatemala", "Honduras", "Panama", "Costa Rica", "Jamaica", "Dominican Republic"],
    "South America": ["Argentina", "Bolivia", "Brazil", "Chile", "Colombia", "Ecuador", "Guyana", "Paraguay", "Peru", "Suriname", "Uruguay", "Venezuela"],
    "LatAm": ["Argentina", "Bolivia", "Brazil", "Chile", "Colombia", "Costa Rica", "Cuba", "Dominican Republic", "Ecuador", "El Salvador", "Guatemala", "Honduras", "Mexico", "Nicaragua", "Panama", "Paraguay", "Peru", "Uruguay", "Venezuela"],
    "Asia": ["China", "Japan", "India", "South Korea", "Indonesia", "Saudi Arabia", "Turkey", "Vietnam", "Thailand", "Malaysia", "Singapore", "Philippines", "Pakistan", "Bangladesh", "Iran", "Iraq", "Israel", "Jordan", "Lebanon", "Oman", "Qatar", "Kuwait", "United Arab Emirates"],
    "Middle East": ["Saudi Arabia", "Turkey", "Iran", "Iraq", "Israel", "Jordan", "Lebanon", "Oman", "Qatar", "Kuwait", "United Arab Emirates", "Yemen", "Syria", "Egypt"],
    "Africa": ["South Africa", "Nigeria", "Egypt", "Algeria", "Morocco", "Kenya", "Ethiopia", "Ghana", "Angola", "Tanzania"],
    "Oceania": ["Australia", "New Zealand", "Papua New Guinea", "Fiji"]
}

SECTOR_COLS = ["Sector: Low Carbon Technology", "Sector: Dual-Use Products", "Sector: Critical Minerals", "Sector: Advanced Technology Products", "Sector: Medical Products", "Sector: Chemicals", "Sector: Includes IT or Digital Services"]
MOTIVE_COLS = ["Motive: National Security or Geopolitical Concern", "Motive: Resilience/Security of Supply (Non-Food)", "Motive: Strategic Competitiveness", "Motive: Climate Change Mitigation", "Motive: Digital Transformation"]
POLICY_COLS = ["Is Export Policy", "Is Import Policy", "Is Trade Defence", "Is Subsidy", "Is Export Incentive", "Is FDI Policy", "Is Procurement Policy", "Is Localisation Policy", "Is Other Policy"]

# ==========================================
# 2. DATA LOAD & UTILITY PIPELINES
# ==========================================
@st.cache_data
def load_source_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df["Announcement Date"] = pd.to_datetime(df["Announcement Date"], errors="coerce")
    df = df.dropna(subset=["Announcement Date"])
    df = df[df["Levels of Policy Intervention"] != "Firm-specific"]
    
    df["Trade Covered (USD Million)"] = pd.to_numeric(df["Trade Covered (USD Million)"], errors='coerce').fillna(0)
    df["Size of Subsidy (USD Million)"] = pd.to_numeric(df["Size of Subsidy (USD Million)"], errors='coerce').fillna(0)
    df["Total_USD_Value"] = df["Trade Covered (USD Million)"] + df["Size of Subsidy (USD Million)"]
    
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
    primary_cats = [c for c in categories if not str(c).startswith("Other")]
    
    if len(primary_cats) > len(base_corporate):
        color_pool = base_corporate + sns.color_palette("Spectral", len(primary_cats) - len(base_corporate)).as_hex()
    else:
        color_pool = base_corporate
        
    palette_map = {}
    idx = 0
    for cat in categories:
        if str(cat).startswith("Other"):
            palette_map[cat] = GREYS["Grey-3"]
        else:
            palette_map[cat] = color_pool[idx]
            idx += 1
    return [palette_map[c] for c in categories]

def resolve_jurisdictions(selected_items):
    resolved = set()
    for item in selected_items:
        if item in COUNTRY_GROUPS:
            resolved.update(COUNTRY_GROUPS[item])
        else:
            resolved.add(item)
    return list(resolved)

# ==========================================
# 3. GRANULAR SELECTION-DRIVEN FILTER ENGINE
# ==========================================
def execute_filter_pipeline(df, config):
    df_out = df.copy()
    
    # 1. Temporal Bounds
    if len(config.get("dates", [])) == 2:
        df_out = df_out[(df_out["Announcement Date"] >= pd.to_datetime(config["dates"][0])) & 
                        (df_out["Announcement Date"] <= pd.to_datetime(config["dates"][1]))]
                        
    # 2. Categorical Strings
    if config.get("gov_level"):
        df_out = df_out[df_out["Level of Government Implementation"].isin(config["gov_level"])]
    if config.get("trade_flow"):
        df_out = df_out[df_out["Affected Trade Flow"].isin(config["trade_flow"])]
    if config.get("assessments"):
        df_out = df_out[df_out["Initial Assessment"].isin(config["assessments"])]
        
    # 3. Jurisdictions
    if config.get("imp_jurisdiction"):
        resolved_imp = resolve_jurisdictions(config["imp_jurisdiction"])
        df_out = df_out[df_out["Implementing Jurisdiction"].isin(resolved_imp)]
    if config.get("aff_jurisdiction"):
        resolved_aff = resolve_jurisdictions(config["aff_jurisdiction"])
        df_out = df_out[df_out["Affected List"].apply(lambda x: any(i in resolved_aff for i in x))]
        
    # 4. Comma-Separated Multi-Label Token Substring Lookups
    if config.get("hs_2d"):
        codes = [item.split("(")[1].replace(")", "").strip() for item in config["hs_2d"]]
        def match_hs(val):
            tokens = [t.strip()[:2].zfill(2) for t in str(val).split(",") if t.strip()]
            return any(c in tokens for c in codes)
        df_out = df_out[df_out["Product: HS 6-digit (2022)"].apply(match_hs)]
        
    if config.get("cpc_2d"):
        codes = [item.split("(")[1].replace(")", "").strip() for item in config["cpc_2d"]]
        def match_cpc(val):
            tokens = [t.strip()[:2].zfill(2) for t in str(val).split(",") if t.strip()]
            return any(c in tokens for c in codes)
        df_out = df_out[df_out["Sector: CPC 3-digit (v2.1)"].apply(match_cpc)]

    # 5. Column Flag Logical Filters
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
            
    return df_out

# ==========================================
# 4. SEGMENTED APPORTIONMENT ALLOCATION MATH
# ==========================================
def apply_fractional_allocation(df, col_type):
    df_temp = df.copy()
    
    if col_type == "Assessment Type":
        df_temp["Active_Categories"] = df_temp["Initial Assessment"].apply(lambda x: [x] if x in ["Liberalising", "Distortive"] else ["Other Assessments"])
        df_temp["Denominator"] = 1.0
    elif col_type in ["Product (CPC v2.1 Sectors)", "Product: HS 6-digit (2022)", "Sector: CPC 3-digit (v2.1)"]:
        target_col = "Sector: CPC 3-digit (v2.1)" if col_type in ["Product (CPC v2.1 Sectors)", "Sector: CPC 3-digit (v2.1)"] else "Product: HS 6-digit (2022)"
        
        def split_all_codes(val):
            val = str(val).strip()
            if val.upper() in ["NAN", "NONE", ""]: return [f"Other {col_type}"]
            tokens = list(set([t.strip() for t in val.split(",") if t.strip()]))
            
            if col_type == "Product (CPC v2.1 Sectors)":
                return list(set([CPC_SECTIONS.get(t[:1], "Other Sections") for t in tokens]))
            elif col_type == "Sector: CPC 3-digit (v2.1)":
                return [f"CPC {t[:3].zfill(3)}" for t in tokens]
            else:
                return [f"HS {t[:2].zfill(2)}" for t in tokens]
                
        df_temp["Active_Categories"] = df_temp[target_col].apply(split_all_codes)
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
# 5. UI CONFIGURATION FOR MENUS
# ==========================================
def render_config_form(df_source, key_prefix):
    all_imp = list(COUNTRY_GROUPS.keys()) + sorted(df_source["Implementing Jurisdiction"].dropna().unique().tolist())
    all_aff = list(COUNTRY_GROUPS.keys()) + sorted(list(set(x for l in df_source["Affected List"].dropna() for x in l)))
    all_gov = sorted(df_source["Level of Government Implementation"].dropna().unique().tolist())
    all_flow = sorted(df_source["Affected Trade Flow"].dropna().unique().tolist())
    
    hs_options = [f"{v} ({k})" for k, v in HS_PRODUCTS_2D.items()]
    cpc_options = [f"{v} ({k})" for k, v in CPC_PRODUCTS_2D.items()]

    dates = st.date_input("Date Range", [df_source["Announcement Date"].min(), df_source["Announcement Date"].max()], key=f"{key_prefix}_dates")
    imp_jurisdiction = st.multiselect("Implementing Jurisdiction/Group", all_imp, key=f"{key_prefix}_imp")
    aff_jurisdiction = st.multiselect("Affected Jurisdiction/Group", all_aff, key=f"{key_prefix}_aff")
    gov_level = st.multiselect("Level of Government", all_gov, key=f"{key_prefix}_gov")
    trade_flow = st.multiselect("Affected Trade Flow", all_flow, key=f"{key_prefix}_flow")
    assessments = st.multiselect("Initial Assessment", ["Liberalising", "Distortive"], key=f"{key_prefix}_assess")
    hs_2d = st.multiselect("Product (HS 2-digit)", hs_options, key=f"{key_prefix}_hs2d")
    cpc_2d = st.multiselect("Product (CPC 2-digit)", cpc_options, key=f"{key_prefix}_cpc2d")
    policies = st.multiselect("Policy Instruments (Active Flags)", POLICY_COLS, key=f"{key_prefix}_pols")
    sectors = st.multiselect("Sectors (Active Flags)", SECTOR_COLS + ["Others"], key=f"{key_prefix}_secs")
    motives = st.multiselect("Motives (Active Flags)", MOTIVE_COLS + ["Others"], key=f"{key_prefix}_mots")

    return {
        "dates": dates, "imp_jurisdiction": imp_jurisdiction, "aff_jurisdiction": aff_jurisdiction,
        "gov_level": gov_level, "trade_flow": trade_flow, "assessments": assessments,
        "hs_2d": hs_2d, "cpc_2d": cpc_2d, "policies": policies, "sectors": sectors, "motives": motives
    }

# ==========================================
# 6. APP EXECUTION RUNNER WORKSPACE
# ==========================================
st.title("🌐 NIPO Industrial Policy Workspace Engine")
uploaded_file = st.file_uploader("Upload NIPO XLSX Source File", type="xlsx")

if uploaded_file is not None:
    raw_df = load_source_data(uploaded_file)
    tab1, tab2 = st.tabs(["🗃️ Data Inspector Hub", "📈 Matrix Visualization Grid"])

    # ------------------------------------------
    # DATA INSPECTOR TAB
    # ------------------------------------------
    with tab1:
        st.subheader("Data Inspector Workspace")
        with st.sidebar.expander("🔍 Data Inspector Filters", expanded=True):
            inspector_config = render_config_form(raw_df, "inspector")
        
        ins_df = execute_filter_pipeline(raw_df, inspector_config)
        st.write(f"Observations Matched: **{len(ins_df):,}** rows")
        
        drop_cols = ["NEW", "Entry ID", "Was First Reported Before This Inventory Month?", "Initial Assessment (Change Relative to 1 Jan 2009)", "Affected List"]
        display_df = ins_df.drop(columns=[c for c in drop_cols if c in ins_df.columns], errors="ignore")
        st.dataframe(display_df, use_container_width=True, hide_index=True)

    # ------------------------------------------
    # VISUALIZATION TAB (2x2 Grid)
    # ------------------------------------------
    with tab2:
        st.subheader("2x2 Layout Matrix Visualization Configurator")
        
        # Dashboard parameters
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            disaggregation = st.selectbox("Split Metric Series By", ["Sector", "Motive", "Policy Instrument", "Assessment Type", "Product (CPC v2.1 Sectors)", "Product: HS 6-digit (2022)", "Sector: CPC 3-digit (v2.1)"])
        with col2:
            freq_choice = st.selectbox("Time Window Frequency", ["Daily", "Monthly", "Quarterly", "Yearly"], index=3)
            freq_code = {"Daily": "D", "Monthly": "M", "Quarterly": "Q", "Yearly": "Y"}[freq_choice]
        with col3:
            smoothing = st.slider("Smoothing Filter Window (Periods)", min_value=1, max_value=100, value=1)
        with col4:
            metric_choice = st.selectbox("Apportionment Matrix Target", ["Policy Count", "Subsidy USD Amount", "Trade Covered USD Amount", "Combined USD Amount"])
            metric_col = {"Policy Count": "Allocated_Count", "Subsidy USD Amount": "Allocated_Subsidy_USD", "Trade Covered USD Amount": "Allocated_Trade_USD", "Combined USD Amount": "Allocated_Combined_USD"}[metric_choice]

        st.markdown("---")
        st.write("### Subplot Configurations (1 to 4)")
        
        # Main form controller logic
        p1_col, p2_col, p3_col, p4_col = st.columns(4)
        
        with p1_col:
            st.markdown("#### **Subplot 1 (Base/Master)**")
            p1_config = render_config_form(raw_df, "p1")
            
        with p2_col:
            override_p2 = st.checkbox("Customize Subplot 2", value=False)
            st.markdown("#### **Subplot 2**")
            p2_config = render_config_form(raw_df, "p2") if override_p2 else p1_config
            
        with p3_col:
            override_p3 = st.checkbox("Customize Subplot 3", value=False)
            st.markdown("#### **Subplot 3**")
            p3_config = render_config_form(raw_df, "p3") if override_p3 else p1_config
            
        with p4_col:
            override_p4 = st.checkbox("Customize Subplot 4", value=False)
            st.markdown("#### **Subplot 4**")
            p4_config = render_config_form(raw_df, "p4") if override_p4 else p1_config

        if st.button("Generate 2x2 Matrix Plots", type="primary", use_container_width=True):
            configs = [p1_config, p2_config, p3_config, p4_config]
            titles = ["Subplot 1 Profile", "Subplot 2 Profile", "Subplot 3 Profile", "Subplot 4 Profile"]
            
            fig, axes = plt.subplots(2, 2, figsize=(20, 11), sharex=False, facecolor='white')
            axes_flat = axes.flatten()
            
            # Identify the unified timeline tracking range across all data subsets
            global_min_yr, global_max_yr = 2010, 2025
            all_periods = pd.period_range(start=f"{global_min_yr}-01-01", end=f"{global_max_yr}-12-31", freq=freq_code)
            
            global_categories = set()
            data_matrices = []
            
            # Loop 1: Filter and apportion to pull complete label ranges safely
            for idx, cfg in enumerate(configs):
                sub_filtered = execute_filter_pipeline(raw_df, cfg)
                if not sub_filtered.empty:
                    allocated_df = apply_fractional_allocation(sub_filtered, disaggregation)
                    global_categories.update(allocated_df["Active_Categories"].dropna().unique())
                    data_matrices.append(allocated_df)
                else:
                    data_matrices.append(pd.DataFrame())
            
            sorted_categories = sorted(list(global_categories), key=lambda x: (str(x).startswith("Other"), x))
            plot_colors = get_dynamic_palette(sorted_categories, disaggregation)
            
            # Loop 2: Compile aggregated timeseries frames and plot
            for idx in range(4):
                ax = axes_flat[idx]
                ax.set_facecolor(GREYS["Sand"])
                df_allocated = data_matrices[idx]
                
                if not df_allocated.empty:
                    df_allocated['Period'] = df_allocated['Announcement Date'].dt.to_period(freq_code)
                    grouped = df_allocated.groupby(["Period", "Active_Categories"])[metric_col].sum().unstack(fill_value=0)
                    plot_data = grouped.reindex(index=all_periods, columns=sorted_categories, fill_value=0)
                    
                    if smoothing > 1:
                        plot_data = plot_data.rolling(window=smoothing, min_periods=1).mean()
                else:
                    plot_data = pd.DataFrame(0.0, index=all_periods, columns=sorted_categories)
                
                # Plot execution sequence
                plot_data.index = plot_data.index.astype(str)
                plot_data.plot(kind='bar', stacked=True, ax=ax, color=plot_colors, width=0.8, legend=False)
                
                ax.set_title(titles[idx], fontsize=11, fontweight='bold', color=PRIMARY_COLORS["Midnight"])
                ax.set_ylabel("USD (Millions)" if "USD" in metric_choice else "Weighted Count", fontsize=9, color=GREYS["Grey-4"])
                ax.grid(axis='y', linestyle='--', alpha=0.4, color=GREYS["Grey-3"])
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['left'].set_color(GREYS["Grey-2"])
                ax.spines['bottom'].set_color(GREYS["Grey-2"])
                ax.tick_params(colors=GREYS["Grey-4"], labelsize=8)
                
                # Limit tick clutter on dense ranges
                if len(plot_data.index) > 20:
                    step = len(plot_data.index) // 10
                    for tick_idx, label in enumerate(ax.xaxis.get_ticklabels()):
                        if tick_idx % step != 0: label.set_visible(False)
                ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')

            # Global layout presentation configuration adjustments
            fig.suptitle(f"2x2 EVALUATION ENGINE: {metric_choice.upper()} BY {disaggregation.upper()}", 
                         fontsize=14, fontweight='bold', color=PRIMARY_COLORS["Midnight"], y=0.98)
            
            # Render a single synchronized chart legend across the bottom row window
            handles, labels = axes[0, 0].get_legend_handles_labels()
            fig.legend(handles, labels, title=disaggregation, loc='lower center', 
                       bbox_to_anchor=(0.5, -0.05), ncol=min(6, len(sorted_categories)), frameon=True, facecolor='white', edgecolor=GREYS["Grey-1"])
            
            plt.tight_layout(rect=[0, 0.05, 1, 0.95])
            st.pyplot(fig)
else:
    st.info("👋 Welcome! Please upload your source XLSX dataset file to boot up the configuration matrices.")
