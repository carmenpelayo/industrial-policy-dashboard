import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl
import extra_streamlit_components as stx

st.set_page_config(page_title="NIPO Industrial Policy Dashboard", layout="wide", page_icon="📊")

# ==========================================
# 1. STYLE CONFIGURATION & LOOKUP DATA
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
# 2. FILE INTERPRETATION LOGIC
# ==========================================
@st.cache_data
def load_source_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df["Announcement Date"] = pd.to_datetime(df["Announcement Date"], errors="coerce")
    df = df.dropna(subset=["Announcement Date"])
    df = df[df["Levels of Policy Intervention"] != "Firm-specific"]
    
    # Currency vectors
    df["Trade Covered (USD Million)"] = pd.to_numeric(df["Trade Covered (USD Million)"], errors='coerce').fillna(0)
    df["Size of Subsidy (USD Million)"] = pd.to_numeric(df["Size of Subsidy (USD Million)"], errors='coerce').fillna(0)
    df["Total_USD_Value"] = df["Trade Covered (USD Million)"] + df["Size of Subsidy (USD Million)"]
    
    # Enforce strict strings for government levels mappings
    df["Level of Government Implementation"] = df["Level of Government Implementation"].replace({
        "IFI": "Independent Fiscal Institutions (IFI)",
        "NFI": "National Framework Implementations (NFI)"
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
    primary_cats = [c for c in categories if not str(c).startswith("Other")]
    color_pool = base_corporate + sns.color_palette("Spectral", max(0, len(primary_cats) - len(base_corporate))).as_hex()
    
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
        clean_item = item.replace("Group: ", "")
        if clean_item in COUNTRY_GROUPS:
            resolved.update(COUNTRY_GROUPS[clean_item])
        else:
            resolved.add(clean_item)
    return list(resolved)

# ==========================================
# 3. INTERACTIVE FILTER SELECTION CONTROL
# ==========================================
def execute_filter_pipeline(df, config):
    df_out = df.copy()
    if len(config.get("dates", [])) == 2:
        df_out = df_out[(df_out["Announcement Date"] >= pd.to_datetime(config["dates"][0])) & 
                        (df_out["Announcement Date"] <= pd.to_datetime(config["dates"][1]))]
    if config.get("gov_level"):
        df_out = df_out[df_out["Level of Government Implementation"].isin(config["gov_level"])]
    if config.get("trade_flow"):
        df_out = df_out[df_out["Affected Trade Flow"].isin(config["trade_flow"])]
    if config.get("assessments"):
        df_out = df_out[df_out["Initial Assessment"].isin(config["assessments"])]
    if config.get("imp_jurisdiction"):
        resolved_imp = resolve_jurisdictions(config["imp_jurisdiction"])
        df_out = df_out[df_out["Implementing Jurisdiction"].isin(resolved_imp)]
    if config.get("aff_jurisdiction"):
        resolved_aff = resolve_jurisdictions(config["aff_jurisdiction"])
        df_out = df_out[df_out["Affected List"].apply(lambda x: any(i in resolved_aff for i in x))]
    if config.get("hs_2d"):
        codes = [item.split("(")[1].replace(")", "").strip() for item in config["hs_2d"]]
        df_out = df_out[df_out["Product: HS 6-digit (2022)"].apply(lambda x: any(c in [t.strip()[:2].zfill(2) for t in str(x).split(",") if t.strip()] for c in codes))]
    if config.get("cpc_2d"):
        codes = [item.split("(")[1].replace(")", "").strip() for item in config["cpc_2d"]]
        df_out = df_out[df_out["Sector: CPC 3-digit (v2.1)"].apply(lambda x: any(c in [t.strip()[:2].zfill(2) for t in str(x).split(",") if t.strip()] for c in codes))]
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

def apply_fractional_allocation(df, col_type):
    df_temp = df.copy()
    if col_type == "Assessment Type":
        df_temp["Active_Categories"] = df_temp["Initial Assessment"].apply(lambda x: [x] if x in ["Liberalising", "Distortive"] else ["Other Assessments"])
        df_temp["Denominator"] = 1.0
    elif col_type in ["Product (CPC v2.1 Sectors)", "Product: HS 6-digit (2022)", "Sector: CPC 3-digit (v2.1)"]:
        target_col = "Sector: CPC 3-digit (v2.1)" if col_type in ["Product (CPC v2.1 Sectors)", "Sector: CPC 3-digit (v2.1)"] else "Product: HS 6-digit (2022)"
        def split_codes(val):
            val = str(val).strip()
            if val.upper() in ["NAN", "NONE", ""]: return [f"Other {col_type}"]
            tokens = list(set([t.strip() for t in val.split(",") if t.strip()]))
            if col_type == "Product (CPC v2.1 Sectors)":
                return list(set([CPC_SECTIONS.get(t[:1], "Other Sections") for t in tokens]))
            return [f"CPC {t[:3].zfill(3)}" for t in tokens] if col_type == "Sector: CPC 3-digit (v2.1)" else [f"HS {t[:2].zfill(2)}" for t in tokens]
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
# 4. RENDER PARAMETER CONFIG MATRIX
# ==========================================
def render_isolated_form(df_source, key_prefix, master_defaults=None):
    # Construct options arrays with explicit string separation mappings
    groups_list = [f"Group: {k}" for k in COUNTRY_GROUPS.keys()]
    all_imp = groups_list + sorted(df_source["Implementing Jurisdiction"].dropna().unique().tolist())
    all_aff = groups_list + sorted(list(set(x for l in df_source["Affected List"].dropna() for x in l)))
    all_gov = ["Independent Fiscal Institutions (IFI)", "National Framework Implementations (NFI)"] + sorted([x for x in df_source["Level of Government Implementation"].dropna().unique().tolist() if x not in ["Independent Fiscal Institutions (IFI)", "National Framework Implementations (NFI)"]])
    all_flow = sorted(df_source["Affected Trade Flow"].dropna().unique().tolist())
    
    hs_options = [f"{v} ({k})" for k, v in HS_PRODUCTS_2D.items()]
    cpc_options = [f"{v} ({k})" for k, v in CPC_PRODUCTS_2D.items()]

    # Helper function to compute fallbacks for form elements
    def get_val(field_key, default_fallback):
        if master_defaults and field_key in master_defaults:
            return master_defaults[field_key]
        return default_fallback

    dates = st.date_input("Date Range", get_val("dates", [df_source["Announcement Date"].min(), df_source["Announcement Date"].max()]), key=f"{key_prefix}_dates")
    imp_jurisdiction = st.multiselect("Implementing Jurisdiction/Group", all_imp, default=get_val("imp_jurisdiction", []), key=f"{key_prefix}_imp")
    aff_jurisdiction = st.multiselect("Affected Jurisdiction/Group", all_aff, default=get_val("aff_jurisdiction", []), key=f"{key_prefix}_aff")
    gov_level = st.multiselect("Level of Government", all_gov, default=get_val("gov_level", []), key=f"{key_prefix}_gov")
    trade_flow = st.multiselect("Affected Trade Flow", all_flow, default=get_val("trade_flow", []), key=f"{key_prefix}_flow")
    assessments = st.multiselect("Initial Assessment", ["Liberalising", "Distortive"], default=get_val("assessments", []), key=f"{key_prefix}_assess")
    hs_2d = st.multiselect("Product (HS 2-digit)", hs_options, default=get_val("hs_2d", []), key=f"{key_prefix}_hs2d")
    cpc_2d = st.multiselect("Product (CPC 2-digit)", cpc_options, default=get_val("cpc_2d", []), key=f"{key_prefix}_cpc2d")
    policies = st.multiselect("Policy Instruments", POLICY_COLS, default=get_val("policies", []), key=f"{key_prefix}_pols")
    sectors = st.multiselect("Sectors", SECTOR_COLS + ["Others"], default=get_val("sectors", []), key=f"{key_prefix}_secs")
    motives = st.multiselect("Motives", MOTIVE_COLS + ["Others"], default=get_val("motives", []), key=f"{key_prefix}_mots")

    return {
        "dates": dates, "imp_jurisdiction": imp_jurisdiction, "aff_jurisdiction": aff_jurisdiction,
        "gov_level": gov_level, "trade_flow": trade_flow, "assessments": assessments,
        "hs_2d": hs_2d, "cpc_2d": cpc_2d, "policies": policies, "sectors": sectors, "motives": motives
    }

# ==========================================
# 5. EXECUTION CONTAINER WORKSPACE
# ==========================================
st.markdown("## NIPO Industrial Policy Workspace Engine")
uploaded_file = st.file_uploader("Upload NIPO XLSX Source File", type="xlsx")

if uploaded_file is not None:
    raw_df = load_source_data(uploaded_file)
    active_tab = stx.tab_bar(data=[
        stx.TabBarItemData(id="inspector", title="🗃️ Data Inspector Hub", description="Raw Records Parsing"),
        stx.TabBarItemData(id="viz", title="📈 Matrix Visualization Grid", description="2x2 Subplot Workspace")
    ])

    # ------------------------------------------
    # DATA INSPECTOR SIDEBAR WORKSPACE
    # ------------------------------------------
    if active_tab == "inspector":
        with st.sidebar:
            st.markdown("### 🔍 Filter Raw Data Records")
            inspector_config = render_isolated_form(raw_df, "inspector")
            
        st.subheader("Data Inspector Workspace")
        ins_df = execute_filter_pipeline(raw_df, inspector_config)
        st.write(f"Observations Matched: **{len(ins_df):,}** rows")
        
        drop_cols = ["NEW", "Entry ID", "Was First Reported Before This Inventory Month?", "Initial Assessment (Change Relative to 1 Jan 2009)", "Affected List"]
        display_df = ins_df.drop(columns=[c for c in drop_cols if c in ins_df.columns], errors="ignore")
        st.dataframe(display_df, use_container_width=True, hide_index=True)

    # ------------------------------------------
    # INTERACTIVE VISUALIZATION GRID (2x2 Matrix)
    # ------------------------------------------
    elif active_tab == "viz":
        # Master Global Plot Configuration Bar Top
        with st.container():
            st.subheader("Global Visualization Matrix Parameters")
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

        # Isolate the filter controls entirely into the left side sidebar
        with st.sidebar:
            st.markdown("### 🎛️ Subplot Matrix Controller")
            config_mode = st.radio("Configuration Mode", ["Synchronized (Inherit Subplot 1)", "Independent Customization per Subplot"])
            
            st.markdown("---")
            st.markdown("### 📊 Subplot 1 (Master Core Configuration)")
            p1_config = render_isolated_form(raw_df, "p1")
            
            if config_mode == "Independent Customization per Subplot":
                st.markdown("---")
                st.markdown("### 📊 Subplot 2 Configuration")
                p2_config = render_isolated_form(raw_df, "p2", master_defaults=p1_config)
                st.markdown("---")
                st.markdown("### 📊 Subplot 3 Configuration")
                p3_config = render_isolated_form(raw_df, "p3", master_defaults=p1_config)
                st.markdown("---")
                st.markdown("### 📊 Subplot 4 Configuration")
                p4_config = render_isolated_form(raw_df, "p4", master_defaults=p1_config)
            else:
                p2_config, p3_config, p4_config = p1_config, p1_config, p1_config

        # Canvas Render Workspace
        st.markdown("---")
        if st.button("Generate 2x2 Matrix Plots", type="primary", use_container_width=True):
            configs = [p1_config, p2_config, p3_config, p4_config]
            titles = ["Subplot 1 Profile", "Subplot 2 Profile", "Subplot 3 Profile", "Subplot 4 Profile"]
            
            fig, axes = plt.subplots(2, 2, figsize=(20, 11), facecolor='white')
            axes_flat = axes.flatten()
            
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
                
                if len(plot_data.index) > 20:
                    step = len(plot_data.index) // 10
                    for tick_idx, label in enumerate(ax.xaxis.get_ticklabels()):
                        if tick_idx % step != 0: label.set_visible(False)
                ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')

            fig.suptitle(f"2x2 EVALUATION ENGINE: {metric_choice.upper()} BY {disaggregation.upper()}", fontsize=14, fontweight='bold', color=PRIMARY_COLORS["Midnight"], y=0.98)
            handles, labels = axes[0, 0].get_legend_handles_labels()
            fig.legend(handles, labels, title=disaggregation, loc='lower center', bbox_to_anchor=(0.5, -0.05), ncol=min(6, len(sorted_categories)), frameon=True, facecolor='white', edgecolor=GREYS["Grey-1"])
            
            plt.tight_layout(rect=[0, 0.05, 1, 0.95])
            st.pyplot(fig)
else:
    st.info("👋 Welcome! Please upload your source XLSX dataset file to display the configuration matrices.")
