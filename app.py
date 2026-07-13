import streamlit as st
import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import seaborn as sns
import openpyxl
import extra_streamlit_components as stx

# ==========================================
# 1. BRANDING, CONFIGURATIONS & LOOKUPS
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
# 2. FILE LOADING AND COMPONENT ALLOCATIONS
# ==========================================
@st.cache_data
def load_source_data(uploaded_file):
    df = pd.read_csv(uploaded_file)
    df["Announcement Date"] = pd.to_datetime(df["Announcement Date"], errors="coerce")
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

# ==========================================
# 3. ADVANCED BOOLEAN LOGIC ENGINE
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
            is_match = "True" if term in target_lower else "False"
            new_parts.append(is_match)
            
    compiled_expr = " ".join(new_parts)
    try:
        return eval(compiled_expr, {"__builtins__": None}, {})
    except:
        return False

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
            
    if config.get("keyword_search"):
        df_out = df_out[df_out["Title"].apply(lambda x: evaluate_boolean_query(x, config["keyword_search"]))]
        
    return df_out

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
            return [f"CPC {t[:3].zfill(3)}" for t in tokens] if col_type == "Sector: CPC 3-digit (v2.1)" else [f"HS {t[:2].zfill(2)}" for t in tokens]
        df_temp["Active_Categories"] = df_temp[target_col].apply(split_codes)
        df_temp["Denominator"] = df_temp["Active_Categories"].apply(len)
    else:
        cols = SECTOR_COLS if col_type == "Sector" else MOTIVE_COLS if col_type == "Motive" else POLICY_COLS
        lbl = "Sectors" if col_type == "Sector" else "Motives" if col_type == "Motive" else "Policies"
        df_temp["True_Count"] = df_temp[cols].sum(axis=1)
        df_temp["Active_Categories"] = df_temp.apply(lambda r: [c.split(": ")[-1].replace("Is ", "") for c in cols if r[c]] or [f"Other {lbl}"], axis=1)
        df_temp["Denominator"] = df_temp["True_Count"].replace(0, 1)

    for m, orig in [("Allocated_Combined_USD", "Total_USD_Value"), ("Allocated_Subsidy_USD", "Size of Subsidy (USD Million)"), ("Allocated_Trade_USD", "Trade Covered (USD Million)")]:
        df_temp[m] = df_temp[orig] / df_temp["Denominator"]
    df_temp["Allocated_Count"] = 1.0 / df_temp["Denominator"]
    return df_temp.explode("Active_Categories")

# ==========================================
# 4. MODULAR SELECTION MENU BUILDER
# ==========================================
def render_inline_filters(df_source, key_prefix, master_ref=None):
    groups_list = [f"Group: {k}" for k in COUNTRY_GROUPS.keys()]
    all_imp = groups_list + sorted(df_source["Implementing Jurisdiction"].dropna().unique().tolist())
    all_aff = groups_list + sorted(list(set(x for l in df_source["Affected List"].dropna() for x in l)))
    all_gov = ["Independent Fiscal Institutions (IFI)", "National Framework Implementations (NFI)"] + sorted([x for x in df_source["Level of Government Implementation"].dropna().unique().tolist() if x not in ["Independent Fiscal Institutions (IFI)", "National Framework Implementations (NFI)"]])
    all_flow = sorted(df_source["Affected Trade Flow"].dropna().unique().tolist())
    
    hs_opts = [f"{v} ({k})" for k, v in HS_PRODUCTS_2D.items()]
    cpc_opts = [f"{v} ({k})" for k, v in CPC_PRODUCTS_2D.items()]

    def get_fallback(field, default):
        return master_ref[field] if master_ref and field in master_ref else default

    keyword_search = st.text_input("Keyword Logic Query", get_fallback("keyword_search", ""), key=f"{key_prefix}_kw", help="Supports queries like: (AI OR Intelligence) AND semiconductor")
    dates = st.date_input("Timeline Window", get_fallback("dates", [df_source["Announcement Date"].min(), df_source["Announcement Date"].max()]), key=f"{key_prefix}_dt")
    imp_jurisdiction = st.multiselect("Implementing Countries", all_imp, default=get_fallback("imp_jurisdiction", []), key=f"{key_prefix}_imp")
    aff_jurisdiction = st.multiselect("Affected Countries", all_aff, default=get_fallback("aff_jurisdiction", []), key=f"{key_prefix}_aff")
    gov_level = st.multiselect("Government Level", all_gov, default=get_fallback("gov_level", []), key=f"{key_prefix}_gov")
    trade_flow = st.multiselect("Trade Flow", all_flow, default=get_fallback("trade_flow", []), key=f"{key_prefix}_flow")
    assessments = st.multiselect("Initial Assessment", ["Liberalising", "Distortive"], default=get_fallback("assessments", []), key=f"{key_prefix}_assess")
    hs_2d = st.multiselect("HS 2-digit Product", hs_opts, default=get_fallback("hs_2d", []), key=f"{key_prefix}_hs2d")
    cpc_2d = st.multiselect("CPC 2-digit Product", cpc_opts, default=get_fallback("cpc_2d", []), key=f"{key_prefix}_cpc2d")
    policies = st.multiselect("Policy Flags", POLICY_COLS, default=get_fallback("policies", []), key=f"{key_prefix}_pols")
    sectors = st.multiselect("Sector Flags", SECTOR_COLS + ["Others"], default=get_fallback("sectors", []), key=f"{key_prefix}_secs")
    motives = st.multiselect("Motive Flags", MOTIVE_COLS + ["Others"], default=get_fallback("motives", []), key=f"{key_prefix}_mots")

    return {
        "keyword_search": keyword_search, "dates": dates, "imp_jurisdiction": imp_jurisdiction, "aff_jurisdiction": aff_jurisdiction,
        "gov_level": gov_level, "trade_flow": trade_flow, "assessments": assessments, "hs_2d": hs_2d, "cpc_2d": cpc_2d,
        "policies": policies, "sectors": sectors, "motives": motives
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

# ==========================================
# 5. DASHBOARD FRAMEWORK APPLICATION INTERFACE
# ==========================================
uploaded_file = st.file_uploader("Upload NIPO CSV Dataset File", type="csv")

if uploaded_file is not None:
    raw_df = load_source_data(uploaded_file)
    
    # Native tab bar configuration alignment
    tab_inspect, tab_viz = st.tabs(["🗃️ Data Inspector Hub", "📈 Matrix Visualization Grid"])

    # ------------------------------------------
    # DATA INSPECTOR CONFIGURATION PANEL
    # ------------------------------------------
    with tab_inspect:
        filter_col, plot_col = st.columns([1, 3])
        with filter_col:
            st.markdown("#### Configuration Filter Menu")
            inspector_config = render_inline_filters(raw_df, "inspector")
            trigger_inspect = st.button("Generate Output Data View", type="primary", use_container_width=True)
            
        with plot_col:
            st.subheader("Filtered Raw Observations Table")
            if trigger_inspect:
                ins_df = execute_filter_pipeline(raw_df, inspector_config)
                st.write(f"Observations Matched: **{len(ins_df):,}** rows")
                drop_fields = ["NEW", "Entry ID", "Was First Reported Before This Inventory Month?", "Initial Assessment (Change Relative to 1 Jan 2009)", "Affected List"]
                display_df = ins_df.drop(columns=[c for c in drop_fields if c in ins_df.columns], errors="ignore")
                st.dataframe(display_df, use_container_width=True, hide_index=True)
            else:
                st.info("Adjust the menu options on the left and click 'Generate Output Data View'.")

    # ------------------------------------------
    # INTERACTIVE 2X2 VISUALIZATION GRID PANEL
    # ------------------------------------------
    with tab_viz:
        st.subheader("Global Visualization Matrix Parameters")
        g1, g2, g3, g4 = st.columns(4)
        with g1: disaggregation = st.selectbox("Split Metric Series By", ["Sector", "Motive", "Policy Instrument", "Assessment Type", "Product (CPC v2.1 Sectors)", "Product: HS 6-digit (2022)", "Sector: CPC 3-digit (v2.1)"])
        with g2: freq_choice = st.selectbox("Time Window Frequency", ["Daily", "Monthly", "Quarterly", "Yearly"], index=3)
        with g3: smoothing = st.slider("Smoothing Filter Window (Periods)", min_value=1, max_value=100, value=1)
        with g4: metric_choice = st.selectbox("Apportionment Matrix Target", ["Policy Count", "Subsidy USD Amount", "Trade Covered USD Amount", "Combined USD Amount"])
        
        freq_code = {"Daily": "D", "Monthly": "M", "Quarterly": "Q", "Yearly": "Y"}[freq_choice]
        metric_col = {"Policy Count": "Allocated_Count", "Subsidy USD Amount": "Allocated_Subsidy_USD", "Trade Covered USD Amount": "Allocated_Trade_USD", "Combined USD Amount": "Allocated_Combined_USD"}[metric_choice]
        
        st.markdown("---")
        st.markdown("#### Subplot Filtering Mappings (1 to 4 From Left to Right)")
        
        # Horizontal layout setup columns alignment
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown("##### **Subplot 1 (Master)**")
            p1_raw = render_inline_filters(raw_df, "v_p1")
        with c2:
            st.markdown("##### **Subplot 2**")
            p2_raw = render_inline_filters(raw_df, "v_p2")
        with c3:
            st.markdown("##### **Subplot 3**")
            p3_raw = render_inline_filters(raw_df, "v_p3")
        with c4:
            st.markdown("##### **Subplot 4**")
            p4_raw = render_inline_filters(raw_df, "v_p4")
            
        # Resolve inheritance constraints systematically
        p1_config = fill_missing_with_master(p1_raw, p1_raw)
        p2_config = fill_missing_with_master(p2_raw, p1_config)
        p3_config = fill_missing_with_master(p3_raw, p1_config)
        p4_config = fill_missing_with_master(p4_raw, p1_config)
        
        st.markdown("---")
        trigger_viz = st.button("Generate Matrix Figures", type="primary", use_container_width=True)
        
        if trigger_viz:
            configs = [p1_config, p2_config, p3_config, p4_config]
            titles = ["Subplot 1 Matrix Profile", "Subplot 2 Matrix Profile", "Subplot 3 Matrix Profile", "Subplot 4 Matrix Profile"]
            
            fig = make_subplots(rows=2, cols=2, subplot_titles=titles, vertical_spacing=0.12, horizontal_spacing=0.08)
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
            
            for idx in range(4):
                row, col = (idx // 2) + 1, (idx % 2) + 1
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
                
                # Append individual stacked traces into subplots frame seamlessly
                for cat in sorted_categories:
                    y_vals = plot_data[cat].tolist()
                    fig.add_trace(
                        go.Bar(
                            x=x_axis_labels, y=y_vals, name=cat, marker_color=color_map[cat],
                            hoverinfo="name+y", showlegend=(idx == 0), legendgroup=cat
                        ),
                        row=row, col=col
                    )
            
            # Interactive formatting configuration mapping (Unified Crosshair Tooltips)
            fig.update_layout(
                barmode='stack', hovermode='x unified', height=750,
                paper_bgcolor="white", plot_bgcolor="white",
                margin=dict(l=50, r=30, t=60, b=100),
                legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="center", x=0.5)
            )
            fig.update_xaxes(showline=True, linewidth=1, linecolor=GREYS["Grey-2"], tickangle=45)
            fig.update_yaxes(showline=True, linewidth=1, linecolor=GREYS["Grey-2"], gridcolor=GREYS["Grey-1"], gridwidth=0.5)
            
            st.plotly_chart(fig, use_container_width=True)
else:
    st.info("👋 Welcome! Please upload your source CSV dataset file to boot up the interactive workspace.")
