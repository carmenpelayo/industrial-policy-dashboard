import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl

# ==========================================
# 1. BBVA BRANDING & CONFIGURATION
# ==========================================
st.set_page_config(page_title="NIPO Policy Dashboard", layout="wide", page_icon="📊")

PRIMARY_COLORS = {"Electric Blue": "#001391", "Serene Blue": "#85C8FF", "Midnight": "#060E46"}
ACCENT_COLORS = ["#88E783", "#FFB56B", "#FFE761", "#8BE1E9", "#9694FF"]
GREYS = {"Sand": "#F7F8F8", "Grey-1": "#E2E6EA", "Grey-2": "#CAD1D8", "Grey-3": "#ADB8C2", "Grey-4": "#46536D"}

# Custom Semantic Colors for Assessment Type
ASSESSMENT_COLORS = {
    "Liberalising": "#1E5631", # Corporate Green
    "Distortive": "#B22222"     # Corporate Crimson Red
}

def get_dynamic_palette(categories, category_type):
    """
    Dynamically maps a unique color to each category ensuring 'Other' categories 
    resolve to grey, and Assessment types stick to strict green/red semantics.
    """
    if category_type == "Assessment Type":
        palette_map = {}
        for cat in categories:
            if cat == "Liberalising":
                palette_map[cat] = ASSESSMENT_COLORS["Liberalising"]
            elif cat == "Distortive":
                palette_map[cat] = ASSESSMENT_COLORS["Distortive"]
            else:
                palette_map[cat] = GREYS["Grey-3"]
        return [palette_map[c] for c in categories]
        
    # Standard groups logic
    base_corporate = [PRIMARY_COLORS["Electric Blue"], PRIMARY_COLORS["Serene Blue"]] + ACCENT_COLORS
    
    # Filter out catch-all labels to preserve sequential palette mapping 
    primary_cats = [c for c in categories if not str(c).startswith("Other")]
    
    # Generate backup overflow colors if needed
    if len(primary_cats) > len(base_corporate):
        overflow_needed = len(primary_cats) - len(base_corporate)
        supplementary = sns.color_palette("Spectral", overflow_needed).as_hex()
        color_pool = base_corporate + supplementary
    else:
        color_pool = base_corporate
        
    palette_map = {}
    idx = 0
    for cat in categories:
        if str(cat).startswith("Other"):
            palette_map[cat] = GREYS["Grey-3"] # Always Grey
        else:
            palette_map[cat] = color_pool[idx]
            idx += 1
            
    return [palette_map[c] for c in categories]

# Classifications Configuration
CPC_SECTIONS = {
    "0": "Agriculture, forestry, fishery", "1": "Ores, minerals, electricity, gas, water",
    "2": "Food, beverages, apparel, leather", "3": "Other transportable goods",
    "4": "Metal products, machinery", "5": "Constructions and services",
    "6": "Distributive trade, transport, hospitality", "7": "Financial, real estate, leasing",
    "8": "Business and production", "9": "Community, social, personal services"
}

# Standard boolean column groups
SECTOR_COLS = ["Sector: Low Carbon Technology", "Sector: Dual-Use Products", "Sector: Critical Minerals", 
               "Sector: Advanced Technology Products", "Sector: Medical Products", "Sector: Chemicals", "Sector: Includes IT or Digital Services"]
MOTIVE_COLS = ["Motive: National Security or Geopolitical Concern", "Motive: Resilience/Security of Supply (Non-Food)", 
               "Motive: Strategic Competitiveness", "Motive: Climate Change Mitigation", "Motive: Digital Transformation"]
POLICY_COLS = ["Is Export Policy", "Is Import Policy", "Is Trade Defence", "Is Subsidy", "Is Export Incentive", 
               "Is FDI Policy", "Is Procurement Policy", "Is Localisation Policy", "Is Other Policy"]

# ==========================================
# 2. CORE DATA PROCESSING
# ==========================================
@st.cache_data
def load_and_clean_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df["Announcement Date"] = pd.to_datetime(df["Announcement Date"], errors="coerce")
    df = df.dropna(subset=["Announcement Date"])
    df = df[df["Levels of Policy Intervention"] != "Firm-specific"]
    
    # Value parsing
    df["Trade Covered (USD Million)"] = pd.to_numeric(df["Trade Covered (USD Million)"], errors='coerce').fillna(0)
    df["Size of Subsidy (USD Million)"] = pd.to_numeric(df["Size of Subsidy (USD Million)"], errors='coerce').fillna(0)
    df["Total_USD_Value"] = df["Trade Covered (USD Million)"] + df["Size of Subsidy (USD Million)"]
    
    # Boolean parsing
    for col in SECTOR_COLS + MOTIVE_COLS + POLICY_COLS:
        if col in df.columns:
            df[col] = df[col].astype(str).str.upper() == "TRUE"
            
    df["Initial Assessment"] = df["Initial Assessment (Change Relative to 1 Jan 2009)"].astype(str).str.capitalize()
    return df

def apply_fractional_allocation(df, col_type, specific_mapping=None):
    df_temp = df.copy()
    
    if col_type == "Assessment Type":
        df_temp["Active_Categories"] = df_temp["Initial Assessment"].apply(
            lambda x: [x] if x in ["Liberalising", "Distortive"] else ["Other Assessments"])
        df_temp["Denominator"] = 1.0
        
    elif col_type in ["Product (CPC v2.1 Sectors)", "HS 2-digit"]:
        target_col = "Sector: CPC 3-digit (v2.1)" if col_type == "Product (CPC v2.1 Sectors)" else "Product: HS 6-digit (2022)"
        other_lbl = "Other Sections" if col_type == "Product (CPC v2.1 Sectors)" else "Other Products"
        
        def map_codes(val):
            val = str(val).strip()
            if val.upper() in ["NAN", "NONE", ""]: return [other_lbl]
            tokens = [t.strip() for t in val.split(",") if t.strip()]
            matched = set()
            for token in tokens:
                prefix = token[:1] if col_type == "Product (CPC v2.1 Sectors)" else token[:2].zfill(2)
                if prefix in specific_mapping: matched.add(specific_mapping[prefix])
            return list(matched) if matched else [other_lbl]
            
        df_temp["Active_Categories"] = df_temp[target_col].apply(map_codes)
        df_temp["Denominator"] = df_temp["Active_Categories"].apply(len)
        
    else: 
        cols = SECTOR_COLS if col_type == "Sector" else MOTIVE_COLS if col_type == "Motive" else POLICY_COLS
        lbl_suffix = "Sectors" if col_type == "Sector" else "Motives" if col_type == "Motive" else "Policies"
        
        df_temp["True_Count"] = df_temp[cols].sum(axis=1)
        def get_active(row):
            active = [c.split(": ")[-1].replace("Is ", "") for c in cols if row[c]]
            return active if active else [f"Other {lbl_suffix}"]
            
        df_temp["Active_Categories"] = df_temp.apply(get_active, axis=1)
        df_temp["Denominator"] = df_temp["True_Count"].replace(0, 1)

    # Calculate allocated column vectors
    df_temp["Allocated_Combined_USD"] = df_temp["Total_USD_Value"] / df_temp["Denominator"]
    df_temp["Allocated_Subsidy_USD"] = df_temp["Size of Subsidy (USD Million)"] / df_temp["Denominator"]
    df_temp["Allocated_Trade_USD"] = df_temp["Trade Covered (USD Million)"] / df_temp["Denominator"]
    df_temp["Allocated_Count"] = 1.0 / df_temp["Denominator"]
    
    return df_temp.explode("Active_Categories")

# ==========================================
# 3. STREAMLIT APP LAYOUT & LOGIC
# ==========================================
st.title("🌐 NIPO Industrial Policy Dashboard")

uploaded_file = st.file_uploader("Upload NIPO Dataset (XLSX)", type="xlsx")

if uploaded_file is not None:
    # Load Data
    with st.spinner("Processing Dataset..."):
        raw_df = load_and_clean_data(uploaded_file)
    
    # --- SIDEBAR FILTERS ---
    st.sidebar.header("🔍 Global Filters")
    
    # 1. Dates
    min_date, max_date = raw_df["Announcement Date"].min(), raw_df["Announcement Date"].max()
    date_range = st.sidebar.date_input("Announcement Date Range", [min_date, max_date], min_value=min_date, max_value=max_date)
    
    # 2. Jurisdictions
    all_imp = sorted(raw_df["Implementing Jurisdiction"].dropna().unique())
    selected_imp = st.sidebar.multiselect("Implementing Jurisdiction", all_imp, default=[])
    
    raw_df["Affected List"] = raw_df["Affected Jurisdiction"].astype(str).str.split(", ")
    all_aff = sorted(list(set(x.strip() for l in raw_df["Affected List"].dropna() for x in l)))
    selected_aff = st.sidebar.multiselect("Affected Jurisdiction", all_aff, default=[])

    # 3. Assessment (Removed "Other" safely)
    assessments = st.sidebar.multiselect("Initial Assessment", ["Liberalising", "Distortive"], default=[])

    # --- APPLY FILTERS ---
    filtered_df = raw_df.copy()
    if len(date_range) == 2:
        filtered_df = filtered_df[(filtered_df["Announcement Date"] >= pd.to_datetime(date_range[0])) & 
                                  (filtered_df["Announcement Date"] <= pd.to_datetime(date_range[1]))]
    if selected_imp:
        filtered_df = filtered_df[filtered_df["Implementing Jurisdiction"].isin(selected_imp)]
    if selected_aff:
        filtered_df = filtered_df[filtered_df["Affected List"].apply(lambda x: any(i in selected_aff for i in x))]
    if assessments:
        filtered_df = filtered_df[filtered_df["Initial Assessment"].isin(assessments)]

    # --- TABS ---
    tab1, tab2 = st.tabs(["🗃️ Data Inspector", "📈 Interactive Visualization"])
    
    with tab1:
        st.subheader("Filtered Raw Observations")
        st.write(f"Showing **{len(filtered_df):,}** policies based on criteria.")
        # Show all fields directly
        st.dataframe(filtered_df.drop(columns=["Affected List"], errors="ignore"), width='stretch')

    with tab2:
        st.subheader("Dynamic Plot Generator")
        
        # Plotting Controls Matrix
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            disaggregation = st.selectbox("Disaggregate By", 
                                          ["Sector", "Motive", "Policy Instrument", "Assessment Type", "Product (CPC v2.1 Sectors)"])
        with col2:
            freq_map = {"Daily": "D", "Monthly": "M", "Quarterly": "Q", "Yearly": "Y"}
            freq_choice = st.selectbox("Time Frequency", list(freq_map.keys()), index=3) # Defaults to Yearly
            freq_code = freq_map[freq_choice]
        with col3:
            smoothing = st.slider("Moving Average (Periods)", min_value=1, max_value=100, value=1, 
                                  help="Smooths the time-series bars using a rolling average filter.")
        with col4:
            metric_map = {
                "Policy Count": "Allocated_Count",
                "Subsidy USD Amount": "Allocated_Subsidy_USD",
                "Trade Covered USD Amount": "Allocated_Trade_USD",
                "Combined USD Amount": "Allocated_Combined_USD"
            }
            metric_choice = st.selectbox("Metric to Plot", list(metric_map.keys()))
            metric_col = metric_map[metric_choice]

        if st.button("Generate Plot", type="primary"):
            with st.spinner("Applying fractional allocations and rendering..."):
                plot_df = apply_fractional_allocation(filtered_df, disaggregation, specific_mapping=CPC_SECTIONS)
                
                # Resample time series
                plot_df['Period'] = plot_df['Announcement Date'].dt.to_period(freq_code)
                grouped = plot_df.groupby(["Period", "Active_Categories"])[metric_col].sum().unstack(fill_value=0)
                
                # Reindex Timeline Frame
                all_periods = pd.period_range(start=plot_df['Period'].min(), end=plot_df['Period'].max(), freq=freq_code)
                grouped = grouped.reindex(all_periods, fill_value=0)
                
                # Rolling calculations
                if smoothing > 1:
                    grouped = grouped.rolling(window=smoothing, min_periods=1).mean()

                # Generate canvas configuration
                fig, ax = plt.subplots(figsize=(14, 6), facecolor='white')
                ax.set_facecolor(GREYS["Sand"])
                
                categories = list(grouped.columns)
                colors = get_dynamic_palette(categories, disaggregation)
                
                grouped.index = grouped.index.astype(str)
                grouped.plot(kind='bar', stacked=True, ax=ax, color=colors, width=0.8)
                
                # Title and labels text configuration
                ax.set_title(f"Filtered Interventions - {metric_choice} by {disaggregation} ({freq_choice})", 
                             fontsize=14, fontweight='bold', color=PRIMARY_COLORS["Midnight"], pad=15)
                ax.set_ylabel("USD (Millions)" if "USD" in metric_choice else "Policy Count (Weighted)", color=GREYS["Grey-4"])
                ax.set_xlabel("Time Period", color=GREYS["Grey-4"])
                
                ax.grid(axis='y', linestyle='--', alpha=0.5, color=GREYS["Grey-3"])
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['left'].set_color(GREYS["Grey-2"])
                ax.spines['bottom'].set_color(GREYS["Grey-2"])
                ax.tick_params(colors=GREYS["Grey-4"])
                
                # Smart label cleanup logic
                if len(grouped.index) > 30:
                    n = len(grouped.index) // 15
                    for index, label in enumerate(ax.xaxis.get_ticklabels()):
                        if index % n != 0: label.set_visible(False)
                plt.xticks(rotation=45, ha='right')
                
                ax.legend(title=disaggregation, bbox_to_anchor=(1.02, 1), loc='upper left', frameon=True)
                plt.tight_layout()
                
                st.pyplot(fig)
else:
    st.info("👋 Welcome to the NIPO Dashboard! Please upload your XLSX dataset to begin.")
