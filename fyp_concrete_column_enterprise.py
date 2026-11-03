import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import datetime
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import zipfile
from fpdf import FPDF

# 1. Professional Page Configuration
st.set_page_config(
    layout="wide",
    page_title="Enterprise Digital Twin & QS Control",
    page_icon="🏢",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    h1, h2, h3 { color: #2c3e50; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .metric-container { background-color: white; padding: 15px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.08); margin-bottom: 15px;}
    .cost-alert { color: #c0392b; font-weight: bold; font-size: 1.05em; }
    .claim-good { color: #1e8449; font-weight: bold; font-size: 1.05em; }
    </style>
""", unsafe_allow_html=True)

CROSS_AREA = np.pi * (0.61 / 2) ** 2  # 0.292247 m2 for 610mm dia column

# --- ROLE-BASED AUTHENTICATION & SIDEBAR ---
if 'role' not in st.session_state:
    st.session_state['role'] = 'Viewer'
if 'username' not in st.session_state:
    st.session_state['username'] = 'Guest'

with st.sidebar:
    st.header("🔐 System Login")
    pwd = st.text_input("Enter Password", type="password")
    if pwd == "admin123":
        st.session_state['role'] = 'Admin'
        st.session_state['username'] = 'Project Manager'
        st.success("Logged in: Admin (Full QS & Export Access)")
    elif pwd == "worker123":
        st.session_state['role'] = 'Editor'
        st.session_state['username'] = 'Site Engineer'
        st.success("Logged in: Site Engineer (Data Entry)")
    elif pwd != "":
        st.error("Invalid password.")
        st.session_state['role'] = 'Viewer'
        st.session_state['username'] = 'Guest'
    else:
        st.info("Log in (`admin123` or `worker123`) to edit.")
        
    st.markdown("---")
    st.header("🛠️ Dashboard Controls")
    stages = ["All", "Not Started", "In Progress", "Drilled", "Completed"]
    selected_stage = st.selectbox("Filter Site Plan by Stage", stages)
    
    st.markdown("---")
    st.header("💰 Commercial & BOQ Rates (HKSMM4)")
    concrete_unit_cost = st.number_input("C45/20D Concrete Rate (HKD/m³)", value=1200.0, step=50.0)
    soil_drill_rate = st.number_input("Soil / Alluvium Drilling (HKD/m)", value=2800.0, step=100.0)
    cdg_drill_rate = st.number_input("Extra-Over CDG Socket (HKD/m)", value=4500.0, step=100.0)
    rig_hourly_rate = st.number_input("Piling Rig Operating Cost (HKD/hr)", value=1800.0, step=100.0)

st.title("🏢 Enterprise Digital Twin: Scheduling, Engineering & QS Cost Control")

with st.expander("📋 View Method Statement & General Procedure"):
    st.markdown("""
    ### General Construction Procedure for Concrete Columns (610mm Ø)
    1. **Setting Out:** Surveyor sets out the column position (X/Y coordinates) based on approved construction drawings.
    2. **Casing Installation:** Pitch and drive the 610mm temporary steel casing into the ground to the required casing top and toe levels to prevent soil collapse.
    3. **Drilling & Excavation:** Drill through the soil and alluvium layers into the Completely Decomposed Granite (CDG) until the actual founding level is reached.
    4. **Inspection & Measurement:** Site Engineer records the *Actual Toe Level* and *Actual Founding Level*, determining the socket length into the CD material.
    5. **Hole Cleaning:** Clean the base of the drilled hole to remove loose debris and sediment.
    6. **Concreting:** Pour **C45/20D** grade concrete using a tremie pipe. The engineer records the *Actual Concrete Volume* and compares it against the theoretical volume to calculate the **Overbreak %**.
    7. **Casing Extraction:** Slowly extract the temporary casing, ensuring the final *Concrete Top Level* is maintained at least 1.0m above the design Cut-off Level.
    """)

st.markdown("---")

# --- DATA LOADING & SANITIZATION (FIXES GHOST ROW 242 & RESTORES THEORETICAL VOLUME) ---
@st.cache_data(show_spinner=False)
def load_data():
    if os.path.exists("FYP_Concrete_Column_Data_v2.csv"):
        df = pd.read_csv("FYP_Concrete_Column_Data_v2.csv")
    elif os.path.exists("FYP_Concrete_Column_Data.csv"):
        df = pd.read_csv("FYP_Concrete_Column_Data.csv")
    else:
        df = pd.DataFrame({"Concrete Column ID": ["BH2_1_C1", "BH2_1_C2"], "Construction_Stage": ["Not Started"]*2})
        
    # 1. Remove Ghost Summary Row (Row 242 where Concrete Column ID is NaN)
    df = df.dropna(subset=['Concrete Column ID']).copy()
    df = df[df['Concrete Column ID'].astype(str).str.strip() != ''].reset_index(drop=True)
    
    required_cols = [
        'Cut off level (mPD)', 'Tentative Concrete Column Bottom level \n(mPD)',
        'Actual Founding level', 'Foundation level\n(mPD)', 
        'Actual Concrete Volumn (m3)', 'X_Coord', 'Y_Coord',
        'Drill Start Date \n(DD/MM/YYYY)', 'Drill Completed Date\n(DD/MM/YYYY)', 
        'Concrete Date\n(DD/MM/YYYY)', 'Casing Top Level\n(mPD)', 
        'Measuring Depth\n(m)', 'Ground Level', 'Actual Toe Level\n(mPD)',
        'Alluvium Bottom\n(mPD)', 'Embedded Depth in CD Material \n(m)',
        'Concrete Column Length \n(m)', 'Estimate Concrete volume\n(m3)', 
        'As-built casing Length (m)', 'Concrete Top Level (mPD)',
        'Rig_Hours', 'NC_Shifts', 'No_of_Casings', 'Equipment_Damage', 'Repair_Cost_HKD',
        'Remark', 'Last_Edited_By', 'Timestamp_of_Edit'
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = np.nan
            
    # 2. Restore Missing Design Length & Theoretical Volume Formula (+3.0m top allowance)
    len_col = 'Concrete Column Length \n(m)'
    est_col = 'Estimate Concrete volume\n(m3)'
    if df[len_col].isna().any():
        m_len = df[len_col].isna()
        df.loc[m_len, len_col] = (
            pd.to_numeric(df.loc[m_len, 'Cut off level (mPD)'], errors='coerce') -
            pd.to_numeric(df.loc[m_len, 'Tentative Concrete Column Bottom level \n(mPD)'], errors='coerce')
        )
    if df[est_col].isna().all():
        df[est_col] = np.round(CROSS_AREA * (pd.to_numeric(df[len_col], errors='coerce') + 3.0), 4)
    else:
        m_est = df[est_col].isna()
        df.loc[m_est, est_col] = np.round(CROSS_AREA * (pd.to_numeric(df.loc[m_est, len_col], errors='coerce') + 3.0), 4)

    # 3. Auto-seed the 13 real site records from Standardized sheet if CSV is still unpopulated
    if (df['Construction_Stage'].fillna('Not Started') == 'Not Started').all():
        seed_records = [
            ("BH12_1_C11", "Completed", "2026-08-05", "2026-08-06", "2026-08-13", -54.430, 4.270, 59.52, 59.41, 3.61, 3.61, 25.0, 39.5, 1, 3.0, "No", 0.0, "", "Engineer", "2026-08-19 16:28:11"),
            ("BH12_1_C3",  "Drilled",   "2026-08-03", "2026-08-10", "",           -52.410, 4.690, 58.14, 59.54, 3.55, 3.55, 23.0, 58.0, 2, 5.0, "Yes", 15000.0, "Casing Broken", "Engineer", "2026-08-12 09:27:41"),
            ("BH12_1_C7",  "Completed", "2026-07-21", "2026-07-27", "2026-07-27", -56.675, 4.425, 61.11, 62.34, 3.59, 3.40, 22.0, 55.4, 0, 4.0, "No", 0.0, "", "Foreman", "2026-08-04 15:54:37"),
            ("BH12_2_C2",  "Completed", "2026-07-30", "2026-07-30", "2026-08-06", -55.180, 4.050, 59.23, 59.28, 3.54, 3.54, 24.0, 19.0, 0, 4.0, "No", 0.0, "", "Engineer", "2026-08-07 07:42:31"),
            ("BH12_2_C3",  "Completed", "2026-08-24", "2026-08-24", "2026-08-25", -46.000, 4.100, 50.10, 50.10, 3.61, 3.61, 30.0, 18.0, 0, 4.0, "No", 0.0, "", "Engineer", "2026-08-26 08:34:03"),
            ("BH13_1_C3",  "Completed", "2026-08-17", "2026-08-18", "2026-08-19", -51.240, 4.560, 55.80, 55.73, 3.68, 3.68, 25.0, 21.0, 0, 4.0, "No", 0.0, "", "Engineer", "2026-08-20 08:48:52"),
            ("BH22_1_C4",  "Completed", "2026-08-21", "2026-08-21", "2026-08-24", -54.130, 4.430, 58.56, 58.89, 3.86, 3.86, 24.0, 17.0, 0, 4.0, "No", 0.0, "", "Engineer", "2026-08-25 09:59:52"),
            ("BH22_2_C1",  "Completed", "2026-08-18", "2026-08-19", "2026-08-21", -54.790, 4.220, 59.60, 59.52, 3.60, 3.62, 22.0, 22.0, 0, 4.0, "No", 0.0, "", "Engineer", "2026-08-22 09:50:05"),
            ("BH5_1_C2",   "Completed", "2026-07-28", "2026-07-28", "2026-08-05", -34.730, 3.920, 39.25, 40.14, 3.54, 3.24, 19.0, 16.0, 0, 4.0, "No", 0.0, "", "Engineer", "2026-08-07 09:41:08"),
            ("BH5_1_C6",   "Completed", "2026-08-12", "2026-08-12", "2026-08-14", -45.990, 3.910, 50.75, 52.69, 3.51, 3.51, 23.0, 17.0, 0, 4.0, "No", 0.0, "", "Engineer", "2026-08-15 08:43:30"),
            ("BH5_2_C1",   "Completed", "2026-08-08", "2026-08-25", "2026-08-26", -32.090, 4.000, 36.09, 36.09, 3.52, 3.52, 18.0, 25.0, 1, 4.0, "No", 0.0, "", "Engineer", "2026-08-26 08:35:16"),
            ("BH5_2_C3",   "Completed", "2026-07-29", "2026-07-30", "2026-08-04", -33.083, 4.127, 38.11, 38.41, 3.65, 2.15, 17.0, 21.0, 0, 4.0, "No", 0.0, "", "Engineer", "2026-08-06 09:59:00"),
            ("BH5_2_C6",   "Completed", "2026-08-07", "2026-08-07", "2026-08-14", -34.100, 3.820, 38.51, 38.96, 3.52, 3.52, 19.0, 24.0, 1, 4.0, "No", 0.0, "", "Engineer", "2026-08-19 16:30:18"),
        ]
        for cid, stg, ds, de, cd, af, ct, md, cl, gl, ctop, avol, rhrs, ncs, cas, dmg, rcost, rem, ed, ts in seed_records:
            idx_m = df['Concrete Column ID'] == cid
            if idx_m.any():
                i = df[idx_m].index[0]
                df.loc[i, 'Construction_Stage'] = stg
                df.loc[i, 'Drill Start Date \n(DD/MM/YYYY)'] = ds
                if de: df.loc[i, 'Drill Completed Date\n(DD/MM/YYYY)'] = de
                if cd: df.loc[i, 'Concrete Date\n(DD/MM/YYYY)'] = cd
                df.loc[i, 'Actual Founding level'] = af
                df.loc[i, 'Casing Top Level\n(mPD)'] = ct
                df.loc[i, 'Measuring Depth\n(m)'] = md
                df.loc[i, 'As-built casing Length (m)'] = cl
                df.loc[i, 'Actual Toe Level\n(mPD)'] = round(ct - cl, 3)
                df.loc[i, 'Ground Level'] = gl
                df.loc[i, 'Concrete Top Level (mPD)'] = ctop
                df.loc[i, 'Actual Concrete Volumn (m3)'] = avol
                df.loc[i, 'Rig_Hours'] = rhrs
                df.loc[i, 'NC_Shifts'] = ncs
                df.loc[i, 'No_of_Casings'] = cas
                df.loc[i, 'Equipment_Damage'] = dmg
                df.loc[i, 'Repair_Cost_HKD'] = rcost
                df.loc[i, 'Remark'] = rem
                df.loc[i, 'Last_Edited_By'] = ed
                df.loc[i, 'Timestamp_of_Edit'] = ts

    df['Equipment_Damage'] = df['Equipment_Damage'].fillna('No')
    df['Repair_Cost_HKD'] = pd.to_numeric(df['Repair_Cost_HKD'], errors='coerce').fillna(0.0)
    df['Rig_Hours'] = pd.to_numeric(df['Rig_Hours'], errors='coerce').fillna(0.0)
    df['NC_Shifts'] = pd.to_numeric(df['NC_Shifts'], errors='coerce').fillna(0).astype(int)
    df['No_of_Casings'] = pd.to_numeric(df['No_of_Casings'], errors='coerce').fillna(4.0)
    
    df = df.sort_values(by='Concrete Column ID').reset_index(drop=True)
    grid_cols = int(np.ceil(np.sqrt(len(df)))) if len(df) > 0 else 1
    spacing = 15.0
    df['X_Coord'] = [(i % grid_cols) * spacing for i in range(len(df))]
    df['Y_Coord'] = [(i // grid_cols) * spacing for i in range(len(df))]
            
    date_cols = ['Concrete Date\n(DD/MM/YYYY)', 'Drill Start Date \n(DD/MM/YYYY)', 'Drill Completed Date\n(DD/MM/YYYY)']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
            
    if 'Construction_Stage' not in df.columns:
        df['Construction_Stage'] = "Not Started"
    return df

# --- SELF-CONTAINED QUANTITY SURVEYING (QS) & EVM ENGINE ---
def compute_commercial_metrics(df_in, conc_rate, soil_rate, cdg_rate, rig_rate):
    work_df = df_in.copy()
    cutoff = pd.to_numeric(work_df['Cut off level (mPD)'], errors='coerce').fillna(0.875)
    alluv = pd.to_numeric(work_df['Alluvium Bottom\n(mPD)'], errors='coerce').fillna(-31.0)
    tent_b = pd.to_numeric(work_df['Tentative Concrete Column Bottom level \n(mPD)'], errors='coerce').fillna(-41.0)
    d_len = pd.to_numeric(work_df['Concrete Column Length \n(m)'], errors='coerce').fillna(cutoff - tent_b)
    d_cdg = pd.to_numeric(work_df['Embedded Depth in CD Material \n(m)'], errors='coerce').fillna(alluv - tent_b)
    d_soil = (cutoff - alluv).clip(lower=0.0)
    
    theo_v = pd.to_numeric(work_df['Estimate Concrete volume\n(m3)'], errors='coerce').fillna(CROSS_AREA * (d_len + 3.0))
    act_f = pd.to_numeric(work_df['Actual Founding level'], errors='coerce')
    c_top = pd.to_numeric(work_df['Concrete Top Level (mPD)'], errors='coerce').fillna(cutoff + 3.0)
    act_v = pd.to_numeric(work_df['Actual Concrete Volumn (m3)'], errors='coerce').fillna(0.0)
    rig_h = pd.to_numeric(work_df['Rig_Hours'], errors='coerce').fillna(0.0)
    rep_c = pd.to_numeric(work_df['Repair_Cost_HKD'], errors='coerce').fillna(0.0)
    
    has_ab = act_f.notna()
    ab_conc_len = np.where(has_ab, (c_top - act_f).clip(lower=0.0), d_len + 3.0)
    ab_cdg_len = np.where(has_ab, (alluv - act_f).clip(lower=0.0), d_cdg)
    ab_drill_len = np.where(has_ab, (cutoff - act_f).clip(lower=1.0), d_len)
    ab_theo_v = ab_conc_len * CROSS_AREA
    
    is_comp = (work_df['Construction_Stage'] == 'Completed') & (act_v > 0)
    is_active = work_df['Construction_Stage'].isin(['Drilled', 'Completed'])
    
    work_df['Design_Soil_Len_m'] = np.round(d_soil, 3)
    work_df['Design_CDG_Len_m'] = np.round(d_cdg, 3)
    work_df['AsBuilt_CDG_Len_m'] = np.round(ab_cdg_len, 3)
    work_df['AsBuilt_Drill_Len_m'] = np.round(ab_drill_len, 3)
    work_df['AsBuilt_Cylinder_Vol_m3'] = np.round(ab_theo_v, 3)
    
    # 2A. Split Claimable BOQ Depth Variation vs. True Radial Overbreak
    work_df['Claimable_BOQ_Var_Vol_m3'] = np.where(is_comp, ab_theo_v - theo_v, 0.0)
    work_df['True_Radial_Overbreak_Vol_m3'] = np.where(is_comp, np.maximum(0.0, act_v - ab_theo_v), 0.0)
    work_df['Apparent_Overbreak_Pct'] = np.where(is_comp & (theo_v > 0), ((act_v - theo_v) / theo_v) * 100.0, 0.0)
    work_df['True_Overbreak_Pct'] = np.where(is_comp & (ab_theo_v > 0), ((act_v - ab_theo_v) / ab_theo_v) * 100.0, 0.0)
    
    # 2B. Multi-Rate BOQ Valuations
    work_df['Baseline_Budget_HKD'] = d_soil * soil_rate + d_cdg * cdg_rate + theo_v * conc_rate
    w = np.where(work_df['Construction_Stage'] == 'Completed', 1.0,
        np.where(work_df['Construction_Stage'] == 'Drilled', 0.60,
        np.where(work_df['Construction_Stage'] == 'In Progress', 0.25, 0.0)))
    work_df['Earned_Value_HKD'] = work_df['Baseline_Budget_HKD'] * w
    
    work_df['Claimable_BOQ_Value_HKD'] = np.where(
        work_df['Construction_Stage'] == 'Completed',
        d_soil * soil_rate + ab_cdg_len * cdg_rate + ab_theo_v * conc_rate,
        np.where(work_df['Construction_Stage'] == 'Drilled', d_soil * soil_rate + ab_cdg_len * cdg_rate, 0.0)
    )
    work_df['Actual_Cost_HKD'] = np.where(
        is_active,
        (d_soil * soil_rate * 0.85) + (ab_cdg_len * cdg_rate * 0.85) + (act_v * conc_rate) + (rig_h * rig_rate) + rep_c,
        0.0
    )
    work_df['Claimable_Var_HKD'] = work_df['Claimable_BOQ_Var_Vol_m3'] * conc_rate + np.where(is_active, (ab_cdg_len - d_cdg) * cdg_rate, 0.0)
    work_df['True_Wastage_Cost_HKD'] = work_df['True_Radial_Overbreak_Vol_m3'] * conc_rate
    
    # 2C. Plant Productivity & Rig Operating Cost per Drilled Meter (HK$/m)
    work_df['Drilling_Speed_m_hr'] = np.where(is_active & (rig_h > 0), ab_drill_len / rig_h, np.nan)
    work_df['Rig_Cost_per_m_HKD'] = np.where(is_active & (ab_drill_len > 0), (rig_h * rig_rate) / ab_drill_len, np.nan)
    
    total_ev = float(work_df['Earned_Value_HKD'].sum())
    total_ac = float(work_df['Actual_Cost_HKD'].sum())
    total_pv = float(work_df['Baseline_Budget_HKD'].iloc[:15].sum())
    
    qs_sum = {
        'total_bac_hkd': float(work_df['Baseline_Budget_HKD'].sum()),
        'total_pv_hkd': total_pv,
        'total_ev_hkd': total_ev,
        'total_ac_hkd': total_ac,
        'total_claimable_ipa_hkd': float(work_df['Claimable_BOQ_Value_HKD'].sum()),
        'cpi': total_ev / total_ac if total_ac > 0 else 1.0,
        'spi': total_ev / total_pv if total_pv > 0 else 1.0,
        'claimable_conc_var_m3': float(work_df['Claimable_BOQ_Var_Vol_m3'].sum()),
        'true_overbreak_m3': float(work_df['True_Radial_Overbreak_Vol_m3'].sum()),
        'claimable_var_hkd': float(work_df['Claimable_Var_HKD'].sum()),
        'true_wastage_hkd': float(work_df['True_Wastage_Cost_HKD'].sum()),
        'total_repair_hkd': float(rep_c.sum()),
        'damaged_casings_count': int((work_df['Equipment_Damage'].astype(str).str.lower() == 'yes').sum()),
        'total_nc_shifts': int(work_df['NC_Shifts'].sum()),
        'avg_drill_speed_m_hr': float(work_df['Drilling_Speed_m_hr'].dropna().mean()) if work_df['Drilling_Speed_m_hr'].notna().any() else 0.0,
        'avg_rig_cost_per_m': float(work_df['Rig_Cost_per_m_HKD'].dropna().mean()) if work_df['Rig_Cost_per_m_HKD'].notna().any() else 0.0,
        'total_rig_hours': float(rig_h.sum())
    }
    return work_df, qs_sum

# --- SELF-CONTAINED DOCUMENT & IPA GENERATORS ---
def format_mpd(val):
    try: return f"{float(val):+.3f}" if pd.notna(val) and str(val).strip() != "" else ""
    except Exception: return val

def format_float(val):
    try: return f"{float(val):.3f}" if pd.notna(val) and str(val).strip() != "" else ""
    except Exception: return val

def generate_excel_report(col_data):
    template_path = "Record Sheet for FYP_3.xlsx"
    if not os.path.exists(template_path): template_path = "Record Sheet for FYP.xlsx"
    if not os.path.exists(template_path): return None, None
    wb = openpyxl.load_workbook(template_path)
    sheet = wb.active
    sheet["H8"] = col_data.get('Concrete Column ID', '')
    sheet["H11"] = "610"; sheet["H12"] = "610"; sheet["H13"] = "0"
    sheet["H14"] = format_mpd(col_data.get('Alluvium Bottom\n(mPD)', np.nan))
    sheet["H15"] = format_mpd(col_data.get('Tentative Concrete Column Bottom level \n(mPD)', np.nan))
    sheet["M21"] = format_mpd(col_data.get('Cut off level (mPD)', np.nan))
    sheet["M17"] = format_mpd(col_data.get('Concrete Top Level (mPD)', np.nan))
    def format_date(d):
        if pd.isna(d) or str(d).strip() == "": return ""
        return d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') else str(d)[:10]
    sheet["H23"] = format_date(col_data.get('Drill Start Date \n(DD/MM/YYYY)', pd.NaT))
    sheet["H24"] = format_date(col_data.get('Drill Completed Date\n(DD/MM/YYYY)', pd.NaT))
    sheet["H25"] = format_mpd(col_data.get('Ground Level', np.nan))
    sheet["H26"] = format_mpd(col_data.get('Casing Top Level\n(mPD)', np.nan))
    sheet["H28"] = format_mpd(col_data.get('Actual Toe Level\n(mPD)', np.nan))
    sheet["H27"] = format_float(col_data.get('As-built casing Length (m)', np.nan))
    sheet["H29"] = format_mpd(col_data.get('Actual Founding level', np.nan))
    sheet["H30"] = ""; sheet["H31"] = ""; sheet["H36"] = "C45/20D"
    theo_vol = col_data.get('Estimate Concrete volume\n(m3)', np.nan)
    act_vol = col_data.get('Actual Concrete Volumn (m3)', np.nan)
    sheet["H37"] = format_float(theo_vol)
    sheet["H38"] = format_float(act_vol)
    try:
        if float(theo_vol) > 0 and pd.notna(act_vol):
            sheet["H39"] = format_float(((float(act_vol) - float(theo_vol)) / float(theo_vol)) * 100)
        else: sheet["H39"] = ""
    except Exception: sheet["H39"] = ""
    sheet["H40"] = format_date(col_data.get('Concrete Date\n(DD/MM/YYYY)', pd.NaT))
    virtual_file = io.BytesIO()
    wb.save(virtual_file)
    virtual_file.seek(0)
    return virtual_file.read(), f"{col_data['Concrete Column ID']}_Record.xlsx"

class PDFReport(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'Formal Concrete Column Submission (HKBD & QS Standard)', 0, 1, 'C')
        self.ln(3)
    def footer(self):
        self.set_y(-20)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()} | Generated by Enterprise Digital Twin', 0, 0, 'C')

def generate_pdf_report(row):
    pdf = PDFReport()
    pdf.add_page()
    col_id = row.get('Concrete Column ID', 'Unknown')
    ds_val = row.get('Drill Start Date \n(DD/MM/YYYY)', '')
    de_val = row.get('Drill Completed Date\n(DD/MM/YYYY)', '')
    cd_val = row.get('Concrete Date\n(DD/MM/YYYY)', '')
    gl_val = format_mpd(row.get('Ground Level', ''))
    co_val = format_mpd(row.get('Cut off level (mPD)', ''))
    ct_val = format_mpd(row.get('Casing Top Level\n(mPD)', ''))
    at_val = format_mpd(row.get('Actual Toe Level\n(mPD)', ''))
    tb_val = format_mpd(row.get('Tentative Concrete Column Bottom level \n(mPD)', ''))
    af_val = format_mpd(row.get('Actual Founding level', ''))
    ctl_val = format_mpd(row.get('Concrete Top Level (mPD)', ''))
    
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 9, f"Column Reference: {col_id}   |   Status: {row.get('Construction_Stage', '')}", 0, 1)
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 7, f"Drill Start: {ds_val} | Drill End: {de_val} | Concrete Date: {cd_val}", 0, 1)
    pdf.ln(3)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 8, "1. Surveying & Level Measurements (mPD)", 0, 1)
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 7, f"Ground Level: {gl_val} mPD   |   Cut-off Level: {co_val} mPD", 0, 1)
    pdf.cell(0, 7, f"Casing Top Level: {ct_val} mPD   |   Actual Toe Level: {at_val} mPD", 0, 1)
    pdf.cell(0, 7, f"Design Bottom: {tb_val} mPD   |   Actual Founding Level: {af_val} mPD", 0, 1)
    pdf.cell(0, 7, f"Concrete Top Level: {ctl_val} mPD", 0, 1)
    pdf.ln(3)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 8, "2. Commercial QS Volume & Variance Analysis (HKSMM4)", 0, 1)
    pdf.set_font("Arial", size=10)
    est_key = 'Estimate Concrete volume\n(m3)'
    t_vol = float(row.get(est_key, 0) or 0)
    ab_vol = float(row.get('AsBuilt_Cylinder_Vol_m3', t_vol) or t_vol)
    a_vol = float(row.get('Actual Concrete Volumn (m3)', 0) or 0)
    boq_var = float(row.get('Claimable_BOQ_Var_Vol_m3', 0) or 0)
    true_ob = float(row.get('True_Radial_Overbreak_Vol_m3', 0) or 0)
    pdf.cell(0, 7, f"Design Theoretical Volume (+3m allowance): {t_vol:.3f} m3", 0, 1)
    pdf.cell(0, 7, f"As-Built Cylinder Volume (Re-measured Depth): {ab_vol:.3f} m3", 0, 1)
    pdf.cell(0, 7, f"Actual Poured Concrete Volume: {a_vol:.3f} m3", 0, 1)
    pdf.cell(0, 7, f"Claimable BOQ Depth Variation: {boq_var:+.3f} m3   |   True Radial Overbreak (Wastage): {true_ob:.3f} m3", 0, 1)
    pdf.ln(3)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 8, "3. Plant Productivity & Audit Trail", 0, 1)
    pdf.set_font("Arial", size=10)
    rig_h = row.get('Rig_Hours', 0)
    nc_s = row.get('NC_Shifts', 0)
    cas_n = row.get('No_of_Casings', 4)
    eq_dmg = row.get('Equipment_Damage', 'No')
    rep_c = float(row.get('Repair_Cost_HKD', 0) or 0)
    pdf.cell(0, 7, f"Rig Hours: {rig_h} hrs (NC Shifts: {nc_s}) | Casings Used: {cas_n} | Damage: {eq_dmg} (HKD ${rep_c:,.2f})", 0, 1)
    pdf.cell(0, 7, f"Last Edited By: {row.get('Last_Edited_By', 'System/Unknown')}   |   Timestamp: {row.get('Timestamp_of_Edit', '')}", 0, 1)
    pdf.ln(12)
    pdf.cell(95, 8, "Prepared By (Site Engineer): ___________________", 0, 0)
    pdf.cell(95, 8, "Certified By (QS / Supervisor): ___________________", 0, 1)
    return pdf.output(dest='S').encode('latin-1'), f"{col_id}_Formal_Report.pdf"

def generate_ipa_workbook(work_df, qs_summary, rates, period_label="All Active / Completed Columns"):
    wb = openpyxl.Workbook()
    navy_header = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    slate_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    zebra_fill = PatternFill(start_color="F9FAFC", end_color="F9FAFC", fill_type="solid")
    accent_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    accent_red = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    font_title = Font(name="Calibri", size=14, bold=True, color="1F3864")
    font_sec = Font(name="Calibri", size=11, bold=True, color="1F3864")
    font_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True, color="000000")
    font_reg = Font(name="Calibri", size=10, color="000000")
    thin_side = Side(border_style="thin", color="D9D9D9")
    double_bottom = Side(border_style="double", color="1F3864")
    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    total_border = Border(top=thin_side, bottom=double_bottom, left=thin_side, right=thin_side)
    
    ws1 = wb.active
    ws1.title = "IPA Summary Certificate"
    ws1.views.sheetView[0].showGridLines = True
    ws1["A1"] = "INTERIM PAYMENT APPLICATION (IPA) & COMMERCIAL CONTROL SUMMARY"
    ws1["A1"].font = font_title
    ws1["A2"] = f"Project: 610mm Dia Socketed Concrete Columns (241 Clusters) | Scope: {period_label}"
    ws1["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")
    
    ws1["A4"] = "1. CONTRACT BILL OF QUANTITIES (BOQ) UNIT RATES (HKSMM4)"
    ws1["A4"].font = font_sec
    for col_idx, h in enumerate(["BOQ Item Ref", "Description", "Unit", "Contract Rate (HKD)"], 1):
        c = ws1.cell(5, col_idx, h)
        c.font = font_hdr; c.fill = navy_header; c.alignment = Alignment(horizontal="center")
    rate_rows = [
        ("Item A.1", "Bored Piling / Casing & Drilling through Soil / Alluvium", "m", rates.get('soil', 2800.0)),
        ("Item A.2", "Extra-Over (EO) Drilling into CDG / Rock Socket", "m", rates.get('cdg', 4500.0)),
        ("Item A.3", "Supply & Pour Grade C45/20D Concrete (Tremie Method)", "m3", rates.get('concrete', 1200.0)),
        ("Internal", "Piling Rig & Crew Operating Cost Allowance", "hr", rates.get('rig', 1800.0)),
    ]
    for r_idx, r_data in enumerate(rate_rows, 6):
        for c_idx, val in enumerate(r_data, 1):
            cell = ws1.cell(r_idx, c_idx, val)
            cell.font = font_reg; cell.border = grid_border
            if c_idx == 4: cell.number_format = '"HK$"#,##0.00'
            
    ws1["A12"] = "2. EARNED VALUE MANAGEMENT (EVM) & QS VARIANCE EXECUTIVE SUMMARY"
    ws1["A12"].font = font_sec
    for col_idx, h in enumerate(["Commercial Metric", "Current Value", "Benchmark / Status", "QS Analytical Note"], 1):
        c = ws1.cell(13, col_idx, h)
        c.font = font_hdr; c.fill = navy_header; c.alignment = Alignment(horizontal="center")
    kpi_rows = [
        ("Total Baseline Contract Budget (BAC - 241 Columns)", qs_summary['total_bac_hkd'], "100.0% Baseline", "Full contract scope across 10,243.37m design length", '"HK$"#,##0.00'),
        ("Planned Value to Date (PV)", qs_summary['total_pv_hkd'], "Schedule Target", "Baseline value of scheduled columns", '"HK$"#,##0.00'),
        ("Earned Value to Date (EV)", qs_summary['total_ev_hkd'], f"SPI = {qs_summary['spi']:.2f}", "Budgeted value of physical work accomplished", '"HK$"#,##0.00'),
        ("Gross Certified IPA Valuation (HKSMM4 Re-measured)", qs_summary['total_claimable_ipa_hkd'], "Billable to Client", "Includes claimable CDG depth & concrete variations", '"HK$"#,##0.00'),
        ("Actual Contractor Cost Incurred (AC)", qs_summary['total_ac_hkd'], f"CPI = {qs_summary['cpi']:.2f}", "Includes actual poured concrete, rig hrs & casing repairs", '"HK$"#,##0.00'),
        ("Claimable BOQ Depth Variation Value", qs_summary['claimable_var_hkd'], f"{qs_summary['claimable_conc_var_m3']:+.2f} m3 Conc", "Extra payment due to deeper actual founding levels", '"HK$"#,##0.00'),
        ("Non-Claimable True Overbreak Wastage Loss", qs_summary['true_wastage_hkd'], f"{qs_summary['true_overbreak_m3']:.2f} m3 Excess", "Contractor commercial risk (radial over-excavation)", '"HK$"#,##0.00'),
        ("Equipment & Casing Damage Repair Cost", qs_summary['total_repair_hkd'], f"{qs_summary['damaged_casings_count']} Incident(s)", "Logged on-site casing damage (e.g. BH12_1_C3)", '"HK$"#,##0.00'),
    ]
    for r_idx, (m_name, m_val, m_stat, m_note, num_fmt) in enumerate(kpi_rows, 14):
        c1 = ws1.cell(r_idx, 1, m_name); c2 = ws1.cell(r_idx, 2, m_val)
        c3 = ws1.cell(r_idx, 3, m_stat); c4 = ws1.cell(r_idx, 4, m_note)
        for c in (c1, c2, c3, c4):
            c.font = font_bold if "Gross Certified" in m_name or "Wastage" in m_name else font_reg
            c.border = grid_border
        c2.number_format = num_fmt
        if "Gross Certified" in m_name or "Claimable BOQ" in m_name: c2.fill = accent_green
        elif "Wastage" in m_name or "Damage" in m_name: c2.fill = accent_red

    ws2 = wb.create_sheet(title="Active Columns QS Schedule")
    ws2.views.sheetView[0].showGridLines = True
    if work_df.empty:
        active_df = pd.DataFrame(columns=['Concrete Column ID', 'Construction_Stage'])
    else:
        active_df = work_df.reset_index(drop=True)
    ws2["A1"] = f"SCHEDULE OF CONCRETE COLUMNS ({period_label}) - BOQ RE-MEASUREMENT & VARIANCE BREAKDOWN"
    ws2["A1"].font = font_title
    
    cols_tab2 = [
        ("Col ID", "Concrete Column ID", "@"),
        ("Stage", "Construction_Stage", "@"),
        ("Design Soil (m)", "Design_Soil_Len_m", "0.000"),
        ("Design CDG (m)", "Design_CDG_Len_m", "0.000"),
        ("As-Built CDG (m)", "AsBuilt_CDG_Len_m", "0.000"),
        ("Design Theo Vol (m3)", "Estimate Concrete volume\n(m3)", "0.000"),
        ("As-Built Cyl Vol (m3)", "AsBuilt_Cylinder_Vol_m3", "0.000"),
        ("Actual Poured (m3)", "Actual Concrete Volumn (m3)", "0.000"),
        ("Claimable Var Vol (m3)", "Claimable_BOQ_Var_Vol_m3", "+0.000;-0.000;0.000"),
        ("True Overbreak Vol (m3)", "True_Radial_Overbreak_Vol_m3", "0.000"),
        ("True Overbreak (%)", "True_Overbreak_Pct", "0.0"),
        ("Rig Hours (hr)", "Rig_Hours", "0.0"),
        ("Rig Cost/m (HKD/m)", "Rig_Cost_per_m_HKD", '"HK$"#,##0.00'),
        ("Repair Cost (HKD)", "Repair_Cost_HKD", '"HK$"#,##0.00'),
        ("Baseline Budget (HKD)", "Baseline_Budget_HKD", '"HK$"#,##0.00'),
        ("Certified IPA Value (HKD)", "Claimable_BOQ_Value_HKD", '"HK$"#,##0.00'),
        ("True Wastage Cost (HKD)", "True_Wastage_Cost_HKD", '"HK$"#,##0.00'),
        ("Actual Total Cost (HKD)", "Actual_Cost_HKD", '"HK$"#,##0.00'),
    ]
    for c_idx, (hdr_text, _, _) in enumerate(cols_tab2, 1):
        cell = ws2.cell(3, c_idx, hdr_text)
        cell.font = font_hdr; cell.fill = navy_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r_idx, row in active_df.iterrows():
        excel_r = r_idx + 4
        for c_idx, (_, col_key, num_fmt) in enumerate(cols_tab2, 1):
            val = row.get(col_key, "")
            if pd.isna(val): val = 0.0 if num_fmt != "@" else ""
            cell = ws2.cell(excel_r, c_idx, val)
            cell.font = font_reg; cell.border = grid_border
            if r_idx % 2 == 1: cell.fill = zebra_fill
            if num_fmt != "@": cell.number_format = num_fmt
            if col_key == "Claimable_BOQ_Value_HKD": cell.fill = accent_green
            elif col_key == "True_Wastage_Cost_HKD" and float(val or 0) > 0: cell.fill = accent_red
            
    tot_r = max(5, len(active_df) + 4)
    ws2.cell(tot_r, 1, "TOTAL / AVERAGE").font = font_bold
    ws2.cell(tot_r, 1).border = total_border; ws2.cell(tot_r, 1).fill = slate_sub
    ws2.cell(tot_r, 2, f"{len(active_df)} Columns").font = font_bold
    ws2.cell(tot_r, 2).border = total_border; ws2.cell(tot_r, 2).fill = slate_sub
    for c_idx in range(3, len(cols_tab2) + 1):
        col_let = get_column_letter(c_idx)
        _, col_key, num_fmt = cols_tab2[c_idx - 1]
        c = ws2.cell(tot_r, c_idx)
        if col_key in ["True_Overbreak_Pct", "Rig_Cost_per_m_HKD"]:
            c.value = f"=AVERAGE({col_let}4:{col_let}{tot_r-1})"
        else:
            c.value = f"=SUM({col_let}4:{col_let}{tot_r-1})"
        c.font = font_bold; c.border = total_border; c.fill = slate_sub
        c.number_format = num_fmt
        
    for ws in [ws1, ws2]:
        ws.freeze_panes = "A4" if ws == ws2 else "A6"
        for col in ws.columns:
            max_l = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(min(max_l + 4, 38), 13)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

# --- RUN APPLICATION ---
df = load_data()
df, qs_summary = compute_commercial_metrics(df, concrete_unit_cost, soil_drill_rate, cdg_drill_rate, rig_hourly_rate)

filtered_df = df.copy()
if selected_stage != "All":
    filtered_df = filtered_df[filtered_df['Construction_Stage'] == selected_stage]

if df.empty:
    st.error("No data.")
else:
    total_cols = len(df)
    completed_df = df[df['Construction_Stage'] == 'Completed'].copy()
    completed_cols = len(completed_df)
    remaining_cols = total_cols - completed_cols
    progress_pct = (completed_cols / total_cols) * 100 if total_cols > 0 else 0
    
    st.markdown("### 📈 Executive Engineering, Schedule & Commercial Control Panel")
    top_col1, top_col2, top_col3, top_col4 = st.columns([1, 1.15, 1.35, 1.2])
    
    with top_col1:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number", value=progress_pct,
            number={'suffix': "%", 'valueformat': '.1f', 'font': {'size': 34, 'color': '#2c3e50'}},
            title={'text': f"Completion ({completed_cols}/{total_cols})", 'font': {'size': 15}},
            gauge={'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': "darkblue"}, 'bar': {'color': "#27ae60"}, 'bgcolor': "#ecf0f1", 'shape': "angular"}
        ))
        fig_gauge.update_layout(height=195, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_gauge, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
    with top_col2:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown("**📊 Earned Value (EVM) & Rig KPIs**")
        st.metric("Gross Certified IPA (Billable)", f"HK${qs_summary['total_claimable_ipa_hkd']:,.0f}")
        st.metric("CPI / SPI Performance", f"CPI: {qs_summary['cpi']:.2f} | SPI: {qs_summary['spi']:.2f}", f"CV: HK${qs_summary['total_ev_hkd'] - qs_summary['total_ac_hkd']:,.0f}")
        st.caption(f"Speed: {qs_summary['avg_drill_speed_m_hr']:.2f} m/h | Rig Cost: HK${qs_summary['avg_rig_cost_per_m']:,.0f}/m | NC Shifts: {qs_summary['total_nc_shifts']}")
        st.markdown("</div>", unsafe_allow_html=True)

    with top_col3:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown("**💸 QS Variance Split (HKSMM4)**")
        st.markdown(f"<div class='claim-good'>✅ Claimable BOQ Variation: +HK${qs_summary['claimable_var_hkd']:,.0f} ({qs_summary['claimable_conc_var_m3']:+.2f} m³)</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='cost-alert'>🚨 True Overbreak Wastage: -HK${qs_summary['true_wastage_hkd']:,.0f} ({qs_summary['true_overbreak_m3']:.2f} m³)</div>", unsafe_allow_html=True)
        st.caption(f"Casing Repairs: HK${qs_summary['total_repair_hkd']:,.0f} ({qs_summary['damaged_casings_count']} broken) | Concrete @ HK${concrete_unit_cost:,.0f}/m³")
        st.markdown("</div>", unsafe_allow_html=True)
        
    with top_col4:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown("**📅 Period & Forecast Analysis**")
        today = datetime.date.today()
        date_range = st.date_input("Filter Production & IPA Date Range", value=(datetime.date(2026, 7, 20), today), format="YYYY-MM-DD")
        if len(date_range) == 2:
            s_d, e_d = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
            d_dates = pd.to_datetime(df['Drill Completed Date\n(DD/MM/YYYY)'], errors='coerce')
            c_dates = pd.to_datetime(df['Concrete Date\n(DD/MM/YYYY)'], errors='coerce')
            d_cnt = int(((d_dates >= s_d) & (d_dates <= e_d)).sum())
            c_cnt = int(((c_dates >= s_d) & (c_dates <= e_d)).sum())
            st.write(f"**In Period:** {d_cnt} Drilled | {c_cnt} Concreted")
        if not completed_df.empty:
            c_dt = pd.to_datetime(completed_df['Concrete Date\n(DD/MM/YYYY)'], errors='coerce').dropna()
            if len(c_dt) > 1:
                elap = max(1, (c_dt.max() - c_dt.min()).days)
                rate = len(c_dt) / elap
                proj = c_dt.max() + datetime.timedelta(days=remaining_cols / rate)
                st.info(f"**Rate:** {rate:.2f}/day → **Est. End:** {proj.strftime('%d %b %Y')}")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")

    col_vis, col_data = st.columns([2.4, 1.6])
    
    with col_vis:
        st.markdown("### 🗺️ Spatial Digital Twin, EVM S-Curve & Plant Productivity")
        color_map = {"Not Started": "lightgrey", "In Progress": "#f1c40f", "Drilled": "#3498db", "Completed": "#2ecc71"}
        tab1, tab2, tab3, tab4 = st.tabs([
            "2D Grid Plan View",
            "3D Column Foundation Model",
            "📈 Time-Phased EVM S-Curve",
            "💰 Column QS & Rig Cost/m"
        ])
        
        with tab1:
            if not filtered_df.empty:
                fig2d = px.scatter(
                    filtered_df, x='X_Coord', y='Y_Coord', color='Construction_Stage',
                    color_discrete_map=color_map, hover_name='Concrete Column ID',
                    hover_data=['Foundation level\n(mPD)', 'Actual Founding level', 'Claimable_BOQ_Value_HKD', 'True_Wastage_Cost_HKD']
                )
                fig2d.update_traces(marker=dict(size=15, line=dict(width=1, color='DarkSlateGrey')))
                fig2d.update_yaxes(scaleanchor="x", scaleratio=1)
                fig2d.update_layout(height=600, plot_bgcolor='white', paper_bgcolor='white')
                fig2d.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGrey')
                fig2d.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGrey')
                st.plotly_chart(fig2d, use_container_width=True)
            else:
                st.info("No data available.")
            
        with tab2:
            if not filtered_df.empty:
                fig3d = px.scatter_3d(
                    filtered_df, x='X_Coord', y='Y_Coord', z='Tentative Concrete Column Bottom level \n(mPD)',
                    color='Construction_Stage', color_discrete_map=color_map, hover_name='Concrete Column ID'
                )
                for _, row in filtered_df.iterrows():
                    bz = row['Actual Founding level'] if pd.notna(row['Actual Founding level']) else row['Tentative Concrete Column Bottom level \n(mPD)']
                    if pd.isna(bz): bz = -40.0
                    tz = row['Cut off level (mPD)'] if pd.notna(row['Cut off level (mPD)']) else 0.875
                    fig3d.add_scatter3d(
                        x=[row['X_Coord'], row['X_Coord']], y=[row['Y_Coord'], row['Y_Coord']], z=[tz, bz],
                        mode='lines', line=dict(color=color_map.get(row['Construction_Stage'], 'grey'), width=7), showlegend=False
                    )
                fig3d.update_layout(scene=dict(zaxis=dict(title='Depth (mPD)'), xaxis=dict(title='X (m)'), yaxis=dict(title='Y (m)')), height=600)
                st.plotly_chart(fig3d, use_container_width=True)
            else:
                st.info("No data available.")
                
        with tab3:
            # --- 2D: TRUE TIME-PHASED EARNED VALUE MANAGEMENT (EVM) S-CURVE ---
            c_dates_all = pd.to_datetime(df['Concrete Date\n(DD/MM/YYYY)'], errors='coerce')
            dr_dates_all = pd.to_datetime(df['Drill Completed Date\n(DD/MM/YYYY)'], errors='coerce')
            min_dt = pd.to_datetime("2026-07-20")
            max_dt = max(c_dates_all.dropna().max() if c_dates_all.notna().any() else pd.to_datetime("2026-08-30"), pd.to_datetime("2026-08-30"))
            timeline = pd.date_range(min_dt, max_dt, freq='D')
            pv_series = np.linspace(0, qs_summary['total_pv_hkd'], len(timeline))
            ev_series, ac_series, ipa_series = [], [], []
            for d_pt in timeline:
                c_m = (df['Construction_Stage'] == 'Completed') & (c_dates_all <= d_pt)
                dr_m = (df['Construction_Stage'].isin(['Drilled', 'Completed'])) & (dr_dates_all <= d_pt) & (~c_m)
                ev_series.append(df.loc[c_m, 'Baseline_Budget_HKD'].sum() + df.loc[dr_m, 'Baseline_Budget_HKD'].sum() * 0.60)
                ac_series.append(df.loc[c_m | dr_m, 'Actual_Cost_HKD'].sum())
                ipa_series.append(df.loc[c_m, 'Claimable_BOQ_Value_HKD'].sum() + df.loc[dr_m, 'Claimable_BOQ_Value_HKD'].sum() * 0.60)
            
            fig_scurve = go.Figure()
            fig_scurve.add_trace(go.Scatter(x=timeline, y=pv_series, mode='lines', name='Planned Value (PV)', line=dict(color='#7f8c8d', dash='dash', width=2.5)))
            fig_scurve.add_trace(go.Scatter(x=timeline, y=ev_series, mode='lines+markers', name='Earned Value (EV)', line=dict(color='#2980b9', width=3)))
            fig_scurve.add_trace(go.Scatter(x=timeline, y=ipa_series, mode='lines+markers', name='Certified BOQ IPA (Billable)', line=dict(color='#27ae60', width=3)))
            fig_scurve.add_trace(go.Scatter(x=timeline, y=ac_series, mode='lines+markers', name='Actual Contractor Cost (AC)', line=dict(color='#c0392b', width=3)))
            fig_scurve.update_layout(
                title="Cumulative Earned Value Management (EVM) S-Curve (PV vs. EV vs. Billable IPA vs. AC)",
                xaxis_title="Project Timeline", yaxis_title="Cumulative Value / Cost (HKD $)",
                height=560, plot_bgcolor='white', hovermode='x unified'
            )
            st.plotly_chart(fig_scurve, use_container_width=True)

        with tab4:
            active_viz = df[df['Construction_Stage'].isin(['Drilled', 'Completed'])].copy()
            if not active_viz.empty:
                fig_qs = go.Figure()
                fig_qs.add_trace(go.Bar(name='Baseline Budget (EV)', x=active_viz['Concrete Column ID'], y=active_viz['Baseline_Budget_HKD'], marker_color='#2980b9'))
                fig_qs.add_trace(go.Bar(name='Certified IPA Value (Billable)', x=active_viz['Concrete Column ID'], y=active_viz['Claimable_BOQ_Value_HKD'], marker_color='#27ae60'))
                fig_qs.add_trace(go.Bar(name='Actual Contractor Cost (AC)', x=active_viz['Concrete Column ID'], y=active_viz['Actual_Cost_HKD'], marker_color='#c0392b'))
                fig_qs.update_layout(
                    barmode='group', title="Column-by-Column Valuation: Baseline vs. Certified IPA vs. Actual Cost (HKD)",
                    yaxis_title="HKD ($)", height=320, plot_bgcolor='white'
                )
                st.plotly_chart(fig_qs, use_container_width=True)
                
                fig_rig = px.bar(
                    active_viz.sort_values('Rig_Cost_per_m_HKD', ascending=False),
                    x='Concrete Column ID', y='Rig_Cost_per_m_HKD', color='Equipment_Damage',
                    color_discrete_map={'Yes': '#c0392b', 'No': '#34495e'},
                    hover_data=['Rig_Hours', 'NC_Shifts', 'No_of_Casings', 'Drilling_Speed_m_hr', 'Remark'],
                    title="Rig Operating Cost per Drilled Meter (HK$/m) & Casing Damage Highlight"
                )
                fig_rig.update_layout(height=300, plot_bgcolor='white', yaxis_title="Rig Cost / m (HKD)")
                st.plotly_chart(fig_rig, use_container_width=True)
            else:
                st.info("Complete columns to view commercial breakdown.")

    with col_data:
        st.markdown("### 📝 Record As-Built & Plant Data")
        
        if st.session_state['role'] == 'Viewer':
            st.warning("🔒 Viewer Mode: Log in via sidebar (`admin123` or `worker123`) to update records or export IPA sheets.")
        else:
            valid_cols = df['Concrete Column ID'].dropna().unique()
            selected_col = st.selectbox("Select Column ID to Update", valid_cols)
            col_design_data = df[df['Concrete Column ID'] == selected_col].iloc[0]
            
            with st.form("update_form"):
                cur_stg = col_design_data.get('Construction_Stage', 'Not Started')
                if cur_stg not in ["Not Started", "In Progress", "Drilled", "Completed"]: cur_stg = "Not Started"
                new_stage = st.selectbox("Update Status", ["Not Started", "In Progress", "Drilled", "Completed"], index=["Not Started", "In Progress", "Drilled", "Completed"].index(cur_stg))
                
                f_col1, f_col2 = st.columns(2)
                def get_val(key, default=0.0):
                    v = col_design_data.get(key, default)
                    return float(v) if pd.notna(v) and str(v).strip() != "" else default
                
                with f_col1:
                    ground_level = st.number_input("Ground Level (mPD)", value=get_val('Ground Level', 3.50), format="%.3f")
                    casing_top = st.number_input("Casing Top Level (mPD)", value=get_val('Casing Top Level\n(mPD)', 4.00), format="%.3f")
                    act_toe = st.number_input("Actual Toe Level (mPD)", value=get_val('Actual Toe Level\n(mPD)', -50.00), format="%.3f")
                    rig_hrs_in = st.number_input("Rig Operating Hours (hr)", value=get_val('Rig_Hours', 18.0), step=1.0, format="%.1f")
                    nc_shifts_in = st.number_input("Incomplete Shifts (Work Type = NC)", value=int(get_val('NC_Shifts', 0)), step=1, min_value=0)
                with f_col2:
                    act_founding = st.number_input("Actual Founding Lvl (mPD)", value=get_val('Actual Founding level', -41.00), format="%.3f")
                    as_built_casing = st.number_input("As-built Casing L. (m)", value=get_val('As-built casing Length (m)', 54.00), format="%.3f")
                    conc_top = st.number_input("Concrete Top Lvl (mPD)", value=get_val('Concrete Top Level (mPD)', 3.50), format="%.3f")
                    concrete_vol = st.number_input("Actual Concrete Vol (m³)", value=get_val('Actual Concrete Volumn (m3)', 0.0), format="%.3f")
                    casings_used_in = st.number_input("No. of Casings Used", value=get_val('No_of_Casings', 4.0), step=1.0, min_value=1.0)
                
                st.markdown("##### 🔧 Plant Downtime, Casing Damage & Photo Log")
                p_col1, p_col2 = st.columns(2)
                with p_col1:
                    eq_dmg_in = st.selectbox("Plant Breakdown / Casing Damaged?", ["No", "Yes"], index=1 if str(col_design_data.get('Equipment_Damage', 'No')).lower() == 'yes' else 0)
                with p_col2:
                    rep_cost_in = st.number_input("Damage / Repair Cost (HKD)", value=get_val('Repair_Cost_HKD', 0.0), step=1000.0)
                remark_in = st.text_input("Engineering / QS Remark (e.g. Casing Broken)", value=str(col_design_data.get('Remark', '') or ''))
                site_photo = st.file_uploader("📷 Upload Casing / Pouring Photo (Optional)", type=['jpg', 'png', 'jpeg'])
                
                d_col1, d_col2, d_col3 = st.columns(3)
                with d_col1: start_date = st.date_input("Drill Start")
                with d_col2: comp_date = st.date_input("Drill End")
                with d_col3: conc_date = st.date_input("Concrete Pour")
                
                submitted = st.form_submit_button("✅ Update System & QS Record")
                
                if submitted:
                    if act_founding > ground_level:
                        st.error("🚨 **QA/QC Error:** Actual Founding Level cannot be higher than Ground Level. Submission blocked.")
                    else:
                        idx_m = df['Concrete Column ID'] == selected_col
                        df.loc[idx_m, 'Construction_Stage'] = new_stage
                        df.loc[idx_m, 'Ground Level'] = ground_level
                        df.loc[idx_m, 'Casing Top Level\n(mPD)'] = casing_top
                        df.loc[idx_m, 'Actual Toe Level\n(mPD)'] = act_toe
                        df.loc[idx_m, 'Actual Founding level'] = act_founding
                        df.loc[idx_m, 'Actual Concrete Volumn (m3)'] = concrete_vol
                        df.loc[idx_m, 'As-built casing Length (m)'] = as_built_casing
                        df.loc[idx_m, 'Concrete Top Level (mPD)'] = conc_top
                        df.loc[idx_m, 'Rig_Hours'] = rig_hrs_in
                        df.loc[idx_m, 'NC_Shifts'] = int(nc_shifts_in)
                        df.loc[idx_m, 'No_of_Casings'] = casings_used_in
                        df.loc[idx_m, 'Equipment_Damage'] = eq_dmg_in
                        df.loc[idx_m, 'Repair_Cost_HKD'] = rep_cost_in
                        df.loc[idx_m, 'Remark'] = remark_in
                        df.loc[idx_m, 'Drill Start Date \n(DD/MM/YYYY)'] = start_date
                        if new_stage in ["Drilled", "Completed"]:
                            df.loc[idx_m, 'Drill Completed Date\n(DD/MM/YYYY)'] = comp_date
                        if new_stage == "Completed":
                            df.loc[idx_m, 'Concrete Date\n(DD/MM/YYYY)'] = conc_date
                        df.loc[idx_m, 'Last_Edited_By'] = st.session_state['username']
                        df.loc[idx_m, 'Timestamp_of_Edit'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        save_path = "FYP_Concrete_Column_Data_v2.csv" if os.path.exists("FYP_Concrete_Column_Data_v2.csv") else "FYP_Concrete_Column_Data.csv"
                        df.to_csv(save_path, index=False)
                        st.cache_data.clear()
                        st.success(f"Column {selected_col} updated! QS & Audit logs refreshed.")
                        st.rerun()

            # --- SELF-CONTAINED EXPORT DOCUMENTATION (ALL 4 EXPORTS ALWAYS WORK) ---
            st.markdown("### 📄 Commercial & Engineering Exports")
            if st.session_state['role'] == 'Admin':
                ex1, ex2, ex3, ex4 = st.tabs(["💰 Monthly IPA (.xlsx)", "📄 HKBD PDF", "📥 Single Excel", "📦 Batch ZIP"])
                with ex1:
                    filter_ipa_by_date = st.checkbox("Filter IPA Schedule by Selected Date Range above", value=False)
                    if filter_ipa_by_date and len(date_range) == 2:
                        s_d, e_d = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
                        c_dt = pd.to_datetime(df['Concrete Date\n(DD/MM/YYYY)'], errors='coerce')
                        dr_dt = pd.to_datetime(df['Drill Completed Date\n(DD/MM/YYYY)'], errors='coerce')
                        p_mask = ((c_dt >= s_d) & (c_dt <= e_d)) | ((dr_dt >= s_d) & (dr_dt <= e_d))
                        ipa_sub_df = df[p_mask & (df['Construction_Stage'] != 'Not Started')].copy()
                        p_label = f"Period {date_range[0]} to {date_range[1]}"
                    else:
                        ipa_sub_df = df[df['Construction_Stage'] != 'Not Started'].copy()
                        p_label = "All Active & Completed Columns"
                        
                    ipa_bytes = generate_ipa_workbook(
                        ipa_sub_df, qs_summary,
                        {'soil': soil_drill_rate, 'cdg': cdg_drill_rate, 'concrete': concrete_unit_cost, 'rig': rig_hourly_rate},
                        period_label=p_label
                    )
                    st.download_button("📊 Download Interim Payment Application (.xlsx)", ipa_bytes, f"IPA_Valuation_{datetime.date.today()}.xlsx")
                with ex2:
                    pdf_b, pdf_n = generate_pdf_report(df[df['Concrete Column ID'] == selected_col].iloc[0])
                    st.download_button("📥 Download Formal HKBD & QS PDF", pdf_b, pdf_n, mime="application/pdf")
                with ex3:
                    ex_b, ex_n = generate_excel_report(df[df['Concrete Column ID'] == selected_col].iloc[0])
                    if ex_b:
                        st.download_button("📥 Download Single Record Sheet", ex_b, ex_n)
                    else:
                        st.error("Upload `Record Sheet for FYP.xlsx` to GitHub to enable template export.")
                with ex4:
                    if st.button("Generate Completed Columns Batch ZIP"):
                        comp_rows = df[df['Construction_Stage'] == 'Completed']
                        if comp_rows.empty:
                            st.warning("No completed columns to export.")
                        else:
                            zip_buf = io.BytesIO()
                            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
                                for _, r_row in comp_rows.iterrows():
                                    fb, fn = generate_excel_report(r_row)
                                    if fb: zf.writestr(fn, fb)
                                    pb, pn = generate_pdf_report(r_row)
                                    if pb: zf.writestr(pn, pb)
                            st.download_button("📦 Download ZIP Archive (Excel + PDF)", zip_buf.getvalue(), f"Batch_Export_{datetime.date.today()}.zip", mime="application/zip")
            else:
                st.info("Commercial IPA, PDF & Batch exports require Admin privileges (`admin123`).")

    # --- ADVANCED DATA VIEW ---
    st.markdown("---")
    with st.expander("📂 View Complete Database (Engineering, QS Valuation, Plant & Audit Trail)"):
        cols_show = [c for c in [
            'Concrete Column ID', 'Construction_Stage', 'Ground Level', 'Actual Founding level',
            'Estimate Concrete volume\n(m3)', 'AsBuilt_Cylinder_Vol_m3', 'Actual Concrete Volumn (m3)',
            'Claimable_BOQ_Var_Vol_m3', 'True_Radial_Overbreak_Vol_m3', 'Claimable_BOQ_Value_HKD',
            'True_Wastage_Cost_HKD', 'Rig_Hours', 'NC_Shifts', 'No_of_Casings', 'Drilling_Speed_m_hr',
            'Rig_Cost_per_m_HKD', 'Equipment_Damage', 'Repair_Cost_HKD', 'Remark',
            'Last_Edited_By', 'Timestamp_of_Edit'
        ] if c in filtered_df.columns]
        st.dataframe(filtered_df[cols_show], use_container_width=True, height=340)
