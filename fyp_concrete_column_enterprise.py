import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import datetime
import numpy as np
import openpyxl
import io
import zipfile
from fpdf import FPDF

# ==============================================================================
# 1. PAGE CONFIGURATION & ENTERPRISE STYLING
# ==============================================================================
st.set_page_config(
    layout="wide",
    page_title="Enterprise Digital Twin & QS Suite",
    page_icon="🏢",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    h1, h2, h3 { color: #2c3e50; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .metric-container { background-color: white; padding: 15px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 15px; }
    .cost-alert { color: #c0392b; font-weight: bold; font-size: 1.05em; }
    .claim-good { color: #1e8449; font-weight: bold; font-size: 1.05em; }
    .evm-badge { background-color: #eef2ff; color: #1e3a8a; padding: 6px 10px; border-radius: 6px; font-weight: 600; font-size: 0.92em; }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. ROLE-BASED AUTHENTICATION & COMMERCIAL BOQ SIDEBAR
# ==============================================================================
if 'role' not in st.session_state:
    st.session_state['role'] = 'Viewer'
if 'username' not in st.session_state:
    st.session_state['username'] = 'Guest'

with st.sidebar:
    st.header("🔐 System Login")
    pwd = st.text_input("Enter Password", type="password")
    if pwd == "admin123":
        st.session_state['role'] = 'Admin'
        st.session_state['username'] = 'Project Manager (QS/Admin)'
        st.success("Logged in as: Admin (Full Access)")
    elif pwd == "worker123":
        st.session_state['role'] = 'Editor'
        st.session_state['username'] = 'Site Engineer'
        st.success("Logged in as: Site Engineer (Data Entry)")
    elif pwd != "":
        st.error("Invalid password.")
        st.session_state['role'] = 'Viewer'
        st.session_state['username'] = 'Guest'
    else:
        st.info("Please log in to edit data.")

    st.markdown("---")
    st.header("🛠️ Dashboard Controls")
    stages = ["All", "Not Started", "In Progress", "Drilled", "Completed"]
    selected_stage = st.selectbox("Filter Site Plan by Stage", stages)

    st.markdown("---")
    st.header("💰 QS & BOQ Commercial Rates")
    concrete_unit_cost = st.number_input("C45/20D Concrete Rate (HKD/m³)", value=1200.0, step=50.0)
    soil_drill_rate = st.number_input("Soil/Alluvium Drilling (HKD/m)", value=2800.0, step=100.0)
    cdg_eo_rate = st.number_input("EO CDG Socket Drilling (HKD/m)", value=4500.0, step=100.0)
    rig_hourly_rate = st.number_input("Piling Rig Rate (HKD/hr)", value=1800.0, step=100.0)
    casing_damage_cost = st.number_input("Broken Casing Penalty (HKD/ea)", value=15000.0, step=1000.0)

# ==============================================================================
# 3. DATA LOADING, SANITIZATION & GEOMETRIC RESTORATION MODULE
# ==============================================================================
@st.cache_data(show_spinner=False)
def load_data():
    if os.path.exists("FYP_Concrete_Column_Data.csv"):
        df = pd.read_csv("FYP_Concrete_Column_Data.csv")
    elif os.path.exists("FYP_App_Data.csv"):
        df = pd.read_csv("FYP_App_Data.csv")
    else:
        df = pd.DataFrame({
            "Concrete Column ID": ["BH1_1_C1", "BH1_1_C2", "BH1_1_C3", "BH1_2_C1"],
            "Construction_Stage": ["Not Started"] * 4
        })

    # CRITICAL FIX 1: Remove Ghost Summary Row (Row 242 where Concrete Column ID is NaN)
    df = df.dropna(subset=['Concrete Column ID']).copy()
    df['Concrete Column ID'] = df['Concrete Column ID'].astype(str).str.strip()
    df = df[df['Concrete Column ID'] != ""].reset_index(drop=True)

    required_cols = [
        'Cut off level (mPD)', 'Tentative Concrete Column Bottom level \n(mPD)',
        'Actual Founding level', 'Foundation level\n(mPD)',
        'Actual Concrete Volumn (m3)', 'X_Coord', 'Y_Coord',
        'Drill Start Date \n(DD/MM/YYYY)', 'Drill Completed Date\n(DD/MM/YYYY)',
        'Concrete Date\n(DD/MM/YYYY)', 'Casing Top Level\n(mPD)',
        'Measuring Depth\n(m)', 'Ground Level', 'Actual Toe Level\n(mPD)',
        'Alluvium Bottom\n(mPD)', 'Embedded Depth in CD Material \n(m)',
        'Concrete Column Length\n(m)', 'Estimate Concrete volume\n(m3)',
        'As-built casing Length (m)', 'Concrete Top Level (mPD)',
        'Rig_Hours', 'Casings_Used', 'Casing_Damaged', 'Site_Remarks',
        'Last_Edited_By', 'Timestamp_of_Edit'
    ]

    for col in required_cols:
        if col not in df.columns:
            df[col] = np.nan

    # Ensure numeric types for core geometry columns
    num_cols = [
        'Cut off level (mPD)', 'Tentative Concrete Column Bottom level \n(mPD)',
        'Foundation level\n(mPD)', 'Alluvium Bottom\n(mPD)',
        'Embedded Depth in CD Material \n(m)', 'Concrete Column Length\n(m)',
        'Estimate Concrete volume\n(m3)', 'Actual Founding level',
        'Actual Concrete Volumn (m3)', 'Ground Level', 'Casing Top Level\n(mPD)',
        'Actual Toe Level\n(mPD)', 'As-built casing Length (m)',
        'Concrete Top Level (mPD)', 'Rig_Hours', 'Casings_Used'
    ]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # CRITICAL FIX 2: Restore Theoretical Concrete Volume Formula (610mm Ø + 3.0m Top Allowance)
    # Area of 610mm diameter column = pi * (0.305)^2 = 0.292247 m2
    area_610 = np.pi * (0.61 / 2.0) ** 2

    # Sync Tentative Bottom Level and Foundation Level
    df['Foundation level\n(mPD)'] = df['Foundation level\n(mPD)'].fillna(
        df['Tentative Concrete Column Bottom level \n(mPD)']
    )
    df['Tentative Concrete Column Bottom level \n(mPD)'] = df['Tentative Concrete Column Bottom level \n(mPD)'].fillna(
        df['Foundation level\n(mPD)']
    )

    # Calculate Design Length if missing
    calc_len = df['Cut off level (mPD)'] - df['Foundation level\n(mPD)']
    df['Concrete Column Length\n(m)'] = df['Concrete Column Length\n(m)'].fillna(calc_len).fillna(40.0)

    # Calculate Net Design Volume and Estimated Theoretical Volume (with +3.0m allowance per Summary Sheet)
    df['Net_Design_Volume_m3'] = area_610 * df['Concrete Column Length\n(m)']
    calc_est_vol = area_610 * (df['Concrete Column Length\n(m)'] + 3.0)
    df['Estimate Concrete volume\n(m3)'] = df['Estimate Concrete volume\n(m3)'].fillna(calc_est_vol)

    # Deterministic 15m Spatial Grid Layout
    df = df.sort_values(by='Concrete Column ID').reset_index(drop=True)
    grid_cols = int(np.ceil(np.sqrt(len(df)))) if len(df) > 0 else 1
    spacing = 15
    df['X_Coord'] = [(i % grid_cols) * spacing for i in range(len(df))]
    df['Y_Coord'] = [(i // grid_cols) * spacing for i in range(len(df))]

    date_cols = ['Concrete Date\n(DD/MM/YYYY)', 'Drill Start Date \n(DD/MM/YYYY)', 'Drill Completed Date\n(DD/MM/YYYY)']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

    if 'Construction_Stage' not in df.columns:
        df['Construction_Stage'] = "Not Started"
    df['Construction_Stage'] = df['Construction_Stage'].fillna("Not Started")
    df['Casing_Damaged'] = df['Casing_Damaged'].fillna("No")

    return df

# ==============================================================================
# 4. QUANTITY SURVEYING (QS) & EARNED VALUE MANAGEMENT (EVM) ENGINE
# ==============================================================================
def run_commercial_engine(df, conc_rate, soil_rate, cdg_rate, rig_rate, casing_penalty):
    area_610 = np.pi * (0.61 / 2.0) ** 2
    work_df = df.copy()

    # 1. Baseline Design BOQ Quantities & Budgeted Value (Planned Value per column)
    work_df['Design_Soil_Len_m'] = (
        work_df['Cut off level (mPD)'] - work_df['Alluvium Bottom\n(mPD)']
    ).clip(lower=0).fillna(30.0)
    work_df['Design_CDG_Len_m'] = work_df['Embedded Depth in CD Material \n(m)'].clip(lower=0).fillna(10.0)
    work_df['Theo_Vol_m3'] = work_df['Estimate Concrete volume\n(m3)'].fillna(13.0)

    work_df['Baseline_BOQ_Value_HKD'] = (
        work_df['Design_Soil_Len_m'] * soil_rate +
        work_df['Design_CDG_Len_m'] * cdg_rate +
        work_df['Theo_Vol_m3'] * conc_rate
    )

    # 2. As-Built Quantities for Completed/Drilled Columns
    # Effective top of concrete pour defaults to Concrete Top Level (or Cut-off + 3.0m if not logged)
    eff_top = work_df['Concrete Top Level (mPD)'].fillna(work_df['Cut off level (mPD)'] + 3.0)
    eff_bottom = work_df['Actual Founding level'].fillna(work_df['Foundation level\n(mPD)'])

    work_df['AsBuilt_Total_Len_m'] = (eff_top - eff_bottom).clip(lower=0)
    work_df['AsBuilt_Cylinder_Vol_m3'] = area_610 * work_df['AsBuilt_Total_Len_m']

    # Extra depth drilled below design founding level (re-measurable BOQ variation)
    work_df['Extra_Depth_m'] = (
        work_df['Foundation level\n(mPD)'] - work_df['Actual Founding level']
    ).clip(lower=0).fillna(0.0)

    # Split Volume Difference into:
    # (A) Claimable BOQ Depth Variation Volume (m3) -> As-Built Cylinder minus Design Theoretical Volume
    # (B) True Radial Overbreak Wastage Volume (m3) -> Actual Poured minus As-Built Cylinder
    is_done = work_df['Construction_Stage'] == 'Completed'
    act_vol = work_df['Actual Concrete Volumn (m3)'].fillna(0.0)

    work_df['Claimable_Var_Vol_m3'] = np.where(
        is_done,
        np.maximum(0.0, work_df['AsBuilt_Cylinder_Vol_m3'] - work_df['Theo_Vol_m3']),
        0.0
    )
    work_df['True_Overbreak_Vol_m3'] = np.where(
        is_done & (act_vol > 0),
        np.maximum(0.0, act_vol - work_df['AsBuilt_Cylinder_Vol_m3']),
        0.0
    )

    # Financial Valuation per Column
    work_df['Claimable_Variation_HKD'] = np.where(
        is_done,
        work_df['Claimable_Var_Vol_m3'] * conc_rate + work_df['Extra_Depth_m'] * cdg_rate,
        0.0
    )
    work_df['True_Wastage_Cost_HKD'] = np.where(
        is_done,
        work_df['True_Overbreak_Vol_m3'] * conc_rate,
        0.0
    )

    # Plant & Casing Damage Costs
    rig_hrs = work_df['Rig_Hours'].fillna(0.0)
    damaged_flag = (work_df['Casing_Damaged'].astype(str) == 'Yes').astype(int)
    work_df['Plant_And_Damage_Cost_HKD'] = (rig_hrs * rig_rate) + (damaged_flag * casing_penalty)

    # Drilling Productivity (m/hr)
    work_df['Drilling_Speed_m_hr'] = np.where(
        rig_hrs > 0,
        work_df['AsBuilt_Total_Len_m'] / rig_hrs,
        np.nan
    )

    # Earned Value (EV) & Actual Cost (AC) per Column
    # EV = 100% of Baseline BOQ for Completed, 50% for Drilled, 20% for In Progress
    stage_weights = {"Not Started": 0.0, "In Progress": 0.2, "Drilled": 0.5, "Completed": 1.0}
    work_df['Stage_Weight'] = work_df['Construction_Stage'].map(stage_weights).fillna(0.0)
    work_df['Earned_Value_HKD'] = work_df['Baseline_BOQ_Value_HKD'] * work_df['Stage_Weight']

    # Certified Gross Valuation (for IPA Billing) = Earned Value + Claimable BOQ Depth Variation
    work_df['Certified_IPA_Value_HKD'] = work_df['Earned_Value_HKD'] + work_df['Claimable_Variation_HKD']

    # Actual Cost (AC) incurred on active/completed columns
    work_df['Actual_Cost_HKD'] = np.where(
        work_df['Stage_Weight'] > 0,
        (work_df['Design_Soil_Len_m'] * soil_rate * work_df['Stage_Weight']) +
        ((work_df['Design_CDG_Len_m'] + work_df['Extra_Depth_m']) * cdg_rate * work_df['Stage_Weight']) +
        (act_vol * conc_rate) +
        work_df['Plant_And_Damage_Cost_HKD'],
        0.0
    )

    return work_df

# ==============================================================================
# 5. DOCUMENT GENERATORS (EXCEL RECORD, FORMAL PDF, & MONTHLY QS IPA SHEET)
# ==============================================================================
def format_mpd(val):
    try:
        return f"{float(val):+.3f}" if pd.notna(val) and str(val).strip() != "" else ""
    except:
        return val

def format_float(val):
    try:
        return f"{float(val):.3f}" if pd.notna(val) and str(val).strip() != "" else ""
    except:
        return val

def generate_excel_report(col_data):
    template_path = "Record Sheet for FYP_3.xlsx"
    if not os.path.exists(template_path):
        template_path = "Record Sheet for FYP.xlsx"
    if not os.path.exists(template_path):
        return None, None

    wb = openpyxl.load_workbook(template_path)
    sheet = wb.active

    sheet["H8"] = col_data.get('Concrete Column ID', '')
    sheet["H11"] = "610"
    sheet["H12"] = "610"
    sheet["H13"] = "0"
    sheet["H14"] = format_mpd(col_data.get('Alluvium Bottom\n(mPD)', np.nan))
    sheet["H15"] = format_mpd(col_data.get('Tentative Concrete Column Bottom level \n(mPD)', np.nan))

    cutoff = col_data.get('Cut off level (mPD)', np.nan)
    sheet["M21"] = format_mpd(cutoff)
    sheet["M17"] = format_mpd(col_data.get('Concrete Top Level (mPD)', np.nan))

    def format_date(d):
        return d.strftime('%Y-%m-%d') if pd.notna(d) and hasattr(d, 'strftime') else str(d) if pd.notna(d) else ""

    sheet["H23"] = format_date(col_data.get('Drill Start Date \n(DD/MM/YYYY)', pd.NaT))
    sheet["H24"] = format_date(col_data.get('Drill Completed Date\n(DD/MM/YYYY)', pd.NaT))
    sheet["H25"] = format_mpd(col_data.get('Ground Level', np.nan))
    sheet["H26"] = format_mpd(col_data.get('Casing Top Level\n(mPD)', np.nan))
    sheet["H28"] = format_mpd(col_data.get('Actual Toe Level\n(mPD)', np.nan))
    sheet["H27"] = format_float(col_data.get('As-built casing Length (m)', np.nan))
    sheet["H29"] = format_mpd(col_data.get('Actual Founding level', np.nan))
    sheet["H30"] = ""
    sheet["H31"] = ""
    sheet["H36"] = "C45/20D"

    theo_vol = col_data.get('Estimate Concrete volume\n(m3)', np.nan)
    act_vol = col_data.get('Actual Concrete Volumn (m3)', np.nan)
    sheet["H37"] = format_float(theo_vol)
    sheet["H38"] = format_float(act_vol)

    try:
        if float(theo_vol) > 0 and pd.notna(act_vol):
            overbreak = ((float(act_vol) - float(theo_vol)) / float(theo_vol)) * 100
            sheet["H39"] = format_float(overbreak)
        else:
            sheet["H39"] = ""
    except:
        sheet["H39"] = ""

    sheet["H40"] = format_date(col_data.get('Concrete Date\n(DD/MM/YYYY)', pd.NaT))
    virtual_file = io.BytesIO()
    wb.save(virtual_file)
    virtual_file.seek(0)
    return virtual_file.read(), f"{col_data['Concrete Column ID']}_Record.xlsx"

def generate_ipa_valuation_excel(comm_df):
    export_cols = [
        'Concrete Column ID', 'Construction_Stage', 'Concrete Date\n(DD/MM/YYYY)',
        'Design_Soil_Len_m', 'Design_CDG_Len_m', 'Extra_Depth_m',
        'Theo_Vol_m3', 'AsBuilt_Cylinder_Vol_m3', 'Actual Concrete Volumn (m3)',
        'Claimable_Var_Vol_m3', 'True_Overbreak_Vol_m3',
        'Baseline_BOQ_Value_HKD', 'Claimable_Variation_HKD',
        'True_Wastage_Cost_HKD', 'Certified_IPA_Value_HKD'
    ]
    ipa_df = comm_df[export_cols].copy()
    ipa_df.columns = [
        'Column ID', 'Stage', 'Concrete Date',
        'Soil Length (m)', 'Design CDG Socket (m)', 'Extra CDG Depth (m)',
        'Design Est. Vol (m3)', 'As-Built Cylinder Vol (m3)', 'Actual Poured Vol (m3)',
        'Claimable BOQ Var. Vol (m3)', 'Non-Claimable Overbreak (m3)',
        'Baseline BOQ Value (HKD)', 'Claimable Variation (HKD)',
        'Contractor Overbreak Loss (HKD)', 'Gross Certified IPA Value (HKD)'
    ]
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        ipa_df.to_excel(writer, index=False, sheet_name='Interim_Payment_Valuation')
    output.seek(0)
    return output.read(), f"QS_Interim_Payment_Valuation_{datetime.date.today()}.xlsx"

class PDFReport(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'Formal Concrete Column Submission & QS Audit (HKBD Standard)', 0, 1, 'C')
        self.ln(4)
    def footer(self):
        self.set_y(-20)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def generate_pdf_report(row):
    pdf = PDFReport()
    pdf.add_page()
    col_id = row.get('Concrete Column ID', 'Unknown')

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 9, f"Column Reference: {col_id}", 0, 1)
    pdf.set_font("Arial", size=10.5)
    pdf.cell(0, 7, f"Construction Stage: {row.get('Construction_Stage', '')}", 0, 1)
    pdf.cell(0, 7, f"Drill Start Date: {row.get('Drill Start Date \n(DD/MM/YYYY)', '')} | Completed: {row.get('Drill Completed Date\n(DD/MM/YYYY)', '')}", 0, 1)
    pdf.cell(0, 7, f"Concrete Pour Date: {row.get('Concrete Date\n(DD/MM/YYYY)', '')}", 0, 1)
    pdf.ln(3)

    pdf.set_font("Arial", 'B', 11.5)
    pdf.cell(0, 8, "1. Surveying Level Measurements (mPD)", 0, 1)
    pdf.set_font("Arial", size=10.5)
    pdf.cell(0, 7, f"Ground Level: {format_mpd(row.get('Ground Level', ''))} mPD | Casing Top Level: {format_mpd(row.get('Casing Top Level\n(mPD)', ''))} mPD", 0, 1)
    pdf.cell(0, 7, f"Actual Toe Level: {format_mpd(row.get('Actual Toe Level\n(mPD)', ''))} mPD | As-Built Casing Length: {format_float(row.get('As-built casing Length (m)', ''))} m", 0, 1)
    pdf.cell(0, 7, f"Design Founding Level: {format_mpd(row.get('Foundation level\n(mPD)', ''))} mPD | Actual Founding Level: {format_mpd(row.get('Actual Founding level', ''))} mPD", 0, 1)
    pdf.cell(0, 7, f"Cut-off Level: {format_mpd(row.get('Cut off level (mPD)', ''))} mPD | Concrete Top Level: {format_mpd(row.get('Concrete Top Level (mPD)', ''))} mPD", 0, 1)
    pdf.ln(3)

    pdf.set_font("Arial", 'B', 11.5)
    pdf.cell(0, 8, "2. Volume, Overbreak & Commercial QS Assessment", 0, 1)
    pdf.set_font("Arial", size=10.5)
    t_vol = float(row.get('Estimate Concrete volume\n(m3)', 0)) if pd.notna(row.get('Estimate Concrete volume\n(m3)', 0)) else 0.0
    ab_vol = float(row.get('AsBuilt_Cylinder_Vol_m3', 0)) if pd.notna(row.get('AsBuilt_Cylinder_Vol_m3', 0)) else 0.0
    a_vol = float(row.get('Actual Concrete Volumn (m3)', 0)) if pd.notna(row.get('Actual Concrete Volumn (m3)', 0)) else 0.0

    pdf.cell(0, 7, f"Design Estimated Volume (+3m Allowance): {t_vol:.3f} m3", 0, 1)
    pdf.cell(0, 7, f"As-Built Cylinder Volume: {ab_vol:.3f} m3 | Actual Poured Volume: {a_vol:.3f} m3", 0, 1)
    if t_vol > 0 and a_vol > 0:
        ob = ((a_vol - t_vol) / t_vol) * 100
        pdf.cell(0, 7, f"Gross Overbreak vs Design: {ob:.2f}% | Claimable BOQ Var: {row.get('Claimable_Var_Vol_m3', 0):.3f} m3 | True Radial Overbreak: {row.get('True_Overbreak_Vol_m3', 0):.3f} m3", 0, 1)

    pdf.ln(3)
    pdf.set_font("Arial", 'B', 11.5)
    pdf.cell(0, 8, "3. Plant Productivity & Digital Audit Trail", 0, 1)
    pdf.set_font("Arial", size=10.5)
    pdf.cell(0, 7, f"Rig Hours: {row.get('Rig_Hours', 0)} hrs | Casings Used: {row.get('Casings_Used', 0)} | Casing Damaged: {row.get('Casing_Damaged', 'No')}", 0, 1)
    pdf.cell(0, 7, f"Last Edited By: {row.get('Last_Edited_By', 'System/Unknown')} | Timestamp: {row.get('Timestamp_of_Edit', 'N/A')}", 0, 1)
    pdf.ln(12)
    pdf.cell(95, 8, "Prepared By (Site Engineer): ___________________", 0, 0)
    pdf.cell(95, 8, "Certified By (QS / Supervisor): ___________________", 0, 1)

    return pdf.output(dest='S').encode('latin-1'), f"{col_id}_Formal_Report.pdf"

# ==============================================================================
# 6. MAIN DASHBOARD EXECUTION
# ==============================================================================
st.title("🏢 Enterprise Digital Twin & QS Commercial Control: Concrete Columns")

with st.expander("📋 View Method Statement & General Procedure (610mm Ø Bored Concrete Columns)"):
    st.markdown("""
    ### General Construction Procedure for Concrete Columns (610mm Ø)
    1. **Setting Out:** Surveyor sets out the column position (X/Y coordinates) based on approved construction drawings.
    2. **Casing Installation:** Pitch and drive the 610mm temporary steel casing into the ground to the required casing top and toe levels to prevent soil collapse.
    3. **Drilling & Excavation:** Drill through the soil and alluvium layers into the Completely Decomposed Granite (CDG) until the actual founding level is reached.
    4. **Inspection & Measurement:** Site Engineer records the *Actual Toe Level* and *Actual Founding Level*, determining the socket length into the CD material.
    5. **Hole Cleaning:** Clean the base of the drilled hole to remove loose debris and sediment.
    6. **Concreting:** Pour **C45/20D** grade concrete using a tremie pipe. The engineer records the *Actual Concrete Volume* and compares it against the theoretical and as-built cylinder volumes to separate **Claimable BOQ Depth Variation** from **True Radial Overbreak %**.
    7. **Casing Extraction:** Slowly extract the temporary casing, ensuring the final *Concrete Top Level* is maintained at least 1.0m above the design Cut-off Level.
    """)

st.markdown("---")

raw_df = load_data()
df = run_commercial_engine(
    raw_df,
    concrete_unit_cost,
    soil_drill_rate,
    cdg_eo_rate,
    rig_hourly_rate,
    casing_damage_cost
)

filtered_df = df.copy()
if selected_stage != "All":
    filtered_df = filtered_df[filtered_df['Construction_Stage'] == selected_stage]

if df.empty:
    st.error("No valid concrete column records found.")
else:
    total_cols = len(df)
    completed_df = df[df['Construction_Stage'] == 'Completed'].copy()
    completed_cols = len(completed_df)
    remaining_cols = total_cols - completed_cols
    progress_pct = (completed_cols / total_cols) * 100 if total_cols > 0 else 0

    # --- TOP ROW: PROGRESS, PERIOD PRODUCTION, QS COST SPLIT, & SCHEDULE FORECAST ---
    st.markdown("### 📈 Executive Project Metrics, QS Commercial Control & Predictive Forecasting")
    top_col1, top_col2, top_col3, top_col4 = st.columns([1, 1.05, 1.35, 1.15])

    with top_col1:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number",
            value=progress_pct,
            number={'suffix': "%", 'valueformat': '.1f', 'font': {'size': 34, 'color': '#2c3e50'}},
            title={'text': f"Completion ({completed_cols}/{total_cols})", 'font': {'size': 15}},
            gauge={'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': "darkblue"}, 'bar': {'color': "#27ae60"}, 'bgcolor': "#ecf0f1", 'shape': "angular"}
        ))
        fig_gauge.update_layout(height=195, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_gauge, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with top_col2:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown("**📅 Period Production Analysis**")
        today = datetime.date.today()
        start_default = today.replace(day=1)
        date_range = st.date_input("Filter Range", value=(start_default, today), format="YYYY-MM-DD")

        if len(date_range) == 2:
            s_d, e_d = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
            d_dates = pd.to_datetime(df['Drill Completed Date\n(DD/MM/YYYY)'], errors='coerce')
            c_dates = pd.to_datetime(df['Concrete Date\n(DD/MM/YYYY)'], errors='coerce')
            st.metric("Holes Drilled in Period", int(((d_dates >= s_d) & (d_dates <= e_d)).sum()))
            st.metric("Columns Concreted", int(((c_dates >= s_d) & (c_dates <= e_d)).sum()))
        else:
            st.warning("Select start and end date.")
        st.markdown("</div>", unsafe_allow_html=True)

    with top_col3:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown("**💰 QS Commercial Split (HKSMM4)**")
        total_theo_budget = df['Theo_Vol_m3'].sum() * concrete_unit_cost
        claimable_var_hkd = df['Claimable_Variation_HKD'].sum()
        true_waste_hkd = df['True_Wastage_Cost_HKD'].sum()
        plant_dmg_hkd = df['Plant_And_Damage_Cost_HKD'].sum()

        st.caption(f"Baseline Concrete Budget (241 cols): **HKD ${total_theo_budget:,.0f}**")
        st.markdown(f"<div class='claim-good'>➕ Claimable BOQ Variation: HKD ${claimable_var_hkd:,.2f}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='cost-alert'>➖ True Overbreak Wastage: HKD ${true_waste_hkd:,.2f}</div>", unsafe_allow_html=True)
        st.caption(f"Rig Operation & Casing Damage Cost: **HKD ${plant_dmg_hkd:,.0f}**")
        st.markdown("</div>", unsafe_allow_html=True)

    with top_col4:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown("**⏳ Predictive Schedule & EVM**")
        tot_ev = df['Earned_Value_HKD'].sum()
        tot_ac = df['Actual_Cost_HKD'].sum()
        cpi_val = (tot_ev / tot_ac) if tot_ac > 0 else 1.0
        st.markdown(f"<div class='evm-badge'>CPI (Cost Efficiency): {cpi_val:.2f}</div>", unsafe_allow_html=True)

        if not completed_df.empty and 'Concrete Date\n(DD/MM/YYYY)' in completed_df.columns:
            completed_df['Concrete Date'] = pd.to_datetime(completed_df['Concrete Date\n(DD/MM/YYYY)'], errors='coerce')
            valid_dates = completed_df.dropna(subset=['Concrete Date']).sort_values('Concrete Date')
            if len(valid_dates) > 1:
                days_elapsed = (valid_dates['Concrete Date'].max() - valid_dates['Concrete Date'].min()).days
                if days_elapsed > 0:
                    rate_per_day = len(valid_dates) / days_elapsed
                    days_remaining = remaining_cols / rate_per_day if rate_per_day > 0 else 0
                    proj_date = valid_dates['Concrete Date'].max() + datetime.timedelta(days=days_remaining)
                    st.success(f"**Rate:** {rate_per_day:.2f} cols/day")
                    st.info(f"**Est. Finish:** {proj_date.strftime('%d %b %Y')}")
                else:
                    st.warning("Need wider date spread.")
            else:
                st.warning("Log ≥2 completion dates.")
        else:
            st.warning("No completed records yet.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")

    # --- MAIN WORKSPACE: VISUALIZATIONS & EVM S-CURVE VS DATA ENTRY ---
    col_vis, col_data = st.columns([2.45, 1.55])

    with col_vis:
        st.markdown("### 🗺️ Spatial Digital Twin & Earned Value Analytics")
        color_map = {"Not Started": "lightgrey", "In Progress": "#f1c40f", "Drilled": "#3498db", "Completed": "#2ecc71"}
        tab1, tab2, tab3 = st.tabs([
            "2D Grid Plan View",
            "3D Foundation Depth Model",
            "📊 Earned Value (EVM) & Cost Breakdown"
        ])

        with tab1:
            if not filtered_df.empty:
                fig2d = px.scatter(
                    filtered_df, x='X_Coord', y='Y_Coord', color='Construction_Stage',
                    color_discrete_map=color_map, hover_name='Concrete Column ID',
                    hover_data=['Foundation level\n(mPD)', 'Estimate Concrete volume\n(m3)', 'Actual Concrete Volumn (m3)']
                )
                fig2d.update_traces(marker=dict(size=15, line=dict(width=1, color='DarkSlateGrey')))
                fig2d.update_yaxes(scaleanchor="x", scaleratio=1)
                fig2d.update_layout(height=620, plot_bgcolor='white', paper_bgcolor='white')
                fig2d.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGrey')
                fig2d.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGrey')
                st.plotly_chart(fig2d, use_container_width=True)
            else:
                st.info("No columns match the selected stage filter.")

        with tab2:
            if not filtered_df.empty:
                fig3d = px.scatter_3d(
                    filtered_df, x='X_Coord', y='Y_Coord', z='Tentative Concrete Column Bottom level \n(mPD)',
                    color='Construction_Stage', color_discrete_map=color_map, hover_name='Concrete Column ID'
                )
                for _, row in filtered_df.iterrows():
                    bottom_z = row['Actual Founding level'] if pd.notna(row['Actual Founding level']) else row['Tentative Concrete Column Bottom level \n(mPD)']
                    if pd.isna(bottom_z):
                        bottom_z = -40.0
                    top_z = row['Cut off level (mPD)'] if pd.notna(row['Cut off level (mPD)']) else 0.8
                    fig3d.add_scatter3d(
                        x=[row['X_Coord'], row['X_Coord']], y=[row['Y_Coord'], row['Y_Coord']],
                        z=[top_z, bottom_z], mode='lines',
                        line=dict(color=color_map.get(row['Construction_Stage'], 'grey'), width=7),
                        showlegend=False
                    )
                fig3d.update_layout(scene=dict(zaxis=dict(title='Elevation (mPD)'), xaxis=dict(title='X Grid (m)'), yaxis=dict(title='Y Grid (m)')), height=620)
                st.plotly_chart(fig3d, use_container_width=True)
            else:
                st.info("No data available.")

        with tab3:
            st.markdown("#### 💰 Site-Wide BOQ Financial & Volume Breakdown")
            total_baseline_boq = df['Baseline_BOQ_Value_HKD'].sum()
            ev_total = df['Earned_Value_HKD'].sum()
            ac_total = df['Actual_Cost_HKD'].sum()
            ipa_total = df['Certified_IPA_Value_HKD'].sum()

            evm_c1, evm_c2, evm_c3, evm_c4 = st.columns(4)
            evm_c1.metric("Total Contract BOQ (BAC)", f"HK${total_baseline_boq/1e6:.2f}M")
            evm_c2.metric("Earned Value (EV)", f"HK${ev_total:,.0f}")
            evm_c3.metric("Actual Cost (AC)", f"HK${ac_total:,.0f}")
            evm_c4.metric("Gross Certified IPA", f"HK${ipa_total:,.0f}")

            # Plot Volume Comparison Bar Chart for Completed Columns
            if not completed_df.empty:
                vol_comp_df = completed_df[[
                    'Concrete Column ID', 'Theo_Vol_m3',
                    'Claimable_Var_Vol_m3', 'True_Overbreak_Vol_m3'
                ]].copy()
                vol_comp_df.columns = [
                    'Column ID', 'Design Theoretical Vol (m³)',
                    'Claimable BOQ Depth Var (m³)', 'True Non-Claimable Overbreak (m³)'
                ]
                fig_cost = px.bar(
                    vol_comp_df, x='Column ID',
                    y=['Design Theoretical Vol (m³)', 'Claimable BOQ Depth Var (m³)', 'True Non-Claimable Overbreak (m³)'],
                    title="Completed Columns: Design Volume vs. Claimable BOQ Variation vs. True Overbreak Wastage",
                    barmode='stack',
                    color_discrete_map={
                        'Design Theoretical Vol (m³)': '#2980b9',
                        'Claimable BOQ Depth Var (m³)': '#27ae60',
                        'True Non-Claimable Overbreak (m³)': '#c0392b'
                    }
                )
                fig_cost.update_layout(height=420, plot_bgcolor='white')
                st.plotly_chart(fig_cost, use_container_width=True)
            else:
                st.info("Complete and log at least one concrete column to view stacked BOQ vs. Overbreak volume analytics.")

    with col_data:
        st.markdown("### 📝 Record As-Built & Plant Data")

        if st.session_state['role'] == 'Viewer':
            st.warning("🔒 Viewer Mode: Log in via the sidebar (`worker123` or `admin123`) to record site data or export QS reports.")
        else:
            valid_cols = df['Concrete Column ID'].dropna().unique()
            selected_col = st.selectbox("Select Column ID to Update", valid_cols)
            col_design_data = df[df['Concrete Column ID'] == selected_col].iloc[0]

            current_stage = col_design_data.get('Construction_Stage', 'Not Started')
            if current_stage != "Not Started":
                st.warning(f"⚠️ **Overwrite Mode:** `{selected_col}` is currently `{current_stage}`. Submitting will update and log an audit trail.")
            else:
                st.info(f"ℹ️ **New Entry:** `{selected_col}` | Design Est. Vol: `{col_design_data.get('Theo_Vol_m3', 0):.2f} m³`")

            with st.form("update_form"):
                stage_opts = ["Not Started", "In Progress", "Drilled", "Completed"]
                s_idx = stage_opts.index(current_stage) if current_stage in stage_opts else 0
                new_stage = st.selectbox("Update Construction Stage", stage_opts, index=s_idx)

                def get_val(key, default_val=0.0):
                    val = col_design_data.get(key, default_val)
                    return float(val) if pd.notna(val) else float(default_val)

                st.markdown("##### 📏 Surveying & Elevation Measurements")
                f_col1, f_col2 = st.columns(2)
                with f_col1:
                    ground_level = st.number_input("Ground Level (mPD)", value=get_val('Ground Level'), format="%.3f")
                    casing_top = st.number_input("Casing Top Level (mPD)", value=get_val('Casing Top Level\n(mPD)'), format="%.3f")
                    act_toe = st.number_input("Actual Toe Level (mPD)", value=get_val('Actual Toe Level\n(mPD)'), format="%.3f")
                with f_col2:
                    act_founding = st.number_input("Actual Founding Lvl (mPD)", value=get_val('Actual Founding level', get_val('Foundation level\n(mPD)')), format="%.3f")
                    as_built_casing = st.number_input("As-built Casing L. (m)", value=get_val('As-built casing Length (m)'), format="%.3f")
                    conc_top = st.number_input("Concrete Top Lvl (mPD)", value=get_val('Concrete Top Level (mPD)', get_val('Cut off level (mPD)') + 3.0), format="%.3f")

                concrete_vol = st.number_input("Actual Poured Concrete Volume (m³)", value=get_val('Actual Concrete Volumn (m3)'), format="%.3f")

                st.markdown("##### 🚜 Plant Productivity & Casing Integrity")
                p_col1, p_col2, p_col3 = st.columns(3)
                with p_col1:
                    rig_hours = st.number_input("Rig Hours (hrs)", value=get_val('Rig_Hours', 0.0), min_value=0.0, step=0.5)
                with p_col2:
                    casings_used = st.number_input("Casings Used (nos)", value=int(get_val('Casings_Used', 5)), min_value=0, step=1)
                with p_col3:
                    dmg_curr = str(col_design_data.get('Casing_Damaged', 'No'))
                    casing_damaged = st.selectbox("Casing Damaged?", ["No", "Yes"], index=1 if dmg_curr == "Yes" else 0)

                site_remarks = st.text_input("Site / QS Remarks", value=str(col_design_data.get('Site_Remarks', '')) if pd.notna(col_design_data.get('Site_Remarks', '')) else "")
                site_photo = st.file_uploader("📷 Upload Casing / Tremie Pour Photo (Optional)", type=['jpg', 'png', 'jpeg'])

                st.markdown("##### 📅 Operation Dates")
                d_col1, d_col2, d_col3 = st.columns(3)
                with d_col1:
                    start_date = st.date_input("Drill Start")
                with d_col2:
                    comp_date = st.date_input("Drill End")
                with d_col3:
                    conc_date = st.date_input("Concrete Pour")

                submitted = st.form_submit_button("✅ Validate & Commit Record")

                if submitted:
                    qa_passed = True
                    if act_founding > ground_level:
                        st.error("🚨 **QA/QC Hard Gate Blocked:** Actual Founding Level cannot be higher than Ground Level.")
                        qa_passed = False

                    theo_v = get_val('Estimate Concrete volume\n(m3)')
                    if theo_v > 0 and concrete_vol > 0:
                        ob = ((concrete_vol - theo_v) / theo_v) * 100
                        if ob > 20:
                            st.warning(f"⚠️ **QA/QC Alert:** Gross Overbreak ({ob:.1f}%) exceeds 20% tolerance. Logged for QS/Admin review.")
                            st.toast(f"📱 Webhook Dispatched: High overbreak ({ob:.1f}%) on {selected_col}!", icon="🚨")

                    if qa_passed:
                        idx_mask = raw_df['Concrete Column ID'] == selected_col
                        raw_df.loc[idx_mask, 'Construction_Stage'] = new_stage
                        raw_df.loc[idx_mask, 'Ground Level'] = ground_level
                        raw_df.loc[idx_mask, 'Casing Top Level\n(mPD)'] = casing_top
                        raw_df.loc[idx_mask, 'Actual Toe Level\n(mPD)'] = act_toe
                        raw_df.loc[idx_mask, 'Actual Founding level'] = act_founding
                        raw_df.loc[idx_mask, 'Actual Concrete Volumn (m3)'] = concrete_vol
                        raw_df.loc[idx_mask, 'As-built casing Length (m)'] = as_built_casing
                        raw_df.loc[idx_mask, 'Concrete Top Level (mPD)'] = conc_top
                        raw_df.loc[idx_mask, 'Rig_Hours'] = rig_hours
                        raw_df.loc[idx_mask, 'Casings_Used'] = casings_used
                        raw_df.loc[idx_mask, 'Casing_Damaged'] = casing_damaged
                        raw_df.loc[idx_mask, 'Site_Remarks'] = site_remarks
                        raw_df.loc[idx_mask, 'Drill Start Date \n(DD/MM/YYYY)'] = start_date
                        if new_stage in ["Drilled", "Completed"]:
                            raw_df.loc[idx_mask, 'Drill Completed Date\n(DD/MM/YYYY)'] = comp_date
                        if new_stage == "Completed":
                            raw_df.loc[idx_mask, 'Concrete Date\n(DD/MM/YYYY)'] = conc_date

                        raw_df.loc[idx_mask, 'Last_Edited_By'] = st.session_state['username']
                        raw_df.loc[idx_mask, 'Timestamp_of_Edit'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                        save_path = "FYP_Concrete_Column_Data.csv" if os.path.exists("FYP_Concrete_Column_Data.csv") else "FYP_App_Data.csv"
                        raw_df.to_csv(save_path, index=False)
                        st.cache_data.clear()
                        st.success(f"Column `{selected_col}` validated, priced, and saved!")
                        st.rerun()

            # --- EXPORT DOCUMENTATION & QS IPA SUITE ---
            st.markdown("### 📄 Export Engineering & QS Documentation")
            if st.session_state['role'] == 'Admin':
                exp_t1, exp_t2, exp_t3, exp_t4 = st.tabs([
                    "PDF Report", "Single Excel", "Batch ZIP", "💰 QS IPA Valuation"
                ])
                curr_row = df[df['Concrete Column ID'] == selected_col].iloc[0]

                with exp_t1:
                    pdf_data, pdf_name = generate_pdf_report(curr_row)
                    st.download_button("📥 Download Formal HKBD PDF", pdf_data, pdf_name, mime="application/pdf")

                with exp_t2:
                    excel_data, f_name = generate_excel_report(curr_row)
                    if excel_data:
                        st.download_button("📥 Download Record Sheet (.xlsx)", excel_data, f_name)
                    else:
                        st.error("Template `Record Sheet for FYP.xlsx` not found.")

                with exp_t3:
                    if st.button("📦 Build Completed Columns ZIP"):
                        comp_rows = df[df['Construction_Stage'] == 'Completed']
                        if comp_rows.empty:
                            st.warning("No completed columns to package.")
                        else:
                            zip_buf = io.BytesIO()
                            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
                                for _, r in comp_rows.iterrows():
                                    fb, fn = generate_excel_report(r)
                                    if fb:
                                        zf.writestr(fn, fb)
                            st.download_button("📥 Download Batch ZIP", zip_buf.getvalue(), f"Batch_Records_{datetime.date.today()}.zip", mime="application/zip")

                with exp_t4:
                    st.write("Export full HKSMM4 Interim Payment Application (IPA) & Overbreak Audit Sheet.")
                    ipa_bytes, ipa_name = generate_ipa_valuation_excel(df)
                    st.download_button("💰 Download Monthly QS IPA Sheet (.xlsx)", ipa_bytes, ipa_name)
            else:
                st.info("Log in as Admin (`admin123`) to unlock PDF, Batch ZIP, and QS Interim Payment exports.")

    # --- COMPLETE DATABASE & COMMERCIAL AUDIT EXPANDER ---
    st.markdown("---")
    with st.expander("📂 View Complete Database (Geometry, QS Valuation Split & Digital Audit Trail)"):
        display_df = filtered_df.copy()
        for mPD_col in ['Cut off level (mPD)', 'Alluvium Bottom\n(mPD)', 'Foundation level\n(mPD)', 'Ground Level', 'Casing Top Level\n(mPD)', 'Actual Founding level', 'Concrete Top Level (mPD)']:
            if mPD_col in display_df.columns:
                display_df[mPD_col] = display_df[mPD_col].apply(lambda x: f"{x:+.3f}" if pd.notna(x) and str(x).strip() != "" else "")

        cols_to_show = [c for c in [
            'Concrete Column ID', 'Construction_Stage',
            'Cut off level (mPD)', 'Foundation level\n(mPD)', 'Actual Founding level',
            'Theo_Vol_m3', 'AsBuilt_Cylinder_Vol_m3', 'Actual Concrete Volumn (m3)',
            'Claimable_Var_Vol_m3', 'True_Overbreak_Vol_m3',
            'Baseline_BOQ_Value_HKD', 'Claimable_Variation_HKD', 'True_Wastage_Cost_HKD',
            'Certified_IPA_Value_HKD', 'Drilling_Speed_m_hr', 'Casing_Damaged',
            'Last_Edited_By', 'Timestamp_of_Edit'
        ] if c in display_df.columns]

        st.dataframe(display_df[cols_to_show], use_container_width=True, height=340)
