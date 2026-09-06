
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

# 1. Professional Page Configuration
st.set_page_config(layout="wide", page_title="Enterprise Concrete Column", page_icon="🏢", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    h1, h2, h3 { color: #2c3e50; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .metric-container { background-color: white; padding: 15px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

# --- 5. ROLE-BASED AUTHENTICATION ---
if 'role' not in st.session_state:
    st.session_state['role'] = 'Viewer'

with st.sidebar:
    st.header("🔐 System Login")
    pwd = st.text_input("Enter Password", type="password")
    if pwd == "admin123":
        st.session_state['role'] = 'Admin'
        st.success("Logged in as: Admin (Full Access)")
    elif pwd == "worker123":
        st.session_state['role'] = 'Editor'
        st.success("Logged in as: Site Engineer (Data Entry Only)")
    elif pwd != "":
        st.error("Invalid password.")
        st.session_state['role'] = 'Viewer'
    else:
        st.info("Please log in to edit data.")

st.title("🏢 Enterprise Digital Twin: Concrete Columns")
st.markdown("---")

# Robust Load data
@st.cache_data(show_spinner=False)
def load_data():
    if os.path.exists("FYP_Concrete_Column_Data.csv"):
        df = pd.read_csv("FYP_Concrete_Column_Data.csv")
    elif os.path.exists("FYP_App_Data.csv"):
        df = pd.read_csv("FYP_App_Data.csv")
    else:
        df = pd.DataFrame({"Concrete Column ID": ["BH1_1_C1", "BH1_1_C2", "BH1_1_C3", "BH1_2_C1"], "Construction_Stage": ["Not Started"]*4})
        
    required_cols = [
        'Cut off level (mPD)', 'Tentative Concrete Column Bottom level \n(mPD)',
        'Actual Founding level', 'Foundation level\n(mPD)', 
        'Actual Concrete Volumn (m3)', 'X_Coord', 'Y_Coord',
        'Drill Start Date \n(DD/MM/YYYY)', 'Drill Completed Date\n(DD/MM/YYYY)', 
        'Concrete Date\n(DD/MM/YYYY)', 'Casing Top Level\n(mPD)', 
        'Measuring Depth\n(m)', 'Ground Level', 'Actual Toe Level\n(mPD)',
        'Alluvium Bottom\n(mPD)', 'Embedded Depth in CD Material \n(m)',
        'Estimate Concrete volume\n(m3)', 
        'As-built casing Length (m)', 'Concrete Top Level (mPD)'
    ]
    
    for col in required_cols:
        if col not in df.columns:
            df[col] = np.nan
            
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
        
    return df

df = load_data()

def format_mpd(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return ""
        return f"{float(val):+.3f}"
    except: return val

def format_float(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return ""
        return f"{float(val):.3f}"
    except: return val

def generate_excel_report(col_data):
    template_path = "Record Sheet for FYP_3.xlsx"
    if not os.path.exists(template_path):
        template_path = "Record Sheet for FYP.xlsx"
        
    if not os.path.exists(template_path):
        return None, None
        
    wb = openpyxl.load_workbook(template_path)
    sheet = wb.active
    
    sheet["H8"] = col_data.get('Concrete Column ID', '')
    sheet["H11"] = "610"; sheet["H12"] = "610"; sheet["H13"] = "0"
    sheet["H14"] = format_mpd(col_data.get('Alluvium Bottom\n(mPD)', np.nan))
    sheet["H15"] = format_mpd(col_data.get('Tentative Concrete Column Bottom level \n(mPD)', np.nan))
    
    cutoff = col_data.get('Cut off level (mPD)', np.nan)
    sheet["M21"] = format_mpd(cutoff)
    sheet["M17"] = format_mpd(col_data.get('Concrete Top Level (mPD)', np.nan))

    def format_date(d): return d.strftime('%Y-%m-%d') if pd.notna(d) else ""
        
    sheet["H23"] = format_date(col_data.get('Drill Start Date \n(DD/MM/YYYY)', pd.NaT))
    sheet["H24"] = format_date(col_data.get('Drill Completed Date\n(DD/MM/YYYY)', pd.NaT))
    sheet["H25"] = format_mpd(col_data.get('Ground Level', np.nan))
    sheet["H26"] = format_mpd(col_data.get('Casing Top Level\n(mPD)', np.nan))
    sheet["H28"] = format_mpd(col_data.get('Actual Toe Level\n(mPD)', np.nan))
    sheet["H27"] = format_float(col_data.get('As-built casing Length (m)', np.nan))
    sheet["H29"] = format_mpd(col_data.get('Actual Founding level', np.nan))
    sheet["H30"] = ""; sheet["H31"] = ""
    sheet["H36"] = "C45/20D"
    
    theo_vol = col_data.get('Estimate Concrete volume\n(m3)', np.nan)
    act_vol = col_data.get('Actual Concrete Volumn (m3)', np.nan)
    sheet["H37"] = format_float(theo_vol)
    sheet["H38"] = format_float(act_vol)
    
    try:
        if float(theo_vol) > 0:
            overbreak = ((float(act_vol) - float(theo_vol)) / float(theo_vol)) * 100
            sheet["H39"] = format_float(overbreak)
        else: sheet["H39"] = ""
    except: sheet["H39"] = ""
        
    sheet["H40"] = format_date(col_data.get('Concrete Date\n(DD/MM/YYYY)', pd.NaT))
    
    virtual_file = io.BytesIO()
    wb.save(virtual_file)
    virtual_file.seek(0)
    return virtual_file.read(), f"{col_data['Concrete Column ID']}_Record.xlsx"

if df.empty:
    st.error("No data.")
else:
    # --- 3. PREDICTIVE PROJECT FORECASTING ---
    total_cols = len(df)
    completed_df = df[df['Construction_Stage'] == 'Completed'].copy()
    completed_cols = len(completed_df)
    remaining_cols = total_cols - completed_cols
    progress_pct = (completed_cols / total_cols) * 100 if total_cols > 0 else 0
    
    st.markdown("### 📈 Project Metrics & Predictive Forecasting")
    top_col1, top_col2, top_col3 = st.columns([1, 1, 1.5])
    
    with top_col1:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.metric("Total Clusters", f"{total_cols} nos")
        st.metric("Completed", f"{completed_cols} nos", f"{progress_pct:.1f}%")
        st.markdown("</div>", unsafe_allow_html=True)
        
    with top_col2:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        vol_col = 'Actual Concrete Volumn (m3)'
        total_vol = pd.to_numeric(df[vol_col], errors='coerce').sum() if vol_col in df.columns else 0.0
        st.metric("Total Injected Volume", f"{total_vol:.2f} m³")
        st.metric("Remaining Columns", f"{remaining_cols} nos")
        st.markdown("</div>", unsafe_allow_html=True)
        
    with top_col3:
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown("**Predictive Completion Timeline**")
        
        # Calculate run rate
        if not completed_df.empty and 'Concrete Date\n(DD/MM/YYYY)' in completed_df.columns:
            completed_df['Concrete Date'] = pd.to_datetime(completed_df['Concrete Date\n(DD/MM/YYYY)'], errors='coerce')
            valid_dates = completed_df.dropna(subset=['Concrete Date']).sort_values('Concrete Date')
            if len(valid_dates) > 1:
                first_date = valid_dates['Concrete Date'].min()
                last_date = valid_dates['Concrete Date'].max()
                days_elapsed = (last_date - first_date).days
                if days_elapsed > 0:
                    rate_per_day = len(valid_dates) / days_elapsed
                    days_remaining = remaining_cols / rate_per_day if rate_per_day > 0 else 0
                    proj_date = last_date + datetime.timedelta(days=days_remaining)
                    st.success(f"**Current Rate:** {rate_per_day:.2f} columns/day")
                    st.info(f"**Projected Completion:** {proj_date.strftime('%d %b %Y')}")
                else:
                    st.warning("Not enough date spread to project timeline.")
            else:
                st.warning("Need more completed records to calculate run rate.")
        else:
            st.warning("No completed records available for forecasting.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")

    col_vis, col_data = st.columns([2.5, 1.5])
    
    with col_vis:
        st.markdown("### 🗺️ Site Plan Visualizations")
        color_map = {"Not Started": "lightgrey", "In Progress": "#f1c40f", "Drilled": "#3498db", "Completed": "#2ecc71"}
        tab1, tab2 = st.tabs(["2D Grid Plan View", "3D Column Foundation Model"])
        with tab1:
            fig2d = px.scatter(df, x='X_Coord', y='Y_Coord', color='Construction_Stage', color_discrete_map=color_map, hover_name='Concrete Column ID')
            fig2d.update_traces(marker=dict(size=16, line=dict(width=1, color='DarkSlateGrey')))
            fig2d.update_yaxes(scaleanchor="x", scaleratio=1)
            fig2d.update_layout(height=650, plot_bgcolor='white', paper_bgcolor='white')
            fig2d.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGrey')
            fig2d.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGrey')
            st.plotly_chart(fig2d, use_container_width=True)
            
        with tab2:
            fig3d = px.scatter_3d(df, x='X_Coord', y='Y_Coord', z='Tentative Concrete Column Bottom level \n(mPD)', color='Construction_Stage', color_discrete_map=color_map, hover_name='Concrete Column ID')
            for i, row in df.iterrows():
                bottom_z = row['Actual Founding level']
                if pd.isna(bottom_z): bottom_z = row['Tentative Concrete Column Bottom level \n(mPD)']
                if pd.isna(bottom_z): bottom_z = 0 
                top_z = row['Cut off level (mPD)']
                if pd.isna(top_z): top_z = 10 
                fig3d.add_scatter3d(x=[row['X_Coord'], row['X_Coord']], y=[row['Y_Coord'], row['Y_Coord']], z=[top_z, bottom_z], mode='lines', line=dict(color=color_map.get(row['Construction_Stage'], 'grey'), width=8), showlegend=False)
            fig3d.update_layout(scene=dict(zaxis=dict(title='Depth (mPD)'), xaxis=dict(title='X Grid'), yaxis=dict(title='Y Grid')), height=650)
            st.plotly_chart(fig3d, use_container_width=True)

    with col_data:
        st.markdown("### 📝 Record As-Built Data")
        
        if st.session_state['role'] == 'Viewer':
            st.warning("🔒 You are in Viewer mode. Please log in from the sidebar to edit data or generate reports.")
        else:
            valid_cols = df['Concrete Column ID'].dropna().unique()
            selected_col = st.selectbox("Select Column ID to Update", valid_cols)
            col_design_data = df[df['Concrete Column ID'] == selected_col].iloc[0]
            
            with st.form("update_form"):
                new_stage = st.selectbox("Update Status", ["Not Started", "In Progress", "Drilled", "Completed"], index=["Not Started", "In Progress", "Drilled", "Completed"].index(col_design_data.get('Construction_Stage', 'Not Started')))
                
                f_col1, f_col2 = st.columns(2)
                def get_val(key): return float(col_design_data.get(key, 0.0)) if pd.notna(col_design_data.get(key, 0.0)) else 0.0
                
                with f_col1:
                    ground_level = st.number_input("Ground Level (mPD)", value=get_val('Ground Level'), format="%.3f")
                    casing_top = st.number_input("Casing Top Level (mPD)", value=get_val('Casing Top Level\n(mPD)'), format="%.3f")
                    act_toe = st.number_input("Actual Toe Level (mPD)", value=get_val('Actual Toe Level\n(mPD)'), format="%.3f")
                with f_col2:
                    act_founding = st.number_input("Actual Founding Lvl (mPD)", value=get_val('Actual Founding level'), format="%.3f")
                    as_built_casing = st.number_input("As-built Casing L. (m)", value=get_val('As-built casing Length (m)'), format="%.3f")
                    conc_top = st.number_input("Concrete Top Lvl (mPD)", value=get_val('Concrete Top Level (mPD)'), format="%.3f")
                
                concrete_vol = st.number_input("Actual Concrete Volume (m3)", value=get_val('Actual Concrete Volumn (m3)'), format="%.3f")
                
                # --- 6. PHOTOGRAPHIC EVIDENCE ---
                st.markdown("##### 📷 Site Photographs (Optional)")
                site_photo = st.file_uploader("Upload Casing / Pouring Photo", type=['jpg', 'png', 'jpeg'])
                
                d_col1, d_col2, d_col3 = st.columns(3)
                with d_col1: start_date = st.date_input("Drill Start")
                with d_col2: comp_date = st.date_input("Drill End")
                with d_col3: conc_date = st.date_input("Concrete Pour")
                
                submitted = st.form_submit_button("✅ Update System Record")
                
                if submitted:
                    # --- 2. AUTOMATED QA/QC VALIDATION ---
                    qa_passed = True
                    if act_founding > ground_level:
                        st.error("🚨 **QA/QC Error:** Actual Founding Level cannot be strictly higher than Ground Level. Submission blocked.")
                        qa_passed = False
                    
                    theo_v = get_val('Estimate Concrete volume\n(m3)')
                    if theo_v > 0 and concrete_vol > 0:
                        ob = ((concrete_vol - theo_v) / theo_v) * 100
                        if ob > 20:
                            st.warning(f"⚠️ **QA/QC Warning:** Overbreak is unusually high ({ob:.1f}%). Record saved, but requires Admin review.")
                    
                    if qa_passed:
                        df.loc[df['Concrete Column ID'] == selected_col, 'Construction_Stage'] = new_stage
                        df.loc[df['Concrete Column ID'] == selected_col, 'Ground Level'] = ground_level
                        df.loc[df['Concrete Column ID'] == selected_col, 'Casing Top Level\n(mPD)'] = casing_top
                        df.loc[df['Concrete Column ID'] == selected_col, 'Actual Toe Level\n(mPD)'] = act_toe
                        df.loc[df['Concrete Column ID'] == selected_col, 'Actual Founding level'] = act_founding
                        df.loc[df['Concrete Column ID'] == selected_col, 'Actual Concrete Volumn (m3)'] = concrete_vol
                        df.loc[df['Concrete Column ID'] == selected_col, 'As-built casing Length (m)'] = as_built_casing
                        df.loc[df['Concrete Column ID'] == selected_col, 'Concrete Top Level (mPD)'] = conc_top
                        df.loc[df['Concrete Column ID'] == selected_col, 'Drill Start Date \n(DD/MM/YYYY)'] = start_date
                        if new_stage in ["Drilled", "Completed"]:
                            df.loc[df['Concrete Column ID'] == selected_col, 'Drill Completed Date\n(DD/MM/YYYY)'] = comp_date
                        if new_stage == "Completed":
                            df.loc[df['Concrete Column ID'] == selected_col, 'Concrete Date\n(DD/MM/YYYY)'] = conc_date
                        
                        save_path = "FYP_Concrete_Column_Data.csv" if os.path.exists("FYP_Concrete_Column_Data.csv") else "FYP_App_Data.csv"
                        df.to_csv(save_path, index=False)
                        st.success(f"Column {selected_col} updated successfully.")
                        st.rerun()

            # --- 4. BATCH EXPORT ---
            st.markdown("### 📄 Export Documentation")
            if st.session_state['role'] == 'Admin':
                st.write("Generate individual or batch Excel reports.")
                
                export_tab1, export_tab2 = st.tabs(["Single Record", "Batch Export (ZIP)"])
                
                with export_tab1:
                    excel_data, f_name = generate_excel_report(df[df['Concrete Column ID'] == selected_col].iloc[0])
                    if excel_data:
                        st.download_button("📥 Download Single Record", excel_data, f_name)
                    else:
                        st.error("Template missing.")
                        
                with export_tab2:
                    st.write("Download all 'Completed' columns as a ZIP archive.")
                    if st.button("Generate Batch Archive"):
                        completed_rows = df[df['Construction_Stage'] == 'Completed']
                        if completed_rows.empty:
                            st.warning("No completed columns to export.")
                        else:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                for idx, row in completed_rows.iterrows():
                                    file_bytes, file_name = generate_excel_report(row)
                                    if file_bytes:
                                        zip_file.writestr(file_name, file_bytes)
                            
                            st.download_button(
                                label="📦 Download ZIP Archive",
                                data=zip_buffer.getvalue(),
                                file_name=f"Batch_Export_{datetime.date.today()}.zip",
                                mime="application/zip"
                            )
            else:
                st.info("Export features are restricted to Admin users.")

