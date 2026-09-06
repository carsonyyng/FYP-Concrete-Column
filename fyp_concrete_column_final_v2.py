
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import datetime
import numpy as np
import openpyxl

st.set_page_config(layout="wide", page_title="FYP Concrete Column")

st.title("🏗️ FYP Concrete Column: Progress Dashboard")

# Robust Load data
@st.cache_data
def load_data():
    if os.path.exists("FYP_Concrete_Column_Data.csv"):
        df = pd.read_csv("FYP_Concrete_Column_Data.csv")
    elif os.path.exists("FYP_App_Data.csv"):
        df = pd.read_csv("FYP_App_Data.csv")
    else:
        df = pd.DataFrame({"Concrete Column ID": ["BH1_1_C1"], "Construction_Stage": ["Not Started"]})
        
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
            
    if df['X_Coord'].isna().all():
        df['X_Coord'] = np.random.uniform(0, 100, len(df))
        df['Y_Coord'] = np.random.uniform(0, 100, len(df))
            
    date_cols = ['Concrete Date\n(DD/MM/YYYY)', 'Drill Start Date \n(DD/MM/YYYY)', 'Drill Completed Date\n(DD/MM/YYYY)']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
            
    if 'Construction_Stage' not in df.columns:
        df['Construction_Stage'] = "Not Started"
        
    return df

df = load_data()

# Helper formatting functions
def format_mpd(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return ""
        return f"{float(val):+.3f}"
    except:
        return val

def format_float(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return ""
        return f"{float(val):.3f}"
    except:
        return val

# Function to generate Excel report
def generate_excel_report(col_data):
    template_path = "Record Sheet for FYP_3.xlsx"
    if not os.path.exists(template_path):
        template_path = "Record Sheet for FYP.xlsx"
        
    os.makedirs("completed_records", exist_ok=True)
    file_name = f"completed_records/{col_data['Concrete Column ID']}_Record.xlsx"
    
    if not os.path.exists(template_path):
        return None, f"Template '{template_path}' not found. Please upload it."
        
    wb = openpyxl.load_workbook(template_path)
    sheet = wb.active
    
    # 1. Hole Data
    sheet["H8"] = col_data.get('Concrete Column ID', '')
    sheet["H11"] = "610" # Casing Size
    sheet["H12"] = "610" # Hole Size
    sheet["H13"] = "0"   # Inclination
    sheet["H14"] = format_mpd(col_data.get('Alluvium Bottom\n(mPD)', np.nan))
    sheet["H15"] = format_mpd(col_data.get('Tentative Concrete Column Bottom level \n(mPD)', np.nan))
    
    # Concrete Top Level 
    cutoff = col_data.get('Cut off level (mPD)', np.nan)
    sheet["M21"] = format_mpd(cutoff)
    sheet["M17"] = format_mpd(col_data.get('Concrete Top Level (mPD)', np.nan))

    # 2. Excavation Data
    def format_date(d):
        return d.strftime('%Y-%m-%d') if pd.notna(d) else ""
        
    sheet["H23"] = format_date(col_data.get('Drill Start Date \n(DD/MM/YYYY)', pd.NaT))
    sheet["H24"] = format_date(col_data.get('Drill Completed Date\n(DD/MM/YYYY)', pd.NaT))
    sheet["H25"] = format_mpd(col_data.get('Ground Level', np.nan))
    
    casing_top = col_data.get('Casing Top Level\n(mPD)', np.nan)
    act_toe = col_data.get('Actual Toe Level\n(mPD)', np.nan)
    sheet["H26"] = format_mpd(casing_top)
    sheet["H28"] = format_mpd(act_toe)
    
    # As-built Casing Length (Now user inputted)
    sheet["H27"] = format_float(col_data.get('As-built casing Length (m)', np.nan))
        
    act_founding = col_data.get('Actual Founding level', np.nan)
    sheet["H29"] = format_mpd(act_founding)
    
    # Intentionally leave Socket Lengths blank
    sheet["H30"] = ""
    sheet["H31"] = ""
        
    # 3. Concreting
    sheet["H36"] = "C45/20D" # Concrete Grade
    
    theo_vol = col_data.get('Estimate Concrete volume\n(m3)', np.nan)
    act_vol = col_data.get('Actual Concrete Volumn (m3)', np.nan)
    sheet["H37"] = format_float(theo_vol)
    sheet["H38"] = format_float(act_vol)
    
    # Calculate Overbreak %
    try:
        if float(theo_vol) > 0:
            overbreak = ((float(act_vol) - float(theo_vol)) / float(theo_vol)) * 100
            sheet["H39"] = format_float(overbreak)
        else:
            sheet["H39"] = ""
    except:
        sheet["H39"] = ""
        
    sheet["H40"] = format_date(col_data.get('Concrete Date\n(DD/MM/YYYY)', pd.NaT))
    
    wb.save(file_name)
    with open(file_name, "rb") as f:
        return f.read(), file_name

if df.empty:
    st.error("Could not load data.")
else:
    # --- TOP DASHBOARD PANEL ---
    st.markdown("### Overall Progress Overview")
    
    total_cols = len(df)
    completed_cols = len(df[df['Construction_Stage'] == 'Completed'])
    progress_pct = (completed_cols / total_cols) * 100 if total_cols > 0 else 0
    remaining_cols = total_cols - completed_cols
    
    top_col1, top_col2, top_col3 = st.columns([1.5, 1.5, 1])
    
    with top_col1:
        fig_gauge = go.Figure(go.Indicator(
            mode = "gauge+number",
            value = progress_pct,
            number = {'suffix': "%", 'valueformat': '.1f'},
            title = {'text': "Start → Today<br>Overall Completion"},
            gauge = {
                'axis': {'range': [0, 100], 'tickwidth': 1},
                'bar': {'color': "#00B0F0"},
                'bgcolor': "#E0F2F7",
                'shape': "angular",
            }
        ))
        fig_gauge.update_layout(height=250, margin=dict(l=10, r=10, t=40, b=10))
        st.plotly_chart(fig_gauge, use_container_width=True)
        
    with top_col2:
        st.markdown("#### Project Summary")
        st.info(f"**Total Clusters:** {total_cols} nos")
        st.info(f"**Remaining:** {remaining_cols} nos")
        
        vol_col = 'Actual Concrete Volumn (m3)'
        total_vol = pd.to_numeric(df[vol_col], errors='coerce').sum() if vol_col in df.columns else 0.0
        st.success(f"**Total Injected Volume:** {total_vol:.2f} m³")
        
    with top_col3:
        st.markdown("#### 📅 Period Analysis")
        today = datetime.date.today()
        start_default = today.replace(day=1)
        
        date_range = st.date_input("Select Date Range", value=(start_default, today), format="YYYY-MM-DD")
        
        if len(date_range) == 2:
            start_date, end_date = date_range
            pd_start, pd_end = pd.to_datetime(start_date), pd.to_datetime(end_date)
            
            drill_dates = pd.to_datetime(df['Drill Completed Date\n(DD/MM/YYYY)'], errors='coerce')
            conc_dates = pd.to_datetime(df['Concrete Date\n(DD/MM/YYYY)'], errors='coerce')
            
            drilled_mask = (drill_dates >= pd_start) & (drill_dates <= pd_end)
            concreted_mask = (conc_dates >= pd_start) & (conc_dates <= pd_end)
            
            st.metric("Holes Drilled in Period", drilled_mask.sum())
            st.metric("Concreted in Period", concreted_mask.sum())
        else:
            st.warning("Select a start and end date.")

    st.divider()

    # --- SIDEBAR & MAIN LAYOUT ---
    st.sidebar.header("Filter Options")
    stages = ["All", "Not Started", "In Progress", "Drilled", "Completed"]
    selected_stage = st.sidebar.selectbox("Filter by Stage", stages)
    
    filtered_df = df.copy()
    if selected_stage != "All":
        filtered_df = filtered_df[filtered_df['Construction_Stage'] == selected_stage]
        
    col_vis, col_data = st.columns([3, 2])

    with col_vis:
        st.markdown("### Site Plan Visualizations")
        color_map = {"Not Started": "lightgrey", "In Progress": "yellow", "Drilled": "blue", "Completed": "green"}
        
        tab1, tab2 = st.tabs(["🗺️ 2D Plan View", "🧊 3D Column Model"])
        
        with tab1:
            if not filtered_df.empty:
                fig2d = px.scatter(
                    filtered_df, x='X_Coord', y='Y_Coord', color='Construction_Stage', color_discrete_map=color_map,
                    hover_name='Concrete Column ID', hover_data=['Construction_Stage'],
                )
                fig2d.update_traces(marker=dict(size=14, line=dict(width=1, color='DarkSlateGrey')))
                fig2d.update_yaxes(scaleanchor="x", scaleratio=1)
                fig2d.update_layout(height=500)
                st.plotly_chart(fig2d, use_container_width=True)
            else:
                st.info("No data.")
            
        with tab2:
            if not filtered_df.empty:
                fig3d = px.scatter_3d(
                    filtered_df, x='X_Coord', y='Y_Coord', z='Tentative Concrete Column Bottom level \n(mPD)',
                    color='Construction_Stage', color_discrete_map=color_map, hover_name='Concrete Column ID',
                    hover_data=['Cut off level (mPD)', 'Foundation level\n(mPD)', 'Actual Founding level'],
                )
                for i, row in filtered_df.iterrows():
                    bottom_z = row['Actual Founding level']
                    if pd.isna(bottom_z): bottom_z = row['Tentative Concrete Column Bottom level \n(mPD)']
                    if pd.isna(bottom_z): bottom_z = 0 
                        
                    top_z = row['Cut off level (mPD)']
                    if pd.isna(top_z): top_z = 10 
                        
                    fig3d.add_scatter3d(
                        x=[row['X_Coord'], row['X_Coord']], y=[row['Y_Coord'], row['Y_Coord']],
                        z=[top_z, bottom_z], mode='lines',
                        line=dict(color=color_map.get(row['Construction_Stage'], 'grey'), width=5), showlegend=False
                    )
                fig3d.update_layout(scene=dict(zaxis=dict(title='Depth (mPD)')), height=500)
                st.plotly_chart(fig3d, use_container_width=True)
            else:
                st.info("No data.")

    with col_data:
        st.markdown("### 📝 Record As-Built Data")
        
        valid_cols = df['Concrete Column ID'].dropna().unique()
        if len(valid_cols) > 0:
            selected_col = st.selectbox("Select Column ID to Update", valid_cols)
            col_design_data = df[df['Concrete Column ID'] == selected_col].iloc[0]
            
            st.info(f"**Design Foundation Level:** {format_mpd(col_design_data.get('Foundation level\n(mPD)', np.nan))} mPD")
            
            with st.form("update_form"):
                current_stage = col_design_data.get('Construction_Stage', 'Not Started')
                if current_stage not in ["Not Started", "In Progress", "Drilled", "Completed"]:
                    current_stage = "Not Started"
                new_stage = st.selectbox("Update Status", ["Not Started", "In Progress", "Drilled", "Completed"], index=["Not Started", "In Progress", "Drilled", "Completed"].index(current_stage))
                
                st.markdown("##### Input Field Measurements")
                
                def get_val(key):
                    v = col_design_data.get(key, 0.0)
                    return float(v) if pd.notna(v) else 0.0
                
                # Retrieve existing values
                ground_level = st.number_input("Ground Level (mPD)", value=get_val('Ground Level'), format="%.3f")
                casing_top = st.number_input("Casing Top Level (mPD)", value=get_val('Casing Top Level\n(mPD)'), format="%.3f")
                act_toe = st.number_input("Actual Toe Level (mPD)", value=get_val('Actual Toe Level\n(mPD)'), format="%.3f")
                act_founding = st.number_input("Actual Founding Level (mPD)", value=get_val('Actual Founding level'), format="%.3f")
                
                # New Direct Inputs
                as_built_casing = st.number_input("As-built Casing Length (m)", value=get_val('As-built casing Length (m)'), format="%.3f")
                conc_top = st.number_input("Concrete Top Level (mPD)", value=get_val('Concrete Top Level (mPD)'), format="%.3f")
                
                concrete_vol = st.number_input("Actual Concrete Volume (m3)", value=get_val('Actual Concrete Volumn (m3)'), format="%.3f")
                
                st.markdown("##### Operation Dates")
                start_date = st.date_input("Drill Start Date")
                comp_date = st.date_input("Drill Completed Date")
                conc_date = st.date_input("Concrete Pour Date")
                
                submitted = st.form_submit_button("Update System Record")
                
                if submitted:
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
                        df.loc[df['Concrete Column ID'] == selected_col, 'column drilled'] = "Yes"
                    if new_stage == "Completed":
                        df.loc[df['Concrete Column ID'] == selected_col, 'Concrete Date\n(DD/MM/YYYY)'] = conc_date
                        df.loc[df['Concrete Column ID'] == selected_col, 'Completed'] = "Yes"
                    
                    save_path = "FYP_Concrete_Column_Data.csv" if os.path.exists("FYP_Concrete_Column_Data.csv") else "FYP_App_Data.csv"
                    df.to_csv(save_path, index=False)
                    st.success(f"Column {selected_col} updated! Refreshing...")
                    st.rerun()

        st.markdown("### 📄 Generate Record Sheet")
        st.write("Automatically calculate metrics and export formal Excel report.")
        
        updated_row = df[df['Concrete Column ID'] == selected_col].iloc[0]
        excel_data, file_name = generate_excel_report(updated_row)
        
        if excel_data is not None:
            st.download_button(
                label=f"📥 Download {selected_col} Record Sheet",
                data=excel_data,
                file_name=os.path.basename(file_name),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.caption(f"Saved locally to: `{file_name}`")
        else:
            st.warning(file_name)

    st.markdown("### As-Built vs Design Data View")
    
    display_df = filtered_df.copy()
    for mPD_col in ['Ground Level', 'Casing Top Level\n(mPD)', 'Foundation level\n(mPD)', 'Actual Founding level', 'Concrete Top Level (mPD)']:
        if mPD_col in display_df.columns:
            display_df[mPD_col] = display_df[mPD_col].apply(lambda x: f"{x:+.3f}" if pd.notna(x) and str(x).strip()!="" else "")

    columns_to_show = [c for c in [
        'Concrete Column ID', 'Construction_Stage', 
        'Ground Level', 'Casing Top Level\n(mPD)',
        'Concrete Top Level (mPD)', 'As-built casing Length (m)',
        'Foundation level\n(mPD)', 'Actual Founding level', 
        'Actual Concrete Volumn (m3)', 'Concrete Date\n(DD/MM/YYYY)'
    ] if c in display_df.columns]
    
    st.dataframe(display_df[columns_to_show], height=250)
