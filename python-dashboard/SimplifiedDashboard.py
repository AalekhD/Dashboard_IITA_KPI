import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
from openpyxl.utils import get_column_letter
import os
import numpy as np

# Page config
st.set_page_config(page_title="IITA KPI Dashboard", layout="wide")

# Header
st.markdown("""
<div style="background-color:#00891a; padding:20px; border-radius:10px;">
    <h1 style="color:#ffffff; text-align:center; margin:0;">🌱 IITA KPI Dashboard</h1>
    <p style="color:#ffffff; text-align:center; margin:5px;">IITA Programs and Service Unit KPIs</p>
</div>
""", unsafe_allow_html=True)

st.write("")

# Load Excel files and convert to HTML with merged cells
@st.cache_data
def load_kpi_data():
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    # Load Program Output KPIs
    program_file = os.path.join(root_dir, 'data', 'Program Output KPIs.xlsx')
    df_programs = pd.read_excel(program_file)
    
    # Load Service Unit KPIs
    service_file = os.path.join(root_dir, 'data', 'Service Unit KPIs.xlsx')
    df_services = pd.read_excel(service_file)
    
    # Load KPI Heat map
    heatmap_file = os.path.join(root_dir, 'data', 'KPI by Nr. Heat map.xlsx')
    df_heatmap = pd.read_excel(heatmap_file)
    
    return df_programs, df_services, df_heatmap

# Function to convert Excel with merged cells to HTML
def excel_to_html_with_merged_cells(excel_file_path):
    # Load workbook with data_only=True to get calculated values instead of formulas
    wb_data = openpyxl.load_workbook(excel_file_path, data_only=True)
    ws_data = wb_data.active
    
    # Load workbook normally to get formatting info (merged cells, number formats)
    wb_format = openpyxl.load_workbook(excel_file_path)
    ws_format = wb_format.active
    
    # Find actual data range (trim empty rows and columns)
    max_row = 0
    max_col = 0
    
    for row_idx, row in enumerate(ws_data.iter_rows(values_only=True), 1):
        has_data = any(cell is not None for cell in row)
        if has_data:
            max_row = row_idx
            # Find max column with data in this row
            for col_idx, cell in enumerate(row, 1):
                if cell is not None:
                    max_col = max(max_col, col_idx)
    
    # Get merged cell ranges from format workbook
    merged_cells = {}
    for merged_range in ws_format.merged_cells.ranges:
        cells = list(merged_range.cells)
        for cell in cells[1:]:  # Skip the first cell (top-left)
            merged_cells[cell] = cells[0]  # Map to top-left cell
    
    # Build HTML table
    html = '<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; background-color: white;">'
    html += '<style>td, th { border: 1px solid #999; padding: 12px; text-align: left; background-color: white; white-space: pre-wrap; word-wrap: break-word; }</style>'
    
    processed = set()
    
    for row_idx in range(1, max_row + 1):
        # Get the actual row
        row_data = list(ws_data.iter_rows(min_row=row_idx, max_row=row_idx, values_only=False))[0]
        
        # Check if this row is completely empty
        has_row_data = False
        for col_idx in range(1, max_col + 1):
            if row_data[col_idx - 1].value is not None:
                has_row_data = True
                break
        
        # Skip completely empty rows
        if not has_row_data:
            continue
        
        html += '<tr>'
        for col_idx in range(1, max_col + 1):
            cell_data = row_data[col_idx - 1]
            cell_coord = cell_data.coordinate
            
            # Get corresponding format cell
            cell_format = ws_format[cell_coord]
            
            # Skip if this cell is part of a merged range (not the top-left)
            if cell_coord in merged_cells and merged_cells[cell_coord] != cell_coord:
                continue
            
            # Skip if already processed
            if cell_coord in processed:
                continue
            
            # Calculate rowspan and colspan for merged cells
            rowspan = 1
            colspan = 1
            
            for merged_range in ws_format.merged_cells.ranges:
                if cell_coord in merged_range:
                    rowspan = merged_range.max_row - merged_range.min_row + 1
                    colspan = merged_range.max_col - merged_range.min_col + 1
                    # Mark all cells in this range as processed
                    for r in range(merged_range.min_row, merged_range.max_row + 1):
                        for c in range(merged_range.min_col, merged_range.max_col + 1):
                            processed.add(f"{get_column_letter(c)}{r}")
                    break
            
            # Get cell value from data workbook (contains calculated values, not formulas)
            cell_value = cell_data.value
            
            # Format based on cell number format
            if cell_value is not None:
                if isinstance(cell_value, (int, float)):
                    # Check if the cell has percentage format
                    if cell_format.number_format and '%' in cell_format.number_format:
                        cell_value = f"{cell_value * 100:.2f}%"
                    else:
                        cell_value = str(cell_value)
                else:
                    cell_value = str(cell_value)
            else:
                cell_value = ""
            
            # Add styling for headers (first row)
            if row_idx == 1:
                html += f'<th style="background-color: #00891a; color: white; font-weight: bold;" rowspan="{rowspan}" colspan="{colspan}">{cell_value}</th>'
            else:
                html += f'<td rowspan="{rowspan}" colspan="{colspan}">{cell_value}</td>'
        
        html += '</tr>'
    
    html += '</table>'
    return html

# Function to create heatmap from KPI heat map file
def create_heatmap_visualization(excel_file_path):
    try:
        # Load with openpyxl to get clean numeric data
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        ws = wb.active
        
        # Extract heatmap data starting from row 3 (headers) and row 4 (data)
        # Column C has program names, columns D onwards have KPI values
        
        programs = []
        kpi_names = []
        data_values = []
        original_values = []
        
        # Debug: Show file info
        # st.write(f"📊 File: {excel_file_path}")
        # st.write(f"Max rows: {ws.max_row}, Max cols: {ws.max_column}")
        
        # Get KPI names from row 3 (starting from column D=4)
        for col_idx in range(4, ws.max_column + 1):
            cell_val = ws.cell(row=3, column=col_idx).value
            if cell_val is not None:
                kpi_names.append(str(cell_val))
        
        # Get data from rows 4 onwards, columns C (program) and D onwards (values)
        for row_idx in range(4, ws.max_row + 1):
            program_cell = ws.cell(row=row_idx, column=3).value
            if program_cell is not None and program_cell != "None":
                programs.append(str(program_cell))
                row_data = []
                orig_data = []
                for col_idx in range(4, 4 + len(kpi_names)):
                    val = ws.cell(row=row_idx, column=col_idx).value
                    # Convert to float if possible
                    try:
                        val = float(val) if val is not None else 0
                    except:
                        val = 0
                    row_data.append(val)
                    orig_data.append(val)
                data_values.append(row_data)
                original_values.append(orig_data)
        
        # Debug output (disabled)
        # st.write(f"📊 **Data Info:**")
        # st.write(f"Programs found: {len(programs)}")
        # st.write(f"KPIs found: {len(kpi_names)}")
        # if programs:
        #     st.write(f"Categories/Programs: {programs[:5]}...")  # Show first 5
        # if kpi_names:
        #     st.write(f"KPI Names: {kpi_names}")
        
        # Create dataframe
        if data_values and len(kpi_names) > 0:
            df_heatmap = pd.DataFrame(data_values, columns=kpi_names[:len(data_values[0])])
            df_heatmap.index = programs
            
            # Replace 0 or NaN values with NaN to exclude them from coloring
            df_heatmap_for_color = df_heatmap.replace(0, np.nan)
            
            # Normalize each column (KPI) independently: scale to 0-1 range per column
            df_normalized = df_heatmap_for_color.copy()
            for col in df_normalized.columns:
                col_values = df_heatmap_for_color[col].dropna()
                if len(col_values) > 0:
                    min_val = col_values.min()
                    max_val = col_values.max()
                    mid_val = (max_val + min_val) / 2
                    
                    # Normalize: values from 0 (min) to 1 (max)
                    if max_val > min_val:
                        df_normalized[col] = (df_heatmap_for_color[col] - min_val) / (max_val - min_val)
                    else:
                        df_normalized[col] = 0.5  # If all values are the same
            
            # Use continuous color scale: Red (0/min) → Yellow (0.5/mid) → Green (1/max)
            colorscale = [
                [0.0, '#FF0000'],      # Red for lowest values
                [0.5, '#FFFF00'],      # Yellow for midpoint
                [1.0, '#00AA00']       # Green for highest values
            ]

            fig = px.imshow(df_normalized,
                            labels=dict(x="KPI", y="Program", color="Value"),
                            color_continuous_scale=colorscale,
                            text_auto=False,
                            aspect="auto",
                            zmin=0,
                            zmax=1)
            
            # Add original values as text
            fig.update_traces(
                text=np.array(original_values).round(1),
                texttemplate='%{text}',
                textfont=dict(size=12, color='black')
            )
            
            fig.update_layout(
                height=600 + (len(programs) * 20),  # Dynamic height based on number of programs
                xaxis_title="KPI Type", 
                yaxis_title="Program",
                title="KPI Heat Map - Programs vs KPI Count (Color scale per column)",
                xaxis_tickangle=-45,
                margin=dict(l=200, r=50, t=80, b=150),  # Better margins for labels
                coloraxis_colorbar=dict(title="Normalized Value")
            )
            
            return fig
        else:
            st.error(f"No data found. Programs: {len(programs)}, KPIs: {len(kpi_names)}")
            return None
    except Exception as e:
        st.error(f"Error creating heatmap: {str(e)}")
        return None

# Load data
df_programs, df_services, df_heatmap = load_kpi_data()

# Tabs
tab1, tab2, tab3 = st.tabs(["📊 Program Output KPIs", "🏢 Service Unit KPIs", "🔥 KPI by Nr, FTE, $ and Time"])

# Programs Tab
with tab1:
    st.subheader("📊 Program Output KPIs")
    
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    program_file = os.path.join(root_dir, 'data', 'Program Output KPIs.xlsx')
    
    try:
        html_programs = excel_to_html_with_merged_cells(program_file)
        st.markdown(html_programs, unsafe_allow_html=True)
    except Exception as e:
        st.warning(f"Could not render with merged cells: {str(e)}")
        st.dataframe(df_programs, width='stretch', height=600)
    
    # Download button
    csv_programs = df_programs.to_csv(index=False)
    st.download_button(
        label="⬇️ Download Program KPIs as CSV",
        data=csv_programs,
        file_name="Program_Output_KPIs.csv",
        mime="text/csv"
    )

# Service Units Tab
with tab2:
    st.subheader("🏢 Service Unit Output KPIs")
    
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    service_file = os.path.join(root_dir, 'data', 'Service Unit KPIs.xlsx')
    
    try:
        html_services = excel_to_html_with_merged_cells(service_file)
        st.markdown(html_services, unsafe_allow_html=True)
    except Exception as e:
        st.warning(f"Could not render with merged cells: {str(e)}")
        st.dataframe(df_services, width='stretch', height=600)
    # Download button
    csv_services = df_services.to_csv(index=False)
    st.download_button(
        label="⬇️ Download Service Unit KPIs as CSV",
        data=csv_services,
        file_name="Service_Unit_KPIs.csv",
        mime="text/csv"
    )

# KPI by Nr, FTE, $ and Time Tab
with tab3:
    st.subheader("🔥 KPI by Nr, FTE, $ and Time")
    
    # Two main sub-tabs
    sub_tab_a, sub_tab_b = st.tabs(["🔬 Research, Training, Product Development", "🏆 Recognition, Societal Impact & Inclusivity"])
    
    # ==================== Research, Training, Product Development ====================
    with sub_tab_a:
        st.markdown("### Research, Training, Product Development")
        
        rtpd_tabs = st.tabs(["KPI by Nr", "KPI by FTE", "KPI by $", "KPI over Time"])
        
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        with rtpd_tabs[0]:
            st.write("**Research, Training, Product Development - KPI by Nr**")
            plot_container = st.empty()
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 1.xlsx')
                if os.path.exists(heatmap_file):
                    fig = create_heatmap_visualization(heatmap_file)
                    if fig:
                        with plot_container.container():
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    plot_container.info("📁 Waiting for: Heat map 1.xlsx")
            except Exception as e:
                plot_container.warning(f"Could not load heatmap: {str(e)}")
        
        with rtpd_tabs[1]:
            st.write("**Research, Training, Product Development - KPI by FTE**")
            plot_container = st.empty()
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 2.xlsx')
                if os.path.exists(heatmap_file):
                    fig = create_heatmap_visualization(heatmap_file)
                    if fig:
                        with plot_container.container():
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    plot_container.info("📁 Waiting for: Heat map 2.xlsx")
            except Exception as e:
                plot_container.warning(f"Could not load heatmap: {str(e)}")
        
        with rtpd_tabs[2]:
            st.write("**Research, Training, Product Development - KPI by $**")
            plot_container = st.empty()
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 3.xlsx')
                if os.path.exists(heatmap_file):
                    fig = create_heatmap_visualization(heatmap_file)
                    if fig:
                        with plot_container.container():
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    plot_container.info("📁 Waiting for: Heat map 3.xlsx")
            except Exception as e:
                plot_container.warning(f"Could not load heatmap: {str(e)}")
        
        with rtpd_tabs[3]:
            st.write("**Research, Training, Product Development - KPI over Time**")
            plot_container = st.empty()
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 4.xlsx')
                if os.path.exists(heatmap_file):
                    fig = create_heatmap_visualization(heatmap_file)
                    if fig:
                        with plot_container.container():
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    plot_container.info("📁 Waiting for: Heat map 4.xlsx")
            except Exception as e:
                plot_container.warning(f"Could not load heatmap: {str(e)}")
    
    # ==================== Recognition, Societal Impact & Inclusivity ====================
    with sub_tab_b:
        st.markdown("### Recognition, Societal Impact & Inclusivity")
        
        rsi_tabs = st.tabs(["KPI by Nr", "KPI by FTE", "KPI by $", "KPI over Time"])
        
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        with rsi_tabs[0]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by Nr**")
            plot_container = st.empty()
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 5.xlsx')
                if os.path.exists(heatmap_file):
                    fig = create_heatmap_visualization(heatmap_file)
                    if fig:
                        with plot_container.container():
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    plot_container.info("📁 Waiting for: Heat map 5.xlsx")
            except Exception as e:
                plot_container.warning(f"Could not load heatmap: {str(e)}")
        
        with rsi_tabs[1]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by FTE**")
            plot_container = st.empty()
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 6.xlsx')
                if os.path.exists(heatmap_file):
                    fig = create_heatmap_visualization(heatmap_file)
                    if fig:
                        with plot_container.container():
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    plot_container.info("📁 Waiting for: Heat map 6.xlsx")
            except Exception as e:
                plot_container.warning(f"Could not load heatmap: {str(e)}")
        
        with rsi_tabs[2]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by $**")
            plot_container = st.empty()
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 7.xlsx')
                if os.path.exists(heatmap_file):
                    fig = create_heatmap_visualization(heatmap_file)
                    if fig:
                        with plot_container.container():
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    plot_container.info("📁 Waiting for: Heat map 7.xlsx")
            except Exception as e:
                plot_container.warning(f"Could not load heatmap: {str(e)}")
        
        with rsi_tabs[3]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI over Time**")

st.markdown("---")
st.caption("Last updated: April 8, 2026 | IITA KPI Dashboard")
