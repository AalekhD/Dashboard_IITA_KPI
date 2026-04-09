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
# Returns (fig, df_below) where df_below contains rows beyond row 17 (or None)
def create_heatmap_visualization(excel_file_path, heatmap_max_row=17):
    try:
        # Load with openpyxl to get clean numeric data
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        ws = wb.active

        programs = []
        program_groups = []   # column B: GI / RAFS / ST etc.
        kpi_names = []
        kpi_type_groups = []  # row 2: Research Outputs / Training etc.
        data_values = []
        original_values = []

        # Programs/rows below heatmap_max_row
        below_programs = []
        below_data = []

        # Build program-group map from column B (fill-forward for merged cells)
        current_b = None
        b_values = {}
        for row_idx in range(4, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=2).value
            if val is not None:
                current_b = str(val)
            b_values[row_idx] = current_b

        # Get KPI names from row 3 (starting from column D=4)
        for col_idx in range(4, ws.max_column + 1):
            cell_val = ws.cell(row=3, column=col_idx).value
            if cell_val is not None:
                kpi_names.append(str(cell_val))

        # Get KPI type groups from row 2 (fill-forward for merged cells)
        current_type = None
        for col_idx in range(4, 4 + len(kpi_names)):
            val = ws.cell(row=2, column=col_idx).value
            if val is not None:
                current_type = str(val)
            kpi_type_groups.append(current_type or "")

        # Helper: wrap a label at word boundaries for ~max_width chars per line
        def wrap_label(text, max_width=15):
            words = text.split()
            lines, current = [], ""
            for w in words:
                if current and len(current) + 1 + len(w) > max_width:
                    lines.append(current)
                    current = w
                else:
                    current = (current + " " + w).strip()
            if current:
                lines.append(current)
            return "<br>".join(lines)

        # Compute KPI type group spans for header annotations
        kpi_type_spans = []
        if kpi_type_groups:
            span_start = 0
            span_name = kpi_type_groups[0]
            for i, grp in enumerate(kpi_type_groups[1:], 1):
                if grp != span_name:
                    kpi_type_spans.append((span_name, span_start, i - 1))
                    span_name = grp
                    span_start = i
            kpi_type_spans.append((span_name, span_start, len(kpi_type_groups) - 1))

        # Get data from rows 4 to heatmap_max_row for the heatmap
        for row_idx in range(4, heatmap_max_row + 1):
            program_cell = ws.cell(row=row_idx, column=3).value
            if program_cell is not None and program_cell != "None":
                programs.append(str(program_cell))
                program_groups.append(b_values.get(row_idx) or "")
                row_data = []
                orig_data = []
                for col_idx in range(4, 4 + len(kpi_names)):
                    val = ws.cell(row=row_idx, column=col_idx).value
                    try:
                        fval = float(val) if val is not None else None
                    except:
                        fval = None
                    # Store None for missing/zero so we can display "NA"
                    row_data.append(fval if fval else np.nan)
                    orig_data.append(fval)
                data_values.append(row_data)
                original_values.append(orig_data)

        # Collect rows beyond heatmap_max_row as "below" data
        for row_idx in range(heatmap_max_row + 1, ws.max_row + 1):
            program_cell = ws.cell(row=row_idx, column=3).value
            if program_cell is not None and program_cell != "None":
                below_programs.append(str(program_cell))
                row_data = []
                for col_idx in range(4, 4 + len(kpi_names)):
                    val = ws.cell(row=row_idx, column=col_idx).value
                    try:
                        val = float(val) if val is not None else 0
                    except:
                        val = 0
                    row_data.append(val)
                below_data.append(row_data)

        # Build below-heatmap dataframe if there is extra data
        df_below = None
        if below_data and kpi_names:
            df_below = pd.DataFrame(below_data, columns=kpi_names[:len(below_data[0])])
            df_below.insert(0, "Program", below_programs)

        # Compute program group spans for left-side bands (mirrors kpi_type_spans on y-axis)
        program_group_spans = []
        if program_groups:
            span_start = 0
            span_name = program_groups[0]
            for i, grp in enumerate(program_groups[1:], 1):
                if grp != span_name:
                    program_group_spans.append((span_name, span_start, i - 1))
                    span_name = grp
                    span_start = i
            program_group_spans.append((span_name, span_start, len(program_groups) - 1))

        # y-axis labels: plain program names only (group shown as side band)
        y_labels = programs

        # Wrapped x-axis tick labels (two lines each)
        wrapped_kpi_names = [wrap_label(k) for k in kpi_names]

        # Create dataframe
        if data_values and len(kpi_names) > 0:
            df_heatmap = pd.DataFrame(data_values, columns=kpi_names[:len(data_values[0])])
            df_heatmap.index = y_labels

            # NaN values (0/None in Excel) are excluded from coloring
            df_heatmap_for_color = df_heatmap.copy()

            # Normalize each column independently: 0 (min) to 1 (max)
            df_normalized = df_heatmap_for_color.copy()
            for col in df_normalized.columns:
                col_values = df_heatmap_for_color[col].dropna()
                if len(col_values) > 0:
                    min_val = col_values.min()
                    max_val = col_values.max()
                    if max_val > min_val:
                        df_normalized[col] = (df_heatmap_for_color[col] - min_val) / (max_val - min_val)
                    else:
                        df_normalized[col] = 0.5

            colorscale = [
                [0.0, '#FF0000'],
                [0.5, '#FFFF00'],
                [1.0, '#00AA00']
            ]

            # Build text display: "NA" for zero/null, rounded number otherwise
            text_display = []
            for row in original_values:
                text_row = []
                for v in row:
                    if v is None or v == 0:
                        text_row.append("NA")
                    else:
                        text_row.append(str(round(v, 1)))
                text_display.append(text_row)

            # Build figure using numpy array with explicit x/y labels
            fig = px.imshow(
                df_normalized.values,
                x=kpi_names,
                y=y_labels,
                labels=dict(x="KPI", y="Program", color="Value"),
                color_continuous_scale=colorscale,
                text_auto=False,
                aspect="auto",
                zmin=0,
                zmax=1
            )

            # Overlay text (NA or rounded value)
            fig.update_traces(
                text=np.array(text_display, dtype=object),
                texttemplate='%{text}',
                textfont=dict(size=11, color='black')
            )

            # Move x-axis to top with wrapped tick labels
            fig.update_layout(
                height=700 + (len(programs) * 25),
                xaxis=dict(
                    side='top',
                    tickangle=0,
                    title='',
                    tickmode='array',
                    tickvals=kpi_names,
                    ticktext=wrapped_kpi_names,
                    tickfont=dict(size=10),
                    automargin=True,
                ),
                yaxis_title="Program",
                title="KPI Heat Map - Programs vs KPI Count (Color scale per column)",
                margin=dict(l=380, r=60, t=450, b=60),
                coloraxis_colorbar=dict(title="Normalized Value")
            )

            # Add KPI type group header rectangles + labels above the x-axis
            # y0/y1 are in paper coords (1.0 = top of plot area).
            # Tick labels (2 lines, size 10, horizontal) occupy ~1.00–1.10.
            # Group bands sit above that at 1.12–1.22.
            group_colors = [
                '#00891a', '#005a8e', '#8e4f00', '#6a008e',
                '#8e0000', '#006e6e', '#4a4a00', '#00456e'
            ]
            for idx, (group_name, start_idx, end_idx) in enumerate(kpi_type_spans):
                color = group_colors[idx % len(group_colors)]
                x_center = (start_idx + end_idx) / 2
                fig.add_shape(
                    type='rect',
                    xref='x', yref='paper',
                    x0=start_idx - 0.5, x1=end_idx + 0.5,
                    y0=1.12, y1=1.22,
                    fillcolor=color,
                    line=dict(color='white', width=1),
                    layer='above'
                )
                fig.add_annotation(
                    xref='x', yref='paper',
                    x=x_center, y=1.17,
                    text=f"<b>{group_name}</b>",
                    showarrow=False,
                    font=dict(color='white', size=10),
                    align='center',
                    bgcolor='rgba(0,0,0,0)'
                )

            # Add program group bands to the LEFT of the y-axis (mirrors top KPI-type bands)
            # xref='paper': x0/x1 are fractions of the plot-area width (negative = in left margin)
            # yref='y': y0/y1 are row indices (0 = first row, reversed axis puts it at top)
            for idx, (group_name, start_idx, end_idx) in enumerate(program_group_spans):
                color = group_colors[idx % len(group_colors)]
                y_center = (start_idx + end_idx) / 2
                fig.add_shape(
                    type='rect',
                    xref='paper', yref='y',
                    x0=-0.22, x1=-0.10,
                    y0=start_idx - 0.5, y1=end_idx + 0.5,
                    fillcolor=color,
                    line=dict(color='white', width=1),
                    layer='above'
                )
                fig.add_annotation(
                    xref='paper', yref='y',
                    x=-0.16, y=y_center,
                    text=f"<b>{group_name}</b>",
                    showarrow=False,
                    font=dict(color='white', size=10),
                    align='center',
                    textangle=-90,
                    bgcolor='rgba(0,0,0,0)'
                )

            return fig, df_below
        else:
            st.error(f"No data found. Programs: {len(programs)}, KPIs: {len(kpi_names)}")
            return None, None
    except Exception as e:
        st.error(f"Error creating heatmap: {str(e)}")
        return None, None

# Load data
df_programs, df_services, df_heatmap = load_kpi_data()

# Tabs
tab1, tab2, tab3 = st.tabs(["📊 Program Output KPIs", "🏢 Service Unit KPIs", "� KPI By Program"])

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

# KPI By Program Tab
with tab3:
    st.subheader("📊 KPI By Program")
    
    # Two main sub-tabs
    sub_tab_a, sub_tab_b = st.tabs(["🔬 Research, Training, Product Development", "🏆 Recognition, Societal Impact & Inclusivity"])
    
    # ==================== Research, Training, Product Development ====================
    with sub_tab_a:
        st.markdown("### Research, Training, Product Development")
        
        rtpd_tabs = st.tabs(["KPI by Nr", "KPI by FTE", "KPI by $", "KPI over Time"])
        
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        with rtpd_tabs[0]:
            st.write("**Research, Training, Product Development - KPI by Nr**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 1.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True)
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        st.dataframe(df_below, use_container_width=True, hide_index=True)
                else:
                    st.info("📁 Waiting for: Heat map 1.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rtpd_tabs[1]:
            st.write("**Research, Training, Product Development - KPI by FTE**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 2.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True)
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        st.dataframe(df_below, use_container_width=True, hide_index=True)
                else:
                    st.info("📁 Waiting for: Heat map 2.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rtpd_tabs[2]:
            st.write("**Research, Training, Product Development - KPI by $**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 3.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True)
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        st.dataframe(df_below, use_container_width=True, hide_index=True)
                else:
                    st.info("📁 Waiting for: Heat map 3.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rtpd_tabs[3]:
            st.write("**Research, Training, Product Development - KPI over Time**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 4.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True)
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        st.dataframe(df_below, use_container_width=True, hide_index=True)
                else:
                    st.info("📁 Waiting for: Heat map 4.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")
    
    # ==================== Recognition, Societal Impact & Inclusivity ====================
    with sub_tab_b:
        st.markdown("### Recognition, Societal Impact & Inclusivity")
        
        rsi_tabs = st.tabs(["KPI by Nr", "KPI by FTE", "KPI by $", "KPI over Time"])
        
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        with rsi_tabs[0]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by Nr**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 5.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True)
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        st.dataframe(df_below, use_container_width=True, hide_index=True)
                else:
                    st.info("📁 Waiting for: Heat map 5.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rsi_tabs[1]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by FTE**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 6.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True)
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        st.dataframe(df_below, use_container_width=True, hide_index=True)
                else:
                    st.info("📁 Waiting for: Heat map 6.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rsi_tabs[2]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by $**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 7.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True)
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        st.dataframe(df_below, use_container_width=True, hide_index=True)
                else:
                    st.info("📁 Waiting for: Heat map 7.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rsi_tabs[3]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI over Time**")

st.markdown("---")
st.caption("Last updated: April 8, 2026 | IITA KPI Dashboard")
