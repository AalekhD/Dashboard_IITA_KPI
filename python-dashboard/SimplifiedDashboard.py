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
    html = '<table style="border-collapse: collapse; width: 100%; table-layout: fixed; font-family: Arial, sans-serif; background-color: white;">'
    html += '<style>td, th { border: 1px solid #999; padding: 12px; text-align: left; background-color: white; white-space: normal; word-wrap: break-word; word-break: break-word; overflow-wrap: break-word; color: black; font-size: 9pt; }</style>'
    
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
# Returns (fig, df_below) where df_below contains rows beyond row 16 (or None)
def create_heatmap_visualization(excel_file_path, heatmap_max_row=16):
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
        below_orig = []

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

        # Truncate to ~34 chars to match Excel row-3 height (191px) at 9pt font (~5.5px/char)
        def truncate_label(text, max_len=34):
            return text[:max_len] + '\u2026' if len(text) > max_len else text

        # Wrap label text at word boundaries, inserting <br> for Plotly HTML rendering
        def wrap_label(text, max_len=15):
            words = text.split()
            lines = []
            current = ''
            for word in words:
                if current and len(current) + 1 + len(word) > max_len:
                    lines.append(current)
                    current = word
                else:
                    current = (current + ' ' + word).strip()
            if current:
                lines.append(current)
            return '<br>'.join(lines)

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
                orig_data = []  # stores (fval, raw_val) tuples
                for col_idx in range(4, 4 + len(kpi_names)):
                    val = ws.cell(row=row_idx, column=col_idx).value
                    try:
                        fval = float(val) if val is not None else None
                    except:
                        fval = None
                    # Store NaN only for truly missing values (None); 0 is a valid value
                    row_data.append(fval if fval is not None else np.nan)
                    orig_data.append((fval, val))  # keep raw value to detect N/A text
                data_values.append(row_data)
                original_values.append(orig_data)

        # Collect rows beyond heatmap_max_row � merged into heatmap as gray rows
        for row_idx in range(heatmap_max_row + 1, ws.max_row + 1):
            program_cell = ws.cell(row=row_idx, column=3).value
            if program_cell is not None and program_cell != "None":
                below_programs.append(str(program_cell))
                row_data = []
                orig_row = []  # stores (fval, raw_val) tuples
                for col_idx in range(4, 4 + len(kpi_names)):
                    val = ws.cell(row=row_idx, column=col_idx).value
                    try:
                        fval = float(val) if val is not None else None
                    except:
                        fval = None
                    row_data.append(fval if fval is not None else np.nan)
                    orig_row.append((fval, val))
                below_data.append(row_data)
                below_orig.append(orig_row)

        df_below = None  # no separate table

        # Merge all rows; track which are "below" for gray coloring
        all_programs = programs + below_programs
        all_data_values = data_values + below_data
        all_original_values = original_values + below_orig
        all_program_groups = program_groups + [''] * len(below_programs)

        # Compute program group spans for left-side bands
        program_group_spans = []
        if all_program_groups:
            span_start = 0
            span_name = all_program_groups[0]
            for i, grp in enumerate(all_program_groups[1:], 1):
                if grp != span_name:
                    program_group_spans.append((span_name, span_start, i - 1))
                    span_name = grp
                    span_start = i
            program_group_spans.append((span_name, span_start, len(all_program_groups) - 1))

        # y-axis labels: plain program names only (group shown as side band)
        y_labels = all_programs

        # X-axis: word-wrapped bold labels at -45�
        MAX_CHARS = 15  # chars per line before wrapping
        CHAR_PX   = 7   # approx px per char at 10pt font
        def bold_wrap(text):
            lines = wrap_label(text, max_len=MAX_CHARS).split('<br>')
            return '<br>'.join(f'<b>{line}</b>' for line in lines)
        kpi_tick_names = [bold_wrap(k) for k in kpi_names]
        # Y-axis: word-wrapped bold labels (horizontal)
        y_labels_wrapped = ['<br>'.join(f'<b>{line}</b>' for line in wrap_label(p, max_len=18).split('<br>')) for p in all_programs]

        # Dynamically find the Per Program Target row by searching col A-C for the label
        per_program_target_row = None
        for row_idx in range(heatmap_max_row + 1, ws.max_row + 1):
            for col_idx in range(1, 4):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val and 'per program' in str(val).lower():
                    per_program_target_row = row_idx
                    break
            if per_program_target_row:
                break

        per_program_targets = []
        if per_program_target_row:
            for col_idx in range(4, 4 + len(kpi_names)):
                val = ws.cell(row=per_program_target_row, column=col_idx).value
                try:
                    per_program_targets.append(float(val) if val is not None else None)
                except:
                    per_program_targets.append(None)
        else:
            per_program_targets = [None] * len(kpi_names)

        # Create dataframe using all rows
        if all_data_values and len(kpi_names) > 0:
            df_heatmap = pd.DataFrame(all_data_values, columns=kpi_names[:len(all_data_values[0])])
            df_heatmap.index = y_labels

            # Normalize using 3-point scale per column:
            #   min_val (lowest in rows 4-16) -> 0.0 (red)
            #   midpoint (per_program_target / 2)  -> 0.5 (yellow)
            #   per_program_target                 -> 1.0 (dark green)
            df_normalized = df_heatmap.copy()
            for col_i, col in enumerate(df_normalized.columns):
                # Use orig values (includes 0.0 correctly) to compute min
                hm_orig_floats = pd.Series(
                    [all_original_values[r][col_i][0] for r in range(len(programs))
                     if r < len(all_original_values) and col_i < len(all_original_values[r])
                     and isinstance(all_original_values[r][col_i], tuple)
                     and all_original_values[r][col_i][0] is not None],
                    dtype=float
                ).dropna()
                min_val  = hm_orig_floats.min() if len(hm_orig_floats) else 0.0
                target   = per_program_targets[col_i] if col_i < len(per_program_targets) and per_program_targets[col_i] else None
                midpoint = target / 2.0 if target else None
                for r in range(len(all_programs)):
                    if r >= len(programs):
                        df_normalized.iloc[r, col_i] = np.nan
                        continue
                    orig_item = all_original_values[r][col_i] if r < len(all_original_values) and col_i < len(all_original_values[r]) else (None, None)
                    orig_fval = orig_item[0] if isinstance(orig_item, tuple) else None
                    orig_raw  = orig_item[1] if isinstance(orig_item, tuple) else orig_item
                    is_na_text = isinstance(orig_raw, str) and orig_raw.strip().upper() in ('N/A', 'NA', '#N/A')
                    is_na = is_na_text or orig_fval is None
                    if is_na:
                        df_normalized.iloc[r, col_i] = np.nan
                    elif target is None or target == min_val:
                        # No target defined — numeric values shown as green
                        df_normalized.iloc[r, col_i] = 1.0
                    elif orig_fval >= target:
                        # At or above target → always green (check before min_val)
                        df_normalized.iloc[r, col_i] = 1.0
                    elif orig_fval <= min_val:
                        df_normalized.iloc[r, col_i] = 0.0
                    elif midpoint and orig_fval <= midpoint:
                        span = midpoint - min_val
                        df_normalized.iloc[r, col_i] = 0.5 * (orig_fval - min_val) / span if span > 0 else 0.25
                    else:
                        span = target - midpoint
                        df_normalized.iloc[r, col_i] = 0.5 + 0.5 * (orig_fval - midpoint) / span if span > 0 else 0.75

            # 3-point colorscale: red → orange/yellow → dark green
            colorscale = [
                [0.0,  '#D73027'],  # dark red  (at min)
                [0.25, '#F46D43'],  # orange
                [0.5,  '#FFFF00'],  # yellow (at midpoint = target/2)
                [0.75, '#A6D96A'],  # light green
                [1.0,  '#1A7A1A'],  # dark green (at or above per-program target)
            ]

            # Smart numeric formatter — preserves significant figures for small values
            def fmt_val(v):
                if v == 0:
                    return '0'
                abs_v = abs(v)
                if abs_v < 0.001:
                    return f"{v:.3e}"          # e.g. 7.519e-04
                elif abs_v < 0.01:
                    return f"{v:.6f}".rstrip('0').rstrip('.')  # e.g. 0.007519
                elif abs_v < 0.1:
                    return f"{v:.4f}".rstrip('0').rstrip('.')  # e.g. 0.0752
                elif abs_v < 1:
                    return f"{v:.3f}".rstrip('0').rstrip('.')  # e.g. 0.752
                elif v == int(v):
                    return str(int(v))                          # e.g. 15
                else:
                    return f"{v:.1f}"                          # e.g. 15.5

            # Build text display
            text_display = []
            for row in all_original_values:
                text_row = []
                for item in row:
                    fval, raw = item if isinstance(item, tuple) else (item, None)
                    # Show NA only for blank cells or cells containing N/A text; 0 displays as 0
                    is_na_text = isinstance(raw, str) and raw.strip().upper() in ('N/A', 'NA', '#N/A')
                    if fval is None and not is_na_text and raw is not None and raw != '':
                        # non-numeric text that isn't N/A — show as-is
                        text_row.append(str(raw))
                    elif fval is None or is_na_text:
                        text_row.append('NA')
                    else:
                        text_row.append(fmt_val(fval))
                text_display.append(text_row)

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

            # Overlay text (NA or rounded value) and add cell grid via gaps
            fig.update_traces(
                text=np.array(text_display, dtype=object),
                texttemplate='%{text}',
                textfont=dict(size=9, color='black'),
                xgap=2,
                ygap=2
            )

            # LABEL_PX: at -45� with wrapped labels, project max lines * line_height * sin(45�)
            LINE_H_PX = 12  # approx line height in px at 9pt
            max_lines = max((len(t.split('<br>')) for t in kpi_tick_names), default=1)
            longest_line = max((len(line.replace('<b>','').replace('</b>','')) for t in kpi_tick_names for line in t.split('<br>')), default=10)
            LABEL_PX = max(60, int((longest_line * CHAR_PX + max_lines * LINE_H_PX) * 0.71) + 10)
            BAND_PX  = 30
            GAP_PX   = 20  # extra space between tick labels and group bands
            dynamic_top = LABEL_PX + GAP_PX + BAND_PX
            row_px = 22  # compact rows to fit on screen
            col_px = 80  # wide enough so diagonal labels don't override each other
            LEFT_M = 300
            RIGHT_M = 20
            chart_height = dynamic_top + 40 + len(all_programs) * row_px
            chart_width  = LEFT_M + RIGHT_M + len(kpi_names) * col_px

            # Group bands sit just above the tick labels in paper coords
            band_y0 = 1.0 + (LABEL_PX + GAP_PX) / chart_height + 0.01
            band_y1 = band_y0 + BAND_PX / chart_height
            band_label_y = (band_y0 + band_y1) / 2

            # Move x-axis to top with wrapped tick labels
            fig.update_layout(
                height=chart_height,
                width=chart_width,
                xaxis=dict(
                    side='top',
                    tickangle=-45,
                    title='',
                    tickmode='array',
                    tickvals=kpi_names,
                    ticktext=kpi_tick_names,
                    tickfont=dict(size=9),
                    automargin=False,
                    showgrid=False,
                    zeroline=False,
                    showline=False,
                ),
                yaxis=dict(
                    tickfont=dict(size=9),
                    tickmode='array',
                    tickvals=all_programs,
                    ticktext=y_labels_wrapped,
                    automargin=True,
                    showgrid=False,
                    zeroline=False,
                    showline=False,
                ),
                yaxis_title="",
                title="",
                margin=dict(l=LEFT_M, r=RIGHT_M, t=dynamic_top, b=20),
                coloraxis_showscale=False
            )

            # Add KPI type group header rectangles + labels above the x-axis
            # Positions are dynamic (band_y0/y1/label_y) to avoid overlapping tick labels
            top_group_colors = [
                '#007a17', '#00891a', '#005c11', '#006b14',
                '#004f0e', '#008a1a', '#003d0b', '#009e1e'
            ]
            left_group_colors = [
                '#005a8e', '#004470', '#2471a3', '#1a5f8a',
                '#003d6b', '#1a6fa8', '#002e52', '#0d5496'
            ]
            for idx, (group_name, start_idx, end_idx) in enumerate(kpi_type_spans):
                color = top_group_colors[idx % len(top_group_colors)]
                x_center = (start_idx + end_idx) / 2
                fig.add_shape(
                    type='rect',
                    xref='x', yref='paper',
                    x0=start_idx - 0.5, x1=end_idx + 0.5,
                    y0=band_y0, y1=band_y1,
                    fillcolor=color,
                    line=dict(color='white', width=1),
                    layer='above'
                )
                fig.add_annotation(
                    xref='x', yref='paper',
                    x=x_center, y=band_label_y,
                    text=f"<b>{group_name}</b>",
                    showarrow=False,
                    font=dict(color='white', size=10),
                    align='center',
                    bgcolor='rgba(0,0,0,0)'
                )

            # Add program group bands to the LEFT of the y-axis (mirrors top KPI-type bands)
            for idx, (group_name, start_idx, end_idx) in enumerate(program_group_spans):
                # Skip bands that are entirely beyond the main heatmap rows or have no group name
                if start_idx >= len(programs) or not group_name:
                    continue
                end_idx = min(end_idx, len(programs) - 1)
                color = left_group_colors[idx % len(left_group_colors)]
                y_center = (start_idx + end_idx) / 2
                fig.add_shape(
                    type='rect',
                    xref='paper', yref='y',
                    x0=-0.32, x1=-0.20,
                    y0=start_idx - 0.5, y1=end_idx + 0.5,
                    fillcolor=color,
                    line=dict(color='white', width=1),
                    layer='above'
                )
                fig.add_annotation(
                    xref='paper', yref='y',
                    x=-0.26, y=y_center,
                    text=f"<b>{group_name}</b>",
                    showarrow=False,
                    font=dict(color='white', size=10),
                    align='center',
                    textangle=-90,
                    bgcolor='rgba(0,0,0,0)'
                )

            # Build raw verification DataFrame (what was read from Excel)
            raw_display_rows = []
            for r, prog in enumerate(all_programs):
                row_dict = {'Program': prog}
                for col_i, kpi in enumerate(kpi_names):
                    if r < len(text_display) and col_i < len(text_display[r]):
                        row_dict[kpi] = text_display[r][col_i]
                    else:
                        row_dict[kpi] = 'NA'
                raw_display_rows.append(row_dict)
            df_raw = pd.DataFrame(raw_display_rows).set_index('Program') if raw_display_rows else None

            return fig, df_below, df_raw
        else:
            st.error(f"No data found. Programs: {len(programs)}, KPIs: {len(kpi_names)}")
            return None, None, None
    except Exception as e:
        st.error(f"Error creating heatmap: {str(e)}")
        return None, None, None

# Helper: render a dataframe as a gray-styled HTML table
def render_gray_table(df):
    header_cells = "".join(
        f'<th style="background-color:#6b7280;color:white;padding:8px 12px;border:1px solid #9ca3af;font-weight:bold;">{col}</th>'
        for col in df.columns
    )
    rows_html = ""
    for i, row in df.iterrows():
        bg = "#f3f4f6" if i % 2 == 0 else "#e5e7eb"
        cells = "".join(
            f'<td style="background-color:{bg};padding:7px 12px;border:1px solid #d1d5db;">{val}</td>'
            for val in row
        )
        rows_html += f"<tr>{cells}</tr>"
    html = f"""
    <div style="overflow-x:auto;">
    <table style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:13px;">
      <thead><tr>{header_cells}</tr></thead>
      <tbody>{rows_html}</tbody>
    </table>
    </div>"""
    st.markdown(html, unsafe_allow_html=True)

# Load data
df_programs, df_services, df_heatmap = load_kpi_data()

# Tabs
tab1, tab2, tab3 = st.tabs(["📊 Program Output KPIs", "🏢 Service Unit KPIs", "🌡️ KPI By Program"])

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
    st.subheader("🌡️ KPI By Program")
    
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
                    fig, df_below, df_raw = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.caption(f"📄 Source: {os.path.basename(heatmap_file)} ({heatmap_file})")
                        st.plotly_chart(fig, use_container_width=False, config={'scrollZoom': False})
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        render_gray_table(df_below)
                    if df_raw is not None:
                        with st.expander("🔍 Raw Excel Data (verification)"):
                            st.dataframe(df_raw, use_container_width=True)
                else:
                    st.info("📁 Waiting for: Heat map 1.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rtpd_tabs[1]:
            st.write("**Research, Training, Product Development - KPI by FTE**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 2.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below, df_raw = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.caption(f"📄 Source: {os.path.basename(heatmap_file)} ({heatmap_file})")
                        st.plotly_chart(fig, use_container_width=False, config={'scrollZoom': False})
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        render_gray_table(df_below)
                    if df_raw is not None:
                        with st.expander("🔍 Raw Excel Data (verification)"):
                            st.dataframe(df_raw, use_container_width=True)
                else:
                    st.info("📁 Waiting for: Heat map 2.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rtpd_tabs[2]:
            st.write("**Research, Training, Product Development - KPI by $**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 3.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below, df_raw = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.caption(f"📄 Source: {os.path.basename(heatmap_file)} ({heatmap_file})")
                        st.plotly_chart(fig, use_container_width=False, config={'scrollZoom': False})
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        render_gray_table(df_below)
                    if df_raw is not None:
                        with st.expander("🔍 Raw Excel Data (verification)"):
                            st.dataframe(df_raw, use_container_width=True)
                else:
                    st.info("📁 Waiting for: Heat map 3.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rtpd_tabs[3]:
            st.write("**Research, Training, Product Development - KPI over Time**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 4.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below, df_raw = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.caption(f"📄 Source: {os.path.basename(heatmap_file)} ({heatmap_file})")
                        st.plotly_chart(fig, use_container_width=False, config={'scrollZoom': False})
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        render_gray_table(df_below)
                    if df_raw is not None:
                        with st.expander("🔍 Raw Excel Data (verification)"):
                            st.dataframe(df_raw, use_container_width=True)
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
                    fig, df_below, df_raw = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.caption(f"📄 Source: {os.path.basename(heatmap_file)} ({heatmap_file})")
                        st.plotly_chart(fig, use_container_width=False, config={'scrollZoom': False})
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        render_gray_table(df_below)
                    if df_raw is not None:
                        with st.expander("🔍 Raw Excel Data (verification)"):
                            st.dataframe(df_raw, use_container_width=True)
                else:
                    st.info("📁 Waiting for: Heat map 5.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rsi_tabs[1]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by FTE**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 6.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below, df_raw = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.caption(f"📄 Source: {os.path.basename(heatmap_file)} ({heatmap_file})")
                        st.plotly_chart(fig, use_container_width=False, config={'scrollZoom': False})
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        render_gray_table(df_below)
                    if df_raw is not None:
                        with st.expander("🔍 Raw Excel Data (verification)"):
                            st.dataframe(df_raw, use_container_width=True)
                else:
                    st.info("📁 Waiting for: Heat map 6.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rsi_tabs[2]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by $**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 7.xlsx')
                if os.path.exists(heatmap_file):
                    fig, df_below, df_raw = create_heatmap_visualization(heatmap_file)
                    if fig:
                        st.caption(f"📄 Source: {os.path.basename(heatmap_file)} ({heatmap_file})")
                        st.plotly_chart(fig, use_container_width=False, config={'scrollZoom': False})
                    if df_below is not None and not df_below.empty:
                        st.markdown("---")
                        st.markdown("**Additional Data**")
                        render_gray_table(df_below)
                    if df_raw is not None:
                        with st.expander("🔍 Raw Excel Data (verification)"):
                            st.dataframe(df_raw, use_container_width=True)
                else:
                    st.info("📁 Waiting for: Heat map 7.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")

        with rsi_tabs[3]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI over Time**")

st.markdown("---")
st.caption("Last updated: April 8, 2026 | IITA KPI Dashboard")

