import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
from openpyxl.utils import get_column_letter
import os
import numpy as np
import re
import base64

# Page config
st.set_page_config(page_title="IITA Key Performance Indicator (KPI) Dashboard", layout="wide")

# Header
st.markdown("""
<div style="background-color:#00891a; padding:20px; border-radius:10px;">
    <h1 style="color:#ffffff; text-align:center; margin:0;">🌱 IITA Key Performance Indicator (KPI) Dashboard</h1>
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
def _compute_kpi_colors_from_base(base_excel_path, target_col=4, actual_col=5, data_start_row=2):
    """Compute red→yellow→green background colours for each data row from the base
    Program Output KPIs.xlsx file using the default colour scheme.
    Returns {row_idx_in_base: (bg_color, text_color)}.
    """
    try:
        wb_d = openpyxl.load_workbook(base_excel_path, data_only=True)
        wb_f = openpyxl.load_workbook(base_excel_path)
        ws_d = wb_d.active
        ws_f = wb_f.active
    except Exception:
        return {}

    def _h2r(h):
        h = h.lstrip('#')
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
    def _r2h(r, g, b):
        return '#%02X%02X%02X' % (int(r), int(g), int(b))
    def _lerp(c1, c2, f):
        return tuple(c1[i] + (c2[i] - c1[i]) * f for i in range(3))

    low_c  = _h2r('#D73027')
    mid_c  = _h2r('#FFFF00')
    high_c = _h2r('#1A7A1A')

    colors = {}
    for row_idx in range(data_start_row, ws_d.max_row + 1):
        try:
            actual_raw = ws_d.cell(row_idx, actual_col).value
            target_raw = ws_d.cell(row_idx, target_col).value
            if actual_raw is None or target_raw is None:
                colors[row_idx] = (None, None)
                continue
            if isinstance(actual_raw, str) and actual_raw.strip().upper() in ('N/A', 'NA', '#N/A'):
                colors[row_idx] = (None, None)
                continue
            actual_fmt = getattr(ws_f.cell(row_idx, actual_col), 'number_format', None)
            target_fmt = getattr(ws_f.cell(row_idx, target_col), 'number_format', None)
            scale = 100 if ((actual_fmt and '%' in str(actual_fmt)) or (target_fmt and '%' in str(target_fmt))) else 1
            _ar = float(actual_raw)
            _tr = float(target_raw)
            a = _ar if (scale == 100 and abs(_ar) > 1) else _ar * scale
            t = _tr if (scale == 100 and abs(_tr) > 1) else _tr * scale
            mid = t / 2.0 if t else None
            if t is None or t == 0:
                bg = '#D73027' if a == 0 else '#1A7A1A'; tc = 'white'
            elif a <= 0:
                bg = '#D73027'; tc = 'white'
            elif mid and a < mid:
                f = a / mid if mid > 0 else 0
                bg = _r2h(*_lerp(low_c, mid_c, f)); tc = 'black'
            elif a < t:
                f = (a - mid) / (t - mid) if mid and (t - mid) > 0 else 0
                bg = _r2h(*_lerp(mid_c, high_c, f)); tc = 'black'
            else:
                bg = '#1A7A1A'; tc = 'white'
            colors[row_idx] = (bg, tc)
        except Exception:
            colors[row_idx] = (None, None)
    return colors


def excel_to_html_with_merged_cells(excel_file_path, no_decimals=False, highlight_row_keyword=None, target_col=4, actual_col=5, only_color_if_target=False, skip_col_indices=None, single_decimal_col_indices=None, alt_color_scheme=False, yellow_green_rows=None, yellow_to_green_rows=None, row_color_overrides=None):
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
    
    # Get merged cell ranges from format workbook (normalize to 'A1' coords)
    merged_cells = {}
    for merged_range in ws_format.merged_cells.ranges:
        cells = list(merged_range.cells)
        for cell in cells[1:]:  # Skip the first cell (top-left)
            merged_cells[cell] = cells[0]  # Map to top-left cell

    # Determine if we should suppress header coloring for specific KPI files
    suppress_header_color = False
    is_service_unit_file = False
    is_program_file = False
    is_program_variant_file = False
    try:
        bn = os.path.basename(excel_file_path).lower()
        if 'program output' in bn or 'service unit' in bn:
            suppress_header_color = True
        if 'service unit' in bn:
            is_service_unit_file = True
        if 'program output' in bn:
            is_program_file = True
            # FTE and USD variants have a secondary header in row 2 that should be gray
            if 'fte' in bn or ' $' in bn or 'by $' in bn or 'by fte' in bn:
                is_program_variant_file = True
        is_fte_file = 'program output' in bn and 'fte' in bn
    except Exception:
        suppress_header_color = False
        is_service_unit_file = False
        is_program_file = False
        is_program_variant_file = False
        is_fte_file = False
    # For service unit files, also suppress coloring for row 2 (secondary header)
    suppress_row2_header = is_service_unit_file
    
    # Build HTML table
    html = '<table style="border-collapse: collapse; width: 100%; table-layout: fixed; font-family: Arial, sans-serif; background-color: white;">'
    # Default styles remain left-aligned; we'll override per-cell below
    # Use a soft black border for all table cells so sections are visually separated
    html += '<style>td, th { border: 1px solid rgba(0,0,0,0.12); padding: 12px; text-align: left; background-color: white; white-space: normal; word-wrap: break-word; word-break: break-word; overflow-wrap: break-word; color: black; font-size: 9pt; }</style>'
    
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

        # Determine if this row should be highlighted (contains keyword)
        highlight_row = False
        if highlight_row_keyword:
            try:
                lowkw = highlight_row_keyword.strip().lower()
                for c in row_data:
                    if c.value is not None and lowkw in str(c.value).lower():
                        highlight_row = True
                        break
            except Exception:
                highlight_row = False

        # If this is a Service Unit file, suppress highlights triggered by the keyword
        if highlight_row and suppress_row2_header:
            highlight_row = False

        # Detect section header rows (merged across columns or bold font) for Program/Service files
        # Only check columns 1-3 for bold — year/data columns (e.g. col 7) can be bold
        # throughout in FTE/$ files and must not falsely mark every row as a section header.
        row_is_section_header = False
        if is_program_file or is_service_unit_file:
            try:
                for col_idx in range(1, min(4, max_col + 1)):
                    fmt_cell = ws_format[f"{get_column_letter(col_idx)}{row_idx}"]
                    if getattr(fmt_cell, 'font', None) and getattr(fmt_cell.font, 'bold', False):
                        row_is_section_header = True
                        break
                if not row_is_section_header:
                    for mr in ws_format.merged_cells.ranges:
                        if mr.min_row == row_idx and (mr.max_col - mr.min_col + 1) >= 2:
                            row_is_section_header = True
                            break
            except Exception:
                row_is_section_header = False

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

            # Skip columns that should not be displayed
            if skip_col_indices and col_idx in skip_col_indices:
                continue

            # Calculate rowspan and colspan for merged cells
            rowspan = 1
            colspan = 1
            
            for merged_range in ws_format.merged_cells.ranges:
                if cell_coord in merged_range:
                    rowspan = merged_range.max_row - merged_range.min_row + 1
                    raw_colspan = merged_range.max_col - merged_range.min_col + 1
                    # Subtract any skipped columns within this merged range
                    if skip_col_indices:
                        skipped_in_range = sum(1 for c in range(merged_range.min_col, merged_range.max_col + 1) if c in skip_col_indices)
                        colspan = max(1, raw_colspan - skipped_in_range)
                    else:
                        colspan = raw_colspan
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
                        try:
                            # Values already stored as whole percentages (e.g. 50 meaning 50%)
                            # should not be multiplied by 100; only decimal fractions (e.g. 0.5) need it.
                            pct = cell_value if abs(cell_value) > 1 else cell_value * 100
                            # Respect Excel percent format decimals when possible
                            fmt = str(cell_format.number_format)
                            dec = None
                            try:
                                m = re.search(r"%(?!.*%)", fmt)
                                # count zeros after decimal point before % (e.g. '0.00%')
                                md = re.search(r"\.(0+)[^%]*%", fmt)
                                if md:
                                    dec = len(md.group(1))
                                else:
                                    # if no explicit decimals, assume 0
                                    dec = 0
                            except Exception:
                                dec = None
                            if dec is None:
                                s = f"{pct:.2f}".rstrip('0').rstrip('.')
                            else:
                                s = f"{pct:.{dec}f}"
                            cell_value = f"{s}%"
                        except Exception:
                            cell_value = str(cell_value)
                    else:
                        if no_decimals:
                            # Round to nearest integer and show without decimals
                            try:
                                cell_value = f"{int(round(cell_value)):,}"
                            except Exception:
                                cell_value = str(cell_value)
                        else:
                            # Up to 3 decimal places; strip trailing zeros; add thousand commas
                            try:
                                if cell_value == int(cell_value):
                                    cell_value = f"{int(cell_value):,}"
                                else:
                                    cell_value = f"{cell_value:,.3f}".rstrip('0').rstrip('.')
                            except Exception:
                                cell_value = str(cell_value)
                else:
                    cell_value = str(cell_value)
            else:
                cell_value = ""

            # Override to 1 decimal place for specified columns (skip header row 1)
            # Do NOT override if the cell has a percentage format — it's already been formatted as "50%"
            _is_pct_cell = bool(cell_format.number_format and '%' in str(cell_format.number_format))
            if single_decimal_col_indices and col_idx in single_decimal_col_indices and row_idx != 1 and not _is_pct_cell:
                if isinstance(cell_data.value, (int, float)):
                    try:
                        if row_idx == 2:
                            # Display row 2 totals with standard thousands separators, no decimals
                            cell_value = f"{int(round(cell_data.value)):,}"
                        else:
                            cell_value = f"{cell_data.value:,.1f}"
                    except Exception:
                        pass

            # Determine if original cell was numeric so we can align numbers/columns
            is_numeric = isinstance(cell_data.value, (int, float))
            # For Service Unit tables: left-align columns 1 and 2, right-align
            # columns 3 and 4, center other columns. For other tables keep existing rules.
            if is_service_unit_file:
                if col_idx in (1, 2):
                    align = 'left'
                elif col_idx in (3, 4):
                    align = 'right'
                else:
                    align = 'center'
            elif is_program_file:
                # For Program Output files ensure first two columns are left-aligned
                # Force column 4 to be right-aligned, column 5 always right, and keep numeric cells right-aligned
                if col_idx in (1, 2):
                    align = 'left'
                elif col_idx in (4, 5):
                    align = 'right'
                elif is_numeric:
                    align = 'right'
                else:
                    align = 'left'
            else:
                # Column-based rule: first three columns should be left-aligned for other files
                if col_idx <= 3:
                    align = 'left'
                elif is_numeric:
                    align = 'right'
                else:
                    align = 'left'

            # Add styling for headers (first row)
            if row_idx == 1:
                # For green header cells: if original value is numeric and >=1000, remove thousands separators
                header_display = cell_value
                try:
                    if isinstance(cell_data.value, (int, float)) and abs(cell_data.value) >= 1000:
                        header_display = str(header_display).replace(',', '')
                except Exception:
                    pass
                # make column header font slightly larger
                # For Service Unit files, use a light grey header; for Program files also use light grey
                # but left-align the first two header cells; for other suppressed files use white; otherwise green
                if is_service_unit_file:
                    # set first and second column widths for service unit tables (col1 a bit narrower, col2 a bit wider)
                    if col_idx == 1:
                        width_style = ' width: 32%;'
                    elif col_idx == 2:
                        width_style = ' width: 20%;'
                    else:
                        width_style = ''
                    # add a stronger bottom border for the top header row in Service Unit tables
                    html += f'<th style="background-color: #e0e0e0; color: black; font-weight: bold; text-align: center; font-size: 11pt; font-family: Arial, sans-serif;{width_style} border-bottom: 3px solid #000;" rowspan="{rowspan}" colspan="{colspan}">{header_display}</th>'
                elif is_program_file:
                    # Program Output: use light-gray header and left-align first two columns
                    # Col 2/3 widths: narrower col 2 and wider col 3 for the base file only;
                    # FTE/USD variants keep their original widths.
                    if col_idx == 1:
                        width_style = ' width: 14%;' if is_fte_file else ' width: 14%;'
                    elif col_idx == 2:
                        width_style = ' width: 7%;' if not is_program_variant_file else ' width: 15%;'
                    elif col_idx == 3:
                        width_style = ' width: 38%;' if not is_program_variant_file else ' width: 32%;'
                    elif col_idx == 4:
                        width_style = ' width: 10%;' if not is_program_variant_file else ''
                    else:
                        width_style = ''
                    text_align = 'left' if col_idx in (1, 2) else 'center'
                    html += f'<th style="background-color: #e0e0e0; color: black; font-weight: bold; text-align: {text_align}; font-size: 11pt; font-family: Arial, sans-serif;{width_style} border-bottom: 3px solid #000;" rowspan="{rowspan}" colspan="{colspan}">{header_display}</th>'
                elif suppress_header_color:
                    html += f'<th style="background-color: white; color: black; font-weight: bold; text-align: center; font-size: 11pt; font-family: Arial, sans-serif;" rowspan="{rowspan}" colspan="{colspan}">{header_display}</th>'
                else:
                    html += f'<th style="background-color: #00891a; color: white; font-weight: bold; text-align: center; font-size: 11pt; font-family: Arial, sans-serif;" rowspan="{rowspan}" colspan="{colspan}">{header_display}</th>'
            else:
                # Build inline style for this cell
                styles = []
                if highlight_row:
                    styles.append('background-color: #00891a')
                    styles.append('font-weight: bold')
                # Center-align the top header row, the Service Unit header row (row 9),
                # and row 2 of Program Output FTE/USD variant files
                cell_align = 'center' if (row_idx == 1 or (is_service_unit_file and row_idx == 9) or (is_program_variant_file and row_idx == 2)) else align
                styles.append(f'text-align: {cell_align}')
                # For Service Unit and Program Output tables add a slightly thicker
                # gray bottom separator for data rows (keeps header/band rows intact).
                if is_service_unit_file or is_program_file:
                    try:
                        # For Service Unit files we want to suppress the special
                        # Service Unit header row (row 9). For Program files this
                        # will be False. Also suppress for program variant row 2.
                        is_srv_header_row = (
                            (suppress_row2_header and (row_idx == 9 or any(c.value is not None and 'service unit key performance' in str(c.value).lower() for c in row_data)))
                            or (is_program_variant_file and row_idx == 2)
                        )
                    except Exception:
                        is_srv_header_row = False
                    # Don't add the gray bottom border for top header (row 1)
                    # or for detected section/header rows.
                    if not is_srv_header_row and row_idx != 1 and not row_is_section_header:
                        styles.append('border-bottom: 2px solid rgba(0,0,0,0.25)')
                # Add section separator for Program and Service Unit tables
                # Do not add a top border before the Service Unit header row (row 9)
                # or before the Program variant header row (row 2);
                # we'll add the stronger border below those rows instead.
                if row_is_section_header and (is_program_file or is_service_unit_file) and row_idx != 1 and not (is_service_unit_file and row_idx == 9) and not (is_program_variant_file and row_idx == 2):
                    styles.append('border-top: 2px solid #000')
                # If this is a Service Unit file and the Service Unit header row (row 9),
                # or a Program Output FTE/USD file and row 2,
                # force no green header and unify font size so the row matches visually
                if (suppress_row2_header and (row_idx == 9 or
                                             any(c.value is not None and 'service unit key performance' in str(c.value).lower() for c in row_data))) \
                        or (is_program_variant_file and row_idx == 2):
                    # remove any green highlight and use light-gray background with black text (Service Unit header rows)
                    styles = [s for s in styles if 'background-color' not in s and 'color:' not in s]
                    styles.append('background-color: #e0e0e0')
                    styles.append('color: black')
                    if not (is_program_variant_file and row_idx == 2):
                        styles.append('font-weight: bold')
                    # ensure the font size for these service-unit header rows matches the main header
                    styles.append('font-size: 11pt')
                    styles.append('font-family: Arial, sans-serif')
                    # add a strong bottom border to separate this header row from the content below
                    styles.append('border-bottom: 3px solid #000')
                # Color coding for Actual column only (default Excel column 5)
                bg_color = None
                text_color = None
                if col_idx == actual_col:
                    try:
                        actual_raw = cell_data.value
                        # Target is in column target_col (1-based) → index target_col-1 in row_data (0-based)
                        tgt_idx = target_col - 1
                        target_cell_obj = row_data[tgt_idx] if len(row_data) > tgt_idx else None
                        target_raw = None
                        target_fmt = None
                        # Resolve merged-anchor for the target cell (scan merged ranges)
                        try:
                            if target_cell_obj is not None:
                                t_row = getattr(target_cell_obj, 'row', None)
                                t_col = getattr(target_cell_obj, 'column', None)
                                anchor_coord = None
                                try:
                                    for mr in ws_format.merged_cells.ranges:
                                        if t_row is not None and t_col is not None and mr.min_row <= t_row <= mr.max_row and mr.min_col <= t_col <= mr.max_col:
                                            anchor_coord = f"{get_column_letter(mr.min_col)}{mr.min_row}"
                                            break
                                except Exception:
                                    anchor_coord = None
                                if anchor_coord is None:
                                    anchor_coord = target_cell_obj.coordinate
                                try:
                                    target_raw = ws_data[anchor_coord].value
                                except Exception:
                                    target_raw = target_cell_obj.value
                                try:
                                    target_fmt = getattr(ws_format[anchor_coord], 'number_format', None)
                                except Exception:
                                    target_fmt = None
                        except Exception:
                            target_raw = target_cell_obj.value if target_cell_obj is not None else None
                            try:
                                if target_cell_obj is not None:
                                    target_fmt = getattr(ws_format[target_cell_obj.coordinate], 'number_format', None)
                            except Exception:
                                target_fmt = None
                        # Detect percent formats on either cell and set scale
                        actual_fmt = getattr(cell_format, 'number_format', None)
                        if (actual_fmt and '%' in str(actual_fmt)) or (target_fmt and '%' in str(target_fmt)):
                            scale = 100
                        else:
                            scale = 1
                        if actual_raw is not None and target_raw is not None:
                            # Skip the *100 scale for values already stored as whole percentages
                            _ar = float(actual_raw)
                            _tr = float(target_raw)
                            a = _ar if (scale == 100 and abs(_ar) > 1) else _ar * scale
                            t = _tr if (scale == 100 and abs(_tr) > 1) else _tr * scale
                            mid = t / 2.0 if t is not None else None
                            # helper: interpolate between two hex colors
                            def hex_to_rgb(h):
                                h = h.lstrip('#')
                                return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
                            def rgb_to_hex(r, g, b):
                                return '#%02X%02X%02X' % (int(r), int(g), int(b))
                            def lerp(c1, c2, f):
                                return tuple(c1[i] + (c2[i] - c1[i]) * f for i in range(3))

                            if is_program_variant_file and row_idx == 11:
                                # Row 11 in FTE/USD files: two-step green only
                                #   >= target       → dark green
                                #   >= target / 2   → light green
                                #   <  target / 2   → no color
                                half = t / 2.0 if (t is not None and t != 0) else None
                                if t is None or t == 0:
                                    bg_color = None; text_color = None
                                elif a >= t:
                                    bg_color = '#1A7A1A'; text_color = 'white'
                                elif half is not None and a >= half:
                                    bg_color = '#A9D18E'; text_color = 'black'
                                else:
                                    bg_color = None; text_color = None
                            elif yellow_to_green_rows and row_idx in yellow_to_green_rows:
                                # Yellow(<=0) -> Lime green(>=target) gradient
                                yellow_c    = hex_to_rgb('#FFFF00')
                                limegreen_c = hex_to_rgb('#92D050')  # lime green at/above target
                                if t is None or t == 0:
                                    bg_color = '#92D050'; text_color = 'black'
                                else:
                                    if a >= t:
                                        bg_color = '#92D050'; text_color = 'black'
                                    elif a <= 0:
                                        bg_color = '#FFFF00'; text_color = 'black'
                                    else:
                                        # Yellow -> Lime green gradient
                                        f = a / t
                                        rgb = lerp(yellow_c, limegreen_c, f)
                                        bg_color = rgb_to_hex(*rgb)
                                        text_color = 'black'
                            elif yellow_green_rows and row_idx in yellow_green_rows:
                                # Special: Yellow(0) -> Orange(target*0.75) -> Green(>=target)
                                yellow_c = hex_to_rgb('#FFFF00')
                                orange_c = hex_to_rgb('#FFA500')
                                green_c  = hex_to_rgb('#1A7A1A')
                                if t is None or t == 0:
                                    bg_color = '#1A7A1A'; text_color = 'white'
                                else:
                                    orange_pt = t * 0.75
                                    if a >= t:
                                        bg_color = '#1A7A1A'; text_color = 'white'
                                    elif a <= 0:
                                        bg_color = '#FFFF00'; text_color = 'black'
                                    elif a < orange_pt:
                                        # Yellow -> Orange
                                        f = a / orange_pt if orange_pt > 0 else 0
                                        rgb = lerp(yellow_c, orange_c, f)
                                        bg_color = rgb_to_hex(*rgb)
                                        text_color = 'black'
                                    else:
                                        # Orange -> Green
                                        f = (a - orange_pt) / (t - orange_pt) if (t - orange_pt) > 0 else 1
                                        rgb = lerp(orange_c, green_c, f)
                                        bg_color = rgb_to_hex(*rgb)
                                        text_color = 'black' if f < 0.7 else 'white'
                            elif alt_color_scheme:
                                # Red(0) -> Yellow(target/2) -> Green(>=target)
                                red_c    = hex_to_rgb('#D73027')
                                yellow_c = hex_to_rgb('#FFFF00')
                                green_c  = hex_to_rgb('#1A7A1A')
                                half = t / 2.0 if (t is not None and t != 0) else None

                                if t is None or t == 0:
                                    if a <= 0:
                                        bg_color = '#D73027'; text_color = 'white'
                                    else:
                                        bg_color = '#1A7A1A'; text_color = 'white'
                                else:
                                    if a <= 0:
                                        bg_color = '#D73027'; text_color = 'white'
                                    elif a >= t:
                                        bg_color = '#1A7A1A'; text_color = 'white'
                                    elif a < half:
                                        # Red -> Yellow
                                        f = a / half if half > 0 else 0
                                        rgb = lerp(red_c, yellow_c, f)
                                        bg_color = rgb_to_hex(*rgb)
                                        text_color = 'black'
                                    else:
                                        # Yellow -> Green
                                        f = (a - half) / (t - half) if (t - half) > 0 else 1
                                        rgb = lerp(yellow_c, green_c, f)
                                        bg_color = rgb_to_hex(*rgb)
                                        text_color = 'black' if f < 0.6 else 'white'
                            else:
                                low   = hex_to_rgb('#D73027')  # red
                                mid_c = hex_to_rgb('#FFFF00')  # yellow
                                high  = hex_to_rgb('#1A7A1A')  # dark green

                                if t is None or t == 0:
                                    if a == 0:
                                        bg_color = '#D73027'; text_color = 'white'
                                    else:
                                        bg_color = '#1A7A1A'; text_color = 'white'
                                else:
                                    if a <= 0:
                                        bg_color = '#D73027'; text_color = 'white'
                                    elif a < mid:
                                        f = (a) / (mid) if mid > 0 else 0
                                        rgb = lerp(low, mid_c, f)
                                        bg_color = rgb_to_hex(*rgb)
                                        text_color = 'black'
                                    elif a < t:
                                        f = (a - mid) / (t - mid) if (t - mid) > 0 else 0
                                        rgb = lerp(mid_c, high, f)
                                        bg_color = rgb_to_hex(*rgb)
                                        text_color = 'black'
                                    else:
                                        bg_color = '#1A7A1A'; text_color = 'white'
                        else:
                            # Fallback: if target missing, only color if flag allows it
                            if not only_color_if_target and actual_raw is not None:
                                a = float(actual_raw)
                                if a == 0:
                                    bg_color = '#D73027'; text_color = 'white'
                                else:
                                    bg_color = '#1A7A1A'; text_color = 'white'
                    except Exception:
                        bg_color = None
                        text_color = None
                    # Apply pre-computed colour override AFTER calculation (overrides target-based result)
                    if row_color_overrides is not None and row_idx in row_color_overrides:
                        _ov_raw = cell_data.value
                        _ov_na = (_ov_raw is None or
                                  (isinstance(_ov_raw, str) and _ov_raw.strip().upper() in ('N/A', 'NA', '#N/A')))
                        if not _ov_na and row_color_overrides[row_idx][0] is not None:
                            bg_color, text_color = row_color_overrides[row_idx]
                        else:
                            bg_color = None; text_color = None

                # First column cells (row headers) should have slightly larger font
                if col_idx == 1:
                    styles.append('font-size: 11pt')
                    # For Service Unit and Program Output tables, set appropriate widths
                    if is_service_unit_file:
                        styles.append('width: 24%')
                    elif is_program_file:
                        styles.append('width: 16%' if is_fte_file else 'width: 14%')
                elif col_idx == 2:
                    # Make column 2 slightly wider for Program Output and Service Unit tables
                    if is_service_unit_file:
                        styles.append('width: 26%')
                    elif is_program_file:
                        # Narrower col 2 for base file only; variants keep original width
                        styles.append('width: 7%' if not is_program_variant_file else 'width: 24%')
                elif col_idx == 3:
                    # Make column 3 wider for Program Output tables
                    if is_program_file:
                        # Wider col 3 for base file only; variants keep original width
                        styles.append('width: 38%' if not is_program_variant_file else 'width: 40%')
                elif col_idx == 4:
                    # Make column 4 wider for the base Program Output file
                    if is_program_file and not is_program_variant_file:
                        styles.append('width: 10%')
                if bg_color:
                    styles.append(f'background-color: {bg_color}')
                # We will force data text color to black for consistency (append below)
                # Force data cells to black text (td). Header <th> handled separately.
                styles = [s for s in styles if not s.strip().startswith('color:')]
                styles.append('color: black')
                style_attr = '; '.join(styles)
                # If this is the Service Unit header row (row 9), contains the phrase,
                # or is row 2 of a Program Output FTE/USD variant file,
                # ensure bold display and consistent font sizing.
                if suppress_row2_header and (row_idx == 9 or any(c.value is not None and 'service unit key performance' in str(c.value).lower() for c in row_data)):
                    # ensure style includes bold
                    if 'font-weight' not in style_attr:
                        style_attr = (style_attr + '; font-weight: bold').strip()
                    # ensure style includes the intended font size for consistency
                    if 'font-size' not in style_attr:
                        style_attr = (style_attr + '; font-size: 11pt').strip()
                    # ensure style includes the intended font family for consistency
                    if 'font-family' not in style_attr:
                        style_attr = (style_attr + '; font-family: Arial, sans-serif').strip()
                    cell_display = f'<strong>{cell_value}</strong>'
                elif is_program_variant_file and row_idx == 2:
                    # row 2 of FTE/USD variant files: styled but not bold
                    if 'font-size' not in style_attr:
                        style_attr = (style_attr + '; font-size: 11pt').strip()
                    if 'font-family' not in style_attr:
                        style_attr = (style_attr + '; font-family: Arial, sans-serif').strip()
                    cell_display = cell_value
                else:
                    cell_display = cell_value
                html += f'<td style="{style_attr};" rowspan="{rowspan}" colspan="{colspan}">{cell_display}</td>'
        
        html += '</tr>'
    
    html += '</table>'
    return html


# Render helpers for Program KPI subtabs — keep logic separate per-tab for future customizations
def render_program_kpi_number(excel_path, df=None):
    st.markdown('<h3 style="font-family: Arial, sans-serif; font-size:16px; margin:4px 0;">Program KPI by Number</h3>', unsafe_allow_html=True)
    try:
        html_programs = excel_to_html_with_merged_cells(excel_path, no_decimals=False)
        st.markdown(get_heatmap_legend_html() + html_programs, unsafe_allow_html=True)
    except Exception as e:
        st.warning(f"Could not render with merged cells: {str(e)}")
        if df is not None:
            display_df = df.copy()
            for col in display_df.select_dtypes(include=["number"]).columns:
                def fmt_cell(x):
                    if pd.isna(x):
                        return ""
                    try:
                        if isinstance(x, (int, float)) and 0 <= x <= 1:
                            s = f"{x * 100:.2f}".rstrip('0').rstrip('.')
                            return s + '%'
                        else:
                            return str(int(round(x)))
                    except Exception:
                        return str(x)
                display_df[col] = display_df[col].apply(fmt_cell)
            st.dataframe(display_df, width='stretch', height=600)


def render_program_kpi_fte_with_color_coding(excel_path):
    st.markdown('<h3 style="font-family: Arial, sans-serif; font-size:16px; margin:4px 0;">Program KPI by Full Time Equivalent (FTE)</h3>', unsafe_allow_html=True)
    legend = get_heatmap_legend_html()
    try:
        if os.path.exists(excel_path):
            # Derive colours from the base Program Output KPIs.xlsx (same folder)
            # Row mapping: FTE row N  →  base row N-1  (FTE has extra gray row 2)
            _base_path = os.path.join(os.path.dirname(excel_path), 'Program Output KPIs.xlsx')
            _base_colors = _compute_kpi_colors_from_base(_base_path) if os.path.exists(_base_path) else {}
            _row_overrides = {1: (None, None), 2: (None, None)}
            _row_overrides.update({fte_row: _base_colors.get(fte_row - 1, (None, None))
                              for fte_row in range(3, 30)})
            # skip_col_indices=[4]: hide Notional Target column from display
            # single_decimal_col_indices=[5]: show 2025 actuals to 1 decimal place
            html = excel_to_html_with_merged_cells(
                excel_path, no_decimals=False, actual_col=5,
                skip_col_indices=[4], single_decimal_col_indices=[5],
                row_color_overrides=_row_overrides
            )
            st.markdown(legend + html, unsafe_allow_html=True)
            try:
                df = pd.read_excel(excel_path)
                st.download_button(label="⬇️ Download 2025 Program KPIs (FTE) as CSV", data=df.to_csv(index=False), file_name="2025_Program_Output_KPIs_FTE.csv", mime="text/csv")
            except Exception:
                pass
        else:
            st.info('📁 Waiting for: Program Output KPIs by FTE.xlsx')
    except Exception as e:
        st.warning(f"Could not render FTE sheet: {str(e)}")
        try:
            df = pd.read_excel(excel_path)
            display_df = df.copy()
            for col in display_df.select_dtypes(include=["number"]).columns:
                display_df[col] = display_df[col].apply(lambda x: "" if pd.isna(x) else str(int(round(x))))
            st.dataframe(display_df, width='stretch', height=600)
        except Exception:
            st.info('No FTE-specific sheet found or it could not be read.')


def render_program_kpi_usd(excel_path):
    st.markdown('<h3 style="font-family: Arial, sans-serif; font-size:16px; margin:4px 0;">Program KPI by Million (USD)</h3>', unsafe_allow_html=True)
    legend = get_heatmap_legend_html()
    try:
        if os.path.exists(excel_path):
            # Derive colours from the base Program Output KPIs.xlsx (same folder)
            # Row mapping: $ row N  →  base row N-1  ($ file has extra gray row 2)
            _base_path = os.path.join(os.path.dirname(excel_path), 'Program Output KPIs.xlsx')
            _base_colors = _compute_kpi_colors_from_base(_base_path) if os.path.exists(_base_path) else {}
            _row_overrides = {1: (None, None), 2: (None, None)}
            _row_overrides.update({usd_row: _base_colors.get(usd_row - 1, (None, None))
                              for usd_row in range(3, 30)})
            # skip_col_indices=[4]: hide Notional Target column from display
            # single_decimal_col_indices=[5]: show 2025 actuals to 1 decimal place
            html = excel_to_html_with_merged_cells(
                excel_path, no_decimals=False, actual_col=5,
                skip_col_indices=[4], single_decimal_col_indices=[5],
                row_color_overrides=_row_overrides
            )
            st.markdown(legend + html, unsafe_allow_html=True)
            try:
                df = pd.read_excel(excel_path)
                st.download_button(label="⬇️ Download 2025 Program KPIs (USD) as CSV", data=df.to_csv(index=False), file_name="2025_Program_Output_KPIs_USD.csv", mime="text/csv")
            except Exception:
                pass
        else:
            st.info('📁 Waiting for: Program Output KPIs by $.xlsx')
    except Exception as e:
        st.warning(f"Could not render USD sheet: {str(e)}")
        try:
            df = pd.read_excel(excel_path)
            display_df = df.copy()
            for col in display_df.select_dtypes(include=["number"]).columns:
                display_df[col] = display_df[col].apply(lambda x: "" if pd.isna(x) else str(int(round(x))))
            st.dataframe(display_df, width='stretch', height=600)
        except Exception:
            st.info('No USD-specific sheet found or it could not be read.')

# Function to create heatmap from KPI heat map file
# Returns (fig, df_below) where df_below contains rows beyond row 16 (or None)
def create_heatmap_visualization(excel_file_path, heatmap_max_row=16,
                                  data_col_start=4, program_col=3, group_col=2,
                                  kpi_row=3, kpi_group_row=2, data_row_start=4,
                                  show_row_groups=True, include_below_rows=True,
                                  left_margin=None, side_cols=None,
                                  zero_decimal_cols=None, one_decimal_cols=None, two_decimal_cols=None,
                                  zero_decimal_rows=None, one_decimal_rows=None, force_decimals=None, suppress_pct_display=False, monospace_numeric=False,
                                  one_decimal_first_col=False, no_gray_first_col=False,
                                  kpi_group_filter=None, force_include_cols=None, kpi_group_source=None, extra_top=0, group_gap=None):
    try:
        # Load with openpyxl to get clean numeric data
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        ws = wb.active

        # Build a map of merged cells -> top-left cell coordinates so we can
        # read values/number formats from the merged-region anchor when a
        # cell belongs to a merged range.
        merged_map = {}
        try:
            for mr in ws.merged_cells.ranges:
                min_r, min_c = mr.min_row, mr.min_col
                for rr in range(mr.min_row, mr.max_row + 1):
                    for cc in range(mr.min_col, mr.max_col + 1):
                        merged_map[(rr, cc)] = (min_r, min_c)
        except Exception:
            merged_map = {}

        def merged_cell_coord(r, c):
            return merged_map.get((r, c), (r, c))

        def merged_val(r, c):
            tr, tc = merged_cell_coord(r, c)
            return ws.cell(row=tr, column=tc).value

        def merged_cell_obj(r, c):
            tr, tc = merged_cell_coord(r, c)
            return ws.cell(row=tr, column=tc)

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

        # Build program-group map from group_col (fill-forward for merged cells)
        current_b = None
        b_values = {}
        if group_col is not None:
            for row_idx in range(data_row_start, ws.max_row + 1):
                val = merged_val(row_idx, group_col)
                if val is not None:
                    current_b = str(val)
                b_values[row_idx] = current_b

        # Get base KPI names and columns from kpi_row (starting from data_col_start)
        # Scan to last non-None header to avoid empty trailing columns. Use
        # merged_val so headers that are part of merged cells are detected.
        base_kpi_cols = []
        base_kpi_names = []
        last_kpi_col = data_col_start
        for col_idx in range(data_col_start, ws.max_column + 1):
            cell_val = merged_val(kpi_row, col_idx) if kpi_row else None
            if cell_val is not None:
                last_kpi_col = col_idx
        for col_idx in range(data_col_start, last_kpi_col + 1):
            if (side_cols or []) and col_idx in (side_cols or []):
                continue  # already handled as a side column — do not add to base KPI list
            cell_val = merged_val(kpi_row, col_idx) if kpi_row else None
            if cell_val is not None:
                base_kpi_cols.append(col_idx)
                base_kpi_names.append(str(cell_val))

        # side_cols are absolute Excel column indices to include as uncolored side columns
        side_cols = side_cols or []

        # Build final included columns and kpi_names: side_cols first, then base KPI cols
        included_cols = []
        kpi_names = []
        for c in side_cols:
            if kpi_row:
                obj = merged_cell_obj(kpi_row, c)
                hdr = obj.value if obj and obj.value is not None else get_column_letter(c)
            else:
                hdr = get_column_letter(c)
            included_cols.append(c)
            kpi_names.append(str(hdr))
        included_cols.extend(base_kpi_cols)
        kpi_names.extend(base_kpi_names)

        # Allow caller to force include specific absolute column indices
        # (useful for wide heatmap variants where data sits in widely spaced columns)
        if force_include_cols:
            for col_idx in force_include_cols:
                if col_idx not in included_cols and col_idx <= ws.max_column:
                    hdr = merged_val(kpi_row, col_idx) if kpi_row else None
                    included_cols.append(col_idx)
                    kpi_names.append(str(hdr) if hdr is not None else get_column_letter(col_idx))

        # Ensure we include any columns that contain data in the data rows even if the
        # KPI header cell is blank (some wide sheets leave header cells empty).
        for col_idx in range(data_col_start, ws.max_column + 1):
            if col_idx in included_cols:
                continue
            has_data = False
            # Check all data rows (including rows beyond heatmap_max_row) for any
            # non-empty cell in this column so we don't omit columns that only
            # have values in the 'below' region.
            for r in range(data_row_start, ws.max_row + 1):
                if merged_val(r, col_idx) is not None:
                    has_data = True
                    break
            if has_data:
                included_cols.append(col_idx)
                hdr = merged_val(kpi_row, col_idx) if kpi_row else None
                kpi_names.append(str(hdr) if hdr is not None else get_column_letter(col_idx))

        # Get KPI type groups from kpi_group_row (fill-forward for merged cells)
        current_type = None
        for c in included_cols:
            if c in side_cols:
                kpi_type_groups.append("")
            else:
                val = merged_val(kpi_group_row, c) if kpi_group_row else None
                if val is not None:
                    current_type = str(val)
                kpi_type_groups.append(current_type or "")

        # If a separate KPI group source file was provided, prefer its grouping
        # text (useful when Capacity/Product sheets omit the top grouping row).
        if kpi_group_source and os.path.exists(kpi_group_source):
            try:
                wb_src = openpyxl.load_workbook(kpi_group_source, data_only=True)
                ws_src = wb_src.active
                # Build merged map for source
                merged_map_src = {}
                try:
                    for mr in ws_src.merged_cells.ranges:
                        min_r, min_c = mr.min_row, mr.min_col
                        for rr in range(mr.min_row, mr.max_row + 1):
                            for cc in range(mr.min_col, mr.max_col + 1):
                                merged_map_src[(rr, cc)] = (min_r, min_c)
                except Exception:
                    merged_map_src = {}

                def merged_src_val(r, c):
                    tr, tc = merged_map_src.get((r, c), (r, c))
                    return ws_src.cell(row=tr, column=tc).value

                # Build group values from the source file with fill-forward semantics
                src_current = None
                src_groups = []
                for c in included_cols:
                    if c in side_cols:
                        src_groups.append("")
                    else:
                        v = merged_src_val(kpi_group_row, c) if kpi_group_row else None
                        if v is not None:
                            src_current = str(v)
                        src_groups.append(src_current or "")

                # If the source provided any non-empty group names, replace the groups
                if any(g for g in src_groups if g):
                    kpi_type_groups = src_groups
            except Exception:
                pass

        # If a KPI group filter was specified, filter included columns and names
        # by case-insensitive substring match on the KPI type group. This allows
        # creating focused "KPI over Time" views for Research Outputs,
        # Capacity Building, Product Development, etc.
        if kpi_group_filter:
            filt = str(kpi_group_filter).strip().lower()
            new_included = []
            new_names = []
            new_type_groups = []
            for col_idx, name, grp in zip(included_cols, kpi_names, kpi_type_groups):
                if grp and filt in grp.lower():
                    new_included.append(col_idx)
                    new_names.append(name)
                    new_type_groups.append(grp)
            if new_included:
                included_cols = new_included
                kpi_names = new_names
                kpi_type_groups = new_type_groups

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

        # Make KPI names unique while preserving visible text by appending
        # zero-width-space suffixes to repeated labels. This prevents
        # dictionary-based row assembly from collapsing duplicate columns.
        seen = {}
        unique_kpi_names = []
        for name in kpi_names:
            cnt = seen.get(name, 0)
            if cnt == 0:
                unique_kpi_names.append(name)
            else:
                unique_kpi_names.append(name + '\u200b' * cnt)
            seen[name] = cnt + 1
        kpi_names = unique_kpi_names

        # Build set of column indices whose KPI header contains '%'.
        # openpyxl returns percentage-formatted cells as decimals (e.g. 0.85 for 85%);
        # multiply by 100 so values display as the user entered them in Excel.
        # A column is treated as percentage if:
        #   (a) its KPI header name contains '%', OR
        #   (b) any data cell in that column has a percentage Excel number format.
        # openpyxl stores percentage values as decimals (1.0 = 100%); we multiply by 100.
        pct_col_indices = set()
        for i, (col_idx, name) in enumerate(zip(included_cols, kpi_names)):
            if '%' in name:
                pct_col_indices.add(col_idx)
            else:
                # Scan first few data rows for a % number format on that column
                for scan_row in range(data_row_start, min(data_row_start + 5, ws.max_row + 1)):
                    cell_obj = merged_cell_obj(scan_row, col_idx)
                    if cell_obj is not None and cell_obj.value is not None and getattr(cell_obj, 'number_format', None) and '%' in cell_obj.number_format:
                        pct_col_indices.add(col_idx)
                        break

        # Get data from data_row_start to heatmap_max_row for the heatmap
        for row_idx in range(data_row_start, heatmap_max_row + 1):
            program_cell = merged_val(row_idx, program_col)
            if program_cell is not None and program_cell != "None":
                programs.append(str(program_cell))
                program_groups.append(b_values.get(row_idx) or "" if group_col is not None else "")
                row_data = []
                orig_data = []  # stores (fval, raw_val) tuples
                for col_idx in included_cols:
                    val = merged_val(row_idx, col_idx)
                    try:
                        fval = float(val) if val is not None else None
                        if fval is not None and col_idx in pct_col_indices:
                            # Only scale up decimal fractions; whole-number percentages are kept as-is
                            if abs(fval) <= 1:
                                fval = fval * 100
                            val = fval  # keep raw consistent
                    except:
                        fval = None
                    # Store NaN only for truly missing values (None); 0 is a valid value
                    row_data.append(fval if fval is not None else np.nan)
                    orig_data.append((fval, val))  # keep raw value to detect N/A text
                data_values.append(row_data)
                original_values.append(orig_data)

        # Collect rows beyond heatmap_max_row — merged into heatmap as gray rows
        if include_below_rows:
            for row_idx in range(heatmap_max_row + 1, ws.max_row + 1):
                program_cell = merged_val(row_idx, program_col)
                # Only include this below-row if it contains any non-blank data in the
                # KPI columns; otherwise skip so we don't render empty grey rows.
                # Treat None or empty/whitespace-only strings as blank.
                if program_cell is not None and str(program_cell).strip() not in ('', 'None'):
                    has_any = False
                    for col_idx in included_cols:
                        v = merged_val(row_idx, col_idx)
                        if v is None:
                            continue
                        if isinstance(v, str) and v.strip() == '':
                            continue
                        has_any = True
                        break
                    if not has_any:
                        continue
                    below_programs.append(str(program_cell))
                    row_data = []
                    orig_row = []  # stores (fval, raw_val) tuples
                    for col_idx in included_cols:
                        val = merged_val(row_idx, col_idx)
                        try:
                            fval = float(val) if val is not None else None
                            if fval is not None and col_idx in pct_col_indices:
                                # Only scale up decimal fractions; whole-number percentages are kept as-is
                                if abs(fval) <= 1:
                                    fval = fval * 100
                                val = fval
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
        if show_row_groups and all_program_groups:
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
        y_labels_wrapped = ['<br>'.join(f'<b>{line}</b>' for line in wrap_label(p, max_len=24).split('<br>')) for p in all_programs]

        # Dynamically find the Per Program Target row by searching col 1 to program_col for the label
        per_program_target_row = None
        for row_idx in range(heatmap_max_row + 1, ws.max_row + 1):
            for col_idx in range(1, (program_col or 3) + 1):
                val = merged_val(row_idx, col_idx)
                if val and 'per program' in str(val).lower():
                    per_program_target_row = row_idx
                    break
            if per_program_target_row:
                break

        per_program_targets = []
        if per_program_target_row:
            for col_idx in range(data_col_start, data_col_start + len(kpi_names)):
                val = merged_val(per_program_target_row, col_idx)
                try:
                    per_program_targets.append(float(val) if val is not None else None)
                except:
                    per_program_targets.append(None)
        else:
            per_program_targets = [None] * len(kpi_names)

        # Fallback: if Per Program target cells are empty (e.g. uncached formulas),
        # compute from Annual Target row divided by number of heatmap programs
        if all(t is None for t in per_program_targets) and len(programs) > 0:
            annual_target_row = None
            for row_idx in range(heatmap_max_row + 1, ws.max_row + 1):
                for col_idx in range(1, (program_col or 3) + 1):
                    val = merged_val(row_idx, col_idx)
                    if val and 'annual target' in str(val).lower():
                        annual_target_row = row_idx
                        break
                if annual_target_row:
                    break
            if annual_target_row:
                n = len(programs)
                per_program_targets = []
                for col_idx in range(data_col_start, data_col_start + len(kpi_names)):
                    val = merged_val(annual_target_row, col_idx)
                    try:
                        annl = float(val) if val is not None else None
                        per_program_targets.append(annl / n if annl is not None and n > 0 else None)
                    except:
                        per_program_targets.append(None)
                # Patch below_orig so display row shows computed values not NA
                for bi, prog in enumerate(below_programs):
                    if 'per program' in prog.lower():
                        below_orig[bi] = [(t, t) for t in per_program_targets]
                        break
                # Re-merge after patch
                all_original_values = original_values + below_orig

        # Create dataframe using all rows
        if all_data_values and len(kpi_names) > 0:
            df_heatmap = pd.DataFrame(all_data_values, columns=kpi_names[:len(all_data_values[0])])
            df_heatmap.index = y_labels

            # Normalize using 3-point scale per column:
            #   min_val (lowest in rows 4-16) -> 0.0 (red)
            #   50 % Progress (per_program_target / 2)  -> 0.5 (yellow)
            #   per_program_target                 -> 1.0 (dark green)
            df_normalized = df_heatmap.copy()
            # Mark side columns (absolute Excel indices in side_cols) as NaN for data rows so they are not colored
            if side_cols:
                try:
                    for sc in side_cols:
                        if sc in included_cols:
                            pos = included_cols.index(sc)
                            if pos < df_normalized.shape[1]:
                                # only NaN for the main heatmap rows; below-rows will get -1 later
                                df_normalized.iloc[:len(programs), pos] = np.nan
                except Exception:
                    pass
            for col_i, col in enumerate(df_normalized.columns):
                # If this column maps to a declared side column, skip normalization (leave NaN)
                try:
                    if side_cols and col_i < len(included_cols) and included_cols[col_i] in side_cols:
                        df_normalized.iloc[:, col_i] = np.nan
                        continue
                except Exception:
                    pass
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
                        # treat below-heatmap rows as missing for coloring (no grey sentinel)
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

            # Previously we forced entire below-row rows to grey. Remove that
            # behavior so that below-row cells containing actual data are
            # colored using the same normalization rules as main rows, while
            # blank/NA cells remain uncolored.

            # Extended colorscale: grey for below-row sentinel (-1), then red→green for data (0..1)
            # With zmin=-1, zmax=1 the normalised position = (v+1)/2:
            #   v=-1  → pos 0.00  (grey, below-row rows)
            #   v= 0  → pos 0.50  (red,  data min)
            #   v= 0.25 → pos 0.625 (orange)
            #   v= 0.5 → pos 0.75  (yellow)
            #   v= 0.75 → pos 0.875 (light green)
            #   v= 1  → pos 1.00  (dark green)
            # Strict red (min) -> yellow (midpoint) -> green (target)
            colorscale = [
                [0.0, '#D73027'],  # red (min)
                [0.5, '#FFFF00'],  # yellow (midpoint)
                [1.0, '#1A7A1A'],  # dark green (target)
            ]

            # Smart numeric formatter — preserves significant figures for small values
            def fmt_val(v):
                """Format a number to at most 3 decimal places with thousand separators.
                Integers display without decimal point; very small values use 2 s.f. scientific."""
                if v == 0:
                    return '0'
                abs_v = abs(v)
                if abs_v < 0.0005:
                    # Too small for 3 dp — use 2 sig-fig scientific
                    return f"{v:.2e}"
                elif v == int(v):
                    return f"{int(v):,}"          # whole number with commas
                else:
                    formatted = f"{v:,.3f}".rstrip('0').rstrip('.')  # up to 3 dp, commas, strip trailing zeros
                    return formatted

            # Build text display
            text_display = []
            # Build set of column positions for fixed decimal formatting
            zero_dp_positions = set()
            if zero_decimal_cols:
                for ci, name in enumerate(kpi_names):
                    if any(sub.lower() in name.lower() for sub in zero_decimal_cols):
                        zero_dp_positions.add(ci)
            one_dp_positions = set()
            if one_decimal_cols:
                for ci, name in enumerate(kpi_names):
                    if any(sub.lower() in name.lower() for sub in one_decimal_cols):
                        one_dp_positions.add(ci)
            two_dp_positions = set()
            if two_decimal_cols:
                for ci, name in enumerate(kpi_names):
                    if any(sub.lower() in name.lower() for sub in two_decimal_cols):
                        two_dp_positions.add(ci)
            # Build set of row indices whose label matches zero_decimal_rows or one_decimal_rows substrings
            zero_dp_row_indices = set()
            if zero_decimal_rows:
                for ri, prog in enumerate(all_programs):
                    if any(sub.lower() in prog.lower() for sub in zero_decimal_rows):
                        zero_dp_row_indices.add(ri)
            one_dp_row_indices = set()
            if one_decimal_rows:
                for ri, prog in enumerate(all_programs):
                    if any(sub.lower() in prog.lower() for sub in one_decimal_rows):
                        one_dp_row_indices.add(ri)
            for row_i, row in enumerate(all_original_values):
                text_row = []
                row_force_zero_dp = row_i in zero_dp_row_indices
                row_force_one_dp = row_i in one_dp_row_indices
                for col_i, item in enumerate(row):
                    fval, raw = item if isinstance(item, tuple) else (item, None)
                    col_is_pct = col_i < len(included_cols) and included_cols[col_i] in pct_col_indices
                    # Show blank for truly empty cells; show 'NA' for N/A text; otherwise show raw text
                    is_na_text = isinstance(raw, str) and raw.strip().upper() in ('N/A', 'NA', '#N/A')
                    if fval is None:
                        # Preserve blank cells
                        if raw is None or (isinstance(raw, str) and raw.strip() == ''):
                            text_row.append('')
                        elif is_na_text:
                            text_row.append('NA')
                        else:
                            # non-numeric text that isn't N/A — show as-is
                            text_row.append(str(raw))
                    else:
                        if col_i in one_dp_positions:
                            formatted = f"{fval:,.1f}"
                        elif col_i in two_dp_positions:
                            formatted = f"{fval:,.2f}"
                        elif one_decimal_first_col and col_i == 0:
                            formatted = f"{fval:,.1f}"
                        elif row_force_one_dp:
                            formatted = f"{fval:,.1f}"
                        elif row_force_zero_dp or col_i in zero_dp_positions:
                            formatted = f"{round(fval):,}"
                        elif force_decimals is not None:
                            formatted = f"{fval:,.{force_decimals}f}"
                        else:
                            formatted = fmt_val(fval)
                        show_pct = col_is_pct and not suppress_pct_display
                        text_row.append(formatted + '%' if show_pct else formatted)
                text_display.append(text_row)

            # If requested, pad numeric text to align right using a monospace font.
            if monospace_numeric:
                ncols = len(kpi_names)
                nrows = len(text_display)
                col_max = [0] * ncols
                for j in range(ncols):
                    for i in range(nrows):
                        if j < len(text_display[i]):
                            s = str(text_display[i][j])
                        else:
                            s = ''
                        if len(s) > col_max[j]:
                            col_max[j] = len(s)
                for i in range(nrows):
                    for j in range(ncols):
                        if j < len(text_display[i]):
                            s = str(text_display[i][j])
                        else:
                            s = ''
                        text_display[i][j] = s.rjust(col_max[j], '\u00A0')
            text_family = 'Courier New, monospace' if monospace_numeric else 'Arial, sans-serif'

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
                textfont=dict(size=16, color='black', family=text_family),
                xgap=2,
                ygap=2
            )

            # Add a colorbar legend for the main heatmap (trace 0)
            try:
                # Primary heatmap is trace 0 from px.imshow — set a colorbar that maps the normalized values (-1..1) to human-friendly tick labels.
                cb = dict(
                    title='Progress',
                    titleside='top',
                    tickmode='array',
                    tickvals=[0.0, 0.5, 1.0],
                    ticktext=['No progress', '50% progress', 'Target achieved'],
                    ticks='outside',
                    thickness=24,
                    lenmode='fraction',
                    len=0.6,
                    outlinewidth=0,
                    tickfont=dict(size=12, color='black'),
                    titlefont=dict(size=12, color='black'),
                    bgcolor='rgba(255,255,255,0.9)',
                    x=0.99,
                    y=0.5,
                    xanchor='left',
                    yanchor='middle',
                )
                if len(fig.data) > 0:
                    # Disable Plotly's built-in colorbar; we use the HTML legend instead
                    try:
                        fig.data[0].update(showscale=False)
                    except Exception:
                        fig.data[0]['showscale'] = False
            except Exception:
                pass

            # Add a gray heatmap trace for below-row cells (rows beyond heatmap_max_row)
            # so those rows render as light-gray while keeping their text readable.
            if len(below_programs) > 0:
                below_text = text_display[len(programs):]
                fig.add_trace(go.Heatmap(
                    z=np.full((len(below_programs), len(kpi_names)), 1.0),
                    x=kpi_names,
                    y=all_programs[len(programs):],
                    text=np.array(below_text, dtype=object),
                    texttemplate='%{text}',
                    textfont=dict(size=16, color='black', family=text_family),
                    showscale=False,
                    coloraxis=None,
                    colorscale=[[0, '#d1d5db'], [1, '#d1d5db']],
                    zmin=0, zmax=1,
                    xgap=2, ygap=2,
                    hoverinfo='skip',
                ))

                # Overlay explicit text scatter for below-row cells so text is always on top
                flat_x = []
                flat_y = []
                flat_text = []
                for ri in range(len(programs), len(all_programs)):
                    row_idx = ri - len(programs)
                    for ci in range(len(kpi_names)):
                        # guard against uneven text_display rows
                        try:
                            t = text_display[ri][ci]
                        except Exception:
                            t = ''
                        flat_x.append(kpi_names[ci])
                        flat_y.append(all_programs[ri])
                        flat_text.append(str(t))

                fig.add_trace(go.Scatter(
                    x=flat_x,
                    y=flat_y,
                    mode='text',
                    text=flat_text,
                    textfont=dict(color='black', size=16, family=text_family),
                    hoverinfo='skip',
                ))

            # LABEL_PX: at -45� with wrapped labels, project max lines * line_height * sin(45�)
            LINE_H_PX = 12  # approx line height in px at 9pt
            max_lines = max((len(t.split('<br>')) for t in kpi_tick_names), default=1)
            longest_line = max((len(line.replace('<b>','').replace('</b>','')) for t in kpi_tick_names for line in t.split('<br>')), default=10)
            # Auto-detect short labels (e.g. years) vs long KPI names
            short_labels = longest_line <= 6 and max_lines == 1
            if short_labels:
                LABEL_PX = 30
                tick_angle = 0
                col_px = 35
            else:
                LABEL_PX = max(60, int((longest_line * CHAR_PX + max_lines * LINE_H_PX) * 0.71) + 10)
                tick_angle = -45
                col_px = 80  # wide enough so diagonal labels don't override each other
            BAND_PX  = 30
            GAP_PX   = group_gap if group_gap is not None else (8 if short_labels else 0)  # extra space between tick labels and group bands
            dynamic_top = LABEL_PX + GAP_PX + BAND_PX + (extra_top or 0)
            row_px = 32  # taller rows so text is more readable
            # Compute left margin based on longest program label so row text fits
            import re
            clean_y_labels = [re.sub(r'<[^>]+>', '', t) for t in y_labels_wrapped]
            longest_label_chars = max((len(s) for s in clean_y_labels), default=18)
            # increase base padding to give more horizontal room for long labels
            computed_left = int(longest_label_chars * CHAR_PX + 160)
            # bump defaults slightly
            LEFT_M = left_margin if left_margin is not None else max(computed_left, 400 if show_row_groups else 260)
            # Reserve enough right margin so the colorbar and its labels are visible
            RIGHT_M = max(140, 20)
            BOTTOM_M = 20
            chart_height = dynamic_top + 40 + len(all_programs) * row_px
            chart_width  = LEFT_M + RIGHT_M + len(kpi_names) * col_px

            # Paper coords above y=1.0 use the PLOT AREA height (not total chart height)
            # plot_area_height = chart_height - top_margin - bottom_margin
            plot_area_h = chart_height - dynamic_top - BOTTOM_M
            band_y0 = 1.0 + (LABEL_PX + GAP_PX) / plot_area_h
            band_y1 = band_y0 + BAND_PX / plot_area_h
            band_label_y = (band_y0 + band_y1) / 2

            # Move x-axis to top with wrapped tick labels
            fig.update_layout(
                height=chart_height,
                width=chart_width,
                xaxis=dict(
                    type='category',
                    side='top',
                    tickangle=tick_angle,
                    title='',
                    tickmode='array',
                    tickvals=kpi_names,
                    ticktext=kpi_tick_names,
                    tickfont=dict(size=10),
                    automargin=False,
                    showgrid=False,
                    zeroline=False,
                    showline=False,
                ),
                yaxis=dict(
                    tickfont=dict(size=10),
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
                margin=dict(l=LEFT_M, r=RIGHT_M, t=dynamic_top, b=BOTTOM_M),
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
                # Top group: no fill, larger black label
                fig.add_shape(
                    type='rect',
                    xref='x', yref='paper',
                    x0=start_idx - 0.5, x1=end_idx + 0.5,
                    y0=band_y0, y1=band_y1,
                    fillcolor='rgba(0,0,0,0)',
                    line=dict(color='rgba(0,0,0,0)', width=0),
                    layer='above'
                )
                # Wrap group name so long labels don't overflow their column span
                wrapped_group = '<br>'.join(f'<b>{line}</b>' for line in wrap_label(group_name, max_len=20).split('<br>'))
                fig.add_annotation(
                    xref='x', yref='paper',
                    x=x_center, y=band_label_y,
                    text=wrapped_group,
                    showarrow=False,
                    font=dict(color='black', size=14, family='Arial Black, Arial, sans-serif'),
                    align='center',
                    xanchor='center',
                    yanchor='middle',
                    bgcolor='rgba(0,0,0,0)'
                )
                # Add a thick black vertical line at the end of this column group
                # Extend from bottom of heatmap (y=1.0 paper) up through the column group band (band_y1)
                fig.add_shape(
                    type='line',
                    xref='x', yref='paper',
                    x0=end_idx + 0.5, x1=end_idx + 0.5,
                    # extend the vertical divider up to the KPI group header band
                    y0=0.0, y1=band_y1,
                    line=dict(color='#000000', width=2),
                    layer='above'
                )

            # Add program group bands to the LEFT of the y-axis (mirrors top KPI-type bands)
            if show_row_groups:
                # Compute band x positions dynamically so they scale correctly with LEFT_M
                # and the plot area width — prevents crowding when plot area is narrow.
                _plot_w_px = len(kpi_names) * col_px
                # Move group band further into the left margin to create a visible gap
                gx_outer = -(LEFT_M * 0.75) / _plot_w_px   # outer edge (furthest into margin)
                gx_inner = -(LEFT_M * 0.25) / _plot_w_px   # inner edge (closest to y-axis)
                gx_ann   = -(LEFT_M * 0.60) / _plot_w_px  # annotation midpoint (further left)
                for idx, (group_name, start_idx, end_idx) in enumerate(program_group_spans):
                    # Skip bands that are entirely beyond the main heatmap rows or have no group name
                    if start_idx >= len(programs) or not group_name:
                        continue
                    end_idx = min(end_idx, len(programs) - 1)
                    color = left_group_colors[idx % len(left_group_colors)]
                    y_center_index = int(round((start_idx + end_idx) / 2))
                    y_center_label = all_programs[y_center_index] if y_center_index < len(all_programs) else all_programs[-1]
                    # Left group: no fill, larger black horizontal label anchored to the right
                    fig.add_shape(
                        type='rect',
                        xref='paper', yref='y',
                        x0=gx_outer, x1=gx_inner,
                        y0=all_programs[start_idx], y1=all_programs[end_idx],
                        fillcolor='rgba(0,0,0,0)',
                        line=dict(color='rgba(0,0,0,0)', width=0),
                        layer='above'
                    )
                    # Wrap long group names so they don't overflow and anchor to the right
                    wrapped_left = '<br>'.join(f'<b>{line}</b>' for line in wrap_label(group_name, max_len=12).split('<br>'))
                    # Position group header: increase gap for specific groups
                    gn = (group_name or '').strip().lower()
                    # file basename to detect specific 4-1 files
                    bn = os.path.basename(excel_file_path).lower() if excel_file_path else ''
                    # If rendering the 4-1 or 4-2 special workbooks (Capacity/Product or Society Impact), open a larger gap
                    if (('4-1' in bn and ('capacity' in bn or 'product' in bn)) or
                        ('4-2' in bn and ('societ' in bn or 'inclusion' in bn or 'impact' in bn))):
                        # increase margin gap for 4-1/4-2 special workbooks
                        ann_x = gx_inner - 0.25
                        ann_anchor = 'left'
                    elif ('capacity build' in gn or 'capacity building' in gn or 'product development' in gn or
                          'societ' in gn or 'inclusion' in gn or 'impact' in gn):
                        # slightly smaller fallback gap for Capacity/Product/Societal groups
                        ann_x = gx_outer - 0.30
                        ann_anchor = 'left'
                    else:
                        # default — place near the precomputed midpoint
                        ann_x = gx_ann
                        ann_anchor = 'left'
                    ann_align = 'right' if ann_anchor == 'right' else 'left'
                    # Use Plotly's align/xanchor settings instead of embedding HTML
                    fig.add_annotation(
                        xref='paper', yref='y',
                        x=ann_x, y=y_center_label,
                        text=wrapped_left,
                        showarrow=False,
                        font=dict(color='black', size=12, family='Arial Black, Arial, sans-serif'),
                        align=ann_align,
                        xanchor=ann_anchor,
                        textangle=0,
                        yanchor='middle',
                        bgcolor='rgba(0,0,0,0)'
                    )
                    # Add a thick black horizontal line at the end of this row group
                    # Extend from outer edge of row-group band to right edge of heatmap
                    fig.add_shape(
                        type='line',
                        xref='paper', yref='y',
                        # Draw the horizontal connector across the entire data block
                        x0=gx_outer, x1=1.0,
                        # place on the row boundary so it does not run through cells
                        y0=end_idx + 0.5, y1=end_idx + 0.5,
                        line=dict(color='#000000', width=2),
                        layer='above'
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
        f'<th style="background-color:#6b7280;color:white;padding:8px 12px;border:1px solid rgba(0,0,0,0.12);font-weight:bold;">{col}</th>'
        for col in df.columns
    )
    rows_html = ""
    for i, row in df.iterrows():
        bg = "#f3f4f6" if i % 2 == 0 else "#e5e7eb"
        cells = "".join(
            f'<td style="background-color:{bg};padding:7px 12px;border:1px solid rgba(0,0,0,0.12);">{val}</td>'
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

# Helper: return HTML legend matching Service KPI styling (red/yellow/green)
def get_heatmap_legend_html():
    return (
        '<div style="display:flex; gap:12px; align-items:center; margin-top:0px; margin-bottom:2px; font-family: Arial, sans-serif;">'
        '<div style="display:flex; align-items:center; gap:6px;"><span style="width:16px;height:16px;background:#D73027;display:inline-block;border-radius:3px;"></span><span>No Progress</span></div>'
        '<div style="display:flex; align-items:center; gap:6px;"><span style="width:16px;height:16px;background:#FFFF00;display:inline-block;border-radius:3px; border:1px solid rgba(0,0,0,0.12);"></span><span>50 % Progress</span></div>'
        '<div style="display:flex; align-items:center; gap:6px;"><span style="width:16px;height:16px;background:#1A7A1A;display:inline-block;border-radius:3px;"></span><span>Target Achieved</span></div>'
        '</div>'
    )

# Load data
df_programs, df_services, df_heatmap = load_kpi_data()

# Tabs
tab1, tab2, tab3 = st.tabs(["📊Program Output KPIs (Aggregate)", "🌡️ Program Output KPI (by Program)", "🏢 Service Unit KPIs"])

# Programs Tab
with tab1:
    st.markdown('<h2 style="font-family: Arial, sans-serif; font-size:20px; margin:6px 0;">📊 Program Output KPIs</h2>', unsafe_allow_html=True)

    # Create three subtabs as requested: Number, FTE, Million (USD)
    prog_sub_1, prog_sub_2, prog_sub_3 = st.tabs(["Program KPI by Number", "Program KPI by Full Time Equivalent (FTE)", "Program KPI by Million (USD)"])

    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    program_file = os.path.join(root_dir, 'data', 'Program Output KPIs.xlsx')

    # --- Program KPI by Number: move previous aggregate table here ---
    with prog_sub_1:
        # Use the dedicated render function so we can easily customize later
        render_program_kpi_number(program_file, df=df_programs)

    # --- Program KPI by FTE: render FTE-specific Excel if present ---
    with prog_sub_2:
        fte_file = os.path.join(root_dir, 'data', 'Program Output KPIs by FTE.xlsx')
        render_program_kpi_fte_with_color_coding(fte_file)

    # --- Program KPI by Million (USD): render USD-specific Excel if present ---
    with prog_sub_3:
        usd_file = os.path.join(root_dir, 'data', 'Program Output KPIs by $.xlsx')
        render_program_kpi_usd(usd_file)

# KPI By Program Tab (now second)
with tab2:
    st.subheader("🌡️ 2025 Program Output KPI (by Program)")
    
    # Two main sub-tabs
    sub_tab_a, sub_tab_b = st.tabs(["🔬 Research, Training, Product Development", "🏆 Recognition, Societal Impact & Inclusivity"])
    
    # ==================== Research, Training, Product Development ====================
    with sub_tab_a:
        st.markdown("### Research, Training, Product Development")
        
        rtpd_tabs = st.tabs(["KPI by Number", "KPI by Full Time Equivalent (FTE)", "KPI by million (USD)", "KPI by Number over time"])
        
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        with rtpd_tabs[0]:
            st.write("**Research, Training, Product Development - KPI by Number**")
            try:
                # Use only the pre-rendered image for Heat Map 1 with the exact filename in the data folder.
                img_path = os.path.join(root_dir, 'data', 'Heat_map_1.png')
                if os.path.exists(img_path):
                    st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                    # Embed the image as a base64 data URL so it scales to the page width responsively
                    try:
                        with open(img_path, 'rb') as _f:
                            _img_bytes = _f.read()
                        _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                        _img_html = (
                            f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                            f'<img src="data:image/png;base64,{_img_b64}" '
                            'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center 10%;" />'
                            '</div>'
                        )
                        st.markdown(_img_html, unsafe_allow_html=True)
                    except Exception:
                        # Fallback to a large fixed width display
                        st.image(img_path, width=1200)
                else:
                    st.info('📁 Waiting for: Heat_map_1.png')
            except Exception as e:
                st.warning(f'Could not load heatmap image: {str(e)}')

        with rtpd_tabs[1]:
            st.write("**Research, Training, Product Development - KPI by Full Time Equivalent (FTE)**")
            try:
                # Use only the pre-rendered image for Heat Map 2 in the data folder.
                img_path = os.path.join(root_dir, 'data', 'Heat_map_2.png')
                if os.path.exists(img_path):
                    st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                    try:
                        with open(img_path, 'rb') as _f:
                            _img_bytes = _f.read()
                        _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                        _img_html = (
                            f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                            f'<img src="data:image/png;base64,{_img_b64}" '
                            'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center 10%;" />'
                            '</div>'
                        )
                        st.markdown(_img_html, unsafe_allow_html=True)
                    except Exception:
                        st.image(img_path, width=1200)
                else:
                    st.info('📁 Waiting for: Heat_map_2.png')
            except Exception as e:
                st.warning(f"Could not load heatmap image: {str(e)}")

        with rtpd_tabs[2]:
            st.write("**Research, Training, Product Development - KPI by million (USD)**")
            try:
                # Use only the pre-rendered image for Heat Map 3 in the data folder.
                img_path = os.path.join(root_dir, 'data', 'Heat_map_3.png')
                if os.path.exists(img_path):
                    st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                    try:
                        with open(img_path, 'rb') as _f:
                            _img_bytes = _f.read()
                        _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                        _img_html = (
                            f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                            f'<img src="data:image/png;base64,{_img_b64}" '
                            'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center 10%;" />'
                            '</div>'
                        )
                        st.markdown(_img_html, unsafe_allow_html=True)
                    except Exception:
                        st.image(img_path, width=1200)
                else:
                    st.info('📁 Waiting for: Heat_map_3.png')
            except Exception as e:
                st.warning(f"Could not load heatmap image: {str(e)}")

        with rtpd_tabs[3]:
            st.write("**Research, Training, Product Development - KPI by Number over time**")
            try:
                heatmap_file = os.path.join(root_dir, 'data', 'Heat map 4.xlsx')
                heatmap_file_4_1 = os.path.join(root_dir, 'data', 'Heat map 4-1 Research Outputs.xlsx')
                heatmap_choice = heatmap_file_4_1 if os.path.exists(heatmap_file_4_1) else heatmap_file
                if os.path.exists(heatmap_choice):
                    # Provide three focused KPI-over-time sub-tabs so users can
                    # view Research Outputs, Capacity Building, and Product
                    # Development separately.
                    ot_tabs = st.tabs(["Research Outputs", "Capacity Building", "Product Development"])

                    with ot_tabs[0]:
                        # Research Outputs — use the pre-rendered image only (no Excel fallback)
                        img_path = os.path.join(root_dir, 'data', 'Heat map 4-1 Research Outputs.png')
                        if os.path.exists(img_path):
                            try:
                                with open(img_path, 'rb') as _f:
                                    _img_bytes = _f.read()
                                _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                                _img_html = (
                                    f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                                    f'<img src="data:image/png;base64,{_img_b64}" '
                                    'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center 10%;" />'
                                    '</div>'
                                )
                                # Reduce spacing by modifying legend's margin and embedding together
                                legend_html = get_heatmap_legend_html().replace('margin-bottom:8px;', 'margin-bottom:2px;')
                                combined_html = f'<div style="margin:0;padding:0;">{legend_html}{_img_html}</div>'
                                st.markdown(combined_html, unsafe_allow_html=True)
                            except Exception:
                                st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                                st.image(img_path, width=1200)
                        else:
                            st.info('📁 Waiting for: Heat map 4-1 Research Outputs.png')

                    with ot_tabs[1]:
                        # Capacity Building — use the pre-rendered image only (no Excel fallback)
                        img_path = os.path.join(root_dir, 'data', 'Heat map 4-1 Capacity Building.png')
                        if os.path.exists(img_path):
                            st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                            try:
                                with open(img_path, 'rb') as _f:
                                    _img_bytes = _f.read()
                                _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                                _img_html = (
                                    f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                                    f'<img src="data:image/png;base64,{_img_b64}" '
                                    'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center 10%;" />'
                                    '</div>'
                                )
                                st.markdown(_img_html, unsafe_allow_html=True)
                            except Exception:
                                st.image(img_path, width=1200)
                        else:
                            st.info('📁 Waiting for: Heat map 4-1 Capacity Building.png')

                    with ot_tabs[2]:
                        # Product Development — use the pre-rendered image only (no Excel fallback)
                        img_path = os.path.join(root_dir, 'data', 'Heat map 4-1 - Product Development.png')
                        if os.path.exists(img_path):
                            st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                            try:
                                with open(img_path, 'rb') as _f:
                                    _img_bytes = _f.read()
                                _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                                _img_html = (
                                    f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                                    f'<img src="data:image/png;base64,{_img_b64}" '
                                    'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center 10%;" />'
                                    '</div>'
                                )
                                st.markdown(_img_html, unsafe_allow_html=True)
                            except Exception:
                                st.image(img_path, width=1200)
                        else:
                            st.info('📁 Waiting for: Heat map 4-1 - Product Development.png')
                else:
                    st.info("📁 Waiting for: Heat map 4.xlsx")
            except Exception as e:
                st.warning(f"Could not load heatmap: {str(e)}")
    
    # ==================== Recognition, Societal Impact & Inclusivity ====================
    with sub_tab_b:
        st.markdown("### Recognition, Societal Impact & Inclusivity")
        
        rsi_tabs = st.tabs(["KPI by Number", "KPI by Full Time Equivalent (FTE)", "KPI by million (USD)", "KPI by Number over time"])
        
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        with rsi_tabs[0]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by Number**")
            try:
                # Use only the pre-rendered image for Heat Map 5 (Recognition/Societal) in data folder.
                img_path = os.path.join(root_dir, 'data', 'Heat_map_5.png')
                if os.path.exists(img_path):
                    st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                    try:
                        with open(img_path, 'rb') as _f:
                            _img_bytes = _f.read()
                        _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                        _img_html = (
                            f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                            f'<img src="data:image/png;base64,{_img_b64}" '
                            'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center 10%;" />'
                            '</div>'
                        )
                        st.markdown(_img_html, unsafe_allow_html=True)
                    except Exception:
                        st.image(img_path, width=1200)
                else:
                    st.info('📁 Waiting for: Heat_map_5.png')
            except Exception as e:
                st.warning(f'Could not load heatmap image: {str(e)}')

        with rsi_tabs[1]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by Full Time Equivalent (FTE)**")
            try:
                # Use only pre-rendered image for Heat Map 6
                img_path = os.path.join(root_dir, 'data', 'Heat_map_6.png')
                if os.path.exists(img_path):
                    st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                    try:
                        with open(img_path, 'rb') as _f:
                            _img_bytes = _f.read()
                        _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                        _img_html = (
                            f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                            f'<img src="data:image/png;base64,{_img_b64}" '
                            'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center 10%;" />'
                            '</div>'
                        )
                        st.markdown(_img_html, unsafe_allow_html=True)
                    except Exception:
                        st.image(img_path, width=1200)
                else:
                    st.info('📁 Waiting for: Heat_map_6.png')
            except Exception as e:
                st.warning(f'Could not load heatmap image: {str(e)}')

        with rsi_tabs[2]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by million (USD)**")
            try:
                # Use only pre-rendered image for Heat Map 7
                img_path = os.path.join(root_dir, 'data', 'Heat_map_7.png')
                if os.path.exists(img_path):
                    st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                    try:
                        with open(img_path, 'rb') as _f:
                            _img_bytes = _f.read()
                        _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                        _img_html = (
                            f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                            f'<img src="data:image/png;base64,{_img_b64}" '
                            'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center 10%;" />'
                            '</div>'
                        )
                        st.markdown(_img_html, unsafe_allow_html=True)
                    except Exception:
                        st.image(img_path, width=1200)
                else:
                    st.info('📁 Waiting for: Heat_map_7.png')
            except Exception as e:
                st.warning(f'Could not load heatmap image: {str(e)}')

        with rsi_tabs[3]:
            st.write("**Recognition, Societal Impact & Inclusivity - KPI by Number over time**")
            # Create focused sub-tabs for Recognition and Societal Impact
            rsi_ot_sub = st.tabs(["Recognition and Reputation", "Societal Impact and Inclusion"]) 
            # Base heatmap file fallback
            # For Recognition and Societal Impact, prefer dedicated 4-2 Excel files; otherwise use Heat_map_5.png image.
            with rsi_ot_sub[0]:
                try:
                    # Use the pre-rendered image for Recognition and Reputation (no Excel fallback)
                    img_path = os.path.join(root_dir, 'data', 'Heat map 4-2 Recognition and Reputation.png')
                    if os.path.exists(img_path):
                        st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                        try:
                            with open(img_path, 'rb') as _f:
                                _img_bytes = _f.read()
                            _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                            _img_html = (
                                f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                                f'<img src="data:image/png;base64,{_img_b64}" '
                                'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center 10%;" />'
                                '</div>'
                            )
                            st.markdown(_img_html, unsafe_allow_html=True)
                        except Exception:
                            st.image(img_path, width=1200)
                    else:
                        st.info('📁 Waiting for: Heat map 4-2 Recognition and Reputation.png')
                except Exception as e:
                    st.warning(f'Could not load Recognition heatmap: {str(e)}')

            with rsi_ot_sub[1]:
                try:
                    # Use the pre-rendered image for Society Impact and Inclusion (no Excel fallback)
                    img_path = os.path.join(root_dir, 'data', 'Heat map 4-2 Society Impact and Inclusion.png')
                    if os.path.exists(img_path):
                        st.markdown(get_heatmap_legend_html(), unsafe_allow_html=True)
                        try:
                            with open(img_path, 'rb') as _f:
                                _img_bytes = _f.read()
                            _img_b64 = base64.b64encode(_img_bytes).decode('utf-8')
                            _img_html = (
                                f'<div style="position:relative; width:100%; max-width:1200px; height:600px; margin:0 auto;">'
                                f'<img src="data:image/png;base64,{_img_b64}" '
                                'style="position:absolute; top:0; left:0; width:100%; height:100%; object-fit:contain; object-position:center top;" />'
                                '</div>'
                            )
                            st.markdown(_img_html, unsafe_allow_html=True)
                        except Exception:
                            st.image(img_path, width=1200)
                    else:
                        st.info('📁 Waiting for: Heat map 4-2 Society Impact and Inclusion.png')
                except Exception as e:
                    st.warning(f'Could not load Societal Impact heatmap: {str(e)}')

# Service Units Tab (now third)
with tab3:
    st.markdown('<h2 style="font-family: Arial, sans-serif; font-size:20px; margin:6px 0;">🏢 Service Unit KPIs</h2>', unsafe_allow_html=True)
    
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    service_file = os.path.join(root_dir, 'data', 'Service Unit KPIs.xlsx')
    
    try:
        # Service Unit KPIs: Target in column C (3) and Actual in column D (4)
        html_services = excel_to_html_with_merged_cells(service_file, no_decimals=True, highlight_row_keyword='service unit key performance', target_col=3, actual_col=4)
        # Legend for Service Unit KPIs
        html_legend_srv = (
            '<div style="display:flex; gap:12px; align-items:center; margin-top:0px; margin-bottom:2px; font-family: Arial, sans-serif;">'
            '<div style="display:flex; align-items:center; gap:6px;"><span style="width:16px;height:16px;background:#D73027;display:inline-block;border-radius:3px;"></span><span>No Progress</span></div>'
            '<div style="display:flex; align-items:center; gap:6px;"><span style="width:16px;height:16px;background:#FFFF00;display:inline-block;border-radius:3px; border:1px solid rgba(0,0,0,0.12);"></span><span>50 % Progress</span></div>'
            '<div style="display:flex; align-items:center; gap:6px;"><span style="width:16px;height:16px;background:#1A7A1A;display:inline-block;border-radius:3px;"></span><span>Target Achieved</span></div>'
            '</div>'
        )


        # Adjust alignment for Service Unit KPIs table:
        # - first two data rows: left-aligned
        # - last two data rows: right-aligned
        # - any row highlighted with the green background: center-aligned
        def adjust_service_alignment(html):
            rows = re.findall(r'(<tr.*?>.*?</tr>)', html, flags=re.DOTALL | re.IGNORECASE)
            if not rows:
                return html

            # find first data row index (first row that contains a <td>)
            data_start = 0
            for idx, r in enumerate(rows):
                if '<td' in r.lower():
                    data_start = idx
                    break

            # determine total logical columns from the header rows (before data_start)
            total_cols = 0
            for hdr_idx in range(data_start):
                hdr = rows[hdr_idx]
                th_tags = re.findall(r'<th[^>]*>', hdr, flags=re.IGNORECASE)
                if th_tags:
                    # use the last header row that contains <th>
                    total_cols = 0
                    for tag in th_tags:
                        cs = re.search(r'colspan\s*=\s*"(\d+)"', tag, flags=re.IGNORECASE)
                        total_cols += int(cs.group(1)) if cs else 1

            if total_cols == 0:
                # fallback: count <td> in first data row ignoring colspan
                first_data = rows[data_start] if data_start < len(rows) else ''
                total_cols = len(re.findall(r'<t[dh][^>]*>', first_data, flags=re.IGNORECASE))

            data_rows = rows[data_start:]
            n = len(data_rows)
            new_rows = rows.copy()

            def update_tag_alignment(tag, align):
                # update or add text-align in the tag's style attribute
                if re.search(r'style\s*=\s*"', tag, flags=re.IGNORECASE):
                    def repl(m):
                        styles = m.group(1)
                        if re.search(r'text-align\s*:', styles, flags=re.IGNORECASE):
                            styles = re.sub(r'text-align\s*:\s*[^;]+', f'text-align: {align}', styles, flags=re.IGNORECASE)
                        else:
                            styles = styles.rstrip() + f'; text-align: {align}'
                        return f'style="{styles}"'
                    return re.sub(r'style\s*=\s*"([^"]*)"', repl, tag, flags=re.IGNORECASE)
                else:
                    # insert style before closing bracket
                    return tag[:-1] + f' style="text-align: {align};">'

            last_two_positions = {total_cols - 1, total_cols} if total_cols >= 2 else {total_cols}

            for i, r in enumerate(data_rows):
                row_idx = data_start + i
                new_r = r

                # If row contains green highlight, center entire row
                if '#00891a' in new_r.lower() or 'background-color: #00891a' in new_r.lower():
                    # set all cell tags to center
                    def center_all(m):
                        return update_tag_alignment(m.group(0), 'center')
                    new_r = re.sub(r'(<t[dh][^>]*>)', center_all, new_r, flags=re.IGNORECASE)
                    new_rows[row_idx] = new_r
                    continue

                # For first two data rows: left-align non-numeric cells, but keep numeric cells right-aligned
                if i < 2:
                    # process full cell tags to decide per-cell alignment
                    cell_pattern_local = re.compile(r'(<t[dh][^>]*>)(.*?)(</t[dh]>)', flags=re.IGNORECASE | re.DOTALL)
                    parts_local = []
                    last_end_local = 0
                    tags_local = list(cell_pattern_local.finditer(new_r))
                    if tags_local:
                        col_pos_local = 1
                        for m2 in tags_local:
                            s2, e2 = m2.span(1)
                            open_tag = m2.group(1)
                            inner = m2.group(2)
                            close_tag = m2.group(3)
                            # determine colspan
                            cs2 = re.search(r'colspan\s*=\s*"(\d+)"', open_tag, flags=re.IGNORECASE)
                            span2 = int(cs2.group(1)) if cs2 else 1
                            tag_start2 = col_pos_local
                            tag_end2 = col_pos_local + span2 - 1

                            inner_text = re.sub(r'<[^>]+>', '', inner or '').strip()
                            cleaned = inner_text.replace('\u00A0', '').replace('\xa0', '').replace(',', '').strip()
                            if cleaned.startswith('(') and cleaned.endswith(')'):
                                cleaned_num = '-' + cleaned[1:-1]
                            else:
                                cleaned_num = cleaned
                            if cleaned_num.endswith('%'):
                                cleaned_num = cleaned_num[:-1]
                            is_numeric_local = False
                            try:
                                if cleaned_num != '':
                                    float(cleaned_num)
                                    is_numeric_local = True
                            except Exception:
                                is_numeric_local = False

                            # If this tag overlaps columns 4-9, preserve center alignment
                            if tag_end2 >= 4 and tag_start2 <= 9:
                                new_open = update_tag_alignment(open_tag, 'center')
                            else:
                                # choose alignment
                                if is_numeric_local:
                                    new_open = update_tag_alignment(open_tag, 'right')
                                else:
                                    new_open = update_tag_alignment(open_tag, 'left')

                            parts_local.append(new_r[last_end_local:s2])
                            parts_local.append(new_open)
                            parts_local.append(inner)
                            parts_local.append(close_tag)
                            last_end_local = e2 + len(m2.group(2)) + len(m2.group(3))
                            col_pos_local += span2
                        parts_local.append(new_r[last_end_local:])
                        new_r = ''.join(parts_local)

                # For last two logical columns and numeric cells: ensure right-alignment
                # Parse full cell tags (open, inner, close) to map logical column positions
                cell_pattern = re.compile(r'(<t[dh][^>]*>)(.*?)(</t[dh]>)', flags=re.IGNORECASE | re.DOTALL)
                tags = list(cell_pattern.finditer(new_r))
                if tags:
                    col_pos = 1
                    parts = []
                    last_end = 0
                    min_last = min(last_two_positions)
                    for m in tags:
                        s, e = m.span(1)
                        open_tag = m.group(1)
                        inner = m.group(2)
                        close_tag = m.group(3)
                        # determine colspan
                        cs = re.search(r'colspan\s*=\s*"(\d+)"', open_tag, flags=re.IGNORECASE)
                        span = int(cs.group(1)) if cs else 1
                        tag_start = col_pos
                        tag_end = col_pos + span - 1

                        # Skip center/left rules already applied for green rows and first two rows
                        # Detect numeric content robustly (commas, NBSP, parentheses, percent)
                        inner_text = re.sub(r'<[^>]+>', '', inner or '').strip()
                        is_numeric = False
                        if inner_text:
                            # normalize whitespace and non-breaking spaces
                            cleaned = inner_text.replace('\u00A0', '').replace('\xa0', '').replace(',', '').strip()
                            # handle parentheses negative like (123)
                            if cleaned.startswith('(') and cleaned.endswith(')'):
                                cleaned_num = '-' + cleaned[1:-1]
                            else:
                                cleaned_num = cleaned
                            # strip percent
                            if cleaned_num.endswith('%'):
                                cleaned_num = cleaned_num[:-1]
                            # try float parse
                            try:
                                float(cleaned_num)
                                is_numeric = True
                            except Exception:
                                is_numeric = False

                        new_open = open_tag
                        # Priority: if cell starts within last-two logical cols -> right ONLY if numeric
                        # Else if numeric and this row isn't in first-two -> right
                        # Preserve center alignment for any tags that overlap columns 4-9
                        if tag_end >= 4 and tag_start <= 9:
                            # force center for these columns
                            new_open = update_tag_alignment(new_open, 'center')
                        else:
                            if tag_start >= min_last:
                                if is_numeric:
                                    new_open = update_tag_alignment(new_open, 'right')
                                else:
                                    # leave alignment as-is (preserve earlier centering for non-numeric)
                                    pass
                            elif is_numeric and i >= 2:
                                new_open = update_tag_alignment(new_open, 'right')

                        parts.append(new_r[last_end:s])
                        parts.append(new_open)
                        parts.append(inner)
                        parts.append(close_tag)
                        last_end = e + len(m.group(2)) + len(m.group(3))
                        col_pos += span
                    parts.append(new_r[last_end:])
                    new_r = ''.join(parts)

                new_rows[row_idx] = new_r

            # replace rows in original HTML sequentially
            new_html = html
            for orig, new in zip(rows, new_rows):
                if orig != new:
                    new_html = new_html.replace(orig, new, 1)
            return new_html

        # skip post-processing alignment adjustments to preserve generator's per-column alignment
        st.markdown(html_legend_srv + html_services, unsafe_allow_html=True)
    except Exception as e:
        st.warning(f"Could not render with merged cells: {str(e)}")
        # Fallback: format numeric columns to have no decimals and convert to strings
        display_df_s = df_services.copy()
        for col in display_df_s.select_dtypes(include=["number"]).columns:
            display_df_s[col] = display_df_s[col].apply(lambda x: "" if pd.isna(x) else str(int(round(x))))
        st.dataframe(display_df_s, width='stretch', height=600)
    # Download button
    csv_services = df_services.to_csv(index=False)
    st.download_button(
        label="⬇️ Download 2025 Service Unit KPIs as CSV",
        data=csv_services,
        file_name="2025_Service_Unit_KPIs.csv",
        mime="text/csv"
    )



st.markdown("---")
st.caption("Last updated: April 8, 2026 | IITA Key Performance Indicator (KPI) Dashboard")
