import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
import io
import re
from datetime import datetime

# ------------------------------------------------------------------------------
# DEFAULT RURAL VEHICLE MASTER LIST
# ------------------------------------------------------------------------------
DEFAULT_RURAL_LIST = [
    # R-Series (All 27 vehicles)
    "R-1", "R-2", "R-3", "R-4", "R-5", "R-6", "R-7", "R-8", "R-9", "R-10",
    "R-11", "R-12", "R-13", "R-14", "R-15", "R-16", "R-17", "R-18", "R-19",
    "R-21", "R-23", "R-24", "R-32", "R-33", "R-35", "R-40", "R-41",
    
    # RIC-Series (All 17 vehicles)
    "RIC-20", "RIC-22", "RIC-25", "RIC-26", "RIC-27", "RIC-28", "RIC-29",
    "RIC-30", "RIC-31", "RIC-34", "RIC-36", "RIC-37", "RIC-38", "RIC-39",
    "RIC-42", "RIC-43", "RIC-44",
    
    # Other Rural Vehicles
    "GAB-4046", "SA-6766", "STR-6637", "MC-5", "TT-06", "GTD-694",
    "GAS-1694", "BN-3932", "DGK-1763", "TT-11", "GAU-8135", "GAJ-19-47",
    "SAA-2946", "ST-663", "GAJ-3943", "TT-05", "OKA-2192", "SAA-7988"
]

st.set_page_config(
    page_title="Mileage Processor | Dev by Ashaan",
    page_icon="🧾",
    layout="wide"
)

# ------------------------------------------------------------------------------
# AESTHETIC RECEIPT UI STYLING
# ------------------------------------------------------------------------------
st.markdown("""
    <style>
    .receipt-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
        padding: 25px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.15);
        margin-bottom: 25px;
        color: white;
    }
    .receipt-header {
        font-family: 'Segoe UI', sans-serif;
        font-size: 28px;
        font-weight: 700;
        text-align: center;
        letter-spacing: 1px;
        color: white;
        margin-bottom: 5px;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    .receipt-sub {
        font-family: 'Segoe UI', sans-serif;
        text-align: center;
        color: rgba(255,255,255,0.9);
        font-size: 15px;
        margin-bottom: 5px;
        font-weight: 300;
    }
    .receipt-sub b {
        color: #FFD700;
        font-weight: 600;
    }
    .stButton>button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 25px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="receipt-card">
    <div class="receipt-header">📊 MILEAGE EXECUTIVE AUDIT DASHBOARD</div>
    <div class="receipt-sub">Developed by: <b>Muhammad Ashaan</b> | Bellanix Tech</div>
</div>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------------------------
def parse_period_to_hours(val):
    if pd.isna(val) or val is None:
        return -1.0
    val_str = str(val).replace('\xa0', '').strip()
    if not val_str:
        return -1.0
    
    if ':' in val_str:
        parts = val_str.split(':')
        try:
            hrs = float(parts[0])
            mins = float(parts[1]) if len(parts) > 1 else 0.0
            secs = float(parts[2]) if len(parts) > 2 else 0.0
            return hrs + (mins / 60.0) + (secs / 3600.0)
        except:
            return -1.0
    try:
        f = float(val_str)
        return f * 24.0 if f <= 1.0 else f
    except:
        return -1.0

def extract_digits(val):
    return ''.join([ch for ch in str(val) if ch.isdigit()])

def normalize_reg(val):
    if pd.isna(val):
        return ""
    return str(val).replace('\xa0', '').replace('-', '').replace(' ', '').strip().upper()

def is_rural_vehicle(reg_str, rural_set, numeric_rural_set, normalized_rural_set):
    raw_clean = str(reg_str).replace('\xa0', '').strip().upper()
    norm_clean = normalize_reg(reg_str)
    
    if raw_clean in rural_set or norm_clean in normalized_rural_set:
        return True
    if raw_clean.startswith('R-') or raw_clean.startswith('RIC-'):
        return True
        
    digits = extract_digits(raw_clean)
    if digits and digits in numeric_rural_set:
        return True
        
    return False

def extract_file_date(df_raw):
    for r in range(min(5, len(df_raw))):
        row_str = " ".join(df_raw.iloc[r].dropna().astype(str).tolist())
        if '[' in row_str and ']' in row_str:
            p1, p2 = row_str.find('['), row_str.find(']')
            extracted = row_str[p1+1:p2].split('-')[0].strip()
            if len(extracted) >= 6:
                return extracted
        match = re.search(r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b', row_str)
        if match:
            return match.group(0)
    return datetime.now().strftime("%m/%d/%Y")

def process_mileage_df(df_raw):
    header_row_idx, col_reg_idx, col_period_idx = -1, -1, -1
    for r_idx in range(min(15, len(df_raw))):
        row_vals = df_raw.iloc[r_idx].astype(str).str.replace('\xa0', '', regex=False).str.strip().str.upper()
        c_reg, c_per = -1, -1
        for c_idx, val in enumerate(row_vals):
            if val in ['REG#', 'REGISTRATION', 'VEHICLE NO', 'REGISTRATION NO', 'REG NO']:
                c_reg = c_idx
            if val in ['PERIOD', 'DURATION', 'HOURS']:
                c_per = c_idx
        if c_reg != -1 and c_per != -1:
            header_row_idx, col_reg_idx, col_period_idx = r_idx, c_reg, c_per
            break
    return header_row_idx, col_reg_idx, col_period_idx

# ------------------------------------------------------------------------------
# ENHANCED EXCEL STYLING FUNCTION
# ------------------------------------------------------------------------------
def apply_professional_styling(ws, headers, num_cols, urban_count, rural_count, total_count, meta_date, selected_city):
    """Apply professional corporate styling to the Excel worksheet"""
    
    # Color Palette - Corporate Blue Theme
    COLORS = {
        'primary_dark': '1C3D5A',      # Deep navy
        'primary': '2A5C8A',            # Corporate blue
        'primary_light': '4A8CC4',      # Light blue
        'accent_gold': 'C5A45D',        # Gold for highlights
        'accent_green': '2E7D32',       # Forest green for rural
        'accent_red': 'C62828',         # Deep red for warnings
        'header_bg': '1A365D',          # Very dark blue for headers
        'alt_row': 'F7FAFC',            # Very light gray
        'white': 'FFFFFF',
        'border': 'CBD5E0',             # Soft gray border
        'text_dark': '2D3748',
        'text_light': 'FFFFFF'
    }
    
    # Fonts
    header_font = Font(name='Segoe UI', size=12, bold=True, color=COLORS['text_light'])
    title_font = Font(name='Segoe UI', size=16, bold=True, color=COLORS['text_light'])
    metric_font = Font(name='Segoe UI', size=11, bold=True, color=COLORS['primary_dark'])
    
    # Alignments
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # Borders
    thin_border = Border(
        left=Side(style='thin', color=COLORS['border']),
        right=Side(style='thin', color=COLORS['border']),
        top=Side(style='thin', color=COLORS['border']),
        bottom=Side(style='thin', color=COLORS['border'])
    )
    
    # Row 1: Main Title (Merged)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    title_cell = ws.cell(row=1, column=1, value=f"🏢 {selected_city} VEHICLE MILEAGE EXECUTIVE REPORT")
    title_cell.font = Font(name='Segoe UI', size=18, bold=True, color=COLORS['text_light'])
    title_cell.fill = PatternFill(start_color=COLORS['header_bg'], end_color=COLORS['header_bg'], fill_type='solid')
    title_cell.alignment = center_align
    ws.row_dimensions[1].height = 45
    
    # Row 2: Subtitle with date and report ID
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
    subtitle_cell = ws.cell(row=2, column=1, value=f"📅 Report Date: {meta_date} | 📊 Audit ID: MILE-{datetime.now().strftime('%Y%m%d')}")
    subtitle_cell.font = Font(name='Segoe UI', size=10, color=COLORS['primary_light'])
    subtitle_cell.fill = PatternFill(start_color='EBF4FC', end_color='EBF4FC', fill_type='solid')
    subtitle_cell.alignment = center_align
    ws.row_dimensions[2].height = 28
    
    # Row 3-4: Metrics Dashboard (4 columns)
    metrics_data = [
        ('Total Fleet (20-24h)', total_count, COLORS['primary'], '🚗'),
        ('Urban Fleet', urban_count, COLORS['primary_light'], '🏙️'),
        ('Rural Fleet', rural_count, COLORS['accent_green'], '🌾'),
        ('Completion Rate', f"{round((total_count/max(1,total_count+urban_count+rural_count))*100)}%", COLORS['accent_gold'], '📈')
    ]
    
    col_width = num_cols // 4
    for i, (label, value, color, icon) in enumerate(metrics_data):
        start_col = i * col_width + 1
        end_col = (i + 1) * col_width
        ws.merge_cells(start_row=3, start_column=start_col, end_row=4, end_column=end_col)
        
        cell = ws.cell(row=3, column=start_col, value=f"{icon} {label}")
        cell.font = Font(name='Segoe UI', size=9, bold=True, color=COLORS['text_dark'])
        cell.alignment = center_align
        cell.fill = PatternFill(start_color='F7FAFC', end_color='F7FAFC', fill_type='solid')
        
        value_cell = ws.cell(row=4, column=start_col, value=str(value))
        value_cell.font = Font(name='Segoe UI', size=16, bold=True, color=color)
        value_cell.alignment = center_align
        value_cell.fill = PatternFill(start_color='F7FAFC', end_color='F7FAFC', fill_type='solid')
    
    ws.row_dimensions[3].height = 28
    ws.row_dimensions[4].height = 32
    
    # Row 5: Separator
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=num_cols)
    sep_cell = ws.cell(row=5, column=1, value="")
    sep_cell.fill = PatternFill(start_color=COLORS['primary'], end_color=COLORS['primary'], fill_type='solid')
    ws.row_dimensions[5].height = 8
    
    # Row 6: Column Headers
    header_row = 6
    ws.row_dimensions[header_row].height = 32
    for c_idx, h_text in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=c_idx, value=h_text)
        cell.font = Font(name='Segoe UI', size=10, bold=True, color=COLORS['text_light'])
        cell.fill = PatternFill(start_color=COLORS['primary'], end_color=COLORS['primary'], fill_type='solid')
        cell.alignment = center_align
        cell.border = thin_border
    
    # Row 7: Filter info
    ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=num_cols)
    filter_cell = ws.cell(row=7, column=1, value="🔍 Filter: Vehicles with mileage between 20-24 hours | Developed by Muhammad Ashaan")
    filter_cell.font = Font(name='Segoe UI', size=9, italic=True, color=COLORS['primary_dark'])
    filter_cell.fill = PatternFill(start_color='EBF4FC', end_color='EBF4FC', fill_type='solid')
    filter_cell.alignment = center_align
    ws.row_dimensions[7].height = 24
    
    # Return styling configuration for data rows
    return {
        'thin_border': thin_border,
        'center_align': center_align,
        'header_font': header_font,
        'COLORS': COLORS
    }

# ------------------------------------------------------------------------------
# CITY SELECTION & FILE UPLOAD SECTION
# ------------------------------------------------------------------------------
st.markdown("### 📍 City & File Selection")

c_col1, c_col2 = st.columns([1, 2])

with c_col1:
    city_option = st.selectbox(
        "Select Report City:",
        options=["Kamoke", "Nowshera Virkan", "Gujranwala", "Other (Custom)"]
    )
    if city_option == "Other (Custom)":
        selected_city = st.text_input("Enter City Name:", value="Kamoke").strip().upper()
    else:
        selected_city = city_option.upper()

col_up1, col_up2 = st.columns(2)

with col_up1:
    current_file = st.file_uploader("1️⃣ Upload TODAY'S Mileage Report (.xlsx)", type=["xlsx"])

with col_up2:
    previous_file = st.file_uploader("2️⃣ Upload PREVIOUS Mileage Report (.xlsx) - Optional", type=["xlsx"])

with st.expander("⚙️ Optional: Master List Override"):
    rural_override_file = st.file_uploader("Upload New Master List (.xlsx)", type=["xlsx"])

if current_file:
    try:
        df_current_raw = pd.read_excel(current_file, header=None)
        h_idx, c_reg_idx, c_per_idx = process_mileage_df(df_current_raw)
        
        if h_idx == -1:
            st.error("❌ Could not auto-detect 'Reg#' and 'Period' columns in Today's Mileage Report.")
        else:
            # Load Master List
            rural_set = set(DEFAULT_RURAL_LIST)
            numeric_rural_set = {extract_digits(x) for x in DEFAULT_RURAL_LIST if extract_digits(x)}
            normalized_rural_set = {normalize_reg(x) for x in DEFAULT_RURAL_LIST if normalize_reg(x)}

            if rural_override_file:
                df_rural_raw = pd.read_excel(rural_override_file)
                rural_set.clear()
                numeric_rural_set.clear()
                normalized_rural_set.clear()
                for col in df_rural_raw.columns:
                    for val in df_rural_raw[col].dropna():
                        val_str = str(val).replace('\xa0', '').strip().upper()
                        if val_str and val_str not in ['TROLLY NO.', 'UC NO.']:
                            rural_set.add(val_str)
                            normalized_rural_set.add(normalize_reg(val_str))
                            digits = extract_digits(val_str)
                            if digits == val_str and digits != '':
                                numeric_rural_set.add(digits)

            # Previous Report Comparison
            prev_vehicles = set()
            if previous_file:
                df_prev_raw = pd.read_excel(previous_file, header=None)
                p_h_idx, p_reg_idx, p_per_idx = process_mileage_df(df_prev_raw)
                if p_h_idx != -1:
                    df_prev_data = df_prev_raw.iloc[p_h_idx + 1:].copy()
                    for _, p_row in df_prev_data.iterrows():
                        p_hrs = parse_period_to_hours(p_row.iloc[p_per_idx])
                        if 20.0 <= p_hrs <= 24.0:
                            p_reg = str(p_row.iloc[p_reg_idx]).replace('\xa0', '').strip()
                            if p_reg and p_reg not in ['-', 'NAN', 'NONE']:
                                prev_vehicles.add(normalize_reg(p_reg))

            # Extract Actual Date
            meta_date = extract_file_date(df_current_raw)

            raw_headers = df_current_raw.iloc[h_idx].astype(str).str.replace('\xa0', '', regex=False).str.strip().tolist()
            has_sno = raw_headers[0].upper() in ['S.NO', 'SR.#', 'SR NO', 'S#', 'NO']
            headers = raw_headers if has_sno else ['S.No'] + raw_headers

            df_data = df_current_raw.iloc[h_idx + 1:].copy()
            
            urban_rows, rural_rows = [], []
            reg_counts = {}
            missing_reg_rows = []
            repeated_vehicles = set()
            
            for idx, row in df_data.iterrows():
                period_val = row.iloc[c_per_idx]
                hrs = parse_period_to_hours(period_val)
                
                if 20.0 <= hrs <= 24.0:
                    reg_val = str(row.iloc[c_reg_idx]).replace('\xa0', '').strip() if pd.notna(row.iloc[c_reg_idx]) else ''
                    
                    if not reg_val or reg_val in ['-', 'NAN', 'NONE', ''] or len(reg_val) < 2:
                        missing_reg_rows.append(row.tolist())
                    else:
                        reg_clean = reg_val.upper()
                        reg_norm = normalize_reg(reg_clean)
                        reg_counts[reg_clean] = reg_counts.get(reg_clean, 0) + 1
                        
                        if reg_norm in prev_vehicles:
                            repeated_vehicles.add(reg_clean)
                        
                        row_list = row.tolist()
                        if not has_sno:
                            row_list = [''] + row_list
                        
                        if is_rural_vehicle(reg_clean, rural_set, numeric_rural_set, normalized_rural_set):
                            rural_rows.append(row_list)
                        else:
                            urban_rows.append(row_list)
            
            duplicate_regs = {k for k, v in reg_counts.items() if v > 1}
            total_valid_vehicles = len(urban_rows) + len(rural_rows)

            # --- RECEIPT METRICS SUMMARY ---
            st.markdown("### 📊 Audit Summary")
            st.info(f"📍 Selected Location: **{selected_city}** | 📅 File Data Date: **{meta_date}**")
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Filtered (20–24h)", total_valid_vehicles)
            m2.metric("Urban Fleet", len(urban_rows))
            m3.metric("Rural Fleet", len(rural_rows))
            m4.metric("Repeated 2-Day Vehicles", len(repeated_vehicles))

            if repeated_vehicles:
                st.info(f"🔁 **2-Day Consecutive 20–24h Repeat Vehicles ({len(repeated_vehicles)}):** {', '.join(repeated_vehicles)}")
            if duplicate_regs:
                st.warning(f"⚠️ **Same-Day Duplicates ({len(duplicate_regs)}):** {', '.join(duplicate_regs)}")

            # ------------------------------------------------------------------
            # ENHANCED EXCEL REPORT BUILDER
            # ------------------------------------------------------------------
            wb = openpyxl.Workbook()
            ws_out = wb.active
            ws_out.title = "Executive Report"
            
            # Apply professional styling
            num_cols = len(headers)
            style_config = apply_professional_styling(
                ws_out, headers, num_cols, 
                len(urban_rows), len(rural_rows), 
                total_valid_vehicles, meta_date, selected_city
            )
            
            thin_border = style_config['thin_border']
            center_align = style_config['center_align']
            COLORS = style_config['COLORS']
            
            # Data rows start from row 8
            curr_r = 8
            sno_tracker = 1
            
            # Fill patterns
            fill_alt = PatternFill(start_color=COLORS['alt_row'], end_color=COLORS['alt_row'], fill_type='solid')
            fill_dup = PatternFill(start_color='FFF3E0', end_color='FFF3E0', fill_type='solid')  # Light orange
            fill_repeat = PatternFill(start_color='FFE5E5', end_color='FFE5E5', fill_type='solid')  # Light red
            
            actual_reg_col = c_reg_idx + (0 if has_sno else 1)

            # --- URBAN FLEET SECTION ---
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=num_cols)
            u_sec = ws_out.cell(row=curr_r, column=1, value=f"🏙️ URBAN FLEET VEHICLES ({len(urban_rows)})")
            u_sec.font = Font(name='Segoe UI', size=12, bold=True, color=COLORS['text_light'])
            u_sec.fill = PatternFill(start_color=COLORS['primary'], end_color=COLORS['primary'], fill_type='solid')
            u_sec.alignment = center_align
            ws_out.row_dimensions[curr_r].height = 30
            curr_r += 1

            for r_data in urban_rows:
                reg_val = str(r_data[actual_reg_col]).replace('\xa0', '').strip() if pd.notna(r_data[actual_reg_col]) else ''
                reg_clean = reg_val.upper()
                is_even = (curr_r % 2 == 0)
                
                r_data[0] = sno_tracker
                sno_tracker += 1
                
                ws_out.row_dimensions[curr_r].height = 22
                for c_idx, val in enumerate(r_data, 1):
                    val_clean = str(val).replace('\xa0', '').strip() if pd.notna(val) else ''
                    cell = ws_out.cell(row=curr_r, column=c_idx, value=val_clean if c_idx > 1 else r_data[0])
                    cell.alignment = center_align
                    cell.border = thin_border
                    cell.font = Font(name='Segoe UI', size=10, color=COLORS['text_dark'])
                    
                    if is_even and c_idx > 0:
                        cell.fill = fill_alt
                    
                    if c_idx == (actual_reg_col + 1):
                        if reg_clean in repeated_vehicles:
                            cell.fill = fill_repeat
                            cell.font = Font(name='Segoe UI', size=10, bold=True, color=COLORS['accent_red'])
                        elif reg_clean in duplicate_regs:
                            cell.fill = fill_dup
                            cell.font = Font(name='Segoe UI', size=10, bold=True, color='E65100')
                curr_r += 1

            curr_r += 1

            # --- RURAL FLEET SECTION ---
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=num_cols)
            r_sec = ws_out.cell(row=curr_r, column=1, value=f"🌾 RURAL FLEET VEHICLES ({len(rural_rows)})")
            r_sec.font = Font(name='Segoe UI', size=12, bold=True, color=COLORS['text_light'])
            r_sec.fill = PatternFill(start_color=COLORS['accent_green'], end_color=COLORS['accent_green'], fill_type='solid')
            r_sec.alignment = center_align
            ws_out.row_dimensions[curr_r].height = 30
            curr_r += 1

            for r_data in rural_rows:
                reg_val = str(r_data[actual_reg_col]).replace('\xa0', '').strip() if pd.notna(r_data[actual_reg_col]) else ''
                reg_clean = reg_val.upper()
                is_even = (curr_r % 2 == 0)
                
                r_data[0] = sno_tracker
                sno_tracker += 1
                
                ws_out.row_dimensions[curr_r].height = 22
                for c_idx, val in enumerate(r_data, 1):
                    val_clean = str(val).replace('\xa0', '').strip() if pd.notna(val) else ''
                    cell = ws_out.cell(row=curr_r, column=c_idx, value=val_clean if c_idx > 1 else r_data[0])
                    cell.alignment = center_align
                    cell.border = thin_border
                    cell.font = Font(name='Segoe UI', size=10, color=COLORS['text_dark'])
                    
                    if is_even and c_idx > 0:
                        cell.fill = fill_alt
                    
                    if c_idx == (actual_reg_col + 1):
                        if reg_clean in repeated_vehicles:
                            cell.fill = fill_repeat
                            cell.font = Font(name='Segoe UI', size=10, bold=True, color=COLORS['accent_red'])
                        elif reg_clean in duplicate_regs:
                            cell.fill = fill_dup
                            cell.font = Font(name='Segoe UI', size=10, bold=True, color='E65100')
                curr_r += 1

            # Auto Column Widths
            for col in ws_out.columns:
                max_len = max(len(str(cell.value or '')) for cell in col if cell.value)
                col_letter = get_column_letter(col[0].column)
                ws_out.column_dimensions[col_letter].width = max(max_len + 5, 15)

            # Add footer with branding
            footer_row = curr_r + 1
            ws_out.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=num_cols)
            footer_cell = ws_out.cell(row=footer_row, column=1, value="📋 Report generated by Bellanix Tech - Mileage Executive Audit System")
            footer_cell.font = Font(name='Segoe UI', size=9, italic=True, color=COLORS['primary_light'])
            footer_cell.fill = PatternFill(start_color='EBF4FC', end_color='EBF4FC', fill_type='solid')
            footer_cell.alignment = center_align
            ws_out.row_dimensions[footer_row].height = 24

            # Download Output
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success("🎉 Executive Audit Report Generated Successfully!")
            st.download_button(
                label=f"📥 Download {selected_city} Executive Report (.xlsx)",
                data=output_buffer,
                file_name=f"{selected_city}_Mileage_Report_{meta_date.replace('/', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing report: {str(e)}")
