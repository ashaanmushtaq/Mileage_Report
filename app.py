import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re
from datetime import datetime

# ------------------------------------------------------------------------------
# DEFAULT RURAL VEHICLE MASTER LIST (UPDATED AS PER YOUR REPORT)
# ------------------------------------------------------------------------------
DEFAULT_RURAL_LIST = [
    "GAB-4046", "SA-6766", "STR-6637", "MC-5", "TT-06", "GTD-694", 
    "GAS-1694", "BN-3932", "DGK-1763", "TT-11", "GAU-8135", "GAJ-19-47", 
    "SAA-2946", "ST-663", "GAJ-3943", "TT-05", "OKA-2192", "SAA-7988",
    "R-1", "R-2", "R-3", "R-4", "R-5", "R-6", "R-7", "R-8", "R-9", "R-10",
    "R-11", "R-12", "R-13", "R-14", "R-15", "R-16", "R-17", "R-18", "R-19",
    "R-21", "R-23", "R-24", "R-32", "R-33", "R-35", "R-40", "R-41",
    "RIC-20", "RIC-22", "RIC-25", "RIC-26", "RIC-27", "RIC-28", "RIC-29",
    "RIC-30", "RIC-31", "RIC-34", "RIC-36", "RIC-37", "RIC-38", "RIC-39",
    "RIC-42", "RIC-43", "RIC-44"
]

st.set_page_config(
    page_title="Mileage Processor | Dev by Ashaan",
    page_icon="🧾",
    layout="wide"
)

# ------------------------------------------------------------------------------
# AESTHETIC UI STYLING
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
# MAIN APP
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

            # --- DISPLAY SUMMARY ---
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
            # EXCEL REPORT BUILDER - COMPLETELY ERROR FREE
            # ------------------------------------------------------------------
            wb = openpyxl.Workbook()
            ws_out = wb.active
            ws_out.title = "Executive Report"
            
            num_cols = len(headers)
            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            # Color Palette
            COLORS = {
                'primary_dark': '1C3D5A',
                'primary': '2A5C8A',
                'primary_light': '4A8CC4',
                'accent_gold': 'C5A45D',
                'accent_green': '2E7D32',
                'accent_red': 'C62828',
                'header_bg': '1A365D',
                'alt_row': 'F7FAFC',
                'white': 'FFFFFF',
                'border': 'CBD5E0',
                'text_dark': '2D3748',
                'text_light': 'FFFFFF'
            }
            
            thin_border = Border(
                left=Side(style='thin', color=COLORS['border']),
                right=Side(style='thin', color=COLORS['border']),
                top=Side(style='thin', color=COLORS['border']),
                bottom=Side(style='thin', color=COLORS['border'])
            )
            
            current_row = 1
            
            # === ROW 1: TITLE ===
            ws_out.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_cols)
            cell = ws_out.cell(row=current_row, column=1)
            cell.value = f"🏢 {selected_city} VEHICLE MILEAGE EXECUTIVE REPORT"
            cell.font = Font(name='Segoe UI', size=18, bold=True, color=COLORS['text_light'])
            cell.fill = PatternFill(start_color=COLORS['header_bg'], end_color=COLORS['header_bg'], fill_type='solid')
            cell.alignment = center_align
            ws_out.row_dimensions[current_row].height = 45
            current_row += 1
            
            # === ROW 2: SUBTITLE ===
            ws_out.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_cols)
            cell = ws_out.cell(row=current_row, column=1)
            cell.value = f"📅 Report Date: {meta_date} | 📊 Audit ID: MILE-{datetime.now().strftime('%Y%m%d')}"
            cell.font = Font(name='Segoe UI', size=10, color=COLORS['primary_light'])
            cell.fill = PatternFill(start_color='EBF4FC', end_color='EBF4FC', fill_type='solid')
            cell.alignment = center_align
            ws_out.row_dimensions[current_row].height = 28
            current_row += 1
            
            # === ROWS 3-4: METRICS ===
            metrics_data = [
                ('Total Fleet', total_valid_vehicles, COLORS['primary']),
                ('Urban Fleet', len(urban_rows), COLORS['primary_light']),
                ('Rural Fleet', len(rural_rows), COLORS['accent_green']),
                ('Completion %', f"{round((total_valid_vehicles/max(1,len(urban_rows)+len(rural_rows)))*100)}%", COLORS['accent_gold'])
            ]
            
            col_width = max(1, num_cols // 4)
            for i, (label, value, color) in enumerate(metrics_data):
                start_col = i * col_width + 1
                end_col = min((i + 1) * col_width, num_cols)
                
                ws_out.merge_cells(start_row=current_row, start_column=start_col, end_row=current_row+1, end_column=end_col)
                
                # Label
                cell = ws_out.cell(row=current_row, column=start_col)
                cell.value = label
                cell.font = Font(name='Segoe UI', size=9, bold=True, color=COLORS['text_dark'])
                cell.alignment = center_align
                cell.fill = PatternFill(start_color='F7FAFC', end_color='F7FAFC', fill_type='solid')
                
                # Value
                cell = ws_out.cell(row=current_row+1, column=start_col)
                cell.value = str(value)
                cell.font = Font(name='Segoe UI', size=16, bold=True, color=color)
                cell.alignment = center_align
                cell.fill = PatternFill(start_color='F7FAFC', end_color='F7FAFC', fill_type='solid')
            
            ws_out.row_dimensions[current_row].height = 28
            ws_out.row_dimensions[current_row+1].height = 32
            current_row += 2
            
            # === ROW 5: SEPARATOR ===
            ws_out.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_cols)
            cell = ws_out.cell(row=current_row, column=1)
            cell.value = ""
            cell.fill = PatternFill(start_color=COLORS['primary'], end_color=COLORS['primary'], fill_type='solid')
            ws_out.row_dimensions[current_row].height = 8
            current_row += 1
            
            # === ROW 6: HEADERS ===
            header_row = current_row
            ws_out.row_dimensions[header_row].height = 32
            for c_idx, h_text in enumerate(headers, 1):
                cell = ws_out.cell(row=header_row, column=c_idx)
                cell.value = h_text
                cell.font = Font(name='Segoe UI', size=10, bold=True, color=COLORS['text_light'])
                cell.fill = PatternFill(start_color=COLORS['primary'], end_color=COLORS['primary'], fill_type='solid')
                cell.alignment = center_align
                cell.border = thin_border
            
            current_row += 1
            
            # === ROW 7: FILTER INFO ===
            ws_out.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_cols)
            cell = ws_out.cell(row=current_row, column=1)
            cell.value = "🔍 Filter: Vehicles with mileage between 20-24 hours | Developed by Muhammad Ashaan"
            cell.font = Font(name='Segoe UI', size=9, italic=True, color=COLORS['primary_dark'])
            cell.fill = PatternFill(start_color='EBF4FC', end_color='EBF4FC', fill_type='solid')
            cell.alignment = center_align
            ws_out.row_dimensions[current_row].height = 24
            current_row += 1
            
            # === DATA ROWS ===
            sno_tracker = 1
            actual_reg_col = c_reg_idx + (0 if has_sno else 1)
            
            fill_alt = PatternFill(start_color=COLORS['alt_row'], end_color=COLORS['alt_row'], fill_type='solid')
            fill_dup = PatternFill(start_color='FFF3E0', end_color='FFF3E0', fill_type='solid')
            fill_repeat = PatternFill(start_color='FFE5E5', end_color='FFE5E5', fill_type='solid')
            
            # --- URBAN FLEET ---
            if urban_rows:
                ws_out.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_cols)
                cell = ws_out.cell(row=current_row, column=1)
                cell.value = f"🏙️ URBAN FLEET VEHICLES ({len(urban_rows)})"
                cell.font = Font(name='Segoe UI', size=12, bold=True, color=COLORS['text_light'])
                cell.fill = PatternFill(start_color=COLORS['primary'], end_color=COLORS['primary'], fill_type='solid')
                cell.alignment = center_align
                ws_out.row_dimensions[current_row].height = 30
                current_row += 1
                
                for r_data in urban_rows:
                    reg_val = str(r_data[actual_reg_col]).replace('\xa0', '').strip() if pd.notna(r_data[actual_reg_col]) else ''
                    reg_clean = reg_val.upper()
                    is_even = (current_row % 2 == 0)
                    
                    r_data[0] = sno_tracker
                    sno_tracker += 1
                    
                    ws_out.row_dimensions[current_row].height = 22
                    for c_idx, val in enumerate(r_data, 1):
                        cell = ws_out.cell(row=current_row, column=c_idx)
                        cell.value = val if c_idx > 1 else r_data[0]
                        cell.alignment = center_align
                        cell.border = thin_border
                        cell.font = Font(name='Segoe UI', size=10, color=COLORS['text_dark'])
                        
                        if is_even:
                            cell.fill = fill_alt
                        
                        if c_idx == (actual_reg_col + 1):
                            if reg_clean in repeated_vehicles:
                                cell.fill = fill_repeat
                                cell.font = Font(name='Segoe UI', size=10, bold=True, color=COLORS['accent_red'])
                            elif reg_clean in duplicate_regs:
                                cell.fill = fill_dup
                                cell.font = Font(name='Segoe UI', size=10, bold=True, color='E65100')
                    current_row += 1
                
                current_row += 1
            
            # --- RURAL FLEET ---
            if rural_rows:
                ws_out.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_cols)
                cell = ws_out.cell(row=current_row, column=1)
                cell.value = f"🌾 RURAL FLEET VEHICLES ({len(rural_rows)})"
                cell.font = Font(name='Segoe UI', size=12, bold=True, color=COLORS['text_light'])
                cell.fill = PatternFill(start_color=COLORS['accent_green'], end_color=COLORS['accent_green'], fill_type='solid')
                cell.alignment = center_align
                ws_out.row_dimensions[current_row].height = 30
                current_row += 1
                
                for r_data in rural_rows:
                    reg_val = str(r_data[actual_reg_col]).replace('\xa0', '').strip() if pd.notna(r_data[actual_reg_col]) else ''
                    reg_clean = reg_val.upper()
                    is_even = (current_row % 2 == 0)
                    
                    r_data[0] = sno_tracker
                    sno_tracker += 1
                    
                    ws_out.row_dimensions[current_row].height = 22
                    for c_idx, val in enumerate(r_data, 1):
                        cell = ws_out.cell(row=current_row, column=c_idx)
                        cell.value = val if c_idx > 1 else r_data[0]
                        cell.alignment = center_align
                        cell.border = thin_border
                        cell.font = Font(name='Segoe UI', size=10, color=COLORS['text_dark'])
                        
                        if is_even:
                            cell.fill = fill_alt
                        
                        if c_idx == (actual_reg_col + 1):
                            if reg_clean in repeated_vehicles:
                                cell.fill = fill_repeat
                                cell.font = Font(name='Segoe UI', size=10, bold=True, color=COLORS['accent_red'])
                            elif reg_clean in duplicate_regs:
                                cell.fill = fill_dup
                                cell.font = Font(name='Segoe UI', size=10, bold=True, color='E65100')
                    current_row += 1
                
                current_row += 1
            
            # === FOOTER ===
            ws_out.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_cols)
            cell = ws_out.cell(row=current_row, column=1)
            cell.value = "📋 Report generated by Bellanix Tech - Mileage Executive Audit System"
            cell.font = Font(name='Segoe UI', size=9, italic=True, color=COLORS['primary_light'])
            cell.fill = PatternFill(start_color='EBF4FC', end_color='EBF4FC', fill_type='solid')
            cell.alignment = center_align
            ws_out.row_dimensions[current_row].height = 24
            
            # === AUTO COLUMN WIDTHS ===
            for col in ws_out.columns:
                max_len = 0
                for cell in col:
                    if cell.value is not None:
                        max_len = max(max_len, len(str(cell.value)))
                col_letter = get_column_letter(col[0].column)
                ws_out.column_dimensions[col_letter].width = max(max_len + 5, 15)
            
            # === SAVE AND DOWNLOAD ===
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
        st.error(f"❌ Error: {str(e)}")
        st.error("Please check your file format and try again.")
