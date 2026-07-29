import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
import io
import re
from datetime import datetime

# ------------------------------------------------------------------------------
# DEFAULT RURAL VEHICLE MASTER LIST
# ------------------------------------------------------------------------------
DEFAULT_RURAL_LIST = [
    "GAB-4046", "SA-6766", "STR-6637", "MC-5", "TT-06", "GTD-694", 
    "GAS-1694", "BN-3932", "DGK-1763", "TT-11", "GAU-8135", "GAJ-19-47", 
    "SAA-2946", "ST-663", "GAJ-3943", "TT-05", "OKA-2192", "SAA-7988",
    "R-1", "R-2", "R-3", "R-4", "R-5", "R-6", "R-7", "R-8", "R-9", "R-10",
    "RIC-1", "RIC-2", "RIC-3", "RIC-4", "RIC-5", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
]

st.set_page_config(
    page_title="Mileage Processor | Dev by Ashaan",
    page_icon="🧾",
    layout="wide"
)

# ------------------------------------------------------------------------------
# CLEAN UI STYLING - Waisa hi jaise pehle tha
# ------------------------------------------------------------------------------
st.markdown("""
    <style>
    .receipt-card {
        background-color: #FFFFFF;
        border: 2px dashed #1F4E79;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        margin-bottom: 20px;
    }
    .receipt-header {
        font-family: 'Courier New', Courier, monospace;
        font-size: 26px;
        font-weight: bold;
        color: #1F4E79;
        text-align: center;
        letter-spacing: 2px;
        margin-bottom: 5px;
    }
    .receipt-sub {
        font-family: 'Calibri', sans-serif;
        text-align: center;
        color: #555;
        font-size: 14px;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="receipt-card">
    <div class="receipt-header">🧾 MILEAGE EXECUTIVE AUDIT DASHBOARD</div>
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
# CITY SELECTION & FILE UPLOAD
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

            # Previous Report Comparison - 2 Day Alert System
            prev_vehicles = set()
            prev_vehicle_details = {}  # Store previous vehicle data for comparison
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
                                reg_norm = normalize_reg(p_reg)
                                prev_vehicles.add(reg_norm)
                                prev_vehicle_details[reg_norm] = {
                                    'reg': p_reg,
                                    'hours': p_hrs,
                                    'row_data': p_row.tolist()
                                }

            # Extract Actual Date
            meta_date = extract_file_date(df_current_raw)

            raw_headers = df_current_raw.iloc[h_idx].astype(str).str.replace('\xa0', '', regex=False).str.strip().tolist()
            has_sno = raw_headers[0].upper() in ['S.NO', 'SR.#', 'SR NO', 'S#', 'NO']
            headers = raw_headers if has_sno else ['S.No'] + raw_headers
            # Add Alert Column
            headers = headers + ['⚠️ 2-DAY ALERT'] if previous_file else headers

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
                        
                        row_list = row.tolist()
                        if not has_sno:
                            row_list = [''] + row_list
                        
                        # Check if vehicle was also in previous report (2-day alert)
                        is_2day_alert = False
                        if previous_file and reg_norm in prev_vehicles:
                            repeated_vehicles.add(reg_clean)
                            is_2day_alert = True
                            row_list.append('⚠️ 2-DAY REPEAT - URGENT REVIEW')
                        else:
                            row_list.append('✅ CLEAR')
                        
                        if is_rural_vehicle(reg_clean, rural_set, numeric_rural_set, normalized_rural_set):
                            rural_rows.append(row_list)
                        else:
                            urban_rows.append(row_list)
            
            duplicate_regs = {k for k, v in reg_counts.items() if v > 1}
            total_valid_vehicles = len(urban_rows) + len(rural_rows)

            # --- METRICS SUMMARY ---
            st.markdown("### 📊 Audit Summary")
            st.info(f"📍 Selected Location: **{selected_city}** | 📅 File Data Date: **{meta_date}**")
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Filtered (20–24h)", total_valid_vehicles)
            m2.metric("Urban Fleet", len(urban_rows))
            m3.metric("Rural Fleet", len(rural_rows))
            m4.metric("⚠️ 2-Day Alerts", len(repeated_vehicles))

            if repeated_vehicles:
                st.error(f"🔴 **CRITICAL ALERT - {len(repeated_vehicles)} Vehicles Repeated 2 Days Consecutively:** {', '.join(repeated_vehicles)}")
            else:
                st.success("✅ No 2-Day repeated vehicles detected. All vehicles are performing optimally!")
            
            if duplicate_regs:
                st.warning(f"⚠️ **Same-Day Duplicates ({len(duplicate_regs)}):** {', '.join(duplicate_regs)}")

            # ------------------------------------------------------------------
            # AESTHETIC EXCEL REPORT BUILDER - PREMIUM DESIGN
            # ------------------------------------------------------------------
            wb = openpyxl.Workbook()
            ws_out = wb.active
            ws_out.title = "Executive Report"
            ws_out.views.sheetView[0].showGridLines = False

            center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
            num_cols = len(headers)
            half_cols = max(2, num_cols // 2)

            # 1. MAIN TITLE - Premium Gradient Header
            ws_out.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
            title_text = f"⭐ {selected_city} VEHICLE MILEAGE EXECUTIVE REPORT ⭐"
            title_cell = ws_out.cell(row=1, column=1, value=title_text)
            title_cell.font = Font(name="Arial", size=20, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            title_cell.alignment = center_align
            ws_out.row_dimensions[1].height = 45

            # 2. SUBTITLE
            ws_out.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
            sub_cell = ws_out.cell(row=2, column=1, value="◆ POWERED BY BELLANIX TECH ◆ AUTOMATED AUDIT SYSTEM ◆")
            sub_cell.font = Font(name="Arial", size=11, italic=True, color="7F8C8D")
            sub_cell.alignment = center_align
            ws_out.row_dimensions[2].height = 22

            # 3. METADATA BLOCK - Premium Styling
            metadata_colors = [
                PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid"),
                PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid"),
                PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid"),
                PatternFill(start_color="F5EEF8", end_color="F5EEF8", fill_type="solid")
            ]
            
            meta_data = [
                f"📅 Report Date: {meta_date}",
                f"🏆 Total Fleet (20-24h): {total_valid_vehicles}",
                f"👨‍💻 Developed by: Muhammad Ashaan",
                f"🏙️ Urban Fleet: {len(urban_rows)}",
                f"⏱️ Filtered Hours: 20–24 Hours",
                f"🌾 Rural Fleet: {len(rural_rows)}",
                f"🔴 2-Day Alerts: {len(repeated_vehicles)}",
                f"📌 Status: {'CRITICAL' if repeated_vehicles else 'CLEAR'}"
            ]
            
            start_row = 4
            for i, text in enumerate(meta_data):
                row = start_row + (i // 2)
                col = 1 + (i % 2) * (half_cols)
                end_col = col + half_cols - 1
                
                ws_out.merge_cells(start_row=row, start_column=col, end_row=row, end_column=end_col)
                cell = ws_out.cell(row=row, column=col, value=text)
                cell.font = Font(bold=True, size=11, color="2C3E50" if i < 6 else ("C0392B" if repeated_vehicles else "27AE60"))
                cell.alignment = center_align
                cell.fill = metadata_colors[i % len(metadata_colors)]
                ws_out.row_dimensions[row].height = 24

            # 4. COLUMN HEADERS - Premium Gradient
            header_row = start_row + 4
            ws_out.row_dimensions[header_row].height = 32
            
            # Gradient colors for headers
            header_gradients = [
                "2471A3", "2980B9", "2E86C1", "3498DB", "5DADE2",
                "85C1E9", "AED6F1", "D4E6F1"
            ]
            
            for c_idx, h_text in enumerate(headers, 1):
                cell = ws_out.cell(row=header_row, column=c_idx, value=h_text)
                color_idx = (c_idx - 1) % len(header_gradients)
                cell.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color=header_gradients[color_idx], 
                                       end_color=header_gradients[color_idx], 
                                       fill_type="solid")
                cell.alignment = center_align
                cell.border = Border(
                    left=Side(style='medium', color='1F4E79'),
                    right=Side(style='medium', color='1F4E79'),
                    top=Side(style='medium', color='1F4E79'),
                    bottom=Side(style='medium', color='1F4E79')
                )

            # 5. DATA ROWS - With Aesthetic Styling
            curr_r = header_row + 1
            sno_tracker = 1
            
            # Border styles
            thin_border = Border(
                left=Side(style='thin', color='BDC3C7'),
                right=Side(style='thin', color='BDC3C7'),
                top=Side(style='thin', color='BDC3C7'),
                bottom=Side(style='thin', color='BDC3C7')
            )
            
            # Fill colors for alternating rows
            fill_light = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
            fill_dark = PatternFill(start_color="E8ECEF", end_color="E8ECEF", fill_type="solid")
            fill_alert = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")  # Red for alerts
            fill_dup = PatternFill(start_color="FCF3CF", end_color="FCF3CF", fill_type="solid")  # Yellow for duplicates

            actual_reg_col = c_reg_idx + (0 if has_sno else 1)

            # URBAN SECTION
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=num_cols)
            u_sec = ws_out.cell(row=curr_r, column=1, value=f"🏙️ URBAN FLEET VEHICLES ({len(urban_rows)})")
            u_sec.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
            u_sec.fill = PatternFill(start_color="1B4F72", end_color="1B4F72", fill_type="solid")
            u_sec.alignment = center_align
            ws_out.row_dimensions[curr_r].height = 32
            curr_r += 1

            for r_data in urban_rows:
                reg_val = str(r_data[actual_reg_col]).replace('\xa0', '').strip() if pd.notna(r_data[actual_reg_col]) else ''
                reg_clean = reg_val.upper()
                
                # Check if this row has an alert
                has_alert = previous_file and len(r_data) > num_cols - 1 and r_data[-1] != '✅ CLEAR'
                
                r_data[0] = sno_tracker
                sno_tracker += 1
                
                ws_out.row_dimensions[curr_r].height = 22
                is_even = (curr_r % 2 == 0)
                
                for c_idx, val in enumerate(r_data, 1):
                    val_clean = str(val).replace('\xa0', '').strip() if pd.notna(val) else ''
                    cell = ws_out.cell(row=curr_r, column=c_idx, value=val_clean if c_idx > 1 else r_data[0])
                    cell.alignment = center_align if c_idx != num_cols else left_align
                    cell.border = thin_border
                    
                    # Apply background colors
                    if has_alert and c_idx == num_cols:
                        cell.fill = fill_alert
                        cell.font = Font(bold=True, color="C0392B", size=11)
                    elif reg_clean in duplicate_regs:
                        cell.fill = fill_dup
                        cell.font = Font(bold=True, color="B7950B")
                    elif is_even:
                        cell.fill = fill_light
                    else:
                        cell.fill = fill_dark
                        
                curr_r += 1

            curr_r += 1

            # RURAL SECTION
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=num_cols)
            r_sec = ws_out.cell(row=curr_r, column=1, value=f"🌾 RURAL FLEET VEHICLES ({len(rural_rows)})")
            r_sec.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
            r_sec.fill = PatternFill(start_color="1E8449", end_color="1E8449", fill_type="solid")
            r_sec.alignment = center_align
            ws_out.row_dimensions[curr_r].height = 32
            curr_r += 1

            for r_data in rural_rows:
                reg_val = str(r_data[actual_reg_col]).replace('\xa0', '').strip() if pd.notna(r_data[actual_reg_col]) else ''
                reg_clean = reg_val.upper()
                
                has_alert = previous_file and len(r_data) > num_cols - 1 and r_data[-1] != '✅ CLEAR'
                
                r_data[0] = sno_tracker
                sno_tracker += 1
                
                ws_out.row_dimensions[curr_r].height = 22
                is_even = (curr_r % 2 == 0)
                
                for c_idx, val in enumerate(r_data, 1):
                    val_clean = str(val).replace('\xa0', '').strip() if pd.notna(val) else ''
                    cell = ws_out.cell(row=curr_r, column=c_idx, value=val_clean if c_idx > 1 else r_data[0])
                    cell.alignment = center_align if c_idx != num_cols else left_align
                    cell.border = thin_border
                    
                    if has_alert and c_idx == num_cols:
                        cell.fill = fill_alert
                        cell.font = Font(bold=True, color="C0392B", size=11)
                    elif reg_clean in duplicate_regs:
                        cell.fill = fill_dup
                        cell.font = Font(bold=True, color="B7950B")
                    elif is_even:
                        cell.fill = fill_light
                    else:
                        cell.fill = fill_dark
                        
                curr_r += 1

            # 6. SUMMARY FOOTER
            curr_r += 1
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=num_cols)
            footer_text = f"📊 REPORT GENERATED: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | STATUS: {'⚠️ 2-DAY REPEATS DETECTED' if repeated_vehicles else '✅ ALL CLEAR'}"
            footer_cell = ws_out.cell(row=curr_r, column=1, value=footer_text)
            footer_cell.font = Font(bold=True, size=11, color="FFFFFF" if repeated_vehicles else "27AE60")
            footer_cell.fill = PatternFill(start_color="C0392B" if repeated_vehicles else "27AE60", 
                                          end_color="C0392B" if repeated_vehicles else "27AE60", 
                                          fill_type="solid")
            footer_cell.alignment = center_align
            ws_out.row_dimensions[curr_r].height = 28

            # 7. AUTO COLUMN WIDTHS
            for col in ws_out.columns:
                max_len = 0
                for cell in col:
                    try:
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max(max_len + 3, 12), 40)
                col_letter = get_column_letter(col[0].column)
                ws_out.column_dimensions[col_letter].width = adjusted_width

            # Download
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success("🎉 Executive Audit Report Generated Successfully!")
            
            # Show preview of alert status
            if repeated_vehicles:
                st.warning(f"⚠️ **{len(repeated_vehicles)} vehicles flagged with 2-Day Repeat Alert!**")
            
            st.download_button(
                label=f"📥 Download {selected_city} Executive Report (.xlsx)",
                data=output_buffer,
                file_name=f"{selected_city}_Mileage_Report_{meta_date.replace('/', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing report: {str(e)}")
