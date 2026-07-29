import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
from datetime import datetime

# ------------------------------------------------------------------------------
# DEFAULT RURAL VEHICLE MASTER LIST (UPDATED WITH IMAGE DATA)
# ------------------------------------------------------------------------------
DEFAULT_RURAL_LIST = [
    # Image Master List Numbers
    "GAB-4046", "SA-6766", "STR-6637", "MC-5", "TT-06", "GTD-694", 
    "GAS-1694", "BN-3932", "DGK-1763", "TT-11", "GAU-8135", "GAJ-19-47", 
    "SAA-2946", "ST-663", "GAJ-3943", "TT-05", "OKA-2192", "SAA-7988",
    # General Standard Rules
    "R-1", "R-2", "R-3", "R-4", "R-5", "R-6", "R-7", "R-8", "R-9", "R-10",
    "RIC-1", "RIC-2", "RIC-3", "RIC-4", "RIC-5", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
]

st.set_page_config(
    page_title="Mileage Processor | Dev by Ashaan",
    page_icon="🧾",
    layout="wide"
)

# ------------------------------------------------------------------------------
# AESTHETIC RECEIPT UI CUSTOM STYLING
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
    .receipt-divider {
        border-top: 2px dashed #CCCCCC;
        margin: 15px 0;
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
    
    # Exact or Normalized Match
    if raw_clean in rural_set or norm_clean in normalized_rural_set:
        return True
    
    # Prefix Match
    if raw_clean.startswith('R-') or raw_clean.startswith('RIC-'):
        return True
        
    digits = extract_digits(raw_clean)
    if digits and digits in numeric_rural_set:
        return True
        
    return False

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
# FILE UPLODERS (CURRENT + PREVIOUS COMPARISON)
# ------------------------------------------------------------------------------
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
            # Prepare Master List Sets
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

            # Extract Previous Report Vehicles for Comparison
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

            # Date Extraction
            meta_date = datetime.now().strftime("%m/%d/%Y")
            top_cell_val = str(df_current_raw.iloc[0, 0])
            if '[' in top_cell_val and ']' in top_cell_val:
                p1, p2 = top_cell_val.find('['), top_cell_val.find(']')
                meta_date = top_cell_val[p1+1:p2].split('-')[0].strip()

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
            # EXCEL REPORT BUILDER (MERGED & CENTERED HEADERS)
            # ------------------------------------------------------------------
            wb = openpyxl.Workbook()
            ws_out = wb.active
            ws_out.title = "Executive Report"
            ws_out.views.sheetView[0].showGridLines = True

            center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            num_cols = len(headers)
            half_cols = max(2, num_cols // 2)

            # 1. Main Title Banner
            ws_out.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
            title_cell = ws_out.cell(row=1, column=1, value="VEHICLE MILEAGE EXECUTIVE REPORT")
            title_cell.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            title_cell.alignment = center_align
            ws_out.row_dimensions[1].height = 36

            # 2. Merged Centralized Metadata Block (2 or 4 Columns merged per block)
            # Row 3: Report Date & Total Fleet
            ws_out.merge_cells(start_row=3, start_column=1, end_row=3, end_column=half_cols)
            r3_left = ws_out.cell(row=3, column=1, value=f"Report Date:  {meta_date}")
            r3_left.font = Font(bold=True)
            r3_left.alignment = center_align

            ws_out.merge_cells(start_row=3, start_column=half_cols+1, end_row=3, end_column=num_cols)
            r3_right = ws_out.cell(row=3, column=half_cols+1, value=f"Total Fleet (20-24h):  {total_valid_vehicles}")
            r3_right.font = Font(bold=True, color="1F4E79")
            r3_right.alignment = center_align

            # Row 4: Developed By & Urban Fleet Count
            ws_out.merge_cells(start_row=4, start_column=1, end_row=4, end_column=half_cols)
            r4_left = ws_out.cell(row=4, column=1, value="Developed by:  Muhammad Ashaan")
            r4_left.font = Font(bold=True, color="1F4E79")
            r4_left.alignment = center_align

            ws_out.merge_cells(start_row=4, start_column=half_cols+1, end_row=4, end_column=num_cols)
            r4_right = ws_out.cell(row=4, column=half_cols+1, value=f"Urban Fleet Count:  {len(urban_rows)}")
            r4_right.font = Font(bold=True, color="154360")
            r4_right.alignment = center_align

            # Row 5: Filtered Hours & Rural Fleet Count
            ws_out.merge_cells(start_row=5, start_column=1, end_row=5, end_column=half_cols)
            r5_left = ws_out.cell(row=5, column=1, value="Filtered Hours:  20–24 Hours")
            r5_left.font = Font(bold=True)
            r5_left.alignment = center_align

            ws_out.merge_cells(start_row=5, start_column=half_cols+1, end_row=5, end_column=num_cols)
            r5_right = ws_out.cell(row=5, column=half_cols+1, value=f"Rural Fleet Count:  {len(rural_rows)}")
            r5_right.font = Font(bold=True, color="145A32")
            r5_right.alignment = center_align

            for r in range(3, 6):
                ws_out.row_dimensions[r].height = 22

            # 3. Table Column Headers
            header_row_out = 7
            ws_out.row_dimensions[header_row_out].height = 28
            for c_idx, h_text in enumerate(headers, 1):
                cell = ws_out.cell(row=header_row_out, column=c_idx, value=h_text)
                cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="2980B9", end_color="2980B9", fill_type="solid")
                cell.alignment = center_align

            curr_r = 8
            sno_tracker = 1
            
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )
            fill_alt = PatternFill(start_color="F2F7FA", end_color="F2F7FA", fill_type="solid")
            fill_dup = PatternFill(start_color="FFE169", end_color="FFE169", fill_type="solid")
            fill_repeat = PatternFill(start_color="FFCC80", end_color="FFCC80", fill_type="solid") # Orange highlight for 2-day repeat

            actual_reg_col = c_reg_idx + (0 if has_sno else 1)

            # --- URBAN FLEET SECTION (MERGED & CENTERED HEADER) ---
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=num_cols)
            u_sec = ws_out.cell(row=curr_r, column=1, value=f"URBAN FLEET VEHICLES ({len(urban_rows)})")
            u_sec.font = Font(bold=True, color="FFFFFF", size=11)
            u_sec.fill = PatternFill(start_color="154360", end_color="154360", fill_type="solid")
            u_sec.alignment = center_align # Center Alignment Fixed
            ws_out.row_dimensions[curr_r].height = 26
            curr_r += 1

            for r_data in urban_rows:
                reg_val = str(r_data[actual_reg_col]).replace('\xa0', '').strip() if pd.notna(r_data[actual_reg_col]) else ''
                reg_clean = reg_val.upper()
                is_even = (curr_r % 2 == 0)
                
                r_data[0] = sno_tracker
                sno_tracker += 1
                
                ws_out.row_dimensions[curr_r].height = 20
                for c_idx, val in enumerate(r_data, 1):
                    val_clean = str(val).replace('\xa0', '').strip() if pd.notna(val) else ''
                    cell = ws_out.cell(row=curr_r, column=c_idx, value=val_clean if c_idx > 1 else r_data[0])
                    cell.alignment = center_align
                    cell.border = thin_border
                    if is_even:
                        cell.fill = fill_alt
                    
                    # Highlights: 2-Day Repeat (Orange) vs Same-Day Duplicate (Yellow)
                    if c_idx == (actual_reg_col + 1):
                        if reg_clean in repeated_vehicles:
                            cell.fill = fill_repeat
                            cell.font = Font(bold=True, color="B7410E")
                        elif reg_clean in duplicate_regs:
                            cell.fill = fill_dup
                            cell.font = Font(bold=True, color="9C5700")
                curr_r += 1

            curr_r += 1

            # --- RURAL FLEET SECTION (MERGED & CENTERED HEADER) ---
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=num_cols)
            r_sec = ws_out.cell(row=curr_r, column=1, value=f"RURAL FLEET VEHICLES ({len(rural_rows)})")
            r_sec.font = Font(bold=True, color="FFFFFF", size=11)
            r_sec.fill = PatternFill(start_color="145A32", end_color="145A32", fill_type="solid")
            r_sec.alignment = center_align # Center Alignment Fixed
            ws_out.row_dimensions[curr_r].height = 26
            curr_r += 1

            for r_data in rural_rows:
                reg_val = str(r_data[actual_reg_col]).replace('\xa0', '').strip() if pd.notna(r_data[actual_reg_col]) else ''
                reg_clean = reg_val.upper()
                is_even = (curr_r % 2 == 0)
                
                r_data[0] = sno_tracker
                sno_tracker += 1
                
                ws_out.row_dimensions[curr_r].height = 20
                for c_idx, val in enumerate(r_data, 1):
                    val_clean = str(val).replace('\xa0', '').strip() if pd.notna(val) else ''
                    cell = ws_out.cell(row=curr_r, column=c_idx, value=val_clean if c_idx > 1 else r_data[0])
                    cell.alignment = center_align
                    cell.border = thin_border
                    if is_even:
                        cell.fill = fill_alt
                    
                    if c_idx == (actual_reg_col + 1):
                        if reg_clean in repeated_vehicles:
                            cell.fill = fill_repeat
                            cell.font = Font(bold=True, color="B7410E")
                        elif reg_clean in duplicate_regs:
                            cell.fill = fill_dup
                            cell.font = Font(bold=True, color="9C5700")
                curr_r += 1

            # Auto Column Widths
            for col in ws_out.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_out.column_dimensions[col_letter].width = max(max_len + 5, 14)

            # Download Output
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success("🎉 Final Receipt-Style Executive Report Generated Successfully!")
            st.download_button(
                label="📥 Download Executive Audit Report (.xlsx)",
                data=output_buffer,
                file_name=f"Mileage_Report_{meta_date.replace('/', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing report: {str(e)}")
