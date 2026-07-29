import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
from datetime import datetime

# ------------------------------------------------------------------------------
# DEFAULT RURAL VEHICLE MASTER LIST (Permanent Storage)
# ------------------------------------------------------------------------------
DEFAULT_RURAL_LIST = [
    # Aap ki Rural List ke tamam numbers/codes yahan direct master set me saved hain
    "R-1", "R-2", "R-3", "R-4", "R-5", "R-6", "R-7", "R-8", "R-9", "R-10",
    "RIC-1", "RIC-2", "RIC-3", "RIC-4", "RIC-5", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
]

st.set_page_config(
    page_title="Vehicle Mileage Processor | Dev by Ashaan",
    page_icon="🚗",
    layout="wide"
)

# Custom Styling for Web Interface
st.markdown("""
    <style>
    .main-header {
        font-size: 30px;
        font-weight: bold;
        background: linear-gradient(90deg, #1F4E79, #2980B9);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin-bottom: 10px;
    }
    .sub-header {
        text-align: center;
        color: #555;
        font-size: 14px;
        margin-bottom: 25px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">⚡ Vehicle Mileage Automated Processor</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Developed by: <b>Muhammad Ashaan</b></div>', unsafe_allow_html=True)

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

def is_rural_vehicle(reg_str, rural_set, numeric_rural_set):
    clean_reg = str(reg_str).replace('\xa0', '').strip().upper()
    
    if clean_reg in ['MC-5', 'MC5']:
        return False
    if clean_reg.startswith('R-') or clean_reg.startswith('RIC-'):
        return True
    if clean_reg in rural_set:
        return True
    digits = extract_digits(clean_reg)
    if digits and digits in numeric_rural_set:
        return True
    return False

# ------------------------------------------------------------------------------
# FILE UPLOAD SECTION
# ------------------------------------------------------------------------------
mileage_file = st.file_uploader("📂 Upload Today's Mileage Report (.xlsx)", type=["xlsx"])

with st.expander("⚙️ Optional: Update Rural Master List (Agra Aap change karna chahein)"):
    rural_file = st.file_uploader("Upload New Master List (.xlsx)", type=["xlsx"], key="rural_master")

if mileage_file:
    try:
        df_mileage_raw = pd.read_excel(mileage_file, header=None)
        
        # Detect Headers
        header_row_idx, col_reg_idx, col_period_idx = -1, -1, -1
        for r_idx in range(min(15, len(df_mileage_raw))):
            row_vals = df_mileage_raw.iloc[r_idx].astype(str).str.replace('\xa0', '', regex=False).str.strip().str.upper()
            c_reg, c_per = -1, -1
            for c_idx, val in enumerate(row_vals):
                if val in ['REG#', 'REGISTRATION', 'VEHICLE NO', 'REGISTRATION NO', 'REG NO']:
                    c_reg = c_idx
                if val in ['PERIOD', 'DURATION', 'HOURS']:
                    c_per = c_idx
            if c_reg != -1 and c_per != -1:
                header_row_idx, col_reg_idx, col_period_idx = r_idx, c_reg, c_per
                break
        
        if header_row_idx == -1:
            st.error("❌ Could not auto-detect 'Reg#' and 'Period' columns in the uploaded Excel file.")
        else:
            # Load Rural Master Data (From Upload OR Saved Default)
            rural_set = set(DEFAULT_RURAL_LIST)
            numeric_rural_set = {extract_digits(x) for x in DEFAULT_RURAL_LIST if extract_digits(x)}
            
            if rural_file:
                df_rural_raw = pd.read_excel(rural_file)
                rural_set.clear()
                numeric_rural_set.clear()
                for col in df_rural_raw.columns:
                    for val in df_rural_raw[col].dropna():
                        val_str = str(val).replace('\xa0', '').strip().upper()
                        if val_str and val_str not in ['TROLLY NO.', 'UC NO.']:
                            rural_set.add(val_str)
                            digits = extract_digits(val_str)
                            if digits == val_str and digits != '':
                                numeric_rural_set.add(digits)

            # Metadata Date Extraction
            meta_date = datetime.now().strftime("%m/%d/%Y")
            top_cell_val = str(df_mileage_raw.iloc[0, 0])
            if '[' in top_cell_val and ']' in top_cell_val:
                p1, p2 = top_cell_val.find('['), top_cell_val.find(']')
                meta_date = top_cell_val[p1+1:p2].split('-')[0].strip()

            raw_headers = df_mileage_raw.iloc[header_row_idx].astype(str).str.replace('\xa0', '', regex=False).str.strip().tolist()
            
            # Check if S.No / Sr.# exists in headers, else prepare column list
            has_sno = raw_headers[0].upper() in ['S.NO', 'SR.#', 'SR NO', 'S#', 'NO']
            headers = raw_headers if has_sno else ['S.No'] + raw_headers

            df_data = df_mileage_raw.iloc[header_row_idx + 1:].copy()
            
            urban_rows, rural_rows = [], []
            reg_counts = {}
            missing_reg_rows = []
            
            for idx, row in df_data.iterrows():
                period_val = row.iloc[col_period_idx]
                hrs = parse_period_to_hours(period_val)
                
                # 20–24 Hours Filter
                if 20.0 <= hrs <= 24.0:
                    reg_val = str(row.iloc[col_reg_idx]).replace('\xa0', '').strip() if pd.notna(row.iloc[col_reg_idx]) else ''
                    
                    if not reg_val or reg_val in ['-', 'NAN', 'NONE', ''] or len(reg_val) < 2:
                        missing_reg_rows.append(row.tolist())
                    else:
                        reg_clean = reg_val.upper()
                        reg_counts[reg_clean] = reg_counts.get(reg_clean, 0) + 1
                        
                        row_list = row.tolist()
                        if not has_sno:
                            row_list = [''] + row_list # Placeholder for S.No
                        
                        if is_rural_vehicle(reg_clean, rural_set, numeric_rural_set):
                            rural_rows.append(row_list)
                        else:
                            urban_rows.append(row_list)
            
            duplicate_regs = {k for k, v in reg_counts.items() if v > 1}
            total_valid_vehicles = len(urban_rows) + len(rural_rows)

            # --- ON-SCREEN DASHBOARD METRICS ---
            st.markdown("### 📊 Summary Dashboard")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Filtered (20–24h)", total_valid_vehicles)
            m2.metric("Urban Fleet", len(urban_rows))
            m3.metric("Rural Fleet", len(rural_rows))
            m4.metric("Invalid/Missing Reg", len(missing_reg_rows))

            if duplicate_regs:
                st.warning(f"⚠️ **Duplicate Vehicle Registrations ({len(duplicate_regs)}):** {', '.join(duplicate_regs)}")
            if missing_reg_rows:
                st.error(f"🚨 **Attention:** Detected {len(missing_reg_rows)} record(s) with **Missing/Invalid Vehicle Registration**. Unassigned entries have been safely excluded.")

            # ------------------------------------------------------------------
            # EXCEL REPORT BUILDER (AESTHETIC & PERFECT CENTERING)
            # ------------------------------------------------------------------
            wb = openpyxl.Workbook()
            ws_out = wb.active
            ws_out.title = "Executive Report"
            ws_out.views.sheetView[0].showGridLines = True

            center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            left_align = Alignment(horizontal="left", vertical="center")
            
            # Title Banner
            ws_out.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
            title_cell = ws_out.cell(row=1, column=1, value="VEHICLE MILEAGE EXECUTIVE REPORT")
            title_cell.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            title_cell.alignment = center_align
            ws_out.row_dimensions[1].height = 36

            # Meta Cards Block
            ws_out.cell(row=3, column=1, value="Report Date:").font = Font(bold=True)
            ws_out.cell(row=3, column=1).alignment = left_align
            ws_out.cell(row=3, column=2, value=meta_date).alignment = center_align
            
            ws_out.cell(row=4, column=1, value="Developer:").font = Font(bold=True)
            ws_out.cell(row=4, column=1).alignment = left_align
            dev_cell = ws_out.cell(row=4, column=2, value="Muhammad Ashaan")
            dev_cell.font = Font(bold=True, color="1F4E79")
            dev_cell.alignment = center_align
            
            ws_out.cell(row=5, column=1, value="Filtered Hours:").font = Font(bold=True)
            ws_out.cell(row=5, column=1).alignment = left_align
            ws_out.cell(row=5, column=2, value="20–24 Hours").alignment = center_align

            ws_out.cell(row=3, column=4, value="Total Fleet:").font = Font(bold=True)
            ws_out.cell(row=3, column=5, value=total_valid_vehicles).font = Font(bold=True, color="1F4E79")
            ws_out.cell(row=4, column=4, value="Urban Fleet:").font = Font(bold=True)
            ws_out.cell(row=4, column=5, value=len(urban_rows)).font = Font(bold=True, color="1F4E79")
            ws_out.cell(row=5, column=4, value="Rural Fleet:").font = Font(bold=True)
            ws_out.cell(row=5, column=5, value=len(rural_rows)).font = Font(bold=True, color="1F4E79")
            
            for r in range(3, 6):
                ws_out.cell(row=r, column=4).alignment = left_align
                ws_out.cell(row=r, column=5).alignment = center_align

            # Main Column Headers
            header_row_out = 7
            ws_out.row_dimensions[header_row_out].height = 28
            for c_idx, h_text in enumerate(headers, 1):
                cell = ws_out.cell(row=header_row_out, column=c_idx, value=h_text)
                cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="2980B9", end_color="2980B9", fill_type="solid")
                cell.alignment = center_align

            curr_r = 8
            global_sno = 1  # Sequential Counter for S.No (1, 2, 3...)
            
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )
            fill_alt = PatternFill(start_color="F2F7FA", end_color="F2F7FA", fill_type="solid")
            fill_dup = PatternFill(start_color="FFE169", end_color="FFE169", fill_type="solid")

            # Function to Write Rows with Perfect Alignment & Continuous S.No
            def write_section(rows_data, section_title, section_color, start_r):
                nonlocal global_sno
                r_pos = start_r
                
                # Section Header Banner
                ws_out.merge_cells(start_row=r_pos, start_column=1, end_row=r_pos, end_column=len(headers))
                sec_cell = ws_out.cell(row=r_pos, column=1, value=f"{section_title} ({len(rows_data)})")
                sec_cell.font = Font(bold=True, color="FFFFFF", size=11)
                sec_cell.fill = PatternFill(start_color=section_color, end_color=section_color, fill_type="solid")
                sec_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                ws_out.row_dimensions[r_pos].height = 25
                r_pos += 1

                actual_reg_col = col_reg_idx + (0 if has_sno else 1)

                for r_data in rows_data:
                    reg_val = str(r_data[actual_reg_col]).replace('\xa0', '').strip() if pd.notna(r_data[actual_reg_col]) else ''
                    reg_clean = reg_val.upper()
                    is_even = (r_pos % 2 == 0)
                    
                    # Force Continuous Serial Numbering
                    r_data[0] = global_sno
                    global_sno += 1
                    
                    ws_out.row_dimensions[r_pos].height = 20
                    for c_idx, val in enumerate(r_data, 1):
                        val_clean = str(val).replace('\xa0', '').strip() if pd.notna(val) else ''
                        cell = ws_out.cell(row=r_pos, column=c_idx, value=val_clean if c_idx > 1 else r_data[0])
                        
                        # 100% CENTERED ALIGNMENT FOR ALL DATA
                        cell.alignment = center_align
                        cell.border = thin_border
                        
                        if is_even:
                            cell.fill = fill_alt
                        if c_idx == (actual_reg_col + 1) and reg_clean in duplicate_regs:
                            cell.fill = fill_dup
                            cell.font = Font(bold=True, color="9C5700")
                    r_pos += 1
                return r_pos

            # Write Urban Fleet
            curr_r = write_section(urban_rows, "URBAN FLEET VEHICLES", "154360", curr_r)
            curr_r += 1
            # Write Rural Fleet (S.No continues automatically)
            curr_r = write_section(rural_rows, "RURAL FLEET VEHICLES", "145A32", curr_r)

            # Auto Auto-fit Column Widths cleanly
            for col in ws_out.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_out.column_dimensions[col_letter].width = max(max_len + 5, 14)

            # Streamlit Download Button
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success("🎉 Executive Excel Report Generated Successfully!")
            st.download_button(
                label="📥 Download Final Executive Report (.xlsx)",
                data=output_buffer,
                file_name=f"Mileage_Report_{meta_date.replace('/', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing report: {str(e)}")
