import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
from datetime import datetime

# ------------------------------------------------------------------------------
# PAGE CONFIGURATION & STYLING
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="Daily Vehicle Mileage Report Processor",
    page_icon="🚗",
    layout="wide"
)

st.markdown("""
    <style>
    .main-header {
        font-size: 28px;
        font-weight: bold;
        color: #1F4E79;
        text-align: center;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🚗 Daily Vehicle Mileage Report Processor</div>', unsafe_allow_html=True)
st.info("Upload today's **Mileage Report** and your **Rural Master List** below. The app will filter for **20–24 Hours**, classify Urban/Rural vehicles, highlight duplicates, and generate the final Executive Excel Report.")

# ------------------------------------------------------------------------------
# HELPER FUNCTIONS (100% ACCURATE FILTERING & CLASSIFICATION)
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
    
    # Rule 1: MC-5 is ALWAYS Urban
    if clean_reg in ['MC-5', 'MC5']:
        return False
    
    # Rule 2: Registration starts with R- or RIC-
    if clean_reg.startswith('R-') or clean_reg.startswith('RIC-'):
        return True
    
    # Rule 3: Exact match in Rural Master List
    if clean_reg in rural_set:
        return True
    
    # Rule 4: Match purely numeric master items (e.g. Master 5 matches R-5)
    digits = extract_digits(clean_reg)
    if digits and digits in numeric_rural_set:
        return True
    
    return False

# ------------------------------------------------------------------------------
# FILE UPLOAD UI
# ------------------------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    mileage_file = st.file_uploader("1️⃣ Upload Today's Mileage Report (.xlsx)", type=["xlsx"])

with col2:
    rural_file = st.file_uploader("2️⃣ Upload Rural Master List (.xlsx)", type=["xlsx"])

if mileage_file and rural_file:
    try:
        # Load Raw Mileage Report
        df_mileage_raw = pd.read_excel(mileage_file, header=None)
        
        # Auto-Detect Column Headers (Reg# and Period)
        header_row_idx = -1
        col_reg_idx = -1
        col_period_idx = -1
        
        for r_idx in range(min(15, len(df_mileage_raw))):
            row_vals = df_mileage_raw.iloc[r_idx].astype(str).str.replace('\xa0', '', regex=False).str.strip().str.upper()
            
            c_reg = -1
            c_per = -1
            for c_idx, val in enumerate(row_vals):
                if val in ['REG#', 'REGISTRATION', 'VEHICLE NO', 'REGISTRATION NO', 'REG NO']:
                    c_reg = c_idx
                if val in ['PERIOD', 'DURATION', 'HOURS']:
                    c_per = c_idx
            
            if c_reg != -1 and c_per != -1:
                header_row_idx = r_idx
                col_reg_idx = c_reg
                col_period_idx = c_per
                break
        
        if header_row_idx == -1:
            st.error("❌ Could not auto-detect 'Reg#' and 'Period' column headers in the uploaded Mileage Report.")
        else:
            # Load Rural Master List into Memory
            df_rural_raw = pd.read_excel(rural_file)
            rural_set = set()
            numeric_rural_set = set()
            
            for col in df_rural_raw.columns:
                for val in df_rural_raw[col].dropna():
                    val_str = str(val).replace('\xa0', '').strip().upper()
                    if val_str and val_str not in ['TROLLY NO.', 'UC NO.']:
                        rural_set.add(val_str)
                        digits = extract_digits(val_str)
                        if digits == val_str and digits != '':
                            numeric_rural_set.add(digits)
            
            # Extract Metadata Date
            meta_date = datetime.now().strftime("%m/%d/%Y")
            top_cell_val = str(df_mileage_raw.iloc[0, 0])
            if '[' in top_cell_val and ']' in top_cell_val:
                p1 = top_cell_val.find('[')
                p2 = top_cell_val.find(']')
                meta_date = top_cell_val[p1+1:p2].split('-')[0].strip()
            
            # Parse Data Table
            headers = df_mileage_raw.iloc[header_row_idx].astype(str).str.replace('\xa0', '', regex=False).str.strip().tolist()
            df_data = df_mileage_raw.iloc[header_row_idx + 1:].copy()
            df_data.columns = headers
            
            urban_rows = []
            rural_rows = []
            reg_counts = {}
            missing_reg_count = 0
            
            for idx, row in df_data.iterrows():
                period_val = row.iloc[col_period_idx]
                hrs = parse_period_to_hours(period_val)
                
                # Rule: Filter Period 20–24 Hours Inclusive
                if 20.0 <= hrs <= 24.0:
                    reg_val = str(row.iloc[col_reg_idx]).replace('\xa0', '').strip() if pd.notna(row.iloc[col_reg_idx]) else ''
                    
                    if not reg_val or reg_val == '-':
                        missing_reg_count += 1
                    else:
                        reg_clean = reg_val.upper()
                        reg_counts[reg_clean] = reg_counts.get(reg_clean, 0) + 1
                        
                        row_list = row.tolist()
                        if is_rural_vehicle(reg_clean, rural_set, numeric_rural_set):
                            rural_rows.append(row_list)
                        else:
                            urban_rows.append(row_list)
            
            duplicate_regs = {k for k, v in reg_counts.items() if v > 1}
            total_vehicles = len(urban_rows) + len(rural_rows)
            
            # Interactive On-Screen Metrics Summary
            st.markdown("### 📊 Live Processing Summary")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Vehicles (20–24 hrs)", total_vehicles)
            m2.metric("Urban Fleet", len(urban_rows))
            m3.metric("Rural Fleet", len(rural_rows))
            m4.metric("Missing Reg Skipped", missing_reg_count)
            
            if duplicate_regs:
                st.warning(f"⚠️ **Duplicate Registrations Detected ({len(duplicate_regs)}):** {', '.join(duplicate_regs)}")
            
            # Build Professional Excel Output in Memory
            wb = openpyxl.Workbook()
            ws_out = wb.active
            ws_out.title = "Final Report"
            ws_out.views.sheetView[0].showGridLines = True
            
            # Title Banner
            ws_out.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
            title_cell = ws_out.cell(row=1, column=1, value="VEHICLE MILEAGE REPORT")
            title_cell.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws_out.row_dimensions[1].height = 35
            
            # Dashboard Cards Header
            ws_out.cell(row=3, column=1, value="Report Date:").font = Font(bold=True)
            ws_out.cell(row=3, column=2, value=meta_date)
            ws_out.cell(row=4, column=1, value="Branch / Unit:").font = Font(bold=True)
            ws_out.cell(row=4, column=2, value="Main Office / Branch")
            ws_out.cell(row=5, column=1, value="Generated On:").font = Font(bold=True)
            ws_out.cell(row=5, column=2, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            ws_out.cell(row=6, column=1, value="Filtered Period:").font = Font(bold=True)
            ws_out.cell(row=6, column=2, value="20–24 Hours (Inclusive)")
            
            ws_out.cell(row=3, column=4, value="Total Vehicles:").font = Font(bold=True)
            ws_out.cell(row=3, column=5, value=total_vehicles).font = Font(bold=True, color="002060")
            ws_out.cell(row=4, column=4, value="Urban Fleet:").font = Font(bold=True)
            ws_out.cell(row=4, column=5, value=len(urban_rows)).font = Font(bold=True, color="002060")
            ws_out.cell(row=5, column=4, value="Rural Fleet:").font = Font(bold=True)
            ws_out.cell(row=5, column=5, value=len(rural_rows)).font = Font(bold=True, color="002060")
            ws_out.cell(row=6, column=4, value="Missing Reg Skipped:").font = Font(bold=True)
            ws_out.cell(row=6, column=5, value=missing_reg_count)
            
            # Table Headers
            header_row_out = 8
            ws_out.row_dimensions[header_row_out].height = 25
            for c_idx, h_text in enumerate(headers, 1):
                cell = ws_out.cell(row=header_row_out, column=c_idx, value=h_text)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="2980B9", end_color="2980B9", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            curr_r = 9
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )
            fill_alt = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
            fill_dup = PatternFill(start_color="FFE169", end_color="FFE169", fill_type="solid")
            
            # Urban Section Header
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=len(headers))
            u_hdr = ws_out.cell(row=curr_r, column=1, value=f"URBAN VEHICLES ({len(urban_rows)})")
            u_hdr.font = Font(bold=True, color="FFFFFF")
            u_hdr.fill = PatternFill(start_color="154360", end_color="154360", fill_type="solid")
            u_hdr.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws_out.row_dimensions[curr_r].height = 24
            curr_r += 1
            
            def write_rows(rows_data, start_r):
                r_pos = start_r
                for r_data in rows_data:
                    reg_val = str(r_data[col_reg_idx]).replace('\xa0', '').strip() if pd.notna(r_data[col_reg_idx]) else ''
                    reg_clean = reg_val.upper()
                    is_even = (r_pos % 2 == 0)
                    
                    for c_idx, val in enumerate(r_data, 1):
                        val_clean = str(val).replace('\xa0', '').strip() if pd.notna(val) else ''
                        cell = ws_out.cell(row=r_pos, column=c_idx, value=val_clean)
                        cell.border = thin_border
                        if is_even:
                            cell.fill = fill_alt
                        if c_idx == (col_reg_idx + 1) and reg_clean in duplicate_regs:
                            cell.fill = fill_dup
                            cell.font = Font(bold=True, color="9C5700")
                    r_pos += 1
                return r_pos
            
            curr_r = write_rows(urban_rows, curr_r)
            curr_r += 1
            
            # Rural Section Header
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=len(headers))
            r_hdr = ws_out.cell(row=curr_r, column=1, value=f"RURAL VEHICLES ({len(rural_rows)})")
            r_hdr.font = Font(bold=True, color="FFFFFF")
            r_hdr.fill = PatternFill(start_color="145A32", end_color="145A32", fill_type="solid")
            r_hdr.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws_out.row_dimensions[curr_r].height = 24
            curr_r += 1
            
            curr_r = write_rows(rural_rows, curr_r)
            
            # Auto-fit columns width
            for col in ws_out.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_out.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
            # Streamlit Download Button
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success("🎉 Final Report Generated Successfully!")
            st.download_button(
                label="📥 Download Formatted Final Report (.xlsx)",
                data=output_buffer,
                file_name=f"Mileage_Report_{meta_date.replace('/', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing report: {str(e)}")
