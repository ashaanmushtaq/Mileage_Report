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
# FUTURISTIC 3D RECEPTIONIST DESIGN - ADVANCED CSS
# ------------------------------------------------------------------------------
st.markdown("""
    <style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700;900&family=Rajdhani:wght@300;400;600;700&display=swap');
    
    /* Global Background */
    .stApp {
        background: linear-gradient(135deg, #0a0e1a 0%, #1a1f3a 50%, #0d1025 100%);
    }
    
    /* Main Container Glass Effect */
    .main-container {
        background: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border-radius: 24px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        padding: 30px;
        margin: 20px 0;
        box-shadow: 
            0 30px 60px rgba(0,0,0,0.5),
            inset 0 1px 0 rgba(255,255,255,0.1);
    }
    
    /* 3D Receptionist Header */
    .reception-header {
        background: linear-gradient(135deg, #0a0e1a, #1a1f3a);
        border-radius: 20px;
        padding: 40px 30px;
        margin-bottom: 30px;
        border: 1px solid rgba(100, 200, 255, 0.2);
        box-shadow: 
            0 20px 60px rgba(0, 100, 255, 0.15),
            inset 0 1px 0 rgba(255,255,255,0.1),
            0 0 40px rgba(0, 100, 255, 0.05);
        position: relative;
        overflow: hidden;
    }
    
    .reception-header::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: radial-gradient(ellipse at center, rgba(100, 200, 255, 0.05) 0%, transparent 70%);
        animation: headerGlow 8s ease-in-out infinite;
    }
    
    @keyframes headerGlow {
        0%, 100% { transform: translate(0, 0); }
        50% { transform: translate(10%, 10%); }
    }
    
    .reception-title {
        font-family: 'Orbitron', sans-serif;
        font-size: 48px;
        font-weight: 900;
        background: linear-gradient(135deg, #00d4ff 0%, #7b2ffc 50%, #ff6b9d 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        text-shadow: none;
        letter-spacing: 4px;
        position: relative;
        z-index: 1;
        animation: titlePulse 3s ease-in-out infinite;
    }
    
    @keyframes titlePulse {
        0%, 100% { transform: scale(1); }
        50% { transform: scale(1.02); }
    }
    
    .reception-sub {
        font-family: 'Rajdhani', sans-serif;
        font-size: 18px;
        color: rgba(255, 255, 255, 0.8);
        letter-spacing: 6px;
        text-transform: uppercase;
        position: relative;
        z-index: 1;
        margin-top: 10px;
    }
    
    .reception-sub b {
        color: #00d4ff;
        font-weight: 700;
    }
    
    /* Holographic Badge */
    .holographic-badge {
        display: inline-block;
        background: rgba(0, 212, 255, 0.1);
        border: 1px solid rgba(0, 212, 255, 0.3);
        border-radius: 50px;
        padding: 8px 24px;
        font-family: 'Rajdhani', sans-serif;
        font-weight: 600;
        color: #00d4ff;
        letter-spacing: 2px;
        backdrop-filter: blur(10px);
        box-shadow: 0 0 30px rgba(0, 212, 255, 0.1);
        position: relative;
        z-index: 1;
    }
    
    /* Glass Cards */
    .glass-card {
        background: rgba(255, 255, 255, 0.03);
        backdrop-filter: blur(15px);
        -webkit-backdrop-filter: blur(15px);
        border-radius: 16px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        padding: 24px;
        box-shadow: 
            0 8px 32px rgba(0,0,0,0.3),
            inset 0 1px 0 rgba(255,255,255,0.05);
        transition: all 0.3s ease;
    }
    
    .glass-card:hover {
        transform: translateY(-2px);
        border-color: rgba(0, 212, 255, 0.2);
        box-shadow: 
            0 12px 48px rgba(0, 100, 255, 0.15),
            inset 0 1px 0 rgba(255,255,255,0.08);
    }
    
    /* Metric Cards */
    .metric-card {
        background: linear-gradient(135deg, rgba(0, 212, 255, 0.05), rgba(123, 47, 252, 0.05));
        border-radius: 16px;
        padding: 20px;
        border: 1px solid rgba(0, 212, 255, 0.1);
        text-align: center;
        backdrop-filter: blur(10px);
        transition: all 0.3s ease;
    }
    
    .metric-card:hover {
        transform: scale(1.02);
        border-color: rgba(0, 212, 255, 0.3);
        box-shadow: 0 0 40px rgba(0, 212, 255, 0.1);
    }
    
    .metric-value {
        font-family: 'Orbitron', sans-serif;
        font-size: 32px;
        font-weight: 700;
        background: linear-gradient(135deg, #00d4ff, #7b2ffc);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    
    .metric-label {
        font-family: 'Rajdhani', sans-serif;
        font-size: 14px;
        color: rgba(255, 255, 255, 0.7);
        text-transform: uppercase;
        letter-spacing: 2px;
        margin-top: 8px;
    }
    
    /* Upload Section */
    .upload-section {
        background: rgba(255, 255, 255, 0.02);
        border: 2px dashed rgba(0, 212, 255, 0.2);
        border-radius: 16px;
        padding: 30px;
        text-align: center;
        transition: all 0.3s ease;
    }
    
    .upload-section:hover {
        border-color: rgba(0, 212, 255, 0.4);
        background: rgba(0, 212, 255, 0.03);
    }
    
    /* Custom File Uploader */
    .stFileUploader > div {
        background: transparent !important;
        border: none !important;
    }
    
    .stFileUploader > div > div {
        background: rgba(255, 255, 255, 0.03) !important;
        border: 2px dashed rgba(0, 212, 255, 0.2) !important;
        border-radius: 16px !important;
        padding: 20px !important;
        transition: all 0.3s ease !important;
    }
    
    .stFileUploader > div > div:hover {
        border-color: rgba(0, 212, 255, 0.5) !important;
        background: rgba(0, 212, 255, 0.05) !important;
    }
    
    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #00d4ff, #7b2ffc) !important;
        border: none !important;
        border-radius: 50px !important;
        color: white !important;
        font-family: 'Rajdhani', sans-serif !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        padding: 12px 36px !important;
        text-transform: uppercase !important;
        letter-spacing: 2px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 20px rgba(0, 212, 255, 0.3) !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px) scale(1.02) !important;
        box-shadow: 0 8px 40px rgba(0, 212, 255, 0.5) !important;
    }
    
    /* Download Button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #00d4ff, #7b2ffc) !important;
        border: none !important;
        border-radius: 50px !important;
        color: white !important;
        font-family: 'Rajdhani', sans-serif !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        padding: 12px 36px !important;
        text-transform: uppercase !important;
        letter-spacing: 2px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 20px rgba(0, 212, 255, 0.3) !important;
    }
    
    .stDownloadButton > button:hover {
        transform: translateY(-2px) scale(1.02) !important;
        box-shadow: 0 8px 40px rgba(0, 212, 255, 0.5) !important;
    }
    
    /* Select Box */
    .stSelectbox > div {
        background: rgba(255, 255, 255, 0.03) !important;
        border: 1px solid rgba(0, 212, 255, 0.2) !important;
        border-radius: 12px !important;
        color: white !important;
    }
    
    .stSelectbox > div:hover {
        border-color: rgba(0, 212, 255, 0.4) !important;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background: rgba(255, 255, 255, 0.03) !important;
        border: 1px solid rgba(0, 212, 255, 0.1) !important;
        border-radius: 12px !important;
        font-family: 'Rajdhani', sans-serif !important;
        color: rgba(255, 255, 255, 0.8) !important;
    }
    
    .streamlit-expanderHeader:hover {
        border-color: rgba(0, 212, 255, 0.3) !important;
    }
    
    /* Info/Warning/Success Boxes */
    .stAlert {
        border-radius: 12px !important;
        border: 1px solid rgba(0, 212, 255, 0.1) !important;
        background: rgba(255, 255, 255, 0.03) !important;
        backdrop-filter: blur(10px) !important;
        color: white !important;
    }
    
    .stAlert > div {
        color: white !important;
    }
    
    /* Custom Scroll */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: rgba(255, 255, 255, 0.05);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #00d4ff, #7b2ffc);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(135deg, #00d4ff, #ff6b9d);
    }
    
    /* Headings */
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Orbitron', sans-serif !important;
        color: white !important;
    }
    
    /* Labels */
    .stMarkdown {
        color: rgba(255, 255, 255, 0.8) !important;
    }
    
    /* Text Input */
    .stTextInput > div {
        background: rgba(255, 255, 255, 0.03) !important;
        border: 1px solid rgba(0, 212, 255, 0.2) !important;
        border-radius: 12px !important;
        color: white !important;
    }
    
    .stTextInput > div:focus-within {
        border-color: rgba(0, 212, 255, 0.5) !important;
        box-shadow: 0 0 20px rgba(0, 212, 255, 0.1) !important;
    }
    
    /* Particle Effect Background */
    .particles {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        pointer-events: none;
        z-index: 0;
        overflow: hidden;
    }
    </style>
    
    <!-- Particle Animation -->
    <div class="particles" id="particles"></div>
    <script>
    (function() {
        const canvas = document.createElement('canvas');
        const container = document.querySelector('.particles');
        container.appendChild(canvas);
        const ctx = canvas.getContext('2d');
        
        canvas.width = window.innerWidth;
        canvas.height = window.innerHeight;
        
        const particles = [];
        const colors = ['#00d4ff', '#7b2ffc', '#ff6b9d', '#00ff88'];
        
        class Particle {
            constructor() {
                this.x = Math.random() * canvas.width;
                this.y = Math.random() * canvas.height;
                this.size = Math.random() * 2 + 1;
                this.speedX = (Math.random() - 0.5) * 0.5;
                this.speedY = (Math.random() - 0.5) * 0.5;
                this.color = colors[Math.floor(Math.random() * colors.length)];
                this.opacity = Math.random() * 0.5 + 0.2;
            }
            
            update() {
                this.x += this.speedX;
                this.y += this.speedY;
                
                if (this.x > canvas.width) this.x = 0;
                if (this.x < 0) this.x = canvas.width;
                if (this.y > canvas.height) this.y = 0;
                if (this.y < 0) this.y = canvas.height;
            }
            
            draw() {
                ctx.beginPath();
                ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
                ctx.fillStyle = this.color;
                ctx.globalAlpha = this.opacity;
                ctx.fill();
                ctx.globalAlpha = 1;
            }
        }
        
        for (let i = 0; i < 80; i++) {
            particles.push(new Particle());
        }
        
        function animate() {
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            particles.forEach(p => {
                p.update();
                p.draw();
            });
            requestAnimationFrame(animate);
        }
        
        animate();
        
        window.addEventListener('resize', () => {
            canvas.width = window.innerWidth;
            canvas.height = window.innerHeight;
        });
    })();
    </script>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# MAIN CONTAINER
# ------------------------------------------------------------------------------
st.markdown("""
<div class="main-container">
    <div class="reception-header">
        <div style="position: relative; z-index: 1; text-align: center;">
            <div style="display: flex; justify-content: center; align-items: center; gap: 20px; margin-bottom: 10px;">
                <span style="font-size: 60px;">🧾</span>
                <span class="reception-title">MILEAGE EXECUTIVE</span>
                <span style="font-size: 60px;">💫</span>
            </div>
            <div class="reception-sub">
                <span style="color: rgba(255,255,255,0.5);">✦</span> 
                DEVELOPED BY <b>MUHAMMAD ASHAAN</b> 
                <span style="color: rgba(255,255,255,0.5);">✦</span>
            </div>
            <div style="margin-top: 15px; display: flex; justify-content: center; gap: 15px; flex-wrap: wrap;">
                <span class="holographic-badge">🔮 BELLANIX TECH</span>
                <span class="holographic-badge" style="border-color: rgba(123, 47, 252, 0.3); color: #7b2ffc;">⚡ AI POWERED</span>
                <span class="holographic-badge" style="border-color: rgba(255, 107, 157, 0.3); color: #ff6b9d;">🎯 3D RECEPTION</span>
            </div>
        </div>
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
# CITY SELECTION & FILE UPLOAD SECTION
# ------------------------------------------------------------------------------
st.markdown("""
    <div style="margin: 30px 0 20px 0;">
        <h3 style="color: #00d4ff; font-family: 'Orbitron', sans-serif; font-size: 20px; letter-spacing: 2px;">📍 LOCATION & DATA INPUT</h3>
    </div>
""", unsafe_allow_html=True)

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

st.markdown('<div style="margin-top: 20px;"></div>', unsafe_allow_html=True)

col_up1, col_up2 = st.columns(2)

with col_up1:
    st.markdown("""
        <div class="upload-section">
            <span style="font-size: 28px;">📊</span>
            <p style="font-family: 'Rajdhani', sans-serif; color: rgba(255,255,255,0.7); margin-top: 10px;">
                <b style="color: #00d4ff;">TODAY'S</b> Mileage Report
            </p>
        </div>
    """, unsafe_allow_html=True)
    current_file = st.file_uploader("1️⃣ Upload TODAY'S Mileage Report (.xlsx)", type=["xlsx"], key="current")

with col_up2:
    st.markdown("""
        <div class="upload-section">
            <span style="font-size: 28px;">📈</span>
            <p style="font-family: 'Rajdhani', sans-serif; color: rgba(255,255,255,0.7); margin-top: 10px;">
                <b style="color: #7b2ffc;">PREVIOUS</b> Mileage Report
            </p>
        </div>
    """, unsafe_allow_html=True)
    previous_file = st.file_uploader("2️⃣ Upload PREVIOUS Mileage Report (.xlsx) - Optional", type=["xlsx"], key="previous")

with st.expander("⚙️ Master List Override", expanded=False):
    st.markdown("""
        <p style="font-family: 'Rajdhani', sans-serif; color: rgba(255,255,255,0.6);">
            Upload a custom vehicle master list to override the default rural fleet configuration.
        </p>
    """, unsafe_allow_html=True)
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

            # --- FUTURISTIC METRICS DISPLAY ---
            st.markdown("""
                <div style="margin: 30px 0 20px 0;">
                    <h3 style="color: #7b2ffc; font-family: 'Orbitron', sans-serif; font-size: 20px; letter-spacing: 2px;">📊 HOLOGRAPHIC METRICS</h3>
                </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
                <div style="background: rgba(255,255,255,0.03); border-radius: 12px; padding: 12px 20px; margin-bottom: 20px; border-left: 3px solid #00d4ff;">
                    <p style="font-family: 'Rajdhani', sans-serif; color: rgba(255,255,255,0.8); font-size: 15px; margin: 0;">
                        📍 <b style="color: #00d4ff;">{selected_city}</b> &nbsp;|&nbsp; 📅 <b style="color: #7b2ffc;">{meta_date}</b>
                    </p>
                </div>
            """, unsafe_allow_html=True)
            
            m1, m2, m3, m4 = st.columns(4)
            
            with m1:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{total_valid_vehicles}</div>
                        <div class="metric-label">Total Fleet (20-24h)</div>
                    </div>
                """, unsafe_allow_html=True)
            
            with m2:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value" style="background: linear-gradient(135deg, #00d4ff, #00ff88); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">{len(urban_rows)}</div>
                        <div class="metric-label">Urban Fleet</div>
                    </div>
                """, unsafe_allow_html=True)
            
            with m3:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value" style="background: linear-gradient(135deg, #7b2ffc, #ff6b9d); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">{len(rural_rows)}</div>
                        <div class="metric-label">Rural Fleet</div>
                    </div>
                """, unsafe_allow_html=True)
            
            with m4:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value" style="background: linear-gradient(135deg, #ff6b9d, #ffd700); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">{len(repeated_vehicles)}</div>
                        <div class="metric-label">2-Day Repeat</div>
                    </div>
                """, unsafe_allow_html=True)

            if repeated_vehicles:
                st.info(f"🔄 **2-Day Consecutive 20-24h Repeat Vehicles ({len(repeated_vehicles)}):** {', '.join(repeated_vehicles)}")
            if duplicate_regs:
                st.warning(f"⚠️ **Same-Day Duplicates ({len(duplicate_regs)}):** {', '.join(duplicate_regs)}")

            # ------------------------------------------------------------------
            # EXCEL REPORT BUILDER - AESTHETIC UPGRADE
            # ------------------------------------------------------------------
            wb = openpyxl.Workbook()
            ws_out = wb.active
            ws_out.title = "Executive Report"
            ws_out.views.sheetView[0].showGridLines = False

            center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            num_cols = len(headers)
            half_cols = max(2, num_cols // 2)

            # 1. Main Title Banner - Premium Gradient
            ws_out.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
            title_text = f"✨ {selected_city} VEHICLE MILEAGE EXECUTIVE REPORT ✨"
            title_cell = ws_out.cell(row=1, column=1, value=title_text)
            title_cell.font = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            title_cell.alignment = center_align
            ws_out.row_dimensions[1].height = 40

            # 2. Decorative Subtitle
            ws_out.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
            sub_cell = ws_out.cell(row=2, column=1, value="◆ BELLANIX TECH ◆ AI POWERED ANALYTICS ◆")
            sub_cell.font = Font(name="Calibri", size=10, italic=True, color="7F8C8D")
            sub_cell.alignment = center_align
            ws_out.row_dimensions[2].height = 20

            # 3. Metadata Block - Glass Effect Style
            ws_out.merge_cells(start_row=4, start_column=1, end_row=4, end_column=half_cols)
            r4_left = ws_out.cell(row=4, column=1, value=f"📅 Report Date:  {meta_date}")
            r4_left.font = Font(bold=True, size=11, color="2C3E50")
            r4_left.alignment = center_align
            r4_left.fill = PatternFill(start_color="ECF0F1", end_color="ECF0F1", fill_type="solid")

            ws_out.merge_cells(start_row=4, start_column=half_cols+1, end_row=4, end_column=num_cols)
            r4_right = ws_out.cell(row=4, column=half_cols+1, value=f"🏆 Total Fleet (20-24h):  {total_valid_vehicles}")
            r4_right.font = Font(bold=True, size=11, color="1F4E79")
            r4_right.alignment = center_align
            r4_right.fill = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")

            ws_out.merge_cells(start_row=5, start_column=1, end_row=5, end_column=half_cols)
            r5_left = ws_out.cell(row=5, column=1, value="👨‍💻 Developed by:  Muhammad Ashaan")
            r5_left.font = Font(bold=True, size=11, color="1F4E79")
            r5_left.alignment = center_align
            r5_left.fill = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")

            ws_out.merge_cells(start_row=5, start_column=half_cols+1, end_row=5, end_column=num_cols)
            r5_right = ws_out.cell(row=5, column=half_cols+1, value=f"🏙️ Urban Fleet Count:  {len(urban_rows)}")
            r5_right.font = Font(bold=True, size=11, color="154360")
            r5_right.alignment = center_align
            r5_right.fill = PatternFill(start_color="D4E6F1", end_color="D4E6F1", fill_type="solid")

            ws_out.merge_cells(start_row=6, start_column=1, end_row=6, end_column=half_cols)
            r6_left = ws_out.cell(row=6, column=1, value="⏱️ Filtered Hours:  20–24 Hours")
            r6_left.font = Font(bold=True, size=11, color="2C3E50")
            r6_left.alignment = center_align
            r6_left.fill = PatternFill(start_color="ECF0F1", end_color="ECF0F1", fill_type="solid")

            ws_out.merge_cells(start_row=6, start_column=half_cols+1, end_row=6, end_column=num_cols)
            r6_right = ws_out.cell(row=6, column=half_cols+1, value=f"🌾 Rural Fleet Count:  {len(rural_rows)}")
            r6_right.font = Font(bold=True, size=11, color="145A32")
            r6_right.alignment = center_align
            r6_right.fill = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")

            for r in range(4, 7):
                ws_out.row_dimensions[r].height = 24

            # 4. Table Column Headers - Premium Gradient
            header_row_out = 8
            ws_out.row_dimensions[header_row_out].height = 30
            header_colors = ["3498DB", "2980B9", "2471A3", "1F618D", "154360"]
            for c_idx, h_text in enumerate(headers, 1):
                cell = ws_out.cell(row=header_row_out, column=c_idx, value=h_text)
                cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color=header_colors[(c_idx-1) % len(header_colors)], 
                                       end_color=header_colors[(c_idx-1) % len(header_colors)], 
                                       fill_type="solid")
                cell.alignment = center_align
                cell.border = Border(
                    left=Side(style='medium', color='1F4E79'),
                    right=Side(style='medium', color='1F4E79'),
                    top=Side(style='medium', color='1F4E79'),
                    bottom=Side(style='medium', color='1F4E79')
                )

            curr_r = 9
            sno_tracker = 1
            
            thin_border = Border(
                left=Side(style='thin', color='BDC3C7'),
                right=Side(style='thin', color='BDC3C7'),
                top=Side(style='thin', color='BDC3C7'),
                bottom=Side(style='thin', color='BDC3C7')
            )
            fill_alt = PatternFill(start_color="F4F6F7", end_color="F4F6F7", fill_type="solid")
            fill_dup = PatternFill(start_color="FCF3CF", end_color="FCF3CF", fill_type="solid")
            fill_repeat = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")

            actual_reg_col = c_reg_idx + (0 if has_sno else 1)

            # --- URBAN FLEET SECTION - Premium Styling ---
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=num_cols)
            u_sec = ws_out.cell(row=curr_r, column=1, value=f"🏙️ URBAN FLEET VEHICLES ({len(urban_rows)})")
            u_sec.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
            u_sec.fill = PatternFill(start_color="1B4F72", end_color="1B4F72", fill_type="solid")
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
                    if is_even:
                        cell.fill = fill_alt
                    
                    if c_idx == (actual_reg_col + 1):
                        if reg_clean in repeated_vehicles:
                            cell.fill = fill_repeat
                            cell.font = Font(bold=True, color="C0392B")
                        elif reg_clean in duplicate_regs:
                            cell.fill = fill_dup
                            cell.font = Font(bold=True, color="B7950B")
                curr_r += 1

            curr_r += 1

            # --- RURAL FLEET SECTION - Premium Styling ---
            ws_out.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=num_cols)
            r_sec = ws_out.cell(row=curr_r, column=1, value=f"🌾 RURAL FLEET VEHICLES ({len(rural_rows)})")
            r_sec.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
            r_sec.fill = PatternFill(start_color="1E8449", end_color="1E8449", fill_type="solid")
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
                    if is_even:
                        cell.fill = fill_alt
                    
                    if c_idx == (actual_reg_col + 1):
                        if reg_clean in repeated_vehicles:
                            cell.fill = fill_repeat
                            cell.font = Font(bold=True, color="C0392B")
                        elif reg_clean in duplicate_regs:
                            cell.fill = fill_dup
                            cell.font = Font(bold=True, color="B7950B")
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
            
            st.markdown("""
                <div style="background: rgba(0, 212, 255, 0.05); border: 1px solid rgba(0, 212, 255, 0.2); border-radius: 16px; padding: 20px; margin-top: 30px; text-align: center;">
                    <p style="font-family: 'Orbitron', sans-serif; color: #00d4ff; font-size: 18px; letter-spacing: 2px;">✅ REPORT GENERATED SUCCESSFULLY</p>
                    <p style="font-family: 'Rajdhani', sans-serif; color: rgba(255,255,255,0.6);">Your executive audit report is ready for download</p>
                </div>
            """, unsafe_allow_html=True)
            
            st.download_button(
                label=f"📥 Download {selected_city} Executive Report",
                data=output_buffer,
                file_name=f"{selected_city}_Mileage_Report_{meta_date.replace('/', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error processing report: {str(e)}")

st.markdown("</div>", unsafe_allow_html=True)
