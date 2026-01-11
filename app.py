import streamlit as st
import pandas as pd
import re
import math
import io
import datetime
import requests
from PIL import Image
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.lib.enums import TA_CENTER

# ---------------- CONFIG ----------------
TG_LINK = "https://t.me/MurlidharAcademy"
IG_LINK = "https://www.instagram.com/murlidhar_academy_official/"

# ✅ Google Drive Image IDs (JPGs)
DEFAULT_DRIVE_ID = "1a1ZK5uiLl0a63Pto1EQDUY0VaIlqp21u"
SIGNATURE_ID = "1U0es4MVJgGniK27rcrA6hiLFFRazmwCs"
LOGO_ID = "1BGvxglcgZ2G6FdVelLjXZVo-_v4e4a42"

# ✅ CHARACTER IMAGE IDs
CHAR_IDS = {
    "VIKRAMADITYA": "19f511argqR5e1ajYIWT4a-PIHJV9pexw",
    "CHANAKYA": "15SiSFtjjvV_G5zl5QBJJy1trb_1YZxie",
    "ARJUNA": "10q8t65_zMvQ9p-4s9BRaZqzVJ1nsnJck",
    "DHRUVA": "1jQ0hzX9Y0bKMGh7htP_mIXOKSmHAR6P_",
    "KARNA": "1DvL4WqmJHlhMhs4TR9O47TyvdqYNpX6u",
    "ANGAD": "1LMT1PfsAxzHrVtexQN2HGkz_d60zmaFd",
    "BHAGIRATH": "1nUQh5P2eAEZE4m7mBJGUFFFzkv1P5BZX"
}

# ==========================================
# 🎛️ CERTIFICATE LAYOUT CONFIGURATION
# ==========================================

# 1. LOGO SETTINGS (Left Side)
CERT_LOGO_WIDTH = 40 * mm       
CERT_LOGO_HEIGHT = 40 * mm      
CERT_LOGO_X_POS = 36 * mm       
CERT_LOGO_Y_POS = 143 * mm      

# 2. SIGNATURE SETTINGS (Bottom Right)
CERT_SIGN_WIDTH = 65 * mm       
CERT_SIGN_HEIGHT = 22 * mm      
CERT_SIGN_X_POS = 235 * mm      
CERT_SIGN_Y_POS = 38 * mm       

# 3. CHARACTER IMAGE SETTINGS (Right Side - Background)
CERT_CHAR_WIDTH = 110 * mm      # Size (Width)
CERT_CHAR_HEIGHT = 110 * mm     # Size (Height)
CERT_CHAR_OPACITY = 0.5        # Opacity (0.35 is perfect as per your request)

# 👇👇👇 (અહીંથી પોઝિશન સેટ કરો) 👇👇👇
CERT_CHAR_MARGIN_RIGHT = 12 * mm   # જમણી બાજુથી કેટલું દૂર રાખવું? (Right Margin)
CERT_CHAR_MARGIN_TOP = 18 * mm    # ઉપરની બાજુથી કેટલું નીચે રાખવું? (Top Margin)

# (આ ઓટોમેટિક ગણતરી કરશે, તમારે આમાં ફેરફાર કરવાની જરૂર નથી)
PAGE_W_MM = 297 # A4 Landscape Width
PAGE_H_MM = 210 # A4 Landscape Height
CERT_CHAR_X_POS = (PAGE_W_MM * mm) - CERT_CHAR_WIDTH - CERT_CHAR_MARGIN_RIGHT
CERT_CHAR_Y_POS = (PAGE_H_MM * mm) - CERT_CHAR_HEIGHT - CERT_CHAR_MARGIN_TOP

# ==========================================

LEFT_MARGIN_mm = 18
RIGHT_MARGIN_mm = 18
TITLE_Y_mm_from_top = 63.5
TABLE_SPACE_AFTER_TITLE_mm = 16
PAGE_NO_Y_mm = 8
ROWS_PER_PAGE = 23
DEFAULT_TEST_MAX_PER_FILE = 50.0

# ✅ THEME COLORS
COLOR_BLUE_HEADER = colors.HexColor("#0f5f9a")
COLOR_GREEN = colors.HexColor("#C8E6C9")
COLOR_YELLOW = colors.HexColor("#FFF9C4")
COLOR_RED = colors.HexColor("#FFCDD2")
COLOR_SAFFRON = colors.HexColor("#FF9933")
COLOR_GOLD = colors.HexColor("#B8860B")

# ✅ SUMMARY COLORS
SUMMARY_HEADER_COLORS = {
    "METRIC": colors.HexColor("#1976D2"),           
    "TOP 5 RANKERS": colors.HexColor("#2E7D32"),    
    "BOTTOM 3 (NEEDS IMPROVEMENT)": colors.HexColor("#C62828") 
}

# ---------------- HELPERS ----------------
def get_drive_url(file_id):
    return f'https://drive.google.com/uc?export=download&id={file_id}'

@st.cache_data(show_spinner=False)
def download_image_from_drive(file_id):
    try:
        url = get_drive_url(file_id)
        response = requests.get(url, allow_redirects=True)
        if response.status_code == 200:
            return io.BytesIO(response.content)
        else:
            return None
    except:
        return None

# ✅ HELPER: PROCESS IMAGE FOR OPACITY (WATERMARK)
def get_transparent_image_reader(img_bytes, opacity=0.5):
    """
    Reads an image, converts to RGBA, applies opacity to Alpha channel,
    and returns an ImageReader object for ReportLab.
    """
    if not img_bytes: return None
    try:
        img_bytes.seek(0)
        img = Image.open(img_bytes).convert("RGBA")
        
        # Adjust Alpha Channel
        r, g, b, alpha = img.split()
        # Evaluate alpha: multiply current alpha by opacity factor
        alpha = alpha.point(lambda p: int(p * opacity))
        img.putalpha(alpha)
        
        # Save to buffer
        new_buffer = io.BytesIO()
        img.save(new_buffer, format='PNG')
        new_buffer.seek(0)
        return ImageReader(new_buffer)
    except Exception as e:
        print(f"Error processing image transparency: {e}")
        return None

def normalize_name(s):
    if pd.isna(s): return ""
    s = str(s).strip()
    s = re.sub(r'\s+', ' ', s)
    return s.lower()

def find_name_series(df):
    cols = list(df.columns)
    lc = [c.lower() for c in cols]
    if 'firstname' in lc and 'lastname' in lc:
        fn = df[cols[lc.index('firstname')]].astype(str).fillna("")
        ln = df[cols[lc.index('lastname')]].astype(str).fillna("")
        return (fn.str.strip() + " " + ln.str.strip()).astype(str)
    
    keywords = ['name','student name','student','full name','studentname', 'candidate']
    for key in keywords:
        for c in cols:
            if key == c.lower().strip():
                return df[c].astype(str).fillna("").str.strip()
    
    for c in cols:
        if 'name' in c.lower() or 'student' in c.lower():
            return df[c].astype(str).fillna("").str.strip()
            
    return pd.Series([f"Student {i+1}" for i in range(len(df))])

def find_possible_pts(df):
    cols = df.columns
    for c in cols:
        if c.strip().lower() in ('possiblepts','possible_pts','possible points','possible', 'max marks'):
            vals = pd.to_numeric(df[c], errors='coerce').dropna()
            if len(vals) > 0: return float(vals.iloc[0])
    for c in cols:
        if any(k in c.lower() for k in ['possible','max','maximum','totalmarks']):
            vals = pd.to_numeric(df[c], errors='coerce').dropna()
            if len(vals) > 0: return float(vals.max())
    return None

def extract_obtained_series(df):
    cols = df.columns
    for c in cols:
        clean = c.lower().replace(" ", "")
        if 'earnedpts' in clean or 'obtainedmarks' in clean or 'score' == clean:
            return pd.to_numeric(df[c], errors='coerce').fillna(0).astype(float)
            
    numeric_candidates = []
    for c in cols:
        if 'phone' in c.lower() or 'id' in c.lower() or 'roll' in c.lower(): continue
        s = pd.to_numeric(df[c], errors='coerce')
        if s.notna().sum() > 0:
            numeric_candidates.append((c, s.mean()))
    
    if numeric_candidates:
        numeric_candidates.sort(key=lambda x: x[1], reverse=True)
        return pd.to_numeric(df[numeric_candidates[0][0]], errors='coerce').fillna(0).astype(float)
        
    return pd.Series(0.0, index=df.index)

def get_smart_row_color(pct, is_even_row, t_green, t_yellow):
    if pct >= t_green: 
        return colors.HexColor("#E8F5E9") if is_even_row else colors.HexColor("#C8E6C9")
    if pct >= t_yellow: 
        return colors.HexColor("#FFFDE7") if is_even_row else colors.HexColor("#FFF9C4")
    return colors.HexColor("#FFEBEE") if is_even_row else colors.HexColor("#FFCDD2")

# ---------------- CERTIFICATE GENERATOR FUNCTION ----------------
def generate_certificates_pdf(out_df, thresh_yellow, thresh_green, report_title, cert_date, logo_bytes, sign_bytes, char_images_bytes):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(A4))
    width, height = landscape(A4)
    
    logo_img = ImageReader(Image.open(logo_bytes)) if logo_bytes else None
    sign_img = ImageReader(Image.open(sign_bytes)) if sign_bytes else None

    # Pre-process Character Images with Opacity
    char_readers = {}
    for key, val_bytes in char_images_bytes.items():
        if val_bytes:
            # Apply Opacity
            reader = get_transparent_image_reader(val_bytes, opacity=CERT_CHAR_OPACITY)
            if reader:
                char_readers[key] = reader

    awards_to_give = [] 

    # --- AWARD LOGIC ---
    rank1 = out_df[out_df['Rank'] == 1]
    for _, r in rank1.iterrows():
        desc = (
            "Like the legendary King Vikramaditya, known for his wisdom and victory, "
            "you have conquered this challenge with supreme excellence! "
            f"Your hard work has placed you at the very top. Keep ruling! [Score: {r['Obtained']}/{r['Total Marks']}]"
        )
        awards_to_give.append((r['Name'], "THE VIKRAMADITYA EXCELLENCE AWARD", desc, COLOR_GOLD, "VIKRAMADITYA"))

    rank2 = out_df[out_df['Rank'] == 2]
    for _, r in rank2.iterrows():
        desc = (
            "With the sharp intellect of Acharya Chanakya, you have proven that strategy determines success. "
            f"Your outstanding intelligence and dedication have secured you the prestigious 2nd Rank. [Score: {r['Obtained']}/{r['Total Marks']}]"
        )
        awards_to_give.append((r['Name'], "THE CHANAKYA NITI AWARD", desc, COLOR_BLUE_HEADER, "CHANAKYA"))

    rank3 = out_df[out_df['Rank'] == 3]
    for _, r in rank3.iterrows():
        desc = (
            "Just like Arjuna saw only the bird's eye, your laser-sharp focus and precision have hit the mark! "
            f"This award celebrates your unwavering concentration and excellent performance (Rank 3). [Score: {r['Obtained']}/{r['Total Marks']}]"
        )
        awards_to_give.append((r['Name'], "THE ARJUNA FOCUS AWARD", desc, COLOR_SAFFRON, "ARJUNA"))

    rank45 = out_df[out_df['Rank'].isin([4, 5])]
    for _, r in rank45.iterrows():
        desc = (
            "Like the eternal Dhruva Tara (Pole Star), your performance shines bright with stability and consistency. "
            f"You are a rising star with immense potential to lead the sky! [Score: {r['Obtained']}]"
        )
        awards_to_give.append((r['Name'], "THE DHRUVA TARA AWARD", desc, COLOR_BLUE_HEADER, "DHRUVA"))

    rank6_10 = out_df[out_df['Rank'].isin([6,7,8,9,10])]
    for _, r in rank6_10.iterrows():
        desc = (
            "A true warrior is defined by their spirit! Like Maharathi Karna, you fought bravely and showed immense talent. "
            f"You are just steps away from the top. Keep fighting, victory is yours! [Score: {r['Obtained']}]"
        )
        awards_to_give.append((r['Name'], "THE KARNA VEERTA AWARD", desc, COLOR_SAFFRON, "KARNA"))

    angad_candidates = out_df[(out_df['Absent'] == 0) & (out_df['Percentage'] >= thresh_yellow) & (out_df['Rank'] > 10)]
    for _, r in angad_candidates.iterrows():
        desc = (
            "Firm as Angad's foot in Ravana's court! Your unshakeable discipline and 100% Attendance prove that consistency is the key to success. "
            "You stood firm in every test! [Attendance: 100%]"
        )
        awards_to_give.append((r['Name'], "THE ANGAD STAMBH AWARD", desc, COLOR_BLUE_HEADER, "ANGAD"))

    bhagirath_candidates = out_df[
        (out_df['Percentage'] >= thresh_yellow) & (out_df['Percentage'] < thresh_green) &
        (out_df['Present'] / out_df['Total Tests'] >= 0.8) & (out_df['Rank'] > 10) & (out_df['Absent'] > 0)
    ]
    for _, r in bhagirath_candidates.iterrows():
        desc = (
            "Like Bhagirath's relentless penance to bring Ganga to Earth, your hard work and persistence are truly inspiring. "
            "This award honors your 'Never Give Up' attitude and continuous improvement."
        )
        awards_to_give.append((r['Name'], "THE BHAGIRATH PRAYAS AWARD", desc, COLOR_SAFFRON, "BHAGIRATH"))

    # DRAWING CERTIFICATES
    for student_name, title, desc, theme_color, char_key in awards_to_give:
        
        # 1. DRAW CHARACTER IMAGE FIRST (BACKGROUND)
        if char_key in char_readers:
            char_img = char_readers[char_key]
            c.drawImage(char_img, CERT_CHAR_X_POS, CERT_CHAR_Y_POS, width=CERT_CHAR_WIDTH, height=CERT_CHAR_HEIGHT, mask='auto', preserveAspectRatio=True)

        # 2. DRAW BORDERS (On top of character)
        c.setStrokeColor(theme_color)
        c.setLineWidth(5); c.rect(15*mm, 15*mm, width-30*mm, height-30*mm)
        c.setLineWidth(1); c.rect(18*mm, 18*mm, width-36*mm, height-36*mm)

        # 3. HEADER SECTION
        center_x = width / 2
        
        # LOGO
        if logo_img:
            c.drawImage(logo_img, CERT_LOGO_X_POS, CERT_LOGO_Y_POS, width=CERT_LOGO_WIDTH, height=CERT_LOGO_HEIGHT, mask='auto', preserveAspectRatio=True)

        # Header Text
        c.setFont("Helvetica-Bold", 32)
        c.setFillColor(COLOR_BLUE_HEADER)
        c.drawCentredString(center_x, height - 52*mm, "MURLIDHAR ACADEMY")
        
        c.setFont("Helvetica", 12)
        c.setFillColor(colors.black)
        c.drawCentredString(center_x, height - 60*mm, "JUNAGADH")

        # Titles
        c.setFont("Helvetica-Oblique", 18)
        c.setFillColor(colors.black)
        c.drawCentredString(center_x, height - 72*mm, "Certificate of Achievement")

        # Main Result Title
        c.setFont("Helvetica-Bold", 14)
        c.setFillColor(colors.darkgrey)
        c.drawCentredString(center_x, height - 82*mm, f"For: {report_title}")

        c.setFont("Helvetica", 12)
        c.setFillColor(colors.gray)
        c.drawCentredString(center_x, height - 92*mm, "This is proudly presented to")

        # Student Name
        c.setFont("Helvetica-Bold", 32)
        c.setFillColor(theme_color)
        c.drawCentredString(center_x, height - 106*mm, student_name.upper()) 
        c.setStrokeColor(colors.black); c.setLineWidth(0.5)
        c.line(center_x - 60*mm, height - 109*mm, center_x + 60*mm, height - 109*mm)

        # Award Title
        c.setFont("Helvetica-Bold", 20)
        c.setFillColor(colors.black)
        c.drawCentredString(center_x, height - 125*mm, title) 

        # Description
        style = ParagraphStyle('Desc', parent=getSampleStyleSheet()['Normal'], fontName='Helvetica', fontSize=13, leading=16, alignment=TA_CENTER, textColor=colors.darkgray)
        p = Paragraph(desc, style)
        w, h = p.wrap(width - 60*mm, 50*mm)
        p.drawOn(c, (width - w)/2, height - 148*mm) 

        # 4. BOTTOM SECTION
        
        # DATE
        c.setFont("Helvetica-Bold", 12)
        c.setFillColor(colors.black)
        c.drawString(30*mm, 35*mm, f"Date: {cert_date}")
        
        # SIGNATURE BLOCK
        if sign_img:
            img_x = CERT_SIGN_X_POS - (CERT_SIGN_WIDTH / 2)
            c.drawImage(sign_img, img_x, CERT_SIGN_Y_POS, width=CERT_SIGN_WIDTH, height=CERT_SIGN_HEIGHT, mask='auto', preserveAspectRatio=True)

        line_width = 50*mm
        line_start_x = CERT_SIGN_X_POS - (line_width / 2)
        line_end_x = CERT_SIGN_X_POS + (line_width / 2)
        line_y = 35*mm
        
        c.setLineWidth(1)
        c.setStrokeColor(colors.black)
        c.line(line_start_x, line_y, line_end_x, line_y)
        
        c.drawCentredString(CERT_SIGN_X_POS, 29*mm, "Director Signature")

        c.showPage()

    c.save()
    buffer.seek(0)
    return buffer

# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="Murlidhar Academy Report System", page_icon="🎓", layout="centered")
st.title("🎓 Murlidhar Academy Report System")

# Load Main Images
if 'default_bg_data' not in st.session_state:
    st.session_state['default_bg_data'] = download_image_from_drive(DEFAULT_DRIVE_ID)
if 'logo_data' not in st.session_state:
    st.session_state['logo_data'] = download_image_from_drive(LOGO_ID)
if 'sign_data' not in st.session_state:
    st.session_state['sign_data'] = download_image_from_drive(SIGNATURE_ID)

# Load Character Images
if 'char_images' not in st.session_state:
    st.session_state['char_images'] = {}
    with st.spinner("Downloading Award Character Images..."):
        for name, file_id in CHAR_IDS.items():
            img_data = download_image_from_drive(file_id)
            if img_data:
                st.session_state['char_images'][name] = img_data

with st.sidebar:
    st.header("🎨 Settings")
    thresh_green = st.number_input("Green Zone (>= %)", min_value=0, max_value=100, value=70)
    thresh_yellow = st.number_input("Yellow Zone (>= %)", min_value=0, max_value=100, value=40)
    st.markdown("---")
    if st.session_state['default_bg_data']: st.success("✅ Background loaded")
    if st.session_state['logo_data']: st.success("✅ Logo loaded")
    if st.session_state['sign_data']: st.success("✅ Signature loaded")
    if len(st.session_state['char_images']) == 7: st.success("✅ Character Images loaded")

col1, col2 = st.columns(2)
report_header_title = col1.text_input("Main Report Header", f"MONTHLY RESULT REPORT - {datetime.date.today().strftime('%B %Y')}")
output_filename = col2.text_input("Output Filename", f"Monthly_Report_{datetime.date.today().strftime('%b_%Y')}")
summary_page_title = st.text_input("Summary Page Title", "SUMMARY & ANALYSIS OF THE MONTH")

uploaded_files = st.file_uploader("Upload CSV Files", type=['csv'], accept_multiple_files=True)

if uploaded_files:
    per_file_data = []
    for uploaded_file in uploaded_files:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file)
            file_max = find_possible_pts(df)
            names = find_name_series(df)
            obtained = extract_obtained_series(df)
            if file_max is None:
                file_max = obtained.max() if not obtained.empty else DEFAULT_TEST_MAX_PER_FILE
                if file_max == 0: file_max = DEFAULT_TEST_MAX_PER_FILE
            name_map = {}
            present_set = set()
            for n, score in zip(names, obtained):
                norm = normalize_name(n)
                if norm:
                    present_set.add(norm)
                    name_map[norm] = float(score)
            per_file_data.append({"file_max": float(file_max), "data": name_map, "present": present_set})
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")

    if per_file_data:
        all_students = set()
        total_max_marks = sum(f['file_max'] for f in per_file_data)
        total_tests_count = len(per_file_data)
        for f in per_file_data: all_students.update(f['present'])
        
        final_records = []
        for student_norm in all_students:
            total_obtained = 0.0
            tests_present = 0
            display_name = student_norm.title()
            for f in per_file_data:
                if student_norm in f['present']:
                    tests_present += 1
                    total_obtained += f['data'].get(student_norm, 0.0)
            pct = (total_obtained / total_max_marks * 100) if total_max_marks > 0 else 0
            final_records.append({
                "Name": display_name, "Total Tests": total_tests_count, "Present": tests_present,
                "Absent": total_tests_count - tests_present, "Total Marks": int(total_max_marks),
                "Obtained": round(total_obtained, 2), "Percentage": round(pct, 1)
            })
        
        out_df = pd.DataFrame(final_records)
        out_df['Rank'] = out_df['Obtained'].rank(method='dense', ascending=False).astype(int)
        out_df = out_df.sort_values(by=['Rank', 'Name']).reset_index(drop=True)

        st.success("✅ Data Processed! Ready to Generate.")
        st.markdown("### 🗓️ Certificate Settings")
        cert_date_input = st.text_input("Enter Date for Certificate (DD-MM-YYYY)", value=datetime.date.today().strftime('%d-%m-%Y'))

        col_btn1, col_btn2 = st.columns(2)

        if col_btn1.button("📄 Generate Report PDF", type="primary"):
            with st.spinner("Generating Report..."):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=A4)
                PAGE_W, PAGE_H = A4
                TEMPLATE_IMG = None
                if st.session_state['default_bg_data']:
                    st.session_state['default_bg_data'].seek(0)
                    TEMPLATE_IMG = ImageReader(Image.open(st.session_state['default_bg_data']))

                def draw_bg_and_header(c, title_text):
                    if TEMPLATE_IMG: c.drawImage(TEMPLATE_IMG, 0, 0, width=PAGE_W, height=PAGE_H)
                    TITLE_Y = PAGE_H - (TITLE_Y_mm_from_top * mm)
                    c.setFont("Helvetica-Bold", 15)
                    c.setFillColor(colors.white if TEMPLATE_IMG else COLOR_BLUE_HEADER)
                    c.drawCentredString(PAGE_W/2, TITLE_Y, title_text)

                def add_social_links(c):
                    if TEMPLATE_IMG:
                        c.linkURL(TG_LINK, (20*mm, 24*mm, 106*mm, 45*mm))
                        c.linkURL(IG_LINK, (110*mm, 24*mm, 190*mm, 45*mm))

                table_header = ["No", "Rank", "Name", "Tests", "Pres", "Abs", "Max", "Obt", "%"]
                TABLE_WIDTH = PAGE_W - (LEFT_MARGIN_mm * mm) - (RIGHT_MARGIN_mm * mm)
                col_widths = [0.06*TABLE_WIDTH, 0.07*TABLE_WIDTH, 0.35*TABLE_WIDTH, 0.08*TABLE_WIDTH, 0.07*TABLE_WIDTH, 0.07*TABLE_WIDTH, 0.10*TABLE_WIDTH, 0.10*TABLE_WIDTH, 0.10*TABLE_WIDTH]
                
                data_rows = []
                for i, r in out_df.iterrows():
                    data_rows.append([str(i+1), str(r['Rank']), str(r['Name']), str(r['Total Tests']), str(r['Present']), str(r['Absent']), str(r['Total Marks']), str(r['Obtained']), f"{r['Percentage']}%"])

                total_pages_main = math.ceil(len(data_rows) / ROWS_PER_PAGE)
                total_pages_approx = total_pages_main + 2
                TABLE_TOP_Y = PAGE_H - (TITLE_Y_mm_from_top * mm) - (TABLE_SPACE_AFTER_TITLE_mm * mm)

                for p in range(total_pages_main):
                    start = p * ROWS_PER_PAGE
                    end = start + ROWS_PER_PAGE
                    page_data = [table_header] + data_rows[start:end]
                    draw_bg_and_header(c, report_header_title)
                    t = Table(page_data, colWidths=col_widths, repeatRows=1)
                    style = TableStyle([
                        ('GRID', (0,0), (-1,-1), 0.25, colors.HexColor("#666666")), ('BACKGROUND', (0,0), (-1,0), COLOR_BLUE_HEADER),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('ALIGN', (2,1), (2,-1), 'LEFT'),
                        ('LEFTPADDING', (2,1), (2,-1), 6), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('FONTSIZE', (0,0), (-1,-1), 9)
                    ])
                    for i in range(1, len(page_data)):
                        try: pct = float(page_data[i][-1].replace('%',''))
                        except: pct = 0
                        bg = get_smart_row_color(pct, i%2==0, thresh_green, thresh_yellow)
                        style.add('BACKGROUND', (0,i), (-1,i), bg)
                        style.add('TEXTCOLOR', (0,i), (-1,i), colors.black)
                    t.setStyle(style); w, h = t.wrapOn(c, TABLE_WIDTH, PAGE_H)
                    t.drawOn(c, (PAGE_W - TABLE_WIDTH)/2, TABLE_TOP_Y - h)
                    c.setFont("Helvetica-Bold", 8); c.setFillColor(colors.white)
                    c.drawRightString(PAGE_W - (RIGHT_MARGIN_mm*mm), PAGE_NO_Y_mm*mm, f"Page {p+1}/{total_pages_approx}")
                    add_social_links(c); c.showPage()

                # Summary
                draw_bg_and_header(c, summary_page_title)
                avg_obt = out_df['Obtained'].mean()
                pass_count = len(out_df[out_df['Percentage'] >= thresh_yellow])
                summary_data = [
                    ["METRIC", "VALUE", "REMARK"],
                    ["Total Candidates", str(len(out_df)), "Unique students appearing"],
                    ["Total Tests", str(total_tests_count), "Number of CSVs processed"],
                    ["Batch Average", f"{avg_obt:.2f}", "Overall class performance"],
                    ["Highest Score", f"{out_df['Obtained'].max()}", "Top ranker score"],
                    ["Pass Candidates", f"{pass_count}", f">={thresh_yellow}% Score"],
                    ["TOP 5 RANKERS", "", ""],
                ]
                for i, r in out_df.head(5).iterrows(): summary_data.append([f"Rank {r['Rank']}", r['Name'], f"{r['Obtained']}/{total_max_marks} ({r['Percentage']}%)"])
                summary_data.append(["BOTTOM 3 (NEEDS IMPROVEMENT)", "", ""])
                for i, r in out_df.tail(3).sort_values(by='Obtained').iterrows(): summary_data.append([f"Rank {r['Rank']}", r['Name'], f"{r['Obtained']}/{total_max_marks} ({r['Percentage']}%)"])
                
                st_table = Table(summary_data, colWidths=[0.25*TABLE_WIDTH, 0.25*TABLE_WIDTH, 0.50*TABLE_WIDTH])
                sum_style = TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.25, colors.HexColor("#666666")), ('BACKGROUND', (0,0), (-1,0), COLOR_BLUE_HEADER),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('LEFTPADDING', (0,0), (-1,-1), 6), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
                ])
                for i, row in enumerate(summary_data):
                    if i==0: continue
                    sum_style.add('BACKGROUND', (0,i), (-1,i), colors.Color(0.96,0.97,1.0) if i%2==0 else colors.white)
                    if row[0] == "TOP 5 RANKERS": sum_style.add('BACKGROUND', (0,i), (-1,i), SUMMARY_HEADER_COLORS["TOP 5 RANKERS"]); sum_style.add('TEXTCOLOR', (0,i), (-1,i), colors.white)
                    elif row[0] == "BOTTOM 3 (NEEDS IMPROVEMENT)": sum_style.add('BACKGROUND', (0,i), (-1,i), SUMMARY_HEADER_COLORS["BOTTOM 3 (NEEDS IMPROVEMENT)"]); sum_style.add('TEXTCOLOR', (0,i), (-1,i), colors.white)
                st_table.setStyle(sum_style); w, h = st_table.wrapOn(c, TABLE_WIDTH, PAGE_H)
                st_table.drawOn(c, (PAGE_W - TABLE_WIDTH)/2, TABLE_TOP_Y - h)
                c.setFont("Helvetica-Bold", 8); c.setFillColor(colors.white)
                c.drawRightString(PAGE_W - (RIGHT_MARGIN_mm*mm), PAGE_NO_Y_mm*mm, f"Page {total_pages_approx-1}/{total_pages_approx}")
                add_social_links(c); c.showPage()

                # Hall of Fame
                draw_bg_and_header(c, "HALL OF FAME")
                styles = getSampleStyleSheet()
                style_an = ParagraphStyle('AN', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.black, alignment=1)
                style_ad = ParagraphStyle('AD', parent=styles['Normal'], fontName='Helvetica', fontSize=9, textColor=colors.black, alignment=1)
                style_aw = ParagraphStyle('AW', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=COLOR_BLUE_HEADER, alignment=1, leading=12)
                
                awards_list = []
                def mk_para(text, style): return Paragraph(text, style)
                
                rank1 = out_df[out_df['Rank'] == 1]
                if not rank1.empty: awards_list.append([mk_para("Vikramaditya Excellence Award", style_an), mk_para("The Batch Topper (Rank 1). Awarded for ruling the result sheet with the highest score and supreme excellence.", style_ad), mk_para("<br/>".join(rank1['Name'].tolist()), style_aw)])
                
                rank2 = out_df[out_df['Rank'] == 2]
                if not rank2.empty: awards_list.append([mk_para("Chanakya Niti Award", style_an), mk_para("The Intellectual Strategist (Rank 2). Awarded for sharp intelligence and securing the second-highest position.", style_ad), mk_para("<br/>".join(rank2['Name'].tolist()), style_aw)])
                
                rank3 = out_df[out_df['Rank'] == 3]
                if not rank3.empty: awards_list.append([mk_para("Arjuna Focus Award", style_an), mk_para("The Focused Archer (Rank 3). Awarded for unwavering focus, precision, and hitting the target score.", style_ad), mk_para("<br/>".join(rank3['Name'].tolist()), style_aw)])
                
                rank45 = out_df[out_df['Rank'].isin([4, 5])]
                if not rank45.empty: awards_list.append([mk_para("Dhruva Tara Award", style_an), mk_para("The Shining Stars (Rank 4 & 5). Awarded for maintaining a high position consistently like the eternal Pole Star.", style_ad), mk_para("<br/>".join(rank45['Name'].tolist()), style_aw)])
                
                rank6_10 = out_df[out_df['Rank'].isin([6,7,8,9,10])]
                if not rank6_10.empty: awards_list.append([mk_para("Karna Veerta Award", style_an), mk_para("The Brave Warriors (Rank 6 to 10). Talented fighters who fought hard and missed the top 5 by a narrow margin.", style_ad), mk_para("<br/>".join(rank6_10['Name'].tolist()), style_aw)])
                
                angad_c = out_df[(out_df['Absent'] == 0) & (out_df['Percentage'] >= thresh_yellow) & (out_df['Rank'] > 10)].sort_values(by='Obtained', ascending=False).head(5)
                if not angad_c.empty: awards_list.append([mk_para("Angad Stambh Award", style_an), mk_para("The Unmovable Pillar. 100% Attendance & Passing All Tests. They stood firm in every exam!", style_ad), mk_para("<br/>".join(angad_c['Name'].tolist()), style_aw)])
                
                bhagirath_c = out_df[(out_df['Percentage'] >= thresh_yellow) & (out_df['Percentage'] < thresh_green) & (out_df['Present'] / out_df['Total Tests'] >= 0.8) & (out_df['Rank'] > 10) & (out_df['Absent'] > 0)].sort_values(by='Obtained', ascending=False).head(5)
                if not bhagirath_c.empty: awards_list.append([mk_para("Bhagirath Prayas Award", style_an), mk_para("The Relentless Effort. High Attendance & Hard Work. Students striving to turn the tide and improve.", style_ad), mk_para("<br/>".join(bhagirath_c['Name'].tolist()), style_aw)])

                aw_table = Table([["AWARD CATEGORY", "DESCRIPTION", "WINNER(S)"]] + awards_list, colWidths=[0.35*TABLE_WIDTH, 0.35*TABLE_WIDTH, 0.30*TABLE_WIDTH])
                aw_style = TableStyle([('GRID', (0,0), (-1,-1), 0.25, colors.HexColor("#666666")), ('BACKGROUND', (0,0), (-1,0), COLOR_BLUE_HEADER), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('FONT', (0,0), (-1,0), 'Helvetica-Bold'), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 6), ('RIGHTPADDING', (0,0), (-1,-1), 6), ('TOPPADDING', (0,0), (-1,-1), 8), ('BOTTOMPADDING', (0,0), (-1,-1), 8)])
                for i in range(1, len(awards_list)+1): aw_style.add('BACKGROUND', (0,i), (-1,i), colors.Color(0.96,0.97,1.0) if i%2==0 else colors.white)
                aw_table.setStyle(aw_style); w, h = aw_table.wrapOn(c, TABLE_WIDTH, PAGE_H)
                aw_table.drawOn(c, (PAGE_W - TABLE_WIDTH)/2, TABLE_TOP_Y - h)
                c.setFont("Helvetica-Bold", 8); c.setFillColor(colors.white)
                c.drawRightString(PAGE_W - (RIGHT_MARGIN_mm*mm), PAGE_NO_Y_mm*mm, f"Page {total_pages_approx}/{total_pages_approx}")
                add_social_links(c); c.showPage()
                c.save(); buffer.seek(0)
                
                final_pdf = output_filename.strip()
                if not final_pdf.endswith('.pdf'): final_pdf += ".pdf"
                st.download_button(label=f"📥 Download {final_pdf}", data=buffer, file_name=final_pdf, mime="application/pdf")

        # --- BUTTON 2: CERTIFICATES ---
        if col_btn2.button("🏆 Generate Certificates PDF", type="secondary"):
            with st.spinner("Generating Certificates..."):
                logo_bytes = st.session_state['logo_data'] if 'logo_data' in st.session_state else None
                sign_bytes = st.session_state['sign_data'] if 'sign_data' in st.session_state else None
                
                # Pass the character images dict
                char_images = st.session_state['char_images'] if 'char_images' in st.session_state else {}

                cert_buffer = generate_certificates_pdf(out_df, thresh_yellow, thresh_green, report_header_title, cert_date_input, logo_bytes, sign_bytes, char_images)
                cert_name = f"Certificates_{output_filename.strip()}.pdf"
                st.download_button(label=f"📥 Download Certificates", data=cert_buffer, file_name=cert_name, mime="application/pdf")
