import streamlit as st
import pandas as pd
import re
import math
import io
import datetime
import requests
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.utils import ImageReader

# ---------------- CONFIG ----------------
TG_LINK = "https://t.me/MurlidharAcademy"
IG_LINK = "https://www.instagram.com/murlidhar_academy_official/"

# ✅ Default Google Drive Background Image
DEFAULT_DRIVE_ID = "1a1ZK5uiLl0a63Pto1EQDUY0VaIlqp21u"

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
def download_default_bg(file_id):
    try:
        url = get_drive_url(file_id)
        response = requests.get(url, allow_redirects=True)
        if response.status_code == 200:
            try:
                img = Image.open(io.BytesIO(response.content))
                img.verify()
                return io.BytesIO(response.content)
            except Exception:
                return None
        else:
            return None
    except:
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

# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="Consolidated Monthly Report", page_icon="📅", layout="centered")
st.title("📅 Consolidated Monthly Report Generator")

if 'default_bg_data' not in st.session_state:
    st.session_state['default_bg_data'] = download_default_bg(DEFAULT_DRIVE_ID)

with st.sidebar:
    st.header("🎨 Color Grading Rules")
    # ✅ Updated Defaults: 70 and 40
    thresh_green = st.number_input("Green Zone (Excellent >= %)", min_value=0, max_value=100, value=70)
    thresh_yellow = st.number_input("Yellow Zone (Average >= %)", min_value=0, max_value=100, value=40)
    st.markdown("---")
    if st.session_state['default_bg_data']:
        st.success("✅ Background loaded from Drive")
    else:
        st.warning("⚠️ Background loading failed")

col1, col2 = st.columns(2)
report_header_title = col1.text_input("Main Report Header (PDF Top)", f"MONTHLY RESULT REPORT - {datetime.date.today().strftime('%B %Y')}")
output_filename = col2.text_input("Output Filename", f"Monthly_Report_{datetime.date.today().strftime('%b_%Y')}")

summary_page_title = st.text_input("Summary Page Title", "SUMMARY & ANALYSIS OF THE MONTH")

uploaded_files = st.file_uploader("Upload CSV Files (Select Multiple)", type=['csv'], accept_multiple_files=True)

if uploaded_files:
    if st.button("Generate Consolidated PDF 🚀", type="primary"):
        with st.spinner("Processing files and compiling data..."):
            
            per_file_data = []
            for uploaded_file in uploaded_files:
                try:
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
                            
                    per_file_data.append({
                        "file_max": float(file_max),
                        "data": name_map,
                        "present": present_set
                    })
                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: {e}")

            if not per_file_data:
                st.stop()

            all_students = set()
            total_max_marks = sum(f['file_max'] for f in per_file_data)
            total_tests_count = len(per_file_data)
            
            for f in per_file_data:
                all_students.update(f['present'])
            
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
                    "Name": display_name,
                    "Total Tests": total_tests_count,
                    "Present": tests_present,
                    "Absent": total_tests_count - tests_present,
                    "Total Marks": int(total_max_marks),
                    "Obtained": round(total_obtained, 2),
                    "Percentage": round(pct, 1)
                })
            
            out_df = pd.DataFrame(final_records)
            out_df['Rank'] = out_df['Obtained'].rank(method='dense', ascending=False).astype(int)
            out_df = out_df.sort_values(by=['Rank', 'Name']).reset_index(drop=True)
            
            # 3. PDF GENERATION
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=A4)
            PAGE_W, PAGE_H = A4
            
            TEMPLATE_IMG = None
            if st.session_state['default_bg_data']:
                try:
                    st.session_state['default_bg_data'].seek(0)
                    TEMPLATE_IMG = ImageReader(Image.open(st.session_state['default_bg_data']))
                except: pass

            def draw_bg_and_header(c, title_text):
                if TEMPLATE_IMG:
                    c.drawImage(TEMPLATE_IMG, 0, 0, width=PAGE_W, height=PAGE_H)
                
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
            col_widths = [
                0.06 * TABLE_WIDTH, 0.07 * TABLE_WIDTH, 0.35 * TABLE_WIDTH,
                0.08 * TABLE_WIDTH, 0.07 * TABLE_WIDTH, 0.07 * TABLE_WIDTH,
                0.10 * TABLE_WIDTH, 0.10 * TABLE_WIDTH, 0.10 * TABLE_WIDTH
            ]

            data_rows = []
            for i, r in out_df.iterrows():
                row = [
                    str(i + 1), str(r['Rank']), str(r['Name']),
                    str(r['Total Tests']), str(r['Present']), str(r['Absent']),
                    str(r['Total Marks']), str(r['Obtained']), f"{r['Percentage']}%"
                ]
                data_rows.append(row)

            total_pages_main = math.ceil(len(data_rows) / ROWS_PER_PAGE)
            # Estimate pages: Main + Summary + Hall of Fame (could be 1 or 2 pages)
            # We'll just calculate total pages at the end or use placeholders.
            # Simple approach: main + 2 fixed
            total_pages_approx = total_pages_main + 2
            
            TABLE_TOP_Y = PAGE_H - (TITLE_Y_mm_from_top * mm) - (TABLE_SPACE_AFTER_TITLE_mm * mm)

            for p in range(total_pages_main):
                start = p * ROWS_PER_PAGE
                end = start + ROWS_PER_PAGE
                page_data = [table_header] + data_rows[start:end]
                
                draw_bg_and_header(c, report_header_title)
                
                t = Table(page_data, colWidths=col_widths, repeatRows=1)
                
                style = TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.25, colors.HexColor("#666666")),
                    ('BACKGROUND', (0,0), (-1,0), COLOR_BLUE_HEADER),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('ALIGN', (2,1), (2,-1), 'LEFT'),
                    ('LEFTPADDING', (2,1), (2,-1), 6),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('FONTSIZE', (0,0), (-1,-1), 9),
                ])
                
                for i in range(1, len(page_data)):
                    try: pct_val = float(page_data[i][-1].replace('%',''))
                    except: pct_val = 0
                    bg_color = get_smart_row_color(pct_val, i%2==0, thresh_green, thresh_yellow)
                    style.add('BACKGROUND', (0,i), (-1,i), bg_color)
                    style.add('TEXTCOLOR', (0,i), (-1,i), colors.black)
                
                t.setStyle(style)
                w, h = t.wrapOn(c, TABLE_WIDTH, PAGE_H)
                t.drawOn(c, (PAGE_W - TABLE_WIDTH)/2, TABLE_TOP_Y - h)
                
                c.setFont("Helvetica-Bold", 8)
                c.setFillColor(colors.white)
                c.drawRightString(PAGE_W - (RIGHT_MARGIN_mm*mm), PAGE_NO_Y_mm*mm, f"Page {p+1}/{total_pages_approx}")
                add_social_links(c)
                c.showPage()

            # --- SUMMARY PAGE ---
            draw_bg_and_header(c, summary_page_title)
            
            avg_obt = out_df['Obtained'].mean()
            pass_count = len(out_df[out_df['Percentage'] >= thresh_yellow])
            
            summary_data = [
                ["METRIC", "VALUE", "REMARK"],
                ["Total Candidates", str(len(out_df)), "Unique students appearing"],
                ["Total Tests", str(total_tests_count), "Number of CSVs processed"],
                ["Total Marks", str(total_max_marks), "Sum of all test max marks"],
                ["Batch Average", f"{avg_obt:.2f}", "Overall class performance"],
                ["Highest Score", f"{out_df['Obtained'].max()}", "Top ranker score"],
                ["Lowest Score", f"{out_df['Obtained'].min()}", "Needs attention"],
                ["Pass Candidates", f"{pass_count}", f">={thresh_yellow}% Score"],
            ]
            
            summary_data.append(["TOP 5 RANKERS", "", ""])
            top_5 = out_df.head(5)
            for i, r in top_5.iterrows():
                summary_data.append([f"Rank {r['Rank']}", r['Name'], f"{r['Obtained']}/{total_max_marks} ({r['Percentage']}%)"])

            summary_data.append(["BOTTOM 3 (NEEDS IMPROVEMENT)", "", ""])
            bot_3 = out_df.tail(3).sort_values(by='Obtained')
            for i, r in bot_3.iterrows():
                summary_data.append([f"Rank {r['Rank']}", r['Name'], f"{r['Obtained']}/{total_max_marks} ({r['Percentage']}%)"])

            st_table = Table(summary_data, colWidths=[0.25*TABLE_WIDTH, 0.25*TABLE_WIDTH, 0.50*TABLE_WIDTH])
            
            sum_style = TableStyle([
                ('GRID', (0,0), (-1,-1), 0.25, colors.HexColor("#666666")),
                ('BACKGROUND', (0,0), (-1,0), COLOR_BLUE_HEADER),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('LEFTPADDING', (0,0), (-1,-1), 6),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ])
            
            for i, row in enumerate(summary_data):
                if i==0: continue
                
                base_bg = colors.Color(0.96,0.97,1.0) if i % 2 == 0 else colors.white
                sum_style.add('BACKGROUND', (0,i), (-1,i), base_bg)
                
                first_col = str(row[0])
                if first_col in SUMMARY_HEADER_COLORS:
                    sum_style.add('BACKGROUND', (0,i), (-1,i), SUMMARY_HEADER_COLORS[first_col])
                    sum_style.add('TEXTCOLOR', (0,i), (-1,i), colors.white)
                    sum_style.add('FONT', (0,i), (-1,i), 'Helvetica-Bold')
                else:
                    val = str(row[1])
                    try:
                        cell_color = None
                        pct = 0
                        if '(' in val and '%' in val: 
                            pct = float(val.split('(')[1].replace('%)',''))
                        elif '/' in val: 
                            num = float(val.split('/')[0])
                            pct = (num / total_max_marks) * 100 if total_max_marks > 0 else 0
                        
                        if pct > 0:
                            if pct >= thresh_green: cell_color = colors.HexColor("#C8E6C9")
                            elif pct >= thresh_yellow: cell_color = colors.HexColor("#FFF9C4")
                            else: cell_color = colors.HexColor("#FFCDD2")
                        
                        if cell_color:
                            sum_style.add('BACKGROUND', (1,i), (1,i), cell_color)
                    except: pass
            
            st_table.setStyle(sum_style)
            w, h = st_table.wrapOn(c, TABLE_WIDTH, PAGE_H)
            st_table.drawOn(c, (PAGE_W - TABLE_WIDTH)/2, TABLE_TOP_Y - h)
            
            c.setFont("Helvetica-Bold", 8)
            c.setFillColor(colors.white)
            c.drawRightString(PAGE_W - (RIGHT_MARGIN_mm*mm), PAGE_NO_Y_mm*mm, f"Page {total_pages_approx-1}/{total_pages_approx}")
            add_social_links(c)
            c.showPage()

            # --- HALL OF FAME PAGE ---
            draw_bg_and_header(c, "HALL OF FAME")

            styles = getSampleStyleSheet()
            style_award_name = ParagraphStyle('AN', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.black, alignment=1)
            style_desc = ParagraphStyle('AD', parent=styles['Normal'], fontName='Helvetica', fontSize=9, textColor=colors.black, alignment=1)
            style_winner = ParagraphStyle('AW', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=COLOR_BLUE_HEADER, alignment=1, leading=12)

            awards_list = []
            
            def mk_para(text, style): return Paragraph(text, style)

            # 1. Vikramaditya (Rank 1)
            rank1 = out_df[out_df['Rank'] == 1]
            if not rank1.empty:
                names_str = "<br/>".join(rank1['Name'].tolist())
                awards_list.append([
                    mk_para("Vikramaditya Excellence Award", style_award_name),
                    mk_para("For Highest Score (Rank 1). The King of the batch.", style_desc),
                    mk_para(names_str, style_winner)
                ])
            
            # 2. Chanakya (Rank 2)
            rank2 = out_df[out_df['Rank'] == 2]
            if not rank2.empty:
                names_str = "<br/>".join(rank2['Name'].tolist())
                awards_list.append([
                    mk_para("Chanakya Niti Award", style_award_name),
                    mk_para("For Outstanding Intelligence & Strategy (Rank 2).", style_desc),
                    mk_para(names_str, style_winner)
                ])
            
            # 3. Arjuna (Rank 3)
            rank3 = out_df[out_df['Rank'] == 3]
            if not rank3.empty:
                names_str = "<br/>".join(rank3['Name'].tolist())
                awards_list.append([
                    mk_para("Arjuna Focus Award", style_award_name),
                    mk_para("For Unwavering Focus & Precision (Rank 3).", style_desc),
                    mk_para(names_str, style_winner)
                ])

            # 4. Dhruva (Rank 4 & 5)
            rank45 = out_df[out_df['Rank'].isin([4, 5])]
            if not rank45.empty:
                 names_str = "<br/>".join(rank45['Name'].tolist())
                 awards_list.append([
                    mk_para("Dhruva Tara Award", style_award_name),
                    mk_para("Consistently Shining Star (Rank 4 & 5).", style_desc),
                    mk_para(names_str, style_winner)
                ])

            # 5. Karna (Rank 6-10) - NEW
            rank6_10 = out_df[out_df['Rank'].isin([6,7,8,9,10])]
            if not rank6_10.empty:
                 names_str = "<br/>".join(rank6_10['Name'].tolist())
                 awards_list.append([
                    mk_para("Karna Veerta Award", style_award_name),
                    mk_para("The Unsung Warriors (Rank 6 to 10). Fighters who just missed the top.", style_desc),
                    mk_para(names_str, style_winner)
                ])

            # 6. Angad Stambh (100% Present + Passed + NOT in Top 10) - NEW
            # Logic: Absent=0, Percentage >= Yellow Threshold, Rank > 10
            angad_candidates = out_df[
                (out_df['Absent'] == 0) & 
                (out_df['Percentage'] >= thresh_yellow) & 
                (out_df['Rank'] > 10)
            ]
            if not angad_candidates.empty:
                 names_str = "<br/>".join(angad_candidates['Name'].tolist())
                 awards_list.append([
                    mk_para("Angad Stambh Award", style_award_name),
                    mk_para("Unmovable Consistency. 100% Attendance & Passed in all.", style_desc),
                    mk_para(names_str, style_winner)
                ])

            # 7. Bhagirath Prayas (Hard Work) - NEW
            # Logic: Percentage between 40-70, Attendance > 80%, Not in Top 10, Not Angad
            # To simplify: % >= 40 AND % < 70 AND Attendance > 80%
            bhagirath_candidates = out_df[
                (out_df['Percentage'] >= thresh_yellow) & 
                (out_df['Percentage'] < thresh_green) &
                (out_df['Present'] / out_df['Total Tests'] >= 0.8) &
                (out_df['Rank'] > 10) & 
                (out_df['Absent'] > 0) # Use Absent>0 to differentiate from Angad
            ]
            if not bhagirath_candidates.empty:
                 names_str = "<br/>".join(bhagirath_candidates['Name'].tolist())
                 awards_list.append([
                    mk_para("Bhagirath Prayas Award", style_award_name),
                    mk_para("Extraordinary Effort. Regular attendance with improving marks.", style_desc),
                    mk_para(names_str, style_winner)
                ])

            award_header = ["AWARD CATEGORY", "DESCRIPTION", "WINNER(S)"]
            table_data = [award_header] + awards_list
            
            aw_widths = [0.35*TABLE_WIDTH, 0.35*TABLE_WIDTH, 0.30*TABLE_WIDTH]
            aw_table = Table(table_data, colWidths=aw_widths)
            
            aw_style = TableStyle([
                ('GRID', (0,0), (-1,-1), 0.25, colors.HexColor("#666666")),
                ('BACKGROUND', (0,0), (-1,0), COLOR_BLUE_HEADER),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('LEFTPADDING', (0,0), (-1,-1), 6),
                ('RIGHTPADDING', (0,0), (-1,-1), 6),
                ('TOPPADDING', (0,0), (-1,-1), 8),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ])

            for i in range(1, len(table_data)):
                base_bg = colors.Color(0.96,0.97,1.0) if i%2==0 else colors.white
                aw_style.add('BACKGROUND', (0,i), (-1,i), base_bg)
            
            aw_table.setStyle(aw_style)
            w, h = aw_table.wrapOn(c, TABLE_WIDTH, PAGE_H)
            
            # Auto-page break logic for awards if list is too long (Simple version: drawing on current page)
            # Given limited awards, it usually fits on one page.
            aw_table.drawOn(c, (PAGE_W - TABLE_WIDTH)/2, TABLE_TOP_Y - h)

            c.setFont("Helvetica-Bold", 8)
            c.setFillColor(colors.white)
            c.drawRightString(PAGE_W - (RIGHT_MARGIN_mm*mm), PAGE_NO_Y_mm*mm, f"Page {total_pages_approx}/{total_pages_approx}")
            add_social_links(c)
            
            c.showPage()
            c.save()
            buffer.seek(0)
            
            final_pdf_name = output_filename.strip()
            if not final_pdf_name.endswith('.pdf'): final_pdf_name += ".pdf"
            
            st.success("✅ Consolidated Report Generated Successfully!")
            st.download_button(label=f"📥 Download {final_pdf_name}", data=buffer, file_name=final_pdf_name, mime="application/pdf")
