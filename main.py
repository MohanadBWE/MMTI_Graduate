import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from rapidfuzz import process, fuzz
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from PIL import Image
import os
import base64
import re
from docxtpl.richtext import RichText
import time

# --- CONFIGURATION ---
MALE_TEMPLATE = "male_template.docx"
FEMALE_TEMPLATE = "female_template.docx"
EXCEL_FILE = "graduate_data.xlsx"
PHOTO_DIR = "photo_uploads"
GENERATED_DOCS_DIR = "generated_docs"
ID_CARD_DIR = "id_card_uploads"
APPOINTMENT_LOG = "appointments_log.csv"
LOGO_LEFT_PATH = "mmti.webp"
LOGO_RIGHT_PATH = "ntu.webp"

# Use a simple password directly or from secrets if available
try:
    EMPLOYEE_PASSWORD = st.secrets.get("passwords", {}).get("employee", "123")
except (KeyError, FileNotFoundError):
    EMPLOYEE_PASSWORD = "123" # Fallback for local development

# Appointment settings
TIME_SLOTS = [
    ("09:00", "10:00"), ("10:00", "11:00"), ("11:00", "12:00"),
    ("13:30", "14:30"), ("14:30", "15:00")
]
MAX_PER_SLOT = 20
MAX_PER_DAY = 100

# --- SETUP ---
os.makedirs(PHOTO_DIR, exist_ok=True)
os.makedirs(GENERATED_DOCS_DIR, exist_ok=True)
os.makedirs(ID_CARD_DIR, exist_ok=True)

# --- CORE FUNCTIONS ---

def normalize_arabic_name(name):
    """Cleans and standardizes an Arabic name for better matching."""
    if not isinstance(name, str): return ""
    name = re.sub(r'^(ال)', '', name)
    name = re.sub(r'[أإآ]', 'ا', name)
    name = re.sub(r'ة$', 'ه', name)
    name = re.sub(r'[^ا-ي\s]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def load_student_data():
    """Loads student data from the local Excel file."""
    try:
        df = pd.read_excel(EXCEL_FILE)
        if 'full_name' in df.columns:
            df['normalized_name_match'] = df['full_name'].astype(str).apply(lambda x: normalize_arabic_name(x).replace(" ", ""))
        return df
    except FileNotFoundError:
        st.error(f"Error: The student data file '{EXCEL_FILE}' was not found.")
        return None
    except Exception as e:
        st.error(f"An error occurred while reading the Excel file: {e}")
        return None

def match_name(input_name, df):
    """Finds the best match for a student's name."""
    if df is None or 'normalized_name_match' not in df.columns:
        st.error("Student data could not be loaded or is missing the 'full_name' column.")
        return None
    normalized_input = normalize_arabic_name(input_name).replace(" ", "")
    names = df['normalized_name_match'].dropna().tolist()
    matches = process.extract(normalized_input, names, limit=1, score_cutoff=90, scorer=fuzz.partial_ratio)
    if matches:
        best_match_name = matches[0][0]
        matched_row = df[df['normalized_name_match'] == best_match_name]
        if not matched_row.empty:
            return matched_row.iloc[0]
    return None

def get_available_slot():
    """Finds the next available day and time slot from the local CSV file."""
    try:
        log_df = pd.read_csv(APPOINTMENT_LOG)
        log_df['date'] = pd.to_datetime(log_df['date']).dt.date
    except FileNotFoundError:
        log_df = pd.DataFrame(columns=["name", "date", "slot"])

    check_date = datetime.today().date() + timedelta(days=1)
    
    while True:
        day_log = log_df[log_df['date'] == check_date] if not log_df.empty else pd.DataFrame()
        if len(day_log) < MAX_PER_DAY:
            for start, end in TIME_SLOTS:
                slot = f"{start}-{end}"
                slot_count = len(day_log[day_log['slot'] == slot]) if not day_log.empty else 0
                if slot_count < MAX_PER_SLOT:
                    return slot, check_date, log_df
        check_date += timedelta(days=1)

def log_appointment(name, slot, date, log_df):
    """Logs a new appointment to the local CSV file."""
    new_row = pd.DataFrame([{"name": name, "date": date.strftime('%Y-%m-%d'), "slot": slot}])
    updated_log = pd.concat([log_df, new_row], ignore_index=True)
    updated_log.to_csv(APPOINTMENT_LOG, index=False)

def generate_certificate(student_data, destination, grad_date, photo_file, gender):
    template_path = MALE_TEMPLATE if gender == "Male" else FEMALE_TEMPLATE
    try:
        doc = DocxTemplate(template_path)
    except Exception as e:
        st.error(f"Error loading template: {e}"); return None
    def get_value(key): return "" if pd.isna(student_data.get(key)) else str(student_data[key])
    style_name = 'pt_bold heading'
    img = Image.open(photo_file)
    safe_full_name = "".join(x for x in get_value('full_name') if x.isalnum())
    img_path = os.path.join(PHOTO_DIR, f"{safe_full_name}_photo.png")
    img.save(img_path)
    image_for_template = InlineImage(doc, img_path, width=Cm(3.5))
    context = {'full_name': RichText(get_value("full_name"), style=style_name), 'type_of_study': RichText(get_value("type_of_study"), style=style_name), 'department': RichText(get_value("department"), style=style_name), 'section': RichText(get_value("section"), style=style_name), 'average': RichText(get_value("average"), style=style_name), 'appreciation': RichText(get_value("appreciation"), style=style_name), 'rank': RichText(get_value("rank"), style=style_name), 'total': RichText(get_value("total"), style=style_name), 'top_rank': RichText(get_value("top_rank"), style=style_name), 'destination': RichText(destination, style=style_name), 'grad_date': RichText(grad_date, style=style_name), 'photo': image_for_template}
    try:
        doc.render(context)
        file_name = f"{safe_full_name}_certificate.docx"
        docx_path = os.path.join(GENERATED_DOCS_DIR, file_name)
        doc.save(docx_path)
        return docx_path
    except Exception as e:
        st.error(f"Failed to render document: {e}"); return None

# --- UI AND STYLING ---

def get_image_as_base64(path):
    try:
        with open(path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except FileNotFoundError:
        st.error(f"Logo file not found: {path}.")
        return None

def apply_custom_styling():
    primary_color = "#003366"
    secondary_color = "#D4AF37"
    background_color = "#F0F2F6"
    text_color = "#31333F"
    sidebar_widget_color = "#B22222"
    light_primary = "#E0E7FF"
    form_bg_color = "#FFFFFF"
    custom_css = f"""
        <style>
            .app-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 2rem; }}
            .app-header h1 {{ color: {primary_color}; text-align: center; font-size: 2.5rem; margin: 0; }}
            .app-header img {{ width: 120px; }}
            .stApp {{ background-color: {background_color}; }}
            [data-testid="stAppViewContainer"] {{ color: {text_color}; }}
            [data-testid="stAppViewContainer"] h1, 
            [data-testid="stAppViewContainer"] h2, 
            [data-testid="stAppViewContainer"] h3 {{ color: {primary_color}; }}
            [data-testid="stAlert"] * {{ color: #004085; }}
            [data-testid="stSidebar"] {{ background-color: {primary_color}; }}
            [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] label,
            [data-testid="stSidebar"] div[data-baseweb="select"] > div > div {{ color: white !important; }}
            [data-testid="stSidebar"] div[data-baseweb="select"] > div {{ background-color: {sidebar_widget_color}; border: 2px solid {secondary_color}; }}
            [data-testid="stSidebar"] div[data-baseweb="select"] svg {{ fill: {secondary_color}; }}
            [data-baseweb="popover"] ul {{ background-color: {sidebar_widget_color}; }}
            [data-baseweb="popover"] ul li {{ color: white !important; }}
            [data-baseweb="popover"] ul li:hover {{ background-color: {secondary_color}; color: {primary_color} !important; }}
            [data-testid="stForm"] {{ background-color: {form_bg_color}; border: 1px solid #E0E0E0; border-radius: 10px; padding: 25px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }}
            .stTextInput>div>div>input, .stTextArea>div>div>textarea {{ background-color: {form_bg_color}; border: 1px solid #A9A9A9; box-shadow: inset 0 1px 2px rgba(0,0,0,0.07); border-radius: 5px; }}
            .stTextInput>div>div>input:focus, .stTextArea>div>div>textarea:focus {{ border-color: {primary_color}; box-shadow: 0 0 0 2px {light_primary}; }}
            [data-testid="stFileUploader"] section {{ background-color: #F0F2F6; border: 2px dashed #A9A9A9; }}
            [data-testid="stFileUploader"] section:hover {{ border-color: {primary_color}; }}
            [data-testid="stFileUploader"] section [data-testid="stText"] {{ color: {text_color}; }}
            .stButton>button {{ background-color: {primary_color}; color: white; border: 2px solid {secondary_color}; border-radius: 8px; padding: 10px 24px; font-weight: bold; }}
            .stButton>button:hover, .stDownloadButton>button:hover {{ background-color: {secondary_color}; color: {primary_color} !important; border-color: {primary_color}; }}
            @media (max-width: 768px) {{
                .app-header {{ flex-direction: column; gap: 1rem; }}
                .app-header h1 {{ font-size: 1.8rem !important; }}
                [data-testid="stForm"] {{ padding: 15px; }}
            }}
        </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)
    logo_left_b64, logo_right_b64 = get_image_as_base64(LOGO_LEFT_PATH), get_image_as_base64(LOGO_RIGHT_PATH)
    if logo_left_b64 and logo_right_b64:
        st.markdown(f'<div class="app-header"><img src="data:image/webp;base64,{logo_left_b64}"><h1>نظام طلب وثيقة التخرج</h1><img src="data:image/webp;base64,{logo_right_b64}"></div>', unsafe_allow_html=True)

# --- PAGE RENDERING ---

def render_student_view():
    student_df = load_student_data()
    if student_df is None:
        return

    st.info("""**ملاحظات هامة عند استلام تأييد التخرج:**\n1. حضور الطالب شخصياً...\n2. جلب نسخة مصورة من البطاقة الموحدة.\n3. وصل تسديد اجور تحديث البيانات في البرنامج الوزاري SIS.\n4. جلب وصل بمبلغ الف دينار من الشعبة المالية كأجور تأييد التخرج.""")

    with st.form("cert_form"):
        st.header("نموذج طلب وثيقة")
        name = st.text_input("الاسم الكامل للطالب (كما في القوائم الرسمية)")
        gender = st.radio("الجنس:", ("Male", "Female"), horizontal=True)
        destination = st.text_input("الجهة المستفيدة من الوثيقة")
        photo = st.file_uploader("الصورة الشخصية", type=["jpg", "jpeg", "png"])
        id_card_front = st.file_uploader("ارفع صورة وجه الهوية (للتأكيد)", type=["jpg", "jpeg", "png"])
        id_card_back = st.file_uploader("ارفع صورة ظهر الهوية (للتأكيد)", type=["jpg", "jpeg", "png"])
        agreement = st.checkbox("أتعهد بإحضار المستمسكات المطلوبة معي")
        submitted = st.form_submit_button("إرسال الطلب")

    if submitted:
        if not agreement:
            st.error("يرجى الموافقة على التعهد للمتابعة.")
            return
        if not all([name, destination, photo, id_card_front, id_card_back]):
            st.error("يرجى ملء جميع الحقول و إرفاق كافة الصور المطلوبة.")
            return
        
        with st.spinner("...جاري التحقق من الطلب"):
            safe_name = re.sub(r'[^A-Za-z0-9ا-ي]', '_', name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Save ID cards to a local (temporary) folder
            id_card_front_bytes = id_card_front.getvalue()
            id_filename_front = f"{safe_name}_{timestamp}_front.png"
            id_filepath_front = os.path.join(ID_CARD_DIR, id_filename_front)
            with open(id_filepath_front, "wb") as f:
                f.write(id_card_front_bytes)

            id_card_back_bytes = id_card_back.getvalue()
            id_filename_back = f"{safe_name}_{timestamp}_back.png"
            id_filepath_back = os.path.join(ID_CARD_DIR, id_filename_back)
            with open(id_filepath_back, "wb") as f:
                f.write(id_card_back_bytes)
            
            time.sleep(3) # Simulate processing time
        
        st.success("✅ تم التحقق من الطلب.")

        with st.spinner("...جاري البحث عن بيانات الطالب"):
            matched_student = match_name(name, student_df)

        if matched_student is None:
            st.error("الاسم غير موجود في قاعدة البيانات.")
        else:
            st.success(f"تم العثور على الطالب: {matched_student['full_name']}")
            
            with st.spinner("...جاري إصدار الوثيقة وحجز الموعد"):
                slot, appointment_date, log_df = get_available_slot()
                if not slot:
                    st.warning("عذراً، جميع المواعيد محجوزة حالياً.")
                    return
                
                grad_date_str = datetime.now().strftime("%d-%m-%Y")
                doc_path = generate_certificate(matched_student, destination, grad_date_str, photo, gender)
                
                if doc_path:
                    log_appointment(matched_student["full_name"], slot, appointment_date, log_df)
                    appointment_date_str = appointment_date.strftime('%Y-%m-%d')
                    st.success(f"✅ تم تقديم طلبك بنجاح. موعدك للمراجعة هو: {slot} بتاريخ {appointment_date_str}")

def render_employee_view():
    st.header("Employee Dashboard")
    password = st.text_input("Enter Password", type="password", label_visibility="collapsed", placeholder="Enter Password")
    if password == EMPLOYEE_PASSWORD:
        st.success("Access Granted")
        
        st.warning("Note: Files listed below are temporary and will be deleted when the app restarts. Please download them daily.")
        
        # Section for Generated Certificates
        st.subheader("Generated Certificates")
        try:
            cert_files = os.listdir(GENERATED_DOCS_DIR)
            if not cert_files: 
                st.info("No certificates have been generated in this session.")
            else:
                for file in sorted(cert_files, reverse=True):
                    file_path = os.path.join(GENERATED_DOCS_DIR, file)
                    with open(file_path, "rb") as f:
                        st.download_button(label=f"Download {file}", data=f, file_name=file)
        except Exception as e: 
            st.error(f"Could not read certificates directory: {e}")
            
        # Section for Uploaded ID Cards
        st.subheader("Uploaded ID Cards")
        try:
            id_files = os.listdir(ID_CARD_DIR)
            if not id_files: 
                st.info("No ID cards have been uploaded in this session.")
            else:
                for file in sorted(id_files, reverse=True):
                    file_path = os.path.join(ID_CARD_DIR, file)
                    with open(file_path, "rb") as f:
                        st.download_button(label=f"Download {file}", data=f, file_name=file)
        except Exception as e: 
            st.error(f"Could not read ID card directory: {e}")

    elif password:
        st.error("Incorrect password.")

# --- MAIN APP LOGIC ---
st.set_page_config(page_title="نظام طلب وثيقة التخرج", page_icon=LOGO_LEFT_PATH, layout="wide")
apply_custom_styling() 
st.sidebar.markdown('<h2 style="color: #D4AF37;">Portal Navigation</h2>', unsafe_allow_html=True)
app_mode = st.sidebar.selectbox("Choose your role:", ["Student Application", "Employee Dashboard"])

if app_mode == "Student Application":
    render_student_view()
elif app_mode == "Employee Dashboard":
    render_employee_view()
