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
import gspread

# --- CONFIGURATION ---
MALE_TEMPLATE = "male_template.docx"
FEMALE_TEMPLATE = "female_template.docx"
PHOTO_DIR = "photo_uploads"
GENERATED_DOCS_DIR = "generated_docs"
ID_CARD_DIR = "id_card_uploads"
LOGO_LEFT_PATH = "mmti.webp"
LOGO_RIGHT_PATH = "ntu.webp"

try:
    EMPLOYEE_PASSWORD = st.secrets.get("passwords", {}).get("employee", "123")
    SPREADSHEET_NAME = st.secrets.get("app_config", {}).get("spreadsheet_name", "")
except (KeyError, FileNotFoundError):
    st.error("Required secrets not found. Please configure them for deployment.")
    EMPLOYEE_PASSWORD, SPREADSHEET_NAME = "123", ""

TIME_SLOTS = [
    ("09:00", "10:00"), ("10:00", "11:00"), ("11:00", "12:00"),
    ("13:30", "14:30"), ("14:30", "15:00")
]
MAX_PER_SLOT, MAX_PER_DAY = 20, 100

os.makedirs(PHOTO_DIR, exist_ok=True)
os.makedirs(GENERATED_DOCS_DIR, exist_ok=True)
os.makedirs(ID_CARD_DIR, exist_ok=True)

# --- CORE FUNCTIONS ---

@st.cache_resource
def get_gsheets_client():
    """Initializes and returns the gspread client."""
    try:
        creds = st.secrets["gcp_service_account"]
        return gspread.service_account_from_dict(creds)
    except Exception as e:
        st.error(f"Failed to connect to Google Sheets. Check secrets. Error: {e}")
        return None

@st.cache_data(ttl=300)
def load_student_data(_gc_client):
    """Loads student data from the Google Sheet."""
    if _gc_client is None: return None
    try:
        spreadsheet = _gc_client.open(SPREADSHEET_NAME)
        worksheet = spreadsheet.worksheet("Sheet1")
        df = pd.DataFrame(worksheet.get_all_records())
        df = df.dropna(how="all")
        if 'full_name' in df.columns:
            df['normalized_name_match'] = df['full_name'].astype(str).apply(lambda x: normalize_arabic_name(x).replace(" ", ""))
        return df
    except Exception as e:
        st.error(f"Failed to read student data from Google Sheet. Error: {e}")
        return None

def normalize_arabic_name(name):
    if not isinstance(name, str): return ""
    name = re.sub(r'^(ال)', '', name)
    name = re.sub(r'[أإآ]', 'ا', name)
    name = re.sub(r'ة$', 'ه', name)
    name = re.sub(r'[^ا-ي\s]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def match_name(input_name, df):
    if df is None or 'normalized_name_match' not in df.columns: return None
    normalized_input = normalize_arabic_name(input_name).replace(" ", "")
    names = df['normalized_name_match'].dropna().tolist()
    matches = process.extract(normalized_input, names, limit=1, score_cutoff=90, scorer=fuzz.partial_ratio)
    if matches:
        best_match_name = matches[0][0]
        matched_row = df[df['normalized_name_match'] == best_match_name]
        if not matched_row.empty:
            return matched_row.iloc[0]
    return None

def get_available_slot(gsheets_client):
    if gsheets_client is None: return None, None
    try:
        spreadsheet = gsheets_client.open(SPREADSHEET_NAME)
        worksheet = spreadsheet.worksheet("Appointments")
        records = worksheet.get_all_records()
        log_df = pd.DataFrame(records)
        if not log_df.empty:
            log_df = log_df.dropna(how="all")
            log_df['date'] = pd.to_datetime(log_df['date']).dt.date
    except gspread.exceptions.WorksheetNotFound:
        st.error("The 'Appointments' tab was not found in your Google Sheet.")
        return None, None
    except Exception:
        log_df = pd.DataFrame(columns=["name", "date", "slot"])
    check_date = datetime.today().date() + timedelta(days=1)
    while True:
        day_log = log_df[log_df['date'] == check_date] if not log_df.empty else pd.DataFrame()
        if len(day_log) < MAX_PER_DAY:
            for start, end in TIME_SLOTS:
                slot = f"{start}-{end}"
                slot_count = len(day_log[day_log['slot'] == slot]) if not day_log.empty else 0
                if slot_count < MAX_PER_SLOT:
                    return slot, check_date
        check_date += timedelta(days=1)

def log_appointment(gsheets_client, name, slot, date):
    if gsheets_client is None: return
    try:
        spreadsheet = gsheets_client.open(SPREADSHEET_NAME)
        worksheet = spreadsheet.worksheet("Appointments")
        worksheet.append_row([name, date.strftime('%Y-%m-%d'), slot])
    except Exception as e:
        st.error(f"Failed to save appointment. Error: {e}")

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

def render_student_view(gsheets_client):
    student_df = load_student_data(gsheets_client)
    if student_df is None:
        st.warning("Connecting to database... Check secrets if this persists.")
        return

    st.info("""**ملاحظات هامة عند استلام تأييد التخرج:**\n1. حضور الطالب شخصياً...\n2. جلب نسخة مصورة من البطاقة الموحدة.\n3. وصل تسديد اجور تحديث البيانات في البرنامج الوزاري SIS.\n4. جلب وصل بمبلغ الف دينار من الشعبة المالية كأجور تأييد التخرج.""")

    with st.form("cert_form"):
        st.header("نموذج طلب وثيقة")
        name = st.text_input("الاسم الكامل للطالب (كما في القوائم الرسمية)")
        gender = st.radio("الجنس:", ("Male", "Female"), horizontal=True)
        destination = st.text_input("الجهة المستفيدة من الوثيقة")
        photo = st.file_uploader("الصورة الشخصية", type=["jpg", "jpeg", "png"])
        id_card_front = st.file_uploader("ارفع صورة وجه الهوية", type=["jpg", "jpeg", "png"])
        id_card_back = st.file_uploader("ارفع صورة ظهر الهوية", type=["jpg", "jpeg", "png"])
        agreement = st.checkbox("أتعهد بإحضار المستمسكات المطلوبة معي")
        submitted = st.form_submit_button("إرسال الطلب")

    if submitted:
        if not agreement:
            st.error("يرجى الموافقة على التعهد للمتابعة.")
            return
        if not all([name, destination, photo, id_card_front, id_card_back]):
            st.error("يرجى ملء جميع الحقول و إرفاق كافة الصور المطلوبة.")
            return
        
        with st.spinner("...جاري حفظ الملفات مؤقتاً"):
            safe_name = re.sub(r'[^A-Za-z0-9ا-ي]', '_', name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            # This section saves ID cards locally for temporary use.
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
        
        with st.spinner("...جاري البحث عن بيانات الطالب"):
            matched_student = match_name(name, student_df)

        if matched_student is None:
            st.error("الاسم غير موجود في قاعدة البيانات.")
        else:
            st.success(f"تم العثور على الطالب: {matched_student['full_name']}")
            
            with st.spinner("...جاري إصدار الوثيقة وحجز الموعد"):
                slot, appointment_date = get_available_slot(gsheets_client)
                if not slot:
                    st.warning("عذراً، جميع المواعيد محجوزة حالياً.")
                    return
                
                grad_date_str = datetime.now().strftime("%d-%m-%Y")
                doc_path = generate_certificate(matched_student, destination, grad_date_str, photo, gender)
                
                if doc_path:
                    # Log appointment to the secure Google Sheet
                    log_appointment(gsheets_client, matched_student["full_name"], slot, appointment_date)
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

# Get the client at the start
gsheets_client = get_gsheets_client()

if app_mode == "Student Application":
    if gsheets_client:
        render_student_view(gsheets_client)
    else:
        st.error("Application is not configured correctly. Missing or invalid Google credentials in secrets.")
elif app_mode == "Employee Dashboard":
    render_employee_view()
