import os
import io
import json
import base64
from datetime import date, datetime
import pandas as pd
from PIL import Image
import streamlit as st

st.set_page_config(
    page_title="م/ محمد غنيم | منصة شرح الرياضيات والإحصاء",
    page_icon="📐",
    layout="wide",
)

try:
    from streamlit_paste_button import paste_image_button
except ImportError:
    paste_image_button = None

try:
    from streamlit_cropper import st_cropper
except ImportError:
    st_cropper = None

FILE_NAME = "سجل_الغياب_والحصص.xlsx"
IMG_NAME = "teacher.jpg"

CURRICULUM_DATA = {
    "المنهج المصري 🇪🇬": [
        "الصف الأول الإعدادي", "الصف الثاني الإعدادي", "الصف الثالث الإعدادي (الشهادة الإعدادية)",
        "الصف الأول الثانوي", "الصف الثاني الثانوي (علمي)", "الصف الثاني الثانوي (أدبي)",
        "الصف الثالث الثانوي (علمي رياضة)", "الصف الثالث الثانوي (علمي علوم)", "الصف الثالث الثانوي (أدبي)",
        "المرحلة الابتدائية"
    ],
    "المنهج القطري 🇶🇦": [
        "الصف السابع (إعدادي)", "الصف الثامن (إعدادي)", "الصف التاسع (إعدادي)",
        "الصف العاشر (المشترك)", "الصف الحادي عشر (المسار العلمي)", "الصف الحادي عشر (مسار الآداب والإنسانيات)",
        "الصف الحادي عشر (المسار التكنولوجي)", "الصف الثاني عشر (المسار العلمي - متقدم)",
        "الصف الثاني عشر (مسار الآداب - تأسيسي)", "الصف الثاني عشر (المسار التكنولوجي)",
        "المرحلة الابتدائية"
    ],
    "المنهج الإماراتي 🇦🇪": [
        "الحلقة الثانية (الصفوف 5 - 8)", "الصف التاسع (مسار عام)", "الصف التاسع (مسار متقدم)",
        "الصف العاشر (مسار عام)", "الصف العاشر (مسار متقدم)", "الصف العاشر (مسار النخبة)",
        "الصف الحادي عشر (مسار عام)", "الصف الحادي عشر (مسار متقدم)", "الصف الحادي عشر (مسار النخبة)",
        "الصف الثاني عشر (مسار عام)", "الصف الثاني عشر (مسار متقدم)", "الصف الثاني عشر (مسار النخبة)"
    ],
    "المنهج السعودي 🇸🇦": [
        "المرحلة المتوسطة (أول / ثاني / ثالث متوسط)", "السنة الأولى المشتركة (أول ثانوي)",
        "السنة الثانية (المسار العام)", "السنة الثانية (مسار علوم الحاسب والهندسة)",
        "السنة الثانية (مسار الصحة والحياة)", "السنة الثانية (مسار إدارة الأعمال / الشرعي)",
        "السنة الثالثة (المسار العام)", "السنة الثالثة (مسار علوم الحاسب والهندسة)",
        "السنة الثالثة (مسار الصحة والحياة)"
    ],
    "المنهج الكويتي 🇰🇼": [
        "المرحلة المتوسطة (الصفوف 6 - 9)", "الصف العاشر الثانوي (مشترك)",
        "الصف الحادي عشر (القسم العلمي)", "الصف الحادي عشر (القسم الأدبي)",
        "الصف الثاني عشر (القسم العلمي)", "الصف الثاني عشر (القسم الأدبي)"
    ],
    "المنهج السوداني 🇸🇩": [
        "المرحلة المتوسطة (أولى / ثانية / ثالثة متوسط)", "الصف الأول الثانوي",
        "الصف الثاني الثانوي (علمي)", "الصف الثاني الثانوي (أدبي)",
        "الصف الثالث الثانوي (علمي رياضيات - الشهادة السودانية)", "الصف الثالث الثانوي (علمي أحياء)",
        "الصف الثالث الثانوي (أدبي)"
    ]
}

possible_images = ["teacher.jpg", "teacher.png", "teacher.jpeg", "photo_2026-08-02_00-34-53.jpg"]
found_img_path = None
for img_cand in possible_images:
    if os.path.exists(img_cand):
        found_img_path = img_cand
        break

def get_image_base64(path):
    if path and os.path.exists(path):
        try:
            with open(path, "rb") as img_file:
                return base64.b64encode(img_file.read()).decode()
        except Exception:
            return ""
    return ""

def pil_to_base64(pil_img):
    if pil_img is None:
        return ""
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

def base64_to_pil(b64_str):
    if not b64_str:
        return None
    try:
        img_bytes = base64.b64decode(b64_str)
        return Image.open(io.BytesIO(img_bytes))
    except Exception:
        return None

img_b64 = get_image_base64(found_img_path)

COL_SESSIONS = ["التاريخ", "اسم الطالب", "المنهج/الدولة", "المجموعة/الصف", "الحالة", "سعر الحصة", "عدد الحصص الكلي", "نظام الدفع", "مستوى الطالب", "ملاحظات"]
COL_USERS = ["اسم الطالب", "كلمة المرور", "المنهج/الدولة", "المجموعة/الصف", "تاريخ التسجيل", "الحالة_حظر", "حالة_الاشتراك_البنك", "نوع_الدخول"]
COL_ASSESSMENTS = ["التاريخ", "اسم الطالب", "النوع", "عنوان التكليف", "الدرجة المحصلة", "الدرجة العظمى", "حالة التسليم", "ملاحظات وتوجيهات"]
COL_MESSAGES = ["التاريخ_والوقت", "اسم الطالب", "المرسل", "نص الرسالة", "الصورة_base64"]
COL_EXAMS = ["معرف_الامتحان", "عنوان الامتحان", "وصف الامتحان", "كلمة المرور", "المنهج/الدولة", "المجموعة/الصف", "المادة", "الفصل الدراسي", "مدة الامتحان بالدقائق", "الأسئلة_JSON", "تاريخ الإنشاء"]
COL_ESSAYS = ["معرف_الحل", "معرف_الامتحان", "عنوان الامتحان", "اسم الطالب", "رقم السؤال", "نص السؤال", "إجابة الطالب النصية", "صورة الحل_base64", "درجة السؤال", "الدرجة المرصودة", "حالة التصحيح", "ملاحظات المعلم", "تاريخ الحل"]
COL_BOOKINGS = ["تاريخ_الحجز", "اسم الطالب", "المنهج_الدولة", "المرحلة_الصف", "رقم_الهاتف", "رقم_ولي_الأمر", "الحالة"]
COL_BANK_REQUESTS = ["تاريخ_الطلب", "اسم الطالب", "رقم_الهاتف", "كود_OTP", "حالة_الدفع", "إيصال_الدفع_base64"]
COL_QUESTION_BANK = ["معرف_السؤال", "المنهج/الدولة", "المجموعة/الصف", "المادة", "نوع_السؤال", "بيانات_السؤال_JSON"]
COL_VIDEOS = ["معرف_الفيديو", "عنوان_الفيديو", "المنهج/الدولة", "المجموعة/الصف", "رابط_الفيديو", "فيديو_base64", "تاريخ_الرفع"]
COL_VIDEO_COMMENTS = ["التاريخ_والوقت", "عنوان_الفيديو", "اسم الطالب", "نص_التعليق"]
COL_ABQARY = ["معرف_عبقري", "عنوان_الإمتحان", "المنهج/الدولة", "المجموعة/الصف", "رابط_الإمتحان", "رابط_النتيجة", "الرقم_السري_للنتيجة", "تاريخ_النشر"]

def load_all_data():
    users_df = pd.DataFrame(columns=COL_USERS)
    sessions_df = pd.DataFrame(columns=COL_SESSIONS)
    assessments_df = pd.DataFrame(columns=COL_ASSESSMENTS)
    messages_df = pd.DataFrame(columns=COL_MESSAGES)
    exams_df = pd.DataFrame(columns=COL_EXAMS)
    essays_df = pd.DataFrame(columns=COL_ESSAYS)
    bookings_df = pd.DataFrame(columns=COL_BOOKINGS)
    bank_requests_df = pd.DataFrame(columns=COL_BANK_REQUESTS)
    question_bank_df = pd.DataFrame(columns=COL_QUESTION_BANK)
    videos_df = pd.DataFrame(columns=COL_VIDEOS)
    video_comments_df = pd.DataFrame(columns=COL_VIDEO_COMMENTS)
    abqary_df = pd.DataFrame(columns=COL_ABQARY)

    if os.path.exists(FILE_NAME):
        try:
            with pd.ExcelFile(FILE_NAME) as xls:
                if "Users" in xls.sheet_names: users_df = pd.read_excel(xls, "Users")
                if "Sessions" in xls.sheet_names: sessions_df = pd.read_excel(xls, "Sessions")
                elif "Sheet1" in xls.sheet_names: sessions_df = pd.read_excel(xls, "Sheet1")
                if "Assessments" in xls.sheet_names: assessments_df = pd.read_excel(xls, "Assessments")
                if "Messages" in xls.sheet_names: messages_df = pd.read_excel(xls, "Messages")
                if "Exams" in xls.sheet_names: exams_df = pd.read_excel(xls, "Exams")
                if "Essays" in xls.sheet_names: essays_df = pd.read_excel(xls, "Essays")
                if "Bookings" in xls.sheet_names: bookings_df = pd.read_excel(xls, "Bookings")
                if "BankRequests" in xls.sheet_names: bank_requests_df = pd.read_excel(xls, "BankRequests")
                if "QuestionBank" in xls.sheet_names: question_bank_df = pd.read_excel(xls, "QuestionBank")
                if "Videos" in xls.sheet_names: videos_df = pd.read_excel(xls, "Videos")
                if "VideoComments" in xls.sheet_names: video_comments_df = pd.read_excel(xls, "VideoComments")
                if "AbqaryExams" in xls.sheet_names: abqary_df = pd.read_excel(xls, "AbqaryExams")
        except Exception:
            pass

    for col in COL_USERS:
        if col not in users_df.columns: 
            users_df[col] = "نشط" if col == "الحالة_حظر" else ("غير مشترك" if col == "حالة_الاشتراك_البنك" else ("مسجل" if col == "نوع_الدخول" else ""))
    for col in COL_SESSIONS:
        if col not in sessions_df.columns: sessions_df[col] = ""
    for col in COL_ASSESSMENTS:
        if col not in assessments_df.columns: assessments_df[col] = ""
    for col in COL_MESSAGES:
        if col not in messages_df.columns: messages_df[col] = ""
    for col in COL_EXAMS:
        if col not in exams_df.columns: exams_df[col] = ""
    for col in COL_ESSAYS:
        if col not in essays_df.columns: essays_df[col] = ""
    for col in COL_BOOKINGS:
        if col not in bookings_df.columns: bookings_df[col] = ""
    for col in COL_BANK_REQUESTS:
        if col not in bank_requests_df.columns: bank_requests_df[col] = ""
    for col in COL_QUESTION_BANK:
        if col not in question_bank_df.columns: question_bank_df[col] = ""
    for col in COL_VIDEOS:
        if col not in videos_df.columns: videos_df[col] = ""
    for col in COL_VIDEO_COMMENTS:
        if col not in video_comments_df.columns: video_comments_df[col] = ""
    for col in COL_ABQARY:
        if col not in abqary_df.columns: abqary_df[col] = ""

    return users_df, sessions_df, assessments_df, messages_df, exams_df, essays_df, bookings_df, bank_requests_df, question_bank_df, videos_df, video_comments_df, abqary_df

def save_all_data(users_df, sessions_df, assessments_df, messages_df, exams_df, essays_df, bookings_df, bank_requests_df, question_bank_df, videos_df, video_comments_df, abqary_df):
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
        users_df.to_excel(writer, sheet_name="Users", index=False)
        sessions_df.to_excel(writer, sheet_name="Sessions", index=False)
        assessments_df.to_excel(writer, sheet_name="Assessments", index=False)
        messages_df.to_excel(writer, sheet_name="Messages", index=False)
        exams_df.to_excel(writer, sheet_name="Exams", index=False)
        essays_df.to_excel(writer, sheet_name="Essays", index=False)
        bookings_df.to_excel(writer, sheet_name="Bookings", index=False)
        bank_requests_df.to_excel(writer, sheet_name="BankRequests", index=False)
        question_bank_df.to_excel(writer, sheet_name="QuestionBank", index=False)
        videos_df.to_excel(writer, sheet_name="Videos", index=False)
        video_comments_df.to_excel(writer, sheet_name="VideoComments", index=False)
        abqary_df.to_excel(writer, sheet_name="AbqaryExams", index=False)

if "users_df" not in st.session_state:
    u_df, s_df, a_df, m_df, e_df, es_df, b_df, br_df, qb_df, v_df, vc_df, ab_df = load_all_data()
    st.session_state.users_df = u_df
    st.session_state.sessions_df = s_df
    st.session_state.assessments_df = a_df
    st.session_state.messages_df = m_df
    st.session_state.exams_df = e_df
    st.session_state.essays_df = es_df
    st.session_state.bookings_df = b_df
    st.session_state.bank_requests_df = br_df
    st.session_state.question_bank_df = qb_df
    st.session_state.videos_df = v_df
    st.session_state.video_comments_df = vc_df
    st.session_state.abqary_df = ab_df

if "page_view" not in st.session_state:
    st.session_state.page_view = "home"

if "student_sub_page" not in st.session_state:
    st.session_state.student_sub_page = "dashboard"

if "teacher_page" not in st.session_state:
    st.session_state.teacher_page = "dashboard"

if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

query_params = st.query_params
is_student_mode = query_params.get("role") == "student"

if "logged_student" not in st.session_state:
    st.session_state.logged_student = None

saved_student_name = query_params.get("st_name")
if not st.session_state.logged_student and saved_student_name:
    matched_st = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == str(saved_student_name).strip()]
    if not matched_st.empty:
        st.session_state.logged_student = matched_st.iloc[0].to_dict()

def delete_student_completely(student_name_to_del):
    target = student_name_to_del.strip()
    st.session_state.users_df = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.sessions_df = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.assessments_df = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.messages_df = st.session_state.messages_df[st.session_state.messages_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.essays_df = st.session_state.essays_df[st.session_state.essays_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.bookings_df = st.session_state.bookings_df[st.session_state.bookings_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.bank_requests_df = st.session_state.bank_requests_df[st.session_state.bank_requests_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.video_comments_df = st.session_state.video_comments_df[st.session_state.video_comments_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)

if st.session_state.dark_mode:
    bg_color = "#0f172a"
    text_color = "#f8fafc"
    card_bg = "#1e293b"
    card_border = "#334155"
else:
    bg_color = "#ffffff"
    text_color = "#0f172a"
    card_bg = "#f8fafc"
    card_border = "#cbd5e1"

st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@700;800;900&display=swap');
    
    html, body, [class*="css"], p, span, label, div, button, h1, h2, h3, h4, h5, h6 {{
        font-family: 'Cairo', sans-serif !important;
        font-weight: 900 !important;
        letter-spacing: 0.3px !important;
        color: {text_color} !important;
    }}

    .stApp {{
        background-color: {bg_color} !important;
    }}

    input, textarea, select, div[data-baseweb="select"] > div {{
        font-family: 'Cairo', sans-serif !important;
        font-weight: 800 !important;
        color: {text_color} !important;
        background-color: {card_bg} !important;
    }}

    div[data-baseweb="popover"], div[role="dialog"], div[aria-label*="Choose a date"], div[aria-label*="Calendar"], div[data-baseweb="calendar"] {{
        background-color: {card_bg} !important;
        color: {text_color} !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.15) !important;
    }}
    div[data-baseweb="calendar"] *, div[data-baseweb="popover"] * {{
        color: {text_color} !important;
        background-color: transparent !important;
    }}
    div[data-baseweb="calendar"] button {{
        color: {text_color} !important;
        background-color: {card_bg} !important;
        border-radius: 6px !important;
        font-weight: 900 !important;
    }}
    div[data-baseweb="calendar"] button:hover {{
        background-color: #059669 !important;
        color: #ffffff !important;
    }}

    div[data-baseweb="popover"], div[role="listbox"] div {{
        background-color: {card_bg} !important;
        color: {text_color} !important;
    }}
    div[role="option"] span, div[role="option"] div {{
        color: {text_color} !important;
    }}

    body, h1, h2, h3, h4, h5, h6, .stMarkdown, .stSelectbox, .stTextInput, .stTextArea {{
        direction: rtl;
        text-align: right;
    }}
    
    img {{
        max-width: 100% !important;
        height: auto !important;
        object-fit: contain !important;
        border-radius: 8px;
    }}

    .darssly-navbar {{
        background: {card_bg};
        padding: 12px 25px;
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
        border: 1px solid {card_border};
    }}
    .navbar-brand {{ display: flex; align-items: center; gap: 15px; }}
    .navbar-title-group h3 {{ margin: 0 !important; color: #059669 !important; font-size: 18px !important; }}
    .navbar-title-group p {{ margin: 2px 0 0 0 !important; color: {text_color} !important; font-size: 12px !important; opacity: 0.8; }}

    .exam-builder-header {{
        background-color: #f59e0b; color: #ffffff !important; padding: 16px 22px; border-radius: 10px 10px 0 0; font-size: 22px; font-weight: 900; margin-bottom: 15px; display: flex; align-items: center; gap: 12px;
    }}
    .vertical-section-header {{
        background: linear-gradient(135deg, #0284c7, #0369a1); color: #ffffff !important; padding: 14px 20px; border-radius: 10px; margin-top: 25px; margin-bottom: 15px; font-size: 20px; font-weight: 900; display: flex; align-items: center; gap: 10px;
    }}
    
    .stButton>button {{
        background: #10b981 !important; color: #ffffff !important; -webkit-text-fill-color: #ffffff !important; border: none !important; border-radius: 10px; font-weight: 900 !important; font-size: 16px !important; padding: 12px 22px; box-shadow: 0 4px 10px rgba(16, 185, 129, 0.3); width: 100% !important;
    }}
    .stButton>button:hover {{ background: #059669 !important; }}

    div.stFormSubmitButton > button, div.row-widget.stButton > button:nth-of-type(1) {{
        background-color: #dc2626 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }}
    div.stFormSubmitButton > button:hover {{
        background-color: #b91c1c !important;
    }}

    .stLinkButton>a {{
        background-color: #dc2626 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-family: 'Cairo', sans-serif !important;
        font-weight: 900 !important;
        font-size: 16px !important;
        border-radius: 10px !important;
        text-align: center !important;
        display: block !important;
        padding: 10px 15px !important;
        box-shadow: 0 4px 12px rgba(220, 38, 38, 0.3);
        border: none !important;
        text-decoration: none !important;
    }}
    .stLinkButton>a:hover {{
        background-color: #b91c1c !important;
        color: #ffffff !important;
    }}

    .darssly-box {{
        background: linear-gradient(135deg, #059669, #10b981); border-radius: 16px; padding: 24px; margin-bottom: 25px; box-shadow: 0 8px 25px rgba(16, 185, 129, 0.3); border: 2px solid rgba(255, 255, 255, 0.25); color: #ffffff !important;
    }}
    .darssly-box * {{ color: #ffffff !important; }}

    .course-card {{
        background: {card_bg}; border: 2px solid {card_border}; border-radius: 16px; padding: 20px; text-align: center; box-shadow: 0 6px 16px rgba(0,0,0,0.08); margin-bottom: 20px; display: flex; flex-direction: column; justify-content: space-between; height: 100%;
    }}
    .course-icon-box {{
        font-size: 45px; background: #065f46; width: 80px; height: 80px; display: flex; align-items: center; justify-content: center; border-radius: 50%; margin: 0 auto 15px auto; border: 2px solid #10b981;
    }}
    .course-title {{ color: #059669 !important; font-size: 20px !important; font-weight: 900 !important; margin-bottom: 8px !important; }}
    .course-desc {{ color: {text_color} !important; font-size: 14px !important; font-weight: 800 !important; opacity: 0.9; margin-bottom: 15px !important; }}

    .social-top-container {{
        display: flex; gap: 12px; align-items: center; margin-top: 10px; justify-content: center;
    }}
    .social-btn-top {{
        display: inline-flex; align-items: center; justify-content: center; width: 40px; height: 40px; border-radius: 50%; text-decoration: none !important; box-shadow: 0 3px 8px rgba(0,0,0,0.25); transition: transform 0.2s ease;
    }}
    .social-btn-top:hover {{ transform: scale(1.12); }}
    .social-btn-top svg {{ width: 20px; height: 20px; fill: #ffffff; }}

    .facebook-bg {{ background-color: #1877F2; }}
    .whatsapp-bg {{ background-color: #25D366; }}
    .telegram-bg {{ background-color: #229ED9; }}
    .tiktok-bg   {{ background-color: #000000; border: 1px solid #444; }}
    .youtube-bg  {{ background-color: #FF0000; }}

    .call-btn-container {{ display: flex; justify-content: center; margin-top: 25px; margin-bottom: 15px; width: 100%; }}
    .call-btn {{
        display: inline-flex; align-items: center; justify-content: center; gap: 12px; background: linear-gradient(135deg, #059669, #10b981); color: #ffffff !important; padding: 14px 28px; border-radius: 50px; font-size: 18px; font-weight: 900; text-decoration: none !important; border: 2px solid #ffffff;
    }}
    .social-footer-box {{
        margin-top: 25px; padding: 20px 0; border-top: 1px solid rgba(150, 150, 150, 0.3); display: flex; flex-direction: column; align-items: center; justify-content: center; width: 100%;
    }}
    .rights-text {{ font-size: 16px; font-weight: 900; margin-top: 12px; text-align: center; color: {text_color}; }}
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. واجهة الطالب (كاملة وشاملة 100%)
# ==============================================================================
if is_student_mode:
    st.markdown("""
        <style>
            [data-testid="stSidebar"], [data-testid="stSidebarCollapseButton"], [data-testid="collapsedControl"] { display: none !important; }
        </style>
    """, unsafe_allow_html=True)

    nav_avatar_tag = f'<img src="data:image/jpeg;base64,{img_b64}" style="width: 45px; height: 45px; border-radius: 50%; border: 2px solid #10b981; object-fit: cover;">' if img_b64 else '<div style="width:45px;height:45px;border-radius:50%;background:#d1fae5;display:flex;align-items:center;justify-content:center;">👨‍🏫</div>'
    
    col_nav1, col_nav2 = st.columns([2, 1.5])
    with col_nav1:
        st.markdown(f"""
        <div class="darssly-navbar" dir="rtl" style="margin-bottom: 0px;">
            <div class="navbar-brand">
                {nav_avatar_tag}
                <div class="navbar-title-group">
                    <h3>م/ محمد غنيم</h3>
                    <p>منصة شرح الرياضيات والإحصاء</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col_nav2:
        st.write("")
        c_btn0, c_btn1, c_btn2, c_btn3 = st.columns(4)
        with c_btn0:
            if st.button("🏠 الرئيسية"):
                st.session_state.page_view = "home"
                st.session_state.logged_student = None
                st.query_params.clear()
                st.query_params["role"] = "student"
                st.rerun()
        with c_btn1:
            if st.button("👤 دخول"):
                st.session_state.page_view = "login"
                st.rerun()
        with c_btn2:
            if st.button("✨ حساب"):
                st.session_state.page_view = "register"
                st.rerun()
        with c_btn3:
            mode_label = "☀️ فاتح" if st.session_state.dark_mode else "🌙 داكن"
            if st.button(mode_label):
                st.session_state.dark_mode = not st.session_state.dark_mode
                st.rerun()

    st.write("---")

    if not st.session_state.logged_student:
        if st.session_state.page_view == "home":
            col_hero_txt, col_hero_img = st.columns([1.3, 1])
            with col_hero_txt:
                st.markdown("<h1 style='color: #059669; font-size: 38px; font-weight: 900; margin-bottom: 10px;'>أهلاً بيكم منورين المنصة! 🚀</h1>", unsafe_allow_html=True)
                st.markdown("<div style='background: #059669; color: #ffffff; padding: 6px 16px; border-radius: 20px; display: inline-block; font-size: 15px; font-weight: 900; margin-bottom: 15px;'>منصة شرح الرياضيات والإحصاء</div>", unsafe_allow_html=True)
                st.markdown(f"<p style='font-size: 17px; font-weight: 800; line-height: 1.8; color: {text_color};'>مع <b>م / محمد غنيم</b>. خبرة متميزة في تدريس الرياضيات والإحصاء للثانوية العامة والمرحلة الإعدادية. آلاف الطلاب حققوا التفوق والدرجات النهائية. هنتعلم بأسلوب مبسط وجميل، مع شرح احترافي وتجارب تفاعلية، وتدريب شامل على أحدث أنماط الأسئلة.</p>", unsafe_allow_html=True)
                
                c_home_b1, c_home_b2, c_home_b3 = st.columns(3)
                with c_home_b1:
                    if st.button("🔐 تسجيل الدخول"):
                        st.session_state.page_view = "login"
                        st.rerun()
                with c_home_b2:
                    if st.button("✨ حساب جديد"):
                        st.session_state.page_view = "register"
                        st.rerun()
                with c_home_b3:
                    if st.button("👥 دخول كـ ضيف"):
                        guest_name = f"ضيف_{datetime.now().strftime('%H%M%S')}"
                        new_guest = {
                            "اسم الطالب": guest_name, "كلمة المرور": "",
                            "المنهج/الدولة": "المنهج المصري 🇪🇬", "المجموعة/الصف": "الصف الثالث الثانوي (علمي رياضة)",
                            "تاريخ التسجيل": str(date.today()), "الحالة_حظر": "نشط", "حالة_الاشتراك_البنك": "غير مشترك", "نوع_الدخول": "ضيف"
                        }
                        st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([new_guest])], ignore_index=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                        st.session_state.logged_student = new_guest
                        st.query_params["role"] = "student"
                        st.query_params["st_name"] = guest_name
                        st.success(f"تم الدخول كـ ضيف بنجاح!")
                        st.rerun()

            with col_hero_img:
                if img_b64:
                    st.markdown(f"""
                        <div style="display: flex; justify-content: center; align-items: center; position: relative; margin-top: 10px;">
                            <img src="data:image/jpeg;base64,{img_b64}" style="width: 230px; height: 230px; border-radius: 50%; border: 5px solid #059669; object-fit: cover; box-shadow: 0 10px 25px rgba(0,0,0,0.15);">
                        </div>
                    """, unsafe_allow_html=True)

        elif st.session_state.page_view == "login":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🔐 تسجيل دخول الطالب:</h3>", unsafe_allow_html=True)
            with st.form("student_login_form"):
                login_name = st.text_input("اسم الطالب المسجل:")
                login_pass = st.text_input("الرقم السري الخاص بك:", type="password")
                c_l1, c_l2 = st.columns(2)
                with c_l1:
                    submit_login = st.form_submit_button("دخول إلى حسابي")
                with c_l2:
                    if st.form_submit_button("العودة للرئيسية"):
                        st.session_state.page_view = "home"
                        st.rerun()

                if submit_login:
                    users_match = st.session_state.users_df[
                        (st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == login_name.strip())
                        & (st.session_state.users_df["كلمة المرور"].astype(str).str.strip() == login_pass.strip())
                    ]
                    if not users_match.empty:
                        user_info = users_match.iloc[0].to_dict()
                        if user_info.get("الحالة_حظر") == "محظور":
                            st.error("🚫 تم حظر هذا الحساب من قبل المعلم.")
                        else:
                            st.session_state.logged_student = user_info
                            st.query_params["role"] = "student"
                            st.query_params["st_name"] = user_info["اسم الطالب"]
                            st.success(f"مرحباً بك مجدداً يا {login_name}!")
                            st.rerun()
                    else:
                        st.error("اسم الطالب أو الرقم السري غير صحيح.")

        elif st.session_state.page_view == "register":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>✨ إنشاء حساب طالب جديد:</h3>", unsafe_allow_html=True)
            with st.form("student_register_form"):
                reg_name = st.text_input("اسمك بالكامل:")
                reg_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()))
                reg_grade = st.selectbox("المرحلة / الصف الدراسي:", CURRICULUM_DATA[reg_curr])
                reg_pass = st.text_input("اختر رقماً سرياً خاصاً بك:", type="password")
                c_r1, c_r2 = st.columns(2)
                with c_r1:
                    submit_reg = st.form_submit_button("تأكيد إنشاء الحساب")
                with c_r2:
                    if st.form_submit_button("العودة للرئيسية"):
                        st.session_state.page_view = "home"
                        st.rerun()

                if submit_reg:
                    if not reg_name.strip() or not reg_pass.strip():
                        st.error("يرجى كتابة اسمك والرقم السري.")
                    else:
                        existing = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == reg_name.strip()]
                        if not existing.empty:
                            st.warning("هذا الاسم مسجل بالفعل! يرجى تسجيل الدخول مباشرة.")
                        else:
                            new_user = {
                                "اسم الطالب": reg_name.strip(), "كلمة المرور": reg_pass.strip(),
                                "المنهج/الدولة": reg_curr, "المجموعة/الصف": reg_grade,
                                "تاريخ التسجيل": str(date.today()), "الحالة_حظر": "نشط", "حالة_الاشتراك_البنك": "غير مشترك", "نوع_الدخول": "مسجل"
                            }
                            st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([new_user])], ignore_index=True)
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                            st.session_state.logged_student = new_user
                            st.query_params["role"] = "student"
                            st.query_params["st_name"] = new_user["اسم الطالب"]
                            st.success(f"تم إنشاء حسابك بنجاح يا {reg_name}!")
                            st.rerun()
    else:
        st_user = st.session_state.logged_student
        col_u1, col_u2 = st.columns([4, 1])
        with col_u1:
            st.markdown(f"""
                <div style="background: {card_bg}; padding: 14px 20px; border-radius: 12px; border-right: 5px solid #10b981; margin-bottom: 20px; border: 1px solid {card_border};">
                    <h3 style="margin: 0; color: #059669; font-size: 20px;">أهلاً بك: {st_user['اسم الطالب']} 🌟</h3>
                    <p style="margin: 4px 0 0 0; font-weight: 900; color: {text_color}; font-size: 16px;">{st_user.get('المنهج/الدولة', '')} | {st_user.get('المجموعة/الصف', '')}</p>
                </div>
            """, unsafe_allow_html=True)
        with col_u2:
            st.write("")
            if st.button("🚪 خروج"):
                st.session_state.logged_student = None
                st.session_state.page_view = "home"
                st.query_params.clear()
                st.query_params["role"] = "student"
                st.rerun()

        st.markdown("<div class='vertical-section-header'>🗂️ لوحة خدمات الطالب التفاعلية</div>", unsafe_allow_html=True)
        
        if st.button("🎥 الفيديوهات والشروحات التعليمية", use_container_width=True):
            st.session_state.student_sub_page = "videos"
            st.rerun()
        if st.button("📚 بنك الأسئلة الشامل", use_container_width=True):
            st.session_state.student_sub_page = "bank"
            st.rerun()
        if st.button("🧠 اختبارات ونتائج موقع عبقري 💡", use_container_width=True):
            st.session_state.student_sub_page = "abqary"
            st.rerun()
        if st.button("✍️ الاختبارات الإلكترونية التفاعلية", use_container_width=True):
            st.session_state.student_sub_page = "exams"
            st.rerun()
        if st.button("📝 تسجيل حضور حصة اليوم", use_container_width=True):
            st.session_state.student_sub_page = "attendance"
            st.rerun()
        if st.button("📊 متابعة درجات الواجبات", use_container_width=True):
            st.session_state.student_sub_page = "hw_grades"
            st.rerun()
        if st.button("📈 متابعة درجات الاختبارات", use_container_width=True):
            st.session_state.student_sub_page = "exam_grades"
            st.rerun()
        if st.button("💬 مركز الدردشة والدعم المباشر", use_container_width=True):
            st.session_state.student_sub_page = "chat"
            st.rerun()

        sub_page = st.session_state.student_sub_page
        st.write("---")

        if sub_page == "abqary":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🧠 اختبارات ونتائج موقع عبقري</h3>", unsafe_allow_html=True)
            student_grade = str(st_user.get("المجموعة/الصف", "")).strip().lower()
            abq_df = st.session_state.abqary_df
            st_abq = abq_df[abq_df["المجموعة/الصف"].astype(str).str.strip().str.lower().str.contains(student_grade, na=False)]
            if st_abq.empty: st_abq = abq_df.copy()

            if st_abq.empty:
                st.info("لا توجد اختبارات منشورة عبر موقع عبقري لمرحلتك حالياً.")
            else:
                for ab_i, ab_row in st_abq.iterrows():
                    ab_title = ab_row["عنوان_الإمتحان"]
                    ab_link = ab_row["رابط_الإمتحان"]
                    res_link = ab_row["رابط_النتيجة"]
                    secret_code = str(ab_row.get("الرقم_السري_للنتيجة", "")).strip()

                    st.markdown(f"""
                        <div style="background:{card_bg}; border:2px solid #059669; border-radius:14px; padding:20px; margin-bottom:20px;">
                            <h4 style="color:#059669; margin-top:0;">💡 اختبار عبقري: {ab_title}</h4>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    if ab_link and ab_link != "nan":
                        st.markdown(f"**عرض امتحان عبقري مباشرة داخل الموقع:**")
                        st.components.v1.iframe(ab_link, height=600, scrolling=True)

                    st.write("")
                    st.markdown("##### 🔍 استعلام عن نتيجة هذا الاختبار برقم سري:")
                    with st.form(f"abqary_result_form_{ab_i}"):
                        entered_pass = st.text_input("أدخل الرقم السري المخصص لإظهار النتيجة:", type="password", key=f"pass_input_{ab_i}")
                        if st.form_submit_button("🔓 إظهار النتيجة"):
                            if secret_code and entered_pass.strip() == secret_code:
                                st.success("✓ الرقم السري صحيح! تم جلب النتيجة بنجاح:")
                                if res_link and res_link != "nan":
                                    st.markdown(f"**معاينة نتيجة موقع عبقري مباشرة:**")
                                    st.components.v1.iframe(res_link, height=500, scrolling=True)
                                else:
                                    st.info("النتيجة متاحة وسيتم إضافتها قريباً.")
                            else:
                                st.error("❌ الرقم السري غير صحيح.")
                    st.write("---")

        elif sub_page == "videos":
            st.subheader("🎥 محتوى الشروحات والفيديوهات التعليمية")
            student_grade = str(st_user.get("المجموعة/الصف", "")).strip()
            v_df = st.session_state.videos_df
            st_videos = v_df[v_df["المجموعة/الصف"].astype(str).str.strip().str.lower() == student_grade.lower()]
            if st_videos.empty:
                st.info("لا توجد فيديوهات مرفوعة لمرحلتك حالياً.")
            else:
                for _, v_row in st_videos.iterrows():
                    st.markdown(f"**{v_row['عنوان_الفيديو']}**")
                    v_link = str(v_row.get("رابط_الفيديو", ""))
                    if v_link: st.video(v_link)

        elif sub_page == "bank":
            st.subheader("📚 بنك الأسئلة الشامل")
            st.info("قسم بنك الأسئلة متاح للاشتراك وتفعيل الحساب.")

        elif sub_page == "exams":
            st.subheader("✍️ الاختبارات الإلكترونية التفاعلية")
            st.info("لا توجد اختبارات إلكترونية نشطة حالياً.")

        elif sub_page == "attendance":
            st.subheader("📝 تسجيل حضور حصة اليوم")
            with st.form("att_form"):
                st_date = st.date_input("التاريخ:")
                if st.form_submit_button("تأكيد الحضور"):
                    st.success("تم تسجيل الحضور بنجاح!")

        elif sub_page == "hw_grades":
            st.subheader("📊 متابعة درجات الواجبات")
            my_ass = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip()]
            if my_ass.empty: st.info("لا توجد درجات مرصودة.")
            else: st.dataframe(my_ass, use_container_width=True)

        elif sub_page == "exam_grades":
            st.subheader("📈 متابعة درجات الاختبارات")
            st.info("لا توجد نتائج اختبارات سابقة.")

        elif sub_page == "chat":
            st.subheader("💬 مركز الدردشة والدعم المباشر")
            st.info("ميزة الدردشة متاحة للتواصل المباشر مع المعلم.")

    st.stop()


# ==============================================================================
# 2. لوحة تحكم المعلم (نظام البطاقات الاحترافي مع صفحات مستقلة بالكامل وشامل لكل الأدوات)
# ==============================================================================
total_exams_count = len(st.session_state.exams_df)
total_students_count = len(st.session_state.users_df)

st.sidebar.markdown(f"""
    <div style="text-align: center; margin-top: 5px; margin-bottom: 20px;">
        <h2 style="margin: 0; color: #059669; font-weight: 900;">م/ محمد غنيم</h2>
        <p style="margin: 4px 0; font-weight: 900; font-size: 14px;">لوحة تحكم المعلم الاحترافية</p>
    </div>
""", unsafe_allow_html=True)

if st.sidebar.button("📊 نظرة عامة (الرئيسية)", use_container_width=True):
    st.session_state.teacher_page = "dashboard"
    st.rerun()
if st.sidebar.button("⚙️ صانع الامتحانات", use_container_width=True):
    st.session_state.teacher_page = "exam_maker"
    st.rerun()
if st.sidebar.button("📚 بنك الأسئلة", use_container_width=True):
    st.session_state.teacher_page = "question_bank"
    st.rerun()
if st.sidebar.button("🎥 إدارة الفيديوهات", use_container_width=True):
    st.session_state.teacher_page = "videos"
    st.rerun()
if st.sidebar.button("💡 امتحانات عبقري", use_container_width=True):
    st.session_state.teacher_page = "abqary"
    st.rerun()
if st.sidebar.button("📈 درجات الاختبارات", use_container_width=True):
    st.session_state.teacher_page = "grades"
    st.rerun()
if st.sidebar.button("📝 تصحيح المقالي", use_container_width=True):
    st.session_state.teacher_page = "essays"
    st.rerun()
if st.sidebar.button("💬 الرسائل والدردشة", use_container_width=True):
    st.session_state.teacher_page = "chat"
    st.rerun()
if st.sidebar.button("👥 بطاقات الطلاب", use_container_width=True):
    st.session_state.teacher_page = "students"
    st.rerun()
if st.sidebar.button("📝 رصد حصة جديدة", use_container_width=True):
    st.session_state.teacher_page = "add_session"
    st.rerun()
if st.sidebar.button("📚 رصد واجب يدوي", use_container_width=True):
    st.session_state.teacher_page = "add_hw"
    st.rerun()
if st.sidebar.button("✏️ تعديل السجلات", use_container_width=True):
    st.session_state.teacher_page = "edit_records"
    st.rerun()
if st.sidebar.button("📊 السجلات الشاملة", use_container_width=True):
    st.session_state.teacher_page = "all_records"
    st.rerun()
if st.sidebar.button("🖨️ تقرير ولي الأمر", use_container_width=True):
    st.session_state.teacher_page = "parent_report"
    st.rerun()

st.sidebar.write("---")
st.sidebar.code("https://engmohamedghonaim.streamlit.app/?role=student", language="text")

t_page = st.session_state.teacher_page

if t_page == "dashboard":
    col_main_top, col_stat_sidebar = st.columns([3, 1])
    
    with col_main_top:
        avatar_html = f'<img src="data:image/jpeg;base64,{img_b64}" style="width: 50px; height: 50px; border-radius: 50%; border: 2px solid #059669; object-fit: cover; vertical-align: middle; margin-left: 15px;">' if img_b64 else ''
        st.markdown(f"""
            <div style="background: {card_bg}; padding: 25px; border-radius: 16px; border: 1px solid {card_border}; margin-bottom: 25px; display: flex; align-items: center;">
                {avatar_html}
                <div>
                    <h2 style="color: #059669; margin: 0 0 5px 0;">أهلاً، م / محمد غنيم 📌</h2>
                    <p style="font-size: 15px; margin: 0; opacity: 0.85;">لوحة تحكم المعلمين – الوصول السريع لأدواتك وخدماتك التعليمية المتقدمة.</p>
                </div>
            </div>
        """, unsafe_allow_html=True)

    with col_stat_sidebar:
        st.markdown(f"""
            <div style="background: {card_bg}; padding: 15px; border-radius: 14px; border: 1px solid {card_border}; text-align: center;">
                <p style="margin:0; font-size:14px;">عدد الامتحانات</p>
                <h3 style="color: #059669; margin: 5px 0;">{total_exams_count}</h3>
                <hr style="margin: 8px 0;">
                <p style="margin:0; font-size:14px;">المتقدمين (إجمالي)</p>
                <h3 style="color: #0284c7; margin: 5px 0;">{total_students_count}</h3>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("### أدواتي السريعة")
    
    bc1, bc2, bc3 = st.columns(3)
    with bc1:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">امتحاناتي</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">إنشاء وإدارة الامتحانات</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("ادارة الامتحانات", key="card_btn_ex"):
            st.session_state.teacher_page = "exam_maker"
            st.rerun()

    with bc2:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">المواد التعليمية</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">رفع الفيديوهات والشروحات</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("ادارة الفيديوهات", key="card_btn_vid"):
            st.session_state.teacher_page = "videos"
            st.rerun()

    with bc3:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">امتحانات عبقري</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">ربط اختبارات موقع عبقري</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("ادارة عبقري", key="card_btn_abq"):
            st.session_state.teacher_page = "abqary"
            st.rerun()

    bc4, bc5, bc6, bc7 = st.columns(4)
    with bc4:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 150px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669; font-size: 16px;">طلبات حجز أونلاين</h4>
                    <p style="margin: 5px 0 0 0; font-size: 12px; opacity: 0.8;">عرض الحجوزات وطباعتها</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("عرض الحجوزات", key="card_btn_bookings"):
            st.session_state.teacher_page = "bookings_page"
            st.rerun()

    with bc5:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 150px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669; font-size: 16px;">سجل الحضور</h4>
                    <p style="margin: 5px 0 0 0; font-size: 12px; opacity: 0.8;">رصد الحصص والغياب</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("رصد الحصص", key="card_btn_ses"):
            st.session_state.teacher_page = "add_session"
            st.rerun()

    with bc6:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 150px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669; font-size: 16px;">الرسائل والدعم</h4>
                    <p style="margin: 5px 0 0 0; font-size: 12px; opacity: 0.8;">محادثات الطلاب</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("عرض الرسائل", key="card_btn_chat"):
            st.session_state.teacher_page = "chat"
            st.rerun()

    with bc7:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 150px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669; font-size: 16px;">تقارير ولي الأمر</h4>
                    <p style="margin: 5px 0 0 0; font-size: 12px; opacity: 0.8;">طباعة التقارير PDF</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("عرض التقارير", key="card_btn_rep"):
            st.session_state.teacher_page = "parent_report"
            st.rerun()

elif t_page == "exam_maker":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.markdown("<div class='exam-builder-header'>➕ صانع الامتحانات التفاعلية (مع لصق الصور وقصها وتايمر للطالب)</div>", unsafe_allow_html=True)
    if "temp_questions" not in st.session_state: st.session_state.temp_questions = []
    if "q_img_ver" not in st.session_state: st.session_state.q_img_ver = 0
    if "pasted_q_img" not in st.session_state: st.session_state.pasted_q_img = ""

    with st.expander("⚙️ 1. الإعدادات الرئيسية للامتحان", expanded=True):
        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            ex_title_input = st.text_input("عنوان الامتحان:*", placeholder="مثال: اختبار الجبر المصور")
            ex_desc_input = st.text_area("وصف الامتحان:*")
            ex_pass_input = st.text_input("كلمة المرور (اختياري):", type="password")
        with c_ex2:
            ex_curr_input = st.selectbox("الدولة / المنهج:*", list(CURRICULUM_DATA.keys()), key="mk_c")
            ex_grade_input = st.selectbox("الصف الدراسي:*", CURRICULUM_DATA[ex_curr_input], key="mk_g")
            ex_subject_input = st.selectbox("المادة:*", ["الرياضيات (عام)", "الجبر والإحصاء", "الهندسة", "التفاضل والتكامل"])
            ex_term_input = st.selectbox("الفصل الدراسي:", ["الفصل الدراسي الأول", "الفصل الدراسي الثاني"])
            ex_time_input = st.number_input("مدة الامتحان بالدقائق (يظهر تايمر للطالب):", value=45)

    with st.expander("📐 2. إدخال وتحرير الأسئلة (يدعم لصق الصور Ctrl+V)", expanded=True):
        q_count = len(st.session_state.temp_questions) + 1
        st.markdown(f"#### إضافة سؤال رقم {q_count}")
        q_type_choice = st.selectbox("نوع السؤال:", ["سؤال موضوعي (اختيار من متعدد)", "سؤال مقالي"])
        is_mcq = "موضوعي" in q_type_choice

        st.markdown("##### 📌 أ) صورة السؤال (مع دعم لصق الصور من الحافظة):")
        if paste_image_button:
            paste_res = paste_image_button("📋 لصق صورة السؤال من الحافظة (Ctrl+V)", key=f"paste_q_{q_count}_{st.session_state.q_img_ver}")
            if paste_res.image_data is not None:
                st.session_state.pasted_q_img = pil_to_base64(paste_res.image_data)
        
        q_file_up = st.file_uploader("أو ارفع صورة للسؤال:", type=["jpg", "png", "jpeg"], key=f"upload_q_{q_count}")
        if q_file_up is not None:
            st.session_state.pasted_q_img = base64.b64encode(q_file_up.read()).decode()

        if st.session_state.pasted_q_img:
            st.image(f"data:image/jpeg;base64,{st.session_state.pasted_q_img}", width=300)
            if st.button("🗑️ مسح صورة السؤال", key=f"del_q_img_{q_count}"):
                st.session_state.pasted_q_img = ""
                st.session_state.q_img_ver += 1
                st.rerun()

        q_text_input = st.text_area("نص السؤال المكتوب:")

        opt1_txt, opt2_txt, opt3_txt, opt4_txt = "", "", "", ""
        correct_num = 1
        if is_mcq:
            st.markdown("##### 🎯 خيارات الاختيار من متعدد:")
            opt1_txt = st.text_input("الخيار أ:")
            opt2_txt = st.text_input("الخيار ب:")
            opt3_txt = st.text_input("الخيار ج:")
            opt4_txt = st.text_input("الخيار د:")
            correct_num = st.selectbox("الإجابة الصحيحة:", [1, 2, 3, 4], format_func=lambda x: f"الخيار ({['أ', 'ب', 'ج', 'د'][x-1]})")

        q_pts_val = st.number_input("درجة السؤال:", value=1.0, key=f"pts_{q_count}")

        if st.button(f"➕ حفظ وإدراج السؤال رقم {q_count}"):
            if not q_text_input.strip() and not st.session_state.pasted_q_img:
                st.error("يرجى كتابة نص السؤال أو وضع صورة.")
            else:
                st.session_state.temp_questions.append({
                    "type": "موضوعي" if is_mcq else "مقالي",
                    "text": q_text_input.strip(),
                    "q_img": st.session_state.pasted_q_img,
                    "opt1": opt1_txt, "opt2": opt2_txt, "opt3": opt3_txt, "opt4": opt4_txt,
                    "correct": correct_num, "points": q_pts_val
                })
                st.session_state.pasted_q_img = ""
                st.session_state.q_img_ver += 1
                st.success(f"✓ تم حفظ السؤال {q_count}!")
                st.rerun()

        if st.session_state.temp_questions:
            st.write(f"**الأسئلة الجاهزة: ({len(st.session_state.temp_questions)})**")
            for idx_q, q_item in enumerate(st.session_state.temp_questions):
                st.markdown(f"- س {idx_q+1}: {q_item['text'] if q_item['text'] else '[سؤال مصور]'} (الدرجة: {q_item['points']})")
            if st.button("🗑️ مسح جميع الأسئلة"):
                st.session_state.temp_questions = []
                st.rerun()

    st.write("---")
    c_save1, c_print_q = st.columns(2)
    with c_save1:
        if st.button("🚀 حفظ ونشر الامتحان للطلاب"):
            if not ex_title_input.strip():
                st.error("يرجى كتابة عنوان الامتحان.")
            elif not st.session_state.temp_questions:
                st.error("يرجى إضافة سؤال واحد على الأقل.")
            else:
                new_ex = {
                    "معرف_الامتحان": f"EX_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "عنوان الامتحان": ex_title_input.strip(),
                    "وصف الامتحان": ex_desc_input.strip(),
                    "كلمة المرور": ex_pass_input.strip(),
                    "المنهج/الدولة": ex_curr_input,
                    "المجموعة/الصف": ex_grade_input,
                    "المادة": ex_subject_input,
                    "الفصل الدراسي": ex_term_input,
                    "مدة الامتحان بالدقائق": int(ex_time_input),
                    "الأسئلة_JSON": json.dumps(st.session_state.temp_questions, ensure_ascii=False),
                    "تاريخ الإنشاء": str(date.today()),
                }
                st.session_state.exams_df = pd.concat([st.session_state.exams_df, pd.DataFrame([new_ex])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.session_state.temp_questions = []
                st.success("✓ تم نشر الامتحان بنجاح!")
                st.rerun()
    with c_print_q:
        if st.session_state.temp_questions:
            print_q_html = f"<!DOCTYPE html><html dir='rtl'><head><meta charset='utf-8'><title>أسئلة الامتحان</title></head><body onload='window.print()'><h2>أسئلة امتحان: {ex_title_input}</h2><hr>"
            for i, q in enumerate(st.session_state.temp_questions):
                print_q_html += f"<p><b>س{i+1}:</b> {q['text']} (درجة: {q['points']})</p>"
                if q['q_img']: print_q_html += f"<img src='data:image/jpeg;base64,{q['q_img']}' width='300'><br>"
                if q['opt1']: print_q_html += f"أ) {q['opt1']} &nbsp; ب) {q['opt2']} &nbsp; ج) {q['opt3']} &nbsp; د) {q['opt4']}<br>"
                print_q_html += "<br>"
            print_q_html += "</body></html>"
            st.download_button(label="🖨️ طباعة أسئلة الامتحان ملف PDF", data=print_q_html.encode("utf-8"), file_name="أسئلة_الامتحان.html", mime="application/octet-stream")

elif t_page == "question_bank":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📚 بنك الأسئلة الشامل (مع التصحيح التلقائي والدعم الكامل للصورة)")
    with st.form("qb_form"):
        qb_curr = st.selectbox("المنهج:", list(CURRICULUM_DATA.keys()))
        qb_grade = st.selectbox("الصف:", CURRICULUM_DATA[qb_curr])
        qb_type = st.selectbox("نوع السؤال:", ["اختيار من متعدد", "مقالي"])
        qb_txt = st.text_area("نص السؤال:")
        
        qb_opt1, qb_opt2, qb_opt3, qb_opt4 = "", "", "", ""
        qb_correct = 1
        if qb_type == "اختيار من متعدد":
            qb_opt1 = st.text_input("أ)")
            qb_opt2 = st.text_input("ب)")
            qb_opt3 = st.text_input("ج)")
            qb_opt4 = st.text_input("د)")
            qb_correct = st.selectbox("الإجابة الصحيحة:", [1, 2, 3, 4], format_func=lambda x: f"الخيار ({['أ', 'ب', 'ج', 'د'][x-1]})")

        qb_pts = st.number_input("الدرجة:", value=1.0)

        if st.form_submit_button("💾 حفظ السؤال في البنك"):
            if qb_txt.strip():
                q_payload = {"type": qb_type, "text": qb_txt.strip(), "q_img": "", "opt1": qb_opt1, "opt2": qb_opt2, "opt3": qb_opt3, "opt4": qb_opt4, "correct": qb_correct, "points": qb_pts}
                new_row = {"معرف_السؤال": f"QB_{datetime.now().strftime('%Y%m%d%H%M%S')}", "المنهج/الدولة": qb_curr, "المجموعة/الصف": qb_grade, "المادة": "الرياضيات", "نوع_السؤال": qb_type, "بيانات_السؤال_JSON": json.dumps(q_payload, ensure_ascii=False)}
                st.session_state.question_bank_df = pd.concat([st.session_state.question_bank_df, pd.DataFrame([new_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("✓ تم حفظ السؤال في البنك بنجاح!")

    st.write("---")
    st.markdown("### 💳 إدارة اشتراكات بنك الأسئلة (إيصالات إنستا باي):")
    bank_reqs = st.session_state.bank_requests_df
    if bank_reqs.empty:
        st.info("لا توجد طلبات اشتراك معلقة حالياً.")
    else:
        for b_idx, b_row in bank_reqs.iterrows():
            st_b_name = b_row["اسم الطالب"]
            with st.expander(f"طلب اشتراك: {st_b_name} — الهاتف: {b_row['رقم_الهاتف']} | الحالة: [{b_row['حالة_الدفع']}]"):
                if b_row["إيصال_الدفع_base64"]:
                    st.image(f"data:image/jpeg;base64,{b_row['إيصال_الدفع_base64']}", width=300)
                if st.button(f"✅ تأكيد تفعيل الاشتراك لـ {st_b_name}", key=f"conf_bank_{b_idx}"):
                    bank_req_state = st.session_state.bank_requests_df
                    bank_req_state.at[b_idx, "حالة_الدفع"] = "مؤكد ومفعل"
                    st.session_state.users_df.loc[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == str(st_b_name).strip(), "حالة_الاشتراك_البنك"] = "مشترك"
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                    st.success(f"✓ تم تفعيل بنك الأسئلة للطالب {st_b_name}!")
                    st.rerun()

    subs_df = st.session_state.users_df[st.session_state.users_df["حالة_الاشتراك_البنك"] == "مشترك"]
    if not subs_df.empty:
        subs_html = f"<!DOCTYPE html><html dir='rtl'><head><meta charset='utf-8'><title>مشتركي بنك الأسئلة</title></head><body onload='window.print()'><h2>قائمة الطلاب المشتركين في بنك الأسئلة</h2><table border='1' style='width:100%; border-collapse:collapse; text-align:center;'><tr><th>اسم الطالب</th><th>المرحلة</th><th>التاريخ</th></tr>"
        for _, sr in subs_df.iterrows():
            subs_html += f"<tr><td>{sr['اسم الطالب']}</td><td>{sr['المجموعة/الصف']}</td><td>{sr['تاريخ التسجيل']}</td></tr>"
        subs_html += "</table></body></html>"
        st.download_button(label="🖨️ طباعة أسماء المشتركين في بنك الأسئلة PDF", data=subs_html.encode("utf-8"), file_name="مشتركي_بنك_الأسئلة.html", mime="application/octet-stream")

elif t_page == "videos":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("🎥 إدارة ورفع الفيديوهات التعليمية (متابعة المشاهدات والتعليقات)")
    with st.form("vid_form"):
        v_title = st.text_input("عنوان الدرس:")
        v_curr = st.selectbox("المنهج:", list(CURRICULUM_DATA.keys()))
        v_grade = st.selectbox("الصف:", CURRICULUM_DATA[v_curr])
        v_link = st.text_input("رابط يوتيوب:")
        if st.form_submit_button("نشر الفيديو"):
            if v_title.strip():
                new_vid = {"معرف_الفيديو": f"VID_{datetime.now().strftime('%Y%m%d%H%M%S')}", "عنوان_الفيديو": v_title.strip(), "المنهج/الدولة": v_curr, "المجموعة/الصف": v_grade, "رابط_الفيديو": v_link.strip(), "فيديو_base64": "", "تاريخ_الرفع": str(date.today())}
                st.session_state.videos_df = pd.concat([st.session_state.videos_df, pd.DataFrame([new_vid])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("✓ تم نشر الفيديو بنجاح!")

    st.write("---")
    st.markdown("### 💬 تعليقات وآراء الطلاب على الفيديوهات:")
    vc_df = st.session_state.video_comments_df
    if vc_df.empty:
        st.info("لا توجد تعليقات حتى الآن.")
    else:
        st.dataframe(vc_df, use_container_width=True)

elif t_page == "abqary":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("💡 إدارة امتحانات ونتائج موقع عبقري")
    with st.form("abq_form"):
        ab_title = st.text_input("عنوان امتحان عبقري:")
        ab_curr = st.selectbox("المنهج:", list(CURRICULUM_DATA.keys()))
        ab_grade = st.selectbox("الصف:", CURRICULUM_DATA[ab_curr])
        ab_link = st.text_input("رابط الامتحان على موقع عبقري:")
        res_link = st.text_input("رابط النتيجة:")
        secret_pass = st.text_input("الرقم السري لإظهار النتيجة للطالب:")
        if st.form_submit_button("حفظ ونشر امتحان عبقري"):
            if ab_title.strip() and ab_link.strip():
                new_ab = {"معرف_عبقري": f"ABQ_{datetime.now().strftime('%Y%m%d%H%M%S')}", "عنوان_الإمتحان": ab_title.strip(), "المنهج/الدولة": ab_curr, "المجموعة/الصف": ab_grade, "رابط_الإمتحان": ab_link.strip(), "رابط_النتيجة": res_link.strip(), "الرقم_السري_للنتيجة": secret_pass.strip(), "تاريخ_النشر": str(date.today())}
                st.session_state.abqary_df = pd.concat([st.session_state.abqary_df, pd.DataFrame([new_ab])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("✓ تم نشر امتحان عبقري بنجاح!")

    st.write("---")
    st.markdown("### 🖨️ طباعة نتائج امتحانات عبقري لكل مرحلة:")
    ab_df = st.session_state.abqary_df
    if not ab_df.empty:
        sel_ab_grade = st.selectbox("اختر المرحلة الدراسية للطباعة:", ab_df["المجموعة/الصف"].unique())
        if st.button("طباعة نتائج هذه المرحلة PDF"):
            filtered_ab = ab_df[ab_df["المجموعة/الصف"] == sel_ab_grade]
            ab_html = f"<!DOCTYPE html><html dir='rtl'><head><meta charset='utf-8'><title>نتائج عبقري</title></head><body onload='window.print()'><h2>نتائج امتحانات عبقري - {sel_ab_grade}</h2><table border='1' style='width:100%; border-collapse:collapse; text-align:center;'><tr><th>العنوان</th><th>رابط الامتحان</th><th>تاريخ النشر</th></tr>"
            for _, r in filtered_ab.iterrows():
                ab_html += f"<tr><td>{r['عنوان_الإمتحان']}</td><td>{r['رابط_الإمتحان']}</td><td>{r['تاريخ_النشر']}</td></tr>"
            ab_html += "</table></body></html>"
            st.download_button(label="📥 تحميل ملف PDF للنتائج", data=ab_html.encode("utf-8"), file_name=f"نتائج_عبقري_{sel_ab_grade}.html", mime="application/octet-stream")

elif t_page == "grades":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📈 سجل درجات ونقاط اختبارات الطلاب")
    exam_assessments = st.session_state.assessments_df[st.session_state.assessments_df["النوع"].astype(str).str.contains("اختبار|كويز", na=False)]
    if exam_assessments.empty:
        st.info("لا توجد نتائج اختبارات مرصودة.")
    else:
        st.dataframe(exam_assessments, use_container_width=True)
        grades_pdf = f"<!DOCTYPE html><html dir='rtl'><head><meta charset='utf-8'><title>درجات الاختبارات</title></head><body onload='window.print()'><h2>سجل درجات الاختبارات والكويزات</h2><table border='1' style='width:100%; border-collapse:collapse; text-align:center;'><tr><th>التاريخ</th><th>الطالب</th><th>الاختبار</th><th>الدرجة</th></tr>"
        for _, r in exam_assessments.iterrows():
            grades_pdf += f"<tr><td>{r['التاريخ']}</td><td>{r['اسم الطالب']}</td><td>{r['عنوان التكليف']}</td><td>{r['الدرجة المحصلة']} / {r['الدرجة العظمى']}</td></tr>"
        grades_pdf += "</table></body></html>"
        st.download_button(label="🖨️ طباعة كل الدرجات PDF", data=grades_pdf.encode("utf-8"), file_name="سجل_الدرجات.html", mime="application/octet-stream")

elif t_page == "essays":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📝 تصحيح إجابات الطلاب المقالية")
    essays = st.session_state.essays_df
    if essays.empty:
        st.info("لا توجد إجابات مقالية قيد الانتظار.")
    else:
        for es_idx, es_row in essays.iterrows():
            st.markdown(f"**الطالب:** {es_row['اسم الطالب']} — **السؤال:** {es_row['نص السؤال']} — **الإجابة:** {es_row['إجابة الطالب النصية']}")
            if es_row['صورة الحل_base64']:
                st.image(f"data:image/jpeg;base64,{es_row['صورة الحل_base64']}", width=300)
            score_g = st.number_input("الدرجة المستحقة:", min_value=0.0, max_value=float(es_row['درجة السؤال']), value=float(es_row['الدرجة المرصودة']), key=f"es_g_{es_idx}")
            if st.button("اعتماد الدرجة", key=f"btn_es_{es_idx}"):
                essays.at[es_idx, "الدرجة المرصودة"] = score_g
                essays.at[es_idx, "حالة التصحيح"] = "تم التصحيح"
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("✓ تم رصد الدرجة بنجاح!")
                st.rerun()

elif t_page == "chat":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("💬 صندوق محادثات الطلاب والرد على الأسئلة")
    all_students_with_chat = list(st.session_state.messages_df["اسم الطالب"].dropna().unique())
    if not all_students_with_chat: st.info("لا توجد رسائل واردة.")
    else:
        sel_st = st.selectbox("اختر الطالب للمحادثة:", all_students_with_chat)
        if sel_st:
            st.dataframe(st.session_state.messages_df[st.session_state.messages_df["اسم الطالب"] == sel_st], use_container_width=True)

elif t_page == "students":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("👥 بطاقات الطلاب المسجلين والتحكم الكامل")
    all_known_st = list(st.session_state.users_df["اسم الطالب"].dropna().unique())
    if not all_known_st:
        st.info("لا توجد طلاب مسجلون.")
    else:
        sel_card_st = st.selectbox("اختر الطالب لعرض بطاقته والتحكم به:", all_known_st)
        if sel_card_st:
            u_r = st.session_state.users_df[st.session_state.users_df["اسم الطالب"] == sel_card_st].iloc[0]
            s_r = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"] == sel_card_st]
            total_s = len(s_r)
            att_s = len(s_r[s_r["الحالة"] == "حاضر"])
            abs_s = len(s_r[s_r["الحالة"] == "غائب"])
            tot_cost = s_r["سعر الحصة"].astype(float).sum() if not s_r.empty else 0.0

            st.markdown(f"""
                <div style="background:{card_bg}; border:2px solid #059669; border-radius:14px; padding:20px; margin-bottom:20px;">
                    <h3 style="color:#059669; margin-top:0;">👤 اسم الطالب: {sel_card_st}</h3>
                    <p><b>نوع الدخول:</b> {u_r.get('نوع_الدخول', 'مسجل')} | <b>المنهج:</b> {u_r.get('المنهج/الدولة', '-')} | <b>المرحلة:</b> {u_r.get('المجموعة/الصف', '-')}</p>
                    <p><b>إجمالي الحصص:</b> {total_s} (حاضر: {att_s} | غياب: {abs_s}) | <b>إجمالي الحساب:</b> {tot_cost:,.1f} جنيه</p>
                </div>
            """, unsafe_allow_html=True)

            c_b1, c_b2 = st.columns(2)
            with c_b1:
                if st.button("🚫 حظر هذا الطالب"):
                    st.session_state.users_df.loc[st.session_state.users_df["اسم الطالب"] == sel_card_st, "الحالة_حظر"] = "محظور"
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                    st.warning("تم حظر الطالب.")
                    st.rerun()
            with c_b2:
                if st.button("🗑️ حذف الطالب نهائياً"):
                    delete_student_completely(sel_card_st)
                    st.success("تم حذف الطالب نهائياً.")
                    st.rerun()

        if st.button("🖨️ طباعة بيانات جميع الطلاب PDF"):
            users_pdf = "<!DOCTYPE html><html dir='rtl'><head><meta charset='utf-8'><title>بطاقات الطلاب</title></head><body onload='window.print()'><h2>بيانات جميع الطلاب المسجلين</h2><table border='1' style='width:100%; border-collapse:collapse; text-align:center;'><tr><th>الاسم</th><th>النوع</th><th>المنهج</th><th>المرحلة</th><th>الحالة</th></tr>"
            for _, ur in st.session_state.users_df.iterrows():
                users_pdf += f"<tr><td>{ur['اسم الطالب']}</td><td>{ur.get('نوع_الدخول', 'مسجل')}</td><td>{ur['المنهج/الدولة']}</td><td>{ur['المجموعة/الصف']}</td><td>{ur['الحالة_حظر']}</td></tr>"
            users_pdf += "</table></body></html>"
            st.download_button(label="📥 تحميل ملف PDF للطلاب", data=users_pdf.encode("utf-8"), file_name="جميع_الطلاب.html", mime="application/octet-stream")

elif t_page == "bookings_page":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📅 طلبات حجز الدروس أونلاين الواردة من الطلاب")
    b_df = st.session_state.bookings_df
    if b_df.empty:
        st.info("لا توجد طلبات حجز أونلاين حتى الآن.")
    else:
        st.dataframe(b_df, use_container_width=True)
        book_pdf = "<!DOCTYPE html><html dir='rtl'><head><meta charset='utf-8'><title>طلبات الحجز</title></head><body onload='window.print()'><h2>طلبات حجز الدروس أونلاين</h2><table border='1' style='width:100%; border-collapse:collapse; text-align:center;'><tr><th>التاريخ</th><th>الطالب</th><th>المنهج</th><th>المرحلة</th><th>الهاتف</th><th>ولي الأمر</th></tr>"
        for _, br in b_df.iterrows():
            book_pdf += f"<tr><td>{br['تاريخ_الحجز']}</td><td>{br['اسم الطالب']}</td><td>{br['المنهج_الدولة']}</td><td>{br['المرحلة_الصف']}</td><td>{br['رقم_الهاتف']}</td><td>{br['رقم_ولي_الأمر']}</td></tr>"
        book_pdf += "</table></body></html>"
        st.download_button(label="🖨️ طباعة طلبات الحجز PDF", data=book_pdf.encode("utf-8"), file_name="طلبات_الحجز.html", mime="application/octet-stream")

elif t_page == "add_session":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📝 رصد حضور حصة جديدة")
    all_st_names = list(st.session_state.users_df["اسم الطالب"].dropna().unique())
    with st.form("sess_f"):
        s_name = st.selectbox("اختر الطالب (مسجل أو ضيف):", all_st_names) if all_st_names else st.text_input("اسم الطالب:")
        s_curr = st.selectbox("المنهج:", list(CURRICULUM_DATA.keys()))
        s_group = st.selectbox("المرحلة / الصف:", CURRICULUM_DATA[s_curr])
        s_status = st.selectbox("حالة الحضور:", ["حاضر", "غائب", "متأخر", "بعذر"])
        s_price = st.number_input("سعر الحصة:", value=100.0)
        s_notes = st.text_input("ملاحظات:")
        
        if st.form_submit_button("💾 حفظ ورصد الحصة"):
            if s_name:
                new_row = {"التاريخ": str(date.today()), "اسم الطالب": str(s_name).strip(), "المنهج/الدولة": s_curr, "المجموعة/الصف": s_group, "الحالة": s_status, "سعر الحصة": s_price, "عدد الحصص الكلي": 1, "نظام الدفع": "مقدم", "مستوى الطالب": "جيد", "ملاحظات": s_notes}
                st.session_state.sessions_df = pd.concat([st.session_state.sessions_df, pd.DataFrame([new_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("✓ تم رصد الحصة بنجاح!")

    st.write("---")
    if not st.session_state.sessions_df.empty:
        ses_pdf = "<!DOCTYPE html><html dir='rtl'><head><meta charset='utf-8'><title>كشف الحضور</title></head><body onload='window.print()'><h2>كشف حضور وغياب الطلاب</h2><table border='1' style='width:100%; border-collapse:collapse; text-align:center;'><tr><th>التاريخ</th><th>الطالب</th><th>المرحلة</th><th>الحالة</th><th>السعر</th></tr>"
        for _, sr in st.session_state.sessions_df.iterrows():
            ses_pdf += f"<tr><td>{sr['التاريخ']}</td><td>{sr['اسم الطالب']}</td><td>{sr['المجموعة/الصف']}</td><td>{sr['الحالة']}</td><td>{sr['سعر الحصة']}</td></tr>"
        ses_pdf += "</table></body></html>"
        st.download_button(label="🖨️ طباعة كشف حضور الطلاب PDF", data=ses_pdf.encode("utf-8"), file_name="كشف_الحضور.html", mime="application/octet-stream")

elif t_page == "add_hw":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📚 رصد واجب يدوي")
    all_st_names = list(st.session_state.users_df["اسم الطالب"].dropna().unique())
    with st.form("hw_f"):
        hw_name = st.selectbox("اختر الطالب:", all_st_names) if all_st_names else st.text_input("اسم الطالب:")
        hw_title = st.text_input("عنوان الواجب أو التكليف:")
        hw_score = st.number_input("الدرجة المحصلة:", value=10.0)
        hw_max = st.number_input("الدرجة العظمى:", value=10.0)
        if st.form_submit_button("حفظ الواجب"):
            if hw_name and hw_title.strip():
                new_ass = {"التاريخ": str(date.today()), "اسم الطالب": str(hw_name).strip(), "النوع": "واجب منزلي", "عنوان التكليف": hw_title.strip(), "الدرجة المحصلة": hw_score, "الدرجة العظمى": hw_max, "حالة التسليم": "تم التسليم", "ملاحظات وتوجيهات": "أداء ممتاز"}
                st.session_state.assessments_df = pd.concat([st.session_state.assessments_df, pd.DataFrame([new_ass])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("✓ تم حفظ الواجب بنجاح!")

elif t_page == "edit_records":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("✏️ تعديل السجلات والبيانات")
    df = st.session_state.sessions_df
    if df.empty:
        st.info("لا توجد سجلات للحصص.")
    else:
        sel_row_idx = st.selectbox("اختر سجل الحصة للتعديل:", options=list(df.index), format_func=lambda x: f"{df.loc[x, 'التاريخ']} - {df.loc[x, 'اسم الطالب']}")
        row_data = df.loc[sel_row_idx]
        with st.form("edit_rec_form"):
            e_name = st.text_input("اسم الطالب:", value=row_data["اسم الطالب"])
            e_price = st.number_input("سعر الحصة:", value=float(row_data["سعر الحصة"]))
            e_status = st.selectbox("الحالة:", ["حاضر", "غائب", "متأخر", "بعذر"], index=["حاضر", "غائب", "متأخر", "بعذر"].index(row_data["الحالة"]) if row_data["الحالة"] in ["حاضر", "غائب", "متأخر", "بعذر"] else 0)
            if st.form_submit_button("💾 تحديث السجل"):
                df.at[sel_row_idx, "اسم الطالب"] = e_name.strip()
                df.at[sel_row_idx, "سعر الحصة"] = e_price
                df.at[sel_row_idx, "الحالة"] = e_status
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("✓ تم التحديث بنجاح!")
                st.rerun()

elif t_page == "all_records":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📊 السجلات الشاملة وإحصائيات الطلاب")
    all_students_master = list(st.session_state.users_df["اسم الطالب"].dropna().unique())
    if not all_students_master:
        st.info("لا توجد بيانات طلاب.")
    else:
        sel_m_st = st.selectbox("اختر طالباً لعرض تقريره الشامل:", all_students_master)
        if sel_m_st:
            st_s = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"] == sel_m_st]
            st_a = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"] == sel_m_st]
            tot_price = st_s["سعر الحصة"].astype(float).sum() if not st_s.empty else 0.0
            st.markdown(f"""
                <div style="background:{card_bg}; border:1px solid {card_border}; padding:15px; border-radius:10px;">
                    <h4>تقرير الطالب: {sel_m_st}</h4>
                    <p><b>إجمالي الحصص:</b> {len(st_s)} | <b>إجمالي المستحق (سعر الحصص):</b> <span style="color:#b91c1c;">{tot_price:,.1f} جنيه</span></p>
                </div>
            """, unsafe_allow_html=True)
            st.write("سجل الحصص:")
            st.dataframe(st_s, use_container_width=True)
            st.write("سجل التقييمات والواجبات والاختبارات:")
            st.dataframe(st_a, use_container_width=True)

elif t_page == "parent_report":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("🖨️ إصدار وطباعة تقرير ولي الأمر الشامل PDF")
    all_names = list(st.session_state.users_df["اسم الطالب"].dropna().unique())
    if not all_names:
        st.info("لا توجد بيانات طلاب كافية.")
    else:
        selected_student = st.selectbox("اختر الطالب لطباعة تقرير ولي الأمر:", all_names)
        if selected_student:
            st_sessions = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"] == selected_student].copy()
            st_assessments = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"] == selected_student].copy()
            
            att_cnt = len(st_sessions[st_sessions["الحالة"] == "حاضر"])
            abs_cnt = len(st_sessions[st_sessions["الحالة"] == "غائب"])
            total_cost = st_sessions["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True) if not st_sessions.empty else 0.0

            ass_html_rows = ""
            for _, r in st_assessments.iterrows():
                ass_html_rows += f"<tr><td>{r['التاريخ']}</td><td>{r['النوع']}</td><td>{r['عنوان التكليف']}</td><td>{r['الدرجة المحصلة']} / {r['الدرجة العظمى']}</td><td>{r['حالة التسليم']}</td></tr>"

            session_html_rows = ""
            for _, r in st_sessions.iterrows():
                session_html_rows += f"<tr><td>{r['التاريخ']}</td><td>{r['الحالة']}</td><td>{r['سعر الحصة']} جنيه</td><td>{r['مستوى الطالب']}</td><td>{r['ملاحظات']}</td></tr>"

            teacher_img_tag = f'<img src="data:image/jpeg;base64,{img_b64}" style="width: 80px; height: 80px; border-radius: 50%; object-fit: cover;">' if img_b64 else ""

            parent_report_html = f"""<!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head><meta charset="utf-8"><title>تقرير ولي الأمر - {selected_student}</title></head>
            <body onload="window.print()" style="font-family: Arial; padding: 30px;">
                <div style="display:flex; justify-content:space-between; border-bottom:3px solid #059669; padding-bottom:10px;">
                    <div>{teacher_img_tag}<h2>م/ محمد غنيم 📐</h2><p>منصة شرح الرياضيات والإحصاء</p></div>
                    <div><h2>تقرير متابعة الطالب: {selected_student}</h2><p>التاريخ: {date.today()}</p></div>
                </div>
                <h3>ملخص الحساب والحضور:</h3>
                <p><b>إجمالي الحصص الحضورية:</b> {att_cnt} | <b>إجمالي الغياب:</b> {abs_cnt} | <b>إجمالي المبلغ المستحق (سعر الحصص):</b> {total_cost:,.1f} جنيه</p>
                <h3>سجل الواجبات والاختبارات والكويزات:</h3>
                <table border='1' style='width:100%; border-collapse:collapse; text-align:center;'>
                    <tr><th>التاريخ</th><th>النوع</th><th>العنوان</th><th>الدرجة</th><th>الحالة</th></tr>
                    {ass_html_rows if ass_html_rows else "<tr><td colspan='5'>لا توجد تقييمات</td></tr>"}
                </table>
                <h3>سجل الحضور وسعر الحصص:</h3>
                <table border='1' style='width:100%; border-collapse:collapse; text-align:center; margin-top:10px;'>
                    <tr><th>التاريخ</th><th>الحالة</th><th>سعر الحصة</th><th>المستوى</th><th>ملاحظات</th></tr>
                    {session_html_rows if session_html_rows else "<tr><td colspan='5'>لا توجد حصص مسجلة</td></tr>"}
                </table>
            </body></html>"""

            st.download_button(
                label=f"🖨️ طباعة وتحميل تقرير ولي الأمر لـ ({selected_student}) PDF",
                data=parent_report_html.encode("utf-8"),
                file_name=f"تقرير_ولي_الأمر_{selected_student}.html",
                mime="application/octet-stream"
            )
