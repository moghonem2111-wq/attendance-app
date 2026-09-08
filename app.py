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
COL_USERS = ["اسم الطالب", "كلمة المرور", "المنهج/الدولة", "المجموعة/الصف", "تاريخ التسجيل", "الحالة_حظر", "حالة_الاشتراك_البنك"]
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
        if col not in users_df.columns: users_df[col] = "نشط" if col == "الحالة_حظر" else ("غير مشترك" if col == "حالة_الاشتراك_البنك" else "")
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
# 1. واجهة الطالب (كما طلبت تماماً ودون أي نقص)
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
                st.markdown(f"<p style='font-size: 17px; font-weight: 800; line-height: 1.8; color: {text_color};'>مع <b>م / محمد غنيم</b>. خبرة متميزة في تدريس الرياضيات والإحصاء للثانوية العامة والمرحلة الإعدادية. آلاف الطلاب حققوا التفوق والدرجات النهائية. هنتعلم بأسلوب مبسط وجميل، مع شرح احترافي وتجارب تفاعلية، وتدريب شامل على أحدث أنماط الأسئلة عشان تدخل الامتحان وأنت جاهز تحقق أفضل نتيجة.</p>", unsafe_allow_html=True)
                
                st.markdown("""
                    <div class="social-top-container">
                        <a href="https://www.facebook.com/share/19fD41rV3H/" target="_blank" title="Facebook" class="social-btn-top facebook-bg"><svg viewBox="0 0 24 24"><path d="M24 12.073c0-6.627-5.373-12-12-12s-12 5.373-12 12c0 5.99 4.388 10.954 10.125 11.854v-8.385H7.078v-3.47h3.047V9.43c0-3.007 1.792-4.669 4.533-4.669 1.312 0 2.686.235 2.686.235v2.953H15.83c-1.491 0-1.956.925-1.956 1.874v2.25h3.328l-.532 3.47h-2.796v8.385C19.612 23.027 24 18.062 24 12.073z"/></svg></a>
                        <a href="https://wa.me/201016361440" target="_blank" title="WhatsApp" class="social-btn-top whatsapp-bg"><svg viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.711 2.599 2.669-.699c.971.53 1.77.822 2.791.823h.002c3.18 0 5.767-2.586 5.768-5.766 0-3.18-2.587-5.766-5.77-5.766zm9.969 5.828c0 5.519-4.481 10-10 10-1.761 0-3.424-.46-4.881-1.267l-5.619 1.474 1.499-5.485c-.911-1.516-1.43-3.285-1.43-5.176 0-5.519 4.481-10 10-10 5.519 0 10 4.481 10 10z"/></svg></a>
                        <a href="https://t.me/mrmaths22" target="_blank" title="Telegram" class="social-btn-top telegram-bg"><svg viewBox="0 0 24 24"><path d="M12 0c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm5.894 8.221l-1.97 9.28c-.145.658-.537.818-1.084.508l-3-2.21-1.446 1.394c-.16.16-.295.295-.605.295l.213-3.053 5.56-5.023c.242-.213-.054-.333-.373-.121l-6.871 4.326-2.962-.924c-.643-.204-.657-.643.136-.953l11.57-4.461c.537-.194 1.006.131.832.942z"/></svg></a>
                        <a href="https://www.tiktok.com/@eng_mohamedghonaim?_r=1&_t=ZS-99VdklZPBUS" target="_blank" title="TikTok" class="social-btn-top tiktok-bg"><svg viewBox="0 0 24 24"><path d="M19.589 6.686a4.793 4.793 0 0 1-3.77-4.245V2h-3.445v13.672a2.896 2.896 0 0 1-5.201 1.743l-.068-.102a2.895 2.895 0 0 1 2.373-4.513c.277 0 .546.039.803.111V9.417a6.338 6.338 0 0 0-.803-.051C6.017 9.366 3.2 12.183 3.2 15.647 3.2 19.11 6.017 22 9.479 22c3.462 0 6.279-2.817 6.279-6.353V9.07c1.378.983 3.054 1.564 4.869 1.584V7.209a4.845 4.845 0 0 1-1.038-.523z"/></svg></a>
                        <a href="https://youtube.com/@engineermaths?si=8C6T808VuAU5OMOt" target="_blank" title="YouTube" class="social-btn-top youtube-bg"><svg viewBox="0 0 24 24"><path d="M23.498 6.186a3.016 3.016 0 0 0-2.122-2.136C19.505 3.545 12 3.545 12 3.545s-7.505 0-9.377.505A3.017 3.017 0 0 0 .502 6.186C0 8.07 0 12 0 12s0 3.93.502 5.814a3.016 3.016 0 0 0 2.122 2.136c1.871.505 9.376.505 9.376.505s7.505 0 9.377-.505a3.015 3.015 0 0 0 2.122-2.136C24 15.93 24 12 24 12s0-3.93-.502-5.814zM9.545 15.568V8.432L15.818 12l-6.273 3.568z"/></svg></a>
                    </div>
                """, unsafe_allow_html=True)
                
                st.write("")
                c_home_b1, c_home_b2 = st.columns(2)
                with c_home_b1:
                    if st.button("🔐 تسجيل الدخول الآن"):
                        st.session_state.page_view = "login"
                        st.rerun()
                with c_home_b2:
                    if st.button("✨ إنشاء حساب جديد"):
                        st.session_state.page_view = "register"
                        st.rerun()

            with col_hero_img:
                if img_b64:
                    st.markdown(f"""
                        <div style="display: flex; justify-content: center; align-items: center; position: relative; margin-top: 10px;">
                            <div style="position: absolute; width: 240px; height: 240px; background: #059669; border-radius: 50%; z-index: 0; filter: blur(15px); opacity: 0.7;"></div>
                            <img src="data:image/jpeg;base64,{img_b64}" style="width: 230px; height: 230px; border-radius: 50%; border: 5px solid #059669; object-fit: cover; z-index: 1; box-shadow: 0 10px 25px rgba(0,0,0,0.15);">
                        </div>
                    """, unsafe_allow_html=True)

            st.write("---")

            # كورسات درسلي
            st.markdown('<div class="darssly-box">', unsafe_allow_html=True)
            st.markdown("<h3 style='color: #ffffff; text-align: center; margin-bottom: 5px; font-size: 22px;'>📢 اشترك الآن في كورسات الرياضيات والإحصاء على منصة درسلي (Darssly)</h3>", unsafe_allow_html=True)
            st.markdown("<p style='color: #ffffff; text-align: center; margin-bottom: 25px; font-size: 16px;'>اختر مرحلتك للاطلاع على الشرح والخطط الكاملة:</p>", unsafe_allow_html=True)

            courses_grid = [
                ("📊", "إحصاء الثالث الثانوي", "شرح مبسط وتدريبات متقدمة لامتحان العزم", "https://darssly.com/courses/mohamed-ghoneim-statistics/plans"),
                ("📖", "رياضيات أول إعدادي", "شرح كامل وتدريبات دورية مبسطة", "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim-3/plans"),
                ("📘", "رياضيات ثاني إعدادي", "متابعة شاملة وأسئلة تفاعلية مميزة", "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim/plans"),
                ("📐", "رياضيات ثالث إعدادي", "تأسيس قوي وضمان الدرجة النهائية", "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim-2/plans")
            ]

            c_cols = st.columns(2)
            for idx, (icon, title, desc, link) in enumerate(courses_grid):
                col_target = c_cols[idx % 2]
                with col_target:
                    st.markdown(f"""
                        <div class="course-card">
                            <div>
                                <div class="course-icon-box">{icon}</div>
                                <div class="course-title">{title}</div>
                                <div class="course-desc">{desc}</div>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    st.link_button(f"معرفة تفاصيل الاشتراك لـ {title} 👈", link, use_container_width=True)
                    st.write("")

            st.markdown('</div>', unsafe_allow_html=True)

            # --- قسم حجز الدروس أونلاين ---
            st.markdown("<div class='vertical-section-header'>📅 حجز دروس أونلاين مباشرة مع م / محمد غنيم</div>", unsafe_allow_html=True)
            with st.form("online_booking_form", clear_on_submit=True):
                book_name = st.text_input("اسم الطالب بالكامل:")
                book_curr = st.selectbox("اختر المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()), key="book_c")
                book_grade = st.selectbox("المرحلة / الصف الدراسي:", CURRICULUM_DATA[book_curr], key="book_g")
                book_phone = st.text_input("رقم هاتف الطالب:")
                book_parent_phone = st.text_input("رقم تليفون ولي الأمر:")
                
                if st.form_submit_button("🚀 إرسال طلب حجز الدرس أونلاين"):
                    if not book_name.strip() or not book_phone.strip():
                        st.error("يرجى كتابة اسم الطالب ورقم الهاتف على الأقل.")
                    else:
                        new_booking = {
                            "تاريخ_الحجز": str(date.today()),
                            "اسم الطالب": book_name.strip(),
                            "المنهج_الدولة": book_curr,
                            "المرحلة_الصف": book_grade,
                            "رقم_الهاتف": book_phone.strip(),
                            "رقم_ولي_الأمر": book_parent_phone.strip(),
                            "الحالة": "قيد المتابعة"
                        }
                        st.session_state.bookings_df = pd.concat([st.session_state.bookings_df, pd.DataFrame([new_booking])], ignore_index=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                        st.success("✓ تم إرسال طلب الحجز بنجاح! سيتم التواصل معك قريباً لتأكيد الموعد.")

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
                reg_name = st.text_input("اسمك بالكامل (ثلاثي أو رباعي):", placeholder="مثال: أحمد محمود علي")
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
                                "تاريخ التسجيل": str(date.today()), "الحالة_حظر": "نشط", "حالة_الاشتراك_البنك": "غير مشترك"
                            }
                            st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([new_user])], ignore_index=True)
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                            st.session_state.logged_student = new_user
                            st.query_params["role"] = "student"
                            st.query_params["st_name"] = new_user["اسم الطالب"]
                            st.success(f"تم إنشاء حسابك بنجاح يا {reg_name}!")
                            st.rerun()

        st.markdown("<div class='vertical-section-header'>🎥 حصص سريعة وملخصات هامة في أقل من دقيقة</div>", unsafe_allow_html=True)
        col_vid1, col_vid2, col_vid3 = st.columns([1, 2, 1])
        with col_vid2:
            st.video("https://www.youtube.com/watch?v=6PleAxZCNZM")

    else:
        st_user = st.session_state.logged_student
        fresh_user = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip()]
        if not fresh_user.empty and fresh_user.iloc[0].get("الحالة_حظر") == "محظور":
            st.error("🚫 عذراً، تم حظر حسابك.")
            st.session_state.logged_student = None
            st.query_params.clear()
            st.query_params["role"] = "student"
            st.rerun()

        col_u1, col_u2 = st.columns([4, 1])
        with col_u1:
            st.markdown(f"""
                <div style="background: {card_bg}; padding: 14px 20px; border-radius: 12px; border-right: 5px solid #10b981; margin-bottom: 20px; border: 1px solid {card_border};">
                    <h3 style="margin: 0; color: #059669; font-size: 20px;">أهلاً بك: {st_user['اسم الطالب']} 🌟</h3>
                    <p style="margin: 4px 0 10px 0; font-weight: 900; color: {text_color}; font-size: 16px;">{st_user.get('المنهج/الدولة', '')} | {st_user.get('المجموعة/الصف', '')}</p>
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

        st.markdown("<div class='vertical-section-header'>🎥 حصص سريعة وملخصات هامة في أقل من دقيقة</div>", unsafe_allow_html=True)
        col_vid1, col_vid2, col_vid3 = st.columns([1, 2, 1])
        with col_vid2:
            st.video("https://www.youtube.com/watch?v=6PleAxZCNZM")

        # ==================== لوحة خدمات الطالب ====================
        st.markdown("<div class='vertical-section-header'>🗂️ لوحة خدمات الطالب التفاعلية</div>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; margin-bottom: 15px; color: {text_color}; font-size: 16px;'>اختر القسم الذي تريد فتحه:</p>", unsafe_allow_html=True)

        if st.button("🎥 الفيديوهات والشروحات التعليمية", use_container_width=True):
            st.session_state.student_sub_page = "videos"
            st.rerun()
        if st.button("📚 بنك الأسئلة الشامل (مغلق للاشتراك)", use_container_width=True):
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
        if st.button("📊 متابعة درجات الواجبات المنزلية", use_container_width=True):
            st.session_state.student_sub_page = "hw_grades"
            st.rerun()
        if st.button("📈 متابعة درجات الاختبارات والكويزات", use_container_width=True):
            st.session_state.student_sub_page = "exam_grades"
            st.rerun()
        if st.button("💬 مركز الدردشة والدعم المباشر", use_container_width=True):
            st.session_state.student_sub_page = "chat"
            st.rerun()

        sub_page = st.session_state.student_sub_page
        st.write("---")

        # --- قسم اختبارات موقع عبقري للطالب ---
        if sub_page == "abqary":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🧠 اختبارات ونتائج موقع عبقري (مرحلتك الدراسية)</h3>", unsafe_allow_html=True)
            student_grade = str(st_user.get("المجموعة/الصف", "")).strip().lower()
            abq_df = st.session_state.abqary_df
            
            st_abq = abq_df[abq_df["المجموعة/الصف"].astype(str).str.strip().str.lower().str.contains(student_grade, na=False)]
            if st_abq.empty:
                st_abq = abq_df.copy()

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
                            <p style="font-size:16px;">اضغط على الزر أدناه للانتقال إلى موقع عبقري وإجراء الاختبار:</p>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    if ab_link and ab_link != "nan":
                        st.link_button(f"🔗 الذهاب لامتحان عبقري: {ab_title} 🚀", ab_link, use_container_width=True)

                    st.write("")
                    st.markdown("##### 🔍 استعلام عن نتيجة هذا الاختبار برقم سري:")
                    with st.form(f"abqary_result_form_{ab_i}"):
                        entered_pass = st.text_input("أدخل الرقم السري المخصص لإظهار النتيجة:", type="password", key=f"pass_input_{ab_i}")
                        if st.form_submit_button("🔓 إظهار النتيجة"):
                            if secret_code and entered_pass.strip() == secret_code:
                                st.success("✓ الرقم السري صحيح! تم جلب النتيجة بنجاح:")
                                if res_link and res_link != "nan":
                                    st.link_button("📊 اضغط هنا لعرض نتيجة امتحان عبقري 🏆", res_link, use_container_width=True)
                                else:
                                    st.info("النتيجة متاحة وسيتم إضافتها قريباً.")
                            else:
                                st.error("❌ الرقم السري غير صحيح. يرجى مراجعة معلم المادة.")
                    st.write("---")

        # --- صفحة الفيديوهات بتصميم "درسلي" الاحترافي مع الأيقونات الاحترافية لاختيار الدروس ---
        elif sub_page == "videos":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🎥 محتوى الشروحات والفيديوهات التعليمية</h3>", unsafe_allow_html=True)
            student_grade = str(st_user.get("المجموعة/الصف", "")).strip()
            v_df = st.session_state.videos_df
            st_videos = v_df[v_df["المجموعة/الصف"].astype(str).str.strip().str.lower() == student_grade.lower()]

            if st_videos.empty:
                st.info(f"لا توجد فيديوهات مرفوعة لمرحلتك الدراسية ({student_grade}) حالياً. ترقبها قريباً!")
            else:
                col_sidebar, col_main_vid = st.columns([1, 2.2])
                
                with col_sidebar:
                    st.markdown(f"""
                        <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:12px; padding:15px; margin-bottom:15px;">
                            <h4 style="margin:0 0 10px 0; color:#059669; font-size:18px;">📚 محتوى الدروس</h4>
                        </div>
                    """, unsafe_allow_html=True)

                    if "selected_video_idx" not in st.session_state:
                        st.session_state.selected_video_idx = 0

                    for v_i, v_row in st_videos.reset_index(drop=True).iterrows():
                        v_title_btn = f"📖 {v_row['عنوان_الفيديو']}"
                        is_active_btn = (st.session_state.selected_video_idx == v_i)
                        
                        if st.button(v_title_btn, key=f"darssly_btn_v_{v_i}", use_container_width=True):
                            st.session_state.selected_video_idx = v_i
                            st.rerun()

                    selected_row = st_videos.iloc[st.session_state.selected_video_idx] if st.session_state.selected_video_idx < len(st_videos) else st_videos.iloc[0]

                with col_main_vid:
                    st.markdown(f"""
                        <div style="background:linear-gradient(135deg, #059669, #10b981); color:#ffffff; padding:15px 20px; border-radius:10px; margin-bottom:15px;">
                            <h3 style="margin:0; color:#ffffff; font-size:20px;">📺 {selected_row['عنوان_الفيديو']}</h3>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    v_link = str(selected_row.get("رابط_الفيديو", "")).strip()
                    v_bytes = selected_row.get("فيديو_base64", "")

                    if v_link and v_link != "nan" and v_link != "":
                        if "youtube.com" in v_link or "youtu.be" in v_link:
                            st.video(v_link)
                        else:
                            st.video(v_link)
                    elif pd.notnull(v_bytes) and str(v_bytes).strip() and str(v_bytes) != "nan":
                        try:
                            vid_bytes_dec = base64.b64decode(v_bytes)
                            st.video(vid_bytes_dec)
                        except Exception:
                            st.error("⚠️ يفضل استخدام روابط يوتيوب أو روابط مباشرة للفيديوهات لضمان سرعة التشغيل.")
                    else:
                        st.info("لا يوجد فيديو متاح لهذا الدرس.")

                    st.write("---")
                    
                    current_vid_title = selected_row['عنوان_الفيديو']
                    vc_df = st.session_state.video_comments_df
                    vid_comments = vc_df[vc_df["عنوان_الفيديو"].astype(str).str.strip() == current_vid_title.strip()]
                    
                    st.markdown(f"<h4 style='font-size:18px;'>❓ الأسئلة والتعليقات ({len(vid_comments)})</h4>", unsafe_allow_html=True)
                    
                    if not vid_comments.empty:
                        for _, c_row in vid_comments.iterrows():
                            st.markdown(f"""
                                <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:10px; padding:12px 15px; margin-bottom:10px;">
                                    <p style="margin:0; font-size:14px; color:#0284c7;"><b>{c_row['اسم الطالب']}</b> — <span style="font-size:12px; opacity:0.7;">{c_row['التاريخ_والوقت']}</span></p>
                                    <p style="margin:5px 0 0 0; font-size:16px;">{c_row['نص_التعليق']}</p>
                                </div>
                            """, unsafe_allow_html=True)
                    
                    with st.form(f"comment_form_{selected_row['معرف_الفيديو']}"):
                        user_comment_text = st.text_area("أكتب سؤالك أو استفسارك حول هذا الدرس:", placeholder="اكتب سؤالك هنا ليجيب عليه البشمهندس...")
                        if st.form_submit_button("إرسال التعليق أو السؤال"):
                            if user_comment_text.strip():
                                new_comment = {
                                    "التاريخ_والوقت": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                    "عنوان_الفيديو": current_vid_title.strip(),
                                    "اسم الطالب": st_user["اسم الطالب"],
                                    "نص_التعليق": user_comment_text.strip()
                                }
                                st.session_state.video_comments_df = pd.concat([st.session_state.video_comments_df, pd.DataFrame([new_comment])], ignore_index=True)
                                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                                st.success("✓ تم إرسال تعليقك أو سؤالك بنجاح وسيظهر للمعلم فوراً!")
                                st.rerun()

        # --- بنك الأسئلة للطالب ---
        elif sub_page == "bank":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📚 بنك الأسئلة الشامل (مرحلتك الدراسية)</h3>", unsafe_allow_html=True)
            
            curr_user_row = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip()]
            sub_status = curr_user_row.iloc[0].get("حالة_الاشتراك_البنك", "غير مشترك") if not curr_user_row.empty else "غير مشترك"

            if sub_status != "مشترك":
                st.warning("🔒 عذراً، بنك الأسئلة مغلق ويحتاج إلى اشتراك خاص بمرحلتك الدراسية بقيمة **100 جنيه** فقط.")
                st.markdown(f"""
                    <div style="background:{card_bg}; border:2px solid #dc2626; border-radius:15px; padding:25px; margin-bottom:20px;">
                        <h3 style="color:#dc2626; margin-top:0;">💳 تعليمات الاشتراك في بنك الأسئلة:</h3>
                        <p style="font-size:18px;">1. قم بالدفع بقيمة <b>100 جنيه</b> عبر تطبيق InstaPay باستخدام زر الدفع السريع بالأسفل.</p>
                        <p style="font-size:18px;">2. اكتب رقم هاتفك المحول منه، وكود التحقق (OTP)، وارفق صورة اسكرين (إيصال) الدفع لتأكيد طلبك.</p>
                    </div>
                """, unsafe_allow_html=True)

                with st.form("bank_subscription_form"):
                    pay_phone = st.text_input("رقم الهاتف المحول منه (فودافون كاش، اتصالات كاش، أو وي كاش):", placeholder="010XXXXXXXX")
                    pay_otp = st.text_input("كود التحقق الخاص بـ InstaPay أو معاملة التحويل (OTP):", placeholder="مثال: 4589")
                    pay_receipt = st.file_uploader("📷 رفع اسكرين (إيصال) الدفع:", type=["jpg", "png", "jpeg"])
                    
                    if st.form_submit_button("📤 إرسال طلب الاشتراك لتأكيد المعلم"):
                        if not pay_phone.strip() or not pay_otp.strip() or pay_receipt is None:
                            st.error("يرجى إدخال رقم الهاتف، وكود الـ OTP، وإرفاق صورة إيصال الدفع.")
                        else:
                            rcpt_str = base64.b64encode(pay_receipt.read()).decode()
                            new_req = {
                                "تاريخ_الطلب": str(date.today()),
                                "اسم الطالب": st_user["اسم الطالب"],
                                "رقم_الهاتف": pay_phone.strip(),
                                "كود_OTP": pay_otp.strip(),
                                "حالة_الدفع": "قيد المراجعة",
                                "إيصال_الدفع_base64": rcpt_str
                            }
                            st.session_state.bank_requests_df = pd.concat([st.session_state.bank_requests_df, pd.DataFrame([new_req])], ignore_index=True)
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                            st.success("✓ تم إرسال طلب اشتراكك بنجاح! سيقوم المعلم بمراجعة الإيصال وتفعيل حسابك خلال دقائق.")
                
                st.link_button("📲 اضغط هنا للدفع من خلال انستا باي", "https://ipn.eg/S/moghonem2002/instapay/6EyvZs")
            else:
                st.success("🎉 أهلاً بك! حسابك مفعل ومسجل في بنك الأسئلة الخاص بمرحلتك الدراسية.")
                student_grade = str(st_user.get("المجموعة/الصف", "")).strip().lower()
                qb_df = st.session_state.question_bank_df
                
                st_qb = qb_df[qb_df["المجموعة/الصف"].astype(str).str.strip().str.lower().str.contains(student_grade, na=False)]
                if st_qb.empty:
                    st_qb = qb_df.copy()

                if st_qb.empty:
                    st.info("لا توجد أسئلة مضافة في بنك الأسئلة لمرحلتك حالياً.")
                else:
                    for qb_i, qb_r in st_qb.iterrows():
                        raw_json_data = qb_r.get("بيانات_السؤال_JSON", "{}")
                        try:
                            q_data = json.loads(raw_json_data) if pd.notnull(raw_json_data) and str(raw_json_data).strip() else {}
                        except Exception:
                            q_data = {}

                        if not q_data:
                            continue

                        st.markdown(f"""
                            <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:12px; padding:20px; margin-bottom:15px;">
                                <p style="font-size:18px; color:{text_color};"><b>سؤال ({qb_i+1}) — الدرجة: {q_data.get('points', 1.0)}</b></p>
                                <p style="font-size:17px; color:{text_color};">{q_data.get('text', '')}</p>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        if q_data.get("q_img"):
                            st.image(f"data:image/jpeg;base64,{q_data['q_img']}", use_container_width=True)

                        q_type_val = q_data.get("type", "اختيار من متعدد")
                        if "اختيار" in q_type_val or q_type_val == "موضوعي":
                            opts = ["أ", "ب", "ج", "د"]
                            ans_choice = st.radio(f"اختر الإجابة الصحيحة للسؤال ({qb_i+1}):", opts, key=f"qb_radio_{qb_i}")
                            
                            opt1 = q_data.get('opt1', '')
                            opt2 = q_data.get('opt2', '')
                            opt3 = q_data.get('opt3', '')
                            opt4 = q_data.get('opt4', '')
                            if opt1 or opt2 or opt3 or opt4:
                                st.markdown(f"""
                                    <div style="padding: 10px; background: #f1f5f9; border-radius: 8px; margin-bottom: 10px;">
                                        أ) {opt1} &nbsp;&nbsp;|&nbsp;&nbsp; ب) {opt2} &nbsp;&nbsp;|&nbsp;&nbsp; ج) {opt3} &nbsp;&nbsp;|&nbsp;&nbsp; د) {opt4}
                                    </div>
                                """, unsafe_allow_html=True)

                            if st.button(f"تحقق من إجابة السؤال ({qb_i+1})", key=f"check_qb_{qb_i}"):
                                correct_idx = int(q_data.get("correct", 1))
                                correct_letter = opts[correct_idx - 1] if 0 <= correct_idx - 1 < 4 else 'أ'
                                if ans_choice == correct_letter:
                                    st.success("إجابة صحيحة تماماً! أحسنت ✅")
                                else:
                                    st.error(f"إجابة خاطئة ❌. الإجابة الصحيحة هي: الخيار ({correct_letter})")
                        else:
                            essay_ans_txt = st.text_area(f"اكتب خطوات حل السؤال المقالي ({qb_i+1}):", key=f"qb_essay_txt_{qb_i}")
                            essay_ans_img = st.file_uploader(f"أو ارفع صورة حل السؤال المقالي ({qb_i+1}):", type=["jpg", "png", "jpeg"], key=f"qb_essay_img_{qb_i}")
                            if st.button(f"إرسال حل السؤال المقالي للمعلم ({qb_i+1})", key=f"submit_qb_essay_{qb_i}"):
                                img_b64_sub = base64.b64encode(essay_ans_img.read()).decode() if essay_ans_img is not None else ""
                                new_qb_sub = {
                                    "معرف_الحل": f"QB_ANS_{datetime.now().strftime('%Y%m%d%H%M%S')}_{qb_i}",
                                    "معرف_الامتحان": "بنك الأسئلة",
                                    "عنوان الامتحان": "بنك الأسئلة الشامل",
                                    "اسم الطالب": st_user["اسم الطالب"],
                                    "رقم السؤال": qb_i + 1,
                                    "نص السؤال": q_data.get("text", ""),
                                    "إجابة الطالب النصية": essay_ans_txt,
                                    "صورة الحل_base64": img_b64_sub,
                                    "درجة السؤال": q_data.get("points", 1.0),
                                    "الدرجة المرصودة": 0.0,
                                    "حالة التصحيح": "قيد التصحيح من المعلم",
                                    "ملاحظات المعلم": "",
                                    "تاريخ الحل": str(date.today()),
                                }
                                st.session_state.essays_df = pd.concat([st.session_state.essays_df, pd.DataFrame([new_qb_sub])], ignore_index=True)
                                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                                st.success("✓ تم إرسال إجابتك المقالية للمعلم بنجاح ليتم تصحيحها!")

        # 2. صفحة الاختبارات
        elif sub_page == "exams":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>✍️ الاختبارات الإلكترونية التفاعلية المتاحة:</h3>", unsafe_allow_html=True)
            available_exams = st.session_state.exams_df.copy()

            if available_exams.empty:
                st.info("لا توجد اختبارات إلكترونية منشورة حالياً.")
            else:
                for _, ex_row in available_exams.iterrows():
                    ex_id = ex_row["معرف_الامتحان"]
                    ex_title = ex_row["عنوان الامتحان"]
                    ex_desc = ex_row["وصف الامتحان"]
                    ex_pass = str(ex_row.get("كلمة المرور", "")).strip()
                    ex_time = int(ex_row.get("مدة الامتحان بالدقائق", 30))

                    with st.container():
                        already_solved = st.session_state.assessments_df[
                            (st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip())
                            & (st.session_state.assessments_df["عنوان التكليف"].astype(str).str.contains(ex_title, na=False))
                        ]

                        if not already_solved.empty:
                            solved_score = already_solved.iloc[-1]["الدرجة المحصلة"]
                            solved_max = already_solved.iloc[-1]["الدرجة العظمى"]
                            pct_old = (solved_score / solved_max * 100) if solved_max > 0 else 0
                            st.markdown(f"""
                                <div style="background:#dc2626; border:3px solid #b91c1c; border-radius:20px; padding:35px; text-align:center; box-shadow:0 15px 30px rgba(0,0,0,0.2); margin-bottom:20px;">
                                    <h2 style="color:#ffffff; margin-bottom:8px; font-size:26px;">🎓 شهادة إتمام الاختبار والنتيجة النهائية</h2>
                                    <p style="font-size:20px; color:#ffffff; margin: 5px 0;"><b>الطالب: {st_user['اسم الطالب']}</b></p>
                                    <p style="font-size:18px; color:#f8fafc; margin: 5px 0;">اجتاز بنجاح امتحان: <b>{ex_title}</b></p>
                                    <div style="width:130px; height:130px; border-radius:50%; border:10px solid #ffffff; display:flex; align-items:center; justify-content:center; margin:25px auto; font-size:30px; font-weight:900; color:#ffffff;">
                                        {pct_old:.0f}%
                                    </div>
                                    <p style="font-size:19px; font-weight:900; color:#ffffff;">الدرجة المحصلة: <b>{solved_score} / {solved_max}</b></p>
                                    <p style="margin-top:15px; font-size:16px; color:#f1f5f9;">مع تحيات معلم المادة: <b>م / محمد غنيم</b></p>
                                </div>
                            """, unsafe_allow_html=True)
                            st.markdown(f"<div style='background:#fef2f2; padding:18px; border-radius:12px; border:1px solid #dc2626; margin-top:10px; color:#991b1b; font-size:16px;'><b>تفاصيل إجاباتك السابقة:</b><br>{already_solved.iloc[-1]['ملاحظات وتوجيهات']}</div>", unsafe_allow_html=True)
                        else:
                            exam_state_key = f"exam_state_{ex_id}"
                            if exam_state_key not in st.session_state:
                                st.session_state[exam_state_key] = {
                                    "started": False,
                                    "cur_idx": 0,
                                    "answers_mcq": {},
                                    "essay_texts": {},
                                    "essay_imgs": {},
                                    "start_time_str": None
                                }

                            st_ex = st.session_state[exam_state_key]

                            if not st_ex["started"]:
                                st.markdown(f"""
                                    <div style="background-color: #f97316; color: #ffffff; padding: 15px 25px; border-radius: 10px 10px 0 0; font-size: 20px; font-weight: 900;">
                                        {ex_title}
                                    </div>
                                    <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:0 0 10px 10px; padding:25px; margin-bottom:25px; color:{text_color};">
                                        <p style="color:{text_color}; font-size:17px;"><b>الوصف:</b> {ex_desc}</p>
                                        <p style="color: #b91c1c; font-weight: 900; font-size:16px;">⚠️ مدة الامتحان ({ex_time} دقيقة)، عند انتهاء الوقت ينتهي الامتحان تلقائياً.</p>
                                        <p style="color: #b91c1c; font-weight: 900; font-size:16px;">⚠️ تم تحديد محاولة واحدة فقط لهذا الإمتحان.</p>
                                """, unsafe_allow_html=True)

                                input_pass = ""
                                if ex_pass and ex_pass != "nan" and ex_pass != "":
                                    input_pass = st.text_input("🔑 أُدخل رقم السر (كلمة المرور):", type="password", key=f"start_pass_{ex_id}")
                                
                                if st.button("🚀 دخول الامتحان", key=f"btn_start_ex_{ex_id}"):
                                    if ex_pass and ex_pass != "nan" and ex_pass != "" and input_pass.strip() != ex_pass:
                                        st.error("كلمة مرور الامتحان غير صحيحة.")
                                    else:
                                        st_ex["started"] = True
                                        st_ex["start_time_str"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        st.rerun()
                                st.markdown("</div>", unsafe_allow_html=True)
                            else:
                                questions = []
                                raw_q = ex_row.get("الأسئلة_JSON", "[]")
                                try:
                                    if isinstance(raw_q, str):
                                        questions = json.loads(raw_q)
                                    elif isinstance(raw_q, list):
                                        questions = raw_q
                                except Exception:
                                    try:
                                        questions = eval(str(raw_q))
                                    except Exception:
                                        questions = []

                                if not questions:
                                    st.error("⚠️ لا توجد أسئلة مضافة في هذا الاختبار.")
                                else:
                                    total_q = len(questions)
                                    cur_i = st_ex["cur_idx"]
                                    q_curr = questions[cur_i]

                                    start_dt = datetime.strptime(st_ex["start_time_str"], "%Y-%m-%d %H:%M:%S") if st_ex.get("start_time_str") else datetime.now()
                                    elapsed_secs = (datetime.now() - start_dt).seconds
                                    total_allowed_secs = ex_time * 60
                                    remaining_secs = max(0, total_allowed_secs - elapsed_secs)
                                    rem_mins = remaining_secs // 60
                                    rem_s = remaining_secs % 60

                                    if remaining_secs == 0:
                                        st.warning("⏰ تنبيه: انتهى وقت الامتحان المحدد! سيتم تسليم إجاباتك تلقائياً.")
                                        st_ex["show_confirm_submit"] = True

                                    st.markdown(f"""
                                        <div style="background-color: #f97316; color: #ffffff; padding: 12px 20px; border-radius: 10px 10px 0 0; display: flex; justify-content: space-between; align-items: center; font-weight: 900; font-size:18px;">
                                            <span>{ex_title}</span>
                                            <span>⏱️ الوقت المتبقي: {rem_mins:02d}:{rem_s:02d}</span>
                                        </div>
                                        <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:0 0 10px 10px; padding:20px; margin-bottom:20px; color:{text_color};">
                                            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; font-weight:900; color:{text_color}; font-size:17px;">
                                                <span style="color:{text_color};">الطالب: {st_user['اسم الطالب']}</span>
                                                <span style="color:{text_color};">السؤال ({cur_i + 1} من {total_q}) — الدرجة: {q_curr['points']}</span>
                                            </div>
                                    """, unsafe_allow_html=True)

                                    if q_curr.get("text"):
                                        st.markdown(f"<span style='color:{text_color}; font-size:18px;'><b>{q_curr['text']}</b></span>", unsafe_allow_html=True)
                                    
                                    if q_curr.get("q_img"):
                                        st.image(f"data:image/jpeg;base64,{q_curr['q_img']}", use_container_width=True)

                                    q_type = q_curr.get("type", "موضوعي")
                                    if q_type == "موضوعي":
                                        opts_labels = ["أ", "ب", "ج", "د"]
                                        saved_choice = st_ex["answers_mcq"].get(cur_i, None)

                                        st.markdown(f"<p style='color:{text_color}; font-size:16px; margin-top:10px;'><b>اختر إجابتك (انقر على الخيار المطلوب):</b></p>", unsafe_allow_html=True)
                                        
                                        for opt_idx, lbl in enumerate(opts_labels, 1):
                                            t_val = q_curr.get(f"opt{opt_idx}", "")
                                            i_val = q_curr.get(f"opt{opt_idx}_img", "")
                                            
                                            is_selected = (saved_choice == opt_idx)
                                            btn_label = f"({lbl}) {t_val}" if t_val else f"الخيار ({lbl})"
                                            if is_selected:
                                                btn_label = f"✅ تم الاختيار: {btn_label}"

                                            if st.button(btn_label, key=f"opt_rect_btn_{ex_id}_{cur_i}_{opt_idx}", use_container_width=True):
                                                st_ex["answers_mcq"][cur_i] = opt_idx
                                                st.rerun()

                                            if i_val:
                                                st.image(f"data:image/jpeg;base64,{i_val}", width=280)

                                    else:
                                        st.markdown(f"<p style='color:{text_color}; font-size:16px;'><b>✍️ سؤال مقالي:</b> اكتب خطوات الحل أو ارفع صورة الحل من كشكولك ليتم تصحيحها من قبل المعلم.</p>", unsafe_allow_html=True)
                                        prev_essay_txt = st_ex["essay_texts"].get(cur_i, "")
                                        in_essay_txt = st.text_area("خطوات الحل المكتوبة:", value=prev_essay_txt, key=f"essay_input_txt_{ex_id}_{cur_i}")
                                        st_ex["essay_texts"][cur_i] = in_essay_txt

                                        up_essay_img = st.file_uploader("📷 إرفاق صورة الحل (من كشكولك):", type=["jpg", "png", "jpeg"], key=f"essay_input_img_{ex_id}_{cur_i}")
                                        if up_essay_img is not None:
                                            st_ex["essay_imgs"][cur_i] = base64.b64encode(up_essay_img.read()).decode()

                                        if st_ex["essay_imgs"].get(cur_i):
                                            st.image(f"data:image/jpeg;base64,{st_ex['essay_imgs'][cur_i]}", use_container_width=True)

                                    st.write("---")
                                    c_prev, c_next, c_finish = st.columns([1, 1, 2])

                                    with c_prev:
                                        if cur_i > 0:
                                            if st.button("➡️ السابق", key=f"btn_prev_{ex_id}"):
                                                st_ex["cur_idx"] -= 1
                                                st.rerun()

                                    with c_next:
                                        if cur_i < total_q - 1:
                                            if st.button("حفظ ومتابعة ⬅️", key=f"btn_next_{ex_id}"):
                                                st_ex["cur_idx"] += 1
                                                st.rerun()

                                    with c_finish:
                                        if st.button("🏁 إنهاء الامتحان", key=f"btn_finish_{ex_id}"):
                                            st_ex["show_confirm_submit"] = True

                                    if st_ex.get("show_confirm_submit"):
                                        st.write("")
                                        unanswered_list = [q_idx+1 for q_idx in range(total_q) if q_idx not in st_ex["answers_mcq"]]
                                        if unanswered_list:
                                            st.warning(f"⚠️ تنبيه: لقد نسيت الإجابة على الأسئلة الموضوعية الآتية: {unanswered_list}")
                                        
                                        if st.button("✅ انهاء الامتحان ومشاهدة النتيجة", key=f"confirm_final_{ex_id}"):
                                            mcq_score = 0.0
                                            total_max = 0.0
                                            detailed_report = ""

                                            for q_idx, q in enumerate(questions):
                                                q_pts = float(q.get("points", 1.0))
                                                total_max += q_pts

                                                if q.get("type", "موضوعي") == "موضوعي":
                                                    student_ans = st_ex["answers_mcq"].get(q_idx, 0)
                                                    correct_ans = int(q.get("correct", 1))
                                                    is_correct = (student_ans == correct_ans)
                                                    if is_correct:
                                                        mcq_score += q_pts
                                                        detailed_report += f"<br>سؤال {q_idx+1}: إجابة صحيحة ✅"
                                                    else:
                                                        detailed_report += f"<br>سؤال {q_idx+1}: إجابة خاطئة ❌ (الصحيحة: الخيار {['أ', 'ب', 'ج', 'د'][correct_ans-1]})"
                                                else:
                                                    new_essay_sub = {
                                                        "معرف_الحل": f"ANS_{datetime.now().strftime('%Y%m%d%H%M%S')}_{q_idx}",
                                                        "معرف_الامتحان": ex_id,
                                                        "عنوان الامتحان": ex_title,
                                                        "اسم الطالب": st_user["اسم الطالب"],
                                                        "رقم السؤال": q_idx + 1,
                                                        "نص السؤال": str(q.get("text", "سؤال مقالي مصور")),
                                                        "إجابة الطالب النصية": str(st_ex["essay_texts"].get(q_idx, "")),
                                                        "صورة الحل_base64": str(st_ex["essay_imgs"].get(q_idx, "")),
                                                        "درجة السؤال": q_pts,
                                                        "الدرجة المرصودة": 0.0,
                                                        "حالة التصحيح": "قيد التصحيح من المعلم",
                                                        "ملاحظات المعلم": "",
                                                        "تاريخ الحل": str(date.today()),
                                                    }
                                                    st.session_state.essays_df = pd.concat([st.session_state.essays_df, pd.DataFrame([new_essay_sub])], ignore_index=True)

                                            pct = (mcq_score / total_max * 100) if total_max > 0 else 0
                                            note_msg = f"الدرجة الموضوعية: {mcq_score}/{total_max} (السؤال المقالي قيد المراجعة). التفاصيل: {detailed_report}"

                                            new_ass = {
                                                "التاريخ": str(date.today()),
                                                "اسم الطالب": st_user["اسم الطالب"],
                                                "النوع": "اختبار دوري إلكتروني",
                                                "عنوان التكليف": ex_title,
                                                "الدرجة المحصلة": mcq_score,
                                                "الدرجة العظمى": total_max,
                                                "حالة التسليم": "تم التسليم بنجاح",
                                                "ملاحظات وتوجيهات": note_msg,
                                            }
                                            st.session_state.assessments_df = pd.concat([st.session_state.assessments_df, pd.DataFrame([new_ass])], ignore_index=True)
                                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                                            st.session_state.pop(exam_state_key, None)
                                            
                                            st.markdown(f"""
                                                <div style="background:#dc2626; border:3px solid #b91c1c; border-radius:20px; padding:35px; text-align:center; box-shadow:0 15px 30px rgba(0,0,0,0.2); margin-top:20px;">
                                                    <h2 style="color:#ffffff; margin-bottom:8px; font-size:26px;">🎓 شهادة إتمام الاختبار والنتيجة النهائية</h2>
                                                    <p style="font-size:20px; color:#ffffff; margin: 5px 0;"><b>الطالب: {st_user['اسم الطالب']}</b></p>
                                                    <p style="font-size:18px; color:#f8fafc; margin: 5px 0;">اجتاز بنجاح امتحان: <b>{ex_title}</b></p>
                                                    <div style="width:130px; height:130px; border-radius:50%; border:10px solid #ffffff; display:flex; align-items:center; justify-content:center; margin:25px auto; font-size:30px; font-weight:900; color:#ffffff;">
                                                        {pct:.0f}%
                                                    </div>
                                                    <p style="font-size:19px; font-weight:900; color:#ffffff;">الدرجة المحصلة: <b>{mcq_score} / {total_max}</b></p>
                                                    <p style="margin-top:10px; font-size:15px; color:#fef2f2;">(تم إرسال إجاباتك المقالية لمعلم المادة لتصحيحها وإضافة درجتها)</p>
                                                    <p style="margin-top:15px; font-size:16px; color:#f1f5f9;">مع تحيات معلم المادة: <b>م / محمد غنيم</b></p>
                                                </div>
                                            """, unsafe_allow_html=True)
                                            st.markdown(f"<div style='background:#fef2f2; padding:18px; border-radius:12px; border:1px solid #dc2626; margin-top:15px; color:#991b1b; font-size:16px;'><b>تفاصيل الإجابات الصحيحة والخاطئة:</b><br>{note_msg}</div>", unsafe_allow_html=True)
                                            st.rerun()
                                    st.markdown("</div>", unsafe_allow_html=True)

        # 2. صفحة الحضور
        elif sub_page == "attendance":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📝 تسجيل حضور حصة اليوم وتقييمها:</h3>", unsafe_allow_html=True)
            with st.form("logged_student_att_form", clear_on_submit=True):
                st_date = st.date_input("تاريخ الحصة:", value=date.today())
                selected_rating = st.selectbox("⭐ قيّم الحصة مع البشمهندس (من 5 نجوم):", [
                    "⭐⭐⭐⭐⭐ (5/5) ممتاز جداً وفهم ممتاز", "⭐⭐⭐⭐ (4/5) جيد جداً",
                    "⭐⭐⭐ (3/5) جيد ومفهوم", "⭐⭐ (2/5) متوسط ويحتاج توضيح", "⭐ (1/5) يحتاج إعادة شرح"
                ])
                if st.form_submit_button("✅ تأكيد حضور الحصة"):
                    new_row = {
                        "التاريخ": str(st_date), "اسم الطالب": st_user["اسم الطالب"],
                        "المنهج/الدولة": st_user.get("المنهج/الدولة", ""), "المجموعة/الصف": st_user.get("المجموعة/الصف", ""),
                        "الحالة": "حاضر", "سعر الحصة": 0.0, "عدد الحصص الكلي": 1, "نظام الدفع": "مؤجل",
                        "مستوى الطالب": "قيد التقييم", "ملاحظات": f"تقييم الحصة: {selected_rating}",
                    }
                    st.session_state.sessions_df = pd.concat([st.session_state.sessions_df, pd.DataFrame([new_row])], ignore_index=True)
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                    st.success(f"تم تسجيل حضورك بنجاح للحصة بتاريخ {st_date}!")

        # 3. درجات الواجبات
        elif sub_page == "hw_grades":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📊 متابعة درجات الواجبات المنزلية:</h3>", unsafe_allow_html=True)
            my_assessments = st.session_state.assessments_df[
                (st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip())
                & (st.session_state.assessments_df["النوع"].astype(str).str.contains("واجب", na=False))
            ].copy()
            if my_assessments.empty:
                st.info("لا توجد درجات واجبات مرصودة لك حتى الآن.")
            else:
                st.dataframe(my_assessments[["التاريخ", "النوع", "عنوان التكليف", "الدرجة المحصلة", "الدرجة العظمى", "حالة التسليم", "ملاحظات وتوجيهات"]], use_container_width=True)

        # 4. درجات الاختبارات مع شهادة حمراء ونصوص بيضاء
        elif sub_page == "exam_grades":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📈 متابعة درجات الاختبارات والكويزات وتفاصيل الإجابات:</h3>", unsafe_allow_html=True)
            my_exams = st.session_state.assessments_df[
                (st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip())
                & (st.session_state.assessments_df["النوع"].astype(str).str.contains("اختبار|كويز", na=False))
            ].copy()
            if my_exams.empty:
                st.info("لا توجد درجات اختبارات مرصودة لك حتى الآن.")
            else:
                for _, ex_rec in my_exams.iterrows():
                    score_val = ex_rec['الدرجة المحصلة']
                    max_val = ex_rec['الدرجة العظمى']
                    pct_val = (score_val / max_val * 100) if max_val > 0 else 0
                    with st.expander(f"📝 {ex_rec['عنوان التكليف']} — النتيجة: ({score_val} / {max_val})"):
                        st.markdown(f"""
                            <div style="background:#dc2626; border:3px solid #b91c1c; border-radius:20px; padding:35px; text-align:center; box-shadow:0 15px 30px rgba(0,0,0,0.2); margin-bottom:20px;">
                                <h2 style="color:#ffffff; margin-bottom:8px; font-size:26px;">🎓 شهادة إتمام الاختبار</h2>
                                <p style="font-size:20px; color:#ffffff; margin: 5px 0;"><b>الطالب: {st_user['اسم الطالب']}</b></p>
                                <p style="font-size:18px; color:#f8fafc; margin: 5px 0;"><b>النتيجة:</b> {pct_val:.0f}% ({score_val} من {max_val})</p>
                                <p style="font-size:16px; color:#f1f5f9;"><b>التاريخ:</b> {ex_rec['التاريخ']}</p>
                            </div>
                        """, unsafe_allow_html=True)
                        st.markdown(f"<span style='color:{text_color}; font-size:16px;'><b>تفاصيل وتوجيهات الإجابة:</b><br>{ex_rec['ملاحظات وتوجيهات']}</span>", unsafe_allow_html=True)

        # 5. الدردشة
        elif sub_page == "chat":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>💬 مركز الدردشة والدعم المباشر:</h3>", unsafe_allow_html=True)
            student_name_key = st_user["اسم الطالب"].strip()
            chat_history = st.session_state.messages_df[st.session_state.messages_df["اسم الطالب"].astype(str).str.strip() == student_name_key].copy()

            with st.container():
                if not chat_history.empty:
                    for _, msg in chat_history.iterrows():
                        sender = msg["المرسل"]
                        t_stamp = msg.get("التاريخ_والوقت", "")
                        content = msg.get("نص الرسالة", "")
                        img_data = msg.get("الصورة_base64", "")
                        if sender == "student":
                            st.markdown(f"<div class='chat-bubble-student'><b>أنت ({t_stamp}):</b><br>{content}</div>", unsafe_allow_html=True)
                        else:
                            st.markdown(f"<div class='chat-bubble-teacher'><b>البشمهندس محمد غنيم ({t_stamp}):</b><br>{content}</div>", unsafe_allow_html=True)
                        if pd.notnull(img_data) and str(img_data).strip():
                            st.image(f"data:image/jpeg;base64,{img_data}", width=300)

            with st.form("student_chat_send_form", clear_on_submit=True):
                msg_text = st.text_area("اكتب رسالتك أو استفسارك هنا:", placeholder="مستر، مش فاهم المسألة رقم...")
                uploaded_msg_img = st.file_uploader("📷 إرفاق صورة للمسألة (اختياري)", type=["jpg", "png", "jpeg"])
                if st.form_submit_button("📤 إرسال الرسالة إلى البشمهندس"):
                    if msg_text.strip() or uploaded_msg_img is not None:
                        img_str = base64.b64encode(uploaded_msg_img.read()).decode() if uploaded_msg_img is not None else ""
                        new_msg = {
                            "التاريخ_والوقت": datetime.now().strftime("%Y-%m-%d %H:%M"), "اسم الطالب": student_name_key,
                            "المرسل": "student", "نص الرسالة": msg_text.strip(), "الصورة_base64": img_str,
                        }
                        st.session_state.messages_df = pd.concat([st.session_state.messages_df, pd.DataFrame([new_msg])], ignore_index=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                        st.success("تم إرسال رسالتك للبشمهندس بنجاح!")
                        st.rerun()

    st.markdown("""
        <div class="call-btn-container">
            <a href="tel:01016361440" class="call-btn">
                <svg viewBox="0 0 24 24" style="width:24px;height:24px;fill:#ffffff;"><path d="M6.62 10.79c1.44 2.83 3.76 5.14 6.59 6.59l2.2-2.2c.27-.27.67-.36 1.02-.24 1.12.37 2.33.57 3.57.57.55 0 1 .45 1 1V20c0 .55-.45 1-1 1-9.39 0-17-7.61-17-17 0-.55.45-1 1-1h3.5c.55 0 1 .45 1 1 0 1.25.2 2.45.57 3.57.11.35.03.74-.25 1.02l-2.2 2.2z"/></svg>
                <span>للتواصل مع م / محمد غنيم: 01016361440</span>
            </a>
        </div>
        <div class="social-footer-box">
            <div class="social-footer-container">
                <a href="https://www.facebook.com/share/19fD41rV3H/" target="_blank" title="Facebook" class="social-btn-top facebook-bg"><svg viewBox="0 0 24 24"><path d="M24 12.073c0-6.627-5.373-12-12-12s-12 5.373-12 12c0 5.99 4.388 10.954 10.125 11.854v-8.385H7.078v-3.47h3.047V9.43c0-3.007 1.792-4.669 4.533-4.669 1.312 0 2.686.235 2.686.235v2.953H15.83c-1.491 0-1.956.925-1.956 1.874v2.25h3.328l-.532 3.47h-2.796v8.385C19.612 23.027 24 18.062 24 12.073z"/></svg></a>
                <a href="https://wa.me/201016361440" target="_blank" title="WhatsApp" class="social-btn-top whatsapp-bg"><svg viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.711 2.599 2.669-.699c.971.53 1.77.822 2.791.823h.002c3.18 0 5.767-2.586 5.768-5.766 0-3.18-2.587-5.766-5.77-5.766zm9.969 5.828c0 5.519-4.481 10-10 10-1.761 0-3.424-.46-4.881-1.267l-5.619 1.474 1.499-5.485c-.911-1.516-1.43-3.285-1.43-5.176 0-5.519 4.481-10 10-10 5.519 0 10 4.481 10 10z"/></svg></a>
                <a href="https://t.me/mrmaths22" target="_blank" title="Telegram" class="social-btn-top telegram-bg"><svg viewBox="0 0 24 24"><path d="M12 0c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm5.894 8.221l-1.97 9.28c-.145.658-.537.818-1.084.508l-3-2.21-1.446 1.394c-.16.16-.295.295-.605.295l.213-3.053 5.56-5.023c.242-.213-.054-.333-.373-.121l-6.871 4.326-2.962-.924c-.643-.204-.657-.643.136-.953l11.57-4.461c.537-.194 1.006.131.832.942z"/></svg></a>
                <a href="https://www.tiktok.com/@eng_mohamedghonaim?_r=1&_t=ZS-99VdklZPBUS" target="_blank" title="TikTok" class="social-btn-top tiktok-bg"><svg viewBox="0 0 24 24"><path d="M19.589 6.686a4.793 4.793 0 0 1-3.77-4.245V2h-3.445v13.672a2.896 2.896 0 0 1-5.201 1.743l-.068-.102a2.895 2.895 0 0 1 2.373-4.513c.277 0 .546.039.803.111V9.417a6.338 6.338 0 0 0-.803-.051C6.017 9.366 3.2 12.183 3.2 15.647 3.2 19.11 6.017 22 9.479 22c3.462 0 6.279-2.817 6.279-6.353V9.07c1.378.983 3.054 1.564 4.869 1.584V7.209a4.845 4.845 0 0 1-1.038-.523z"/></svg></a>
                <a href="https://youtube.com/@engineermaths?si=8C6T808VuAU5OMOt" target="_blank" title="YouTube" class="social-btn-top youtube-bg"><svg viewBox="0 0 24 24"><path d="M23.498 6.186a3.016 3.016 0 0 0-2.122-2.136C19.505 3.545 12 3.545 12 3.545s-7.505 0-9.377.505A3.017 3.017 0 0 0 .502 6.186C0 8.07 0 12 0 12s0 3.93.502 5.814a3.016 3.016 0 0 0 2.122 2.136c1.871.505 9.376.505 9.376.505s7.505 0 9.377-.505a3.015 3.015 0 0 0 2.122-2.136C24 15.93 24 12 24 12s0-3.93-.502-5.814zM9.545 15.568V8.432L15.818 12l-6.273 3.568z"/></svg></a>
            </div>
            <div class="rights-text">جميع الحقوق محفوظة لدي م / محمد غنيم 2026</div>
         </div>
    """, unsafe_allow_html=True)
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
        st.markdown(f"""
            <div style="background: {card_bg}; padding: 25px; border-radius: 16px; border: 1px solid {card_border}; margin-bottom: 25px;">
                <h2 style="color: #059669; margin: 0 0 10px 0;">أهلاً، م / محمد غنيم 📌</h2>
                <p style="font-size: 16px; margin: 0; opacity: 0.85;">لوحة تحكم المعلمين – الوصول السريع لأدواتك وخدماتك التعليمية المتقدمة.</p>
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

    bc4, bc5, bc6 = st.columns(3)
    with bc4:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">سجل الحضور والغياب</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">رصد ومتابعة الحصص</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("رصد الحصص", key="card_btn_ses"):
            st.session_state.teacher_page = "add_session"
            st.rerun()

    with bc5:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">الرسائل والدعم</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">صندوق محادثات الطلاب</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("عرض الرسائل", key="card_btn_chat"):
            st.session_state.teacher_page = "chat"
            st.rerun()

    with bc6:
        st.markdown(f"""
            <div style="background: {card_bg}; border: 1px solid {card_border}; border-radius: 14px; padding: 20px; margin-bottom: 15px; height: 160px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h4 style="margin: 0; color: #059669;">تقارير ولي الأمر</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">طباعة التقارير بصيغة PDF</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("عرض التقارير", key="card_btn_rep"):
            st.session_state.teacher_page = "parent_report"
            st.rerun()

elif t_page == "exam_maker":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.markdown("<div class='exam-builder-header'>➕ صانع الامتحانات التفاعلية (مع محرر وقص الصور)</div>", unsafe_allow_html=True)
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
            ex_time_input = st.number_input("مدة الامتحان بالدقائق:", value=45)

    with st.expander("📐 2. إدخال وتحرير الأسئلة", expanded=True):
        q_type_choice = st.selectbox("نوع السؤال:", ["سؤال موضوعي (اختيار من متعدد)", "سؤال مقالي"])
        is_mcq = "موضوعي" in q_type_choice
        q_text_input = st.text_area("نص السؤال المكتوب:")
        
        opt1_txt, opt2_txt, opt3_txt, opt4_txt = "", "", "", ""
        correct_num = 1
        if is_mcq:
            opt1_txt = st.text_input("الخيار أ:")
            opt2_txt = st.text_input("الخيار ب:")
            opt3_txt = st.text_input("الخيار ج:")
            opt4_txt = st.text_input("الخيار د:")
            correct_num = st.selectbox("الإجابة الصحيحة:", [1, 2, 3, 4], format_func=lambda x: f"الخيار ({['أ', 'ب', 'ج', 'د'][x-1]})")

        q_pts_val = st.number_input("درجة السؤال:", value=1.0)

        if st.button("➕ حفظ وإدراج السؤال في الامتحان"):
            if not q_text_input.strip():
                st.error("يرجى كتابة نص السؤال أولاً.")
            else:
                st.session_state.temp_questions.append({
                    "type": "موضوعي" if is_mcq else "مقالي",
                    "text": q_text_input.strip(),
                    "q_img": "", "opt1": opt1_txt, "opt2": opt2_txt, "opt3": opt3_txt, "opt4": opt4_txt,
                    "correct": correct_num, "points": q_pts_val
                })
                st.success("✓ تم حفظ السؤال!")
                st.rerun()

        if st.session_state.temp_questions:
            st.write(f"**الأسئلة الجاهزة: ({len(st.session_state.temp_questions)})**")
            for idx_q, q_item in enumerate(st.session_state.temp_questions):
                st.markdown(f"- س {idx_q+1}: {q_item['text']} (الدرجة: {q_item['points']})")
            if st.button("🗑️ مسح الأسئلة"):
                st.session_state.temp_questions = []
                st.rerun()

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

elif t_page == "question_bank":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📚 بنك الأسئلة الشامل")
    with st.form("qb_form"):
        qb_curr = st.selectbox("المنهج:", list(CURRICULUM_DATA.keys()))
        qb_grade = st.selectbox("الصف:", CURRICULUM_DATA[qb_curr])
        qb_txt = st.text_area("نص السؤال:")
        if st.form_submit_button("حفظ السؤال في البنك"):
            if qb_txt.strip():
                q_payload = {"type": "اختيار من متعدد", "text": qb_txt.strip(), "q_img": "", "opt1": "أ", "opt2": "ب", "opt3": "ج", "opt4": "د", "correct": 1, "points": 1.0}
                new_row = {"معرف_السؤال": f"QB_{datetime.now().strftime('%Y%m%d%H%M%S')}", "المنهج/الدولة": qb_curr, "المجموعة/الصف": qb_grade, "المادة": "الرياضيات", "نوع_السؤال": "اختيار من متعدد", "بيانات_السؤال_JSON": json.dumps(q_payload, ensure_ascii=False)}
                st.session_state.question_bank_df = pd.concat([st.session_state.question_bank_df, pd.DataFrame([new_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("✓ تم حفظ السؤال في البنك بنجاح!")

elif t_page == "videos":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("🎥 إدارة ورفع الفيديوهات التعليمية")
    with st.form("vid_form"):
        v_title = st.text_input("عنوان الدرس:")
        v_curr = st.selectbox("المنهج:", list(CURRICULUM_DATA.keys()))
        v_grade = st.selectbox("الصف:", CURRICULUM_DATA[v_curr])
        v_link = st.text_input("رابط يوتيوب أو فيديو مباشر:")
        if st.form_submit_button("نشر الفيديو للطالب"):
            if v_title.strip():
                new_vid = {"معرف_الفيديو": f"VID_{datetime.now().strftime('%Y%m%d%H%M%S')}", "عنوان_الفيديو": v_title.strip(), "المنهج/الدولة": v_curr, "المجموعة/الصف": v_grade, "رابط_الفيديو": v_link.strip(), "فيديو_base64": "", "تاريخ_الرفع": str(date.today())}
                st.session_state.videos_df = pd.concat([st.session_state.videos_df, pd.DataFrame([new_vid])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("✓ تم نشر الفيديو بنجاح!")

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

elif t_page == "grades":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📈 سجل درجات ونقاط اختبارات الطلاب")
    exam_assessments = st.session_state.assessments_df[st.session_state.assessments_df["النوع"].astype(str).str.contains("اختبار|كويز", na=False)]
    if exam_assessments.empty: st.info("لا توجد نتائج اختبارات مرصودة.")
    else: st.dataframe(exam_assessments, use_container_width=True)

elif t_page == "essays":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📝 تصحيح إجابات الطلاب المقالية")
    essays = st.session_state.essays_df
    if essays.empty: st.info("لا توجد إجابات مقالية قيد الانتظار.")
    else: st.dataframe(essays, use_container_width=True)

elif t_page == "chat":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("💬 صندوق محادثات الطلاب والرد على الأسئلة")
    all_students_with_chat = list(st.session_state.messages_df["اسم الطالب"].dropna().unique())
    if not all_students_with_chat: st.info("لا توجد رسائل واردة.")
    else:
        sel_st = st.selectbox("اختر الطالب:", all_students_with_chat)
        if sel_st:
            st.dataframe(st.session_state.messages_df[st.session_state.messages_df["اسم الطالب"] == sel_st], use_container_width=True)

elif t_page == "students":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("👥 بطاقات الطلاب المسجلين والتحكم الكامل")
    st.dataframe(st.session_state.users_df, use_container_width=True)

elif t_page == "add_session":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📝 رصد حصة جديدة")
    with st.form("sess_f"):
        s_name = st.text_input("اسم الطالب:")
        s_price = st.number_input("سعر الحصة:", value=100.0)
        if st.form_submit_button("حفظ الحصة"):
            if s_name.strip():
                new_row = {"التاريخ": str(date.today()), "اسم الطالب": s_name.strip(), "المنهج/الدولة": "المنهج المصري 🇪🇬", "المجموعة/الصف": "الصف الثالث الثانوي (علمي رياضة)", "الحالة": "حاضر", "سعر الحصة": s_price, "عدد الحصص الكلي": 1, "نظام الدفع": "مقدم", "مستوى الطالب": "ممتاز", "ملاحظات": ""}
                st.session_state.sessions_df = pd.concat([st.session_state.sessions_df, pd.DataFrame([new_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("تم الحفظ بنجاح!")

elif t_page == "add_hw":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📚 رصد واجب يدوي")
    with st.form("hw_f"):
        hw_name = st.text_input("اسم الطالب:")
        hw_title = st.text_input("عنوان الواجب:")
        hw_score = st.number_input("الدرجة المحصلة:", value=10.0)
        if st.form_submit_button("حفظ الواجب"):
            if hw_name.strip() and hw_title.strip():
                new_ass = {"التاريخ": str(date.today()), "اسم الطالب": hw_name.strip(), "النوع": "واجب منزلي", "عنوان التكليف": hw_title.strip(), "الدرجة المحصلة": hw_score, "الدرجة العظمى": 10.0, "حالة التسليم": "تم التسليم", "ملاحظات وتوجيهات": ""}
                st.session_state.assessments_df = pd.concat([st.session_state.assessments_df, pd.DataFrame([new_ass])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df)
                st.success("تم حفظ الواجب بنجاح!")

elif t_page == "edit_records":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("✏️ تعديل السجلات")
    st.dataframe(st.session_state.sessions_df, use_container_width=True)

elif t_page == "all_records":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📊 السجلات الشاملة وإحصائيات الطلاب")
    st.dataframe(st.session_state.users_df, use_container_width=True)

elif t_page == "parent_report":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("🖨️ تقرير ولي الأمر")
    all_names = list(st.session_state.users_df["اسم الطالب"].dropna().unique())
    if all_names:
        sel_st = st.selectbox("اختر الطالب لطباعة تقريره:", all_names)
        if sel_st: st.success(f"جاهز لطباعة تقرير الطالب: {sel_st}")
    else:
        st.info("لا توجد بيانات طلاب كافية.")
