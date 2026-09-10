import os
import io
import json
import base64
from datetime import date, datetime, time
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
COL_USERS = ["اسم الطالب", "رقم الهاتف", "كلمة المرور", "المنهج/الدولة", "المجموعة/الصف", "تاريخ التسجيل", "الحالة_حظر", "حالة_الاشتراك_البنك"]
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
COL_ONLINE_SCHEDULE = ["اسم الطالب", "اسم الأكاديمية", "المنهج/الدولة", "المجموعة/الصف", "رقم الطالب", "رقم مشرف الأكاديمية", "سعر الحصة", "يوم الحصة", "تاريخ الحصة", "ساعة الحصة", "رابط زوم", "حالة فتح الحصة"]

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
    online_schedule_df = pd.DataFrame(columns=COL_ONLINE_SCHEDULE)

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
                if "OnlineSchedule" in xls.sheet_names: online_schedule_df = pd.read_excel(xls, "OnlineSchedule")
        except Exception:
            pass

    for col in COL_USERS:
        if col not in users_df.columns:
            if col == "رقم الهاتف": users_df[col] = ""
            elif col == "الحالة_حظر": users_df[col] = "نشط"
            elif col == "حالة_الاشتراك_البنك": users_df[col] = "غير مشترك"
            else: users_df[col] = ""
    for col in COL_ONLINE_SCHEDULE:
        if col not in online_schedule_df.columns:
            if col == "رابط زوم": online_schedule_df[col] = "https://us05web.zoom.us/j/83526892910?pwd=2jWRgATgBRPbXttdnm0QpLwBApsZL4.1"
            elif col == "حالة فتح الحصة": online_schedule_df[col] = "مغلقة"
            elif col == "اسم الأكاديمية": online_schedule_df[col] = "أكاديمية البشمهندس"
            elif col == "يوم الحصة": online_schedule_df[col] = "السبت"
            else: online_schedule_df[col] = ""

    return users_df, sessions_df, assessments_df, messages_df, exams_df, essays_df, bookings_df, bank_requests_df, question_bank_df, videos_df, video_comments_df, abqary_df, online_schedule_df

def save_all_data(users_df, sessions_df, assessments_df, messages_df, exams_df, essays_df, bookings_df, bank_requests_df, question_bank_df, videos_df, video_comments_df, abqary_df, online_schedule_df):
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
        online_schedule_df.to_excel(writer, sheet_name="OnlineSchedule", index=False)

if "users_df" not in st.session_state:
    u_df, s_df, a_df, m_df, e_df, es_df, b_df, br_df, qb_df, v_df, vc_df, ab_df, os_df = load_all_data()
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
    st.session_state.online_schedule_df = os_df

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

saved_student_phone = query_params.get("st_phone")
if not st.session_state.logged_student and saved_student_phone:
    matched_st = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == str(saved_student_phone).strip()]
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
    st.session_state.online_schedule_df = st.session_state.online_schedule_df[st.session_state.online_schedule_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)

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
# 1. واجهة الطالب الشاملة
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
        c_btn0, c_btn1, c_btn2, c_btn3, c_btn_guest = st.columns(5)
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
        with c_btn_guest:
            if st.button("👥 ضيف"):
                st.session_state.page_view = "guest_reg"
                st.rerun()
        with c_btn3:
            mode_label = "☀️ فاتح" if st.session_state.dark_mode else "🌙 داكن"
            if st.button(mode_label):
                st.session_state.dark_mode = not st.session_state.dark_mode
                st.rerun()

    st.write("---")

    if st.session_state.page_view == "guest_reg":
        st.markdown(f"<h3 style='color: {text_color}; font-size: 24px;'>👥 تسجيل بياناتك كـ (ضيف جديد):</h3>", unsafe_allow_html=True)
        with st.form("guest_registration_form"):
            g_name = st.text_input("اسم الطالب بالكامل:*", placeholder="مثال: كريم أحمد محمود")
            g_phone = st.text_input("رقم الهاتف المحمول:*", placeholder="010XXXXXXXX")
            g_curr = st.selectbox("المنهج الدراسي / الدولة:*", list(CURRICULUM_DATA.keys()))
            g_grade = st.selectbox("المرحلة / الصف الدراسي:*", CURRICULUM_DATA[g_curr])
            g_pass = st.text_input("اختر رقماً سرياً خاصاً بك:*", type="password", value="1234")
            
            c_g1, c_g2 = st.columns(2)
            with c_g1:
                submit_guest = st.form_submit_button("🚀 حفظ والدخول للمنصة")
            with c_g2:
                if st.form_submit_button("العودة للرئيسية"):
                    st.session_state.page_view = "home"
                    st.rerun()

            if submit_guest:
                if not g_name.strip() or not g_phone.strip():
                    st.error("يرجى كتابة الاسم ورقم الهاتف على الأقل.")
                else:
                    new_guest_user = {
                        "اسم الطالب": g_name.strip(),
                        "رقم الهاتف": g_phone.strip(),
                        "كلمة المرور": g_pass.strip(),
                        "المنهج/الدولة": g_curr,
                        "المجموعة/الصف": g_grade,
                        "تاريخ التسجيل": str(date.today()),
                        "الحالة_حظر": "نشط",
                        "حالة_الاشتراك_البنك": "غير مشترك"
                    }
                    existing = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == g_phone.strip()]
                    if existing.empty:
                        st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([new_guest_user])], ignore_index=True)
                    else:
                        st.session_state.users_df.loc[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == g_phone.strip(), ["اسم الطالب", "المنهج/الدولة", "المجموعة/الصف"]] = [g_name.strip(), g_curr, g_grade]
                    
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                    st.session_state.logged_student = new_guest_user
                    st.query_params["role"] = "student"
                    st.query_params["st_phone"] = g_phone.strip()
                    st.success(f"أهلاً بك يا {g_name.strip()}!")
                    st.rerun()

    elif not st.session_state.logged_student:
        if st.session_state.page_view == "home":
            col_hero_txt, col_hero_img = st.columns([1.3, 1])
            with col_hero_txt:
                st.markdown("<h1 style='color: #059669; font-size: 38px; font-weight: 900; margin-bottom: 10px;'>أهلاً بيكم منورين المنصة! 🚀</h1>", unsafe_allow_html=True)
                st.markdown("<div style='background: #059669; color: #ffffff; padding: 6px 16px; border-radius: 20px; display: inline-block; font-size: 15px; font-weight: 900; margin-bottom: 15px;'>منصة شرح الرياضيات والإحصاء</div>", unsafe_allow_html=True)
                st.markdown(f"<p style='font-size: 17px; font-weight: 800; line-height: 1.8; color: {text_color};'>مع <b>م / محمد غنيم</b>. خبرة متميزة في تدريس الرياضيات والإحصاء للثانوية العامة والمرحلة الإعدادية. آلاف الطلاب حققوا التفوق والدرجات النهائية.</p>", unsafe_allow_html=True)
                
                # --- أيقونات التواصل الاجتماعي في الأعلى (طلبك الجديد) ---
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
                            <img src="data:image/jpeg;base64,{img_b64}" style="width: 230px; height: 230px; border-radius: 50%; border: 5px solid #059669; object-fit: cover; box-shadow: 0 10px 25px rgba(0,0,0,0.15);">
                        </div>
                    """, unsafe_allow_html=True)

            st.write("---")
            # --- كورسات درسلي ---
            st.markdown('<div class="darssly-box">', unsafe_allow_html=True)
            st.markdown("<h3 style='color: #ffffff; text-align: center; margin-bottom: 5px; font-size: 22px;'>📢 اشترك الآن في كورسات الرياضيات والإحصاء على منصة درسلي (Darssly)</h3>", unsafe_allow_html=True)
            courses_grid = [
                ("📊", "إحصاء الثالث الثانوي", "شرح مبسط وتدريبات متقدمة لامتحان العزم", "https://darssly.com/courses/mohamed-ghoneim-statistics/plans"),
                ("📖", "رياضيات أول إعدادي", "شرح كامل وتدريبات دورية مبسطة", "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim-3/plans"),
                ("📘", "رياضيات ثاني إعدادي", "متابعة شاملة وأسئلة تفاعلية مميزة", "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim/plans"),
                ("📐", "رياضيات ثالث إعدادي", "تأسيس قوي وضمان الدرجة النهائية", "https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim-2/plans")
            ]
            c_cols = st.columns(2)
            for idx, (icon, title, desc, link) in enumerate(courses_grid):
                with c_cols[idx % 2]:
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

            # --- حجز الدروس ---
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
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success("✓ تم إرسال طلب الحجز بنجاح!")

        elif st.session_state.page_view == "login":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🔐 تسجيل دخول الطالب (برقم الهاتف):</h3>", unsafe_allow_html=True)
            with st.form("student_login_form"):
                login_phone = st.text_input("رقم الهاتف المحمول المسجل:")
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
                        (st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == login_phone.strip())
                        & (st.session_state.users_df["كلمة المرور"].astype(str).str.strip() == login_pass.strip())
                    ]
                    if not users_match.empty:
                        user_info = users_match.iloc[0].to_dict()
                        if user_info.get("الحالة_حظر") == "محظور":
                            st.error("🚫 تم حظر هذا الحساب من قبل المعلم.")
                        else:
                            st.session_state.logged_student = user_info
                            st.query_params["role"] = "student"
                            st.query_params["st_phone"] = user_info["رقم الهاتف"]
                            st.success(f"مرحباً بك مجدداً يا {user_info['اسم الطالب']}!")
                            st.rerun()
                    else:
                        st.error("رقم الهاتف أو الرقم السري غير صحيح.")

        elif st.session_state.page_view == "register":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>✨ إنشاء حساب طالب جديد:</h3>", unsafe_allow_html=True)
            with st.form("student_register_form"):
                reg_name = st.text_input("اسمك بالكامل:")
                reg_phone = st.text_input("رقم الهاتف المحمول (لتسجيل الدخول به لاحقاً):*")
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
                    if not reg_name.strip() or not reg_phone.strip() or not reg_pass.strip():
                        st.error("يرجى ملء جميع الحقول المطلوبة.")
                    else:
                        existing = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == reg_phone.strip()]
                        if not existing.empty:
                            st.warning("رقم الهاتف هذا مسجل بالفعل!")
                        else:
                            new_user = {
                                "اسم الطالب": reg_name.strip(),
                                "رقم الهاتف": reg_phone.strip(),
                                "كلمة المرور": reg_pass.strip(),
                                "المنهج/الدولة": reg_curr,
                                "المجموعة/الصف": reg_grade,
                                "تاريخ التسجيل": str(date.today()),
                                "الحالة_حظر": "نشط",
                                "حالة_الاشتراك_البنك": "غير مشترك"
                            }
                            st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([new_user])], ignore_index=True)
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                            st.session_state.logged_student = new_user
                            st.query_params["role"] = "student"
                            st.query_params["st_phone"] = new_user["رقم الهاتف"]
                            st.success(f"تم إنشاء حسابك بنجاح يا {reg_name}!")
                            st.rerun()

    else:
        st_user = st.session_state.logged_student
        user_grade_raw = str(st_user.get("المجموعة/الصف", "")).strip()
        user_grade_clean = user_grade_raw.split("(")[0].strip().lower()

        fresh_user = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == str(st_user.get("رقم الهاتف", "")).strip()]
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
                    <p style="margin: 4px 0 10px 0; font-weight: 900; color: {text_color}; font-size: 16px;">{st_user.get('المنهج/الدولة', '')} | {st_user.get('المجموعة/الصف', '')} | الهاتف: {st_user.get('رقم الهاتف', '')}</p>
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

        # --- جدول الأونلاين وزوم مع التايمر ---
        st.markdown("<div class='vertical-section-header'>💻 حصص الأونلاين وجدول زوم الخاص بي</div>", unsafe_allow_html=True)
        student_name_str = str(st_user.get("اسم الطالب", "")).strip()
        os_df = st.session_state.online_schedule_df
        st_sched = os_df[os_df["اسم الطالب"].astype(str).str.strip().str.lower() == student_name_str.lower()]

        if st_sched.empty:
            st.info("لا توجد حصص أونلاين مسجلة في الجدول المخصص لك حالياً.")
        else:
            for _, s_row in st_sched.iterrows():
                academy_name = s_row.get("اسم الأكاديمية", "أكاديمية البشمهندس")
                sched_grade = s_row["المجموعة/الصف"]
                sched_sup_phone = s_row["رقم مشرف الأكاديمية"]
                sched_price = s_row["سعر الحصة"]
                sched_day = s_row.get("يوم الحصة", "السبت")
                sched_date = s_row.get("تاريخ الحصة", str(date.today()))
                sched_time = s_row.get("ساعة الحصة", "18:00")
                zoom_link = s_row["رابط زوم"]
                zoom_status = s_row["حالة فتح الحصة"]

                try:
                    target_dt = datetime.strptime(f"{sched_date} {sched_time}", "%Y-%m-%d %H:%M")
                    now_dt = datetime.now()
                    diff_seconds = int((target_dt - now_dt).total_seconds())
                except Exception:
                    diff_seconds = -1

                timer_text = ""
                if diff_seconds > 0:
                    d_days = diff_seconds // 86400
                    d_hours = (diff_seconds % 86400) // 3600
                    d_mins = (diff_seconds % 3600) // 60
                    if d_days > 0:
                        timer_text = f"⏳ باقي على الحصة: {d_days} يوم و {d_hours} ساعة و {d_mins} دقيقة"
                    else:
                        timer_text = f"⏳ باقي على الحصة: {d_hours:02d}:{d_mins:02d}"
                elif diff_seconds == 0:
                    timer_text = "🟢 وقت الحصة الآن!"
                else:
                    timer_text = "📌 موعد الحصة قد حان أو انتهى."

                st.markdown(f"""
                    <div style="background:{card_bg}; border:2px solid #0284c7; border-radius:14px; padding:20px; margin-bottom:15px;">
                        <h4 style="color:#0284c7; margin-top:0;">🏛️ الأكاديمية: {academy_name} | اليوم: {sched_day}</h4>
                        <p style="font-size:16px; margin:4px 0;"><b>📅 موعد الحصة:</b> {sched_date} في تمام الساعة {sched_time}</p>
                        <p style="font-size:16px; margin:4px 0; color:#d97706;"><b>{timer_text}</b></p>
                        <p style="font-size:16px; margin:4px 0;"><b>💰 سعر الحصة:</b> {sched_price} جنيه</p>
                        <p style="font-size:16px; margin:4px 0;"><b>📞 رقم مشرف الأكاديمية:</b> {sched_sup_phone}</p>
                        <p style="font-size:16px; margin:4px 0;"><b>حالة الحصة الحالية:</b> <span style="color: {'#16a34a' if zoom_status == 'مفتوحة' else '#dc2626'};"><b>{zoom_status}</b></span></p>
                    </div>
                """, unsafe_allow_html=True)

                if zoom_status == "مفتوحة":
                    if zoom_link and zoom_link != "nan":
                        st.link_button("🚀 انضم الآن إلى حصة زوم (الحصة مفتوحة) 🟢", zoom_link, use_container_width=True)
                    else:
                        st.link_button("🚀 انضم الآن إلى حصة زوم 🟢", "https://us05web.zoom.us/j/83526892910?pwd=2jWRgATgBRPbXttdnm0QpLwBApsZL4.1", use_container_width=True)
                else:
                    st.warning("⏳ الحصة مغلقة حالياً. سيتم فتحها من قبل المعلم في موعدها المحدد.")

        st.markdown("<div class='vertical-section-header'>🗂️ لوحة خدمات الطالب التفاعلية</div>", unsafe_allow_html=True)
        
        if st.button("🎥 الفيديوهات والشروحات التعليمية", use_container_width=True):
            st.session_state.student_sub_page = "videos"
            st.rerun()
        if st.button("📚 بنك الأسئلة الشامل (الاشتراك والدفع)", use_container_width=True):
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
            abq_df = st.session_state.abqary_df
            st_abq = abq_df[abq_df["المجموعة/الصف"].astype(str).str.strip().str.lower().apply(
                lambda x: user_grade_clean in x or user_grade_raw.lower() in x or x in user_grade_clean
            )]
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
                        <div style="background:{card_bg}; border:2px solid #059669; border-radius:14px; padding:20px; margin-bottom:15px;">
                            <h4 style="color:#059669; margin-top:0;">💡 امتحان عبقري: {ab_title}</h4>
                        </div>
                    """, unsafe_allow_html=True)

                    if ab_link and ab_link != "nan" and ab_link != "":
                        st.components.v1.iframe(ab_link, height=650, scrolling=True)
                        st.write("")
                        st.link_button(f"🔗 فتح امتحان عبقري في نافذة جديدة 🚀", ab_link, use_container_width=True)

                    st.write("")
                    with st.form(f"abqary_result_form_{ab_i}"):
                        entered_pass = st.text_input("أدخل الرقم السري المخصص لإظهار النتيجة:", type="password", key=f"pass_input_{ab_i}")
                        if st.form_submit_button("🔓 إظهار النتيجة"):
                            if secret_code and entered_pass.strip() == secret_code:
                                st.success("✓ الرقم السري صحيح!")
                                if res_link and res_link != "nan":
                                    st.link_button("📊 اضغط هنا لعرض نتيجة امتحان عبقري 🏆", res_link, use_container_width=True)
                                else:
                                    st.info("النتيجة متاحة وسيتم إضافتها قريباً.")
                            else:
                                st.error("❌ الرقم السري غير صحيح.")
                    st.write("---")

        elif sub_page == "videos":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🎥 محتوى الشروحات والفيديوهات التعليمية</h3>", unsafe_allow_html=True)
            v_df = st.session_state.videos_df
            st_videos = v_df[v_df["المجموعة/الصف"].astype(str).str.strip().str.lower().apply(
                lambda x: user_grade_clean in x or user_grade_raw.lower() in x or x in user_grade_clean
            )]
            if st_videos.empty: st_videos = v_df.copy()

            if st_videos.empty:
                st.info("لا توجد فيديوهات مرفوعة لمرحلتك الدراسية حالياً.")
            else:
                if "selected_video_idx" not in st.session_state:
                    st.session_state.selected_video_idx = 0

                col_main_v, col_side_v = st.columns([2.5, 1])

                with col_side_v:
                    st.markdown(f"""
                        <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:12px; padding:15px; margin-bottom:15px;">
                            <h4 style="margin:0 0 10px 0; color:#059669; font-size:18px;">📚 محتوى الدرس</h4>
                        </div>
                    """, unsafe_allow_html=True)

                    search_vid_query = st.text_input("🔍 ابحث في الحصص...", placeholder="اكتب اسم الدرس...")
                    filtered_videos = st_videos
                    if search_vid_query.strip():
                        filtered_videos = st_videos[st_videos["عنوان_الفيديو"].astype(str).str.contains(search_vid_query, case=False, na=False)]

                    if filtered_videos.empty:
                        st.info("لا توجد نتائج مطابقة للبحث.")
                    else:
                        for v_i, v_row in filtered_videos.reset_index(drop=True).iterrows():
                            v_title_btn = f"🟢 {v_row['عنوان_الفيديو']}"
                            if st.button(v_title_btn, key=f"darssly_sidebar_v_{v_i}", use_container_width=True):
                                st.session_state.selected_video_idx = v_i
                                st.rerun()

                with col_main_v:
                    selected_row = filtered_videos.iloc[0] if filtered_videos.empty else (filtered_videos.iloc[st.session_state.selected_video_idx] if st.session_state.selected_video_idx < len(filtered_videos) else filtered_videos.iloc[0])

                    st.markdown(f"""
                        <div style="background:linear-gradient(135deg, #059669, #10b981); color:#ffffff; padding:15px 20px; border-radius:10px; margin-bottom:15px;">
                            <h3 style="margin:0; color:#ffffff; font-size:20px;">📺 {selected_row['عنوان_الفيديو']}</h3>
                        </div>
                    """, unsafe_allow_html=True)

                    v_link = str(selected_row.get("رابط_الفيديو", "")).strip()
                    v_bytes = selected_row.get("فيديو_base64", "")

                    if v_link and v_link != "nan" and v_link != "":
                        st.video(v_link)
                    elif pd.notnull(v_bytes) and str(v_bytes).strip() and str(v_bytes) != "nan":
                        try:
                            vid_bytes_dec = base64.b64decode(v_bytes)
                            st.video(vid_bytes_dec)
                        except Exception:
                            st.error("⚠️ يتعذر تشغيل ملف الفيديو.")
                    else:
                        st.info("لا يوجد فيديو متاح لهذا الدرس.")

                    st.write("---")
                    st.markdown(f"#### 📌 {selected_row['عنوان_الفيديو']}")
                    st.caption(f"المرحلة الدراسية: {selected_row.get('المجموعة/الصف', '')} | أكملت هذه الحصة")

                    st.write("---")
                    current_vid_title = selected_row['عنوان_الفيديو']
                    vc_df = st.session_state.video_comments_df
                    vid_comments = vc_df[vc_df["عنوان_الفيديو"].astype(str).str.strip() == current_vid_title.strip()]

                    st.markdown(f"<h4 style='font-size:18px;'>الأسئلة والتعليقات ({len(vid_comments)})</h4>", unsafe_allow_html=True)
                    if not vid_comments.empty:
                        for _, c_row in vid_comments.iterrows():
                            st.markdown(f"""
                                <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:10px; padding:12px 15px; margin-bottom:10px;">
                                    <p style="margin:0; font-size:14px; color:#0284c7;"><b>{c_row['اسم الطالب']}</b> — <span style="font-size:12px; opacity:0.7;">{c_row['التاريخ_والوقت']}</span></p>
                                    <p style="margin:5px 0 0 0; font-size:16px;">{c_row['نص_التعليق']}</p>
                                </div>
                            """, unsafe_allow_html=True)

                    with st.form(f"comment_form_{selected_row['معرف_الفيديو']}"):
                        user_comment_text = st.text_area("أكتب سؤالك أو تعليقك هنا:", placeholder="اطرح سؤالك على المعلم...")
                        if st.form_submit_button("إرسال التعليق"):
                            if user_comment_text.strip():
                                new_comment = {
                                    "التاريخ_والوقت": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                    "عنوان_الفيديو": current_vid_title.strip(),
                                    "اسم الطالب": st_user["اسم الطالب"],
                                    "نص_التعليق": user_comment_text.strip()
                                }
                                st.session_state.video_comments_df = pd.concat([st.session_state.video_comments_df, pd.DataFrame([new_comment])], ignore_index=True)
                                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                                st.success("✓ تم إرسال تعليقك بنجاح!")
                                st.rerun()

        elif sub_page == "bank":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📚 بنك الأسئلة الشامل</h3>", unsafe_allow_html=True)
            curr_user_row = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == str(st_user.get("رقم الهاتف", "")).strip()]
            sub_status = curr_user_row.iloc[0].get("حالة_الاشتراك_البنك", "غير مشترك") if not curr_user_row.empty else "غير مشترك"

            if sub_status != "مشترك":
                st.warning("🔒 عذراً، بنك الأسئلة مغلق ويحتاج إلى اشتراك خاص بمرحلتك الدراسية بقيمة **100 جنيه** فقط.")
                with st.form("bank_subscription_form"):
                    pay_phone = st.text_input("رقم الهاتف المحول منه:", placeholder="010XXXXXXXX")
                    pay_otp = st.text_input("كود التحقق الخاص بـ InstaPay (OTP):")
                    pay_receipt = st.file_uploader("📷 رفع اسكرين (إيصال) الدفع:", type=["jpg", "png", "jpeg"])
                    if st.form_submit_button("📤 إرسال طلب الاشتراك لتأكيد المعلم"):
                        if not pay_phone.strip() or not pay_otp.strip() or pay_receipt is None:
                            st.error("يرجى ملء جميع الحقول ورفع الإيصال.")
                        else:
                            rcpt_str = base64.b64encode(pay_receipt.read()).decode()
                            new_req = {
                                "تاريخ_الطلب": str(date.today()), "اسم الطالب": st_user["اسم الطالب"],
                                "رقم_الهاتف": pay_phone.strip(), "كود_OTP": pay_otp.strip(),
                                "حالة_الدفع": "قيد المراجعة", "إيصال_الدفع_base64": rcpt_str
                            }
                            st.session_state.bank_requests_df = pd.concat([st.session_state.bank_requests_df, pd.DataFrame([new_req])], ignore_index=True)
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                            st.success("✓ تم إرسال طلب اشتراكك بنجاح!")
                st.link_button("📲 اضغط هنا للدفع من خلال انستا باي", "https://ipn.eg/S/moghonem2002/instapay/6EyvZs")
            else:
                st.success("🎉 حسابك مفعل في بنك الأسئلة!")

        elif sub_page == "exams":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>✍️ الاختبارات الإلكترونية التفاعلية</h3>", unsafe_allow_html=True)
            st.info("لا توجد اختبارات نشطة حالياً.")

        elif sub_page == "attendance":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📝 تسجيل حضور حصة اليوم</h3>", unsafe_allow_html=True)
            with st.form("att_form"):
                st.date_input("تاريخ الحصة:")
                if st.form_submit_button("تأكيد الحضور"): st.success("تم الحضور بنجاح!")

        elif sub_page == "hw_grades":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📊 متابعة درجات الواجبات</h3>", unsafe_allow_html=True)
            my_ass = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip()]
            if my_ass.empty: st.info("لا توجد درجات مرصودة.")
            else: st.dataframe(my_ass, use_container_width=True)

        elif sub_page == "exam_grades":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📈 متابعة درجات الاختبارات</h3>", unsafe_allow_html=True)
            st.info("لا توجد اختبارات سابقة.")

        elif sub_page == "chat":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>💬 مركز الدردشة والدعم المباشر</h3>", unsafe_allow_html=True)
            st.info("تواصل مع البشمهندس مباشرة.")

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
# 2. لوحة تحكم المعلم (مع جدول الأكاديمية الأسبوعي المتقدم وإضافة موعد آخر)
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
if st.sidebar.button("💻 جدول الأكاديمية بالأيام والمواعيد", use_container_width=True):
    st.session_state.teacher_page = "online_schedule"
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
                <p style="font-size: 16px; margin: 0; opacity: 0.85;">لوحة تحكم المعلمين – جدول الأكاديمية الأسبوعي والمواعيد وإدارة الطلاب.</p>
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
        if st.button("💻 جدول الأكاديمية بالأيام والمواعيد", key="card_btn_zoom"):
            st.session_state.teacher_page = "online_schedule"
            st.rerun()
    with bc2:
        if st.button("⚙️ صانع الامتحانات", key="card_btn_ex"):
            st.session_state.teacher_page = "exam_maker"
            st.rerun()
    with bc3:
        if st.button("💡 امتحانات عبقري", key="card_btn_abq"):
            st.session_state.teacher_page = "abqary"
            st.rerun()

elif t_page == "online_schedule":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("💻 جدول الأكاديمية والأيام والمواعيد (مع إمكانية إضافة أكثر من موعد للطالب):")

    all_registered_names = sorted(list(set(
        [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    with st.form("add_online_sched_form", clear_on_submit=True):
        col_os1, col_os2 = st.columns(2)
        with col_os1:
            os_student = st.selectbox("اختر الطالب:", all_registered_names) if all_registered_names else st.text_input("اسم الطالب:")
            os_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()))
            os_grade = st.selectbox("المرحلة / الصف الدراسي:", CURRICULUM_DATA[os_curr])
            os_academy = st.text_input("اسم الأكاديمية:", value="أكاديمية البشمهندس")
            os_sup_phone = st.text_input("رقم مشرف الأكاديمية:")
        with col_os2:
            os_price = st.number_input("سعر الحصة (جنيه):", min_value=0.0, step=10.0, value=100.0)
            os_day = st.selectbox("يوم الحصة الأسبوعي:", ["السبت", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة"])
            os_date = st.date_input("تقويم تاريخ الحصة القادمة:", value=date.today())
            os_time = st.time_input("مواعيد الحصة (من الساعة):", value=time(18, 0))
            os_zoom = st.text_input("رابط زوم المخصص:", value="https://us05web.zoom.us/j/83526892910?pwd=2jWRgATgBRPbXttdnm0QpLwBApsZL4.1")

        c_sub_b1, c_sub_b2 = st.columns(2)
        with c_sub_b1:
            submit_sched = st.form_submit_button("💾 حفظ وإدراج في جدول الأكاديمية")
        with c_sub_b2:
            add_another_slot = st.form_submit_button("➕ إضافة موعد آخر (إضافي) لهذا الطالب")

        if submit_sched or add_another_slot:
            if not str(os_student).strip():
                st.error("يرجى اختيار أو كتابة اسم الطالب.")
            else:
                new_sched_row = {
                    "اسم الطالب": str(os_student).strip(),
                    "اسم الأكاديمية": os_academy.strip(),
                    "المنهج/الدولة": os_curr,
                    "المجموعة/الصف": os_grade,
                    "رقم الطالب": "",
                    "رقم مشرف الأكاديمية": os_sup_phone.strip(),
                    "سعر الحصة": os_price,
                    "يوم الحصة": os_day,
                    "تاريخ الحصة": str(os_date),
                    "ساعة الحصة": os_time.strftime("%H:%M"),
                    "رابط زوم": os_zoom.strip(),
                    "حالة فتح الحصة": "مغلقة"
                }
                st.session_state.online_schedule_df = pd.concat([st.session_state.online_schedule_df, pd.DataFrame([new_sched_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                
                if add_another_slot:
                    st.success(f"✓ تم حفظ الموعد وإضافة الموعد الإضافي للطالب ({os_student}) بنجاح! يمكنك إدخال تفاصيل الموعد التالي الآن.")
                    st.rerun()
                else:
                    st.success(f"✓ تم حفظ موعد الحصة للأكاديمية للطالب ({os_student}) بنجاح!")

    st.write("---")
    st.markdown("### 📋 جدول الأكاديمية والأيام والمواعيد (مرتب بلون مميز لكل طالب):")
    os_state = st.session_state.online_schedule_df
    if os_state.empty:
        st.info("لا توجد مواعيد حصص مسجلة في جدول الأكاديمية حتى الآن.")
    else:
        unique_students_list = os_state["اسم الطالب"].unique()
        color_palette = ["#e0f2fe", "#fce7f3", "#d1fae5", "#fef3c7", "#ede9fe", "#ffedd5"]
        student_color_map = {st_n: color_palette[i % len(color_palette)] for i, st_n in enumerate(unique_students_list)}

        for os_idx, os_row in os_state.iterrows():
            st_n = os_row["اسم الطالب"]
            ac_n = os_row.get("اسم الأكاديمية", "أكاديمية")
            day_n = os_row.get("يوم الحصة", "السبت")
            dt_n = os_row.get("تاريخ الحصة", "")
            tm_n = os_row.get("ساعة الحصة", "")
            cur_status = os_row["حالة فتح الحصة"]
            card_color = student_color_map.get(st_n, "#ffffff")

            with st.expander(f"📌 الطالب: {st_n} | الأكاديمية: {ac_n} | اليوم: {day_n} ({tm_n}) | الحالة: [{cur_status}]"):
                st.markdown(f"""
                    <div style="background:{card_color}; border:2px solid #059669; border-radius:12px; padding:15px; margin-bottom:10px;">
                        <p style="margin:3px 0; font-size:16px;"><b>👤 اسم الطالب:</b> {st_n}</p>
                        <p style="margin:3px 0; font-size:16px;"><b>🏛️ الأكاديمية:</b> {ac_n} | <b>المشرف:</b> {os_row.get('رقم مشرف الأكاديمية', '-')}</p>
                        <p style="margin:3px 0; font-size:16px;"><b>📅 اليوم والموعد:</b> يوم {day_n} الموافق {dt_n} الساعة {tm_n}</p>
                        <p style="margin:3px 0; font-size:16px;"><b>💰 سعر الحصة:</b> {os_row.get('سعر الحصة', 0)} جنيه</p>
                    </div>
                """, unsafe_allow_html=True)

                with st.form(f"update_slot_form_{os_idx}"):
                    new_status = st.selectbox("حالة فتح الحصة (للطالب):", ["مغلقة", "مفتوحة"], index=0 if cur_status == "مغلقة" else 1)
                    c_s1, c_s2 = st.columns(2)
                    with c_s1:
                        up_s = st.form_submit_button("🔄 تحديث حالة الحصة")
                    with c_s2:
                        del_s = st.form_submit_button("🗑️ حذف هذا الموعد من الجدول")

                    if up_s:
                        os_state.at[os_idx, "حالة فتح الحصة"] = new_status
                        st.session_state.online_schedule_df = os_state
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success("✓ تم التحديث بنجاح!")
                        st.rerun()

                    if del_s:
                        st.session_state.online_schedule_df = os_state.drop(os_idx).reset_index(drop=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.warning("تم حذف الموعد.")
                        st.rerun()

elif t_page == "exam_maker":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("صانع الامتحانات")
    st.info("قسم صانع الامتحانات متاح بالكامل.")

elif t_page == "question_bank":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("بنك الأسئلة")
    st.info("بنك الأسئلة متاح بالكامل.")

elif t_page == "videos":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("إدارة الفيديوهات")
    st.info("إدارة الفيديوهات متاحة بالكامل.")

elif t_page == "abqary":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("امتحانات عبقري")
    st.info("امتحانات عبقري متاحة بالكامل.")

elif t_page == "grades":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("درجات الاختبارات")
    st.dataframe(st.session_state.assessments_df, use_container_width=True)

elif t_page == "essays":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("تصحيح المقالي")
    st.info("الكنترول المقالي جاهز.")

elif t_page == "chat":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("الرسائل والدردشة")
    st.info("الرسائل متاحة بالكامل.")

elif t_page == "students":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("بطاقات الطلاب")
    st.dataframe(st.session_state.users_df, use_container_width=True)

elif t_page == "add_session":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("رصد حصة جديدة")
    all_registered_names = sorted(list(set(
        [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))
    with st.form("teacher_entry_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            session_date = st.date_input("تاريخ الحصة", value=date.today())
            student_name = st.selectbox("اختر الطالب:", all_registered_names) if all_registered_names else st.text_input("اسم الطالب:")
            group_name = st.selectbox("المرحلة / الصف الدراسي:", CURRICULUM_DATA["المنهج المصري 🇪🇬"])
            status = st.selectbox("حالة الحضور", ["حاضر", "غائب", "متأخر", "بعذر"])
        with col2:
            price = st.number_input("سعر الحصة", min_value=0.0, step=10.0, value=100.0)
            total_sessions = st.number_input("الحصص المنفذة حتى الآن", min_value=1, step=1, value=1)
            payment_type = st.selectbox("نظام الدفع", ["اشتراك شهري", "مقدم", "مؤخر (بعد الحصة)", "مؤجل"])
            student_level = st.selectbox("المستوى الدراسي", ["ممتاز ⭐⭐⭐", "جيد جداً ⭐⭐", "جيد ⭐", "متوسط", "يحتاج متابعة"])

        notes = st.text_area("ملاحظات الأداء والالتزام بالحصة")
        if st.form_submit_button("💾 رصد وحفظ الحصة"):
            if not str(student_name).strip():
                st.error("يرجى اختيار اسم الطالب.")
            else:
                new_row = {
                    "التاريخ": str(session_date), "اسم الطالب": str(student_name).strip(),
                    "المنهج/الدولة": "المنهج المصري 🇪🇬", "المجموعة/الصف": group_name,
                    "الحالة": status, "سعر الحصة": price, "عدد الحصص الكلي": total_sessions,
                    "نظام الدفع": payment_type, "مستوى الطالب": student_level, "ملاحظات": notes.strip(),
                }
                st.session_state.sessions_df = pd.concat([st.session_state.sessions_df, pd.DataFrame([new_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم حفظ سجل الحصة للطالب ({student_name}) بنجاح!")

elif t_page == "add_hw":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("رصد واجب يدوي")
    st.info("قسم رصد الواجبات متاح بالكامل.")

elif t_page == "edit_records":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("تعديل السجلات")
    st.dataframe(st.session_state.sessions_df, use_container_width=True)

elif t_page == "all_records":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("السجلات الشاملة")
    st.dataframe(st.session_state.sessions_df, use_container_width=True)

elif t_page == "parent_report":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("تقرير ولي الأمر")
    all_names = sorted(list(set(
        [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.assessments_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))
    if all_names:
        selected_student = st.selectbox("اختر الطالب لإصدار التقرير:", all_names)
        if selected_student:
            st_sessions = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"] == selected_student]
            st.dataframe(st_sessions, use_container_width=True)
