
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
COL_ONLINE_SCHEDULE = ["اسم الطالب", "اسم الأكاديمية", "المنهج/الدولة", "المجموعة/الصف", "رقم الطالب", "رقم مشرف الأكاديمية", "سعر الحصة", "تاريخ الحصة", "ساعة الحصة", "رابط زوم", "حالة فتح الحصة"]

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
    .tiktok-bg   {{ background-color: #000000; border: 1px solid #444; }}
    .youtube-bg  {{ background-color: #FF0000; }}

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
        st.markdown(f"<p style='color: {text_color}; font-size: 16px;'>أدخل اسمك، رقم هاتفك، ومنهجك لتظهر في سجلات المعلم فوراً:</p>", unsafe_allow_html=True)
        
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
                    st.success(f"أهلاً بك يا {g_name.strip()}! تم تسجيل بياناتك بنجاح في سجلات المعلم.")
                    st.rerun()

    elif not st.session_state.logged_student:
        if st.session_state.page_view == "home":
            col_hero_txt, col_hero_img = st.columns([1.3, 1])
            with col_hero_txt:
                st.markdown("<h1 style='color: #059669; font-size: 38px; font-weight: 900; margin-bottom: 10px;'>أهلاً بيكم منورين المنصة! 🚀</h1>", unsafe_allow_html=True)
                st.markdown("<div style='background: #059669; color: #ffffff; padding: 6px 16px; border-radius: 20px; display: inline-block; font-size: 15px; font-weight: 900; margin-bottom: 15px;'>منصة شرح الرياضيات والإحصاء</div>", unsafe_allow_html=True)
                st.markdown(f"<p style='font-size: 17px; font-weight: 800; line-height: 1.8; color: {text_color};'>مع <b>م / محمد غنيم</b>. خبرة متميزة في تدريس الرياضيات والإحصاء للثانوية العامة والمرحلة الإعدادية. آلاف الطلاب حققوا التفوق والدرجات النهائية.</p>", unsafe_allow_html=True)
                
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
                            st.warning("رقم الهاتف هذا مسجل بالفعل! يرجى تسجيل الدخول مباشرة برقم هاتفك.")
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

        st.markdown("<div class='vertical-section-header'>🎥 حصص سريعة وملخصات هامة في أقل من دقيقة</div>", unsafe_allow_html=True)
        col_vid1, col_vid2, col_vid3 = st.columns([1, 2, 1])
        with col_vid2:
            st.video("https://www.youtube.com/watch?v=6PleAxZCNZM")

    else:
        st_user = st.session_state.logged_student
        
        # فلترة ذكية ومرنة لمرحلة الطالب بصرف النظر عن أي فروق بسيطة في الأحرف أو الأقواس
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
                    d_secs = diff_seconds % 60
                    if d_days > 0:
                        timer_text = f"⏳ باقي على الحصة: {d_days} يوم و {d_hours} ساعة و {d_mins} دقيقة"
                    else:
                        timer_text = f"⏳ باقي على الحصة: {d_hours:02d}:{d_mins:02d}:{d_secs:02d}"
                elif diff_seconds == 0:
                    timer_text = "🟢 وقت الحصة الآن!"
                else:
                    timer_text = "📌 موعد الحصة قد حان أو انتهى."

                st.markdown(f"""
                    <div style="background:{card_bg}; border:2px solid #0284c7; border-radius:14px; padding:20px; margin-bottom:15px;">
                        <h4 style="color:#0284c7; margin-top:0;">🏛️ الأكاديمية: {academy_name} | المرحلة: {sched_grade}</h4>
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

        # --- قسم اختبارات عبقري مع فلترة مرنة ومطابقة شاملة ---
        if sub_page == "abqary":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🧠 اختبارات ونتائج موقع عبقري</h3>", unsafe_allow_html=True)
            abq_df = st.session_state.abqary_df
            
            # فلترة مرنة جداً تتأكد من ظهور الامتحان بغض النظر عن الأقواس أو التشكيل
            st_abq = abq_df[abq_df["المجموعة/الصف"].astype(str).str.strip().str.lower().apply(
                lambda x: user_grade_clean in x or user_grade_raw.lower() in x or x in user_grade_clean
            )]
            if st_abq.empty:
                st_abq = abq_df.copy() # لو مفيش تطابق حرفي تام، اعرض الكل عشان الطالب ميتعطلش

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
                            <p style="font-size:15px;">إليك امتحان موقع عبقري معروض بالكامل داخل المنصة:</p>
                        </div>
                    """, unsafe_allow_html=True)

                    if ab_link and ab_link != "nan" and ab_link != "":
                        st.components.v1.iframe(ab_link, height=650, scrolling=True)
                        st.write("")
                        st.link_button(f"🔗 فتح امتحان عبقري في نافذة جديدة 🚀", ab_link, use_container_width=True)

                    st.write("")
                    st.markdown("##### 🔍 استعلام عن نتيجة هذا الاختبار برقم سري:")
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

        # --- الفيديوهات بتصميم درسلي (Sidebar يمين، والفيديو شمال) ---
        elif sub_page == "videos":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>🎥 محتوى الشروحات والفيديوهات التعليمية</h3>", unsafe_allow_html=True)
            v_df = st.session_state.videos_df
            
            # فلترة مرنة للفيديوهات حسب المرحلة
            st_videos = v_df[v_df["المجموعة/الصف"].astype(str).str.strip().str.lower().apply(
                lambda x: user_grade_clean in x or user_grade_raw.lower() in x or x in user_grade_clean
            )]
            if st_videos.empty:
                st_videos = v_df.copy()

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

        # --- بنك الأسئلة والاشتراك والدفع عبر انستا باي ---
        elif sub_page == "bank":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📚 بنك الأسئلة الشامل (مرحلتك الدراسية)</h3>", unsafe_allow_html=True)
            
            curr_user_row = st.session_state.users_df[st.session_state.users_df["رقم الهاتف"].astype(str).str.strip() == str(st_user.get("رقم الهاتف", "")).strip()]
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
                    pay_phone = st.text_input("رقم الهاتف المحول منه:", placeholder="010XXXXXXXX")
                    pay_otp = st.text_input("كود التحقق الخاص بـ InstaPay (OTP):", placeholder="مثال: 4589")
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
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                            st.success("✓ تم إرسال طلب اشتراكك بنجاح! سيقوم المعلم بمراجعة الإيصال وتفعيل حسابك خلال دقائق.")
                
                st.link_button("📲 اضغط هنا للدفع من خلال انستا باي", "https://ipn.eg/S/moghonem2002/instapay/6EyvZs")
            else:
                st.success("🎉 أهلاً بك! حسابك مفعل ومسجل في بنك الأسئلة الخاص بمرحلتك الدراسية.")
                qb_df = st.session_state.question_bank_df
                st_qb = qb_df[qb_df["المجموعة/الصف"].astype(str).str.strip().str.lower().apply(
                    lambda x: user_grade_clean in x or user_grade_raw.lower() in x or x in user_grade_clean
                )]
                if st_qb.empty: st_qb = qb_df.copy()

                if st_qb.empty:
                    st.info("لا توجد أسئلة مضافة في بنك الأسئلة لمرحلتك حالياً.")
                else:
                    for qb_i, qb_r in st_qb.iterrows():
                        raw_json_data = qb_r.get("بيانات_السؤال_JSON", "{}")
                        try:
                            q_data = json.loads(raw_json_data) if pd.notnull(raw_json_data) and str(raw_json_data).strip() else {}
                        except Exception:
                            q_data = {}

                        if not q_data: continue

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

        elif sub_page == "exams":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>✍️ الاختبارات الإلكترونية التفاعلية</h3>", unsafe_allow_html=True)
            st.info("لا توجد اختبارات تفاعلية نشطة حالياً.")

        elif sub_page == "attendance":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📝 تسجيل حضور حصة اليوم</h3>", unsafe_allow_html=True)
            with st.form("att_form"):
                st.date_input("تاريخ الحصة:")
                if st.form_submit_button("تأكيد الحضور"):
                    st.success("تم تسجيل الحضور بنجاح!")

        elif sub_page == "hw_grades":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📊 متابعة درجات الواجبات المنزلية</h3>", unsafe_allow_html=True)
            my_ass = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip()]
            if my_ass.empty: st.info("لا توجد درجات مرصودة.")
            else: st.dataframe(my_ass, use_container_width=True)

        elif sub_page == "exam_grades":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>📈 متابعة درجات الاختبارات والكويزات</h3>", unsafe_allow_html=True)
            st.info("لا توجد اختبارات سابقة.")

        elif sub_page == "chat":
            st.markdown(f"<h3 style='color: {text_color}; font-size: 22px;'>💬 مركز الدردشة والدعم المباشر</h3>", unsafe_allow_html=True)
            st.info("تواصل مع البشمهندس مباشرة عبر مركز الدردشة.")

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
# 2. لوحة تحكم المعلم (الشاملة بجميع الأقسام وزوم الأونلاين والتقارير المالية)
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
if st.sidebar.button("💻 جدول حصص الأونلاين (Zoom)", use_container_width=True):
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
                    <h4 style="margin: 0; color: #059669;">جدول حصص الأونلاين</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.8;">إدارة جداول زوم والطلاب</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("ادارة زوم", key="card_btn_zoom"):
            st.session_state.teacher_page = "online_schedule"
            st.rerun()

    with bc2:
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

elif t_page == "online_schedule":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("💻 إدارة جدول حصص الأونلاين (Zoom) وتحديد مواعيد الطلاب:")

    all_registered_names = sorted(list(set(
        [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    with st.form("add_online_sched_form", clear_on_submit=True):
        col_os1, col_os2 = st.columns(2)
        with col_os1:
            os_student = st.selectbox("اختر الطالب أو اكتبه:", all_registered_names) if all_registered_names else st.text_input("اسم الطالب:")
            os_academy = st.text_input("اسم الأكاديمية:", value="أكاديمية البشمهندس")
            os_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()))
            os_grade = st.selectbox("المرحلة / الصف الدراسي:", CURRICULUM_DATA[os_curr])
            os_phone = st.text_input("رقم الهاتف المحمول للطالب:")
        with col_os2:
            os_sup_phone = st.text_input("رقم مشرف الأكاديمية:")
            os_price = st.number_input("سعر الحصة (جنيه):", min_value=0.0, step=10.0, value=100.0)
            os_date = st.date_input("تقويم موعد الحصة:", value=date.today())
            os_time = st.time_input("ساعة الحصة:", value=time(18, 0))
            os_zoom = st.text_input("رابط زوم المخصص:", value="https://us05web.zoom.us/j/83526892910?pwd=2jWRgATgBRPbXttdnm0QpLwBApsZL4.1")

        if st.form_submit_button("💾 حفظ وإدراج الطالب في جدول الأونلاين"):
            if not str(os_student).strip() or not os_phone.strip():
                st.error("يرجى كتابة اسم الطالب ورقم الهاتف.")
            else:
                new_sched_row = {
                    "اسم الطالب": str(os_student).strip(),
                    "اسم الأكاديمية": os_academy.strip(),
                    "المنهج/الدولة": os_curr,
                    "المجموعة/الصف": os_grade,
                    "رقم الطالب": os_phone.strip(),
                    "رقم مشرف الأكاديمية": os_sup_phone.strip(),
                    "سعر الحصة": os_price,
                    "تاريخ الحصة": str(os_date),
                    "ساعة الحصة": os_time.strftime("%H:%M"),
                    "رابط زوم": os_zoom.strip(),
                    "حالة فتح الحصة": "مغلقة"
                }
                st.session_state.online_schedule_df = pd.concat([st.session_state.online_schedule_df, pd.DataFrame([new_sched_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم إضافة الطالب ({os_student}) إلى جدول الأونلاين بنجاح!")

    st.write("---")
    st.markdown("### 📋 جدول الطلاب المسجلين أونلاين والتحكم بحالة الحصة (فتح/إغلاق):")
    os_state = st.session_state.online_schedule_df
    if os_state.empty:
        st.info("لا توجد حصص أونلاين مسجلة حتى الآن.")
    else:
        for os_idx, os_row in os_state.iterrows():
            st_n = os_row["اسم الطالب"]
            ac_n = os_row.get("اسم الأكاديمية", "أكاديمية")
            dt_n = os_row.get("تاريخ الحصة", "")
            tm_n = os_row.get("ساعة الحصة", "")
            cur_status = os_row["حالة فتح الحصة"]

            with st.expander(f"طالب: {st_n} — أكاديمية: ({ac_n}) | الموعد: {dt_n} {tm_n} | الحالة: [{cur_status}]"):
                with st.form(f"update_zoom_status_{os_idx}"):
                    new_status = st.selectbox("تغيير حالة الحصة (ليراها الطالب زر فتح):", ["مغلقة", "مفتوحة"], index=0 if cur_status == "مغلقة" else 1)
                    edit_zoom_link = st.text_input("تحديث رابط زوم:", value=str(os_row["رابط زوم"]))
                    
                    c_os1, c_os2 = st.columns(2)
                    with c_os1:
                        update_sub = st.form_submit_button("🔄 تحديث حالة الحصة ولينك زوم")
                    with c_os2:
                        del_sub = st.form_submit_button("🗑️ حذف من جدول الأونلاين")

                    if update_sub:
                        os_state.at[os_idx, "حالة فتح الحصة"] = new_status
                        os_state.at[os_idx, "رابط زوم"] = edit_zoom_link.strip()
                        st.session_state.online_schedule_df = os_state
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success("✓ تم تحديث حالة الحصة بنجاح وسيظهر للطالب فوراً!")
                        st.rerun()

                    if del_sub:
                        st.session_state.online_schedule_df = os_state.drop(os_idx).reset_index(drop=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.warning("تم حذف السجل.")
                        st.rerun()

elif t_page == "exam_maker":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.markdown("<div class='exam-builder-header'>➕ إضافة امتحان جديد / محرر وقص الصور وميزة الطباعة PDF</div>", unsafe_allow_html=True)

    if "temp_questions" not in st.session_state: st.session_state.temp_questions = []
    if "q_img_ver" not in st.session_state: st.session_state.q_img_ver = 0
    if "opt1_ver" not in st.session_state: st.session_state.opt1_ver = 0
    if "opt2_ver" not in st.session_state: st.session_state.opt2_ver = 0
    if "opt3_ver" not in st.session_state: st.session_state.opt3_ver = 0
    if "opt4_ver" not in st.session_state: st.session_state.opt4_ver = 0

    if "pasted_q_img" not in st.session_state: st.session_state.pasted_q_img = ""
    if "pasted_opt1_img" not in st.session_state: st.session_state.pasted_opt1_img = ""
    if "pasted_opt2_img" not in st.session_state: st.session_state.pasted_opt2_img = ""
    if "pasted_opt3_img" not in st.session_state: st.session_state.pasted_opt3_img = ""
    if "pasted_opt4_img" not in st.session_state: st.session_state.pasted_opt4_img = ""

    with st.expander("⚙️ 1. الإعدادات الرئيسية للامتحان", expanded=True):
        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            ex_title_input = st.text_input("عنوان الامتحان:*", placeholder="مثال: اختبار الجبر المصور - شهر أكتوبر")
            ex_desc_input = st.text_area("وصف الامتحان:*", placeholder="تعليمات هامة للطلاب قبل البدء...")
            ex_pass_input = st.text_input("كلمة المرور (اختياري):", placeholder="اتركه فارغاً إن لم ترغب بكلمة سر")
        with c_ex2:
            ex_curr_input = st.selectbox("اختر الدولة / المنهج:*", list(CURRICULUM_DATA.keys()), key="mk_c")
            ex_grade_input = st.selectbox("اختر الصف الدراسي:*", CURRICULUM_DATA[ex_curr_input], key="mk_g")
            ex_subject_input = st.selectbox("اختر مادة الامتحان:*", ["الرياضيات (عام)", "الجبر والإحصاء", "الهندسة وحساب المثلثات", "التفاضل والتكامل", "الاستاتيكا والديناميكا", "الرياضيات التطبيقية"])
            ex_term_input = st.selectbox("اختر الفصل الدراسي:*", ["الفصل الدراسي الأول", "الفصل الدراسي الثاني", "مراجعة نهائية"])
            ex_time_input = st.number_input("مدة الامتحان بالدقائق:", min_value=5, max_value=180, value=45, step=5)

    with st.expander("📐 2. إدخال الأسئلة وتحرير وقص الصور", expanded=True):
        q_count = len(st.session_state.temp_questions) + 1
        st.markdown(f"#### إضافة سؤال رقم {q_count}")

        q_type_choice = st.selectbox("نوع السؤال:", ["سؤال موضوعي (اختيار من متعدد)", "سؤال مقالي (خطوات حل ورفع صورة)"])
        is_mcq = "موضوعي" in q_type_choice

        st.markdown("##### 📌 أ) صورة السؤال:")
        col_qp1, col_qp2 = st.columns([1, 1])
        with col_qp1:
            if paste_image_button:
                paste_res = paste_image_button("📋 لصق صورة السؤال من الحافظة", key=f"paste_q_{q_count}_{st.session_state.q_img_ver}")
                if paste_res.image_data is not None:
                    st.session_state.pasted_q_img = pil_to_base64(paste_res.image_data)
        with col_qp2:
            q_file_up = st.file_uploader("أو اسحب صورة السؤال هنا:", type=["jpg", "png", "jpeg"], key=f"upload_q_{q_count}_{st.session_state.q_img_ver}")
            if q_file_up is not None:
                st.session_state.pasted_q_img = base64.b64encode(q_file_up.read()).decode()

        if st.session_state.pasted_q_img:
            st.write("---")
            col_preview, col_del = st.columns([3, 1])
            with col_preview:
                st.markdown("**المعاينة الحالية لصورة السؤال:**")
                st.image(f"data:image/jpeg;base64,{st.session_state.pasted_q_img}", width=400)
            with col_del:
                st.write("")
                st.write("")
                if st.button("🗑️ مسح الصورة بالكامل", key=f"del_q_main_btn_{q_count}_{st.session_state.q_img_ver}"):
                    st.session_state.pasted_q_img = ""
                    st.session_state.q_img_ver += 1
                    st.rerun()

            st.markdown("##### ✂️ أداة قص الصورة بالماوس مباشرة:")
            pil_q = base64_to_pil(st.session_state.pasted_q_img)
            if pil_q:
                st.markdown('<div class="crop-container">', unsafe_allow_html=True)
                st.info("👇 حرّك المربع بالماوس لتحديد الجزء المطلوب قصه ثم اضغط زر 'اعتماد القص':")
                if st_cropper:
                    cropped_img = st_cropper(pil_q, realtime_update=True, box_color='#0052cc', aspect_ratio=None, key=f"cropper_q_{q_count}_{st.session_state.q_img_ver}")
                    if st.button("✂️ اعتماد وحفظ الجزء المقصوص", key=f"btn_confirm_crop_{q_count}"):
                        st.session_state.pasted_q_img = pil_to_base64(cropped_img)
                        st.session_state.q_img_ver += 1
                        st.success("✓ تم قص وتحديث صورة السؤال!")
                        st.rerun()
                else:
                    st.warning("يرجى التأكد من إضافة `streamlit-cropper` في requirements.txt.")
                st.markdown('</div>', unsafe_allow_html=True)

        q_text_input = st.text_area("نص السؤال المكتوب (اختياري):", placeholder="مثال: أوجد قيمة س الموضحة بالرسم:")

        opt1_txt, opt2_txt, opt3_txt, opt4_txt = "", "", "", ""
        correct_num = 1

        if is_mcq:
            st.write("---")
            st.markdown("##### 🎯 ب) خيارات الإجابة الأربعة ومسح صورها:")

            st.markdown("**الخيار الأول (أ):**")
            co1_a, co1_b, co1_c = st.columns([3, 2, 1])
            with co1_a: opt1_txt = st.text_input("نص الخيار (أ):", key=f"t_opt1_{q_count}")
            with co1_b:
                if paste_image_button:
                    p1 = paste_image_button("📋 لصق صورة (أ)", key=f"p_opt1_{q_count}_{st.session_state.opt1_ver}")
                    if p1.image_data is not None: st.session_state.pasted_opt1_img = pil_to_base64(p1.image_data)
                f1 = st.file_uploader("ارفع صورة (أ):", type=["jpg", "png", "jpeg"], key=f"u_opt1_{q_count}_{st.session_state.opt1_ver}")
                if f1 is not None: st.session_state.pasted_opt1_img = base64.b64encode(f1.read()).decode()
            with co1_c:
                st.write("")
                if st.session_state.pasted_opt1_img and st.button("🗑️ مسح", key=f"del_o1_{q_count}_{st.session_state.opt1_ver}"):
                    st.session_state.pasted_opt1_img = ""
                    st.session_state.opt1_ver += 1
                    st.rerun()
            if st.session_state.pasted_opt1_img: st.image(f"data:image/jpeg;base64,{st.session_state.pasted_opt1_img}", width=180)

            st.markdown("**الخيار الثاني (ب):**")
            co2_a, co2_b, co2_c = st.columns([3, 2, 1])
            with co2_a: opt2_txt = st.text_input("نص الخيار (ب):", key=f"t_opt2_{q_count}")
            with co2_b:
                if paste_image_button:
                    p2 = paste_image_button("📋 لصق صورة (ب)", key=f"p_opt2_{q_count}_{st.session_state.opt2_ver}")
                    if p2.image_data is not None: st.session_state.pasted_opt2_img = pil_to_base64(p2.image_data)
                f2 = st.file_uploader("ارفع صورة (ب):", type=["jpg", "png", "jpeg"], key=f"u_opt2_{q_count}_{st.session_state.opt2_ver}")
                if f2 is not None: st.session_state.pasted_opt2_img = base64.b64encode(f2.read()).decode()
            with co2_c:
                st.write("")
                if st.session_state.pasted_opt2_img and st.button("🗑️ مسح", key=f"del_o2_{q_count}_{st.session_state.opt2_ver}"):
                    st.session_state.pasted_opt2_img = ""
                    st.session_state.opt2_ver += 1
                    st.rerun()
            if st.session_state.pasted_opt2_img: st.image(f"data:image/jpeg;base64,{st.session_state.pasted_opt2_img}", width=180)

            st.markdown("**الخيار الثالث (ج):**")
            co3_a, co3_b, co3_c = st.columns([3, 2, 1])
            with co3_a: opt3_txt = st.text_input("نص الخيار (ج):", key=f"t_opt3_{q_count}")
            with co3_b:
                if paste_image_button:
                    p3 = paste_image_button("📋 لصق صورة (ج)", key=f"p_opt3_{q_count}_{st.session_state.opt3_ver}")
                    if p3.image_data is not None: st.session_state.pasted_opt3_img = pil_to_base64(p3.image_data)
                f3 = st.file_uploader("ارفع صورة (ج):", type=["jpg", "png", "jpeg"], key=f"u_opt3_{q_count}_{st.session_state.opt3_ver}")
                if f3 is not None: st.session_state.pasted_opt3_img = base64.b64encode(f3.read()).decode()
            with co3_c:
                st.write("")
                if st.session_state.pasted_opt3_img and st.button("🗑️ مسح", key=f"del_o3_{q_count}_{st.session_state.opt3_ver}"):
                    st.session_state.pasted_opt3_img = ""
                    st.session_state.opt3_ver += 1
                    st.rerun()
            if st.session_state.pasted_opt3_img: st.image(f"data:image/jpeg;base64,{st.session_state.pasted_opt3_img}", width=180)

            st.markdown("**الخيار الرابع (د):**")
            co4_a, co4_b, co4_c = st.columns([3, 2, 1])
            with co4_a: opt4_txt = st.text_input("نص الخيار (د):", key=f"t_opt4_{q_count}")
            with co4_b:
                if paste_image_button:
                    p4 = paste_image_button("📋 لصق صورة (د)", key=f"p_opt4_{q_count}_{st.session_state.opt4_ver}")
                    if p4.image_data is not None: st.session_state.pasted_opt4_img = pil_to_base64(p4.image_data)
                f4 = st.file_uploader("ارفع صورة (د):", type=["jpg", "png", "jpeg"], key=f"u_opt4_{q_count}_{st.session_state.opt4_ver}")
                if f4 is not None: st.session_state.pasted_opt4_img = base64.b64encode(f4.read()).decode()
            with co4_c:
                st.write("")
                if st.session_state.pasted_opt4_img and st.button("🗑️ مسح", key=f"del_o4_{q_count}_{st.session_state.opt4_ver}"):
                    st.session_state.pasted_opt4_img = ""
                    st.session_state.opt4_ver += 1
                    st.rerun()
            if st.session_state.pasted_opt4_img: st.image(f"data:image/jpeg;base64,{st.session_state.pasted_opt4_img}", width=180)

            st.write("---")
            correct_num = st.selectbox("الإجابة الصحيحة هي:*", [1, 2, 3, 4], format_func=lambda x: f"الخيار ({['أ', 'ب', 'ج', 'د'][x-1]})", key=f"cor_{q_count}")

        st.write("---")
        q_pts_val = st.selectbox("علامة هذا السؤال (الدرجة):*", [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0, 10.0], index=1, key=f"pts_{q_count}")

        if st.button(f"➕ حفظ وإدراج السؤال رقم {q_count} في الامتحان"):
            if not q_text_input.strip() and not st.session_state.pasted_q_img:
                st.error("يرجى كتابة نص السؤال أو لصق صورة للسؤال أولاً.")
            else:
                st.session_state.temp_questions.append({
                    "type": "موضوعي" if is_mcq else "مقالي",
                    "text": q_text_input.strip(),
                    "q_img": st.session_state.pasted_q_img,
                    "opt1": opt1_txt.strip(), "opt1_img": st.session_state.pasted_opt1_img,
                    "opt2": opt2_txt.strip(), "opt2_img": st.session_state.pasted_opt2_img,
                    "opt3": opt3_txt.strip(), "opt3_img": st.session_state.pasted_opt3_img,
                    "opt4": opt4_txt.strip(), "opt4_img": st.session_state.pasted_opt4_img,
                    "correct": correct_num,
                    "points": q_pts_val,
                })
                st.session_state.pasted_q_img = ""
                st.session_state.pasted_opt1_img = ""
                st.session_state.pasted_opt2_img = ""
                st.session_state.pasted_opt3_img = ""
                st.session_state.pasted_opt4_img = ""
                st.session_state.q_img_ver += 1
                st.session_state.opt1_ver += 1
                st.session_state.opt2_ver += 1
                st.session_state.opt3_ver += 1
                st.session_state.opt4_ver += 1
                st.success(f"✓ تم حفظ السؤال رقم {q_count} بنجاح!")
                st.rerun()

        if st.session_state.temp_questions:
            st.write(f"**الأسئلة الجاهزة في هذا الاختبار: ({len(st.session_state.temp_questions)}) سؤال**")
            for idx_q, q_item in enumerate(st.session_state.temp_questions):
                st.markdown(f"- **س {idx_q+1} ({q_item['type']}):** {q_item['text'] if q_item['text'] else '[سؤال مصور]'} — *(الدرجة: {q_item['points']})*")
            if st.button("🗑️ مسح كل الأسئلة والبدء من جديد"):
                st.session_state.temp_questions = []
                st.rerun()

    st.write("---")
    if st.button("🚀 حفظ ونشر الامتحان للطلاب"):
        if not ex_title_input.strip():
            st.error("يرجى كتابة عنوان الامتحان أولاً.")
        elif not st.session_state.temp_questions:
            st.error("يرجى إضافة سؤال واحد على الأقل قبل النشر.")
        else:
            new_ex = {
                "معرف_الامتحان": f"EX_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                "عنوان الامتحان": ex_title_input.strip(),
                "وصف الامتحان": ex_desc_input.strip(),
                "كلمة المرور": str(ex_pass_input).strip(),
                "المنهج/الدولة": ex_curr_input,
                "المجموعة/الصف": ex_grade_input,
                "المادة": ex_subject_input,
                "الفصل الدراسي": ex_term_input,
                "مدة الامتحان بالدقائق": int(ex_time_input),
                "الأسئلة_JSON": json.dumps(st.session_state.temp_questions, ensure_ascii=False),
                "تاريخ الإنشاء": str(date.today()),
            }
            st.session_state.exams_df = pd.concat([st.session_state.exams_df, pd.DataFrame([new_ex])], ignore_index=True)
            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
            st.session_state.temp_questions = []
            st.success(f"✓ تم نشر امتحان ({ex_title_input}) بنجاح لصف ({ex_grade_input})!")
            st.rerun()

elif t_page == "question_bank":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📚 بنك الأسئلة الشامل (إضافة أسئلة اختر ومقالي بصور ومواصفات كاملة):")

    if "qb_q_img" not in st.session_state: st.session_state.qb_q_img = ""
    if "qb_ver" not in st.session_state: st.session_state.qb_ver = 0

    qb_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()), key="qbc_c")
    qb_grade = st.selectbox("المرحلة / الصف الدراسي المستهدف:", CURRICULUM_DATA[qb_curr], key="qbc_g")
    qb_subject = st.selectbox("المادة:", ["الرياضيات (عام)", "الجبر والإحصاء", "الهندسة وحساب المثلثات", "التفاضل والتكامل", "الاستاتيكا والديناميكا"], key="qbc_s")
    
    qb_type = st.selectbox("نوع السؤال:", ["اختيار من متعدد", "مقالي"], key="qbc_t", on_change=lambda: st.rerun())
    is_qbank_mcq = (qb_type == "اختيار من متعدد")

    with st.form("add_qbank_form"):
        st.markdown("##### 📌 أ) صورة السؤال (اختياري):")
        qb_file_up = st.file_uploader("رفع صورة السؤال:", type=["jpg", "png", "jpeg"], key=f"qbc_img_up_{st.session_state.qb_ver}")
        if qb_file_up is not None:
            st.session_state.qb_q_img = base64.b64encode(qb_file_up.read()).decode()

        if st.session_state.qb_q_img:
            st.image(f"data:image/jpeg;base64,{st.session_state.qb_q_img}", width=350)
            if st.form_submit_button("🗑️ مسح صورة السؤال"):
                st.session_state.qb_q_img = ""
                st.session_state.qb_ver += 1
                st.rerun()

        qb_text = st.text_area("نص السؤال المكتوب:", key="qbc_txt", placeholder="أكتب نص السؤال هنا...")

        qb_opt1, qb_opt2, qb_opt3, qb_opt4 = "", "", "", ""
        qb_correct = 1

        if is_qbank_mcq:
            st.markdown("---")
            st.markdown("##### 🎯 ب) خيارات الاختيار من متعدد:")
            qb_opt1 = st.text_input("الخيار (أ):", key="qbc_o1")
            qb_opt2 = st.text_input("الخيار (ب):", key="qbc_o2")
            qb_opt3 = st.text_input("الخيار (ج):", key="qbc_o3")
            qb_opt4 = st.text_input("الخيار (د):", key="qbc_o4")
            qb_correct = st.selectbox("الإجابة الصحيحة هي:*", [1, 2, 3, 4], format_func=lambda x: f"الخيار ({['أ', 'ب', 'ج', 'د'][x-1]})", key="qbc_cor")

        st.write("---")
        qb_points = st.number_input("درجة السؤال:", min_value=0.5, max_value=10.0, value=1.0, step=0.5, key="qbc_pts")

        if st.form_submit_button("💾 حفظ وإدراج السؤال في بنك الأسئلة"):
            if not qb_text.strip() and not st.session_state.qb_q_img:
                st.error("يرجى كتابة نص السؤال أو رفع صورة السؤال على الأقل.")
            else:
                q_payload = {
                    "type": qb_type,
                    "text": qb_text.strip(),
                    "q_img": st.session_state.qb_q_img,
                    "opt1": qb_opt1.strip() if is_qbank_mcq else "",
                    "opt2": qb_opt2.strip() if is_qbank_mcq else "",
                    "opt3": qb_opt3.strip() if is_qbank_mcq else "",
                    "opt4": qb_opt4.strip() if is_qbank_mcq else "",
                    "correct": qb_correct if is_qbank_mcq else 1,
                    "points": qb_points
                }
                new_qbank_row = {
                    "معرف_السؤال": f"QB_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "المنهج/الدولة": qb_curr,
                    "المجموعة/الصف": qb_grade,
                    "المادة": qb_subject,
                    "نوع_السؤال": qb_type,
                    "بيانات_السؤال_JSON": json.dumps(q_payload, ensure_ascii=False)
                }
                st.session_state.question_bank_df = pd.concat([st.session_state.question_bank_df, pd.DataFrame([new_qbank_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.session_state.qb_q_img = ""
                st.session_state.qb_ver += 1
                st.success("✓ تم حفظ وإدراج السؤال في بنك الأسئلة بنجاح!")

    st.write("---")
    st.markdown("### 📋 الأسئلة الحالية في بنك الأسئلة:")
    qbf_df = st.session_state.question_bank_df
    if qbf_df.empty:
        st.info("لا توجد أسئلة مضافة في بنك الأسئلة بعد.")
    else:
        for qbi, qbr in qbf_df.iterrows():
            raw_json_data = qbr.get("بيانات_السؤال_JSON", "{}")
            try:
                q_data = json.loads(raw_json_data) if pd.notnull(raw_json_data) and str(raw_json_data).strip() else {}
            except Exception:
                q_data = {}

            if not q_data:
                continue

            with st.expander(f"[{qbr['المجموعة/الصف']}] — {q_data.get('type', 'اختيار من متعدد')} (الدرجة: {q_data.get('points', 1.0)})"):
                st.write(f"**نص السؤال:** {q_data.get('text', '')}")
                if q_data.get("q_img"):
                    st.image(f"data:image/jpeg;base64,{q_data['q_img']}", width=300)
                if q_data.get("type") == "اختيار من متعدد":
                    st.write(f"أ) {q_data.get('opt1','')} | ب) {q_data.get('opt2','')} | ج) {q_data.get('opt3','')} | د) {q_data.get('opt4','')}")
                    corr_idx = int(q_data.get('correct', 1)) - 1
                    corr_letter = ['أ', 'ب', 'ج', 'د'][corr_idx] if 0 <= corr_idx < 4 else 'أ'
                    st.write(f"<b>الإجابة الصحيحة:</b> الخيار ({corr_letter})")
                
                if st.button(f"حذف هذا السؤال 🗑️", key=f"del_qb_{qbi}"):
                    st.session_state.question_bank_df = qbf_df.drop(qbi).reset_index(drop=True)
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                    st.warning("تم حذف السؤال.")
                    st.rerun()

elif t_page == "videos":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("🎥 إدارة ورفع الفيديوهات التعليمية للطلاب:")
    with st.form("upload_video_form", clear_on_submit=True):
        vid_title = st.text_input("عنوان الفيديو / الدرس:")
        vid_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()), key="vid_c")
        vid_grade = st.selectbox("المرحلة / الصف الدراسي المستهدف:", CURRICULUM_DATA[vid_curr], key="vid_g")
        
        vid_upload_file = st.file_uploader("رفع ملف الفيديو (MP4 أو ما شابه):", type=["mp4", "mov", "avi", "mkv", "webm"])
        vid_link_input = st.text_input("أو ضع رابط فيديو (يوتيوب أو رابط مباشر):", placeholder="https://www.youtube.com/watch?v=...")

        if st.form_submit_button("💾 حفظ ونشر الفيديو للطالب"):
            if not vid_title.strip():
                st.error("يرجى كتابة عنوان الفيديو.")
            else:
                v_bytes_str = ""
                if vid_upload_file is not None:
                    file_bytes = vid_upload_file.read()
                    v_bytes_str = base64.b64encode(file_bytes).decode()

                new_vid = {
                    "معرف_الفيديو": f"VID_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "عنوان_الفيديو": vid_title.strip(),
                    "المنهج/الدولة": vid_curr,
                    "المجموعة/الصف": vid_grade,
                    "رابط_الفيديو": vid_link_input.strip() if vid_link_input else "",
                    "فيديو_base64": v_bytes_str,
                    "تاريخ_الرفع": str(date.today())
                }
                st.session_state.videos_df = pd.concat([st.session_state.videos_df, pd.DataFrame([new_vid])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم نشر الفيديو ({vid_title}) بنجاح للمرحلة ({vid_grade})!")

    st.write("---")
    st.markdown("### 📋 الفيديوهات المنشورة مسبقاً:")
    v_df_state = st.session_state.videos_df
    if v_df_state.empty:
        st.info("لا توجد فيديوهات منشورة حالياً.")
    else:
        st.dataframe(v_df_state[["عنوان_الفيديو", "المنهج/الدولة", "المجموعة/الصف", "تاريخ_الرفع"]], use_container_width=True)
        with st.expander("🗑️ حذف فيديو منشور"):
            del_vid_opts = {vi: f"[{vr['المجموعة/الصف']}] {vr['عنوان_الفيديو']}" for vi, vr in v_df_state.iterrows()}
            sel_del_vid = st.selectbox("اختر الفيديو المراد حذفه:", options=list(del_vid_opts.keys()), format_func=lambda x: del_vid_opts[x], key="sel_del_vid")
            if st.button("🚨 تأكيد حذف الفيديو المختار"):
                st.session_state.videos_df = v_df_state.drop(sel_del_vid).reset_index(drop=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("✓ تم حذف الفيديو بنجاح!")
                st.rerun()

    # --- قسم متابعة تعليقات وأسئلة الطلاب على الفيديوهات ---
    st.write("---")
    st.markdown("### 💬 متابعة تعليقات وأسئلة الطلاب على الفيديوهات:")
    vc_state = st.session_state.video_comments_df
    if vc_state.empty:
        st.info("لا توجد تعليقات أو أسئلة من الطلاب على الفيديوهات حتى الآن.")
    else:
        for c_idx, c_row in vc_state.iterrows():
            st.markdown(f"""
                <div style="background:{card_bg}; border:1px solid {card_border}; border-radius:10px; padding:15px; margin-bottom:12px;">
                    <p style="margin:0; color:#10b981; font-size:16px;"><b>درس: {c_row['عنوان_الفيديو']}</b></p>
                    <p style="margin:4px 0; color:#0284c7; font-size:14px;">الطالب: <b>{c_row['اسم الطالب']}</b> — <span style="font-size:12px; opacity:0.7;">{c_row['التاريخ_والوقت']}</span></p>
                    <p style="margin:8px 0 0 0; font-size:16px;">{c_row['نص_التعليق']}</p>
                </div>
            """, unsafe_allow_html=True)
            if st.button(f"🗑️ حذف هذا التعليق", key=f"del_vcomm_{c_idx}"):
                st.session_state.video_comments_df = vc_state.drop(c_idx).reset_index(drop=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("تم حذف التعليق.")
                st.rerun()

elif t_page == "abqary":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("💡 إدارة امتحانات ونتائج موقع عبقري:")
    with st.form("upload_abqary_form", clear_on_submit=True):
        ab_title = st.text_input("عنوان امتحان عبقري:")
        ab_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()), key="ab_c")
        ab_grade = st.selectbox("المرحلة / الصف الدراسي المستهدف:", CURRICULUM_DATA[ab_curr], key="ab_g")
        ab_link = st.text_input("رابط الامتحان على موقع عبقري:", placeholder="https://abqary.com/exam/...")
        res_link = st.text_input("رابط النتيجة (اختياري):", placeholder="https://abqary.com/result/...")
        secret_pass = st.text_input("الرقم السري لإظهار النتيجة للطالب:", placeholder="مثال: 1234 أو كود خاص")

        if st.form_submit_button("💾 حفظ ونشر امتحان عبقري للطالب"):
            if not ab_title.strip() or not ab_link.strip():
                st.error("يرجى كتابة عنوان الامتحان ورابط موقع عبقري.")
            else:
                new_ab = {
                    "معرف_عبقري": f"ABQ_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "عنوان_الإمتحان": ab_title.strip(),
                    "المنهج/الدولة": ab_curr,
                    "المجموعة/الصف": ab_grade,
                    "رابط_الإمتحان": ab_link.strip(),
                    "رابط_النتيجة": res_link.strip() if res_link else "",
                    "الرقم_السري_للنتيجة": secret_pass.strip(),
                    "تاريخ_النشر": str(date.today())
                }
                st.session_state.abqary_df = pd.concat([st.session_state.abqary_df, pd.DataFrame([new_ab])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم نشر امتحان عبقري ({ab_title}) بنجاح!")

    st.write("---")
    st.markdown("### 📋 امتحانات عبقري المنشورة مسبقاً:")
    ab_df_state = st.session_state.abqary_df
    if ab_df_state.empty:
        st.info("لا توجد امتحانات عبقري منشورة حالياً.")
    else:
        st.dataframe(ab_df_state[["عنوان_الإمتحان", "المجموعة/الصف", "رابط_الإمتحان", "تاريخ_النشر"]], use_container_width=True)
        with st.expander("🗑️ حذف امتحان عبقري"):
            del_ab_opts = {ai: f"[{ar['المجموعة/الصف']}] {ar['عنوان_الإمتحان']}" for ai, ar in ab_df_state.iterrows()}
            sel_del_ab = st.selectbox("اختر الامتحان المراد حذفه:", options=list(del_ab_opts.keys()), format_func=lambda x: del_ab_opts[x], key="sel_del_ab")
            if st.button("🚨 تأكيد حذف امتحان عبقري المختار"):
                st.session_state.abqary_df = ab_df_state.drop(sel_del_ab).reset_index(drop=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("✓ تم حذف الامتحان بنجاح!")
                st.rerun()

elif t_page == "grades":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📈 سجل درجات ونقاط اختبارات الطلاب:")
    exam_assessments = st.session_state.assessments_df[
        st.session_state.assessments_df["النوع"].astype(str).str.contains("اختبار|كويز", na=False)
    ].copy()

    if exam_assessments.empty:
        st.info("لا توجد نتائج اختبارات مرصودة للطلاب حتى الآن.")
    else:
        st.dataframe(exam_assessments[["التاريخ", "اسم الطالب", "عنوان التكليف", "الدرجة المحصلة", "الدرجة العظمى", "حالة التسليم", "ملاحظات وتوجيهات"]], use_container_width=True)

        with st.expander("🗑️ حذف نتيجة امتحان لطالب محدد"):
            del_exam_opts = {i: f"{r['اسم الطالب']} - {r['عنوان التكليف']} ({r['التاريخ']})" for i, r in exam_assessments.iterrows()}
            sel_del_exam_idx = st.selectbox("اختر السجل المراد حذفه:", options=list(del_exam_opts.keys()), format_func=lambda x: del_exam_opts[x], key="sel_del_exam")
            if st.button("🚨 تأكيد حذف سجل الامتحان المختار"):
                real_idx = exam_assessments.loc[sel_del_exam_idx].name
                st.session_state.assessments_df = st.session_state.assessments_df.drop(real_idx).reset_index(drop=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("✓ تم حذف النتيجة بنجاح!")
                st.rerun()

        buf_grades = io.BytesIO()
        with pd.ExcelWriter(buf_grades, engine="openpyxl") as writer:
            exam_assessments.to_excel(writer, sheet_name="Exam_Grades", index=False)
        
        c_ex_dl1, c_ex_dl2 = st.columns(2)
        with c_ex_dl1:
            st.download_button(
                label="📥 تصدير درجات الطلاب لملف Excel",
                data=buf_grades.getvalue(),
                file_name="درجات_الاختبارات_للطلاب.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c_ex_dl2:
            grades_pdf_html = """<!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head><meta charset="utf-8"><title>سجل درجات الاختبارات</title></head>
            <body style="font-family: Arial; padding: 25px;" onload="window.print()">
                <h2>سجل درجات ونقاط اختبارات الطلاب - م/ محمد غنيم</h2>
                <table border="1" style="width:100%; border-collapse:collapse; text-align:center; margin-top:20px;">
                    <tr style="background:#f1f5f9;">
                        <th style="padding:10px;">التاريخ</th>
                        <th style="padding:10px;">اسم الطالب</th>
                        <th style="padding:10px;">عنوان الاختبار</th>
                        <th style="padding:10px;">الدرجة</th>
                        <th style="padding:10px;">الحالة</th>
                    </tr>
            """
            for _, r in exam_assessments.iterrows():
                grades_pdf_html += f"""
                    <tr>
                        <td style="padding:8px;">{r['التاريخ']}</td>
                        <td style="padding:8px;">{r['اسم الطالب']}</td>
                        <td style="padding:8px;">{r['عنوان التكليف']}</td>
                        <td style="padding:8px;">{r['الدرجة المحصلة']} / {r['الدرجة العظمى']}</td>
                        <td style="padding:8px;">{r['حالة التسليم']}</td>
                    </tr>
                """
            grades_pdf_html += "</table></body></html>"

            st.download_button(
                label="🖨️ طباعة وتصدير درجات الطلاب PDF",
                data=grades_pdf_html.encode("utf-8"),
                file_name="سجل_درجات_الاختبارات.html",
                mime="application/octet-stream"
            )

elif t_page == "essays":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📝 كنترول تصحيح إجابات الطلاب المقالية وصور الحل:")
    essays = st.session_state.essays_df

    if essays.empty:
        st.info("لا توجد إجابات مقالية مرسلة من الطلاب بعد.")
    else:
        for es_idx, es_row in essays.iterrows():
            st_name = str(es_row.get("اسم الطالب", ""))
            ex_name = str(es_row.get("عنوان الامتحان", ""))
            q_num = es_row.get("رقم السؤال", 1)
            q_text = str(es_row.get("نص السؤال", ""))
            ans_txt = es_row.get("إجابة الطالب النصية", "")
            img_b64_ans = es_row.get("صورة الحل_base64", "")
            q_max = float(es_row.get("درجة السؤال", 1.0))
            status = str(es_row.get("حالة التصحيح", "قيد التصحيح من المعلم"))
            
            raw_fb = es_row.get("ملاحظات المعلم", "")
            t_feedback = str(raw_fb) if pd.notnull(raw_fb) and str(raw_fb) != "nan" else ""

            with st.expander(f"📌 حل الطالب: {st_name} — {ex_name} (س {q_num}) | [{status}]"):
                st.markdown(f"**نص السؤال:** {q_text}")
                st.markdown(f"**إجابة الطالب المكتوبة:**")
                st.write(str(ans_txt) if pd.notnull(ans_txt) and str(ans_txt).strip() and str(ans_txt) != "nan" else "لم يكتب نصاً (أرفق صورة بالأسفل)")

                if pd.notnull(img_b64_ans) and str(img_b64_ans).strip() and str(img_b64_ans) != "nan":
                    st.markdown("**📷 صورة خطوات الحل المرفوعة من الطالب:**")
                    st.image(f"data:image/jpeg;base64,{img_b64_ans}", use_container_width=True)

                with st.form(f"grade_essay_form_{es_idx}"):
                    c_g1, c_g2 = st.columns(2)
                    with c_g1:
                        awarded_score = st.number_input("الدرجة المستحقة:*", min_value=0.0, max_value=q_max, value=float(es_row.get("الدرجة المرصودة", 0.0)) if pd.notnull(es_row.get("الدرجة المرصودة")) else 0.0, step=0.5)
                    with c_g2:
                        teacher_feedback = st.text_input("ملاحظات المعلم وتوجيهه للطالب:", value=t_feedback)

                    c_sub1, c_sub2 = st.columns(2)
                    with c_sub1:
                        submit_grade = st.form_submit_button("💾 اعتماد ورصد الدرجة للطالب")
                    with c_sub2:
                        delete_essay = st.form_submit_button("🗑️ مسح وإلغاء هذه الإجابة المقالية")

                    if submit_grade:
                        essays.at[es_idx, "الدرجة المرصودة"] = awarded_score
                        essays.at[es_idx, "حالة التصحيح"] = "تم التصحيح والاعتماد"
                        essays.at[es_idx, "ملاحظات المعلم"] = str(teacher_feedback).strip()

                        match_ass = st.session_state.assessments_df[
                            (st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_name.strip())
                            & (st.session_state.assessments_df["عنوان التكليف"].astype(str).str.contains(ex_name, na=False))
                        ]
                        if not match_ass.empty:
                            ass_idx = match_ass.index[-1]
                            current_score = float(st.session_state.assessments_df.at[ass_idx, "الدرجة المحصلة"])
                            st.session_state.assessments_df.at[ass_idx, "الدرجة المحصلة"] = current_score + awarded_score
                            st.session_state.assessments_df.at[ass_idx, "ملاحظات وتوجيهات"] = f"تم تصحيح المقالي: +{awarded_score} درجة. {teacher_feedback}"

                        st.session_state.essays_df = essays
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success(f"✓ تم رصد درجة الطالب ({st_name}) بنجاح!")
                        st.rerun()

                    if delete_essay:
                        st.session_state.essays_df = essays.drop(es_idx).reset_index(drop=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.warning("⚠️ تم مسح إجابة المقالي بنجاح.")
                        st.rerun()

elif t_page == "chat":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("💬 صندوق محادثات الطلاب والرد على الأسئلة:")
    all_students_with_chat = sorted(list(set(
        [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.messages_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    if not all_students_with_chat:
        st.info("لا توجد رسائل واردة بعد.")
    else:
        selected_chat_student = st.selectbox("اختر الطالب لمشاهدة محادثته والرد عليه:", all_students_with_chat)
        if selected_chat_student:
            student_thread = st.session_state.messages_df[st.session_state.messages_df["اسم الطالب"].astype(str).str.strip() == selected_chat_student.strip()].copy()
            st.write(f"### سجل المحادثة مع الطالب: **{selected_chat_student}**")

            with st.container():
                if not student_thread.empty:
                    for _, m in student_thread.iterrows():
                        sender = m["المرسل"]
                        t_stamp = m.get("التاريخ_والوقت", "")
                        content = m.get("نص الرسالة", "")
                        img_data = m.get("الصورة_base64", "")
                        if sender == "student":
                            st.markdown(f"<div class='chat-bubble-student'><b>الطالب ({t_stamp}):</b><br>{content}</div>", unsafe_allow_html=True)
                        else:
                            st.markdown(f"<div class='chat-bubble-teacher'><b>أنت (البشمهندس) ({t_stamp}):</b><br>{content}</div>", unsafe_allow_html=True)
                        if pd.notnull(img_data) and str(img_data).strip():
                            st.image(f"data:image/jpeg;base64,{img_data}", width=350)

            with st.form("teacher_reply_form", clear_on_submit=True):
                reply_text = st.text_area("اكتب ردك أو التوجيه للطالب:")
                if st.form_submit_button("📤 إرسال الرد إلى الطالب"):
                    if reply_text.strip():
                        new_rep = {
                            "التاريخ_والوقت": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "اسم الطالب": selected_chat_student.strip(),
                            "المرسل": "teacher", "نص الرسالة": reply_text.strip(), "الصورة_base64": "",
                        }
                        st.session_state.messages_df = pd.concat([st.session_state.messages_df, pd.DataFrame([new_rep])], ignore_index=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success("تم إرسال الرد للبشمهندس بنجاح!")
                        st.rerun()

elif t_page == "students":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("👥 بطاقات الطلاب المسجلين والتحكم الكامل:")
    
    bookings_df_state = st.session_state.bookings_df
    if not bookings_df_state.empty:
        st.markdown("#### 📅 طلبات حجز الدروس أونلاين الواردة:")
        st.dataframe(bookings_df_state, use_container_width=True)
        with st.expander("🗑️ حذف أو إدارة طلب حجز"):
            del_book_opts = {bi: f"{br['اسم الطالب']} - {br['المرحلة_الصف']} ({br['رقم_الهاتف']})" for bi, br in bookings_df_state.iterrows()}
            sel_del_book = st.selectbox("اختر الطلب:", options=list(del_book_opts.keys()), format_func=lambda x: del_book_opts[x], key="sel_del_book")
            if st.button("🚨 تأكيد حذف طلب الحجز المختار"):
                st.session_state.bookings_df = bookings_df_state.drop(sel_del_book).reset_index(drop=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("✓ تم حذف طلب الحجز بنجاح!")
                st.rerun()
        st.write("---")

    bank_req_state = st.session_state.bank_requests_df
    if not bank_req_state.empty:
        st.markdown("#### 💳 طلبات اشتراكات بنك الأسئلة الواردة (تأكيد الدفع):")
        for br_idx, br_row in bank_req_state.iterrows():
            st_req_name = br_row["اسم الطالب"]
            req_phone = br_row["رقم_الهاتف"]
            req_otp = br_row["كود_OTP"]
            req_status = br_row["حالة_الدفع"]
            req_receipt = br_row["إيصال_الدفع_base64"]

            with st.expander(f"طلب اشتراك بنك الأسئلة: الطالب ({st_req_name}) — الهاتف: ({req_phone}) — كود OTP: ({req_otp}) | الحالة: [{req_status}]"):
                if pd.notnull(req_receipt) and str(req_receipt).strip() and str(req_receipt) != "nan":
                    st.markdown("**📷 صورة إيصال التحويل المرفق:**")
                    st.image(f"data:image/jpeg;base64,{req_receipt}", width=350)
                
                c_bk1, c_bk2 = st.columns(2)
                with c_bk1:
                    if st.button(f"✅ تأكيد وتفعيل الاشتراك للطالب {st_req_name}", key=f"confirm_bank_{br_idx}"):
                        bank_req_state.at[br_idx, "حالة_الدفع"] = "مؤكد ومفعل"
                        st.session_state.users_df.loc[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == str(st_req_name).strip(), "حالة_الاشتراك_البنك"] = "مشترك"
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success(f"✓ تم تفعيل اشتراك بنك الأسئلة للطالب {st_req_name} بنجاح!")
                        st.rerun()
                with c_bk2:
                    if st.button(f"🗑️ حذف طلب الاشتراك", key=f"del_bank_req_{br_idx}"):
                        st.session_state.bank_requests_df = bank_req_state.drop(br_idx).reset_index(drop=True)
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.warning("تم حذف طلب الاشتراك.")
                        st.rerun()
        st.write("---")

    all_known_students = sorted(list(set(
        [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.assessments_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    if all_known_students:
        with st.expander("🗑️ حذف طالب محدد نهائياً من كافة السجلات"):
            del_selected_st = st.selectbox("اختر الطالب المراد حذفه نهائياً:", all_known_students, key="del_box_select")
            if st.button("🚨 تأكيد حذف هذا الطالب نهائياً", key="btn_confirm_del_box"):
                delete_student_completely(del_selected_st)
                st.success(f"✓ تم مسح الطالب ({del_selected_st}) نهائياً!")
                st.rerun()

    st.write("---")
    if not all_known_students:
        st.info("لا توجد طلاب مسجلون حالياً.")
    else:
        for idx, st_name in enumerate(all_known_students):
            u_row = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == st_name.strip()]
            s_row = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == st_name.strip()]

            st_curr = u_row.iloc[0].get("المنهج/الدولة", "-") if not u_row.empty else (s_row.iloc[-1].get("المنهج/الدولة", "-") if not s_row.empty else "-")
            st_grade = u_row.iloc[0].get("المجموعة/الصف", "-") if not u_row.empty else (s_row.iloc[-1].get("المجموعة/الصف", "-") if not s_row.empty else "-")
            st_date = u_row.iloc[0].get("تاريخ التسجيل", "-") if not u_row.empty else "-"
            st_pass = u_row.iloc[0].get("كلمة المرور", "-") if not u_row.empty else "-"
            st_phone = u_row.iloc[0].get("رقم الهاتف", "-") if not u_row.empty else "-"
            is_banned = u_row.iloc[0].get("الحالة_حظر") == "محظور" if not u_row.empty else False
            bank_sub_state = u_row.iloc[0].get("حالة_الاشتراك_البنك", "غير مشترك") if not u_row.empty else "غير مشترك"

            status_badge = "🚫 محظور" if is_banned else "✅ نشط"
            badge_color = "#dc2626" if is_banned else "#16a34a"
            bank_badge = "📚 مشترك بالبنك" if bank_sub_state == "مشترك" else "🔒 غير مشترك بالبنك"

            st_sessions_card = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == st_name.strip()]
            total_st_sessions = len(st_sessions_card)
            attended_st_sessions = len(st_sessions_card[st_sessions_card["الحالة"] == "حاضر"])
            total_st_cost = st_sessions_card["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True) if not st_sessions_card.empty else 0.0

            with st.container():
                col_c1, col_c2, col_c3, col_c4 = st.columns([3, 2, 2, 2])
                with col_c1:
                    st.markdown(f"""
                        <h4 style="margin: 0; color: #0052cc;">{st_name}</h4>
                        <p style="margin: 3px 0; font-size: 14px; font-weight: 900;">{st_curr} — {st_grade}</p>
                        <p style="margin: 0; font-size: 13px; color: #64748b;">الهاتف: {st_phone} | كلمة المرور: <b>{st_pass}</b></p>
                        <p style="margin: 2px 0 0 0; font-size: 13px; color: #0284c7;"><b>{bank_badge}</b></p>
                    """, unsafe_allow_html=True)
                with col_c2:
                    st.markdown(f"""
                        <p style="margin:0; font-size:14px; font-weight:900;">حضور: <b>{attended_st_sessions}/{total_st_sessions}</b></p>
                        <p style="margin:0; font-size:14px; font-weight:900; color:#b91c1c;">إجمالي السعر: <b>{total_st_cost:,.1f}</b></p>
                    """, unsafe_allow_html=True)
                with col_c3:
                    st_sessions_html = f"""<!DOCTYPE html>
                    <html dir="rtl" lang="ar">
                    <head><meta charset="utf-8"><title>تقرير الحضور والأسعار - {st_name}</title></head>
                    <body style="font-family: Arial; padding: 25px;" onload="window.print()">
                        <h2>تقرير الحضور وحساب الحصص - م/ محمد غنيم</h2>
                        <p><b>اسم الطالب:</b> {st_name} | <b>المجموعة/الصف:</b> {st_grade}</p>
                        <p><b>إجمالي الحصص الحضورية:</b> {attended_st_sessions} من {total_st_sessions} | <b>المبلغ الإجمالي المستحق:</b> {total_st_cost:,.1f}</p>
                        <hr>
                        <table border="1" style="width:100%; border-collapse:collapse; text-align:center; margin-top:15px;">
                            <tr style="background:#f1f5f9;">
                                <th style="padding:8px;">التاريخ</th>
                                <th style="padding:8px;">الحالة</th>
                                <th style="padding:8px;">سعر الحصة</th>
                                <th style="padding:8px;">المستوى</th>
                                <th style="padding:8px;">ملاحظات</th>
                            </tr>
                    """
                    for _, s_row in st_sessions_card.iterrows():
                        st_sessions_html += f"""
                            <tr>
                                <td style="padding:6px;">{s_row['التاريخ']}</td>
                                <td style="padding:6px;">{s_row['الحالة']}</td>
                                <td style="padding:6px;">{s_row['سعر الحصة']}</td>
                                <td style="padding:6px;">{s_row['مستوى الطالب']}</td>
                                <td style="padding:6px;">{s_row['ملاحظات']}</td>
                            </tr>
                        """
                    st_sessions_html += "</table></body></html>"

                    st.download_button(
                        label="🖨️ طباعة الحضور والأسعار",
                        data=st_sessions_html.encode("utf-8"),
                        file_name=f"حضور_وأسعار_{st_name}.html",
                        mime="application/octet-stream",
                        key=f"print_att_{idx}"
                    )
                with col_c4:
                    if st.button("📊 درجات الطالب", key=f"btn_grades_{idx}"):
                        st.session_state[f"show_grades_{idx}"] = not st.session_state.get(f"show_grades_{idx}", False)

                if st.session_state.get(f"show_grades_{idx}", False):
                    st.markdown(f"**سجل درجات الطالب: {st_name}**")
                    st_grades_df = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_name.strip()]
                    if st_grades_df.empty:
                        st.info("لا توجد درجات مرصودة لهذا الطالب حتى الآن.")
                    else:
                        st.dataframe(st_grades_df[["التاريخ", "النوع", "عنوان التكليف", "الدرجة المحصلة", "الدرجة العظمى", "حالة التسليم", "ملاحظات وتوجيهات"]], use_container_width=True)

                        with st.form(f"clear_st_grades_form_{idx}"):
                            del_grade_choice = st.selectbox("اختر النتيجة المراد مسحها لهذا الطالب:", options=list(st_grades_df.index), format_func=lambda x: f"{st_grades_df.loc[x, 'عنوان التكليف']} ({st_grades_df.loc[x, 'التاريخ']})")
                            if st.form_submit_button("🗑️ مسح هذه الدرجة المحددة للطالب"):
                                st.session_state.assessments_df = st.session_state.assessments_df.drop(del_grade_choice).reset_index(drop=True)
                                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                                st.success("✓ تم مسح الدرجة بنجاح!")
                                st.rerun()

                if is_banned:
                    if st.button("فك الحظر 🔓", key=f"unban_{idx}"):
                        st.session_state.users_df.loc[st.session_state.users_df["اسم الطالب"] == st_name, "الحالة_حظر"] = "نشط"
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                        st.success(f"تم فك حظر {st_name}!")
                        st.rerun()
                else:
                    if st.button("حظر 🚫", key=f"ban_{idx}"):
                        if not u_row.empty:
                            st.session_state.users_df.loc[st.session_state.users_df["اسم الطالب"] == st_name, "الحالة_حظر"] = "محظور"
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                            st.warning(f"تم حظر {st_name}.")
                            st.rerun()
                
                if st.button("حذف نهائي للطالب 🗑️", key=f"del_card_btn_{idx}"):
                    delete_student_completely(st_name)
                    st.success(f"✓ تم مسح الطالب ({st_name}) نهائياً!")
                    st.rerun()
                st.write("---")

elif t_page == "add_session":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("إدخال بيانات الحصة")
    t_curriculum = st.selectbox("اختر المنهج الدراسي:", list(CURRICULUM_DATA.keys()), key="teacher_curr_select")
    t_grades = CURRICULUM_DATA[t_curriculum]

    with st.form("teacher_entry_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            session_date = st.date_input("تاريخ الحصة", value=date.today())
            student_name = st.text_input("اسم الطالب", placeholder="مثال: أحمد محمد")
            group_name = st.selectbox("المرحلة / الصف الدراسي:", t_grades)
            status = st.selectbox("حالة الحضور", ["حاضر", "غائب", "متأخر", "بعذر"])
        with col2:
            price = st.number_input("سعر الحصة", min_value=0.0, step=10.0, value=100.0)
            total_sessions = st.number_input("الحصص المنفذة حتى الآن", min_value=1, step=1, value=1)
            payment_type = st.selectbox("نظام الدفع", ["اشتراك شهري", "مقدم", "مؤخر (بعد الحصة)", "مؤجل"])
            student_level = st.selectbox("المستوى الدراسي", ["ممتاز ⭐⭐⭐", "جيد جداً ⭐⭐", "جيد ⭐", "متوسط", "يحتاج متابعة"])

        notes = st.text_area("ملاحظات الأداء والالتزام بالحصة", placeholder="أداء الطالب في الحصة...")
        if st.form_submit_button("💾 رصد وحفظ الحصة"):
            if not student_name.strip():
                st.error("يرجى كتابة اسم الطالب أولاً.")
            else:
                new_row = {
                    "التاريخ": str(session_date), "اسم الطالب": student_name.strip(),
                    "المنهج/الدولة": t_curriculum, "المجموعة/الصف": group_name,
                    "الحالة": status, "سعر الحصة": price, "عدد الحصص الكلي": total_sessions,
                    "نظام الدفع": payment_type, "مستوى الطالب": student_level, "ملاحظات": notes.strip(),
                }
                st.session_state.sessions_df = pd.concat([st.session_state.sessions_df, pd.DataFrame([new_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم حفظ سجل الحصة للطالب ({student_name}) بنجاح!")

elif t_page == "add_hw":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("إضافة درجات الواجبات المنزلية والاختبارات يدوياً")
    all_registered_names = sorted(list(set(
        [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    with st.form("assessment_form", clear_on_submit=True):
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            ass_student = st.selectbox("اختر الطالب:", all_registered_names) if all_registered_names else st.text_input("اسم الطالب:")
            ass_type = st.selectbox("نوع التكليف الأكاديمي:", ["واجب منزلي", "اختبار دوري", "كويز سريع", "مهمة إضافية"])
            ass_title = st.text_input("عنوان الدرس / التكليف:", placeholder="مثال: تمارين الهندسة صـ 25")
            ass_date = st.date_input("تاريخ الرصد:", value=date.today())
        with col_a2:
            ass_status = st.selectbox("حالة التسليم والحل:", ["تم التسليم كاملاً وبشكل ممتاز", "تم التسليم مع بعض الأخطاء", "تسليم جزئي / ناقص", "لم يتم التسليم (مقصّر)", "غائب عن الاختبار"])
            ass_score = st.number_input("الدرجة المحصلة:", min_value=0.0, step=0.5, value=10.0)
            ass_max = st.number_input("الدرجة العظمى:", min_value=1.0, step=1.0, value=10.0)
            ass_notes = st.text_input("ملاحظة وتوجيه ولي الأمر:", placeholder="توجيه مباشر يظهر في تقرير ولي الأمر...")

        if st.form_submit_button("💾 رصد التقييم وحفظه"):
            if not str(ass_student).strip() or not ass_title.strip():
                st.error("يرجى التأكد من اختيار اسم الطالب وكتابة عنوان التكليف.")
            else:
                new_ass = {
                    "التاريخ": str(ass_date), "اسم الطالب": str(ass_student).strip(),
                    "النوع": ass_type, "عنوان التكليف": ass_title.strip(),
                    "الدرجة المحصلة": ass_score, "الدرجة العظمى": ass_max,
                    "حالة التسليم": ass_status, "ملاحظات وتوجيهات": ass_notes.strip(),
                }
                st.session_state.assessments_df = pd.concat([st.session_state.assessments_df, pd.DataFrame([new_ass])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success(f"✓ تم رصد {ass_type} بنجاح للطالب ({ass_student})!")

    st.write("---")
    st.dataframe(st.session_state.assessments_df, use_container_width=True)

elif t_page == "edit_records":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("مراجعة وتعديل بيانات الحصص")
    df = st.session_state.sessions_df

    if df.empty:
        st.info("لا توجد سجلات حصص مسجلة بعد.")
    else:
        record_options = {idx: f"[{row['التاريخ']}] - {row['اسم الطالب']} ({row.get('المنهج/الدولة', '')} | {row['المجموعة/الصف']}) - {row['الحالة']}" for idx, row in df.iterrows()}
        selected_idx = st.selectbox("اختر السجل المراد تعديله:", options=list(record_options.keys()), format_func=lambda x: record_options[x])
        selected_row = df.loc[selected_idx]

        try: curr_date = datetime.strptime(str(selected_row["التاريخ"]), "%Y-%m-%d").date()
        except Exception: curr_date = date.today()

        with st.form("edit_record_form"):
            c1, c2 = st.columns(2)
            with c1:
                edit_name = st.text_input("اسم الطالب:", value=str(selected_row["اسم الطالب"]))
                saved_curr = selected_row.get("المنهج/الدولة", "")
                curr_list = list(CURRICULUM_DATA.keys())
                curr_idx = curr_list.index(saved_curr) if saved_curr in curr_list else 0
                edit_curr = st.selectbox("المنهج / الدولة:", curr_list, index=curr_idx)
                edit_group = st.text_input("المرحلة / الصف الدراسي:", value=str(selected_row["المجموعة/الصف"]))
                edit_date = st.date_input("التاريخ:", value=curr_date)
                s_opts = ["حاضر", "غائب", "متأخر", "بعذر"]
                edit_status = st.selectbox("الحالة:", s_opts, index=(s_opts.index(selected_row["الحالة"]) if selected_row["الحالة"] in s_opts else 0))
            with c2:
                edit_price = st.number_input("سعر الحصة:", min_value=0.0, step=10.0, value=float(selected_row["سعر الحصة"]) if pd.notnull(selected_row["سعر الحصة"]) else 0.0)
                edit_sessions = st.number_input("عدد الحصص الكلي:", min_value=1, step=1, value=int(selected_row["عدد الحصص الكلي"]) if pd.notnull(selected_row["عدد الحصص الكلي"]) else 1)
                p_opts = ["اشتراك شهري", "مقدم", "مؤخر (بعد الحصة)", "مؤجل"]
                edit_pay = st.selectbox("نظام الدفع:", p_opts, index=(p_opts.index(selected_row["نظام الدفع"]) if selected_row["نظام الدفع"] in p_opts else 0))
                l_opts = ["ممتاز ⭐⭐⭐", "جيد جداً ⭐⭐", "جيد ⭐", "متوسط", "يحتاج متابعة", "قيد التقييم"]
                edit_level = st.selectbox("المستوى:", l_opts, index=(l_opts.index(selected_row["مستوى الطالب"]) if selected_row["مستوى الطالب"] in l_opts else 0))

            edit_notes = st.text_area("الملاحظات والتقييم:", value=str(selected_row["ملاحظات"]))
            b1, b2 = st.columns([1, 4])
            with b1: update_btn = st.form_submit_button("💾 تحديث البيانات")
            with b2: delete_btn = st.form_submit_button("🗑️ حذف السجل")

            if update_btn:
                df.at[selected_idx, "التاريخ"] = str(edit_date)
                df.at[selected_idx, "اسم الطالب"] = edit_name.strip()
                df.at[selected_idx, "المنهج/الدولة"] = edit_curr
                df.at[selected_idx, "المجموعة/الصف"] = edit_group.strip()
                df.at[selected_idx, "الحالة"] = edit_status
                df.at[selected_idx, "سعر الحصة"] = edit_price
                df.at[selected_idx, "عدد الحصص الكلي"] = edit_sessions
                df.at[selected_idx, "نظام الدفع"] = edit_pay
                df.at[selected_idx, "مستوى الطالب"] = edit_level
                df.at[selected_idx, "ملاحظات"] = edit_notes.strip()
                st.session_state.sessions_df = df
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.success("✓ تم تحديث بيانات الحصة بنجاح!")
                st.rerun()

            if delete_btn:
                df = df.drop(selected_idx).reset_index(drop=True)
                st.session_state.sessions_df = df
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df, st.session_state.bookings_df, st.session_state.bank_requests_df, st.session_state.question_bank_df, st.session_state.videos_df, st.session_state.video_comments_df, st.session_state.abqary_df, st.session_state.online_schedule_df)
                st.warning("⚠️ تم حذف السجل.")
                st.rerun()

elif t_page == "all_records":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📊 نظرة شاملة على السجلات وإحصائيات الطلاب")
    
    all_students_master = sorted(list(set(
        [str(s).strip() for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [str(s).strip() for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [str(s).strip() for s in st.session_state.assessments_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    if not all_students_master:
        st.info("لا توجد بيانات طلاب مسجلة حتى الآن.")
    else:
        selected_master_student = st.selectbox("🔍 اختر طالباً لعرض بياناته التفصيلية:", options=all_students_master, key="master_student_sel")
        
        if selected_master_student:
            u_r = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == selected_master_student]
            s_r = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == selected_master_student]
            a_r = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == selected_master_student]

            reg_state = "نعم (مسجل على المنصة)" if not u_r.empty else "لا (مسجل يدويًا)"
            
            grade_val = "-"
            curr_val = "-"
            if not u_r.empty:
                grade_val = str(u_r.iloc[0].get("المجموعة/الصف", "-"))
                curr_val = str(u_r.iloc[0].get("المنهج/الدولة", "-"))
            elif not s_r.empty:
                grade_val = str(s_r.iloc[-1].get("المجموعة/الصف", "-"))
                curr_val = str(s_r.iloc[-1].get("المنهج/الدولة", "-"))

            total_sess = len(s_r)
            attended_sess = len(s_r[s_r["الحالة"] == "حاضر"])
            total_due = s_r["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True) if not s_r.empty else 0.0

            st.markdown(f"""
                <div style="background:{card_bg}; border:2px solid #10b981; border-radius:12px; padding:20px; margin-bottom:10px;">
                    <h4 style="color:#059669; margin-top:0;">👤 ملف الطالب: {selected_master_student}</h4>
                    <p style="font-size:16px; margin:5px 0; color:#0f172a;"><b>حالة التسجيل:</b> {reg_state} | <b>المرحلة/الصف:</b> {grade_val} ({curr_val})</p>
                    <p style="font-size:16px; margin:5px 0; color:#0f172a;"><b>إجمالي الحصص:</b> {total_sess} (حاضر: {attended_sess}) | <b>إجمالي المبلغ المستحق:</b> <span style="color:#dc2626;">{total_due:,.1f} جنيه</span></p>
                </div>
            """, unsafe_allow_html=True)

            st_sessions_pdf = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == selected_master_student].copy()
            st_assessments_pdf = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == selected_master_student].copy()

            sess_rows_html = ""
            for _, sr in st_sessions_pdf.iterrows():
                sess_rows_html += f"<tr><td>{sr['التاريخ']}</td><td>{sr['الحالة']}</td><td>{sr['سعر الحصة']}</td><td>{sr['مستوى الطالب']}</td><td>{sr['ملاحظات']}</td></tr>"

            ass_rows_html = ""
            for _, ar in st_assessments_pdf.iterrows():
                ass_rows_html += f"<tr><td>{ar['التاريخ']}</td><td>{ar['النوع']}</td><td>{ar['عنوان التكليف']}</td><td>{ar['الدرجة المحصلة']} / {ar['الدرجة العظمى']}</td><td>{ar['حالة التسليم']}</td></tr>"

            single_student_pdf_html = f"""<!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head><meta charset="utf-8"><title>تقرير الطالب - {selected_master_student}</title></head>
            <body style="font-family: Arial; padding: 25px;" onload="window.print()">
                <h2>تقرير متابعة الطالب: {selected_master_student}</h2>
                <p><b>المرحلة / الصف:</b> {grade_val} | <b>المنهج:</b> {curr_val}</p>
                <p><b>إجمالي الحصص:</b> {total_sess} | <b>إجمالي الرصيد المستحق:</b> {total_due:,.1f} جنيه</p>
                <hr>
                <h3>سجل الحصص والحضور:</h3>
                <table border="1" style="width:100%; border-collapse:collapse; text-align:center; margin-bottom:20px;">
                    <tr style="background:#f1f5f9;"><th style="padding:8px;">التاريخ</th><th>الحالة</th><th>السعر</th><th>المستوى</th><th>ملاحظات</th></tr>
                    {sess_rows_html if sess_rows_html else "<tr><td colspan='5'>لا توجد حصص مسجلة</td></tr>"}
                </table>
                <h3>سجل الاختبارات والواجبات:</h3>
                <table border="1" style="width:100%; border-collapse:collapse; text-align:center;">
                    <tr style="background:#f1f5f9;"><th style="padding:8px;">التاريخ</th><th>النوع</th><th>العنوان</th><th>الدرجة</th><th>الحالة</th></tr>
                    {ass_rows_html if ass_rows_html else "<tr><td colspan='5'>لا توجد اختبارات مسجلة</td></tr>"}
                </table>
            </body>
            </html>"""

            st.download_button(
                label=f"🖨️ طباعة وتصدير ملف PDF خاص بالطالب ({selected_master_student})",
                data=single_student_pdf_html.encode("utf-8"),
                file_name=f"تقرير_الطالب_{selected_master_student}.html",
                mime="application/octet-stream",
                key=f"dl_single_pdf_{selected_master_student}"
            )

        st.write("---")
        st.markdown("### 📋 جدول ملخص الطلاب الشامل:")
        
        master_data_list = []
        for st_name in all_students_master:
            u_r = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == st_name]
            s_r = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == st_name]
            a_r = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_name]
            
            is_registered = "نعم (مسجل على المنصة)" if not u_r.empty else "لا (مسجل يدويًا)"
            
            grade_val = "-"
            if not u_r.empty:
                grade_val = str(u_r.iloc[0].get("المجموعة/الصف", "-"))
            elif not s_r.empty:
                grade_val = str(s_r.iloc[-1].get("المجموعة/الصف", "-"))

            total_sess = len(s_r)
            attended_sess = len(s_r[s_r["الحالة"] == "حاضر"])
            avg_price = s_r["سعر الحصة"].astype(float, errors="ignore").mean() if not s_r.empty else 0.0
            total_due = s_r["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True) if not s_r.empty else 0.0
            
            exams_count = len(a_r[a_r["النوع"].astype(str).str.contains("اختبار|كويز", na=False)])
            hws_count = len(a_r[a_r["النوع"].astype(str).str.contains("واجب", na=False)])

            master_data_list.append({
                "اسم الطالب": st_name,
                "حالة التسجيل": is_registered,
                "المرحلة/الصف": grade_val,
                "إجمالي الحصص": total_sess,
                "الحصص الحاضرة": attended_sess,
                "متوسط سعر الحصة": f"{avg_price:,.1f}",
                "إجمالي الحساب المستحق": f"{total_due:,.1f}",
                "عدد الاختبارات": exams_count,
                "عدد الواجبات": hws_count
            })

        master_df = pd.DataFrame(master_data_list)
        st.dataframe(master_df, use_container_width=True)

        all_students_pdf_html = """<!DOCTYPE html>
        <html dir="rtl" lang="ar">
        <head><meta charset="utf-8"><title>تقرير السجلات الشاملة لجميع الطلاب</title></head>
        <body style="font-family: Arial; padding: 25px;" onload="window.print()">
            <h2>تقرير السجلات الشاملة وإحصائيات جميع الطلاب - م/ محمد غنيم</h2>
            <table border="1" style="width:100%; border-collapse:collapse; text-align:center; margin-top:15px;">
                <tr style="background:#f1f5f9;">
                    <th style="padding:8px;">اسم الطالب</th>
                    <th>حالة التسجيل</th>
                    <th>المرحلة/الصف</th>
                    <th>إجمالي الحصص</th>
                    <th>الحصص الحاضرة</th>
                    <th>إجمالي الحساب</th>
                    <th>الاختبارات</th>
                    <th>الواجبات</th>
                </tr>
        """
        for item in master_data_list:
            all_students_pdf_html += f"""
                <tr>
                    <td style="padding:6px;">{item['اسم الطالب']}</td>
                    <td>{item['حالة التسجيل']}</td>
                    <td>{item['المرحلة/الصف']}</td>
                    <td>{item['إجمالي الحصص']}</td>
                    <td>{item['الحصص الحاضرة']}</td>
                    <td>{item['إجمالي الحساب المستحق']}</td>
                    <td>{item['عدد الاختبارات']}</td>
                    <td>{item['عدد الواجبات']}</td>
                </tr>
            """
        all_students_pdf_html += "</table></body></html>"

        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                master_df.to_excel(writer, sheet_name="Master_Students_Report", index=False)
                st.session_state.users_df.to_excel(writer, sheet_name="Users", index=False)
                st.session_state.sessions_df.to_excel(writer, sheet_name="Sessions", index=False)
                st.session_state.assessments_df.to_excel(writer, sheet_name="Assessments", index=False)
                st.session_state.exams_df.to_excel(writer, sheet_name="Exams", index=False)
                st.session_state.videos_df.to_excel(writer, sheet_name="Videos", index=False)

            st.download_button(
                label="📥 تصدير تقرير السجلات الشاملة للطلاب (Excel)",
                data=buf.getvalue(), file_name="تقرير_السجلات_الشاملة_للطلاب.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with c_dl2:
            st.download_button(
                label="🖨️ طباعة وتصدير تقرير السجلات الشاملة لجميع الطلاب (PDF)",
                data=all_students_pdf_html.encode("utf-8"),
                file_name="تقرير_السجلات_الشاملة_لجميع_الطلاب.html",
                mime="application/octet-stream"
            )

elif t_page == "parent_report":
    if st.button("⬅️ العودة للرئيسية"): st.session_state.teacher_page = "dashboard"; st.rerun()
    st.subheader("📑 إصدار وطباعة تقرير متابعة الطالب لولي الأمر (شامل الأسعار)")
    all_names = sorted(list(set(
        [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.assessments_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    if not all_names:
        st.info("لا توجد بيانات كافية لإصدار التقرير بعد.")
    else:
        selected_student = st.selectbox("اختر الطالب لإصدار وطباعة تقريره بصيغة PDF:", all_names)

        if selected_student:
            st_sessions = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"] == selected_student].copy().sort_values(by="التاريخ")
            st_assessments = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"] == selected_student].copy().sort_values(by="التاريخ")
            
            u_r_rep = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == selected_student]

            curr_val = "-"
            group_val = "-"
            if not u_r_rep.empty:
                group_val = str(u_r_rep.iloc[0].get("المجموعة/الصف", "-"))
                curr_val = str(u_r_rep.iloc[0].get("المنهج/الدولة", "-"))
            elif not st_sessions.empty:
                group_val = str(st_sessions.iloc[-1].get("المجموعة/الصف", "-"))
                curr_val = str(st_sessions.iloc[-1].get("المنهج/الدولة", "-"))

            level_val = st_sessions.iloc[-1].get("مستوى الطالب", "جيد") if not st_sessions.empty else "جيد"
            pay_val = st_sessions.iloc[-1].get("نظام الدفع", "مؤجل") if not st_sessions.empty else "مؤجل"

            att_cnt = len(st_sessions[st_sessions["الحالة"] == "حاضر"])
            abs_cnt = len(st_sessions[st_sessions["الحالة"] == "غائب"])
            total_sessions_cnt = len(st_sessions)
            att_percentage = (att_cnt / total_sessions_cnt * 100) if total_sessions_cnt > 0 else 0
            total_cost = st_sessions["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True) if not st_sessions.empty else 0.0

            exams = st_assessments[st_assessments["النوع"].str.contains("اختبار", na=False)]
            hws = st_assessments[st_assessments["النوع"].str.contains("واجب", na=False)]
            total_exam_score = exams["الدرجة المحصلة"].astype(float).sum() if not exams.empty else 0
            total_exam_max = exams["الدرجة العظمى"].astype(float).sum() if not exams.empty else 0
            exam_percentage = (total_exam_score / total_exam_max * 100) if total_exam_max > 0 else 100.0

            col_rep1, col_rep2, col_rep3, col_rep4 = st.columns(4)
            col_rep1.metric("نسبة الحضور والالتزام", f"{att_percentage:.1f}%")
            col_rep2.metric("عدد الواجبات المستلمة", len(hws))
            col_rep3.metric("متوسط درجات الاختبارات", f"{exam_percentage:.1f}%")
            col_rep4.metric("إجمالي المبلغ المستحق", f"{total_cost:,.1f}")

            ass_html_rows = ""
            if not st_assessments.empty:
                for _, r in st_assessments.iterrows():
                    ass_html_rows += f"""
                    <tr>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['التاريخ']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #0052cc;">{r['النوع']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['عنوان التكليف']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['الدرجة المحصلة']} / {r['الدرجة العظمى']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['حالة التسليم']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #047857;">{r['ملاحظات وتوجيهات']}</td>
                    </tr>
                    """
            else:
                ass_html_rows = "<tr><td colspan='6' style='padding: 15px; font-weight: 900; border: 2px solid #000;'>لا توجد واجبات أو اختبارات مرصودة حتى الآن.</td></tr>"

            session_html_rows = ""
            if not st_sessions.empty:
                for _, r in st_sessions.iterrows():
                    st_color = "#0f766e" if r["الحالة"] == "حاضر" else ("#b91c1c" if r["الحالة"] == "غائب" else "#b45309")
                    p_curr = float(r.get('سعر الحصة', 0)) if pd.notnull(r.get('سعر الحصة')) else 0.0
                    session_html_rows += f"""
                    <tr>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['التاريخ']}</td>
                        <td style="padding: 10px; border: 2px solid #000; color:{st_color}; font-weight: 900;">{r['الحالة']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #16a34a; font-size: 17px;">{p_curr:,.1f}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #0052cc;">{r['مستوى الطالب']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['ملاحظات']}</td>
                    </tr>
                    """
                session_html_rows += f"""
                <tr style="background-color: #f1f5f9;">
                    <td colspan="2" style="padding: 12px; border: 2px solid #000; font-weight: 900; font-size: 17px;">الإجمالي الكلي المستحق للحصص:</td>
                    <td style="padding: 12px; border: 2px solid #000; font-weight: 900; font-size: 19px; color: #b91c1c;">{total_cost:,.1f}</td>
                    <td colspan="2" style="padding: 12px; border: 2px solid #000; font-weight: 900; font-size: 16px;">نظام الدفع المعتمد: {pay_val}</td>
                </tr>
                """
            else:
                session_html_rows = "<tr><td colspan='5' style='padding: 15px; font-weight: 900; border: 2px solid #000;'>لا توجد حصص مسجلة بعد.</td></tr>"

            teacher_img_tag = f'<img src="data:image/jpeg;base64,{img_b64}" style="width: 100px; height: 100px; border-radius: 50%; border: 3px solid #0052cc; object-fit: cover;">' if img_b64 else ""

            parent_report_html = f"""<!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head>
                <meta charset="utf-8">
                <title>تقرير متابعة ولي الأمر - {selected_student}</title>
                <style>
                    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@800;900&display=swap');
                    body {{ font-family: 'Cairo', Tahoma, Arial, sans-serif; padding: 30px; color: #000000 !important; background-color: #ffffff !important; font-weight: 900; }}
                    .header-box {{ border-bottom: 3px solid #0052cc; padding-bottom: 15px; margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; }}
                    .brand-name {{ color: #0052cc; margin: 0; font-size: 32px; font-weight: 900; }}
                    .brand-sub {{ color: #000000; margin: 4px 0; font-size: 16px; font-weight: 900; }}
                    .parent-notice {{ background-color: #f0fdf4; border: 2px solid #16a34a; padding: 12px 18px; border-radius: 8px; margin-bottom: 25px; font-size: 17px; font-weight: 900; color: #14532d; }}
                    .stats-box {{ width: 100%; border-collapse: collapse; margin-bottom: 25px; }}
                    .stats-box th, .stats-box td {{ border: 2px solid #000000; padding: 10px; text-align: center; font-size: 16px; font-weight: 900; }}
                    .stats-box th {{ background: #f1f5f9; color: #000000; font-weight: 900; }}
                    .section-title {{ margin-top: 25px; margin-bottom: 10px; color: #0052cc; font-size: 20px; font-weight: 900; border-bottom: 2px solid #0052cc; padding-bottom: 6px; display: inline-block; }}
                    .table-main {{ width: 100%; border-collapse: collapse; text-align: center; margin-top: 10px; margin-bottom: 20px; }}
                    .table-main th {{ background-color: #f8fafc; color: #000000; font-weight: 900; font-size: 16px; padding: 10px; border: 2px solid #000000; }}
                    .footer-note {{ margin-top: 35px; text-align: center; color: #000000; font-size: 16px; font-weight: 900; border-top: 2px solid #000000; padding-top: 15px; }}
                </style>
            </head>
            <body onload="window.print()">
                <div class="header-box">
                    <div style="display: flex; align-items: center; gap: 20px;">
                        {teacher_img_tag}
                        <div>
                            <h2 class="brand-name">م/ محمد غنيم 📐</h2>
                            <p class="brand-sub">تقرير التقييم الدوري والحساب المالي الشامل لولي الأمر</p>
                            <p style="margin: 2px 0; color: #000000; font-size: 16px; font-weight: 900;"><b>المنهج والمرحلة:</b> {curr_val} — {group_val}</p>
                        </div>
                    </div>
                    <div style="text-align: left;">
                        <h2 style="color: #0052cc; margin: 0; font-size: 26px; font-weight: 900;">الطالب: {selected_student}</h2>
                        <p style="margin: 5px 0 0 0; color: #000000; font-size: 15px; font-weight: 900;">تاريخ إصدار التقرير: {date.today()}</p>
                    </div>
                </div>
                <table class="stats-box">
                    <tr>
                        <th>إجمالي الحصص</th><th>مرات الحضور</th><th>مرات الغياب</th><th>نسبة الالتزام بالحضور</th><th>إجمالي الحساب المستحق</th><th>المستوى العام</th>
                    </tr>
                    <tr>
                        <td style="font-weight: 900;">{total_sessions_cnt}</td>
                        <td style="color: #0f766e; font-weight: 900;">{att_cnt}</td>
                        <td style="color: #b91c1c; font-weight: 900;">{abs_cnt}</td>
                        <td style="font-weight: 900; color: #0052cc;">{att_percentage:.1f}%</td>
                        <td style="font-weight: 900; font-size: 18px; color: #b91c1c;">{total_cost:,.1f}</td>
                        <td style="color: #0052cc; font-weight: 900;">{level_val}</td>
                    </tr>
                </table>
                <div class="section-title">1. تقرير الواجبات المنزلية والاختبارات الدورية:</div>
                <table class="table-main">
                    <tr><th>التاريخ</th><th>النوع</th><th>عنوان التكليف / الاختبار</th><th>الدرجة المحصلة</th><th>حالة التسليم والالتزام</th><th>ملاحظات وتوجيهات</th></tr>
                    {ass_html_rows}
                </table>
                <div class="section-title">2. سجل الحضور وتفاصيل سعر كل حصة:</div>
                <table class="table-main">
                    <tr><th>التاريخ</th><th>حالة الحضور</th><th>سعر الحصة</th><th>مستوى الطالب بالحصة</th><th>ملاحظات التفاعل والاستيعاب</th></tr>
                    {session_html_rows}
                </table>
                <div class="footer-note">مع تحيات: <b>م/ محمد غنيم | منصة الرياضيات والإحصاء</b> — رقم التواصل المباشر: 01016361440 🌟</div>
            </body>
            </html>"""

            st.download_button(
                label=f"🖨️ تحميل وطباعة تقرير ولي الأمر لـ ({selected_student}) بصيغة PDF",
                data=parent_report_html.encode("utf-8"),
                file_name=f"تقرير_ولي_الأمر_{selected_student}.html",
                mime="application/octet-stream",
            )
