import os
import io
import json
import base64
from datetime import date, datetime
import pandas as pd
from PIL import Image
import streamlit as st

FILE_NAME = "سجل_الغياب_والحصص.xlsx"
IMG_NAME = "teacher.jpg"

st.set_page_config(
    page_title="البشمهندس X الرياضة | متابعة الطلاب والامتحانات",
    page_icon="📐",
    layout="wide",
)

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

img_b64 = get_image_base64(found_img_path)

COL_SESSIONS = ["التاريخ", "اسم الطالب", "المنهج/الدولة", "المجموعة/الصف", "الحالة", "سعر الحصة", "عدد الحصص الكلي", "نظام الدفع", "مستوى الطالب", "ملاحظات"]
COL_USERS = ["اسم الطالب", "كلمة المرور", "المنهج/الدولة", "المجموعة/الصف", "تاريخ التسجيل", "الحالة_حظر"]
COL_ASSESSMENTS = ["التاريخ", "اسم الطالب", "النوع", "عنوان التكليف", "الدرجة المحصلة", "الدرجة العظمى", "حالة التسليم", "ملاحظات وتوجيهات"]
COL_MESSAGES = ["التاريخ_والوقت", "اسم الطالب", "المرسل", "نص الرسالة", "الصورة_base64"]
COL_EXAMS = ["معرف_الامتحان", "عنوان الامتحان", "وصف الامتحان", "كلمة المرور", "المنهج/الدولة", "المجموعة/الصف", "المادة", "الفصل الدراسي", "مدة الامتحان بالدقائق", "الأسئلة_JSON", "تاريخ الإنشاء"]
COL_ESSAYS = ["معرف_الحل", "معرف_الامتحان", "عنوان الامتحان", "اسم الطالب", "رقم السؤال", "نص السؤال", "إجابة الطالب النصية", "صورة الحل_base64", "درجة السؤال", "الدرجة المرصودة", "حالة التصحيح", "ملاحظات المعلم", "تاريخ الحل"]

def load_all_data():
    users_df = pd.DataFrame(columns=COL_USERS)
    sessions_df = pd.DataFrame(columns=COL_SESSIONS)
    assessments_df = pd.DataFrame(columns=COL_ASSESSMENTS)
    messages_df = pd.DataFrame(columns=COL_MESSAGES)
    exams_df = pd.DataFrame(columns=COL_EXAMS)
    essays_df = pd.DataFrame(columns=COL_ESSAYS)

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
        except Exception:
            pass

    for col in COL_USERS:
        if col not in users_df.columns: users_df[col] = "نشط" if col == "الحالة_حظر" else ""
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

    return users_df, sessions_df, assessments_df, messages_df, exams_df, essays_df

def save_all_data(users_df, sessions_df, assessments_df, messages_df, exams_df, essays_df):
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
        users_df.to_excel(writer, sheet_name="Users", index=False)
        sessions_df.to_excel(writer, sheet_name="Sessions", index=False)
        assessments_df.to_excel(writer, sheet_name="Assessments", index=False)
        messages_df.to_excel(writer, sheet_name="Messages", index=False)
        exams_df.to_excel(writer, sheet_name="Exams", index=False)
        essays_df.to_excel(writer, sheet_name="Essays", index=False)

if "users_df" not in st.session_state:
    u_df, s_df, a_df, m_df, e_df, es_df = load_all_data()
    st.session_state.users_df = u_df
    st.session_state.sessions_df = s_df
    st.session_state.assessments_df = a_df
    st.session_state.messages_df = m_df
    st.session_state.exams_df = e_df
    st.session_state.essays_df = es_df

def delete_student_completely(student_name_to_del):
    target = student_name_to_del.strip()
    st.session_state.users_df = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.sessions_df = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.assessments_df = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.messages_df = st.session_state.messages_df[st.session_state.messages_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    st.session_state.essays_df = st.session_state.essays_df[st.session_state.essays_df["اسم الطالب"].astype(str).str.strip() != target].reset_index(drop=True)
    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@600;700;800;900&display=swap');
    html, body, [class*="css"], p, span, label, div {
        font-family: 'Cairo', sans-serif !important;
        direction: rtl;
        text-align: right;
    }
    .brand-banner {
        background: linear-gradient(135deg, #0b1f3a, #1e3c72) !important;
        padding: 24px 30px;
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 25px;
        box-shadow: 0 6px 18px rgba(11, 31, 58, 0.35);
    }
    .brand-title {
        font-size: 34px !important;
        font-weight: 900 !important;
        color: #ffffff !important;
        margin: 0 !important;
    }
    .brand-subtitle {
        color: #e0edff !important;
        font-size: 16px !important;
        font-weight: 700 !important;
        margin-top: 6px !important;
    }
    .exam-builder-header {
        background-color: #f59e0b;
        color: #ffffff !important;
        padding: 14px 20px;
        border-radius: 10px 10px 0 0;
        font-size: 20px;
        font-weight: 900;
        margin-bottom: 15px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .editor-toolbar {
        background: #f1f5f9;
        border: 1px solid #cbd5e1;
        border-radius: 8px 8px 0 0;
        padding: 8px 12px;
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
        font-size: 14px;
        font-weight: 800;
    }
    .vertical-section-header {
        background: linear-gradient(135deg, #0284c7, #0369a1);
        color: #ffffff !important;
        padding: 12px 18px;
        border-radius: 10px;
        margin-top: 25px;
        margin-bottom: 15px;
        font-size: 18px;
        font-weight: 900;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .chat-bubble-student {
        background-color: #e0f2fe;
        color: #0369a1;
        padding: 12px 16px;
        border-radius: 14px 14px 0 14px;
        margin-bottom: 10px;
        max-width: 75%;
        margin-right: auto;
        border: 1px solid #bae6fd;
    }
    .chat-bubble-teacher {
        background-color: #f0fdf4;
        color: #166534;
        padding: 12px 16px;
        border-radius: 14px 14px 14px 0;
        margin-bottom: 10px;
        max-width: 75%;
        margin-left: auto;
        border: 1px solid #bbf7d0;
    }
    .stButton>button {
        background: #0052cc !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px;
        font-weight: 900 !important;
        font-size: 16px !important;
        padding: 10px 24px;
    }
    .stButton>button:hover { background: #003d99 !important; }
    .call-btn-container {
        display: flex;
        justify-content: center;
        margin-top: 25px;
        margin-bottom: 15px;
        width: 100%;
    }
    .call-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 12px;
        background: linear-gradient(135deg, #059669, #10b981);
        color: #ffffff !important;
        padding: 12px 24px;
        border-radius: 50px;
        font-size: 16px;
        font-weight: 900;
        text-decoration: none !important;
        border: 2px solid #ffffff;
    }
    .social-footer-box {
        margin-top: 20px;
        padding: 20px 0;
        border-top: 1px solid rgba(150, 150, 150, 0.3);
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        width: 100%;
    }
    .social-icons-container {
        display: flex;
        gap: 15px;
        justify-content: center;
        align-items: center;
        margin-bottom: 12px;
    }
    .social-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 44px;
        height: 44px;
        border-radius: 50%;
        text-decoration: none !important;
    }
    .social-btn svg { width: 24px; height: 24px; }
    .facebook-bg { background-color: #1877F2; }
    .whatsapp-bg { background-color: #25D366; }
    .telegram-bg { background-color: #229ED9; }
    .tiktok-bg   { background-color: #000000; border: 1px solid #444; }
    .youtube-bg  { background-color: #FF0000; }
    .rights-text { font-size: 15px; font-weight: 900; margin-top: 10px; text-align: center; }
    @media (prefers-color-scheme: light) {
        body, .stApp { background-color: #ffffff !important; }
        p, span, label, h1, h2, h3, h4, h5, h6, .stMarkdown { color: #111111 !important; font-weight: 800 !important; }
        div[data-testid="stMetric"] { background: #f8fafc !important; border: 2px solid #0052cc !important; border-radius: 12px; padding: 14px 18px; }
        div[data-testid="stMetric"] label { color: #1e293b !important; font-size: 15px !important; font-weight: 800 !important; }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] { color: #0052cc !important; font-weight: 900 !important; }
        input, select, textarea { color: #000000 !important; background-color: #ffffff !important; border: 2px solid #94a3b8 !important; font-weight: 700 !important; }
        .editor-toolbar { background: #f1f5f9; border-color: #cbd5e1; color: #1e293b; }
    }
    @media (prefers-color-scheme: dark) {
        body, .stApp { background-color: #0e1117 !important; }
        p, span, label, h1, h2, h3, h4, h5, h6, .stMarkdown { color: #ffffff !important; font-weight: 700 !important; }
        div[data-testid="stMetric"] { background: #1e232d !important; border: 2px solid #3b82f6 !important; border-radius: 12px; padding: 14px 18px; }
        div[data-testid="stMetric"] label { color: #e2e8f0 !important; font-size: 15px !important; font-weight: 800 !important; }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] { color: #60a5fa !important; font-weight: 900 !important; }
        input, select, textarea { color: #ffffff !important; background-color: #262730 !important; border: 2px solid #475569 !important; font-weight: 700 !important; }
        .editor-toolbar { background: #1e293b; border-color: #475569; color: #f8fafc; }
        .chat-bubble-student { background-color: #075985; color: #f0f9ff; border-color: #0284c7; }
        .chat-bubble-teacher { background-color: #065f46; color: #ecfdf5; border-color: #059669; }
    }
    </style>
""", unsafe_allow_html=True)

query_params = st.query_params
is_student_mode = query_params.get("role") == "student"

# ==============================================================================
# 1. واجهة الطالب (حل الاختبارات الموضوعية والمقالية مع رفع صور الحل)
# ==============================================================================
if is_student_mode:
    st.markdown("""
        <style>
            [data-testid="stSidebar"], [data-testid="stSidebarCollapseButton"], [data-testid="collapsedControl"] { display: none !important; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="brand-banner" dir="rtl">
        <div>
            <h1 class="brand-title">البشمهندس X الرياضة 📐</h1>
            <p class="brand-subtitle">بوابة الطالب الذكية • الحضور، الاختبارات الإلكترونية، والدردشة مع البشمهندس</p>
        </div>
        {'<img src="data:image/jpeg;base64,' + img_b64 + '" style="width: 90px; height: 90px; border-radius: 50%; border: 3px solid #ffffff; object-fit: cover;">' if img_b64 else ''}
    </div>
    """, unsafe_allow_html=True)

    if "logged_student" not in st.session_state:
        st.session_state.logged_student = None

    if not st.session_state.logged_student:
        auth_tab1, auth_tab2 = st.tabs(["🔐 تسجيل دخول الطالب", "✨ إنشاء حساب طالب جديد"])

        with auth_tab1:
            st.subheader("سجل دخولك لمتابعة اختباراتك وحضورك:")
            with st.form("student_login_form"):
                login_name = st.text_input("اسم الطالب المسجل:")
                login_pass = st.text_input("الرقم السري الخاص بك:", type="password")
                btn_login = st.form_submit_button("دخول إلى حسابي")

                if btn_login:
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
                            st.success(f"مرحباً بك مجدداً يا {login_name}!")
                            st.rerun()
                    else:
                        st.error("اسم الطالب أو الرقم السري غير صحيح.")

        with auth_tab2:
            st.subheader("إنشاء حساب لأول مرة:")
            with st.form("student_register_form"):
                reg_name = st.text_input("اسمك بالكامل (ثلاثي أو رباعي):", placeholder="مثال: أحمد محمود علي")
                reg_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()))
                reg_grade = st.selectbox("المرحلة / الصف الدراسي:", CURRICULUM_DATA[reg_curr])
                reg_pass = st.text_input("اختر رقماً سرياً خاصاً بك:", type="password")
                btn_reg = st.form_submit_button("تأكيد إنشاء الحساب")

                if btn_reg:
                    if not reg_name.strip() or not reg_pass.strip():
                        st.error("يرجى كتابة اسمك والرقم السري.")
                    else:
                        existing = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == reg_name.strip()]
                        if not existing.empty:
                            st.warning("هذا الاسم مسجل بالفعل! يرجى تسجيل الدخول مباشرة.")
                        else:
                            new_user = {
                                "اسم الطالب": reg_name.strip(),
                                "كلمة المرور": reg_pass.strip(),
                                "المنهج/الدولة": reg_curr,
                                "المجموعة/الصف": reg_grade,
                                "تاريخ التسجيل": str(date.today()),
                                "الحالة_حظر": "نشط",
                            }
                            st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([new_user])], ignore_index=True)
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                            st.session_state.logged_student = new_user
                            st.success(f"تم إنشاء حسابك بنجاح يا {reg_name}!")
                            st.rerun()

    else:
        st_user = st.session_state.logged_student
        fresh_user = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip()]
        if not fresh_user.empty and fresh_user.iloc[0].get("الحالة_حظر") == "محظور":
            st.error("🚫 عذراً، تم حظر حسابك.")
            st.session_state.logged_student = None
            st.stop()

        col_u1, col_u2 = st.columns([4, 1])
        with col_u1:
            st.markdown(f"""
                <div style="background: rgba(0, 82, 204, 0.08); padding: 12px 18px; border-radius: 10px; border-right: 5px solid #0052cc;">
                    <h3 style="margin: 0; color: #0052cc;">أهلاً بك: {st_user['اسم الطالب']} 🌟</h3>
                    <p style="margin: 3px 0 0 0; font-weight: 800;">{st_user.get('المنهج/الدولة', '')} | {st_user.get('المجموعة/الصف', '')}</p>
                </div>
            """, unsafe_allow_html=True)
        with col_u2:
            st.write("")
            if st.button("🚪 خروج"):
                st.session_state.logged_student = None
                st.rerun()

        # 1. الاختبارات الإلكترونية المتاحة
        st.markdown("<div class='vertical-section-header'>✍️ أولاً: الاختبارات الإلكترونية المتاحة لصفك</div>", unsafe_allow_html=True)
        my_grade = st_user.get("المجموعة/الصف", "")
        available_exams = st.session_state.exams_df[st.session_state.exams_df["المجموعة/الصف"].astype(str).str.strip() == my_grade.strip()].copy()

        if available_exams.empty:
            st.info("لا توجد اختبارات إلكترونية مخصصة لصفك الدراسي حالياً.")
        else:
            for _, ex_row in available_exams.iterrows():
                ex_id = ex_row["معرف_الامتحان"]
                ex_title = ex_row["عنوان الامتحان"]
                ex_desc = ex_row["وصف الامتحان"]
                ex_pass = str(ex_row.get("كلمة المرور", "")).strip()
                ex_time = ex_row.get("مدة الامتحان بالدقائق", 30)

                with st.expander(f"📝 {ex_title} (المدة: {ex_time} دقيقة)"):
                    st.write(f"**الوصف:** {ex_desc}")

                    already_solved = st.session_state.assessments_df[
                        (st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip())
                        & (st.session_state.assessments_df["عنوان التكليف"].astype(str).str.contains(ex_title, na=False))
                    ]

                    if not already_solved.empty:
                        solved_score = already_solved.iloc[-1]["الدرجة المحصلة"]
                        solved_max = already_solved.iloc[-1]["الدرجة العظمى"]
                        st.success(f"✅ لقد قمت بحل هذا الاختبار مسبقاً بنتيجة: ({solved_score} من {solved_max})")
                    else:
                        with st.form(f"take_exam_{ex_id}"):
                            input_pass = ""
                            if ex_pass and ex_pass != "nan":
                                input_pass = st.text_input("🔐 أدخل كلمة مرور الاختبار للبدء:", type="password", key=f"p_{ex_id}")

                            questions = []
                            try:
                                questions = json.loads(ex_row["الأسئلة_JSON"])
                            except Exception:
                                pass

                            answers_mcq = {}
                            essay_texts = {}
                            essay_imgs = {}

                            if questions:
                                st.write("---")
                                for q_idx, q in enumerate(questions):
                                    q_type = q.get("type", "موضوعي")
                                    st.markdown(f"**س {q_idx + 1} ({'اختيار من متعدد' if q_type=='موضوعي' else 'سؤال مقالي'}):** {q['text']} — *(الدرجة: {q['points']})*")

                                    if q_type == "موضوعي":
                                        opts = [q['opt1'], q['opt2'], q['opt3'], q['opt4']]
                                        answers_mcq[q_idx] = st.radio(f"إجابتك للسؤال {q_idx + 1}:", opts, key=f"q_{ex_id}_{q_idx}")
                                    else:
                                        st.info("✍️ يمكنك كتابة خطوات الحل في المربع أو إرفاق صورة لحلك من كشكولك:")
                                        essay_texts[q_idx] = st.text_area(f"اكتب خطوات الحل هنا (س {q_idx + 1}):", key=f"essay_txt_{ex_id}_{q_idx}")
                                        essay_imgs[q_idx] = st.file_uploader(f"📷 إرفاق صورة خطوات الحل (س {q_idx + 1}):", type=["jpg", "png", "jpeg"], key=f"essay_img_{ex_id}_{q_idx}")
                                    st.write("")

                            submit_exam_btn = st.form_submit_button("🏁 إنهاء الاختبار وتسليم الإجابات")

                            if submit_exam_btn:
                                if ex_pass and ex_pass != "nan" and input_pass.strip() != ex_pass:
                                    st.error("كلمة مرور الاختبار غير صحيحة.")
                                elif not questions:
                                    st.error("لا توجد أسئلة داخل هذا الاختبار.")
                                else:
                                    mcq_score = 0.0
                                    total_max = 0.0
                                    has_essay = False

                                    for q_idx, q in enumerate(questions):
                                        q_pts = float(q.get("points", 1.0))
                                        total_max += q_pts

                                        if q.get("type", "موضوعي") == "موضوعي":
                                            correct_opt = q.get(f"opt{q['correct']}")
                                            if answers_mcq.get(q_idx) == correct_opt:
                                                mcq_score += q_pts
                                        else:
                                            has_essay = True
                                            img_str = ""
                                            up_file = essay_imgs.get(q_idx)
                                            if up_file is not None:
                                                img_str = base64.b64encode(up_file.read()).decode()

                                            new_essay_sub = {
                                                "معرف_الحل": f"ANS_{datetime.now().strftime('%Y%m%d%H%M%S')}_{q_idx}",
                                                "معرف_الامتحان": ex_id,
                                                "عنوان الامتحان": ex_title,
                                                "اسم الطالب": st_user["اسم الطالب"],
                                                "رقم السؤال": q_idx + 1,
                                                "نص السؤال": q["text"],
                                                "إجابة الطالب النصية": essay_texts.get(q_idx, ""),
                                                "صورة الحل_base64": img_str,
                                                "درجة السؤال": q_pts,
                                                "الدرجة المرصودة": 0.0,
                                                "حالة التصحيح": "قيد التصحيح من المعلم",
                                                "ملاحظات المعلم": "",
                                                "تاريخ الحل": str(date.today()),
                                            }
                                            st.session_state.essays_df = pd.concat([st.session_state.essays_df, pd.DataFrame([new_essay_sub])], ignore_index=True)

                                    note_msg = f"الأسئلة الموضوعية: {mcq_score} درجة"
                                    if has_essay:
                                        note_msg += " + الأسئلة المقالية قيد المراجعة والتصحيح"

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
                                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                                    st.success(f"🎉 تم تسليم إجاباتك بنجاح! نتيجتك في الاختياري: ({mcq_score} من {total_max}) وسيتم تصحيح المقالي قريباً.")
                                    st.rerun()

        # 2. تسجيل حضور الحصة
        st.markdown("<div class='vertical-section-header'>📝 ثانياً: تسجيل حضور حصة اليوم وتقييمها</div>", unsafe_allow_html=True)
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
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                st.success(f"تم تسجيل حضورك بنجاح للحصة بتاريخ {st_date}!")

        # 3. سجل الدرجات
        st.markdown("<div class='vertical-section-header'>📊 ثالثاً: متابعة سجل الواجبات والاختبارات</div>", unsafe_allow_html=True)
        my_assessments = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"].astype(str).str.strip() == st_user["اسم الطالب"].strip()].copy()
        if my_assessments.empty:
            st.info("لا توجد درجات مرصودة لك حتى الآن.")
        else:
            st.dataframe(my_assessments[["التاريخ", "النوع", "عنوان التكليف", "الدرجة المحصلة", "الدرجة العظمى", "حالة التسليم", "ملاحظات وتوجيهات"]], use_container_width=True)

        # 4. الدردشة
        st.markdown("<div class='vertical-section-header'>💬 رابعاً: الدردشة والتواصل المباشر مع البشمهندس</div>", unsafe_allow_html=True)
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
            uploaded_msg_img = st.file_uploader("📷 إرفاق صورة للمسألة (اختياري):", type=["jpg", "png", "jpeg"])
            if st.form_submit_button("📤 إرسال الرسالة إلى البشمهندس"):
                if msg_text.strip() or uploaded_msg_img is not None:
                    img_str = base64.b64encode(uploaded_msg_img.read()).decode() if uploaded_msg_img is not None else ""
                    new_msg = {
                        "التاريخ_والوقت": datetime.now().strftime("%Y-%m-%d %H:%M"), "اسم الطالب": student_name_key,
                        "المرسل": "student", "نص الرسالة": msg_text.strip(), "الصورة_base64": img_str,
                    }
                    st.session_state.messages_df = pd.concat([st.session_state.messages_df, pd.DataFrame([new_msg])], ignore_index=True)
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                    st.success("تم إرسال رسالتك للبشمهندس بنجاح!")
                    st.rerun()

    st.markdown("""
        <div class="call-btn-container">
            <a href="tel:01016361440" class="call-btn">
                <svg viewBox="0 0 24 24"><path d="M6.62 10.79c1.44 2.83 3.76 5.14 6.59 6.59l2.2-2.2c.27-.27.67-.36 1.02-.24 1.12.37 2.33.57 3.57.57.55 0 1 .45 1 1V20c0 .55-.45 1-1 1-9.39 0-17-7.61-17-17 0-.55.45-1 1-1h3.5c.55 0 1 .45 1 1 0 1.25.2 2.45.57 3.57.11.35.03.74-.25 1.02l-2.2 2.2z"/></svg>
                <span>للتواصل مع م / محمد غنيم: 01016361440</span>
            </a>
        </div>
        <div class="social-footer-box">
            <div class="rights-text">جميع الحقوق محفوظة لدي م / محمد غنيم 2026</div>
        </div>
    """, unsafe_allow_html=True)
    st.stop()

# ==============================================================================
# 2. لوحة تحكم المعلم الرئيسية
# ==============================================================================
st.sidebar.markdown("### 📷 صورة الشعار والمعلم")
uploaded_photo = st.sidebar.file_uploader("ارفع صورتك هنا إذا لم تظهر تلقائياً:", type=["jpg", "png", "jpeg"])
if uploaded_photo is not None:
    with open(IMG_NAME, "wb") as f:
        f.write(uploaded_photo.getbuffer())
    found_img_path = IMG_NAME
    st.sidebar.success("✓ تم حفظ صورتك بنجاح!")
    st.rerun()

if found_img_path and os.path.exists(found_img_path):
    st.sidebar.image(found_img_path, width=220)

st.sidebar.markdown("""
    <div style="text-align: center; margin-top: 5px; margin-bottom: 20px;">
        <h2 style="margin: 0; color: #0052cc; font-weight: 900;">البشمهندس X الرياضة</h2>
        <p style="margin: 4px 0; font-weight: 800; font-size: 15px;">لوحة المعلم المتقدمة</p>
    </div>
""", unsafe_allow_html=True)

st.sidebar.header("🔗 رابط تسجيل الطلاب")
st.sidebar.code("https://engmohamedghonaim.streamlit.app/?role=student", language="text")

st.markdown(f"""
<div class="brand-banner" dir="rtl">
    <div>
        <h1 class="brand-title">البشمهندس X الرياضة 📐</h1>
        <p class="brand-subtitle">نظام صانع الامتحانات التفاعلي • تصحيح المقالي بالصور • إدارة الحصص والمحادثات</p>
    </div>
    {'<img src="data:image/jpeg;base64,' + img_b64 + '" style="width: 90px; height: 90px; border-radius: 50%; border: 3px solid #ffffff; object-fit: cover;">' if img_b64 else ''}
</div>
""", unsafe_allow_html=True)

tab_exam_maker, tab_essay_grade, tab_chat, tab_cards, tab1, tab_hw, tab2, tab3, tab4 = st.tabs([
    "⚙️ صانع الامتحانات والأسئلة",
    "📝 تصحيح الأسئلة المقالية والصور",
    "💬 مركز الدردشة والرسائل",
    "👥 بطاقات الطلاب والتحكم",
    "📝 رصد حصة جديدة",
    "📚 رصد واجب يدوي",
    "✏️ تعديل السجلات",
    "📊 السجلات الشاملة",
    "🖨️ تقرير ولي الأمر للطباعة",
])

# ----------------- تبويب صانع الامتحانات (تطابق كامل لواجهة الصورة) -----------------
with tab_exam_maker:
    st.markdown("<div class='exam-builder-header'>➕ إضافة امتحان جديد / إدخال بنك الأسئلة</div>", unsafe_allow_html=True)

    if "temp_questions" not in st.session_state:
        st.session_state.temp_questions = []

    with st.expander("⚙️ 1. الإعدادات الرئيسية للامتحان", expanded=True):
        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            ex_title_input = st.text_input("عنوان الامتحان:*", placeholder="مثال: امتحان مراجعة الجبر وحساب المثلثات")
            ex_desc_input = st.text_area("وصف الامتحان:*", placeholder="تعليمات هامة للطلاب قبل البدء...")
            ex_pass_input = st.text_input("كلمة المرور (اختياري):", placeholder="اتركه فارغاً إن لم ترغب بكلمة سر")
        with c_ex2:
            ex_curr_input = st.selectbox("اختر الدولة / المنهج:*", list(CURRICULUM_DATA.keys()), key="mk_c")
            ex_grade_input = st.selectbox("اختر الصف الدراسي:*", CURRICULUM_DATA[ex_curr_input], key="mk_g")
            ex_subject_input = st.selectbox("اختر مادة الامتحان:*", ["الرياضيات (عام)", "الجبر والإحصاء", "الهندسة وحساب المثلثات", "التفاضل والتكامل", "الاستاتيكا والديناميكا", "الرياضيات التطبيقية"])
            ex_term_input = st.selectbox("اختر الفصل الدراسي:*", ["الفصل الدراسي الأول", "الفصل الدراسي الثاني", "مراجعة نهائية"])
            ex_time_input = st.number_input("مدة الامتحان (بالدقائق):", min_value=5, max_value=180, value=45, step=5)

    with st.expander("📐 2. محرر وإدخال الأسئلة (موضوعي ومقالي)", expanded=True):
        q_count = len(st.session_state.temp_questions) + 1
        st.markdown(f"#### إضافة سؤال رقم {q_count}")

        # نوع السؤال
        q_type_choice = st.selectbox("الرجاء اختيار نوع السؤال:", ["سؤال موضوعي (اختيار من متعدد)", "سؤال مقالي (خطوات حل وكتابة / رفع صورة)"])
        is_mcq = "موضوعي" in q_type_choice

        # شريط أدوات الرموز الرياضية
        st.markdown("""
            <div class="editor-toolbar">
                <span>➕ شريط الرموز الرياضية السريعة:</span>
                <b>B</b> | <i>I</i> | <u>U</u> | 
                <code>x²</code> <code>√x</code> <code>½</code> <code>π</code> <code>θ</code> <code>Δ</code> <code>≤</code> <code>≥</code> <code>≠</code> <code>∞</code>
            </div>
        """, unsafe_allow_html=True)

        with st.form(f"add_question_form_{q_count}", clear_on_submit=True):
            st.markdown("**نص السؤال الرياضي:***")
            q_text = st.text_area("اكتب نص السؤال هنا (يدعم LaTeX مثل $x^2 + 5 = 9$):", label_visibility="collapsed")

            if is_mcq:
                st.markdown("---")
                st.markdown("**خيارات الإجابة الأربعة:**")

                col_o1, col_o2 = st.columns([4, 1])
                with col_o1:
                    opt1 = st.text_input("الإجابة الأولى (أ):")
                with col_o2:
                    st.write("")
                    c1 = st.checkbox("الإجابة الصحيحة", key="chk_1")

                col_o3, col_o4 = st.columns([4, 1])
                with col_o3:
                    opt2 = st.text_input("الإجابة الثانية (ب):")
                with col_o4:
                    st.write("")
                    c2 = st.checkbox("الإجابة الصحيحة", key="chk_2")

                col_o5, col_o6 = st.columns([4, 1])
                with col_o5:
                    opt3 = st.text_input("الإجابة الثالثة (ج):")
                with col_o6:
                    st.write("")
                    c3 = st.checkbox("الإجابة الصحيحة", key="chk_3")

                col_o7, col_o8 = st.columns([4, 1])
                with col_o7:
                    opt4 = st.text_input("الإجابة الرابعة (د):")
                with col_o8:
                    st.write("")
                    c4 = st.checkbox("الإجابة الصحيحة", key="chk_4")
            else:
                opt1, opt2, opt3, opt4 = "", "", "", ""
                c1, c2, c3, c4 = False, False, False, False
                st.info("💡 في السؤال المقالي، ستظهر للطالب مساحة لكتابة خطوات الحل أو زر لرفع صورة كشكوله لحلها، لتصححها أنت يدوياً.")

            st.write("---")
            q_points = st.selectbox("علامة هذا السؤال (الدرجة):*", [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0, 10.0], index=1)
            add_q_submit = st.form_submit_button("إضافة هذا السؤال")

            if add_q_submit:
                if not q_text.strip():
                    st.error("يرجى كتابة نص السؤال أولاً.")
                elif is_mcq and (not opt1.strip() or not opt2.strip()):
                    st.error("يرجى كتابة خيارين على الأقل للسؤال الموضوعي.")
                else:
                    correct_num = 1
                    if c2: correct_num = 2
                    elif c3: correct_num = 3
                    elif c4: correct_num = 4

                    st.session_state.temp_questions.append({
                        "type": "موضوعي" if is_mcq else "مقالي",
                        "text": q_text.strip(),
                        "opt1": opt1.strip(), "opt2": opt2.strip(),
                        "opt3": opt3.strip(), "opt4": opt4.strip(),
                        "correct": correct_num,
                        "points": q_points,
                    })
                    st.success(f"✓ تم إضافة السؤال رقم {q_count} بنجاح!")
                    st.rerun()

        if st.session_state.temp_questions:
            st.write(f"**الأسئلة المجهزة حالياً: ({len(st.session_state.temp_questions)}) سؤال**")
            for idx_q, q_item in enumerate(st.session_state.temp_questions):
                st.markdown(f"- **س {idx_q+1} ({q_item['type']}):** {q_item['text']} — *(الدرجة: {q_item['points']})*")
            if st.button("🗑️ مسح الأسئلة والبدء من جديد"):
                st.session_state.temp_questions = []
                st.rerun()

    st.write("---")
    if st.button("🚀 حفظ ونشر الامتحان للطلاب"):
        if not ex_title_input.strip():
            st.error("يرجى كتابة عنوان الامتحان أولاً.")
        elif not st.session_state.temp_questions:
            st.error("يرجى إضافة سؤال واحد على الأقل قبل الحفظ.")
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
                "مدة الامتحان بالدقائق": ex_time_input,
                "الأسئلة_JSON": json.dumps(st.session_state.temp_questions, ensure_ascii=False),
                "تاريخ الإنشاء": str(date.today()),
            }
            st.session_state.exams_df = pd.concat([st.session_state.exams_df, pd.DataFrame([new_ex])], ignore_index=True)
            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
            st.session_state.temp_questions = []
            st.success(f"✓ تم نشر امتحان ({ex_title_input}) بنجاح لطلاب صف ({ex_grade_input})!")
            st.rerun()

    # قائمة الامتحانات المنشورة
    st.write("---")
    st.markdown("### 📋 الامتحانات المنشورة مسبقاً:")
    if st.session_state.exams_df.empty:
        st.info("لا توجد امتحانات منشورة بعد.")
    else:
        for ex_i, ex_r in st.session_state.exams_df.iterrows():
            c_e1, c_e2 = st.columns([5, 1])
            with c_e1:
                st.markdown(f"**{ex_r['عنوان الامتحان']}** — الصف: {ex_r['المجموعة/الصف']} | المدة: {ex_r['مدة الامتحان بالدقائق']} دقيقة")
            with c_e2:
                if st.button("حذف 🗑️", key=f"del_ex_btn_{ex_i}"):
                    st.session_state.exams_df = st.session_state.exams_df.drop(ex_i).reset_index(drop=True)
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                    st.warning("تم حذف الامتحان.")
                    st.rerun()

# ----------------- تبويب تصحيح الأسئلة المقالية والصور -----------------
with tab_essay_grade:
    st.subheader("📝 كنترول تصحيح إجابات الطلاب المقالية وصور الحل:")
    essays = st.session_state.essays_df

    if essays.empty:
        st.info("لا توجد إجابات مقالية مرسلة من الطلاب بعد.")
    else:
        for es_idx, es_row in essays.iterrows():
            st_name = es_row["اسم الطالب"]
            ex_name = es_row["عنوان الامتحان"]
            q_num = es_row["رقم السؤال"]
            q_text = es_row["نص السؤال"]
            ans_txt = es_row.get("إجابة الطالب النصية", "")
            img_b64_ans = es_row.get("صورة الحل_base64", "")
            q_max = float(es_row.get("درجة السؤال", 1.0))
            status = es_row.get("حالة التصحيح", "قيد التصحيح من المعلم")

            with st.expander(f"📌 حل الطالب: {st_name} — {ex_name} (س {q_num}) | [{status}]"):
                st.markdown(f"**نص السؤال:** {q_text}")
                st.markdown(f"**إجابة الطالب المكتوبة:**")
                st.write(ans_txt if ans_txt.strip() else "لم يكتب نصاً (أرفق صورة بالأسفل)")

                if pd.notnull(img_b64_ans) and str(img_b64_ans).strip():
                    st.markdown("**📷 صورة خطوات الحل المرفوعة من الطالب:**")
                    st.image(f"data:image/jpeg;base64,{img_b64_ans}", width=500)

                with st.form(f"grade_essay_form_{es_idx}"):
                    c_g1, c_g2 = st.columns(2)
                    with c_g1:
                        awarded_score = st.number_input("الدرجة المستحقة:*", min_value=0.0, max_value=q_max, value=float(es_row.get("الدرجة المرصودة", 0.0)), step=0.5)
                    with c_g2:
                        teacher_feedback = st.text_input("ملاحظات المعلم وتوجيهه للطالب:", value=str(es_row.get("ملاحظات المعلم", "")))

                    if st.form_submit_button("💾 اعتماد ورصد الدرجة للطالب"):
                        essays.at[es_idx, "الدرجة المرصودة"] = awarded_score
                        essays.at[es_idx, "حالة التصحيح"] = "تم التصحيح والاعتماد"
                        essays.at[es_idx, "ملاحظات المعلم"] = teacher_feedback.strip()

                        # تحديث درجة الطالب الإجمالية في Assessments
                        match_ass = st.session_state.assessments_df[
                            (st.session_state.assessments_df["اسم الطالب"] == st_name)
                            & (st.session_state.assessments_df["عنوان التكليف"] == ex_name)
                        ]
                        if not match_ass.empty:
                            ass_idx = match_ass.index[-1]
                            current_score = float(st.session_state.assessments_df.at[ass_idx, "الدرجة المحصلة"])
                            st.session_state.assessments_df.at[ass_idx, "الدرجة المحصلة"] = current_score + awarded_score
                            st.session_state.assessments_df.at[ass_idx, "ملاحظات وتوجيهات"] = f"تم تصحيح المقالي: +{awarded_score} درجة. {teacher_feedback}"

                        st.session_state.essays_df = essays
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                        st.success(f"✓ تم رصد درجة الطالب ({st_name}) بنجاح وإضافتها لشهادة درجاته!")
                        st.rerun()

# ----------------- تبويب الدردشة -----------------
with tab_chat:
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
                        save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                        st.success("تم إرسال الرد للطالب فوراً!")
                        st.rerun()

# ----------------- تبويب بطاقات الطلاب -----------------
with tab_cards:
    st.subheader("👥 بطاقات الطلاب المسجلين والتحكم الكامل:")
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
        st.info("لا يوجد طلاب مسجلون حالياً.")
    else:
        for idx, st_name in enumerate(all_known_students):
            u_row = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].astype(str).str.strip() == st_name.strip()]
            s_row = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"].astype(str).str.strip() == st_name.strip()]

            st_curr = u_row.iloc[0].get("المنهج/الدولة", "-") if not u_row.empty else (s_row.iloc[-1].get("المنهج/الدولة", "-") if not s_row.empty else "-")
            st_grade = u_row.iloc[0].get("المجموعة/الصف", "-") if not u_row.empty else (s_row.iloc[-1].get("المجموعة/الصف", "-") if not s_row.empty else "-")
            st_date = u_row.iloc[0].get("تاريخ التسجيل", "-") if not u_row.empty else "-"
            st_pass = u_row.iloc[0].get("كلمة المرور", "-") if not u_row.empty else "-"
            is_banned = u_row.iloc[0].get("الحالة_حظر") == "محظور" if not u_row.empty else False

            status_badge = "🚫 محظور" if is_banned else "✅ نشط"
            badge_color = "#dc2626" if is_banned else "#16a34a"

            with st.container():
                col_c1, col_c2, col_c3, col_c4 = st.columns([4, 2, 2, 2])
                with col_c1:
                    st.markdown(f"""
                        <h4 style="margin: 0; color: #0052cc;">{st_name}</h4>
                        <p style="margin: 3px 0; font-size: 14px; font-weight: 800;">{st_curr} — {st_grade}</p>
                        <p style="margin: 0; font-size: 13px; color: #64748b;">تاريخ التسجيل: {st_date} | كلمة المرور: <b>{st_pass}</b></p>
                    """, unsafe_allow_html=True)
                with col_c2:
                    st.markdown(f"<span style='color: {badge_color}; font-weight: 900; font-size: 16px;'>{status_badge}</span>", unsafe_allow_html=True)
                with col_c3:
                    if is_banned:
                        if st.button("فك الحظر 🔓", key=f"unban_{idx}"):
                            st.session_state.users_df.loc[st.session_state.users_df["اسم الطالب"] == st_name, "الحالة_حظر"] = "نشط"
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                            st.success(f"تم فك حظر {st_name}!")
                            st.rerun()
                    else:
                        if st.button("حظر 🚫", key=f"ban_{idx}"):
                            if not u_row.empty:
                                st.session_state.users_df.loc[st.session_state.users_df["اسم الطالب"] == st_name, "الحالة_حظر"] = "محظور"
                                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                                st.warning(f"تم حظر {st_name}.")
                                st.rerun()
                with col_c4:
                    if st.button("حذف نهائي 🗑️", key=f"del_card_btn_{idx}"):
                        delete_student_completely(st_name)
                        st.success(f"✓ تم حذف الطالب ({st_name}) نهائياً!")
                        st.rerun()
                st.write("---")

# ----------------- تبويب رصد الحصة -----------------
with tab1:
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
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                st.success(f"✓ تم حفظ سجل الحصة للطالب ({student_name}) بنجاح!")

# ----------------- تبويب رصد الواجبات يدوياً -----------------
with tab_hw:
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
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                st.success(f"✓ تم رصد {ass_type} بنجاح للطالب ({ass_student})!")

    st.write("---")
    st.dataframe(st.session_state.assessments_df, use_container_width=True)

# ----------------- تبويب تعديل السجلات -----------------
with tab2:
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
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                st.success("✓ تم تحديث بيانات الحصة بنجاح!")
                st.rerun()

            if delete_btn:
                df = df.drop(selected_idx).reset_index(drop=True)
                st.session_state.sessions_df = df
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df, st.session_state.messages_df, st.session_state.exams_df, st.session_state.essays_df)
                st.warning("⚠️ تم حذف السجل.")
                st.rerun()

# ----------------- تبويب السجلات العامة -----------------
with tab3:
    st.subheader("نظرة شاملة على السجلات والإحصائيات")
    current_df = st.session_state.sessions_df

    if current_df.empty:
        st.info("لا توجد بيانات حصص متاحة.")
    else:
        c_f1, c_f2, c_f3 = st.columns(3)
        with c_f1:
            all_currs = ["الكل"] + [c for c in current_df["المنهج/الدولة"].dropna().unique() if str(c).strip()]
            filter_curr = st.selectbox("تصفية حسب المنهج / الدولة:", all_currs)
        with c_f2:
            all_groups = ["الكل"] + [g for g in current_df["المجموعة/الصف"].dropna().unique() if str(g).strip()]
            filter_group = st.selectbox("تصفية حسب المرحلة / الصف:", all_groups)
        with c_f3:
            search_name = st.text_input("بحث باسم الطالب:")

        filtered = current_df.copy()
        if filter_curr != "الكل": filtered = filtered[filtered["المنهج/الدولة"] == filter_curr]
        if filter_group != "الكل": filtered = filtered[filtered["المجموعة/الصف"] == filter_group]
        if search_name.strip(): filtered = filtered[filtered["اسم الطالب"].str.contains(search_name.strip(), na=False)]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("إجمالي الحصص", len(filtered))
        m2.metric("عدد مرات الحضور", len(filtered[filtered["الحالة"] == "حاضر"]))
        m3.metric("عدد مرات الغياب", len(filtered[filtered["الحالة"] == "غائب"]))
        total_cash = filtered["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True)
        m4.metric("إجمالي المبالغ المستحقة", f"{total_cash:,.1f}")

        st.dataframe(filtered, use_container_width=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            st.session_state.users_df.to_excel(writer, sheet_name="Users", index=False)
            st.session_state.sessions_df.to_excel(writer, sheet_name="Sessions", index=False)
            st.session_state.assessments_df.to_excel(writer, sheet_name="Assessments", index=False)
            st.session_state.messages_df.to_excel(writer, sheet_name="Messages", index=False)
            st.session_state.exams_df.to_excel(writer, sheet_name="Exams", index=False)
            st.session_state.essays_df.to_excel(writer, sheet_name="Essays", index=False)

        st.download_button(
            label="📥 تصدير ملف Excel الشامل (مستخدمين + حصص + واجبات + امتحانات + مقالي)",
            data=buf.getvalue(), file_name=FILE_NAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ----------------- تبويب تقرير ولي الأمر للطباعة -----------------
with tab4:
    st.subheader("📑 إصدار وطباعة تقرير متابعة الطالب لولي الأمر (شامل الأسعار)")
    all_names = sorted(list(set(
        [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.assessments_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    if not all_names:
        st.info("لا توجد بيانات كافية لإصدار التقرير بعد.")
    else:
        selected_student = st.selectbox("اختر الطالب لإصدار تقرير ولي الأمر:", all_names)

        if selected_student:
            st_sessions = st.session_state.sessions_df[st.session_state.sessions_df["اسم الطالب"] == selected_student].copy().sort_values(by="التاريخ")
            st_assessments = st.session_state.assessments_df[st.session_state.assessments_df["اسم الطالب"] == selected_student].copy().sort_values(by="التاريخ")

            curr_val = st_sessions.iloc[-1].get("المنهج/الدولة", "-") if not st_sessions.empty else "-"
            group_val = st_sessions.iloc[-1].get("المجموعة/الصف", "-") if not st_sessions.empty else "-"
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
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 800;">{r['التاريخ']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #0052cc;">{r['النوع']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 800;">{r['عنوان التكليف']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900;">{r['الدرجة المحصلة']} / {r['الدرجة العظمى']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 800;">{r['حالة التسليم']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 800; color: #047857;">{r['ملاحظات وتوجيهات']}</td>
                    </tr>
                    """
            else:
                ass_html_rows = "<tr><td colspan='6' style='padding: 15px; font-weight: 800; border: 2px solid #000;'>لا توجد واجبات أو اختبارات مرصودة حتى الآن.</td></tr>"

            session_html_rows = ""
            if not st_sessions.empty:
                for _, r in st_sessions.iterrows():
                    st_color = "#0f766e" if r["الحالة"] == "حاضر" else ("#b91c1c" if r["الحالة"] == "غائب" else "#b45309")
                    p_curr = float(r.get('سعر الحصة', 0)) if pd.notnull(r.get('سعر الحصة')) else 0.0
                    session_html_rows += f"""
                    <tr>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 800;">{r['التاريخ']}</td>
                        <td style="padding: 10px; border: 2px solid #000; color:{st_color}; font-weight: 900;">{r['الحالة']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #16a34a; font-size: 16px;">{p_curr:,.1f}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 900; color: #0052cc;">{r['مستوى الطالب']}</td>
                        <td style="padding: 10px; border: 2px solid #000; font-weight: 800;">{r['ملاحظات']}</td>
                    </tr>
                    """
                session_html_rows += f"""
                <tr style="background-color: #f1f5f9;">
                    <td colspan="2" style="padding: 12px; border: 2px solid #000; font-weight: 900; font-size: 16px;">الإجمالي الكلي المستحق للحصص:</td>
                    <td style="padding: 12px; border: 2px solid #000; font-weight: 900; font-size: 18px; color: #b91c1c;">{total_cost:,.1f}</td>
                    <td colspan="2" style="padding: 12px; border: 2px solid #000; font-weight: 800; font-size: 15px;">نظام الدفع المعتمد: {pay_val}</td>
                </tr>
                """
            else:
                session_html_rows = "<tr><td colspan='5' style='padding: 15px; font-weight: 800; border: 2px solid #000;'>لا توجد حصص مسجلة بعد.</td></tr>"

            teacher_img_tag = f'<img src="data:image/jpeg;base64,{img_b64}" style="width: 100px; height: 100px; border-radius: 50%; border: 3px solid #0052cc; object-fit: cover;">' if img_b64 else ""

            parent_report_html = f"""<!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head>
                <meta charset="utf-8">
                <title>تقرير متابعة ولي الأمر - {selected_student}</title>
                <style>
                    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@700;800;900&display=swap');
                    body {{ font-family: 'Cairo', Tahoma, Arial, sans-serif; padding: 30px; color: #000000 !important; background-color: #ffffff !important; font-weight: 800; }}
                    .header-box {{ border-bottom: 3px solid #0052cc; padding-bottom: 15px; margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; }}
                    .brand-name {{ color: #0052cc; margin: 0; font-size: 30px; font-weight: 900; }}
                    .brand-sub {{ color: #000000; margin: 4px 0; font-size: 15px; font-weight: 800; }}
                    .parent-notice {{ background-color: #f0fdf4; border: 2px solid #16a34a; padding: 12px 18px; border-radius: 8px; margin-bottom: 25px; font-size: 16px; font-weight: 800; color: #14532d; }}
                    .stats-box {{ width: 100%; border-collapse: collapse; margin-bottom: 25px; }}
                    .stats-box th, .stats-box td {{ border: 2px solid #000000; padding: 10px; text-align: center; font-size: 15px; font-weight: 800; }}
                    .stats-box th {{ background: #f1f5f9; color: #000000; font-weight: 900; }}
                    .section-title {{ margin-top: 25px; margin-bottom: 10px; color: #0052cc; font-size: 19px; font-weight: 900; border-bottom: 2px solid #0052cc; padding-bottom: 6px; display: inline-block; }}
                    .table-main {{ width: 100%; border-collapse: collapse; text-align: center; margin-top: 10px; margin-bottom: 20px; }}
                    .table-main th {{ background-color: #f8fafc; color: #000000; font-weight: 900; font-size: 15px; padding: 10px; border: 2px solid #000000; }}
                    .footer-note {{ margin-top: 35px; text-align: center; color: #000000; font-size: 15px; font-weight: 900; border-top: 2px solid #000000; padding-top: 15px; }}
                </style>
            </head>
            <body onload="window.print()">
                <div class="header-box">
                    <div style="display: flex; align-items: center; gap: 20px;">
                        {teacher_img_tag}
                        <div>
                            <h2 class="brand-name">البشمهندس X الرياضة 📐</h2>
                            <p class="brand-sub">تقرير التقييم الدوري والحساب المالي الشامل لولي الأمر</p>
                            <p style="margin: 2px 0; color: #000000; font-size: 15px; font-weight: 800;"><b>المنهج والمرحلة:</b> {curr_val} — {group_val}</p>
                        </div>
                    </div>
                    <div style="text-align: left;">
                        <h2 style="color: #0052cc; margin: 0; font-size: 24px; font-weight: 900;">الطالب: {selected_student}</h2>
                        <p style="margin: 5px 0 0 0; color: #000000; font-size: 14px; font-weight: 800;">تاريخ إصدار التقرير: {date.today()}</p>
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
                        <td style="font-weight: 900; font-size: 17px; color: #b91c1c;">{total_cost:,.1f}</td>
                        <td style="color: #0052cc; font-weight: 900;">{level_val}</td>
                    </tr>
                </table>
                <div class="section-title">1. تقرير الواجبات المنزلية والاختبارات الدورية:</div>
                <table class="table-main">
                    <tr><th>التاريخ</th><th>النوع</th><th>عنوان التكليف / الاختبار</th><th>الدرجة المحصلة</th><th>حالة التسليم والالتزام</th><th>ملاحظات المعلم وتوجيهاته</th></tr>
                    {ass_html_rows}
                </table>
                <div class="section-title">2. سجل الحضور وتفاصيل سعر كل حصة:</div>
                <table class="table-main">
                    <tr><th>التاريخ</th><th>حالة الحضور</th><th>سعر الحصة</th><th>مستوى الطالب بالحصة</th><th>ملاحظات التفاعل والاستيعاب</th></tr>
                    {session_html_rows}
                </table>
                <div class="footer-note">مع تحيات: <b>البشمهندس X الرياضة | م / محمد غنيم</b> — رقم التواصل المباشر: 01016361440 🌟</div>
            </body>
            </html>"""

            st.download_button(
                label=f"🖨️ تحميل وطباعة تقرير ولي الأمر لـ ({selected_student}) بصيغة PDF",
                data=parent_report_html.encode("utf-8"),
                file_name=f"تقرير_ولي_الأمر_{selected_student}.html",
                mime="text/html",
            )
