import base64
import io
import os
from datetime import date, datetime
import pandas as pd
from PIL import Image
import streamlit as st

FILE_NAME = "سجل_الغياب_والحصص.xlsx"
IMG_NAME = "teacher.jpg"

st.set_page_config(
    page_title="البشمهندس X الرياضة | متابعة الطلاب",
    page_icon="📐",
    layout="wide",
)

CURRICULUM_DATA = {
    "المنهج المصري 🇪🇬": [
        "الصف الأول الإعدادي",
        "الصف الثاني الإعدادي",
        "الصف الثالث الإعدادي (الشهادة الإعدادية)",
        "الصف الأول الثانوي",
        "الصف الثاني الثانوي (علمي)",
        "الصف الثاني الثانوي (أدبي)",
        "الصف الثالث الثانوي (علمي رياضة)",
        "الصف الثالث الثانوي (علمي علوم)",
        "الصف الثالث الثانوي (أدبي)",
        "المرحلة الابتدائية",
    ],
    "المنهج القطري 🇶🇦": [
        "الصف السابع (إعدادي)",
        "الصف الثامن (إعدادي)",
        "الصف التاسع (إعدادي)",
        "الصف العاشر (المشترك)",
        "الصف الحادي عشر (المسار العلمي)",
        "الصف الحادي عشر (مسار الآداب والإنسانيات)",
        "الصف الحادي عشر (المسار التكنولوجي)",
        "الصف الثاني عشر (المسار العلمي - متقدم)",
        "الصف الثاني عشر (مسار الآداب - تأسيسي)",
        "الصف الثاني عشر (المسار التكنولوجي)",
        "المرحلة الابتدائية",
    ],
    "المنهج الإماراتي 🇦🇪": [
        "الحلقة الثانية (الصفوف 5 - 8)",
        "الصف التاسع (مسار عام)",
        "الصف التاسع (مسار متقدم)",
        "الصف العاشر (مسار عام)",
        "الصف العاشر (مسار متقدم)",
        "الصف العاشر (مسار النخبة)",
        "الصف الحادي عشر (مسار عام)",
        "الصف الحادي عشر (مسار متقدم)",
        "الصف الحادي عشر (مسار النخبة)",
        "الصف الثاني عشر (مسار عام)",
        "الصف الثاني عشر (مسار متقدم)",
        "الصف الثاني عشر (مسار النخبة)",
    ],
    "المنهج السعودي 🇸🇦": [
        "المرحلة المتوسطة (أول / ثاني / ثالث متوسط)",
        "السنة الأولى المشتركة (أول ثانوي)",
        "السنة الثانية (المسار العام)",
        "السنة الثانية (مسار علوم الحاسب والهندسة)",
        "السنة الثانية (مسار الصحة والحياة)",
        "السنة الثانية (مسار إدارة الأعمال / الشرعي)",
        "السنة الثالثة (المسار العام)",
        "السنة الثالثة (مسار علوم الحاسب والهندسة)",
        "السنة الثالثة (مسار الصحة والحياة)",
    ],
    "المنهج الكويتي 🇰🇼": [
        "المرحلة المتوسطة (الصفوف 6 - 9)",
        "الصف العاشر الثانوي (مشترك)",
        "الصف الحادي عشر (القسم العلمي)",
        "الصف الحادي عشر (القسم الأدبي)",
        "الصف الثاني عشر (القسم العلمي)",
        "الصف الثاني عشر (القسم الأدبي)",
    ],
    "المنهج السوداني 🇸🇩": [
        "المرحلة المتوسطة (أولى / ثانية / ثالثة متوسط)",
        "الصف الأول الثانوي",
        "الصف الثاني الثانوي (علمي)",
        "الصف الثاني الثانوي (أدبي)",
        "الصف الثالث الثانوي (علمي رياضيات - الشهادة السودانية)",
        "الصف الثالث الثانوي (علمي أحياء)",
        "الصف الثالث الثانوي (أدبي)",
    ],
}

possible_images = [
    "teacher.jpg",
    "teacher.png",
    "teacher.jpeg",
    "photo_2026-08-02_00-34-53.jpg",
]
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

# ==================== إدارة قواعد البيانات (Excel Sheets) ====================
COL_SESSIONS = [
    "التاريخ",
    "اسم الطالب",
    "المنهج/الدولة",
    "المجموعة/الصف",
    "الحالة",
    "سعر الحصة",
    "عدد الحصص الكلي",
    "نظام الدفع",
    "مستوى الطالب",
    "ملاحظات",
]

COL_USERS = ["اسم الطالب", "كلمة المرور", "المنهج/الدولة", "المجموعة/الصف", "تاريخ التسجيل"]

COL_ASSESSMENTS = [
    "التاريخ",
    "اسم الطالب",
    "النوع",
    "عنوان التكليف",
    "الدرجة المحصلة",
    "الدرجة العظمى",
    "حالة التسليم",
    "ملاحظات وتوجيهات",
]


def load_all_data():
    users_df = pd.DataFrame(columns=COL_USERS)
    sessions_df = pd.DataFrame(columns=COL_SESSIONS)
    assessments_df = pd.DataFrame(columns=COL_ASSESSMENTS)

    if os.path.exists(FILE_NAME):
        try:
            with pd.ExcelFile(FILE_NAME) as xls:
                if "Users" in xls.sheet_names:
                    users_df = pd.read_excel(xls, "Users")
                if "Sessions" in xls.sheet_names:
                    sessions_df = pd.read_excel(xls, "Sessions")
                elif "Sheet1" in xls.sheet_names:
                    sessions_df = pd.read_excel(xls, "Sheet1")
                if "Assessments" in xls.sheet_names:
                    assessments_df = pd.read_excel(xls, "Assessments")
        except Exception:
            pass

    for col in COL_USERS:
        if col not in users_df.columns:
            users_df[col] = ""
    for col in COL_SESSIONS:
        if col not in sessions_df.columns:
            sessions_df[col] = ""
    for col in COL_ASSESSMENTS:
        if col not in assessments_df.columns:
            assessments_df[col] = ""

    return users_df, sessions_df, assessments_df


def save_all_data(users_df, sessions_df, assessments_df):
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
        users_df.to_excel(writer, sheet_name="Users", index=False)
        sessions_df.to_excel(writer, sheet_name="Sessions", index=False)
        assessments_df.to_excel(writer, sheet_name="Assessments", index=False)


if "users_df" not in st.session_state:
    u_df, s_df, a_df = load_all_data()
    st.session_state.users_df = u_df
    st.session_state.sessions_df = s_df
    st.session_state.assessments_df = a_df

# ==================== التنسيق المتجاوب مع الوضع الفاتح والداكن ====================
st.markdown(
    """
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
        -webkit-text-fill-color: #ffffff !important;
        margin: 0 !important;
        text-shadow: 0 2px 6px rgba(0, 0, 0, 0.5);
    }

    .brand-subtitle {
        color: #e0edff !important;
        -webkit-text-fill-color: #e0edff !important;
        font-size: 16px !important;
        font-weight: 700 !important;
        margin-top: 6px !important;
    }

    .darssly-banner {
        background: linear-gradient(135deg, #4f46e5, #7c3aed);
        border-radius: 16px;
        padding: 22px 20px;
        margin-bottom: 25px;
        text-align: center;
        box-shadow: 0 8px 20px rgba(99, 102, 241, 0.28);
        border: 2px solid rgba(255, 255, 255, 0.2);
    }
    .darssly-title {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-size: 22px !important;
        font-weight: 900 !important;
        margin: 0 0 6px 0 !important;
        text-align: center !important;
        text-shadow: 0 2px 4px rgba(0,0,0,0.3);
    }
    .darssly-sub {
        color: #e0e7ff !important;
        -webkit-text-fill-color: #e0e7ff !important;
        font-size: 15px !important;
        font-weight: 700 !important;
        margin-bottom: 16px !important;
        text-align: center !important;
    }
    .darssly-buttons {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        justify-content: center;
        align-items: center;
    }
    .darssly-btn {
        background-color: #ffffff;
        color: #4338ca !important;
        -webkit-text-fill-color: #4338ca !important;
        padding: 10px 20px;
        border-radius: 10px;
        font-size: 15px;
        font-weight: 800;
        text-decoration: none !important;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.15);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        display: inline-flex;
        align-items: center;
        gap: 8px;
    }
    .darssly-btn:hover {
        transform: translateY(-3px) scale(1.03);
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.25);
        background-color: #f8fafc;
    }

    .stButton>button {
        background: #0052cc !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px;
        font-weight: 900 !important;
        font-size: 16px !important;
        padding: 10px 24px;
        box-shadow: 0 4px 10px rgba(0, 82, 204, 0.25);
    }
    .stButton>button:hover {
        background: #003d99 !important;
    }

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
        -webkit-text-fill-color: #ffffff !important;
        padding: 14px 28px;
        border-radius: 50px;
        font-size: 18px;
        font-weight: 900;
        text-decoration: none !important;
        box-shadow: 0 6px 16px rgba(16, 185, 129, 0.35);
        transition: all 0.25s ease;
        border: 2px solid #ffffff;
    }
    .call-btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 22px rgba(16, 185, 129, 0.45);
    }
    .call-btn svg {
        width: 24px;
        height: 24px;
        fill: #ffffff;
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
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        transition: transform 0.2s ease;
    }
    .social-btn:hover {
        transform: scale(1.12);
    }
    .social-btn svg {
        width: 24px;
        height: 24px;
    }
    .facebook-bg { background-color: #1877F2; }
    .whatsapp-bg { background-color: #25D366; }
    .telegram-bg { background-color: #229ED9; }
    .tiktok-bg   { background-color: #000000; border: 1px solid #444; }
    .youtube-bg  { background-color: #FF0000; }

    .rights-text {
        font-size: 15px;
        font-weight: 900;
        margin-top: 10px;
        text-align: center !important;
        direction: rtl;
    }

    @media (prefers-color-scheme: light) {
        body, .stApp { background-color: #ffffff !important; }
        p, span, label, h1, h2, h3, h4, h5, h6, .stMarkdown { color: #111111 !important; font-weight: 800 !important; }
        div[data-testid="stMetric"] { background: #f8fafc !important; border: 2px solid #0052cc !important; border-radius: 12px; padding: 14px 18px; }
        div[data-testid="stMetric"] label { color: #1e293b !important; font-size: 15px !important; font-weight: 800 !important; }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] { color: #0052cc !important; font-weight: 900 !important; }
        input, select, textarea { color: #000000 !important; background-color: #ffffff !important; border: 2px solid #94a3b8 !important; font-weight: 700 !important; }
        .rights-text { color: #1e293b !important; }
    }

    @media (prefers-color-scheme: dark) {
        body, .stApp { background-color: #0e1117 !important; }
        p, span, label, h1, h2, h3, h4, h5, h6, .stMarkdown { color: #ffffff !important; font-weight: 700 !important; }
        div[data-testid="stMetric"] { background: #1e232d !important; border: 2px solid #3b82f6 !important; border-radius: 12px; padding: 14px 18px; }
        div[data-testid="stMetric"] label { color: #e2e8f0 !important; font-size: 15px !important; font-weight: 800 !important; }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] { color: #60a5fa !important; font-weight: 900 !important; }
        input, select, textarea { color: #ffffff !important; background-color: #262730 !important; border: 2px solid #475569 !important; font-weight: 700 !important; }
        .rights-text { color: #e2e8f0 !important; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

query_params = st.query_params
is_student_mode = query_params.get("role") == "student"

# ==============================================================================
# 1. واجهة الطالب
# ==============================================================================
if is_student_mode:
    st.markdown(
        """
        <style>
            [data-testid="stSidebar"] { display: none !important; }
            [data-testid="stSidebarCollapseButton"] { display: none !important; }
            [data-testid="collapsedControl"] { display: none !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
    <div class="brand-banner" dir="rtl">
        <div>
            <h1 class="brand-title">البشمهندس X الرياضة 📐</h1>
            <p class="brand-subtitle">بوابة الطالب الذكية • الحضور والتقييم والواجبات والاختبارات</p>
        </div>
        {'<img src="data:image/jpeg;base64,' + img_b64 + '" style="width: 90px; height: 90px; border-radius: 50%; border: 3px solid #ffffff; object-fit: cover;">' if img_b64 else ''}
    </div>
    """,
        unsafe_allow_html=True,
    )

    # بنر درسلي
    st.markdown(
        """
        <div class="darssly-banner" dir="rtl">
            <h3 class="darssly-title">📢 اشترك الآن في كورسات الرياضيات على منصة درسلي (Darssly)</h3>
            <p class="darssly-sub">اختر مرحلتك للاطلاع على الشرح والخطط الكاملة:</p>
            <div class="darssly-buttons">
                <a href="https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim-3/plans?fbclid=IwY2xjawUKZOBwZG9mAWV4dG4DYWVtAjEwAGJyaWQRMTlWRlpNM3FsN3ViUTA1blBzcnRjBmFwcF9pZBAyMjIwMzkxNzg4MjAwODkyAAEe9bKTqqqGR-Dm5vuexgnPWzsVYzOxhf0apYaJrKgsvqstKjjwojK94y6SkXE_aem_s6Zn4zipvtVUTsL6u7T4ww" target="_blank" class="darssly-btn">📚 الصف الأول الإعدادي</a>
                <a href="https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim/plans?fbclid=IwY2xjawUKZSFwZG9mAWV4dG4DYWVtAjEwAGJyaWQRMTlWRlpNM3FsN3ViUTA1blBzcnRjBmFwcF9pZBAyMjIwMzkxNzg4MjAwODkyAAEeV3ubjKouoEGuz4whFUVEraSk1byAy7b3Mm6DipOCFjFTI9bFIn4o3xHzkGI_aem_AtBFlzF8bD1TSjD6fNW5uw" target="_blank" class="darssly-btn">📚 الصف الثاني الإعدادي</a>
                <a href="https://darssly.com/courses/mathematics-for-preparatory-stage-mr-mohamed-ghonaim-2/plans?fbclid=IwY2xjawUKZT1wZG9mAWV4dG4DYWVtAjEwAGJyaWQRMTlWRlpNM3FsN3ViUTA1blBzcnRjBmFwcF9pZBAyMjIwMzkxNzg4MjAwODkyAAEe-zyDmeptpawIByu0yeN8QqfHqYDVlWgOVHoQI-Thweh_tYxL17oQONjII7w_aem_bEcNaJAAnCRbg7a77YDWxw" target="_blank" class="darssly-btn">📚 الصف الثالث الإعدادي</a>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if "logged_student" not in st.session_state:
        st.session_state.logged_student = None

    if not st.session_state.logged_student:
        auth_tab1, auth_tab2 = st.tabs(["🔐 تسجيل دخول الطالب", "✨ إنشاء حساب طالب جديد"])

        with auth_tab1:
            st.subheader("سجل دخولك لمتابعة درجاتك وتسجيل حضورك:")
            with st.form("student_login_form"):
                login_name = st.text_input("اسم الطالب المسجل:")
                login_pass = st.text_input("الرقم السري الخاص بك:", type="password")
                btn_login = st.form_submit_button("دخول إلى حسابي")

                if btn_login:
                    users_match = st.session_state.users_df[
                        (st.session_state.users_df["اسم الطالب"].str.strip() == login_name.strip())
                        & (st.session_state.users_df["كلمة المرور"].astype(str).str.strip() == login_pass.strip())
                    ]
                    if not users_match.empty:
                        user_info = users_match.iloc[0].to_dict()
                        st.session_state.logged_student = user_info
                        st.success(f"مرحباً بك مجدداً يا {login_name}!")
                        st.rerun()
                    else:
                        st.error("اسم الطالب أو الرقم السري غير صحيح. إذا لم يكن لديك حساب، يمكنك إنشاؤه من التبويب المجاور.")

        with auth_tab2:
            st.subheader("إنشاء حساب لأول مرة:")
            with st.form("student_register_form"):
                reg_name = st.text_input("اسمك بالكامل (ثلاثي أو رباعي):", placeholder="مثال: أحمد محمود علي")
                reg_curr = st.selectbox("المنهج الدراسي / الدولة:", list(CURRICULUM_DATA.keys()))
                reg_grade = st.selectbox("المرحلة / الصف الدراسي:", CURRICULUM_DATA[reg_curr])
                reg_pass = st.text_input("اختر رقماً سرياً خاصاً بك (احفظه جيداً):", type="password")
                btn_reg = st.form_submit_button("تأكيد إنشاء الحساب")

                if btn_reg:
                    if not reg_name.strip() or not reg_pass.strip():
                        st.error("يرجى كتابة اسمك والرقم السري.")
                    else:
                        existing = st.session_state.users_df[st.session_state.users_df["اسم الطالب"].str.strip() == reg_name.strip()]
                        if not existing.empty:
                            st.warning("هذا الاسم مسجل بالفعل! يرجى تسجيل الدخول مباشرة أو كتابة الاسم ثلاثياً.")
                        else:
                            new_user = {
                                "اسم الطالب": reg_name.strip(),
                                "كلمة المرور": reg_pass.strip(),
                                "المنهج/الدولة": reg_curr,
                                "المجموعة/الصف": reg_grade,
                                "تاريخ التسجيل": str(date.today()),
                            }
                            st.session_state.users_df = pd.concat([st.session_state.users_df, pd.DataFrame([new_user])], ignore_index=True)
                            save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df)
                            st.session_state.logged_student = new_user
                            st.success(f"تم إنشاء حسابك بنجاح يا {reg_name}! تم تسجيل دخولك تلقائياً.")
                            st.rerun()

    else:
        st_user = st.session_state.logged_student
        st.markdown(
            f"""
        <div style="background: rgba(0, 82, 204, 0.08); padding: 15px 20px; border-radius: 10px; border-right: 5px solid #0052cc; margin-bottom: 20px;">
            <h3 style="margin: 0; color: #0052cc;">أهلاً بك: {st_user['اسم الطالب']} 🌟</h3>
            <p style="margin: 5px 0 0 0; font-weight: 800;">{st_user.get('المنهج/الدولة', '')} | {st_user.get('المجموعة/الصف', '')}</p>
        </div>
        """,
            unsafe_allow_html=True,
        )

        st_menu1, st_menu2 = st.tabs(["📝 تسجيل حضور حصة اليوم", "📊 متابعة درجات الواجبات والاختبارات"])

        with st_menu1:
            with st.form("logged_student_att_form", clear_on_submit=True):
                st_date = st.date_input("تاريخ الحصة:", value=date.today())
                rating_options = [
                    "⭐⭐⭐⭐⭐ (5/5) ممتاز جداً وفهم ممتاز",
                    "⭐⭐⭐⭐ (4/5) جيد جداً",
                    "⭐⭐⭐ (3/5) جيد ومفهوم",
                    "⭐⭐ (2/5) متوسط ويحتاج توضيح",
                    "⭐ (1/5) يحتاج إعادة شرح",
                ]
                selected_rating = st.selectbox("⭐ قيّم الحصة مع البشمهندس (من 5 نجوم):", options=rating_options, index=0)
                submit_btn = st.form_submit_button("✅ تأكيد حضور الحصة")

                if submit_btn:
                    new_row = {
                        "التاريخ": str(st_date),
                        "اسم الطالب": st_user["اسم الطالب"],
                        "المنهج/الدولة": st_user.get("المنهج/الدولة", ""),
                        "المجموعة/الصف": st_user.get("المجموعة/الصف", ""),
                        "الحالة": "حاضر",
                        "سعر الحصة": 0.0,
                        "عدد الحصص الكلي": 1,
                        "نظام الدفع": "مؤجل",
                        "مستوى الطالب": "قيد التقييم",
                        "ملاحظات": f"تقييم الحصة: {selected_rating}",
                    }
                    st.session_state.sessions_df = pd.concat([st.session_state.sessions_df, pd.DataFrame([new_row])], ignore_index=True)
                    save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df)
                    st.success(f"تم تسجيل حضورك وتقييمك بنجاح للحصة بتاريخ {st_date}!")

        with st_menu2:
            st.subheader("سجل الواجبات والاختبارات الخاصة بك:")
            my_assessments = st.session_state.assessments_df[
                st.session_state.assessments_df["اسم الطالب"].str.strip() == st_user["اسم الطالب"].strip()
            ].copy()

            if my_assessments.empty:
                st.info("لم يتم رصد واجبات أو اختبارات لك حتى الآن من قبل البشمهندس.")
            else:
                c1, c2 = st.columns(2)
                hw_count = len(my_assessments[my_assessments["النوع"] == "واجب منزلي"])
                exam_count = len(my_assessments[my_assessments["النوع"] == "اختبار دوري"])
                c1.metric("عدد الواجبات المستلمة", hw_count)
                c2.metric("عدد الاختبارات", exam_count)

                st.dataframe(my_assessments[["التاريخ", "النوع", "عنوان التكليف", "الدرجة المحصلة", "الدرجة العظمى", "حالة التسليم", "ملاحظات وتوجيهات"]], use_container_width=True)

        if st.button("🚪 تسجيل الخروج"):
            st.session_state.logged_student = None
            st.rerun()

    # زر الاتصال المباشر بالموبايل
    st.markdown(
        """
        <div class="call-btn-container">
            <a href="tel:01016361440" class="call-btn">
                <svg viewBox="0 0 24 24"><path d="M6.62 10.79c1.44 2.83 3.76 5.14 6.59 6.59l2.2-2.2c.27-.27.67-.36 1.02-.24 1.12.37 2.33.57 3.57.57.55 0 1 .45 1 1V20c0 .55-.45 1-1 1-9.39 0-17-7.61-17-17 0-.55.45-1 1-1h3.5c.55 0 1 .45 1 1 0 1.25.2 2.45.57 3.57.11.35.03.74-.25 1.02l-2.2 2.2z"/></svg>
                <span>للتواصل مع م / محمد غنيم: 01016361440</span>
            </a>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # شريط التواصل وحفظ الحقوق
    st.markdown(
        """
        <div class="social-footer-box">
            <div class="social-icons-container">
                <a href="https://www.facebook.com/share/19fD41rV3H/" target="_blank" title="Facebook" class="social-btn facebook-bg"><svg fill="#ffffff" viewBox="0 0 24 24"><path d="M24 12.073c0-6.627-5.373-12-12-12s-12 5.373-12 12c0 5.99 4.388 10.954 10.125 11.854v-8.385H7.078v-3.47h3.047V9.43c0-3.007 1.792-4.669 4.533-4.669 1.312 0 2.686.235 2.686.235v2.953H15.83c-1.491 0-1.956.925-1.956 1.874v2.25h3.328l-.532 3.47h-2.796v8.385C19.612 23.027 24 18.062 24 12.073z"/></svg></a>
                <a href="https://wa.me/201016361440" target="_blank" title="WhatsApp" class="social-btn whatsapp-bg"><svg fill="#ffffff" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.711 2.599 2.669-.699c.971.53 1.77.822 2.791.823h.002c3.18 0 5.767-2.586 5.768-5.766 0-3.18-2.587-5.766-5.77-5.766zm9.969 5.828c0 5.519-4.481 10-10 10-1.761 0-3.424-.46-4.881-1.267l-5.619 1.474 1.499-5.485c-.911-1.516-1.43-3.285-1.43-5.176 0-5.519 4.481-10 10-10 5.519 0 10 4.481 10 10z"/></svg></a>
                <a href="https://t.me/mrmaths22" target="_blank" title="Telegram" class="social-btn telegram-bg"><svg fill="#ffffff" viewBox="0 0 24 24"><path d="M12 0c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm5.894 8.221l-1.97 9.28c-.145.658-.537.818-1.084.508l-3-2.21-1.446 1.394c-.16.16-.295.295-.605.295l.213-3.053 5.56-5.023c.242-.213-.054-.333-.373-.121l-6.871 4.326-2.962-.924c-.643-.204-.657-.643.136-.953l11.57-4.461c.537-.194 1.006.131.832.942z"/></svg></a>
                <a href="https://www.tiktok.com/@eng_mohamedghonaim?_r=1&_t=ZS-99VdklZPBUS" target="_blank" title="TikTok" class="social-btn tiktok-bg"><svg fill="#ffffff" viewBox="0 0 24 24"><path d="M19.589 6.686a4.793 4.793 0 0 1-3.77-4.245V2h-3.445v13.672a2.896 2.896 0 0 1-5.201 1.743l-.068-.102a2.895 2.895 0 0 1 2.373-4.513c.277 0 .546.039.803.111V9.417a6.338 6.338 0 0 0-.803-.051C6.017 9.366 3.2 12.183 3.2 15.647 3.2 19.11 6.017 22 9.479 22c3.462 0 6.279-2.817 6.279-6.353V9.07c1.378.983 3.054 1.564 4.869 1.584V7.209a4.845 4.845 0 0 1-1.038-.523z"/></svg></a>
                <a href="https://youtube.com/@engineermaths?si=8C6T808VuAU5OMOt" target="_blank" title="YouTube" class="social-btn youtube-bg"><svg fill="#ffffff" viewBox="0 0 24 24"><path d="M23.498 6.186a3.016 3.016 0 0 0-2.122-2.136C19.505 3.545 12 3.545 12 3.545s-7.505 0-9.377.505A3.017 3.017 0 0 0 .502 6.186C0 8.07 0 12 0 12s0 3.93.502 5.814a3.016 3.016 0 0 0 2.122 2.136c1.871.505 9.376.505 9.376.505s7.505 0 9.377-.505a3.015 3.015 0 0 0 2.122-2.136C24 15.93 24 12 24 12s0-3.93-.502-5.814zM9.545 15.568V8.432L15.818 12l-6.273 3.568z"/></svg></a>
            </div>
            <div class="rights-text">جميع الحقوق محفوظة لدي م / محمد غنيم 2026</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
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

st.sidebar.markdown(
    """
<div style="text-align: center; margin-top: 5px; margin-bottom: 20px;">
    <h2 style="margin: 0; color: #0052cc; font-weight: 900;">البشمهندس X الرياضة</h2>
    <p class="sidebar-desc" style="margin: 4px 0; font-weight: 800; font-size: 15px;">نظام إدارة الحصص والتقييمات</p>
</div>
""",
    unsafe_allow_html=True,
)

st.sidebar.header("🔗 رابط تسجيل الطلاب")
st.sidebar.markdown("شارك الرابط مع الطلاب لإنشاء حساباتهم ومتابعة واجباتهم:")
student_url = "https://engmohamedghonaim.streamlit.app/?role=student"
st.sidebar.code(student_url, language="text")

st.markdown(
    f"""
<div class="brand-banner" dir="rtl">
    <div>
        <h1 class="brand-title">البشمهندس X الرياضة 📐</h1>
        <p class="brand-subtitle">لوحة المعلم والإدارة • رصد الحصص والواجبات والاختبارات وإصدار تقارير أولياء الأمور</p>
    </div>
    {'<img src="data:image/jpeg;base64,' + img_b64 + '" style="width: 90px; height: 90px; border-radius: 50%; border: 3px solid #ffffff; object-fit: cover;">' if img_b64 else ''}
</div>
""",
    unsafe_allow_html=True,
)

tab1, tab_hw, tab2, tab3, tab4, tab_users = st.tabs([
    "📝 رصد حصة جديدة",
    "📚 رصد واجب / اختبار",
    "✏️ تعديل ومراجعة الحصص",
    "📊 السجلات الشاملة",
    "🖨️ تقرير ولي الأمر للطباعة",
    "👥 حسابات الطلاب",
])

# ----------------- تبويب 1: رصد الحصة -----------------
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
        submitted = st.form_submit_button("💾 رصد وحفظ الحصة")

        if submitted:
            if not student_name.strip():
                st.error("يرجى كتابة اسم الطالب أولاً.")
            else:
                new_row = {
                    "التاريخ": str(session_date),
                    "اسم الطالب": student_name.strip(),
                    "المنهج/الدولة": t_curriculum,
                    "المجموعة/الصف": group_name,
                    "الحالة": status,
                    "سعر الحصة": price,
                    "عدد الحصص الكلي": total_sessions,
                    "نظام الدفع": payment_type,
                    "مستوى الطالب": student_level,
                    "ملاحظات": notes.strip(),
                }
                st.session_state.sessions_df = pd.concat([st.session_state.sessions_df, pd.DataFrame([new_row])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df)
                st.success(f"✓ تم حفظ سجل الحصة للطالب ({student_name}) بنجاح!")

# ----------------- تبويب 2: رصد الواجبات والاختبارات -----------------
with tab_hw:
    st.subheader("إضافة درجات الواجبات المنزلية والاختبارات الدورية")

    all_registered_names = sorted(list(set(
        [s for s in st.session_state.users_df["اسم الطالب"].dropna().unique() if str(s).strip()]
        + [s for s in st.session_state.sessions_df["اسم الطالب"].dropna().unique() if str(s).strip()]
    )))

    if not all_registered_names:
        st.info("لا يوجد طلاب مسجلون بعد. يمكنك تسجيل طالب يدوياً أدناه:")

    with st.form("assessment_form", clear_on_submit=True):
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            ass_student = st.selectbox("اختر الطالب المسجل:", all_registered_names) if all_registered_names else st.text_input("اسم الطالب:")
            ass_type = st.selectbox("نوع التكليف الأكاديمي:", ["واجب منزلي", "اختبار دوري", "كويز سريع", "مهمة إضافية"])
            ass_title = st.text_input("عنوان الدرس / التكليف:", placeholder="مثال: حل تمارين نظرية فيثاغورس صـ 14")
            ass_date = st.date_input("تاريخ الرصد:", value=date.today())

        with col_a2:
            ass_status = st.selectbox("حالة التسليم والحل:", ["تم التسليم كاملاً وبشكل ممتاز", "تم التسليم مع بعض الأخطاء", "تسليم جزئي / ناقص", "لم يتم التسليم (مقصّر)", "غائب عن الاختبار"])
            ass_score = st.number_input("الدرجة التي حصل عليها الطالب:", min_value=0.0, step=0.5, value=10.0)
            ass_max = st.number_input("الدرجة العظمى (النهائية):", min_value=1.0, step=1.0, value=10.0)
            ass_notes = st.text_input("ملاحظة وتوجيه ولي الأمر:", placeholder="توجيه مباشر يظهر في تقرير ولي الأمر...")

        save_ass_btn = st.form_submit_button("💾 رصد التقييم وحفظه")

        if save_ass_btn:
            if not str(ass_student).strip() or not ass_title.strip():
                st.error("يرجى التأكد من اختيار اسم الطالب وكتابة عنوان التكليف.")
            else:
                new_ass = {
                    "التاريخ": str(ass_date),
                    "اسم الطالب": str(ass_student).strip(),
                    "النوع": ass_type,
                    "عنوان التكليف": ass_title.strip(),
                    "الدرجة المحصلة": ass_score,
                    "الدرجة العظمى": ass_max,
                    "حالة التسليم": ass_status,
                    "ملاحظات وتوجيهات": ass_notes.strip(),
                }
                st.session_state.assessments_df = pd.concat([st.session_state.assessments_df, pd.DataFrame([new_ass])], ignore_index=True)
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df)
                st.success(f"✓ تم رصد {ass_type} بنجاح للطالب ({ass_student}) بنتيجة ({ass_score} / {ass_max})!")

    st.write("---")
    st.write("### سجل الواجبات والاختبارات المرصودة مؤخراً:")
    st.dataframe(st.session_state.assessments_df, use_container_width=True)

# ----------------- تبويب 3: التعديل والمراجعة -----------------
with tab2:
    st.subheader("مراجعة وتعديل بيانات الطلاب والحصص")
    df = st.session_state.sessions_df

    if df.empty:
        st.info("لا توجد سجلات حصص مسجلة بعد.")
    else:
        record_options = {
            idx: f"[{row['التاريخ']}] - {row['اسم الطالب']} ({row.get('المنهج/الدولة', '')} | {row['المجموعة/الصف']}) - {row['الحالة']}"
            for idx, row in df.iterrows()
        }
        selected_idx = st.selectbox("اختر السجل المراد تعديل بياناته:", options=list(record_options.keys()), format_func=lambda x: record_options[x])

        selected_row = df.loc[selected_idx]
        try:
            curr_date = datetime.strptime(str(selected_row["التاريخ"]), "%Y-%m-%d").date()
        except Exception:
            curr_date = date.today()

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
                try:
                    p_val = float(selected_row["سعر الحصة"])
                except Exception:
                    p_val = 0.0
                edit_price = st.number_input("سعر الحصة:", min_value=0.0, step=10.0, value=p_val)
                try:
                    s_val = int(selected_row["عدد الحصص الكلي"])
                except Exception:
                    s_val = 1
                edit_sessions = st.number_input("عدد الحصص الكلي:", min_value=1, step=1, value=s_val)
                p_opts = ["اشتراك شهري", "مقدم", "مؤخر (بعد الحصة)", "مؤجل"]
                edit_pay = st.selectbox("نظام الدفع:", p_opts, index=(p_opts.index(selected_row["نظام الدفع"]) if selected_row["نظام الدفع"] in p_opts else 0))
                l_opts = ["ممتاز ⭐⭐⭐", "جيد جداً ⭐⭐", "جيد ⭐", "متوسط", "يحتاج متابعة", "قيد التقييم"]
                edit_level = st.selectbox("المستوى:", l_opts, index=(l_opts.index(selected_row["مستوى الطالب"]) if selected_row["مستوى الطالب"] in l_opts else 0))

            edit_notes = st.text_area("الملاحظات والتقييم:", value=str(selected_row["ملاحظات"]))

            b1, b2 = st.columns([1, 4])
            with b1:
                update_btn = st.form_submit_button("💾 تحديث البيانات")
            with b2:
                delete_btn = st.form_submit_button("🗑️ حذف السجل")

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
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df)
                st.success("✓ تم تحديث بيانات الحصة بنجاح!")
                st.rerun()

            if delete_btn:
                df = df.drop(selected_idx).reset_index(drop=True)
                st.session_state.sessions_df = df
                save_all_data(st.session_state.users_df, st.session_state.sessions_df, st.session_state.assessments_df)
                st.warning("⚠️ تم حذف السجل.")
                st.rerun()

# ----------------- تبويب 4: السجلات العامة -----------------
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
        if filter_curr != "الكل":
            filtered = filtered[filtered["المنهج/الدولة"] == filter_curr]
        if filter_group != "الكل":
            filtered = filtered[filtered["المجموعة/الصف"] == filter_group]
        if search_name.strip():
            filtered = filtered[filtered["اسم الطالب"].str.contains(search_name.strip(), na=False)]

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

        st.download_button(
            label="📥 تصدير ملف Excel الشامل (مستخدمين + حصص + واجبات)",
            data=buf.getvalue(),
            file_name=FILE_NAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ----------------- تبويب 5: تقرير ولي الأمر للطباعة (مُحدّث بالأسعار والإجمالي) -----------------
with tab4:
    st.subheader("📑 إصدار وطباعة تقرير متابعة الطالب لولي الأمر (شامل الأسعار والحساب المالي)")
    
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

            curr_val = "-"
            group_val = "-"
            level_val = "جيد"
            pay_val = "مؤجل"
            if not st_sessions.empty:
                latest = st_sessions.iloc[-1]
                curr_val = latest.get("المنهج/الدولة", "-")
                group_val = latest.get("المجموعة/الصف", "-")
                level_val = latest.get("مستوى الطالب", "جيد")
                pay_val = latest.get("نظام الدفع", "مؤجل")

            att_cnt = len(st_sessions[st_sessions["الحالة"] == "حاضر"])
            abs_cnt = len(st_sessions[st_sessions["الحالة"] == "غائب"])
            total_sessions_cnt = len(st_sessions)
            att_percentage = (att_cnt / total_sessions_cnt * 100) if total_sessions_cnt > 0 else 0

            # حساب إجمالي الحساب المالي
            total_cost = 0.0
            if not st_sessions.empty and "سعر الحصة" in st_sessions.columns:
                total_cost = st_sessions["سعر الحصة"].astype(float, errors="ignore").sum(numeric_only=True)

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

            # جدول الواجبات والاختبارات
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

            # جدول الحصص شاملاً سعر كل حصة وسطراً للإجمالي النهائي
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

            # قالب PDF الجاهز للطباعة
            parent_report_html = f"""<!DOCTYPE html>
            <html dir="rtl" lang="ar">
            <head>
                <meta charset="utf-8">
                <title>تقرير متابعة ولي الأمر - {selected_student}</title>
                <style>
                    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@700;800;900&display=swap');
                    body {{
                        font-family: 'Cairo', Tahoma, Arial, sans-serif;
                        padding: 30px;
                        color: #000000 !important;
                        background-color: #ffffff !important;
                        font-weight: 800;
                    }}
                    .header-box {{
                        border-bottom: 3px solid #0052cc;
                        padding-bottom: 15px;
                        margin-bottom: 25px;
                        display: flex;
                        justify-content: space-between;
                        align-items: center;
                    }}
                    .brand-name {{ color: #0052cc; margin: 0; font-size: 30px; font-weight: 900; }}
                    .brand-sub {{ color: #000000; margin: 4px 0; font-size: 15px; font-weight: 800; }}
                    .parent-notice {{
                        background-color: #f0fdf4;
                        border: 2px solid #16a34a;
                        padding: 12px 18px;
                        border-radius: 8px;
                        margin-bottom: 25px;
                        font-size: 16px;
                        font-weight: 800;
                        color: #14532d;
                    }}
                    .stats-box {{ width: 100%; border-collapse: collapse; margin-bottom: 25px; }}
                    .stats-box th, .stats-box td {{ border: 2px solid #000000; padding: 10px; text-align: center; font-size: 15px; font-weight: 800; }}
                    .stats-box th {{ background: #f1f5f9; color: #000000; font-weight: 900; }}
                    .section-title {{
                        margin-top: 25px;
                        margin-bottom: 10px;
                        color: #0052cc;
                        font-size: 19px;
                        font-weight: 900;
                        border-bottom: 2px solid #0052cc;
                        padding-bottom: 6px;
                        display: inline-block;
                    }}
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

                <div class="parent-notice">
                    📌 <b>رسالة لولي الأمر الفاضل:</b> نحرص دائماً على إحاطة سيادتكم بالمستوى الفعلي للطالب أولاً بأول، وتفاصيل الحساب المالي للحصص، من أجل التعاون المستمر نحو التفوق بإذن الله.
                </div>

                <table class="stats-box">
                    <tr>
                        <th>إجمالي الحصص</th>
                        <th>مرات الحضور</th>
                        <th>مرات الغياب</th>
                        <th>نسبة الالتزام بالحضور</th>
                        <th>إجمالي الحساب المستحق</th>
                        <th>المستوى العام</th>
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
                    <tr>
                        <th>التاريخ</th>
                        <th>النوع</th>
                        <th>عنوان التكليف / الاختبار</th>
                        <th>الدرجة المحصلة</th>
                        <th>حالة التسليم والالتزام</th>
                        <th>ملاحظات المعلم وتوجيهاته</th>
                    </tr>
                    {ass_html_rows}
                </table>

                <div class="section-title">2. سجل الحضور وتفاصيل سعر كل حصة:</div>
                <table class="table-main">
                    <tr>
                        <th>التاريخ</th>
                        <th>حالة الحضور</th>
                        <th>سعر الحصة</th>
                        <th>مستوى الطالب بالحصة</th>
                        <th>ملاحظات التفاعل والاستيعاب</th>
                    </tr>
                    {session_html_rows}
                </table>

                <div class="footer-note">
                    مع تحيات: <b>البشمهندس X الرياضة | م / محمد غنيم</b> — رقم التواصل المباشر: 01016361440 🌟
                </div>
            </body>
            </html>"""

            st.download_button(
                label=f"🖨️ تحميل وطباعة تقرير ولي الأمر لـ ({selected_student}) بصيغة PDF",
                data=parent_report_html.encode("utf-8"),
                file_name=f"تقرير_ولي_الأمر_{selected_student}.html",
                mime="text/html",
            )

# ----------------- تبويب 6: حسابات الطلاب المسجلين -----------------
with tab_users:
    st.subheader("👥 قائمة حسابات الطلاب المسجلين وكلمات المرور:")
    st.write("يمكنك الاطلاع على كلمات مرور الطلاب لمساعدتهم في حال نسيانها:")
    st.dataframe(st.session_state.users_df, use_container_width=True)
